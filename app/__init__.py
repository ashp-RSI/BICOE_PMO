import atexit
import logging
import threading
import time

from flask import Flask
from flask_caching import Cache

cache = Cache()
logger = logging.getLogger(__name__)
_boot_ts = int(time.time())

# Module-level singletons populated by create_app() so other modules can
# import them after init.
notification_store = None
_scheduler = None

DEMAND_HEADERS = [
    "Requisition ID", "Yrs of Exp", "Skillset", "Demand Status", "Notes",
    "Customer Name", "Fulfillment Type", "Mapped Emp Code", "Mapped Emp Name",
    "Mapping Date",
]

HEADCOUNT_BILLABLE_COLUMNS = [
    "Customer interview happened(Yes/No)",
    "Customer Selected(Yes/No)",
]

HEADCOUNT_EXTRA_COLUMNS = [
    "Comments",
]


def create_app():
    app = Flask(__name__)
    app.config.from_object("app.config.Config")
    app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

    cache.init_app(app)

    @app.context_processor
    def inject_cache_bust():
        return {"cache_bust": _boot_ts}

    from app.routes.views import views_bp
    from app.routes.api import api_bp

    app.register_blueprint(views_bp)
    app.register_blueprint(api_bp, url_prefix="/api")

    _init_notifications(app)

    def _deferred_init():
        time.sleep(2)
        with app.app_context():
            _init_demand_sheet(app.config)

    threading.Thread(target=_deferred_init, daemon=True).start()

    return app


def _init_notifications(app):
    """Initialize the SQLite notification store and start the reminder
    scheduler. Safe to call once per process."""
    global notification_store, _scheduler

    if not app.config.get("NOTIF_ENABLED", True):
        logger.info("Manager notifications disabled (NOTIF_ENABLED=false)")
        return

    try:
        from app.services.notification_store import NotificationStore
        notification_store = NotificationStore(app.config["NOTIF_DB_PATH"])
        logger.info("Notification store initialized at %s",
                    app.config["NOTIF_DB_PATH"])
    except Exception:
        logger.exception("Failed to initialize notification store — "
                         "manager notifications will be disabled")
        return

    try:
        from apscheduler.schedulers.background import BackgroundScheduler
        from app.services.reminder_worker import ReminderWorker
        from app.services.sharepoint_service import SharePointService

        config = app.config

        def _sp_factory():
            return SharePointService(config)

        worker = ReminderWorker(
            store=notification_store,
            sp_service_factory=_sp_factory,
            reminder_days=config["NOTIF_REMINDER_DAYS"],
            max_reminders=config["NOTIF_MAX_REMINDERS"],
        )

        _scheduler = BackgroundScheduler(daemon=True, timezone="UTC")
        _scheduler.add_job(
            worker.run_once,
            trigger="interval",
            minutes=config["NOTIF_POLL_INTERVAL_MINUTES"],
            id="notification_reminder_worker",
            replace_existing=True,
            max_instances=1,
            coalesce=True,
        )
        _scheduler.start()
        logger.info(
            "Reminder scheduler started — every %d min, "
            "reminder cadence %d days, max %d reminders",
            config["NOTIF_POLL_INTERVAL_MINUTES"],
            config["NOTIF_REMINDER_DAYS"],
            config["NOTIF_MAX_REMINDERS"],
        )

        atexit.register(lambda: _scheduler.shutdown(wait=False)
                        if _scheduler else None)
    except Exception:
        logger.exception("Failed to start reminder scheduler — reminders "
                         "will not run automatically")


def _init_demand_sheet(config):
    """Ensure the Demand Requisition sheet and required headcount columns
    exist in the SharePoint workbook on startup."""
    try:
        from app.services.sharepoint_service import SharePointService

        sp = SharePointService(config)
        demand_sheet = config.get("DEMAND_SHEET_NAME", "Demand Requisition")
        headcount_sheet = config.get("SHEET_NAME", "Sheet1")

        sp.ensure_multiple_sheets({
            demand_sheet: DEMAND_HEADERS,
            headcount_sheet: HEADCOUNT_BILLABLE_COLUMNS + HEADCOUNT_EXTRA_COLUMNS,
        })
        logger.info("Demand Requisition sheet and headcount columns verified in SharePoint")

        updated = sp.backfill_billable_columns(HEADCOUNT_BILLABLE_COLUMNS)
        if updated:
            logger.info("Backfilled %d Billable employees with Yes/Yes", updated)

        cleared = sp.clear_column_value("Comments", "Yes")
        if cleared:
            logger.info("Cleared %d stale 'Yes' values from Comments column", cleared)
    except Exception:
        logger.exception(
            "Failed to initialize sheets — "
            "missing columns will be created on first write"
        )

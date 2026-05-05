"""Background reminder worker driven by APScheduler.

Runs every NOTIF_POLL_INTERVAL_MINUTES. For each notification still in
`awaiting_reply` state:
  - If now - sent_at (or last_reminder_at) >= NOTIF_REMINDER_DAYS days
    AND reminder_count < NOTIF_MAX_REMINDERS  →  send a reminder email.
  - If reminder_count >= NOTIF_MAX_REMINDERS and overdue → mark `no_response`.
"""
import logging
from datetime import datetime, timedelta, timezone

from app.services.email_service import EmailService, EmailServiceError
from app.services.notification_store import (
    NotificationStore, STATUS_NO_RESPONSE,
)

logger = logging.getLogger(__name__)


def _parse_iso(s):
    """Parse an ISO timestamp string (timezone-naive UTC) into datetime."""
    if not s:
        return None
    try:
        # Accept both 'Z' and naive forms
        if s.endswith("Z"):
            s = s[:-1]
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except (ValueError, TypeError):
        return None


def _build_action_buttons(notification_id, base_url, secret_key,
                          notification_type="proposed"):
    """Generate signed action button HTML for a reminder email."""
    from itsdangerous import URLSafeTimedSerializer

    s = URLSafeTimedSerializer(secret_key)
    salt = "notif-action"

    if notification_type == "blocked":
        confirm_token = s.dumps(
            {"nid": notification_id, "act": "confirm_blocked"}, salt=salt)
        reject_token = s.dumps(
            {"nid": notification_id, "act": "reject_blocked"}, salt=salt)
        confirm_url = f"{base_url}/api/notifications/blocked-action?token={confirm_token}"
        reject_url = f"{base_url}/api/notifications/blocked-action?token={reject_token}"
        return f"""
            <table cellpadding="0" cellspacing="0" border="0" style="margin-top:20px">
              <tr>
                <td style="padding-right:12px">
                  <a href="{confirm_url}"
                     style="background-color:#28a745;color:#ffffff;padding:12px 28px;
                            text-decoration:none;border-radius:5px;font-weight:bold;
                            display:inline-block;font-size:14px">
                    &#10004; Yes, Allocation Completed
                  </a>
                </td>
                <td>
                  <a href="{reject_url}"
                     style="background-color:#dc3545;color:#ffffff;padding:12px 28px;
                            text-decoration:none;border-radius:5px;font-weight:bold;
                            display:inline-block;font-size:14px">
                    &#10008; No
                  </a>
                </td>
              </tr>
            </table>
            <p style="color:#888;font-size:11px;margin-top:8px">
              Click a button above to respond directly, or reply to this email.
            </p>
        """

    approve_token = s.dumps({"nid": notification_id, "act": "approve"}, salt=salt)
    reject_token = s.dumps({"nid": notification_id, "act": "reject"}, salt=salt)
    approve_url = f"{base_url}/api/notifications/action?token={approve_token}"
    reject_url = f"{base_url}/api/notifications/action?token={reject_token}"
    return f"""
        <table cellpadding="0" cellspacing="0" border="0" style="margin-top:20px">
          <tr>
            <td style="padding-right:12px">
              <a href="{approve_url}"
                 style="background-color:#28a745;color:#ffffff;padding:12px 28px;
                        text-decoration:none;border-radius:5px;font-weight:bold;
                        display:inline-block;font-size:14px">
                &#10004; Approve
              </a>
            </td>
            <td>
              <a href="{reject_url}"
                 style="background-color:#dc3545;color:#ffffff;padding:12px 28px;
                        text-decoration:none;border-radius:5px;font-weight:bold;
                        display:inline-block;font-size:14px">
                &#10008; Reject
              </a>
            </td>
          </tr>
        </table>
        <p style="color:#888;font-size:11px;margin-top:8px">
          Click a button above to respond directly, or reply to this email.
        </p>
    """


def _build_reminder_html(n, reminder_idx, max_reminders,
                         base_url=None, secret_key=None):
    emp_name = n.get("emp_name") or n["emp_code"]
    req_id = n.get("demand_req_id") or "N/A"
    customer = n.get("customer_name") or "N/A"
    sent_at_dt = _parse_iso(n.get("sent_at"))
    sent_str = sent_at_dt.strftime("%d %b %Y") if sent_at_dt else "earlier"
    notification_type = n.get("notification_type") or "proposed"

    buttons = ""
    if base_url and secret_key:
        buttons = _build_action_buttons(
            n["id"], base_url, secret_key, notification_type=notification_type)

    if notification_type == "blocked":
        return f"""
            <p>Hi,</p>
            <p>This is a friendly reminder ({reminder_idx} of {max_reminders})
               regarding <b>{emp_name}</b> ({n['emp_code']}) who has been
               <b>Blocked</b> — we haven't received your response yet.</p>
            <p>Please confirm the allocation start date.</p>
            <table cellpadding="6" style="border-collapse:collapse;font-size:13px;border:1px solid #ddd">
              <tr><td><b>Originally sent:</b></td><td>{sent_str}</td></tr>
            </table>
            {buttons}
            <p style="color:#888;font-size:12px">
               — Internal Project Management Tool (automated reminder)
            </p>
        """

    return f"""
        <p>Hi,</p>
        <p>This is a friendly reminder ({reminder_idx} of {max_reminders})
           regarding the proposed allocation for <b>{emp_name}</b> ({n['emp_code']})
           — we haven't received your confirmation yet.</p>
        <p>Please confirm whether the allocation should proceed?</p>
        <p>Reply with <b>Yes</b> to approve or <b>No</b> to reject.</p>
        <table cellpadding="6" style="border-collapse:collapse;font-size:13px;border:1px solid #ddd">
          <tr><td><b>Originally sent:</b></td><td>{sent_str}</td></tr>
        </table>
        {buttons}
        <p>Please <u>Note</u> : Allocation priority will be considered based on the earliest confirmed start
           date in case of multiple proposals.</p>
        <p style="color:#888;font-size:12px">
           — Internal Project Management Tool (automated reminder)
        </p>
    """


class ReminderWorker:
    def __init__(self, store: NotificationStore, sp_service_factory,
                 reminder_days: int, max_reminders: int,
                 app_base_url: str = "", secret_key: str = ""):
        """
        Args:
            store: NotificationStore instance.
            sp_service_factory: callable returning a SharePointService
                                (used by EmailService for MSAL auth).
            reminder_days: days between reminder emails.
            max_reminders: total reminder emails after the initial one.
            app_base_url: public URL for building approve/reject links.
            secret_key: Flask SECRET_KEY for signing action tokens.
        """
        self.store = store
        self._sp_factory = sp_service_factory
        self.reminder_days = reminder_days
        self.max_reminders = max_reminders
        self.app_base_url = app_base_url
        self.secret_key = secret_key

    def run_once(self):
        """One pass over all awaiting notifications. Safe to call from a
        scheduler or manually for testing."""
        try:
            pending = self.store.list_pending_for_reminder()
        except Exception:
            logger.exception("Reminder worker: failed to load pending list")
            return

        if not pending:
            return

        now = datetime.now(timezone.utc)
        threshold = timedelta(days=self.reminder_days)

        email_svc = None
        for n in pending:
            try:
                last_ts = _parse_iso(
                    n.get("last_reminder_at") or n.get("sent_at")
                )
                if not last_ts:
                    continue
                if now - last_ts < threshold:
                    continue

                if n["reminder_count"] >= self.max_reminders:
                    self.store.update_status(
                        n["id"], STATUS_NO_RESPONSE,
                        decided_by="system",
                        decision_note=(
                            f"No reply after {self.max_reminders} reminders"
                        ),
                    )
                    logger.info(
                        "Notification %d marked no_response (emp=%s)",
                        n["id"], n["emp_code"],
                    )
                    continue

                if email_svc is None:
                    email_svc = EmailService(self._sp_factory())

                reminder_idx = n["reminder_count"] + 1
                subject = "Reminder: " + (
                    n.get("subject") or "Proposed allocation needs your approval"
                )
                if not subject.lower().startswith("reminder"):
                    subject = "Reminder: " + subject
                html = _build_reminder_html(
                    n, reminder_idx, self.max_reminders,
                    base_url=self.app_base_url,
                    secret_key=self.secret_key,
                )

                email_svc.send_mail(
                    to_email=n["manager_email"],
                    subject=subject,
                    html_body=html,
                    cc_emails=n.get("cc_emails") or [],
                    save_to_sent=True,
                )
                self.store.record_reminder_sent(n["id"])
                logger.info(
                    "Reminder %d/%d sent for notification %d (emp=%s, mgr=%s)",
                    reminder_idx, self.max_reminders, n["id"],
                    n["emp_code"], n["manager_email"],
                )
            except EmailServiceError:
                logger.exception(
                    "Reminder worker: send failed for notification %d",
                    n.get("id"),
                )
            except Exception:
                logger.exception(
                    "Reminder worker: unexpected error on notification %d",
                    n.get("id"),
                )

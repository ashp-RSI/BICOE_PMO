import os
from dotenv import load_dotenv

load_dotenv()

_IS_AZURE = bool(os.getenv("WEBSITE_SITE_NAME"))
_PROJECT_ROOT = os.path.dirname(os.path.dirname(__file__))


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "dev-secret-key-change-in-prod")
    SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
    SHAREPOINT_USERNAME = os.getenv("SHAREPOINT_USERNAME")
    SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")
    TENANT_ID = os.getenv("TENANT_ID")
    FILE_RELATIVE_URL = os.getenv("FILE_RELATIVE_URL")
    SHEET_NAME = os.getenv("SHEET_NAME", "Head Count Report")
    DEMAND_SHEET_NAME = os.getenv("DEMAND_SHEET_NAME", "Demand Requisition")
    SHAREPOINT_HOST = "rsystemsiltd.sharepoint.com"
    MSAL_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    IS_AZURE = _IS_AZURE
    TOKEN_CACHE_FILE = (
        "/home/.token_cache.bin" if _IS_AZURE
        else os.path.join(_PROJECT_ROOT, ".token_cache.bin")
    )
    CACHE_TYPE = "SimpleCache"
    CACHE_DEFAULT_TIMEOUT = 300

    # ── Manager Notification Settings ─────────────────────────────
    NOTIF_DB_PATH = (
        "/home/notifications.db" if _IS_AZURE
        else os.path.join(_PROJECT_ROOT, "notifications.db")
    )
    NOTIF_REMINDER_DAYS = int(os.getenv("NOTIF_REMINDER_DAYS", "2"))
    NOTIF_MAX_REMINDERS = int(os.getenv("NOTIF_MAX_REMINDERS", "3"))
    NOTIF_POLL_INTERVAL_MINUTES = int(
        os.getenv("NOTIF_POLL_INTERVAL_MINUTES", "30")
    )
    NOTIF_FUZZY_THRESHOLD = int(os.getenv("NOTIF_FUZZY_THRESHOLD", "90"))
    NOTIF_FROM_NAME = os.getenv(
        "NOTIF_FROM_NAME", "Internal Project Management Tool"
    )
    NOTIF_ENABLED = os.getenv("NOTIF_ENABLED", "true").lower() in ("1", "true", "yes")

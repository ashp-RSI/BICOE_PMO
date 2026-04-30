"""SQLite-backed store for manager notifications.

Tracks every "Notify Manager" email sent, the resolution method used,
the manager's reply state, and reminder cadence. Schema is forward-compatible
with the future Option A (Graph Mail.Read polling) upgrade — we already
persist conversation_id and message_id when Graph returns them.
"""
import json
import logging
import os
import sqlite3
import threading
from datetime import datetime

logger = logging.getLogger(__name__)

# Notification status constants
STATUS_AWAITING = "awaiting_reply"
STATUS_APPROVED = "approved"
STATUS_REJECTED = "rejected"
STATUS_NO_RESPONSE = "no_response"
STATUS_CANCELLED = "cancelled"

ACTIVE_STATUSES = {STATUS_AWAITING}

_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS notifications (
  id                INTEGER PRIMARY KEY AUTOINCREMENT,
  emp_code          TEXT NOT NULL,
  emp_name          TEXT,
  emp_row_index     INTEGER,
  demand_row_index  INTEGER,
  demand_req_id     TEXT,
  customer_name     TEXT,
  manager_name      TEXT,
  manager_email     TEXT NOT NULL,
  resolution_method TEXT,
  cc_emails         TEXT,
  subject           TEXT,
  body_html         TEXT,
  conversation_id   TEXT,
  message_id        TEXT,
  status            TEXT NOT NULL,
  reminder_count    INTEGER DEFAULT 0,
  sent_at           TEXT NOT NULL,
  last_reminder_at  TEXT,
  decided_by        TEXT,
  decided_at        TEXT,
  decision_note     TEXT
);

CREATE INDEX IF NOT EXISTS idx_notifications_status        ON notifications(status);
CREATE INDEX IF NOT EXISTS idx_notifications_emp_code      ON notifications(emp_code);
CREATE INDEX IF NOT EXISTS idx_notifications_demand_req_id ON notifications(demand_req_id);
CREATE INDEX IF NOT EXISTS idx_notifications_emp_row       ON notifications(emp_row_index);
"""


class NotificationStore:
    """Thread-safe SQLite wrapper. APScheduler runs in a background thread
    and the Flask request thread also writes — so a single per-store lock
    keeps things simple. Volume is low (one row per Notify action)."""

    def __init__(self, db_path):
        self.db_path = db_path
        self._lock = threading.Lock()
        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
        self._init_schema()

    def _connect(self):
        conn = sqlite3.connect(self.db_path, timeout=10, isolation_level=None)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    def _init_schema(self):
        with self._lock, self._connect() as conn:
            conn.executescript(_SCHEMA_SQL)

    @staticmethod
    def _row_to_dict(row):
        if row is None:
            return None
        d = dict(row)
        if d.get("cc_emails"):
            try:
                d["cc_emails"] = json.loads(d["cc_emails"])
            except (TypeError, ValueError):
                d["cc_emails"] = []
        else:
            d["cc_emails"] = []
        return d

    # ── CRUD ────────────────────────────────────────────────────

    def create(self, *, emp_code, emp_name, emp_row_index, demand_row_index,
               demand_req_id, customer_name, manager_name, manager_email,
               resolution_method, cc_emails, subject, body_html,
               conversation_id=None, message_id=None, sent_at=None):
        sent_at = sent_at or datetime.utcnow().isoformat()
        cc_json = json.dumps(cc_emails or [])
        with self._lock, self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO notifications (
                    emp_code, emp_name, emp_row_index,
                    demand_row_index, demand_req_id, customer_name,
                    manager_name, manager_email, resolution_method,
                    cc_emails, subject, body_html,
                    conversation_id, message_id,
                    status, reminder_count, sent_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
                """,
                (
                    str(emp_code), emp_name, emp_row_index,
                    demand_row_index, demand_req_id, customer_name,
                    manager_name, manager_email, resolution_method,
                    cc_json, subject, body_html,
                    conversation_id, message_id,
                    STATUS_AWAITING, sent_at,
                ),
            )
            return cur.lastrowid

    def get(self, notification_id):
        with self._lock, self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM notifications WHERE id = ?", (notification_id,)
            ).fetchone()
            return self._row_to_dict(row)

    def get_active_for_employee(self, emp_code):
        """Return the most recent active (awaiting_reply) notification, if any.
        Used by the per-row badge in the headcount table."""
        with self._lock, self._connect() as conn:
            row = conn.execute(
                """SELECT * FROM notifications
                   WHERE emp_code = ? AND status = ?
                   ORDER BY id DESC LIMIT 1""",
                (str(emp_code), STATUS_AWAITING),
            ).fetchone()
            return self._row_to_dict(row)

    def get_latest_for_employee(self, emp_code):
        """Return the most recent notification of any status."""
        with self._lock, self._connect() as conn:
            row = conn.execute(
                """SELECT * FROM notifications
                   WHERE emp_code = ?
                   ORDER BY id DESC LIMIT 1""",
                (str(emp_code),),
            ).fetchone()
            return self._row_to_dict(row)

    def list(self, status=None, emp_code=None, limit=500):
        with self._lock, self._connect() as conn:
            sql = "SELECT * FROM notifications WHERE 1=1"
            params = []
            if status:
                sql += " AND status = ?"
                params.append(status)
            if emp_code:
                sql += " AND emp_code = ?"
                params.append(str(emp_code))
            sql += " ORDER BY id DESC LIMIT ?"
            params.append(limit)
            rows = conn.execute(sql, params).fetchall()
            return [self._row_to_dict(r) for r in rows]

    def list_pending_for_reminder(self):
        """All notifications still awaiting reply, oldest first."""
        with self._lock, self._connect() as conn:
            rows = conn.execute(
                """SELECT * FROM notifications
                   WHERE status = ?
                   ORDER BY sent_at ASC""",
                (STATUS_AWAITING,),
            ).fetchall()
            return [self._row_to_dict(r) for r in rows]

    def list_all_active_codes(self):
        """Return the set of emp_codes that currently have an active notification.
        Used by the headcount endpoint to attach badge state efficiently."""
        with self._lock, self._connect() as conn:
            rows = conn.execute(
                """SELECT emp_code, status, sent_at, reminder_count
                   FROM notifications
                   WHERE id IN (
                       SELECT MAX(id) FROM notifications GROUP BY emp_code
                   )"""
            ).fetchall()
            return {
                str(r["emp_code"]): {
                    "status": r["status"],
                    "sent_at": r["sent_at"],
                    "reminder_count": r["reminder_count"],
                }
                for r in rows
            }

    def update_status(self, notification_id, status, *,
                      decided_by=None, decision_note=None):
        decided_at = datetime.utcnow().isoformat()
        with self._lock, self._connect() as conn:
            conn.execute(
                """UPDATE notifications
                   SET status = ?, decided_by = ?, decided_at = ?, decision_note = ?
                   WHERE id = ?""",
                (status, decided_by, decided_at, decision_note, notification_id),
            )

    def record_reminder_sent(self, notification_id):
        with self._lock, self._connect() as conn:
            conn.execute(
                """UPDATE notifications
                   SET reminder_count = reminder_count + 1,
                       last_reminder_at = ?
                   WHERE id = ?""",
                (datetime.utcnow().isoformat(), notification_id),
            )

    def delete(self, notification_id):
        """Remove a notification row entirely (used when email send fails
        immediately after creating the record)."""
        with self._lock, self._connect() as conn:
            conn.execute(
                "DELETE FROM notifications WHERE id = ?", (notification_id,)
            )

    def record_send_ids(self, notification_id, *, conversation_id=None,
                        message_id=None):
        """Store the Graph conversation/message IDs captured after sending."""
        with self._lock, self._connect() as conn:
            conn.execute(
                """UPDATE notifications
                   SET conversation_id = COALESCE(?, conversation_id),
                       message_id = COALESCE(?, message_id)
                   WHERE id = ?""",
                (conversation_id, message_id, notification_id),
            )

    def reset_for_resend(self, notification_id, *, conversation_id=None,
                        message_id=None, sent_at=None):
        """Reset a notification to awaiting_reply on a manual resend."""
        sent_at = sent_at or datetime.utcnow().isoformat()
        with self._lock, self._connect() as conn:
            conn.execute(
                """UPDATE notifications
                   SET status = ?, reminder_count = 0,
                       sent_at = ?, last_reminder_at = NULL,
                       conversation_id = COALESCE(?, conversation_id),
                       message_id = COALESCE(?, message_id),
                       decided_by = NULL, decided_at = NULL,
                       decision_note = NULL
                   WHERE id = ?""",
                (
                    STATUS_AWAITING, sent_at, conversation_id, message_id,
                    notification_id,
                ),
            )

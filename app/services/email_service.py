"""Microsoft Graph email service. Uses the same MSAL public client as
SharePointService, with a separate `Mail.Send` scope.

For the initial Option C release we only need Mail.Send. The captured
conversationId/messageId are saved with each notification so an Option A
(Mail.Read polling) upgrade requires no schema change later.
"""
import logging

import requests

logger = logging.getLogger(__name__)

GRAPH_SENDMAIL_URL = "https://graph.microsoft.com/v1.0/me/sendMail"
GRAPH_SENT_ITEMS_URL = (
    "https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages"
)
GRAPH_SCOPES = ["https://graph.microsoft.com/Mail.Send"]


class EmailServiceError(RuntimeError):
    """Raised when the Graph API rejects a sendMail request."""


class EmailService:
    """Send mail via Microsoft Graph reusing the SharePoint MSAL session."""

    def __init__(self, sp_service):
        self._sp = sp_service
        self._app = sp_service._app

    def _get_token(self):
        """Acquire a Graph token, falling back to interactive device flow.

        Note: the cached account from the SharePoint sign-in is reused, so the
        user typically isn't prompted again unless Mail.Send needs explicit
        consent.
        """
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(
                GRAPH_SCOPES, account=accounts[0]
            )
            if result and "access_token" in result:
                return result["access_token"]

        flow = self._app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            raise EmailServiceError(f"Device flow failed: {flow}")

        print("\n" + "=" * 50)
        print("  GRAPH MAIL.SEND SIGN-IN REQUIRED")
        print("=" * 50)
        print(f"  1. Open: {flow['verification_uri']}")
        print(f"  2. Enter code: {flow['user_code']}")
        print("=" * 50 + "\n")

        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise EmailServiceError(
                result.get("error_description")
                or result.get("error")
                or "Token acquisition failed"
            )
        return result["access_token"]

    def send_mail(self, to_email, subject, html_body,
                  cc_emails=None, save_to_sent=True):
        """Send a single email through Microsoft Graph.

        Returns a dict with keys: conversation_id, message_id, sent_at_iso.
        Either may be None if the lookup of the saved sent item fails — that's
        non-fatal for Option C (we don't need them yet) but useful for Option A
        when it's added later.
        """
        token = self._get_token()
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [
                {"emailAddress": {"address": to_email}}
            ],
        }
        if cc_emails:
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_emails if addr
            ]

        payload = {"message": message, "saveToSentItems": save_to_sent}
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        resp = requests.post(
            GRAPH_SENDMAIL_URL, json=payload, headers=headers, timeout=30
        )
        if resp.status_code not in (200, 202):
            logger.error("Graph sendMail failed: %s %s",
                         resp.status_code, resp.text)
            raise EmailServiceError(
                f"Graph sendMail failed ({resp.status_code}): {resp.text}"
            )

        conversation_id = None
        message_id = None
        if save_to_sent:
            try:
                conversation_id, message_id = self._lookup_sent_message(
                    token, subject, to_email
                )
            except Exception:
                logger.exception(
                    "Could not retrieve conversationId for sent mail "
                    "(non-fatal)"
                )

        logger.info("Mail sent to %s — subject=%r conv=%s",
                    to_email, subject, conversation_id)
        return {
            "conversation_id": conversation_id,
            "message_id": message_id,
        }

    def _lookup_sent_message(self, token, subject, to_email):
        """Look up the most recent sent message matching subject + recipient
        to capture conversationId / messageId for future reply tracking."""
        headers = {"Authorization": f"Bearer {token}"}
        # Filter is best-effort: subject + recipient + sort by sent date desc
        params = {
            "$top": 5,
            "$orderby": "sentDateTime desc",
            "$select": "id,conversationId,subject,toRecipients,sentDateTime",
        }
        resp = requests.get(
            GRAPH_SENT_ITEMS_URL, headers=headers, params=params, timeout=15
        )
        resp.raise_for_status()
        items = resp.json().get("value", [])
        target = (to_email or "").lower()
        for item in items:
            if item.get("subject") != subject:
                continue
            recipients = [
                r.get("emailAddress", {}).get("address", "").lower()
                for r in item.get("toRecipients", [])
            ]
            if target in recipients:
                return item.get("conversationId"), item.get("id")
        return None, None

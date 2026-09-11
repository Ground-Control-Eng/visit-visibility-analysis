"""Last-resort fallback alert via a Microsoft Teams channel webhook - used only when Microsoft
Graph can't send the failure-alert email either (see send_summary.send_failure_email).
Deliberately independent of Graph/Outlook: a plain HTTPS POST via the stdlib, so it still works
when the whole reason we're here is that Graph itself is unreachable or misconfigured.
"""
from __future__ import annotations

import json
import logging
import urllib.request

from .config import Config

logger = logging.getLogger("visit_reconciliation")


def send_teams_alert(subject: str, message: str, cfg: Config) -> None:
    """Posts subject/message to cfg.teams.webhook_url as an Adaptive Card. Best-effort and never
    raises - a missing/unset webhook URL logs a warning and returns; a request failure is logged
    and swallowed, since there's no further fallback channel after this one."""
    webhook_url = cfg.teams.webhook_url
    if not webhook_url:
        logger.warning(
            "No notifications.teams.webhook_url configured in config.yaml - cannot post a Teams "
            "fallback alert for %r. See README for how to set one up.", subject,
        )
        return

    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.4",
                    "body": [
                        {"type": "TextBlock", "text": subject, "weight": "Bolder", "size": "Medium", "wrap": True},
                        {"type": "TextBlock", "text": message, "wrap": True},
                    ],
                },
            }
        ],
    }
    request = urllib.request.Request(
        webhook_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            if response.status >= 300:
                logger.warning("Teams webhook responded with HTTP %s for %r.", response.status, subject)
            else:
                logger.info("Posted Teams fallback alert for %r.", subject)
    except Exception:  # noqa: BLE001
        logger.exception("Could not post Teams fallback alert for %r.", subject)

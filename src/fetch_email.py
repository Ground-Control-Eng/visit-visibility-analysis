"""Finds today's Hubscape trigger email via Microsoft Graph and saves its attachment."""
from __future__ import annotations

import base64
import fnmatch
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path

from . import graph_client
from .config import Config

logger = logging.getLogger("visit_reconciliation")


class EmailNotFoundError(Exception):
    pass


class AttachmentNotFoundError(Exception):
    pass


def _resolve_folder_id(cfg: Config, folder_path: str) -> str:
    """Maps a "Inbox" / "Inbox/Subfolder" style path to a Graph mail folder id, walking
    childFolders by displayName one segment at a time. "Inbox" (or blank) resolves to the
    well-known folder name "inbox" directly - no lookup call needed."""
    if folder_path.lower() in ("inbox", ""):
        return "inbox"

    parts = [p for p in folder_path.split("/") if p.lower() != "inbox"]
    current_id = "inbox"
    for part in parts:
        escaped = part.replace("'", "''")
        resp = graph_client.graph_get(
            f"/users/{cfg.graph.mailbox}/mailFolders/{current_id}/childFolders",
            cfg,
            params={"$filter": f"displayName eq '{escaped}'", "$select": "id,displayName"},
        )
        matches = resp.get("value", [])
        if len(matches) != 1:
            raise EmailNotFoundError(
                f"Could not resolve folder path {folder_path!r}: expected exactly one child "
                f"folder named {part!r} under folder id {current_id!r}, found {len(matches)}."
            )
        current_id = matches[0]["id"]
    return current_id


def find_todays_hubscape_email(cfg: Config, now: datetime | None = None) -> dict:
    """Returns the most recent matching Graph message resource (a dict), or raises
    EmailNotFoundError. A naive `now` is assumed to already be UTC."""
    now = now or datetime.now(timezone.utc)
    if now.tzinfo is None:
        now = now.replace(tzinfo=timezone.utc)

    trigger = cfg.email.trigger
    cutoff = now - timedelta(hours=trigger.lookback_hours)
    cutoff_iso = cutoff.strftime("%Y-%m-%dT%H:%M:%SZ")

    folder_id = _resolve_folder_id(cfg, trigger.search_folder)
    params = {
        "$filter": f"receivedDateTime ge {cutoff_iso}",
        "$orderby": "receivedDateTime desc",
        "$top": cfg.graph.mail_search_page_size,
        "$select": "id,receivedDateTime,subject,from,hasAttachments",
    }

    matches = []
    path: str | None = f"/users/{cfg.graph.mailbox}/mailFolders/{folder_id}/messages"
    page = 0
    stop = False
    while path and not stop and page < cfg.graph.mail_search_max_pages:
        page += 1
        resp = graph_client.graph_get(path, cfg, params=params if page == 1 else None)

        for msg in resp.get("value", []):
            received = datetime.strptime(msg["receivedDateTime"], "%Y-%m-%dT%H:%M:%SZ").replace(
                tzinfo=timezone.utc
            )
            if received < cutoff:
                stop = True  # sorted descending, nothing older (here or on later pages) will match
                break

            sender_ok = True
            if trigger.sender_filter:
                sender_address = ((msg.get("from") or {}).get("emailAddress") or {}).get("address") or ""
                sender_ok = trigger.sender_filter.lower() in sender_address.lower()

            subject_ok = True
            if trigger.subject_contains:
                subject_ok = trigger.subject_contains.lower() in (msg.get("subject") or "").lower()

            if sender_ok and subject_ok:
                matches.append(msg)

        path = resp.get("@odata.nextLink")

    if not matches:
        raise EmailNotFoundError(
            f"No email found in folder '{trigger.search_folder}' from the last "
            f"{trigger.lookback_hours}h matching sender_filter={trigger.sender_filter!r} "
            f"subject_contains={trigger.subject_contains!r}."
        )

    if len(matches) > 1:
        logger.warning(
            "Found %d candidate emails matching the trigger filter; using the most recent (received %s).",
            len(matches), matches[0]["receivedDateTime"],
        )

    return matches[0]


def extract_attachment(message: dict, cfg: Config, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    pattern = cfg.email.trigger.attachment_name_pattern

    if not message.get("hasAttachments"):
        raise AttachmentNotFoundError(
            f"No attachment matching pattern {pattern!r} found. Message has no attachments."
        )

    resp = graph_client.graph_get(
        f"/users/{cfg.graph.mailbox}/messages/{message['id']}/attachments", cfg,
    )
    all_attachments = resp.get("value", [])
    file_attachments = [
        a for a in all_attachments if a.get("@odata.type") == "#microsoft.graph.fileAttachment"
    ]
    candidates = [a for a in file_attachments if fnmatch.fnmatch(a["name"], pattern)]
    if not candidates:
        all_names = [a.get("name", "<unnamed>") for a in all_attachments]
        raise AttachmentNotFoundError(
            f"No attachment matching pattern {pattern!r} found. Attachments present: {all_names}"
        )

    attachment = candidates[0]
    dest_path = dest_dir / f"raw_attachment_{attachment['name']}"
    dest_path.write_bytes(base64.b64decode(attachment["contentBytes"]))
    logger.info("Saved Hubscape attachment to %s", dest_path)
    return dest_path

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


def _parse_graph_datetime(value: str) -> datetime:
    """Parses a Graph DateTimeOffset string (always UTC, "Z"-suffixed) into an aware UTC
    datetime. Graph commonly includes fractional seconds of varying precision (e.g.
    "2026-09-10T10:00:00.1234567Z"), which a fixed-format strptime("%Y-%m-%dT%H:%M:%SZ") can't
    parse - normalize the fractional part to exactly 6 digits (Python's %f) before parsing."""
    value = value.rstrip("Z")
    if "." in value:
        base, frac = value.split(".", 1)
        frac = (frac + "000000")[:6]
        parsed = datetime.strptime(f"{base}.{frac}", "%Y-%m-%dT%H:%M:%S.%f")
    else:
        parsed = datetime.strptime(value, "%Y-%m-%dT%H:%M:%S")
    return parsed.replace(tzinfo=timezone.utc)


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
    path = f"/users/{cfg.graph.mailbox}/mailFolders/{folder_id}/messages"

    matches = []
    stop = False
    truncated = False
    pages_seen = 0
    for resp in graph_client.graph_get_pages(path, cfg, params=params):
        pages_seen += 1
        for msg in resp.get("value", []):
            received = _parse_graph_datetime(msg["receivedDateTime"])
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

        if stop:
            break
        if pages_seen >= cfg.graph.mail_search_max_pages:
            if resp.get("@odata.nextLink"):
                # More results existed beyond graph.mail_search_max_pages and we never reached
                # the lookback cutoff - this search was bounded, not exhausted. Surface that
                # distinction rather than reporting it identically to a genuine no-match.
                truncated = True
            break

    if not matches:
        if truncated:
            logger.warning(
                "Trigger-email search hit graph.mail_search_max_pages (%d) before reaching the "
                "%dh lookback cutoff in folder '%s' - some messages were not checked.",
                cfg.graph.mail_search_max_pages, trigger.lookback_hours, trigger.search_folder,
            )
            raise EmailNotFoundError(
                f"No email found in folder '{trigger.search_folder}' matching "
                f"sender_filter={trigger.sender_filter!r} subject_contains={trigger.subject_contains!r} "
                f"- but the search was truncated after {pages_seen} page(s) "
                f"({pages_seen * cfg.graph.mail_search_page_size} messages checked) before "
                f"reaching the {trigger.lookback_hours}h lookback cutoff. Raise "
                "graph.mail_search_max_pages/mail_search_page_size and retry before concluding "
                "no matching email exists."
            )
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

    all_attachments = []
    for resp in graph_client.graph_get_pages(
        f"/users/{cfg.graph.mailbox}/messages/{message['id']}/attachments", cfg,
    ):
        all_attachments.extend(resp.get("value", []))

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

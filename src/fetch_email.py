"""Finds today's Hubscape trigger email in Outlook (via COM) and saves its attachment."""
from __future__ import annotations

import fnmatch
import logging
from datetime import datetime, timedelta
from pathlib import Path

from .config import Config

logger = logging.getLogger("visit_reconciliation")


class EmailNotFoundError(Exception):
    pass


class AttachmentNotFoundError(Exception):
    pass


def _get_folder(namespace, folder_name: str):
    inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
    if folder_name.lower() in ("inbox", ""):
        return inbox
    # Support "Inbox/Subfolder" style paths.
    parts = [p for p in folder_name.split("/") if p.lower() != "inbox"]
    folder = inbox
    for part in parts:
        folder = folder.Folders[part]
    return folder


def find_todays_hubscape_email(cfg: Config, now: datetime | None = None):
    """Returns the most recent matching Outlook MailItem, or raises EmailNotFoundError.

    Imports win32com lazily so this module can be imported (e.g. by tests) on machines
    without Outlook/pywin32 installed.
    """
    import win32com.client  # noqa: PLC0415

    now = now or datetime.now()
    trigger = cfg.email.trigger
    cutoff = now - timedelta(hours=trigger.lookback_hours)

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if cfg.email.outlook_profile:
        namespace.Logon(cfg.email.outlook_profile)

    folder = _get_folder(namespace, trigger.search_folder)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # descending, newest first

    matches = []
    for item in items:
        try:
            received = item.ReceivedTime
        except Exception:  # noqa: BLE001
            continue
        # pywin32 COM datetimes compare fine against naive datetimes via .replace(tzinfo=None)
        received_naive = received.replace(tzinfo=None) if received.tzinfo else received
        if received_naive < cutoff:
            break  # sorted descending, nothing older will match either

        sender_ok = True
        if trigger.sender_filter:
            sender_address = getattr(item, "SenderEmailAddress", "") or ""
            sender_ok = trigger.sender_filter.lower() in sender_address.lower()

        subject_ok = True
        if trigger.subject_contains:
            subject_ok = trigger.subject_contains.lower() in (item.Subject or "").lower()

        if sender_ok and subject_ok:
            matches.append(item)

    if not matches:
        raise EmailNotFoundError(
            f"No email found in folder '{trigger.search_folder}' from the last "
            f"{trigger.lookback_hours}h matching sender_filter={trigger.sender_filter!r} "
            f"subject_contains={trigger.subject_contains!r}."
        )

    if len(matches) > 1:
        logger.warning(
            "Found %d candidate emails matching the trigger filter; using the most recent (received %s).",
            len(matches), matches[0].ReceivedTime,
        )

    return matches[0]


def extract_attachment(mail_item, cfg: Config, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    pattern = cfg.email.trigger.attachment_name_pattern

    candidates = [a for a in mail_item.Attachments if fnmatch.fnmatch(a.FileName, pattern)]
    if not candidates:
        all_names = [a.FileName for a in mail_item.Attachments]
        raise AttachmentNotFoundError(
            f"No attachment matching pattern {pattern!r} found. Attachments present: {all_names}"
        )

    attachment = candidates[0]
    dest_path = dest_dir / f"raw_attachment_{attachment.FileName}"
    attachment.SaveAsFile(str(dest_path))
    logger.info("Saved Hubscape attachment to %s", dest_path)
    return dest_path

import base64
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import fetch_email


def _fake_cfg(sender_filter="ci@hubscape.co.uk", subject_contains="", search_folder="Inbox",
              lookback_hours=24, attachment_name_pattern="production_uncompleted_visits_*.csv",
              mailbox="alex.clark@ground-control.co.uk"):
    return SimpleNamespace(
        graph=SimpleNamespace(mailbox=mailbox, mail_search_page_size=50, mail_search_max_pages=5),
        email=SimpleNamespace(
            trigger=SimpleNamespace(
                sender_filter=sender_filter, subject_contains=subject_contains,
                search_folder=search_folder, lookback_hours=lookback_hours,
                attachment_name_pattern=attachment_name_pattern,
            )
        ),
    )


def _msg(msg_id, received_at, sender="ci@hubscape.co.uk", subject="trigger", has_attachments=True):
    return {
        "id": msg_id,
        "receivedDateTime": received_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "subject": subject,
        "from": {"emailAddress": {"address": sender}},
        "hasAttachments": has_attachments,
    }


def test_finds_most_recent_matching_email_by_sender_and_subject(monkeypatch):
    cfg = _fake_cfg()
    now = datetime(2026, 9, 10, 12, 0, 0, tzinfo=timezone.utc)
    messages = [
        _msg("newest", now - timedelta(hours=1)),
        _msg("older-match", now - timedelta(hours=2)),
        _msg("wrong-sender", now - timedelta(hours=1, minutes=30), sender="someone-else@example.com"),
    ]
    monkeypatch.setattr(fetch_email.graph_client, "graph_get", lambda path, cfg, params=None: {"value": messages})

    result = fetch_email.find_todays_hubscape_email(cfg, now=now)
    assert result["id"] == "newest"


def test_lookback_window_boundary_excludes_older_messages(monkeypatch):
    cfg = _fake_cfg()
    now = datetime(2026, 9, 10, 12, 0, 0, tzinfo=timezone.utc)
    cutoff = now - timedelta(hours=cfg.email.trigger.lookback_hours)

    old_message = _msg("too-old", cutoff - timedelta(minutes=1))
    monkeypatch.setattr(fetch_email.graph_client, "graph_get", lambda path, cfg, params=None: {"value": [old_message]})
    with pytest.raises(fetch_email.EmailNotFoundError):
        fetch_email.find_todays_hubscape_email(cfg, now=now)

    boundary_message = _msg("boundary", cutoff)
    monkeypatch.setattr(fetch_email.graph_client, "graph_get", lambda path, cfg, params=None: {"value": [boundary_message]})
    result = fetch_email.find_todays_hubscape_email(cfg, now=now)
    assert result["id"] == "boundary"


def test_raises_email_not_found_error_with_diagnostic_message(monkeypatch):
    cfg = _fake_cfg()
    monkeypatch.setattr(fetch_email.graph_client, "graph_get", lambda path, cfg, params=None: {"value": []})

    with pytest.raises(fetch_email.EmailNotFoundError) as exc_info:
        fetch_email.find_todays_hubscape_email(cfg, now=datetime(2026, 9, 10, 12, 0, 0, tzinfo=timezone.utc))

    message = str(exc_info.value)
    assert "Inbox" in message
    assert "ci@hubscape.co.uk" in message
    assert "24" in message


def test_nested_folder_path_resolves_child_folder_by_display_name(monkeypatch):
    cfg = _fake_cfg(search_folder="Inbox/Hubscape")
    calls = []

    def fake_graph_get(path, cfg, params=None):
        calls.append((path, params))
        if path.endswith("/childFolders"):
            assert params["$filter"] == "displayName eq 'Hubscape'"
            return {"value": [{"id": "child-folder-id", "displayName": "Hubscape"}]}
        return {"value": [_msg("found", datetime(2026, 9, 10, 10, 0, 0, tzinfo=timezone.utc))]}

    monkeypatch.setattr(fetch_email.graph_client, "graph_get", fake_graph_get)

    result = fetch_email.find_todays_hubscape_email(cfg, now=datetime(2026, 9, 10, 12, 0, 0, tzinfo=timezone.utc))
    assert result["id"] == "found"
    assert calls[-1][0] == "/users/alex.clark@ground-control.co.uk/mailFolders/child-folder-id/messages"


def test_extract_attachment_matches_pattern_and_decodes_base64_content(monkeypatch, tmp_path):
    cfg = _fake_cfg()
    content = b"VisitID,Status\n1,OK\n"
    encoded = base64.b64encode(content).decode("ascii")
    attachments_response = {
        "value": [
            {"@odata.type": "#microsoft.graph.itemAttachment", "name": "not-a-file"},
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": "production_uncompleted_visits_20260910.csv",
                "contentBytes": encoded,
            },
        ]
    }
    monkeypatch.setattr(fetch_email.graph_client, "graph_get", lambda path, cfg: attachments_response)

    message = {"id": "msg-1", "hasAttachments": True}
    dest_path = fetch_email.extract_attachment(message, cfg, tmp_path)

    assert dest_path.name == "raw_attachment_production_uncompleted_visits_20260910.csv"
    assert dest_path.read_bytes() == content


def test_extract_attachment_raises_with_full_attachment_name_list_on_no_match(monkeypatch, tmp_path):
    cfg = _fake_cfg()
    attachments_response = {
        "value": [
            {"@odata.type": "#microsoft.graph.fileAttachment", "name": "unrelated.txt", "contentBytes": ""},
        ]
    }
    monkeypatch.setattr(fetch_email.graph_client, "graph_get", lambda path, cfg: attachments_response)

    message = {"id": "msg-1", "hasAttachments": True}
    with pytest.raises(fetch_email.AttachmentNotFoundError) as exc_info:
        fetch_email.extract_attachment(message, cfg, tmp_path)

    assert "unrelated.txt" in str(exc_info.value)


def test_extract_attachment_skips_network_call_when_has_attachments_false(monkeypatch, tmp_path):
    cfg = _fake_cfg()

    def _fail_if_called(*_args, **_kwargs):
        raise AssertionError("graph_get should not be called when hasAttachments is False")

    monkeypatch.setattr(fetch_email.graph_client, "graph_get", _fail_if_called)

    message = {"id": "msg-1", "hasAttachments": False}
    with pytest.raises(fetch_email.AttachmentNotFoundError):
        fetch_email.extract_attachment(message, cfg, tmp_path)

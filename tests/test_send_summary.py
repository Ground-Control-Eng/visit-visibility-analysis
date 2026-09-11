import base64
import sys
from datetime import date
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import send_summary
from src.graph_client import GraphApiError
from src.send_summary import (
    _alert_via_event_log,
    _build_attachment_payloads,
    _prioritized_attachments,
    _send_via_graph,
    send_failure_email,
    send_success_email,
)


def test_alert_via_event_log_reports_event(monkeypatch):
    import win32evtlogutil

    calls = []
    monkeypatch.setattr(win32evtlogutil, "ReportEvent", lambda *args, **kwargs: calls.append((args, kwargs)))
    _alert_via_event_log("something went wrong")
    assert len(calls) == 1
    assert calls[0][1]["strings"] == ["something went wrong"]


def test_alert_via_event_log_swallows_reporting_failure(monkeypatch):
    import win32evtlogutil

    def _raise(*_args, **_kwargs):
        raise OSError("event source 'Visit Reconciliation' is not registered")

    monkeypatch.setattr(win32evtlogutil, "ReportEvent", _raise)
    _alert_via_event_log("something went wrong")  # must not raise


def _make_cfg(**overrides):
    defaults = dict(
        email=SimpleNamespace(to=["a@b.com"], cc=[]),
        graph=SimpleNamespace(mailbox="alex.clark@ground-control.co.uk", max_inline_attachment_bytes=3_000_000),
    )
    defaults.update(overrides)
    return SimpleNamespace(**defaults)


def test_prioritized_attachments_keeps_all_when_under_limit(tmp_path):
    summary = tmp_path / "summary.csv"
    detail = tmp_path / "detail.csv"
    summary.write_bytes(b"a" * 100)
    detail.write_bytes(b"b" * 100)

    result = _prioritized_attachments([summary, detail], max_bytes=1_000)
    assert result == [summary, detail]


def test_prioritized_attachments_drops_lowest_priority_sankey_when_over_limit(tmp_path):
    summary = tmp_path / "summary.csv"
    sankey = tmp_path / "sankey.html"
    summary.write_bytes(b"a" * 100)
    sankey.write_bytes(b"b" * 5_000)

    result = _prioritized_attachments([summary, sankey], max_bytes=1_000)
    assert result == [summary]


def test_prioritized_attachments_keeps_highest_priority_even_if_alone_it_exceeds_limit(tmp_path):
    summary = tmp_path / "summary.csv"
    summary.write_bytes(b"a" * 5_000)

    result = _prioritized_attachments([summary], max_bytes=1_000)
    assert result == [summary]


def test_prioritized_attachments_skips_missing_files(tmp_path):
    summary = tmp_path / "summary.csv"
    summary.write_bytes(b"a" * 10)
    missing = tmp_path / "missing.csv"

    result = _prioritized_attachments([summary, missing], max_bytes=1_000)
    assert result == [summary]


def test_build_attachment_payloads_encodes_base64_file_attachment(tmp_path):
    csv_path = tmp_path / "summary.csv"
    csv_path.write_bytes(b"metric,count\nfoo,1\n")

    payloads = _build_attachment_payloads([csv_path])
    assert len(payloads) == 1
    payload = payloads[0]
    assert payload["@odata.type"] == "#microsoft.graph.fileAttachment"
    assert payload["name"] == "summary.csv"
    assert base64.b64decode(payload["contentBytes"]) == csv_path.read_bytes()


def test_send_via_graph_builds_expected_sendmail_payload(monkeypatch, tmp_path):
    csv_path = tmp_path / "summary.csv"
    csv_path.write_bytes(b"metric,count\nfoo,1\n")
    cfg = _make_cfg(email=SimpleNamespace(to=["a@b.com", "c@d.com"], cc=["e@f.com"]))

    calls = []
    monkeypatch.setattr(send_summary.graph_client, "graph_post", lambda path, cfg, json_body: calls.append((path, json_body)))

    _send_via_graph("Subject line", "<html></html>", cfg, [csv_path])

    assert len(calls) == 1
    path, body = calls[0]
    assert path == "/users/alex.clark@ground-control.co.uk/sendMail"
    message = body["message"]
    assert message["subject"] == "Subject line"
    assert message["body"] == {"contentType": "HTML", "content": "<html></html>"}
    assert message["toRecipients"] == [{"emailAddress": {"address": "a@b.com"}}, {"emailAddress": {"address": "c@d.com"}}]
    assert message["ccRecipients"] == [{"emailAddress": {"address": "e@f.com"}}]
    assert len(message["attachments"]) == 1
    assert body["saveToSentItems"] is True


def test_send_via_graph_omits_cc_key_when_no_cc_recipients(monkeypatch, tmp_path):
    cfg = _make_cfg(email=SimpleNamespace(to=["a@b.com"], cc=[]))
    calls = []
    monkeypatch.setattr(send_summary.graph_client, "graph_post", lambda path, cfg, json_body: calls.append(json_body))

    _send_via_graph("Subject", "<html></html>", cfg, [])

    assert "ccRecipients" not in calls[0]["message"]


def test_send_via_graph_alerts_event_log_and_raises_on_graph_api_error(monkeypatch, tmp_path):
    cfg = _make_cfg()
    event_log_calls = []
    monkeypatch.setattr(send_summary, "_alert_via_event_log", lambda msg: event_log_calls.append(msg))

    def _raise(*_args, **_kwargs):
        raise GraphApiError(403, "ErrorAccessDenied", "The mailbox is not accessible.")

    monkeypatch.setattr(send_summary.graph_client, "graph_post", _raise)

    with pytest.raises(GraphApiError):
        _send_via_graph("Subject", "<html></html>", cfg, [])

    assert len(event_log_calls) == 1
    assert "Subject" in event_log_calls[0]


def test_send_success_email_test_mode_writes_local_html_without_calling_graph(monkeypatch, tmp_path):
    cfg = SimpleNamespace(run=SimpleNamespace(test_mode=True), email=SimpleNamespace(send_on_success=True))
    monkeypatch.setattr(
        send_summary, "_send_via_graph",
        lambda *a, **k: (_ for _ in ()).throw(AssertionError("Graph should not be called in test_mode")),
    )

    summary_df = pd.DataFrame([{"metric": "TOTAL_ICE2_ROWS", "count": 0}])
    detail_df = pd.DataFrame(columns=["issue_type", "VisitID"])
    send_success_email(summary_df, detail_df, cfg, date(2026, 9, 5), tmp_path)

    assert (tmp_path / "would_be_email.html").exists()


def test_send_success_email_send_on_success_false_writes_local_html_without_calling_graph(monkeypatch, tmp_path):
    cfg = SimpleNamespace(run=SimpleNamespace(test_mode=False), email=SimpleNamespace(send_on_success=False))
    monkeypatch.setattr(
        send_summary, "_send_via_graph",
        lambda *a, **k: (_ for _ in ()).throw(AssertionError("Graph should not be called")),
    )

    summary_df = pd.DataFrame([{"metric": "TOTAL_ICE2_ROWS", "count": 0}])
    detail_df = pd.DataFrame(columns=["issue_type", "VisitID"])
    send_success_email(summary_df, detail_df, cfg, date(2026, 9, 5), tmp_path)

    assert (tmp_path / "would_be_email.html").exists()


def test_send_failure_email_falls_back_to_teams_when_graph_send_fails(monkeypatch, tmp_path):
    cfg = SimpleNamespace(run=SimpleNamespace(test_mode=False), email=SimpleNamespace(send_on_failure=True))
    monkeypatch.setattr(
        send_summary, "_send_via_graph",
        lambda *a, **k: (_ for _ in ()).throw(GraphApiError(500, "InternalServerError", "boom")),
    )
    teams_calls = []
    monkeypatch.setattr(send_summary.notify_teams, "send_teams_alert", lambda *a, **k: teams_calls.append(a))

    send_failure_email("query_ice2", RuntimeError("db unreachable"), cfg, date(2026, 9, 5), tmp_path)

    assert len(teams_calls) == 1

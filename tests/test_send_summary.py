import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import send_summary
from src.send_summary import (
    EMAIL_SUBJECT_PREFIX,
    _alert_via_event_log,
    _confirm_outbox_transmit,
    _exchange_connection_problem,
    _inbox_last_received,
    _outbox_entry_ids,
    is_stale_pipeline_outbox_item,
)

SUBJECT = f"{EMAIL_SUBJECT_PREFIX} Summary - 2026-09-05"
THRESHOLD = 3600


def test_stale_pipeline_item_is_flagged():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(hours=2)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is True


def test_fresh_pipeline_item_is_not_flagged():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(seconds=5)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is False


def test_non_pipeline_subject_is_never_flagged_regardless_of_age():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(days=30)
    assert is_stale_pipeline_outbox_item("Some unrelated email", created_at, now, THRESHOLD) is False


def test_boundary_at_exact_threshold():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(seconds=THRESHOLD)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is True


def test_tz_aware_inputs_both_sides_consistent():
    now = datetime(2026, 9, 7, 12, 0, 0, tzinfo=timezone.utc)
    created_at = now - timedelta(hours=2)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is True


def test_naive_vs_aware_mismatch_raises_typeerror():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = datetime(2026, 9, 7, 10, 0, 0, tzinfo=timezone.utc)
    with pytest.raises(TypeError):
        is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD)


class _StubItem:
    def __init__(self, entry_id):
        self._entry_id = entry_id

    @property
    def EntryID(self):
        if self._entry_id is None:
            raise Exception("Outlook has already begun transmitting this message.")
        return self._entry_id


class _StubOutbox:
    def __init__(self, items):
        self.Items = items


def test_outbox_entry_ids_skips_item_mid_transmission():
    outbox = _StubOutbox([_StubItem("readable-1"), _StubItem(None), _StubItem("readable-2")])
    assert _outbox_entry_ids(outbox) == {"readable-1", "readable-2"}


@pytest.mark.parametrize("mode", [0, 100, 200, 300, 400])
def test_disconnected_exchange_modes_are_flagged(mode):
    assert _exchange_connection_problem(mode) is not None


@pytest.mark.parametrize("mode", [500, 600, 700, 800])
def test_connected_looking_exchange_modes_are_not_flagged(mode):
    assert _exchange_connection_problem(mode) is None


class _StubItems:
    def __init__(self, items):
        self._items = items

    def Sort(self, *_args, **_kwargs):
        pass

    def GetFirst(self):
        return self._items[0] if self._items else None


class _StubInboxFolder:
    def __init__(self, items):
        self.Items = _StubItems(items)


class _StubReceivedItem:
    def __init__(self, received_time):
        self.ReceivedTime = received_time


class _StubNamespaceForInbox:
    def __init__(self, folder=None, raise_on_get_folder=False):
        self._folder = folder
        self._raise = raise_on_get_folder

    def GetDefaultFolder(self, _folder_id):
        if self._raise:
            raise Exception("Could not access Inbox")
        return self._folder


def test_inbox_last_received_empty_inbox_returns_none():
    namespace = _StubNamespaceForInbox(folder=_StubInboxFolder([]))
    assert _inbox_last_received(namespace) is None


def test_inbox_last_received_naive_time_passed_through():
    received = datetime(2026, 9, 9, 9, 24, 18, 182000)
    namespace = _StubNamespaceForInbox(folder=_StubInboxFolder([_StubReceivedItem(received)]))
    assert _inbox_last_received(namespace) == received


def test_inbox_last_received_tz_aware_time_has_tzinfo_stripped():
    received = datetime(2026, 9, 9, 9, 24, 18, tzinfo=timezone.utc)
    namespace = _StubNamespaceForInbox(folder=_StubInboxFolder([_StubReceivedItem(received)]))
    result = _inbox_last_received(namespace)
    assert result == received.replace(tzinfo=None)
    assert result.tzinfo is None


def test_inbox_last_received_folder_access_failure_returns_none():
    namespace = _StubNamespaceForInbox(raise_on_get_folder=True)
    assert _inbox_last_received(namespace) is None


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


class _FakeClock:
    """Lets the retry loop's timeout logic run deterministically with no real waiting."""

    def __init__(self):
        self.now = 0.0

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


class _AttemptTrackingNamespace:
    def __init__(self):
        self.send_and_receive_calls = 0

    def SendAndReceive(self, _sync):
        self.send_and_receive_calls += 1


class _DynamicStubOutbox:
    """Outbox item disappears (transmits) once send_and_receive_calls reaches leaves_on_attempt."""

    def __init__(self, namespace, entry_id, leaves_on_attempt):
        self._namespace = namespace
        self._entry_id = entry_id
        self._leaves_on_attempt = leaves_on_attempt

    @property
    def Items(self):
        if self._namespace.send_and_receive_calls >= self._leaves_on_attempt:
            return []
        return [_StubItem(self._entry_id)]


def _make_cfg(timeout_seconds, retries):
    return SimpleNamespace(
        email=SimpleNamespace(
            send_confirm_timeout_seconds=timeout_seconds, send_confirm_retries=retries, to=["a@b.com"],
        )
    )


def test_confirm_outbox_transmit_succeeds_without_needing_a_retry(monkeypatch):
    clock = _FakeClock()
    monkeypatch.setattr(send_summary.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(send_summary.time, "sleep", clock.sleep)
    namespace = _AttemptTrackingNamespace()
    outbox = _DynamicStubOutbox(namespace, "id-1", leaves_on_attempt=1)

    assert _confirm_outbox_transmit(namespace, outbox, "subj", {"id-1"}, _make_cfg(5, 2)) is True
    assert namespace.send_and_receive_calls == 1


def test_confirm_outbox_transmit_succeeds_on_a_later_retry(monkeypatch):
    clock = _FakeClock()
    monkeypatch.setattr(send_summary.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(send_summary.time, "sleep", clock.sleep)
    namespace = _AttemptTrackingNamespace()
    outbox = _DynamicStubOutbox(namespace, "id-1", leaves_on_attempt=2)

    assert _confirm_outbox_transmit(namespace, outbox, "subj", {"id-1"}, _make_cfg(5, 2)) is True
    assert namespace.send_and_receive_calls == 2


def test_confirm_outbox_transmit_gives_up_after_retries_exhausted(monkeypatch):
    clock = _FakeClock()
    monkeypatch.setattr(send_summary.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(send_summary.time, "sleep", clock.sleep)
    namespace = _AttemptTrackingNamespace()
    outbox = _DynamicStubOutbox(namespace, "id-1", leaves_on_attempt=99)

    assert _confirm_outbox_transmit(namespace, outbox, "subj", {"id-1"}, _make_cfg(5, 2)) is False
    assert namespace.send_and_receive_calls == 3  # 1 initial attempt + 2 retries

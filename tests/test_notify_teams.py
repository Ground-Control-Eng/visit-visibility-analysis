import json
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import notify_teams

SUBJECT = "Visit Reconciliation FAILED (send_summary) - 2026-09-09"
MESSAGE = "Stage: send_summary\nError: RuntimeError: still in Outbox"


def _make_cfg(webhook_url):
    return SimpleNamespace(teams=SimpleNamespace(webhook_url=webhook_url))


class _FakeResponse:
    def __init__(self, status):
        self.status = status

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


def test_no_webhook_configured_does_not_attempt_a_request(monkeypatch):
    calls = []
    monkeypatch.setattr(notify_teams.urllib.request, "urlopen", lambda *a, **k: calls.append((a, k)))

    notify_teams.send_teams_alert(SUBJECT, MESSAGE, _make_cfg(None))

    assert calls == []


def test_posts_expected_payload_when_webhook_configured(monkeypatch):
    captured = {}

    def fake_urlopen(request, timeout=None):
        captured["url"] = request.full_url
        captured["headers"] = request.headers
        captured["body"] = json.loads(request.data.decode("utf-8"))
        captured["timeout"] = timeout
        return _FakeResponse(200)

    monkeypatch.setattr(notify_teams.urllib.request, "urlopen", fake_urlopen)

    notify_teams.send_teams_alert(SUBJECT, MESSAGE, _make_cfg("https://example.com/webhook"))

    assert captured["url"] == "https://example.com/webhook"
    assert captured["headers"]["Content-type"] == "application/json"
    card = captured["body"]["attachments"][0]["content"]
    text_blocks = [block["text"] for block in card["body"]]
    assert SUBJECT in text_blocks
    assert MESSAGE in text_blocks


def test_non_success_response_is_logged_not_raised(monkeypatch):
    monkeypatch.setattr(notify_teams.urllib.request, "urlopen", lambda *a, **k: _FakeResponse(500))

    notify_teams.send_teams_alert(SUBJECT, MESSAGE, _make_cfg("https://example.com/webhook"))  # must not raise


def test_request_exception_is_swallowed(monkeypatch):
    def _raise(*_args, **_kwargs):
        raise OSError("network unreachable")

    monkeypatch.setattr(notify_teams.urllib.request, "urlopen", _raise)

    notify_teams.send_teams_alert(SUBJECT, MESSAGE, _make_cfg("https://example.com/webhook"))  # must not raise

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import graph_client


class _FakeResponse:
    def __init__(self, status_code, json_data=None, text="", headers=None):
        self.status_code = status_code
        self._json_data = json_data
        self.text = text
        self.headers = headers or {}

    @property
    def ok(self):
        return 200 <= self.status_code < 300

    def json(self):
        if self._json_data is None:
            raise ValueError("response has no JSON body")
        return self._json_data


def _fake_cfg(max_retries=2):
    return SimpleNamespace(
        graph=SimpleNamespace(
            tenant_id="tenant-1", client_id="client-1", client_secret="secret-1",
            request_timeout_seconds=5, max_retries=max_retries,
        )
    )


@pytest.fixture(autouse=True)
def _reset_msal_singleton():
    graph_client._msal_app = None
    yield
    graph_client._msal_app = None


def test_acquire_token_returns_access_token_on_success(monkeypatch):
    cfg = _fake_cfg()
    fake_app = SimpleNamespace(acquire_token_for_client=lambda scopes: {"access_token": "tok-123"})
    monkeypatch.setattr(graph_client, "_get_msal_app", lambda cfg: fake_app)

    assert graph_client.acquire_token(cfg) == "tok-123"


def test_acquire_token_raises_graph_auth_error_on_failure(monkeypatch):
    cfg = _fake_cfg()
    fake_app = SimpleNamespace(
        acquire_token_for_client=lambda scopes: {"error": "invalid_client", "error_description": "bad secret"}
    )
    monkeypatch.setattr(graph_client, "_get_msal_app", lambda cfg: fake_app)

    with pytest.raises(graph_client.GraphAuthError):
        graph_client.acquire_token(cfg)


def test_graph_get_sends_bearer_header_and_base_url(monkeypatch):
    cfg = _fake_cfg()
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")
    captured = {}

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        captured.update(method=method, url=url, headers=headers, timeout=timeout, kwargs=kwargs)
        return _FakeResponse(200, json_data={"value": []})

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    result = graph_client.graph_get("/users/x/messages", cfg, params={"$top": 10})
    assert captured["method"] == "GET"
    assert captured["url"] == "https://graph.microsoft.com/v1.0/users/x/messages"
    assert captured["headers"]["Authorization"] == "Bearer tok-abc"
    assert captured["kwargs"]["params"] == {"$top": 10}
    assert result == {"value": []}


def test_graph_post_sends_json_body(monkeypatch):
    cfg = _fake_cfg()
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")
    captured = {}

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        captured.update(method=method, json=kwargs.get("json"))
        return _FakeResponse(202)

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    graph_client.graph_post("/users/x/sendMail", cfg, json_body={"message": {}})
    assert captured["method"] == "POST"
    assert captured["json"] == {"message": {}}


def test_non_2xx_raises_graph_api_error_with_status_and_graph_message(monkeypatch):
    cfg = _fake_cfg()
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")

    def fake_request(*_args, **_kwargs):
        return _FakeResponse(400, json_data={"error": {"code": "BadRequest", "message": "oops"}})

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    with pytest.raises(graph_client.GraphApiError) as exc_info:
        graph_client.graph_get("/x", cfg)
    assert exc_info.value.status_code == 400
    assert exc_info.value.graph_code == "BadRequest"
    assert exc_info.value.graph_message == "oops"


def test_429_retries_honoring_retry_after_header(monkeypatch):
    cfg = _fake_cfg()
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")
    sleep_calls = []
    monkeypatch.setattr(graph_client.time, "sleep", lambda s: sleep_calls.append(s))

    responses = [
        _FakeResponse(429, json_data={"error": {}}, headers={"Retry-After": "2"}),
        _FakeResponse(200, json_data={"ok": True}),
    ]
    call_count = {"n": 0}

    def fake_request(*_args, **_kwargs):
        resp = responses[call_count["n"]]
        call_count["n"] += 1
        return resp

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    result = graph_client.graph_get("/x", cfg)
    assert result == {"ok": True}
    assert call_count["n"] == 2
    assert sleep_calls == [2.0]


def test_5xx_retries_with_backoff_then_raises_after_max_retries(monkeypatch):
    cfg = _fake_cfg(max_retries=2)
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")
    monkeypatch.setattr(graph_client.time, "sleep", lambda s: None)

    call_count = {"n": 0}

    def fake_request(*_args, **_kwargs):
        call_count["n"] += 1
        return _FakeResponse(503, json_data={"error": {"code": "ServiceUnavailable", "message": "down"}})

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    with pytest.raises(graph_client.GraphApiError):
        graph_client.graph_get("/x", cfg)
    assert call_count["n"] == cfg.graph.max_retries + 1


def test_non_retryable_4xx_fails_immediately(monkeypatch):
    cfg = _fake_cfg()
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")

    call_count = {"n": 0}

    def fake_request(*_args, **_kwargs):
        call_count["n"] += 1
        return _FakeResponse(403, json_data={"error": {"code": "Forbidden", "message": "nope"}})

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    with pytest.raises(graph_client.GraphApiError):
        graph_client.graph_get("/x", cfg)
    assert call_count["n"] == 1


def test_401_triggers_one_forced_refresh_retry_even_with_max_retries_zero(monkeypatch):
    cfg = _fake_cfg(max_retries=0)
    force_refresh_values = []

    def fake_acquire_token(cfg, force_refresh=False):
        force_refresh_values.append(force_refresh)
        return "tok-abc"

    monkeypatch.setattr(graph_client, "acquire_token", fake_acquire_token)

    responses = [
        _FakeResponse(401, json_data={"error": {}}),
        _FakeResponse(200, json_data={"ok": True}),
    ]
    call_count = {"n": 0}

    def fake_request(*_args, **_kwargs):
        resp = responses[call_count["n"]]
        call_count["n"] += 1
        return resp

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    result = graph_client.graph_get("/x", cfg)
    assert result == {"ok": True}
    assert call_count["n"] == 2
    assert force_refresh_values == [False, True]


def test_request_exception_is_retried_then_raises_graph_api_error(monkeypatch):
    cfg = _fake_cfg(max_retries=2)
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")
    monkeypatch.setattr(graph_client.time, "sleep", lambda s: None)

    call_count = {"n": 0}

    def fake_request(*_args, **_kwargs):
        call_count["n"] += 1
        raise graph_client.requests.exceptions.ConnectionError("boom")

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    with pytest.raises(graph_client.GraphApiError) as exc_info:
        graph_client.graph_get("/x", cfg)

    assert call_count["n"] == cfg.graph.max_retries + 1
    assert exc_info.value.status_code == 0
    assert exc_info.value.graph_code == "RequestException"


def test_3xx_response_is_not_treated_as_success(monkeypatch):
    cfg = _fake_cfg(max_retries=0)
    monkeypatch.setattr(graph_client, "acquire_token", lambda cfg, force_refresh=False: "tok-abc")

    def fake_request(*_args, **_kwargs):
        return _FakeResponse(302, json_data={"error": {"code": "Redirect", "message": "moved"}})

    monkeypatch.setattr(graph_client.requests, "request", fake_request)

    with pytest.raises(graph_client.GraphApiError) as exc_info:
        graph_client.graph_get("/x", cfg)
    assert exc_info.value.status_code == 302


def test_acquire_token_raises_graph_auth_error_when_client_secret_missing():
    cfg = _fake_cfg()
    cfg.graph.client_secret = ""

    with pytest.raises(graph_client.GraphAuthError):
        graph_client.acquire_token(cfg)


def test_graph_get_pages_follows_next_link_until_exhausted(monkeypatch):
    cfg = _fake_cfg()
    pages = {
        "/x": {"value": [1], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"},
        "https://graph.microsoft.com/v1.0/next": {"value": [2]},
    }
    monkeypatch.setattr(graph_client, "graph_get", lambda path, cfg, params=None: pages[path])

    results = list(graph_client.graph_get_pages("/x", cfg, params={"$top": 5}))
    assert [r["value"] for r in results] == [[1], [2]]

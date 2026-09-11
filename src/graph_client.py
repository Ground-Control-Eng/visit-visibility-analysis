"""Shared Microsoft Graph auth + HTTP layer used by fetch_email.py and send_summary.py.

Uses app-only (client-credentials) auth via MSAL - no interactive user/session, so none of the
stale-Exchange-session failure modes that used to plague the Outlook COM path can occur here.
"""
from __future__ import annotations

import logging
import time

import msal
import requests

from .config import Config

logger = logging.getLogger("visit_reconciliation")

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

_RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}

_msal_app: msal.ConfidentialClientApplication | None = None


class GraphAuthError(Exception):
    """Raised when acquiring an app-only Graph access token fails."""


class GraphApiError(Exception):
    """Raised when a Graph API call returns a non-2xx response after retries are exhausted."""

    def __init__(self, status_code: int, graph_code: str | None, graph_message: str | None):
        self.status_code = status_code
        self.graph_code = graph_code
        self.graph_message = graph_message
        super().__init__(f"Graph API error {status_code} ({graph_code}): {graph_message}")


def _get_msal_app(cfg: Config) -> msal.ConfidentialClientApplication:
    global _msal_app
    if _msal_app is None:
        _msal_app = msal.ConfidentialClientApplication(
            client_id=cfg.graph.client_id,
            authority=f"https://login.microsoftonline.com/{cfg.graph.tenant_id}",
            client_credential=cfg.graph.client_secret,
        )
    return _msal_app


def acquire_token(cfg: Config, *, force_refresh: bool = False) -> str:
    """Acquires an app-only Graph access token via MSAL's client-credentials flow.

    MSAL's in-memory cache transparently reuses the token across the fetch + send calls within
    one run - this only makes a real token request once per process under normal conditions.

    The client secret is checked here rather than at config-load time, so an offline
    --dry-run-email-path run (which never touches Graph) doesn't need a .env file at all.
    """
    if not cfg.graph.client_secret:
        raise GraphAuthError(
            "GRAPH_CLIENT_SECRET environment variable is not set. Copy .env.example to .env "
            "(in the project root) and fill in the real client secret from the Entra ID app "
            "registration, or set GRAPH_CLIENT_SECRET as a real environment variable."
        )

    app = _get_msal_app(cfg)
    if force_refresh:
        # MSAL has no direct "drop this token" call for the client-credentials flow - removing
        # the cached app instance forces the next acquire_token_for_client to hit the wire again.
        global _msal_app
        _msal_app = None
        app = _get_msal_app(cfg)

    result = app.acquire_token_for_client(scopes=GRAPH_SCOPES)
    if "access_token" not in result:
        raise GraphAuthError(
            f"{result.get('error')}: {result.get('error_description')}"
        )
    return result["access_token"]


def _parse_graph_error(response: requests.Response) -> tuple[str | None, str | None]:
    try:
        body = response.json()
        error = body.get("error", {})
        return error.get("code"), error.get("message")
    except ValueError:
        return None, response.text[:500] if response.text else None


def _request_with_retry(method: str, url: str, cfg: Config, **kwargs) -> requests.Response:
    """Retries retryable HTTP statuses (429/5xx) and transport-level failures (DNS/timeout/
    connection errors) up to cfg.graph.max_retries times, with backoff. The one-time 401
    forced-token-refresh retry is tracked separately (did_401_refresh) so it always gets to run
    once even when max_retries is 0 - it's a distinct concern (a possibly-stale cached token),
    not an ordinary transient-failure retry."""
    max_retries = cfg.graph.max_retries
    retries_used = 0
    did_401_refresh = False
    force_refresh = False
    last_response: requests.Response | None = None
    request_kwargs = dict(kwargs)

    while True:
        token = acquire_token(cfg, force_refresh=force_refresh)
        force_refresh = False
        headers = request_kwargs.pop("headers", {}) or {}
        headers["Authorization"] = f"Bearer {token}"

        try:
            response = requests.request(
                method, url, headers=headers, timeout=cfg.graph.request_timeout_seconds,
                **request_kwargs,
            )
        except requests.RequestException as exc:
            if retries_used < max_retries:
                retries_used += 1
                wait_seconds = min(2 ** (retries_used - 1), 30)
                logger.warning(
                    "Graph request to %s failed (%s) - retrying in %.0fs (%d/%d).",
                    url, exc, wait_seconds, retries_used, max_retries,
                )
                time.sleep(wait_seconds)
                continue
            raise GraphApiError(0, "RequestException", str(exc)) from exc

        if 200 <= response.status_code < 300:
            return response

        last_response = response
        if response.status_code == 401 and not did_401_refresh:
            did_401_refresh = True
            logger.warning("Graph API returned 401 - forcing a token refresh and retrying once.")
            force_refresh = True
            continue
        if response.status_code in _RETRYABLE_STATUS_CODES and retries_used < max_retries:
            retries_used += 1
            retry_after = response.headers.get("Retry-After")
            wait_seconds = float(retry_after) if retry_after else min(2 ** (retries_used - 1), 30)
            logger.warning(
                "Graph API returned %d - retrying in %.0fs (%d/%d).",
                response.status_code, wait_seconds, retries_used, max_retries,
            )
            time.sleep(wait_seconds)
            continue
        break

    graph_code, graph_message = _parse_graph_error(last_response)
    raise GraphApiError(last_response.status_code, graph_code, graph_message)


def _resolve_url(path: str) -> str:
    return path if path.startswith("http") else f"{GRAPH_BASE_URL}{path}"


def graph_get(path: str, cfg: Config, params: dict | None = None) -> dict:
    response = _request_with_retry("GET", _resolve_url(path), cfg, params=params)
    return response.json()


def graph_post(path: str, cfg: Config, json_body: dict) -> requests.Response:
    return _request_with_retry("POST", _resolve_url(path), cfg, json=json_body)


def graph_get_pages(path: str, cfg: Config, params: dict | None = None):
    """Yields each page's response dict from a Graph list endpoint, following @odata.nextLink
    until the collection is exhausted. The first request uses `params`; subsequent requests use
    the full next-link URL as-is (it already encodes the original query). Callers that want a
    bounded search (e.g. a page-count cap) should break out of the loop themselves - this
    generator doesn't limit how many pages it will fetch."""
    next_path: str | None = path
    page = 0
    while next_path:
        page += 1
        resp = graph_get(next_path, cfg, params=params if page == 1 else None)
        yield resp
        next_path = resp.get("@odata.nextLink")

"""Loads and validates config.yaml so no code changes are needed to retune the pipeline."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config.yaml"
DEFAULT_ENV_PATH = PROJECT_ROOT / ".env"

_PLACEHOLDER_MARKERS = ("PLACEHOLDER", "CONFIRM")


class ConfigError(Exception):
    """Raised when config.yaml is missing, malformed, or contains unfilled placeholders."""


@dataclass
class SqlServerConfig:
    server: str
    database: str
    driver: str
    auth: str
    extra: dict = field(default_factory=dict)


@dataclass
class ExcelSourceConfig:
    path: Path
    sheet_name: str
    table_name: str
    refresh_timeout_seconds: float


@dataclass
class EmailTriggerConfig:
    sender_filter: str
    subject_contains: str
    search_folder: str
    lookback_hours: int
    attachment_name_pattern: str


@dataclass
class EmailConfig:
    trigger: EmailTriggerConfig
    to: list
    cc: list
    send_on_success: bool
    send_on_failure: bool


@dataclass
class GraphConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    mailbox: str
    request_timeout_seconds: float = 30.0
    max_retries: int = 3
    max_inline_attachment_bytes: int = 3_000_000
    mail_search_page_size: int = 50
    mail_search_max_pages: int = 5


@dataclass
class TeamsConfig:
    webhook_url: str | None


@dataclass
class RunConfig:
    output_dir: Path
    keep_days: int
    test_mode: bool


@dataclass
class Config:
    ice2: SqlServerConfig
    visits_api_excel: ExcelSourceConfig
    ice2_query_start_date: str
    ice2_excluded_contractor_ids: list[int]
    ice2_exclude_de_teams: bool
    visits_api_status_title_column: str
    recent_window_days: int
    legitimately_excluded_from_hubscape: set
    status_legend: dict
    email: EmailConfig
    hubscape_api_id_column: str
    run: RunConfig
    # Trailing + defaulted so existing call sites (tests building a Config by hand) don't need
    # to know about it - real runs get the loaded value from load_config() below regardless.
    teams: TeamsConfig = field(default_factory=lambda: TeamsConfig(webhook_url=None))
    graph: GraphConfig = field(
        default_factory=lambda: GraphConfig(tenant_id="", client_id="", client_secret="", mailbox="")
    )


def _check_placeholder(value, path: str) -> None:
    if isinstance(value, str) and any(marker in value.upper() for marker in _PLACEHOLDER_MARKERS):
        raise ConfigError(
            f"config.yaml key '{path}' still contains a placeholder value ({value!r}). "
            "Fill it in with the real value before running the pipeline."
        )


def _load_dotenv(path: Path) -> None:
    """Minimal KEY=VALUE .env loader - sets os.environ for keys not already set (never
    overrides a real environment variable). A dependency like python-dotenv is disproportionate
    for the single secret (GRAPH_CLIENT_SECRET) this project needs."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def load_config(path: Path | str = DEFAULT_CONFIG_PATH) -> Config:
    path = Path(path)
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")

    _load_dotenv(DEFAULT_ENV_PATH)

    with path.open("r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)

    try:
        sql_raw = raw["sql"]
        ice2_raw = sql_raw["ice2"]
        api_raw = sql_raw["visits_api_excel"]
        recency_raw = raw.get("recency", {})
        statuses_raw = raw["statuses"]
        email_raw = raw["email"]
        trigger_raw = email_raw["trigger"]
        recipients_raw = email_raw["recipients"]
        hub_cols_raw = raw["hubscape_columns"]
        run_raw = raw["run"]
        graph_raw = raw["graph"]
    except KeyError as exc:
        raise ConfigError(f"config.yaml is missing required section: {exc}") from exc

    for key in ("server", "database"):
        _check_placeholder(ice2_raw.get(key), f"sql.ice2.{key}")
    _check_placeholder(api_raw.get("path"), "sql.visits_api_excel.path")
    _check_placeholder(trigger_raw.get("sender_filter"), "email.trigger.sender_filter")
    _check_placeholder(hub_cols_raw.get("api_id_column"), "hubscape_columns.api_id_column")
    for key in ("tenant_id", "client_id", "mailbox"):
        _check_placeholder(graph_raw.get(key), f"graph.{key}")

    if not trigger_raw.get("sender_filter") and not trigger_raw.get("subject_contains"):
        raise ConfigError(
            "email.trigger must specify at least one of sender_filter / subject_contains, "
            "otherwise every email in the folder would match."
        )

    ice2 = SqlServerConfig(
        server=ice2_raw["server"],
        database=ice2_raw["database"],
        driver=ice2_raw.get("driver", "ODBC Driver 17 for SQL Server"),
        auth=ice2_raw.get("auth", "trusted"),
    )
    excel_path = Path(api_raw["path"])
    if not excel_path.is_absolute():
        excel_path = (PROJECT_ROOT / excel_path).resolve()
    visits_api_excel = ExcelSourceConfig(
        path=excel_path,
        sheet_name=api_raw.get("sheet_name", "Query1"),
        table_name=api_raw.get("table_name", "Query1"),
        refresh_timeout_seconds=float(api_raw.get("refresh_timeout_seconds", 180)),
    )

    email = EmailConfig(
        trigger=EmailTriggerConfig(
            sender_filter=trigger_raw.get("sender_filter", ""),
            subject_contains=trigger_raw.get("subject_contains", ""),
            search_folder=trigger_raw.get("search_folder", "Inbox"),
            lookback_hours=int(trigger_raw.get("lookback_hours", 30)),
            attachment_name_pattern=trigger_raw.get("attachment_name_pattern", "*"),
        ),
        to=recipients_raw.get("to", []),
        cc=recipients_raw.get("cc", []),
        send_on_success=bool(email_raw.get("send_on_success", True)),
        send_on_failure=bool(email_raw.get("send_on_failure", True)),
    )

    if not email.to:
        raise ConfigError("email.recipients.to must contain at least one recipient.")

    graph_client_secret = os.environ.get("GRAPH_CLIENT_SECRET", "")
    if not graph_client_secret:
        raise ConfigError(
            "GRAPH_CLIENT_SECRET environment variable is not set. Copy .env.example to .env "
            "(in the project root) and fill in the real client secret from the Entra ID app "
            "registration, or set GRAPH_CLIENT_SECRET as a real environment variable."
        )
    graph = GraphConfig(
        tenant_id=graph_raw["tenant_id"],
        client_id=graph_raw["client_id"],
        client_secret=graph_client_secret,
        mailbox=graph_raw["mailbox"],
        request_timeout_seconds=float(graph_raw.get("request_timeout_seconds", 30)),
        max_retries=int(graph_raw.get("max_retries", 3)),
        max_inline_attachment_bytes=int(graph_raw.get("max_inline_attachment_bytes", 3_000_000)),
        mail_search_page_size=int(graph_raw.get("mail_search_page_size", 50)),
        mail_search_max_pages=int(graph_raw.get("mail_search_max_pages", 5)),
    )

    run = RunConfig(
        output_dir=(PROJECT_ROOT / run_raw.get("output_dir", "./output")).resolve(),
        keep_days=int(run_raw.get("keep_days", 90)),
        test_mode=bool(run_raw.get("test_mode", True)),
    )

    teams_raw = raw.get("notifications", {}).get("teams", {}) or {}
    teams = TeamsConfig(webhook_url=teams_raw.get("webhook_url") or None)

    return Config(
        ice2=ice2,
        visits_api_excel=visits_api_excel,
        ice2_query_start_date=ice2_raw.get("query_start_date", "2024-04-01"),
        ice2_excluded_contractor_ids=list(ice2_raw.get("excluded_contractor_ids", [])),
        ice2_exclude_de_teams=bool(ice2_raw.get("exclude_de_teams", False)),
        visits_api_status_title_column=api_raw.get("status_title_column", "Title"),
        recent_window_days=int(recency_raw.get("recent_window_days", 30)),
        legitimately_excluded_from_hubscape=set(
            statuses_raw.get("legitimately_excluded_from_hubscape", [])
        ),
        status_legend=statuses_raw.get("legend", {}),
        email=email,
        hubscape_api_id_column=hub_cols_raw.get("api_id_column", "External Visit API Id"),
        teams=teams,
        graph=graph,
        run=run,
    )

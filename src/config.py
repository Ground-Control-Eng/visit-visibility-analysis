"""Loads and validates config.yaml so no code changes are needed to retune the pipeline."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config.yaml"

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
    outlook_profile: str | None
    trigger: EmailTriggerConfig
    to: list
    cc: list
    send_on_success: bool
    send_on_failure: bool


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


def _check_placeholder(value, path: str) -> None:
    if isinstance(value, str) and any(marker in value.upper() for marker in _PLACEHOLDER_MARKERS):
        raise ConfigError(
            f"config.yaml key '{path}' still contains a placeholder value ({value!r}). "
            "Fill it in with the real value before running the pipeline."
        )


def load_config(path: Path | str = DEFAULT_CONFIG_PATH) -> Config:
    path = Path(path)
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")

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
    except KeyError as exc:
        raise ConfigError(f"config.yaml is missing required section: {exc}") from exc

    for key in ("server", "database"):
        _check_placeholder(ice2_raw.get(key), f"sql.ice2.{key}")
    _check_placeholder(api_raw.get("path"), "sql.visits_api_excel.path")
    _check_placeholder(trigger_raw.get("sender_filter"), "email.trigger.sender_filter")
    _check_placeholder(hub_cols_raw.get("api_id_column"), "hubscape_columns.api_id_column")

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
        outlook_profile=email_raw.get("outlook_profile"),
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

    run = RunConfig(
        output_dir=(PROJECT_ROOT / run_raw.get("output_dir", "./output")).resolve(),
        keep_days=int(run_raw.get("keep_days", 90)),
        test_mode=bool(run_raw.get("test_mode", True)),
    )

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
        run=run,
    )

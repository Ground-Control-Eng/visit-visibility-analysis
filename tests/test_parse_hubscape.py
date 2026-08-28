import sys
from pathlib import Path

import pandas as pd


sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config import Config, EmailConfig, EmailTriggerConfig, ExcelSourceConfig, RunConfig, SqlServerConfig
from src.parse_hubscape import parse_attachment
from src.reconcile import reconcile


def make_cfg() -> Config:
    return Config(
        ice2=SqlServerConfig(server="s", database="d", driver="drv", auth="trusted"),
        visits_api_excel=ExcelSourceConfig(
            path=Path("API Status.xlsx"), sheet_name="Query1", table_name="Query1",
            refresh_timeout_seconds=180,
        ),
        ice2_query_start_date="2024-04-01",
        ice2_excluded_contractor_ids=[],
        ice2_exclude_de_teams=False,
        visits_api_status_title_column="Title",
        recent_window_days=30,
        legitimately_excluded_from_hubscape={50, 60, 70},
        status_legend={},
        email=EmailConfig(
            outlook_profile=None,
            trigger=EmailTriggerConfig(
                sender_filter="ci@hubscape.co.uk", subject_contains="", search_folder="Inbox",
                lookback_hours=30, attachment_name_pattern="production_uncompleted_visits_*.csv",
            ),
            to=["alex.clark@ground-control.co.uk"], cc=[], send_on_success=True, send_on_failure=True,
        ),
        hubscape_api_id_column="External Visit API Id",
        run=RunConfig(output_dir=Path("./output"), keep_days=90, test_mode=True),
    )


def test_blank_api_id_cell_stays_real_nan_not_the_string_nan(tmp_path):
    # A numeric API_ID column with one blank cell alongside other real columns forces pandas to
    # infer float64 for the column (the same shape a real Hubscape export produces) - the exact
    # dtype path that used to risk stringifying the blank cell to the literal "nan" instead of
    # preserving it as real NaN, which would silently defeat reconcile.py's
    # HUBSCAPE_MISSING_API_ID accounting downstream (reconcile.py's own _to_nullable_str_id then
    # cleans up the ".0" float formatting on the non-blank values - not parse_hubscape's job).
    csv_path = tmp_path / "hubscape.csv"
    csv_path.write_text(
        "External Visit API Id,Status\n123,SCHEDULED\n456,SCHEDULED\n,UNSCHEDULED\n789,SCHEDULED\n",
        encoding="utf-8",
    )
    cfg = make_cfg()

    df = parse_attachment(csv_path, cfg)

    assert df["API_ID"].isna().sum() == 1
    assert pd.isna(df.loc[2, "API_ID"])
    assert df.loc[0, "API_ID"] == "123.0"
    assert df.loc[1, "API_ID"] == "456.0"
    assert df.loc[3, "API_ID"] == "789.0"


def test_blank_hubscape_api_id_reaches_reconcile_as_hubscape_missing_api_id(tmp_path):
    # End-to-end version of the above: routes a real Hubscape CSV with a blank API_ID cell
    # through parse_attachment() (not a synthetic hub_df built by hand) and confirms it still
    # gets counted as HUBSCAPE_MISSING_API_ID by reconcile.py, rather than silently surviving as
    # the literal string "nan" and being merged into HUBSCAPE_MATCHED/an orphan instead.
    csv_path = tmp_path / "hubscape.csv"
    csv_path.write_text(
        "External Visit API Id,Status\n123,SCHEDULED\n,UNSCHEDULED\n",
        encoding="utf-8",
    )
    cfg = make_cfg()
    hub_df = parse_attachment(csv_path, cfg)

    ice2 = pd.DataFrame([{
        "VisitID": "1", "ExpectedStartDate": "2026-01-01", "ExpectedEndDate": "2026-01-01",
        "VisitStatusID": 20, "StatusDescription": "In Progress",
    }])
    api_df = pd.DataFrame([{"API_ID": "123", "VisitID": "1", "Title": "In Progress"}])

    summary, _ = reconcile(ice2, api_df, hub_df, cfg)
    counts = dict(zip(summary["metric"], summary["count"]))

    assert counts["HUBSCAPE_MISSING_API_ID"] == 1
    assert counts["TOTAL_HUBSCAPE_ROWS"] == 2

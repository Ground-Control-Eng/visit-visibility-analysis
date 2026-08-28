import sys
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config import Config, EmailConfig, EmailTriggerConfig, ExcelSourceConfig, RunConfig, SqlServerConfig
from src.reconcile import classify_team_type, missing_from_hubscape_breakdown, orphan_candidate_visit_ids, reconcile


def make_cfg() -> Config:
    return Config(
        ice2=SqlServerConfig(server="s", database="d", driver="drv", auth="trusted"),
        visits_api_excel=ExcelSourceConfig(
            path=Path("API Status.xlsx"), sheet_name="Query1", table_name="Query1",
            refresh_timeout_seconds=180,
        ),
        ice2_query_start_date="2024-04-01",
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


def make_ice2_row(visit_id, status_id, status_desc):
    return {
        "VisitID": visit_id,
        "ExpectedStartDate": "2026-01-01",
        "ExpectedEndDate": "2026-01-01",
        "VisitStatusID": status_id,
        "StatusDescription": status_desc,
    }


def hub_row(api_id):
    # Real Hubscape exports key on API_ID alone - they never expose ICe2's VisitID.
    return {"API_ID": api_id}


def test_status_mismatch():
    ice2 = pd.DataFrame([make_ice2_row("1", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "1", "Title": "Completed"}])
    hub = pd.DataFrame([hub_row("A1")])
    cfg = make_cfg()

    summary, detail = reconcile(ice2, api, hub, cfg)

    assert detail.loc[detail["VisitID"] == "1", "issue_type"].iloc[0] == "ICE2_API_STATUS_MISMATCH"
    assert int(summary.loc[summary["metric"] == "ICE2_API_STATUS_MISMATCH", "count"].iloc[0]) == 1


def test_missing_api_mapping():
    # No row in the Visits API extract (i.e. no Entitymapping.SourceID) matches this VisitID -
    # the API itself has no record of this visit, so API_ID can't be resolved at all.
    ice2 = pd.DataFrame([make_ice2_row("2", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "1", "Title": "In Progress"}])
    hub = pd.DataFrame([hub_row("A1")])
    cfg = make_cfg()

    summary, detail = reconcile(ice2, api, hub, cfg)

    assert detail.loc[detail["VisitID"] == "2", "issue_type"].iloc[0] == "ICE2_MISSING_API_MAPPING"
    assert int(summary.loc[summary["metric"] == "ICE2_MISSING_API_MAPPING", "count"].iloc[0]) == 1


def test_missing_from_hubscape_when_active_and_recent():
    ice2 = pd.DataFrame([make_ice2_row("3", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "3", "Title": "In Progress"}])
    hub = pd.DataFrame([], columns=["API_ID"])
    cfg = make_cfg()

    # ice2 row's ExpectedStartDate/EndDate default to 2026-01-01 - within 30 days of this ref date.
    summary, detail = reconcile(ice2, api, hub, cfg, reference_date=date(2026, 1, 10))

    assert detail.loc[detail["VisitID"] == "3", "issue_type"].iloc[0] == "MISSING_FROM_HUBSCAPE"


def test_missing_from_hubscape_when_active_but_historic():
    ice2 = pd.DataFrame([make_ice2_row("3h", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "3h", "Title": "In Progress"}])
    hub = pd.DataFrame([], columns=["API_ID"])
    cfg = make_cfg()

    # ice2 row defaults to 2026-01-01 - far more than 30 days before this ref date.
    summary, detail = reconcile(ice2, api, hub, cfg, reference_date=date(2026, 12, 1))

    assert "3h" not in set(detail["VisitID"])
    assert int(summary.loc[summary["metric"] == "MISSING_FROM_HUBSCAPE_HISTORIC", "count"].iloc[0]) == 1


def test_detail_df_surfaces_internalcontractor_as_bool():
    ice2 = pd.DataFrame([
        {**make_ice2_row("3de", 20, "In Progress"), "internalcontractor": 1},
        {**make_ice2_row("3sub", 20, "In Progress"), "internalcontractor": 0},
    ])
    api = pd.DataFrame([
        {"API_ID": "A1", "VisitID": "3de", "Title": "In Progress"},
        {"API_ID": "A2", "VisitID": "3sub", "Title": "In Progress"},
    ])
    hub = pd.DataFrame([], columns=["API_ID"])
    cfg = make_cfg()

    # ice2 rows default to 2026-01-01 - within 30 days of this ref date, so both land in detail_df.
    summary, detail = reconcile(ice2, api, hub, cfg, reference_date=date(2026, 1, 10))

    assert detail.loc[detail["VisitID"] == "3de", "internalcontractor"].iloc[0] is True
    assert detail.loc[detail["VisitID"] == "3sub", "internalcontractor"].iloc[0] is False


def test_legitimately_absent_when_completed():
    ice2 = pd.DataFrame([make_ice2_row("4", 60, "Completed")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "4", "Title": "Completed"}])
    hub = pd.DataFrame([], columns=["API_ID"])
    cfg = make_cfg()

    summary, detail = reconcile(ice2, api, hub, cfg)

    # Informational only - counted in summary, but deliberately excluded from row-level detail
    # (it would otherwise be dominated by every historically completed/cancelled visit).
    assert "4" not in set(detail["VisitID"])
    assert int(summary.loc[summary["metric"] == "LEGITIMATELY_ABSENT_FROM_HUBSCAPE", "count"].iloc[0]) == 1


def test_orphan_in_hubscape():
    ice2 = pd.DataFrame([make_ice2_row("5", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "5", "Title": "In Progress"}])
    hub = pd.DataFrame([hub_row("A1"), hub_row("A999")])
    cfg = make_cfg()

    summary, detail = reconcile(ice2, api, hub, cfg)

    orphan_rows = detail[detail["issue_type"] == "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2"]
    assert list(orphan_rows["API_ID"]) == ["A999"]


def test_orphan_candidate_visit_ids_resolves_via_api_extract():
    ice2 = pd.DataFrame([make_ice2_row("5", 20, "In Progress")])
    api = pd.DataFrame([
        {"API_ID": "A1", "VisitID": "5", "Title": "In Progress"},
        {"API_ID": "A999", "VisitID": "999", "Title": "Completed"},
    ])
    hub = pd.DataFrame([hub_row("A1"), hub_row("A999")])
    cfg = make_cfg()

    # A999 is an orphan candidate (in Hubscape, not in the filtered ICe2 extract) whose VisitID
    # (999) is resolvable via the Visits API extract - worth a follow-up ICe2 status lookup.
    assert orphan_candidate_visit_ids(ice2, api, hub, cfg) == ["999"]


def test_orphan_reclassified_when_ice2_status_lookup_resolves_it():
    ice2 = pd.DataFrame([make_ice2_row("5", 20, "In Progress")])
    api = pd.DataFrame([
        {"API_ID": "A1", "VisitID": "5", "Title": "In Progress"},
        {"API_ID": "A999", "VisitID": "999", "Title": "Completed"},
    ])
    hub = pd.DataFrame([hub_row("A1"), hub_row("A999")])
    cfg = make_cfg()
    # ICe2 does know about VisitID 999 - it's just Completed (60), outside the extract's filter.
    lookup = pd.DataFrame([{"VisitID": "999", "VisitStatusID": 60}])

    summary, detail = reconcile(ice2, api, hub, cfg, orphan_status_lookup=lookup)

    assert "A999" not in set(detail["API_ID"])
    assert int(summary.loc[summary["metric"] == "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2", "count"].iloc[0]) == 0
    assert int(summary.loc[summary["metric"] == "ORPHAN_ICE2_STATUS_EXCLUDED", "count"].iloc[0]) == 1


def test_orphan_stays_true_orphan_when_status_lookup_does_not_resolve_it():
    ice2 = pd.DataFrame([make_ice2_row("5", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "5", "Title": "In Progress"}])
    hub = pd.DataFrame([hub_row("A1"), hub_row("A999")])
    cfg = make_cfg()
    # A999 has no VisitID in the Visits API extract at all, so the lookup (even if non-empty for
    # other candidates) can never resolve it - it's genuinely unknown to ICe2.
    lookup = pd.DataFrame([{"VisitID": "1234", "VisitStatusID": 60}])

    summary, detail = reconcile(ice2, api, hub, cfg, orphan_status_lookup=lookup)

    orphan_rows = detail[detail["issue_type"] == "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2"]
    assert list(orphan_rows["API_ID"]) == ["A999"]
    assert int(summary.loc[summary["metric"] == "ORPHAN_ICE2_STATUS_EXCLUDED", "count"].iloc[0]) == 0


def test_classify_team_type_distinguishes_de_subcontractor_and_unknown():
    cfg = make_cfg()
    flags = pd.Series([1, 1, 0, 0, None])

    result = classify_team_type(flags, cfg)

    assert list(result) == ["DE", "DE", "Subcontractor", "Subcontractor", "Unknown"]


def test_missing_from_hubscape_breakdown_spans_recent_and_historic_years():
    cfg = make_cfg()

    ice2 = pd.DataFrame([
        {**make_ice2_row("10", 20, "In Progress"), "ExpectedStartDate": "2026-01-15",
         "ExpectedEndDate": "2026-01-15", "ContractorID": 1, "FieldTeamName": "DE SEN",
         "internalcontractor": 1},
        {**make_ice2_row("11", 20, "In Progress"), "ExpectedStartDate": "2024-06-01",
         "ExpectedEndDate": "2024-06-01", "ContractorID": 3, "FieldTeamName": "JC Landscapes",
         "internalcontractor": 0},
        {**make_ice2_row("12", 20, "In Progress"), "ExpectedStartDate": "2026-01-15",
         "ExpectedEndDate": "2026-01-15", "ContractorID": 3, "FieldTeamName": "JC Landscapes",
         "internalcontractor": 0},
    ])
    api = pd.DataFrame([
        {"API_ID": "A10", "VisitID": "10", "Title": "In Progress"},
        {"API_ID": "A11", "VisitID": "11", "Title": "In Progress"},
        {"API_ID": "A12", "VisitID": "12", "Title": "In Progress"},
    ])
    # A12 is present in Hubscape - should be excluded from the breakdown.
    hub = pd.DataFrame([hub_row("A12")])

    # reference_date is deliberately far from ExpectedStartDate for VisitID 11 (>30 days), so
    # reconcile() would bucket it as MISSING_FROM_HUBSCAPE_HISTORIC - it must still appear here.
    breakdown = missing_from_hubscape_breakdown(ice2, api, hub, cfg)
    breakdown_lookup = {(row.year, row.team_type): row.count for row in breakdown.itertuples()}

    assert breakdown_lookup == {(2026, "DE"): 1, (2024, "Subcontractor"): 1}
    assert (2026, "Subcontractor") not in breakdown_lookup


def test_clean_match_produces_no_issue():
    ice2 = pd.DataFrame([make_ice2_row("6", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "6", "Title": "In Progress"}])
    hub = pd.DataFrame([hub_row("A1")])
    cfg = make_cfg()

    summary, detail = reconcile(ice2, api, hub, cfg)

    assert "6" not in set(detail["VisitID"])
    total_alerts = summary.loc[summary["metric"].isin([
        "ICE2_MISSING_API_MAPPING", "ICE2_API_STATUS_MISMATCH",
        "MISSING_FROM_HUBSCAPE", "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2",
    ]), "count"].sum()
    assert total_alerts == 0
    assert int(summary.loc[summary["metric"] == "ICE2_MATCHED_OK", "count"].iloc[0]) == 1
    assert int(summary.loc[summary["metric"] == "HUBSCAPE_MATCHED", "count"].iloc[0]) == 1


def test_hubscape_row_missing_api_id_is_counted_not_dropped():
    ice2 = pd.DataFrame([make_ice2_row("7", 20, "In Progress")])
    api = pd.DataFrame([{"API_ID": "A1", "VisitID": "7", "Title": "In Progress"}])
    # A Hubscape row with no API_ID at all used to vanish from every bucket - it must now be
    # counted explicitly, and TOTAL_HUBSCAPE_ROWS accounting must still balance.
    hub = pd.DataFrame([hub_row("A1"), {"API_ID": None}])
    cfg = make_cfg()

    summary, detail = reconcile(ice2, api, hub, cfg)

    assert int(summary.loc[summary["metric"] == "HUBSCAPE_MISSING_API_ID", "count"].iloc[0]) == 1
    assert int(summary.loc[summary["metric"] == "TOTAL_HUBSCAPE_ROWS", "count"].iloc[0]) == 2


def test_reconciliation_accounting_covers_every_bucket_and_balances():
    # One row per bucket (except ICE2_MISSING_API_MAPPING, which by construction has no API_ID
    # and so can never appear in Hubscape or the orphan-status lookup either).
    ice2 = pd.DataFrame([
        make_ice2_row("miss_map", 20, "In Progress"),  # no matching API row -> ICE2_MISSING_API_MAPPING
        make_ice2_row("mismatch", 20, "In Progress"),  # status differs from API -> ICE2_API_STATUS_MISMATCH
        make_ice2_row("missing_recent", 20, "In Progress"),  # not in hub, recent -> MISSING_FROM_HUBSCAPE
        make_ice2_row("missing_hist", 20, "In Progress"),  # not in hub, old -> MISSING_FROM_HUBSCAPE_HISTORIC
        make_ice2_row("legit_absent", 60, "Completed"),  # excluded status, not in hub -> LEGITIMATELY_ABSENT_FROM_HUBSCAPE
        make_ice2_row("matched", 20, "In Progress"),  # in hub, status matches -> ICE2_MATCHED_OK
    ])
    ice2.loc[ice2["VisitID"] == "missing_hist", "ExpectedStartDate"] = "2020-01-01"
    ice2.loc[ice2["VisitID"] == "missing_hist", "ExpectedEndDate"] = "2020-01-01"

    api = pd.DataFrame([
        {"API_ID": "A_mismatch", "VisitID": "mismatch", "Title": "Completed"},
        {"API_ID": "A_missing_recent", "VisitID": "missing_recent", "Title": "In Progress"},
        {"API_ID": "A_missing_hist", "VisitID": "missing_hist", "Title": "In Progress"},
        {"API_ID": "A_legit_absent", "VisitID": "legit_absent", "Title": "Completed"},
        {"API_ID": "A_matched", "VisitID": "matched", "Title": "In Progress"},
        {"API_ID": "A_orphan_status_excluded", "VisitID": "orphan_status_excluded", "Title": "Completed"},
    ])

    hub = pd.DataFrame([
        {"API_ID": "A_matched"},
        {"API_ID": "A_orphan"},  # unknown to ICe2 at all -> ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2
        {"API_ID": "A_orphan_status_excluded"},  # known to ICe2, excluded by status -> ORPHAN_ICE2_STATUS_EXCLUDED
        {"API_ID": None},  # -> HUBSCAPE_MISSING_API_ID
    ])
    orphan_status_lookup = pd.DataFrame([{"VisitID": "orphan_status_excluded", "VisitStatusID": 60}])
    cfg = make_cfg()

    summary, _ = reconcile(
        ice2, api, hub, cfg, reference_date=date(2026, 1, 10), orphan_status_lookup=orphan_status_lookup,
    )
    counts = dict(zip(summary["metric"], summary["count"]))

    ice2_side_total = sum(counts[m] for m in [
        "ICE2_MISSING_API_MAPPING", "ICE2_API_STATUS_MISMATCH", "MISSING_FROM_HUBSCAPE",
        "MISSING_FROM_HUBSCAPE_HISTORIC", "LEGITIMATELY_ABSENT_FROM_HUBSCAPE", "ICE2_MATCHED_OK",
    ])
    hubscape_side_total = sum(counts[m] for m in [
        "HUBSCAPE_MATCHED", "HUBSCAPE_MISSING_API_ID",
        "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2", "ORPHAN_ICE2_STATUS_EXCLUDED",
    ])

    assert ice2_side_total == counts["TOTAL_ICE2_ROWS"] == len(ice2)
    assert hubscape_side_total == counts["TOTAL_HUBSCAPE_ROWS"] == len(hub)

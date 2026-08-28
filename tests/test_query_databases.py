import sys
from pathlib import Path

import openpyxl
import pandas as pd
import sqlalchemy
from sqlalchemy.dialects import mssql

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config import ExcelSourceConfig
from src.query_databases import (
    ICE2_CONTRACTOR_JOIN_SQL,
    ICE2_QUERY_TEMPLATE,
    ICE2_STATUS_LOOKUP_TEMPLATE,
    get_ice2_status_by_visit_ids,
    _read_excel_table,
)


def _make_excel_cfg(path):
    return ExcelSourceConfig(
        path=path, sheet_name="Query1", table_name="Query1", refresh_timeout_seconds=5,
    )


def test_read_excel_table_matches_visits_api_shape(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Query1"
    ws.append(["API_ID", "VisitID", "Id", "Title"])
    ws.append([1, "3235412", 8, "Completed"])
    ws.append([2, "3235994", 8, "Completed"])
    path = tmp_path / "API Status.xlsx"
    wb.save(path)

    excel_cfg = _make_excel_cfg(path)
    df = _read_excel_table(excel_cfg)

    assert list(df.columns) == ["API_ID", "VisitID", "Id", "Title"]
    assert len(df) == 2
    assert df.loc[0, "VisitID"] == "3235412"
    assert df.loc[0, "Title"] == "Completed"


def _compile_ice2_query(excluded_contractor_ids, exclude_de_teams=True):
    query = sqlalchemy.text(ICE2_QUERY_TEMPLATE).bindparams(
        sqlalchemy.bindparam("excluded_contractor_ids", expanding=True)
    )
    bound = query.bindparams(
        start_date="2025-10-01",
        excluded_contractor_ids=excluded_contractor_ids,
        exclude_de_teams=1 if exclude_de_teams else 0,
    )
    compiled = bound.compile(dialect=mssql.dialect(), compile_kwargs={"render_postcompile": True})
    return str(compiled), compiled.params


def test_ice2_query_excludes_configured_contractor_ids():
    sql, params = _compile_ice2_query([3244, 8030, 8031])

    assert "j.ContractorID IS NULL OR j.ContractorID NOT IN" in sql
    assert set(params.values()) >= {3244, 8030, 8031}


def test_ice2_query_with_no_excluded_contractor_ids_excludes_nothing():
    # An empty exclusion list must compile to an always-true NOT IN, not a syntax error and not
    # a NOT IN that (via SQL's NULL/empty-set handling) accidentally excludes every row.
    sql, _ = _compile_ice2_query([])

    assert "NOT IN (SELECT 1 WHERE 1!=1)" in sql


def test_ice2_query_excludes_de_teams_when_enabled():
    sql, params = _compile_ice2_query([], exclude_de_teams=True)

    assert "ft.internalcontractor IS NULL OR ft.internalcontractor <> 1" in sql
    assert params["exclude_de_teams"] == 1


def test_ice2_query_keeps_de_teams_when_disabled():
    sql, params = _compile_ice2_query([], exclude_de_teams=False)

    # The clause is static SQL either way (the toggle is the bind value, not a string swap) -
    # `:exclude_de_teams = 0` short-circuits the OR so no row gets dropped on this condition.
    assert "ft.internalcontractor IS NULL OR ft.internalcontractor <> 1" in sql
    assert params["exclude_de_teams"] == 0


def test_ice2_query_and_status_lookup_share_identical_contractor_join():
    # Drift tripwire: ICE2_STATUS_LOOKUP_TEMPLATE's contractor/DE-team resolution must stay
    # byte-for-byte identical to ICE2_QUERY_TEMPLATE's own join - if someone edits one without
    # mirroring the other, reconcile.py's ORPHAN_ICE2_TEAM_EXCLUDED classification would silently
    # stop matching the main extract's own exclusion criteria.
    assert ICE2_CONTRACTOR_JOIN_SQL in ICE2_QUERY_TEMPLATE
    assert ICE2_CONTRACTOR_JOIN_SQL in ICE2_STATUS_LOOKUP_TEMPLATE


def test_ice2_status_lookup_query_includes_contractor_join_and_columns():
    query = sqlalchemy.text(ICE2_STATUS_LOOKUP_TEMPLATE).bindparams(
        sqlalchemy.bindparam("visit_ids", expanding=True)
    )
    bound = query.bindparams(visit_ids=["123", "456"])
    compiled = bound.compile(dialect=mssql.dialect(), compile_kwargs={"render_postcompile": True})
    sql = str(compiled)

    assert "LEFT JOIN tblContractor ft" in sql
    assert "j.ContractorID" in sql
    assert "ft.internalcontractor" in sql
    assert "WHERE v.VisitID IN" in sql
    assert set(compiled.params.values()) >= {"123", "456"}


def test_get_ice2_status_by_visit_ids_empty_input_returns_expected_columns():
    # The early-return branch never touches cfg.ice2, so a dummy cfg is safe here.
    df = get_ice2_status_by_visit_ids(cfg=None, visit_ids=[])

    assert list(df.columns) == ["VisitID", "VisitStatusID", "ContractorID", "internalcontractor"]
    assert df.empty

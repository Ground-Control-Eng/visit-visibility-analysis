"""Runs the ICe2 SQL query via integrated auth and reads the Visits API extract from its
live-connection Excel workbook, returning pandas DataFrames."""
from __future__ import annotations

import logging
import threading
import urllib.parse

import openpyxl
import pandas as pd
import sqlalchemy

from .config import Config, ExcelSourceConfig, SqlServerConfig

logger = logging.getLogger("visit_reconciliation")

ICE2_QUERY_TEMPLATE = """
SELECT v.VisitID, v.ExpectedStartDate, v.ExpectedEndDate, v.StartDate, v.EndDate,
       v.VisitStatusID, vs.StatusDescription, j.ContractorID, ft.strName AS FieldTeamName, ft.internalcontractor
FROM OM_Visit v
    INNER JOIN OM_Site_Contract sc ON sc.sitecontractid = v.sitecontractid
    INNER JOIN OM_Client_Contract cc ON cc.ClientContractID = sc.clientcontractid
    INNER JOIN OM_Visit_Status vs ON v.visitstatusID = vs.visitstatusID
    LEFT JOIN (
        SELECT j.VisitID, jlu.ContractorID,
               ROW_NUMBER() OVER (PARTITION BY j.VisitID ORDER BY jlu.LabourUsedID DESC) AS rn
        FROM OM_Job j
        INNER JOIN Om_Job_Labour_Used jlu ON j.jobid = jlu.jobid
    ) j ON v.VisitID = j.VisitID AND j.rn = 1
    LEFT JOIN tblContractor ft ON j.ContractorID = ft.intContractorID
WHERE v.VisitStatusID IN (5,10,15,20,35) -- 50,55,60,70 removed
  AND v.ExpectedStartDate >= :start_date
  AND ft.istest = 0
ORDER BY v.VisitID;
"""
# API_ID is no longer sourced from ICe2's own GCIS_API_Record_Mapping table (which could
# hold a stale/dangling ExternalRecordID pointing at no real Visits API record). Instead,
# reconcile.py looks it up by matching VisitID against the Visits API extract's own VisitID
# column, which is itself sourced from the Visits API's Entitymapping.SourceID - i.e. the
# API's own authoritative record of which VisitID maps to which API_ID.


class DatabaseQueryError(Exception):
    def __init__(self, source_name: str, original_exception: Exception):
        self.source_name = source_name
        self.original_exception = original_exception
        super().__init__(f"Query against '{source_name}' failed: {original_exception}")


def _build_connection_url(cfg: SqlServerConfig) -> str:
    # Built as a raw ODBC connection string passed through the `odbc_connect` escape hatch,
    # because sqlalchemy's mssql+pyodbc dialect unconditionally injects Trusted_Connection=Yes
    # into the connect string when no username/password is present in the URL, which conflicts
    # with (and cannot be overridden alongside) the Authentication=ActiveDirectoryIntegrated
    # keyword needed for Azure AD integrated auth.
    if cfg.auth == "trusted":
        raw = f"DRIVER={{{cfg.driver}}};Server={cfg.server};Database={cfg.database};Trusted_Connection=yes;"
    elif cfg.auth == "activedirectoryintegrated":
        raw = (
            f"DRIVER={{{cfg.driver}}};Server={cfg.server};Database={cfg.database};"
            "Authentication=ActiveDirectoryIntegrated;"
        )
    else:
        raise ValueError(f"Unsupported auth mode '{cfg.auth}' for server {cfg.server}")

    return "mssql+pyodbc:///?odbc_connect=" + urllib.parse.quote_plus(raw)


def _connect_and_query(cfg: SqlServerConfig, query: str | sqlalchemy.sql.ClauseElement,
                        params: dict | None) -> pd.DataFrame:
    stmt = query if isinstance(query, sqlalchemy.sql.ClauseElement) else sqlalchemy.text(query)
    engine = sqlalchemy.create_engine(_build_connection_url(cfg))
    with engine.connect() as conn:
        return pd.read_sql(stmt, conn, params=params or {})


def _run_query(source_name: str, cfg: SqlServerConfig, query: str | sqlalchemy.sql.ClauseElement,
                params: dict | None = None, max_attempts: int = 2) -> pd.DataFrame:
    last_exc: Exception | None = None
    for attempt in range(1, max_attempts + 1):
        try:
            df = _connect_and_query(cfg, query, params)
            logger.info("Query against '%s' returned %d rows (attempt %d)", source_name, len(df), attempt)
            return df
        except Exception as exc:  # noqa: BLE001 - deliberately broad, wrapped below
            last_exc = exc
            logger.warning("Query against '%s' failed on attempt %d: %s", source_name, attempt, exc)
    raise DatabaseQueryError(source_name, last_exc)


def get_ice2_visits(cfg: Config) -> pd.DataFrame:
    return _run_query(
        "ICe2",
        cfg.ice2,
        ICE2_QUERY_TEMPLATE,
        params={"start_date": cfg.ice2_query_start_date},
    )


ICE2_STATUS_LOOKUP_TEMPLATE = """
SELECT v.VisitID, v.VisitStatusID
FROM OM_Visit v
WHERE v.VisitID IN :visit_ids
"""
# No join to OM_Site_Contract/OM_Client_Contract here deliberately - this only confirms whether
# ICe2 has any record of these specific VisitIDs and what status they're currently in, not
# re-scoping which visits count as in-business the way ICE2_QUERY_TEMPLATE does.


# SQL Server/ODBC caps a single query at ~2,100 parameters. orphan_candidate_visit_ids can
# return thousands of VisitIDs (each becomes its own bound parameter via expanding=True), so
# this lookup must be chunked to stay well under that limit.
MAX_IN_CLAUSE_PARAMS = 2000


def get_ice2_status_by_visit_ids(cfg: Config, visit_ids: list[str]) -> pd.DataFrame:
    """Targeted follow-up lookup for reconcile.orphan_candidate_visit_ids - resolves whether ICe2
    has any record of a Hubscape-only VisitID at all, and under what current status, regardless
    of ICE2_QUERY_TEMPLATE's status/date filter."""
    if not visit_ids:
        return pd.DataFrame(columns=["VisitID", "VisitStatusID"])
    query = sqlalchemy.text(ICE2_STATUS_LOOKUP_TEMPLATE).bindparams(
        sqlalchemy.bindparam("visit_ids", expanding=True)
    )
    chunks = [
        _run_query(
            "ICe2 (orphan status lookup)",
            cfg.ice2,
            query,
            params={"visit_ids": visit_ids[i:i + MAX_IN_CLAUSE_PARAMS]},
        )
        for i in range(0, len(visit_ids), MAX_IN_CLAUSE_PARAMS)
    ]
    return pd.concat(chunks, ignore_index=True)


def _refresh_excel_workbook(excel_cfg: ExcelSourceConfig) -> None:
    """Drives Excel via COM to open the live-connection workbook and block until its data
    connection(s) finish refreshing, then saves and closes it.

    Runs the actual COM work on a daemon thread bounded by excel_cfg.refresh_timeout_seconds,
    the same pattern the old SQL path used to bound a stuck interactive sign-in: an unattended
    overnight run must fail fast rather than hang if Excel gets stuck (e.g. a stray dialog).
    """
    import pythoncom  # noqa: PLC0415
    import win32com.client  # noqa: PLC0415

    box: dict = {}

    def _target():
        pythoncom.CoInitialize()
        excel_app = None
        try:
            # DispatchEx starts a fresh, invisible Excel instance rather than attaching to (and
            # potentially disrupting) a copy the user already has open interactively.
            excel_app = win32com.client.DispatchEx("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.AskToUpdateLinks = False

            workbook = excel_app.Workbooks.Open(str(excel_cfg.path), UpdateLinks=0, ReadOnly=False)
            try:
                for connection in workbook.Connections:
                    try:
                        connection.OLEDBConnection.BackgroundQuery = False
                    except Exception:  # noqa: BLE001 - not every connection is OLEDB-backed
                        pass
                workbook.RefreshAll()
                # Blocks until every async query kicked off by RefreshAll actually completes,
                # rather than returning immediately with the workbook's stale cached data.
                excel_app.CalculateUntilAsyncQueriesDone()
                workbook.Save()
            finally:
                workbook.Close(SaveChanges=False)
        except Exception as inner_exc:  # noqa: BLE001
            box["exc"] = inner_exc
        finally:
            if excel_app is not None:
                excel_app.Quit()
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=_target, daemon=True)
    thread.start()
    thread.join(timeout=excel_cfg.refresh_timeout_seconds)
    if thread.is_alive():
        raise DatabaseQueryError(
            "Visits API (Excel)",
            TimeoutError(
                f"Refreshing '{excel_cfg.path}' did not complete within "
                f"{excel_cfg.refresh_timeout_seconds:.0f}s."
            ),
        )
    if "exc" in box:
        raise DatabaseQueryError("Visits API (Excel)", box["exc"])


def _read_excel_table(excel_cfg: ExcelSourceConfig, sheet_name: str | None = None) -> pd.DataFrame:
    workbook = openpyxl.load_workbook(excel_cfg.path, data_only=True, read_only=True)
    try:
        worksheet = workbook[sheet_name or excel_cfg.sheet_name]
        rows = worksheet.iter_rows(values_only=True)
        header = list(next(rows))
        return pd.DataFrame(rows, columns=header)
    finally:
        workbook.close()


def get_visits_api_from_excel(cfg: Config) -> pd.DataFrame:
    """Refreshes the live-connection Excel workbook (which stays authenticated against the
    Azure SQL Visits API via a human's occasional manual use, unlike an unattended direct SQL
    connection) and returns its refreshed table as a DataFrame."""
    excel_cfg = cfg.visits_api_excel
    try:
        _refresh_excel_workbook(excel_cfg)
        df = _read_excel_table(excel_cfg)
    except DatabaseQueryError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise DatabaseQueryError("Visits API (Excel)", exc) from exc
    logger.info("Visits API Excel extract ('%s') returned %d rows", excel_cfg.path, len(df))
    return df

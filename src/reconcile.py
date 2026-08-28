"""Pure pandas 3-way reconciliation of ICe2, Visits API, and Hubscape visit data.

No I/O here by design - fully unit-testable against synthetic DataFrames.
"""
from __future__ import annotations

from datetime import date

import pandas as pd

from .config import Config

# Alert-worthy issue types, in the order they should be presented in the summary email.
ALERT_ISSUE_TYPES = [
    "ICE2_MISSING_API_MAPPING",
    "ICE2_API_STATUS_MISMATCH",
    "MISSING_FROM_HUBSCAPE",
    "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2",
]
# Hubscape's rollout was gradual, so an old visit missing from Hubscape is expected/low
# priority, not a defect - it's counted but kept out of the alert-worthy total.
INFORMATIONAL_ISSUE_TYPES = [
    "LEGITIMATELY_ABSENT_FROM_HUBSCAPE", "MISSING_FROM_HUBSCAPE_HISTORIC", "ORPHAN_ICE2_STATUS_EXCLUDED",
]
ALL_ISSUE_TYPES = ALERT_ISSUE_TYPES + INFORMATIONAL_ISSUE_TYPES

# Directly-computed (never derived by subtraction) counts proving every ICe2 row and every
# Hubscape row lands in exactly one bucket - see ReconciliationCompletenessError below.
RECONCILIATION_ACCOUNTING_TYPES = ["ICE2_MATCHED_OK", "HUBSCAPE_MATCHED", "HUBSCAPE_MISSING_API_ID"]


class ReconciliationCompletenessError(Exception):
    """Raised when the classified buckets don't sum to the source row counts - i.e. some row
    went unclassified, was double-classified, or was silently dropped by a merge. Deliberately
    fatal (propagates through main.py's failure-email path) rather than a warning, since a wrong
    bucket count with no error is worse than a failed run."""


def _normalize_id(value):
    if pd.isna(value):
        return None
    try:
        as_float = float(value)
    except (TypeError, ValueError):
        return str(value)
    return str(int(as_float)) if as_float.is_integer() else str(value)


def _to_nullable_str_id(series: pd.Series) -> pd.Series:
    # API_ID arrives as float/object (with NaNs) from ICe2 but int64 (no NaNs) from the
    # Visits API; normalize both to strings, preserving NaN, so the merge key types match.
    return series.map(_normalize_id)


def _merge_and_classify_base(ice2_df: pd.DataFrame, api_df: pd.DataFrame, hub_df: pd.DataFrame, cfg: Config):
    """Shared ICe2/Visits-API/Hubscape merge, ID normalization, and orphan-candidate detection.

    Used by both `orphan_candidate_visit_ids` (called before the optional orphan-status
    follow-up query) and `reconcile` (called after it) so the two never drift apart.
    """
    ice2 = ice2_df.copy()
    ice2["VisitID"] = ice2["VisitID"].astype(str).str.strip()

    api = api_df.copy()
    api["VisitID"] = api["VisitID"].astype(str).str.strip()
    api["API_ID"] = _to_nullable_str_id(api["API_ID"])
    status_title_col = cfg.visits_api_status_title_column
    if status_title_col not in api.columns:
        raise KeyError(
            f"Expected Visits API status-label column '{status_title_col}' not found. "
            f"Columns present: {list(api.columns)}. "
            "Update sql.visits_api.status_title_column in config.yaml."
        )
    api = api.rename(columns={status_title_col: "API_StatusTitle"})

    hub = hub_df.copy()
    hub["API_ID"] = _to_nullable_str_id(hub["API_ID"])

    # --- Step 1: ICe2 vs Visits API, joined on VisitID ---
    # API_ID is resolved from the Visits API's own data (ultimately Entitymapping.SourceID ->
    # API_ID, baked into the Query1 extract) rather than from ICe2's own, potentially-stale
    # GCIS_API_Record_Mapping copy - so a match here means the API itself knows this VisitID.
    merged = ice2.merge(
        api[["API_ID", "VisitID", "API_StatusTitle"]],
        on="VisitID",
        how="left",
    )

    merged["issue_type"] = None
    merged.loc[merged["API_ID"].isna(), "issue_type"] = "ICE2_MISSING_API_MAPPING"

    status_mismatch_mask = (
        merged["API_ID"].notna()
        & (merged["StatusDescription"].astype(str) != merged["API_StatusTitle"].astype(str))
    )
    merged.loc[status_mismatch_mask, "issue_type"] = "ICE2_API_STATUS_MISMATCH"

    # --- Step 2: which rows should legitimately appear in Hubscape ---
    merged["expected_in_hubscape"] = ~merged["VisitStatusID"].isin(cfg.legitimately_excluded_from_hubscape)

    # --- Step 3: join against Hubscape extract on API_ID ---
    # Hubscape only ever sees the API_ID it received via the Visits API - it has no knowledge
    # of ICe2's own VisitID - so presence/absence must be checked via API_ID, not VisitID.
    # Rows with no API_ID at all (ICE2_MISSING_API_MAPPING) can never be matched against
    # Hubscape either way, but that's fine: they're already flagged and skipped below via
    # the `issue_type.isna()` guard.
    hub_api_ids = set(hub["API_ID"].dropna())
    merged["in_hubscape"] = merged["API_ID"].isin(hub_api_ids)

    # --- Orphan candidates: an API_ID in Hubscape that ICe2's (status/date-filtered) extract
    # has no mapping to at all. Uses merged["API_ID"] (resolved post-merge), not ice2's own data
    # - ICe2 no longer carries an API_ID column of its own.
    known_api_ids = set(merged["API_ID"].dropna())
    orphan_ids = hub_api_ids - known_api_ids

    return merged, api, hub, hub_api_ids, orphan_ids


def orphan_candidate_visit_ids(ice2_df: pd.DataFrame, api_df: pd.DataFrame, hub_df: pd.DataFrame, cfg: Config) -> list[str]:
    """VisitIDs (per the Visits API extract) behind Hubscape-only API_IDs that ICe2's
    status/date-filtered extract excluded.

    NOT all such orphan API_IDs are "unknown to ICe2" - most are visits whose current
    VisitStatusID (or ExpectedStartDate) has since moved outside ICE2_QUERY_TEMPLATE's WHERE
    filter, while Hubscape's export hasn't caught up yet. This returns the VisitIDs worth a
    follow-up ICe2 query (via `query_databases.get_ice2_status_by_visit_ids`) to tell the two
    cases apart before calling `reconcile`. Orphan API_IDs with no VisitID at all in the Visits
    API extract have nothing to look up and are always genuine orphans.
    """
    _, api, _, _, orphan_ids = _merge_and_classify_base(ice2_df, api_df, hub_df, cfg)
    resolvable = api[api["API_ID"].isin(orphan_ids)]
    return sorted(set(resolvable["VisitID"].dropna()))


def classify_team_type(internal_contractor_flags: pd.Series, cfg: Config) -> pd.Series:
    """Classifies each visit as "DE" (directly-employed), "Subcontractor", or "Unknown" from
    ft.internalcontractor (1/0), sourced authoritatively from ICe2's tblContractor rather than
    inferred from FieldTeamName. A missing flag (no Om_Job_Labour_Used match for the visit) is
    Unknown.
    """
    def _classify(flag) -> str:
        if pd.isna(flag):
            return "Unknown"
        return "DE" if int(flag) == 1 else "Subcontractor"

    return internal_contractor_flags.map(_classify)


def missing_from_hubscape_breakdown(
    ice2_df: pd.DataFrame,
    api_df: pd.DataFrame,
    hub_df: pd.DataFrame,
    cfg: Config,
) -> pd.DataFrame:
    """Year x team-type counts for every visit expected in Hubscape but not found there.

    Covers MISSING_FROM_HUBSCAPE and MISSING_FROM_HUBSCAPE_HISTORIC together (i.e. every row
    where expected_in_hubscape & ~in_hubscape & issue_type.isna() in the shared merge, ignoring
    reconcile()'s recency split deliberately) - a by-year view needs the full historic
    population, not just the last recent_window_days. year is ExpectedStartDate's calendar year.
    Subcontractor visits are expected to be ~100% ingested into Hubscape, so a persistent
    Subcontractor share here would be a genuine defect, unlike the expected DE gap.

    Returns columns: year, team_type, count.
    """
    merged, _, _, _, _ = _merge_and_classify_base(ice2_df, api_df, hub_df, cfg)
    missing_mask = merged["expected_in_hubscape"] & ~merged["in_hubscape"] & merged["issue_type"].isna()
    missing = merged[missing_mask].copy()

    missing["year"] = pd.to_datetime(missing["ExpectedStartDate"], errors="coerce").dt.year
    missing["team_type"] = classify_team_type(missing["internalcontractor"], cfg)

    breakdown = (
        missing.groupby(["year", "team_type"]).size()
        .reset_index(name="count")
        .sort_values(["year", "team_type"])
        .reset_index(drop=True)
    )
    return breakdown


def reconcile(
    ice2_df: pd.DataFrame,
    api_df: pd.DataFrame,
    hub_df: pd.DataFrame,
    cfg: Config,
    reference_date: date | None = None,
    orphan_status_lookup: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Returns (summary_df, detail_df).

    summary_df: one row per issue_type with a count column, plus source row-count rows.
    detail_df: one row per flagged VisitID with its issue_type and key context columns.

    orphan_status_lookup: optional (VisitID, VisitStatusID) DataFrame from
    `query_databases.get_ice2_status_by_visit_ids`, covering the VisitIDs returned by
    `orphan_candidate_visit_ids` for this same input. When supplied, an orphan candidate whose
    VisitID resolves here is reclassified from ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2 (implying ICe2 has
    never heard of it) to the informational ORPHAN_ICE2_STATUS_EXCLUDED (ICe2 knows it, it's just
    excluded from the extract by its current status/date - most likely Completed/Cancelled/
    On-Hold and Hubscape hasn't synced yet). When omitted, every orphan candidate is classified
    as ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2, matching this function's pre-split behavior.
    """
    merged, api, hub, hub_api_ids, orphan_ids = _merge_and_classify_base(ice2_df, api_df, hub_df, cfg)

    # The merge is left-join-on-VisitID from ice2_df; it must not fan out or drop rows, or every
    # bucket count below would be silently wrong. A duplicate VisitID appearing in the Visits API
    # extract is the most likely cause of a fan-out.
    if len(merged) != len(ice2_df):
        raise ReconciliationCompletenessError(
            f"ICe2 merge row count changed: {len(ice2_df)} input rows became {len(merged)} after "
            "merging against the Visits API extract on VisitID. Check for duplicate VisitIDs in "
            "the Visits API extract."
        )

    # Split "missing from Hubscape" by recency: Hubscape's integration rolled out gradually,
    # so an old visit that was never synced is expected/low-priority, while a *recent* one
    # missing is a genuine, actionable gap. Recency is based on ExpectedEndDate, falling back
    # to ExpectedStartDate, compared to a rolling window ending at reference_date (today).
    ref_date = pd.Timestamp(reference_date) if reference_date else pd.Timestamp.today().normalize()
    cutoff = ref_date - pd.Timedelta(days=cfg.recent_window_days)
    recency_date = pd.to_datetime(merged["ExpectedEndDate"], errors="coerce").fillna(
        pd.to_datetime(merged["ExpectedStartDate"], errors="coerce")
    )
    is_recent = recency_date >= cutoff

    missing_mask = merged["expected_in_hubscape"] & ~merged["in_hubscape"] & merged["issue_type"].isna()
    merged.loc[missing_mask & is_recent, "issue_type"] = "MISSING_FROM_HUBSCAPE"
    merged.loc[missing_mask & ~is_recent, "issue_type"] = "MISSING_FROM_HUBSCAPE_HISTORIC"

    legit_absent_mask = ~merged["expected_in_hubscape"] & ~merged["in_hubscape"] & merged["issue_type"].isna()
    merged.loc[legit_absent_mask, "issue_type"] = "LEGITIMATELY_ABSENT_FROM_HUBSCAPE"

    # Every row still unclassified at this point was checked by every mask above and tripped
    # none of them - i.e. it's fully matched (mapped, status matches, present/absent as expected).
    # Counted directly (not derived by subtraction) so a future gap in the masks above shows up
    # as a failed invariant below instead of silently changing this number.
    ice2_matched_ok = int(merged["issue_type"].isna().sum())

    # --- Step 4: classify orphans ---
    # An orphan candidate whose VisitID resolves in orphan_status_lookup is known to ICe2 - it's
    # just excluded from the (status/date-filtered) extract, most likely because its status has
    # since moved to Completed/Cancelled/On-Hold and Hubscape hasn't synced yet. Only a candidate
    # that resolves to no VisitID at all, or whose VisitID isn't found even in this broader
    # lookup, is genuinely unknown to ICe2.
    orphans = hub[hub["API_ID"].isin(orphan_ids)].copy()
    orphans["issue_type"] = "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2"
    if orphan_status_lookup is not None and not orphan_status_lookup.empty:
        api_id_to_visit_id = dict(zip(api["API_ID"], api["VisitID"]))
        lookup = orphan_status_lookup.copy()
        lookup["VisitID"] = lookup["VisitID"].astype(str).str.strip()
        status_by_visit_id = dict(zip(lookup["VisitID"], lookup["VisitStatusID"]))

        orphans["VisitID"] = orphans["API_ID"].map(api_id_to_visit_id)
        orphans["VisitStatusID"] = orphans["VisitID"].map(status_by_visit_id)
        resolved_mask = orphans["VisitStatusID"].notna()
        orphans.loc[resolved_mask, "issue_type"] = "ORPHAN_ICE2_STATUS_EXCLUDED"

    # --- Hubscape-side accounting, computed straight from hub (not from merged) so it's an
    # independent check rather than a re-derivation of the same numbers ---
    # Hubscape rows with no API_ID at all were previously invisible: excluded from hub_api_ids,
    # so they never appeared as "visible", "orphan", or anywhere else. Counted explicitly now.
    hubscape_missing_api_id = int(hub["API_ID"].isna().sum())
    hubscape_matched = int((hub["API_ID"].notna() & ~hub["API_ID"].isin(orphan_ids)).sum())

    ice2_side_total = (
        int((merged["issue_type"] == "ICE2_MISSING_API_MAPPING").sum())
        + int((merged["issue_type"] == "ICE2_API_STATUS_MISMATCH").sum())
        + int((merged["issue_type"] == "MISSING_FROM_HUBSCAPE").sum())
        + int((merged["issue_type"] == "MISSING_FROM_HUBSCAPE_HISTORIC").sum())
        + int((merged["issue_type"] == "LEGITIMATELY_ABSENT_FROM_HUBSCAPE").sum())
        + ice2_matched_ok
    )
    if ice2_side_total != len(ice2_df):
        raise ReconciliationCompletenessError(
            f"ICe2-side buckets sum to {ice2_side_total}, expected {len(ice2_df)} "
            f"(TOTAL_ICE2_ROWS). Some row was left unclassified or double-classified."
        )

    hubscape_orphan_total = int((orphans["issue_type"] == "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2").sum()) + int(
        (orphans["issue_type"] == "ORPHAN_ICE2_STATUS_EXCLUDED").sum()
    )
    hubscape_side_total = hubscape_matched + hubscape_missing_api_id + hubscape_orphan_total
    if hubscape_side_total != len(hub):
        raise ReconciliationCompletenessError(
            f"Hubscape-side buckets sum to {hubscape_side_total}, expected {len(hub)} "
            f"(TOTAL_HUBSCAPE_ROWS). Some row was left unclassified or double-classified."
        )

    # --- Step 5: assemble detail_df ---
    # Informational issue types (LEGITIMATELY_ABSENT_FROM_HUBSCAPE, MISSING_FROM_HUBSCAPE_HISTORIC,
    # ORPHAN_ICE2_STATUS_EXCLUDED) can be huge - they're counted in summary_df but excluded from
    # the row-level detail_df so it stays small and actionable.
    flagged = merged[merged["issue_type"].isin(ALERT_ISSUE_TYPES)].copy()
    orphans_flagged = orphans[orphans["issue_type"].isin(ALERT_ISSUE_TYPES)].copy()
    informational_counts = {
        issue_type: int((merged["issue_type"] == issue_type).sum()) + int((orphans["issue_type"] == issue_type).sum())
        for issue_type in INFORMATIONAL_ISSUE_TYPES
    }

    # Surface ft.internalcontractor as a clean nullable bool (True/False/blank) rather than the
    # raw 1/0/NaN from SQL, so subcontractor visits can be filtered directly in detail.csv.
    if "internalcontractor" in flagged.columns:
        flagged["internalcontractor"] = flagged["internalcontractor"].map(
            lambda v: None if pd.isna(v) else bool(int(v))
        )

    detail_cols = [
        "VisitID", "issue_type", "VisitStatusID", "StatusDescription",
        "API_ID", "API_StatusTitle", "ExpectedStartDate", "ExpectedEndDate",
        "ContractorID", "FieldTeamName", "internalcontractor",
    ]
    detail_cols = [c for c in detail_cols if c in flagged.columns]
    detail_df = pd.concat(
        [flagged[detail_cols], orphans_flagged.reindex(columns=detail_cols)],
        ignore_index=True,
    )

    # --- Step 6: summary_df ---
    counts = detail_df["issue_type"].value_counts().reindex(ALERT_ISSUE_TYPES, fill_value=0)
    summary_rows = [{"metric": issue_type, "count": int(count)} for issue_type, count in counts.items()]
    for issue_type, count in informational_counts.items():
        summary_rows.append({"metric": issue_type, "count": count})
    summary_rows.append({"metric": "ICE2_MATCHED_OK", "count": ice2_matched_ok})
    summary_rows.append({"metric": "HUBSCAPE_MATCHED", "count": hubscape_matched})
    summary_rows.append({"metric": "HUBSCAPE_MISSING_API_ID", "count": hubscape_missing_api_id})
    summary_rows.append({"metric": "TOTAL_ICE2_ROWS", "count": len(ice2_df)})
    summary_rows.append({"metric": "TOTAL_VISITS_API_ROWS", "count": len(api)})
    summary_rows.append({"metric": "TOTAL_HUBSCAPE_ROWS", "count": len(hub)})
    summary_df = pd.DataFrame(summary_rows)

    return summary_df, detail_df

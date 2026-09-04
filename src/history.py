"""Reads prior runs' saved output so today's run can compare against them.

The only place in this project that reads a *previous* day's output - reconcile.py is
deliberately I/O-free (see its module docstring), so this stays separate rather than blurring
that boundary.
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import pandas as pd


def find_previous_detail_csv(output_dir: Path, run_date: date, max_days_back: int = 7) -> Path | None:
    """Most recent run_date-N folder (N=1..max_days_back) containing a detail.csv, or None.

    Searches backward day by day rather than requiring exactly run_date - 1, so a day the
    pipeline didn't run at all (failure, Outlook unreachable) doesn't lose track of an ID that's
    genuinely still stuck - the Visits API's own nightly retry runs regardless of whether this
    reconciliation script ran that day.
    """
    for days_back in range(1, max_days_back + 1):
        candidate = output_dir / (run_date - timedelta(days=days_back)).isoformat() / "detail.csv"
        if candidate.exists():
            return candidate
    return None


def load_missing_api_mapping_ids(detail_csv_path: Path | None) -> set[str]:
    """VisitIDs flagged ICE2_MISSING_API_MAPPING in a previously saved detail.csv.

    VisitID is read as a plain string column (not inferred) so it isn't coerced through float64
    the way an unguarded read could - the same failure mode _to_nullable_str_id in reconcile.py
    guards against on the live extracts. Returns an empty set if the path is None, the file is
    missing/empty, or lacks the expected columns (never raises - a missing prior day just means
    nothing to compare against).
    """
    if detail_csv_path is None or not detail_csv_path.exists():
        return set()
    try:
        previous = pd.read_csv(detail_csv_path, dtype={"VisitID": str})
    except pd.errors.EmptyDataError:
        return set()
    if "VisitID" not in previous.columns or "issue_type" not in previous.columns:
        return set()
    missing = previous.loc[previous["issue_type"] == "ICE2_MISSING_API_MAPPING", "VisitID"]
    return set(missing.dropna())

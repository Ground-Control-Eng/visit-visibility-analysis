import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.history import find_previous_detail_csv, load_missing_api_mapping_ids


def _write_detail_csv(output_dir: Path, day: date, rows: list[dict]) -> Path:
    folder = output_dir / day.isoformat()
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "detail.csv"
    pd.DataFrame(rows).to_csv(path, index=False)
    return path


def test_find_previous_detail_csv_finds_immediate_prior_day(tmp_path):
    run_date = date(2026, 9, 4)
    expected = _write_detail_csv(tmp_path, date(2026, 9, 3), [{"VisitID": "1", "issue_type": "ICE2_MISSING_API_MAPPING"}])

    found = find_previous_detail_csv(tmp_path, run_date)

    assert found == expected


def test_find_previous_detail_csv_skips_gap_from_a_missed_run(tmp_path):
    # No output for 2026-09-03 (e.g. the pipeline failed that day) - must still find 2026-09-02
    # rather than giving up, so a genuinely stuck ID doesn't lose its repeat-failure tracking.
    run_date = date(2026, 9, 4)
    expected = _write_detail_csv(tmp_path, date(2026, 9, 2), [{"VisitID": "1", "issue_type": "ICE2_MISSING_API_MAPPING"}])

    found = find_previous_detail_csv(tmp_path, run_date)

    assert found == expected


def test_find_previous_detail_csv_returns_none_when_nothing_within_max_days_back(tmp_path):
    run_date = date(2026, 9, 4)
    _write_detail_csv(tmp_path, date(2026, 8, 1), [{"VisitID": "1", "issue_type": "ICE2_MISSING_API_MAPPING"}])

    found = find_previous_detail_csv(tmp_path, run_date, max_days_back=7)

    assert found is None


def test_load_missing_api_mapping_ids_filters_to_the_right_issue_type(tmp_path):
    path = _write_detail_csv(tmp_path, date(2026, 9, 3), [
        {"VisitID": "1", "issue_type": "ICE2_MISSING_API_MAPPING"},
        {"VisitID": "2", "issue_type": "ICE2_API_STATUS_MISMATCH"},
        {"VisitID": "3", "issue_type": "ICE2_MISSING_API_MAPPING"},
    ])

    ids = load_missing_api_mapping_ids(path)

    assert ids == {"1", "3"}


def test_load_missing_api_mapping_ids_preserves_visitid_as_string(tmp_path):
    # A prior run's detail.csv can carry numeric-looking VisitIDs - must not be coerced through
    # float64 on read (which would append ".0" and break the set comparison against today's
    # string-normalized VisitIDs from reconcile.py's _to_nullable_str_id).
    path = _write_detail_csv(tmp_path, date(2026, 9, 3), [
        {"VisitID": "12345", "issue_type": "ICE2_MISSING_API_MAPPING"},
    ])

    ids = load_missing_api_mapping_ids(path)

    assert ids == {"12345"}


def test_load_missing_api_mapping_ids_returns_empty_set_for_missing_path(tmp_path):
    ids = load_missing_api_mapping_ids(tmp_path / "does-not-exist" / "detail.csv")

    assert ids == set()


def test_load_missing_api_mapping_ids_returns_empty_set_for_none_path():
    assert load_missing_api_mapping_ids(None) == set()


def test_load_missing_api_mapping_ids_returns_empty_set_for_empty_file(tmp_path):
    path = tmp_path / "empty_detail.csv"
    path.write_text("", encoding="utf-8")

    assert load_missing_api_mapping_ids(path) == set()

import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.visualize import build_sankey_html


def make_summary_df():
    rows = {
        "ICE2_MISSING_API_MAPPING": 10,
        "ICE2_API_STATUS_MISMATCH": 5,
        "MISSING_FROM_HUBSCAPE": 20,
        "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2": 3,
        "LEGITIMATELY_ABSENT_FROM_HUBSCAPE": 100,
        "MISSING_FROM_HUBSCAPE_HISTORIC": 15,
        "ICE2_MATCHED_OK": 350,
        "HUBSCAPE_MATCHED": 347,
        "HUBSCAPE_MISSING_API_ID": 0,
        "TOTAL_ICE2_ROWS": 500,
        "TOTAL_VISITS_API_ROWS": 600,
        "TOTAL_HUBSCAPE_ROWS": 350,
    }
    return pd.DataFrame([{"metric": k, "count": v} for k, v in rows.items()])


def test_build_sankey_html_contains_expected_labels_and_values(tmp_path):
    summary_df = make_summary_df()
    path = build_sankey_html(summary_df, cfg=None, run_date=date(2026, 1, 1), output_dir=tmp_path)

    assert path.exists()
    content = path.read_text(encoding="utf-8")

    for label in ["ICe2 Visits", "Missing API Mapping",
                  "Status Mismatch", "Missing from Hubscape", "Visible in Hubscape",
                  "Untraced in ICe2"]:
        assert label in content

    # visible is read directly from ICE2_MATCHED_OK (350), +3 orphans = 353 in_hubscape_total
    assert "500" in content
    assert "353" in content  # in_hubscape_total shown in the table


def test_build_sankey_html_handles_zero_orphans(tmp_path):
    summary_df = make_summary_df()
    summary_df.loc[summary_df["metric"] == "ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2", "count"] = 0
    path = build_sankey_html(summary_df, cfg=None, run_date=date(2026, 1, 1), output_dir=tmp_path)
    content = path.read_text(encoding="utf-8")
    assert "Untraced in ICe2" not in content

import sys
from pathlib import Path

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config import ExcelSourceConfig
from src.query_databases import _read_excel_table


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

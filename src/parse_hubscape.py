"""Parses the Hubscape attachment (CSV or Excel) into a normalized DataFrame of active visits."""
from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from .config import Config

logger = logging.getLogger("visit_reconciliation")


class AttachmentFormatError(Exception):
    pass


def parse_attachment(path: Path, cfg: Config) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(path)
    elif suffix in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    else:
        raise AttachmentFormatError(
            f"Unsupported attachment type '{suffix}' for {path}. Expected .csv, .xlsx, or .xls."
        )

    api_id_col = cfg.hubscape_api_id_column
    if api_id_col not in df.columns:
        raise AttachmentFormatError(
            f"Expected column '{api_id_col}' not found in Hubscape attachment. "
            f"Columns present: {list(df.columns)}. "
            "Update hubscape_columns.api_id_column in config.yaml to match the real export."
        )

    # Hubscape only ever exposes the API_ID it received via the Visits API - it has no
    # knowledge of ICe2's own VisitID.
    df = df.rename(columns={api_id_col: "API_ID"})
    # A blank API_ID cell must stay real NaN so reconcile.py's HUBSCAPE_MISSING_API_ID accounting
    # can count it - a bare astype(str) risks stringifying it to the literal "nan" on some pandas
    # versions, which pd.isna()/.dropna() downstream would no longer recognize as missing.
    api_id = df["API_ID"]
    df["API_ID"] = api_id.where(api_id.isna(), api_id.astype(str).str.strip())
    logger.info("Parsed %d active visit rows from Hubscape attachment %s", len(df), path)
    return df

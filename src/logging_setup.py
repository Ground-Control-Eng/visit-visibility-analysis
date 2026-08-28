"""Rotating file + console logger, one log file per run under output/YYYY-MM-DD/."""
from __future__ import annotations

import logging
from pathlib import Path


def setup_logging(run_output_dir: Path) -> logging.Logger:
    run_output_dir.mkdir(parents=True, exist_ok=True)
    log_path = run_output_dir / "run_log.txt"

    logger = logging.getLogger("visit_reconciliation")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)
    file_handler.setLevel(logging.DEBUG)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt)
    console_handler.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger

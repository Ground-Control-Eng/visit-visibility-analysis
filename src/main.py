"""Orchestrates the daily visit reconciliation pipeline.

Every stage is wrapped so a failure anywhere still results in a diagnostic
failure email (or, if Outlook itself is unreachable, at least a local log entry)
rather than the run silently doing nothing.

Usage:
    python -m src.main
    python -m src.main --dry-run-email-path samples/some_saved_attachment.csv
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from . import fetch_email, parse_hubscape, query_databases, reconcile, send_summary, visualize
from .config import Config, ConfigError, load_config
from .logging_setup import setup_logging


def _cleanup_old_output(cfg: Config, logger) -> None:
    cutoff_days = cfg.run.keep_days
    if not cfg.run.output_dir.exists():
        return
    now = datetime.now()
    for child in cfg.run.output_dir.iterdir():
        if not child.is_dir():
            continue
        try:
            folder_date = datetime.strptime(child.name, "%Y-%m-%d")
        except ValueError:
            continue
        if (now - folder_date).days > cutoff_days:
            logger.info("Cleaning up old output folder %s (older than %d days)", child, cutoff_days)
            for f in child.glob("*"):
                f.unlink(missing_ok=True)
            child.rmdir()


def run(dry_run_email_path: Path | None = None) -> int:
    run_date = date.today()
    try:
        cfg = load_config()
    except ConfigError as exc:
        print(f"CONFIG ERROR: {exc}", file=sys.stderr)
        return 2

    run_output_dir = cfg.run.output_dir / run_date.isoformat()
    logger = setup_logging(run_output_dir)
    logger.info("Starting daily visit reconciliation run for %s (test_mode=%s)", run_date, cfg.run.test_mode)

    stage = "startup"
    try:
        # --- Stage: fetch + parse Hubscape data ---
        stage = "fetch_email"
        if dry_run_email_path is not None:
            attachment_path = dry_run_email_path
            logger.info("Dry-run mode: using saved attachment at %s instead of live Outlook search.", attachment_path)
        else:
            mail_item = fetch_email.find_todays_hubscape_email(cfg)
            stage = "extract_attachment"
            attachment_path = fetch_email.extract_attachment(mail_item, cfg, run_output_dir)

        stage = "parse_hubscape"
        hub_df = parse_hubscape.parse_attachment(attachment_path, cfg)
        hub_df.to_csv(run_output_dir / "hubscape.csv", index=False)

        # --- Stage: query databases ---
        # Raw source dumps (ICe2/Visits API) are large (100MB+) - only persisted in test_mode
        # for debugging, to avoid ballooning this OneDrive-synced output/ folder on every
        # production run. summary.csv and detail.csv (the actionable artifacts) are always kept.
        stage = "query_ice2"
        ice2_df = query_databases.get_ice2_visits(cfg)
        if cfg.run.test_mode:
            ice2_df.to_csv(run_output_dir / "ice2.csv", index=False)

        stage = "query_visits_api"
        api_df = query_databases.get_visits_api_from_excel(cfg)
        if cfg.run.test_mode:
            api_df.to_csv(run_output_dir / "visits_api.csv", index=False)

        # --- Stage: reconcile ---
        stage = "reconcile"
        orphan_visit_ids = reconcile.orphan_candidate_visit_ids(ice2_df, api_df, hub_df, cfg)
        orphan_status_df = query_databases.get_ice2_status_by_visit_ids(cfg, orphan_visit_ids)
        summary_df, detail_df = reconcile.reconcile(
            ice2_df, api_df, hub_df, cfg, orphan_status_lookup=orphan_status_df,
        )
        summary_df.to_csv(run_output_dir / "summary.csv", index=False)
        detail_df.to_csv(run_output_dir / "detail.csv", index=False)
        logger.info("Reconciliation complete: %s", summary_df.to_dict(orient="records"))

        stage = "missing_breakdown"
        breakdown_df = reconcile.missing_from_hubscape_breakdown(ice2_df, api_df, hub_df, cfg)
        breakdown_df.to_csv(run_output_dir / "missing_from_hubscape_by_year_and_team.csv", index=False)
        logger.info("Missing-from-Hubscape breakdown: %s", breakdown_df.to_dict(orient="records"))

        stage = "visualize"
        sankey_path = visualize.build_sankey_html(summary_df, cfg, run_date, run_output_dir)
        logger.info("Wrote Sankey visualization to %s", sankey_path)

        # --- Stage: send summary ---
        stage = "send_summary"
        send_summary.send_success_email(summary_df, detail_df, cfg, run_date, run_output_dir, breakdown_df=breakdown_df)

        _cleanup_old_output(cfg, logger)
        logger.info("Run completed successfully.")
        return 0

    except Exception as exc:  # noqa: BLE001 - top-level guard, must not let anything escape silently
        logger.exception("Run failed at stage '%s'", stage)
        try:
            send_summary.send_failure_email(stage, exc, cfg, run_date, run_output_dir)
        except Exception:  # noqa: BLE001
            logger.exception("Additionally failed to send/write the failure email.")
        return 1


def main() -> None:
    parser = argparse.ArgumentParser(description="Daily Hubscape/Visits API/ICe2 visit reconciliation.")
    parser.add_argument(
        "--dry-run-email-path",
        type=Path,
        default=None,
        help="Path to a saved Hubscape attachment; skips live Outlook search and uses this file instead.",
    )
    args = parser.parse_args()
    sys.exit(run(dry_run_email_path=args.dry_run_email_path))


if __name__ == "__main__":
    main()

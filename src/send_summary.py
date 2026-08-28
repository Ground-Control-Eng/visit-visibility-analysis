"""Builds and sends the daily summary/failure email via Outlook COM."""
from __future__ import annotations

import logging
from datetime import date
from pathlib import Path

import pandas as pd

from .config import Config
from .reconcile import ALERT_ISSUE_TYPES, INFORMATIONAL_ISSUE_TYPES, RECONCILIATION_ACCOUNTING_TYPES

logger = logging.getLogger("visit_reconciliation")


def _df_to_html_table(df: pd.DataFrame, max_rows: int | None = None) -> str:
    if df.empty:
        return "<p><em>No rows.</em></p>"
    shown = df.head(max_rows) if max_rows else df
    html = shown.to_html(index=False, border=1)
    if max_rows and len(df) > max_rows:
        html += f"<p><em>Showing first {max_rows} of {len(df)} rows - see attached CSV for full detail.</em></p>"
    return html


def _build_breakdown_html(breakdown_df: pd.DataFrame) -> str:
    if breakdown_df.empty:
        return "<p><em>No visits missing from Hubscape.</em></p>"
    pivoted = breakdown_df.pivot(index="year", columns="team_type", values="count").fillna(0).astype(int)
    pivoted["Total"] = pivoted.sum(axis=1)
    pivoted.loc["Total"] = pivoted.sum(axis=0)
    return pivoted.to_html(border=1)


def _build_success_html(summary_df: pd.DataFrame, detail_df: pd.DataFrame, run_date: date,
                         extra_notes: list[str], breakdown_df: pd.DataFrame | None = None) -> str:
    counts = dict(zip(summary_df["metric"], summary_df["count"]))
    total_alerts = sum(counts.get(t, 0) for t in ALERT_ISSUE_TYPES)

    rows_html = "".join(
        f"<tr><td>{t}</td><td style='text-align:right'>{counts.get(t, 0)}</td></tr>"
        for t in ALERT_ISSUE_TYPES
    )
    info_rows_html = "".join(
        f"<tr><td>{t}</td><td style='text-align:right'>{counts.get(t, 0)}</td></tr>"
        for t in INFORMATIONAL_ISSUE_TYPES
    )
    source_rows_html = "".join(
        f"<tr><td>{t}</td><td style='text-align:right'>{counts.get(t, 0)}</td></tr>"
        for t in ("TOTAL_ICE2_ROWS", "TOTAL_VISITS_API_ROWS", "TOTAL_HUBSCAPE_ROWS")
    )
    accounting_rows_html = "".join(
        f"<tr><td>{t}</td><td style='text-align:right'>{counts.get(t, 0)}</td></tr>"
        for t in RECONCILIATION_ACCOUNTING_TYPES
    )

    notes_html = ""
    if extra_notes:
        notes_html = "<ul>" + "".join(f"<li>{n}</li>" for n in extra_notes) + "</ul>"

    alert_detail = detail_df[detail_df["issue_type"].isin(ALERT_ISSUE_TYPES)]
    breakdown_html = _build_breakdown_html(breakdown_df) if breakdown_df is not None else ""
    breakdown_section = (
        f"<h3>Missing from Hubscape - by year and team type</h3>{breakdown_html}"
        if breakdown_df is not None else ""
    )

    return f"""
    <html><body>
    <h2>Visit Reconciliation Summary - {run_date.isoformat()}</h2>
    <p><b>{total_alerts}</b> total alert-worthy discrepancies found.</p>
    {notes_html}
    <h3>Discrepancies</h3>
    <table>{rows_html}</table>
    <h3>Informational (expected absences)</h3>
    <table>{info_rows_html}</table>
    <h3>Source row counts (sanity check)</h3>
    <table>{source_rows_html}</table>
    <h3>Full accounting (every ICe2/Hubscape row is in exactly one bucket, enforced)</h3>
    <table>{accounting_rows_html}</table>
    {breakdown_section}
    <h3>Discrepancy detail (first 200 rows)</h3>
    {_df_to_html_table(alert_detail, max_rows=200)}
    <p>Full detail and logs are attached / saved to the output folder for this run.</p>
    </body></html>
    """


def _build_failure_html(stage: str, exception: Exception, run_date: date, log_path: Path) -> str:
    return f"""
    <html><body>
    <h2>Visit Reconciliation FAILED - {run_date.isoformat()}</h2>
    <p><b>Stage:</b> {stage}</p>
    <p><b>Error:</b> {type(exception).__name__}: {exception}</p>
    <p>See the full log for this run at: {log_path}</p>
    </body></html>
    """


def _send_via_outlook(subject: str, html_body: str, cfg: Config, attachments: list[Path]) -> None:
    import win32com.client  # noqa: PLC0415

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = "; ".join(cfg.email.to)
    if cfg.email.cc:
        mail.CC = "; ".join(cfg.email.cc)
    mail.Subject = subject
    mail.HTMLBody = html_body
    for attachment_path in attachments:
        if attachment_path.exists():
            mail.Attachments.Add(str(attachment_path))
    mail.Send()
    logger.info("Sent email '%s' to %s", subject, cfg.email.to)


def send_success_email(summary_df: pd.DataFrame, detail_df: pd.DataFrame, cfg: Config,
                        run_date: date, run_output_dir: Path, extra_notes: list[str] | None = None,
                        breakdown_df: pd.DataFrame | None = None) -> None:
    html = _build_success_html(summary_df, detail_df, run_date, extra_notes or [], breakdown_df=breakdown_df)
    subject = f"Visit Reconciliation Summary - {run_date.isoformat()}"
    attachments = [run_output_dir / "summary.csv", run_output_dir / "detail.csv", run_output_dir / "sankey.html"]
    if breakdown_df is not None:
        attachments.append(run_output_dir / "missing_from_hubscape_by_year_and_team.csv")

    if cfg.run.test_mode or not cfg.email.send_on_success:
        (run_output_dir / "would_be_email.html").write_text(html, encoding="utf-8")
        logger.info(
            "test_mode/send_on_success is off - wrote would-be email body to %s instead of sending.",
            run_output_dir / "would_be_email.html",
        )
        return

    _send_via_outlook(subject, html, cfg, attachments)


def send_failure_email(stage: str, exception: Exception, cfg: Config, run_date: date,
                        run_output_dir: Path) -> None:
    log_path = run_output_dir / "run_log.txt"
    html = _build_failure_html(stage, exception, run_date, log_path)
    subject = f"Visit Reconciliation FAILED ({stage}) - {run_date.isoformat()}"

    if cfg.run.test_mode or not cfg.email.send_on_failure:
        (run_output_dir / "would_be_failure_email.html").write_text(html, encoding="utf-8")
        logger.info(
            "test_mode/send_on_failure is off - wrote would-be failure email to %s instead of sending.",
            run_output_dir / "would_be_failure_email.html",
        )
        return

    try:
        _send_via_outlook(subject, html, cfg, attachments=[log_path])
    except Exception:  # noqa: BLE001
        logger.exception(
            "Could not send failure email via Outlook COM (Outlook may be unreachable). "
            "Failure details are in the run log at %s.",
            log_path,
        )

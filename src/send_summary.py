"""Builds and sends the daily summary/failure email via Outlook COM."""
from __future__ import annotations

import logging
import time
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from .config import Config
from .reconcile import ALERT_ISSUE_TYPES, INFORMATIONAL_ISSUE_TYPES, RECONCILIATION_ACCOUNTING_TYPES

logger = logging.getLogger("visit_reconciliation")

_OUTLOOK_FOLDER_OUTBOX = 4  # olFolderOutbox

EMAIL_SUBJECT_PREFIX = "Visit Reconciliation"

# Matches COLOR_CRITICAL in visualize.py - kept as a local literal rather than an import so the
# two modules stay decoupled.
REPEAT_FAILURE_ROW_STYLE = "background-color:#fbe1e1;font-weight:bold"


def _df_to_html_table(df: pd.DataFrame, max_rows: int | None = None) -> str:
    if df.empty:
        return "<p><em>No rows.</em></p>"
    shown = df.head(max_rows) if max_rows else df
    html = shown.to_html(index=False, border=1)
    if max_rows and len(df) > max_rows:
        html += f"<p><em>Showing first {max_rows} of {len(df)} rows - see attached CSV for full detail.</em></p>"
    return html


def _alert_detail_html(df: pd.DataFrame, max_rows: int | None = None) -> str:
    """Like _df_to_html_table, but bold/red-highlights rows where repeat_failure is True - a
    VisitID that's still ICE2_MISSING_API_MAPPING after also being missing in the prior run
    (i.e. the API's overnight retry failed for it again)."""
    if df.empty:
        return "<p><em>No rows.</em></p>"
    shown = df.head(max_rows) if max_rows else df
    display_cols = [c for c in shown.columns if c != "repeat_failure"]
    has_repeat_flag = "repeat_failure" in shown.columns

    header = "".join(f"<th>{col}</th>" for col in display_cols)
    body_rows = []
    for _, row in shown.iterrows():
        is_repeat = has_repeat_flag and bool(row["repeat_failure"])
        tr_style = f' style="{REPEAT_FAILURE_ROW_STYLE}"' if is_repeat else ""
        cells = "".join(f"<td>{row[col]}</td>" for col in display_cols)
        body_rows.append(f"<tr{tr_style}>{cells}</tr>")

    html = f"<table border=\"1\"><thead><tr>{header}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"
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

    repeat_count = int(counts.get("ICE2_MISSING_API_MAPPING_REPEAT", 0))
    repeat_callout_html = ""
    if repeat_count and "repeat_failure" in detail_df.columns:
        repeat_ids = sorted(detail_df.loc[detail_df["repeat_failure"], "VisitID"].astype(str))
        repeat_callout_html = (
            f'<p style="{REPEAT_FAILURE_ROW_STYLE}">'
            f"&#9888; {repeat_count} VisitID(s) still missing an API mapping after also being "
            f"missing in the prior run - the overnight retry has failed for these two days "
            f"running: {', '.join(repeat_ids)}</p>"
        )

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
    {repeat_callout_html}
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
    {_alert_detail_html(alert_detail, max_rows=200)}
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


def is_stale_pipeline_outbox_item(
    subject: str, created_at: datetime, now: datetime, threshold_seconds: float,
) -> bool:
    """Whether an Outbox item is this pipeline's and old enough to be an orphan from a
    previous run's stuck send, rather than a normal in-flight one.

    created_at and now must both be naive or both tz-aware (mismatched inputs raise
    TypeError on subtraction) - callers passing Outlook COM's tz-aware CreationTime should
    strip its tzinfo first, as src/fetch_email.py already does for ReceivedTime.
    """
    if not subject.startswith(EMAIL_SUBJECT_PREFIX):
        return False
    return (now - created_at).total_seconds() >= threshold_seconds


def _outbox_entry_ids(outbox) -> set[str]:
    """EntryIDs of all readable Outbox items - skips any item mid-transmission (touching one
    raises 'Outlook has already begun transmitting this message', which isn't an actual error,
    just a sign the item is actively going out)."""
    ids = set()
    for item in outbox.Items:
        try:
            ids.add(item.EntryID)
        except Exception:  # noqa: BLE001
            continue
    return ids


def _cleanup_stale_outbox_items(outbox, threshold_seconds: float) -> None:
    """Deletes leftover Outbox items from a previous failed run so they don't sit there
    indefinitely requiring manual cleanup before the next send can get through."""
    # Snapshot into a plain list first - deleting while iterating Outlook's live COM
    # collection re-indexes it mid-loop and silently skips entries.
    for item in list(outbox.Items):
        try:
            subject = item.Subject
            created_at = item.CreationTime
        except Exception:  # noqa: BLE001
            continue
        created_at = created_at.replace(tzinfo=None) if created_at.tzinfo else created_at
        now = datetime.now()
        if not is_stale_pipeline_outbox_item(subject, created_at, now, threshold_seconds):
            continue

        age_seconds = (now - created_at).total_seconds()
        logger.warning(
            "Deleting stale Outbox item left over from a previous run: subject=%r, age=%.0fs "
            "(threshold=%.0fs). It will not be sent - see that day's output folder for the "
            "original report content if it's still needed.",
            subject, age_seconds, threshold_seconds,
        )
        try:
            item.Delete()
        except Exception:  # noqa: BLE001
            logger.exception("Failed to delete stale Outbox item with subject=%r.", subject)


def _send_via_outlook(subject: str, html_body: str, cfg: Config, attachments: list[Path]) -> None:
    import win32com.client  # noqa: PLC0415

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if cfg.email.outlook_profile:
        namespace.Logon(cfg.email.outlook_profile)
    if namespace.Offline:
        raise RuntimeError(
            f"Outlook is in Work Offline mode - cannot send '{subject}'. Switch it to online "
            "(Send/Receive tab) and re-run."
        )

    outbox = namespace.GetDefaultFolder(_OUTLOOK_FOLDER_OUTBOX)
    _cleanup_stale_outbox_items(outbox, cfg.email.stale_outbox_cleanup_seconds)

    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = "; ".join(cfg.email.to)
    if cfg.email.cc:
        mail.CC = "; ".join(cfg.email.cc)
    mail.Subject = subject
    mail.HTMLBody = html_body
    for attachment_path in attachments:
        if attachment_path.exists():
            mail.Attachments.Add(str(attachment_path))

    # mail.Send() only queues the item into the local Outbox - it does not confirm the message
    # was actually transmitted, and doesn't raise if Outlook's connection to Exchange has
    # silently gone stale. Snapshot the Outbox's EntryIDs before/after Send() to identify the
    # freshly-queued copy (subject text alone isn't reliable - stale stuck items from past runs
    # can share the same subject), then poll for it to actually leave.
    entry_ids_before = _outbox_entry_ids(outbox)
    mail.Send()
    new_entry_ids = _outbox_entry_ids(outbox) - entry_ids_before

    if not new_entry_ids:
        logger.info("Confirmed '%s' left the Outbox immediately (transmitted) to %s", subject, cfg.email.to)
        return

    namespace.SendAndReceive(False)  # force transmission now rather than waiting on Outlook's timer
    timeout = cfg.email.send_confirm_timeout_seconds
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        still_stuck = new_entry_ids & _outbox_entry_ids(outbox)
        if not still_stuck:
            logger.info("Confirmed '%s' left the Outbox (transmitted) to %s", subject, cfg.email.to)
            return
        time.sleep(2)

    raise RuntimeError(
        f"Email '{subject}' is still sitting in the Outbox after {timeout:.0f}s - Outlook accepted "
        "it but never transmitted it. Check Outlook is actually connected/signed in to Exchange "
        "(e.g. no new mail arriving in the Inbox is a red flag, even if it doesn't report 'offline')."
    )


def send_success_email(summary_df: pd.DataFrame, detail_df: pd.DataFrame, cfg: Config,
                        run_date: date, run_output_dir: Path, extra_notes: list[str] | None = None,
                        breakdown_df: pd.DataFrame | None = None) -> None:
    html = _build_success_html(summary_df, detail_df, run_date, extra_notes or [], breakdown_df=breakdown_df)
    subject = f"{EMAIL_SUBJECT_PREFIX} Summary - {run_date.isoformat()}"
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
    subject = f"{EMAIL_SUBJECT_PREFIX} FAILED ({stage}) - {run_date.isoformat()}"

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

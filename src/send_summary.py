"""Builds and sends the daily summary/failure email via Outlook COM."""
from __future__ import annotations

import logging
import time
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from . import notify_teams
from .config import Config
from .reconcile import ALERT_ISSUE_TYPES, INFORMATIONAL_ISSUE_TYPES, RECONCILIATION_ACCOUNTING_TYPES

logger = logging.getLogger("visit_reconciliation")

_OUTLOOK_FOLDER_OUTBOX = 4  # olFolderOutbox
_OUTLOOK_FOLDER_INBOX = 6  # olFolderInbox

# OlExchangeConnectionMode values that unambiguously mean "not actually connected to Exchange"
# (per Microsoft's OlExchangeConnectionMode enum). namespace.Offline only ever catches the two
# "Offline" rows below - it stays False for a stale/disconnected Modern Auth session, which is
# exactly the recurring failure mode this guards against (see README "Known limitation").
_EXCHANGE_DISCONNECTED_MODES = {
    0: "olNoExchange - this profile isn't on an Exchange account",
    100: "olOffline - classic offline mode",
    200: "olCachedOffline - Work Offline is selected",
    300: "olDisconnected - disconnected from the Exchange server",
    400: "olCachedDisconnected - cached mode, disconnected from the Exchange server",
}

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


def _exchange_connection_problem(connection_mode: int) -> str | None:
    """None if connection_mode looks healthy; otherwise a human-readable diagnosis. Note: a
    value not in _EXCHANGE_DISCONNECTED_MODES ("looks connected") does NOT guarantee Outlook is
    actually syncing - a stalled cached-mode session can still report e.g.
    olCachedConnectedFull - see _send_via_outlook's Outbox-poll fallback for that case."""
    return _EXCHANGE_DISCONNECTED_MODES.get(connection_mode)


def _read_exchange_connection_mode(namespace) -> int | None:
    """Best-effort read of namespace.ExchangeConnectionMode - None if the COM property access
    itself fails (e.g. a transient COM error), so a read failure degrades to 'inconclusive'
    rather than raising a secondary exception that would mask the real diagnosis and skip the
    Event Log alert."""
    try:
        return namespace.ExchangeConnectionMode
    except Exception:  # noqa: BLE001
        logger.warning("Could not read namespace.ExchangeConnectionMode.", exc_info=True)
        return None


def _inbox_last_received(namespace) -> datetime | None:
    """Best-effort read of the most recent Inbox ReceivedTime, or None on any failure (e.g. an
    empty folder). An independent corroborating signal: a stale Exchange session can still
    report a 'connected'-looking ExchangeConnectionMode while no new mail is actually
    arriving."""
    try:
        inbox = namespace.GetDefaultFolder(_OUTLOOK_FOLDER_INBOX)
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        latest = items.GetFirst()
        if latest is None:
            return None
        received_at = latest.ReceivedTime
        return received_at.replace(tzinfo=None) if received_at.tzinfo else received_at
    except Exception:  # noqa: BLE001
        return None


def _alert_via_event_log(message: str) -> None:
    """Best-effort write to the Windows Application Event Log - a channel independent of
    Outlook, so it still fires when Outlook itself is what's broken (see README "Known
    limitation"). Requires the "Visit Reconciliation" event source to have been registered once
    (scheduled_task/register_task.ps1, elevated) - if that hasn't been done, or the write fails
    for any other reason, this just logs locally rather than masking the original error."""
    try:
        import win32evtlog  # noqa: PLC0415
        import win32evtlogutil  # noqa: PLC0415

        win32evtlogutil.ReportEvent(
            "Visit Reconciliation", 1, eventType=win32evtlog.EVENTLOG_ERROR_TYPE, strings=[message],
        )
    except Exception:  # noqa: BLE001
        logger.exception("Could not write alert to the Windows Event Log either.")


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


def _attempt_reconnect(namespace, wait_seconds: float) -> int | None:
    """Best-effort nudge to clear a stale Exchange connection before giving up: re-issues
    Namespace.Logon (a no-op if a MAPI session is already active - harmless either way) and
    forces a Send/Receive, which is what actually prompts Outlook's own identity module to
    either silently refresh an expiring token (the "transient blip" case - see README "Known
    limitation") or surface its own sign-in UI for whoever's at the machine to complete. Neither
    outcome is guaranteed - no script can complete an MFA challenge on a person's behalf - so
    this never raises; callers re-check the returned connection mode themselves."""
    try:
        namespace.Logon("", "", True, False)
    except Exception:  # noqa: BLE001
        logger.debug("Namespace.Logon reconnect nudge failed (often harmless if already logged on).", exc_info=True)
    try:
        namespace.SendAndReceive(True)
    except Exception:  # noqa: BLE001
        logger.debug("SendAndReceive reconnect nudge failed.", exc_info=True)

    deadline = time.monotonic() + wait_seconds
    connection_mode = _read_exchange_connection_mode(namespace)
    while (
        connection_mode is not None
        and _exchange_connection_problem(connection_mode) is not None
        and time.monotonic() < deadline
    ):
        time.sleep(2)
        connection_mode = _read_exchange_connection_mode(namespace)
    return connection_mode


def _send_via_outlook(subject: str, html_body: str, cfg: Config, attachments: list[Path]) -> None:
    import win32com.client  # noqa: PLC0415

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if cfg.email.outlook_profile:
        namespace.Logon(cfg.email.outlook_profile)

    connection_mode = _read_exchange_connection_mode(namespace)
    problem = _exchange_connection_problem(connection_mode) if connection_mode is not None else None
    if problem:
        logger.warning(
            "ExchangeConnectionMode=%s (%s) before sending '%s' - attempting a reconnect nudge "
            "(Logon + forced Send/Receive) before giving up.",
            connection_mode, problem, subject,
        )
        connection_mode = _attempt_reconnect(namespace, cfg.email.send_confirm_timeout_seconds)
        problem = _exchange_connection_problem(connection_mode) if connection_mode is not None else None

    if problem:
        diagnosis = (
            f"Outlook is not connected to Exchange - cannot send '{subject}'. "
            f"ExchangeConnectionMode={connection_mode} ({problem}). A reconnect nudge (forced "
            "Send/Receive) didn't clear it - sign in / reconnect Outlook (Send/Receive tab, or "
            "close and reopen it) and re-run."
        )
        _alert_via_event_log(diagnosis)
        raise RuntimeError(diagnosis)

    outbox = namespace.GetDefaultFolder(_OUTLOOK_FOLDER_OUTBOX)
    _cleanup_stale_outbox_items(outbox, cfg.email.stale_outbox_cleanup_seconds)
    inbox_before = _inbox_last_received(namespace)

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

    if _confirm_outbox_transmit(namespace, outbox, subject, new_entry_ids, cfg):
        return

    total_attempts = cfg.email.send_confirm_retries + 1
    inbox_after = _inbox_last_received(namespace)
    connection_mode_after = _read_exchange_connection_mode(namespace)
    if connection_mode_after is None:
        mode_str = "unknown"
        problem_after = "could not be determined - reading ExchangeConnectionMode failed"
    else:
        mode_str = str(connection_mode_after)
        problem_after = _exchange_connection_problem(connection_mode_after) or "reports connected"
    diagnosis = (
        f"Email '{subject}' is still sitting in the Outbox after {total_attempts} attempt(s) of "
        f"{cfg.email.send_confirm_timeout_seconds:.0f}s each - Outlook accepted it but never "
        f"transmitted it. ExchangeConnectionMode={mode_str} ({problem_after}). Last Inbox mail "
        f"received at {inbox_after} (was {inbox_before} before this send attempt) - if that "
        "hasn't advanced, Outlook is not actually syncing with Exchange despite not reporting "
        "itself offline; an interactive re-sign-in is required."
    )
    _alert_via_event_log(diagnosis)
    raise RuntimeError(diagnosis)


def _confirm_outbox_transmit(namespace, outbox, subject: str, new_entry_ids: set[str], cfg: Config) -> bool:
    """Polls for new_entry_ids to leave the Outbox, retrying the wait (never a second Send() -
    so there's no risk of a duplicate email) up to cfg.email.send_confirm_retries extra times if
    it's still stuck. A transient Exchange blip can self-recover within seconds - observed
    2026-09-09: a failure-alert for this exact timeout transmitted instantly moments later - so
    it's worth riding out a couple of retries before giving up and alerting."""
    attempts = cfg.email.send_confirm_retries + 1
    for attempt in range(1, attempts + 1):
        namespace.SendAndReceive(False)  # force transmission now rather than waiting on Outlook's timer
        deadline = time.monotonic() + cfg.email.send_confirm_timeout_seconds
        while time.monotonic() < deadline:
            still_stuck = new_entry_ids & _outbox_entry_ids(outbox)
            if not still_stuck:
                logger.info(
                    "Confirmed '%s' left the Outbox (transmitted) to %s%s",
                    subject, cfg.email.to, f" on attempt {attempt}/{attempts}" if attempt > 1 else "",
                )
                return True
            time.sleep(2)
        if attempt < attempts:
            logger.warning(
                "'%s' still in Outbox after %.0fs (attempt %d/%d) - retrying the wait rather than "
                "giving up immediately.",
                subject, cfg.email.send_confirm_timeout_seconds, attempt, attempts,
            )
    return False


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
        teams_message = (
            f"Stage: {stage}\nError: {type(exception).__name__}: {exception}\n"
            f"Outlook could not send either the original alert or this failure notification - "
            f"see {log_path} for full details."
        )
        notify_teams.send_teams_alert(subject, teams_message, cfg)

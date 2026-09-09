# Daily Hubscape / Visits API / ICe2 Visit Reconciliation

Automates the daily reconciliation across ICe2 (source of truth), the Visits API (nightly
mapped copy), and Hubscape (what delivery teams see, active visits only).

## Setup (one-time)

```powershell
pip install -r requirements.txt
```

The ICe2 connection (`GCV-PROD-SQL01`) uses Windows-integrated auth and needs no setup.

The Visits API data (Azure SQL, `sql-gc-services-prod.database.windows.net`) is **not** queried
directly. A direct `ActiveDirectoryInteractive` connection needed a fresh browser/MFA sign-in
roughly once a day under this tenant's conditional access policy, which an unattended Scheduled
Task run can't complete on its own. Instead, `API Status.xlsx` (in the project root) has a live
Power Query connection to the same database, kept authenticated by ordinary occasional manual use
of the workbook. Each run drives Excel via COM (`pywin32`, already a dependency) to open that
workbook invisibly, refresh its connection, and read the refreshed table - see
`sql.visits_api_excel` in `config.yaml` for the path/sheet/table settings.

This means:
- Excel must be installed on the machine running the Scheduled Task.
- `API Status.xlsx` must stay at the configured path (or `config.yaml` updated if it moves), and
  should periodically be opened/refreshed by a person so its cached Azure AD session stays warm.
- It must not be left open/locked by another process (e.g. someone editing it by hand) when the
  scheduled run fires, or the automated refresh will fail.
- If the refresh doesn't complete within `sql.visits_api_excel.refresh_timeout_seconds` (default
  180s), the run fails with a diagnostic error instead of hanging indefinitely.

## Configuration

All settings live in `config.yaml` - no code changes needed to retune:
- SQL server/database names and auth mode
- Which `VisitStatusID`s are legitimately excluded from Hubscape's active-only export
- The Outlook search filter for the daily trigger email (sender, subject, folder, lookback)
- Recipients for the summary email
- `run.test_mode`: when `true`, the pipeline runs for real (real SQL, real Outlook search)
  but writes the would-be email to `output/YYYY-MM-DD/` instead of sending it, and also
  saves the raw ICe2/Visits API extracts there for debugging. Set to `false` once you're
  confident in the output, before relying on the live scheduled run.
- DE (directly-employed) vs Subcontractor classification comes straight from ICe2's
  `ft.internalcontractor` flag (`1` = DE, `0` = Subcontractor) in `ICE2_QUERY_TEMPLATE`.
  A visit whose team couldn't be resolved (no `Om_Job_Labour_Used` match) is classified `Unknown`.
- `sql.ice2.exclude_de_teams`: DE teams are known not to be ingested into Hubscape yet, so by
  default (`true`) their visits are omitted from the ICe2 extract entirely - they won't appear in
  `detail.csv`, `summary.csv`, or `missing_from_hubscape_by_year_and_team.csv`. Set to `false`
  once Hubscape starts ingesting DE team visits, to resume tracking them like Subcontractor
  visits.
- `sql.ice2.excluded_contractor_ids`: specific Field Team `ContractorID`s known not to be ingested
  into Hubscape yet - same rationale and same omit-from-the-extract-entirely treatment as
  `sql.ice2.exclude_de_teams` above, just for individual contractors rather than the whole DE population.
- A visit excluded by either setting above that turns out to *already have a matching visit in
  Hubscape* (i.e. the "not yet ingested" assumption doesn't fully hold for it) isn't silently
  dropped: it's counted informationally as `ORPHAN_ICE2_TEAM_EXCLUDED` in `summary.csv`/the email,
  kept distinct from `ORPHAN_ICE2_STATUS_EXCLUDED` (a visit ICe2 knows about that's just currently
  outside the date/status window - most likely Completed/Cancelled/On-Hold, Hubscape not yet
  synced). Neither is row-level in `detail.csv`.

The attachment column name/format (`External Visit API Id`) has been confirmed against real
Hubscape exports. The ICe2 query's `ExpectedStartDate` cutoff is controlled by
`sql.ice2.query_start_date` in `config.yaml` (currently `2025-10-01`) and can be retuned there
without a code change.

## Running manually

```powershell
python -m src.main
```

To test against a saved sample attachment instead of searching Outlook live:

```powershell
python -m src.main --dry-run-email-path samples\some_saved_attachment.csv
```

## Running tests

```powershell
python -m pytest tests\
```

`tests/test_reconcile.py` covers every reconciliation issue type against synthetic data -
no database or Outlook connection required.

## Output

Each run writes to `output/YYYY-MM-DD/`:
- `summary.csv` - counts per issue type (always written)
- `detail.csv` - one row per flagged VisitID for the alert-worthy issue types (always written;
  the informational `LEGITIMATELY_ABSENT_FROM_HUBSCAPE` count is in `summary.csv` only, since
  it covers every historically completed/cancelled visit and would otherwise dominate the file)
- `missing_from_hubscape_by_year_and_team.csv` - year x team-type (DE / Subcontractor / Unknown)
  counts of every visit missing from Hubscape (recent and historic combined), to catch any
  Subcontractor visits slipping through (which shouldn't happen). While `sql.ice2.exclude_de_teams`
  is `true` (the default), DE visits are excluded upstream in the ICe2 extract itself, so the DE
  column here will read zero/near-zero - this is expected, not a sign the gap closed
- `detail.csv`'s `repeat_failure` column flags an `ICE2_MISSING_API_MAPPING` VisitID that was
  *also* missing in the most recent prior run (searching back up to 7 days, to ride out a day the
  pipeline didn't run) - i.e. the Visits API's overnight retry failed to map it again. These rows
  are bold/red-highlighted in the summary email's discrepancy table, with a callout near the top
  listing the affected VisitIDs, and counted in `summary.csv` as `ICE2_MISSING_API_MAPPING_REPEAT`.
- `summary.csv`'s two orphan-adjacent counts (a Hubscape API_ID with no matching ICe2 extract row)
  are deliberately split: `ORPHAN_ICE2_STATUS_EXCLUDED` (known to ICe2, just outside the current
  date/status window) vs `ORPHAN_ICE2_TEAM_EXCLUDED` (would have been excluded by
  `excluded_contractor_ids`/`exclude_de_teams` policy despite already being in Hubscape) - see the
  Configuration section above
- `run_log.txt` - full log for this run
- `hubscape.csv` - parsed Hubscape extract
- `ice2.csv` / `visits_api.csv` - raw source dumps, only written when `test_mode: true`
  (they're 100MB+ each and this folder is OneDrive-synced, so they're skipped in production
  runs to avoid needless sync/storage load)
- `would_be_email.html` / `would_be_failure_email.html` - only written when the corresponding
  real email wasn't sent (test mode, or `send_on_success`/`send_on_failure` off)

Folders older than `run.keep_days` (default 90) are cleaned up automatically on each run.

## Scheduling the daily run

Once you've manually reviewed a few days of `test_mode: true` output and are happy with it:

1. Set `run.test_mode: false` in `config.yaml`.
2. Register the Scheduled Task (run from an elevated PowerShell if prompted):
   ```powershell
   scheduled_task\register_task.ps1
   ```
   This registers a daily 21:15 task (the Hubscape email typically arrives ~20:30) running
   as your Windows user, with "run only when logged on" - required because the pipeline
   drives Outlook via COM, which needs an interactive desktop session. A locked (but
   logged-in) session should be fine; a fully logged-off session will not run the task.

To change the run time: `scheduled_task\register_task.ps1 -RunTime "22:00"`.

## Known limitation

If Outlook itself is unreachable when a run fails (not signed in, or closed), the script
can't send its own failure-alert email through Outlook. In that case the failure is still
recorded in `output/YYYY-MM-DD/run_log.txt` - check there if a day's summary email doesn't
arrive at all.

A second, easier-to-miss failure mode: `mail.Send()` only queues a message into Outlook's
local Outbox - it does **not** confirm the message was actually transmitted, and it doesn't
raise an exception even if Outlook's connection to Exchange has silently gone stale (e.g. a
Modern Auth/MFA session that expired with no one present to re-authenticate it). Outlook can
sit in this state indefinitely without reporting itself as "offline" - the only visible sign
is that no new mail arrives in the Inbox either. Before this fix, that meant the log would
say `Sent email '...'` while the message sat unsent in the Outbox forever.

`_send_via_outlook()` (`src/send_summary.py`) now guards against this on two levels:

- Before attempting a send, it checks `Namespace.ExchangeConnectionMode` (not just
  `Namespace.Offline`, which only ever catches a manual "Work Offline" toggle and stays
  `False` for a stale/disconnected session). If the mode is unambiguously bad
  (`olNoExchange`/`olOffline`/`olCachedOffline`/`olDisconnected`/`olCachedDisconnected`), it
  fails immediately with that diagnosis instead of waiting out the full Outbox-poll timeout.
- Otherwise, it snapshots the Outbox before/after `Send()` to identify the freshly-queued
  item, forces an immediate Send/Receive, and polls for up to
  `email.send_confirm_timeout_seconds` (default 30s, `config.yaml`) for that item to actually
  leave the Outbox. A stuck-but-"connected"-looking session can be a transient blip that
  self-recovers within seconds (observed in practice: a failure-alert for this exact timeout
  transmitted instantly moments later), so this wait is retried up to
  `email.send_confirm_retries` extra times (default 2, i.e. 3 attempts total) before giving
  up - always re-polling the *same* queued item, never calling `Send()` again, so there's no
  risk of a duplicate email going out. Only once every attempt has timed out does the
  resulting error report `ExchangeConnectionMode` at that point and the Inbox's most-recent
  `ReceivedTime` before vs. after the send attempt - operationalizing "no new mail arriving is
  a red flag" into the error text itself, since a stalled cached-mode session can still report
  a connected-looking mode.

Either way the run now fails loudly (non-zero exit, logged error, a failure-email attempt)
instead of falsely logging success, and a diagnostic entry is written to the Windows
Application Event Log (source "Visit Reconciliation") - a channel independent of Outlook, so
the alert still lands even if the failure-notification email itself can't send either. This
requires the event source to have been registered once via an elevated
`scheduled_task\register_task.ps1` run; if that hasn't been done, the write is skipped
harmlessly (logged locally) rather than masking the original error.

If Outlook's connection has gone stale like this, it typically needs an interactive
re-sign-in (closing and reopening Outlook, or completing an MFA prompt) - something only a
person at the keyboard can do; there's no way around this in this scenario, so if the
pipeline keeps failing this way, check Event Viewer (Windows Logs > Application, source
"Visit Reconciliation") or `run_log.txt` for the diagnosis, and reconnect Outlook.

If a run *still* can't get a failure email out at all (Outlook fully unreachable), `main.py`
still returns a non-zero exit code - check the Scheduled Task's "Last Run Result" in Task
Scheduler as a backstop, alongside `run_log.txt` and the Event Log, if no email arrives and
you're not sure why.

On top of detecting a stuck send, every send attempt now also auto-deletes any leftover
Outbox item matching this pipeline's subject prefix ("Visit Reconciliation...") that's older
than `email.stale_outbox_cleanup_seconds` (default 1 hour, `config.yaml`), logging a warning
with its subject and age. This means a prior run's stuck item no longer has to be deleted by
hand before the next run can get its own email out - previously, if the connection recovered
on its own, the next scheduled run would still find the same stuck item sitting in the Outbox
and someone had to clear it manually before a retry could succeed. This only clears the
symptom (a blocked Outbox); if the underlying stale Exchange session keeps recurring, it
still typically needs the interactive re-sign-in described above.

Briefly being unable to read an Outbox item's properties (Outlook raises "already begun
transmitting this message") is expected right as a message starts sending, and is handled
gracefully - it's not itself a sign of a problem, just that the item is actively going out.

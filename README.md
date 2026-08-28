# Daily Hubscape / Visits API / ICe2 Visit Reconciliation

Automates the daily reconciliation across ICe2 (source of truth), the Visits API (nightly
mapped copy), and Hubscape (what delivery teams see, active visits only). See
`instructions.md` for the business background.

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
  `ft.internalcontractor` flag (`1` = DE, `0` = Subcontractor) in `ICE2_QUERY_TEMPLATE` -
  no config needed. A visit whose team couldn't be resolved (no `Om_Job_Labour_Used` match)
  is classified `Unknown`.

**Still to confirm once a real Hubscape email has been seen** (see `config.yaml` comments):
the exact attachment column name/format, and whether the ICe2 query's
`ExpectedStartDate >= 2024-04-01` cutoff should stay.

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
  counts of every visit missing from Hubscape (recent and historic combined), to track the known
  DE ingestion gap and catch any subcontractor visits slipping through (which shouldn't happen)
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

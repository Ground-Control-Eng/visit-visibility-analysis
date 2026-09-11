# Daily Hubscape / Visits API / ICe2 Visit Reconciliation

Automates the daily reconciliation across ICe2 (source of truth), the Visits API (nightly
mapped copy), and Hubscape (what delivery teams see, active visits only).

## Setup (one-time)

```powershell
pip install -r requirements.txt
```

`src/send_summary.py` and `src/fetch_email.py` send/fetch email via **Microsoft Graph API**
(app-only / client-credentials auth), not Outlook - no Outlook installation or interactive
sign-in is needed on the machine running this pipeline.

This needs an Entra ID app registration with **Application** (not Delegated) permissions
`Mail.Send` and `Mail.Read`. Verify with ICT that admin consent shows a green "Granted for
[tenant]" against **both** permissions specifically - it's easy for one of the two to be added
but not actually consented while still looking set up at a glance. Once you have one:
1. Fill in `graph.tenant_id` / `graph.client_id` / `graph.mailbox` in `config.yaml` (`mailbox`
   is the address the pipeline sends from and searches in - app-only auth has no "me").
2. Copy `.env.example` to `.env` (in the project root) and set `GRAPH_CLIENT_SECRET` to the
   app registration's client secret. `.env` is gitignored - never commit the real secret.

`pywin32` is still a dependency, but no longer for Outlook - it's kept for the Windows Event
Log alert (`src/send_summary.py`, `_alert_via_event_log`) and for the Excel-COM workaround
described below (`src/query_databases.py`), both unchanged by this.

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

Optional: a Microsoft Teams fallback alert (see "Known limitation" below) for when Graph can't
send the failure-alert email either. To set it up, in the target Teams channel: "..." > Workflows
> "Post to a channel when a webhook request is received" > complete the wizard > paste the URL it
gives you into `notifications.teams.webhook_url` in `config.yaml`. Leave it `null` to skip this -
it's optional and the run won't fail because it's unset.

## Configuration

All settings live in `config.yaml` - no code changes needed to retune:
- SQL server/database names and auth mode
- Which `VisitStatusID`s are legitimately excluded from Hubscape's active-only export
- The Graph search filter for the daily trigger email (sender, subject, folder, lookback)
- Recipients for the summary email
- `run.test_mode`: when `true`, the pipeline runs for real (real SQL, real Graph search)
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

To test against a saved sample attachment instead of searching live via Graph:

```powershell
python -m src.main --dry-run-email-path samples\some_saved_attachment.csv
```

## Running tests

```powershell
python -m pytest tests\
```

`tests/test_reconcile.py` covers every reconciliation issue type against synthetic data -
no database or Graph connection required.

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
   as your Windows user, with "run only when logged on". This was originally required because
   the pipeline drove Outlook via COM, which needed an interactive desktop session - that's now
   gone (migrated to Microsoft Graph, see Setup above, which needs no interactive session at
   all). `src/query_databases.py` still drives Excel via COM for the Visits API workaround
   (unchanged, out of scope for the Graph migration) and **may** have the same interactive-
   session requirement - this hasn't been verified empirically, so don't change `-LogonType`
   away from `Interactive` until that's confirmed safe. A locked (but logged-in) session should
   be fine either way; a fully logged-off session will not run the task.

To change the run time: `scheduled_task\register_task.ps1 -RunTime "22:00"`.

## Known limitation

Microsoft Graph's app-only (client-credentials) auth has no interactive user or session at
all - there's no desktop Outlook connection to go stale, so the entire class of "stuck in the
Outbox with a silently expired Exchange session" failures this section used to document
(reconnect nudges, `ExchangeConnectionMode` checks, Outbox-transmit polling) cannot occur here
and no longer exists in the code.

`_send_via_graph()` (`src/send_summary.py`) sends via a single synchronous `POST
/users/{mailbox}/sendMail` call - a 202 response means Graph has accepted the message for
delivery immediately, no local queue to get stuck in. Transient errors (HTTP 429, honoring
`Retry-After`, or 5xx) are retried automatically with backoff, up to `graph.max_retries`
(default 3) extra attempts (`src/graph_client.py`). A persistent failure raises, is logged to
`run_log.txt`, and a diagnostic entry is written to the Windows Application Event Log (source
"Visit Reconciliation") - a channel independent of Graph/email entirely, so the alert still
lands even if the failure-notification email itself can't send. This requires the event source
to have been registered once via an elevated `scheduled_task\register_task.ps1` run; if that
hasn't been done, the write is skipped harmlessly (logged locally) rather than masking the
original error.

If a run still can't get a failure email out at all (e.g. the app registration's credentials
are wrong or revoked), `send_failure_email()` additionally posts to a Teams channel webhook
(`notifications.teams.webhook_url` in `config.yaml` - see Setup above) via a plain HTTPS
request (`src/notify_teams.py`) with no Graph/Outlook involved at all - so there's still a
notification path when Graph itself is what's broken. This is a fallback only: it doesn't fire
on a normal successful run, and if the webhook isn't configured yet, it just logs a warning
rather than blocking anything. `main.py` also always returns a non-zero exit code on failure -
check the Scheduled Task's "Last Run Result" in Task Scheduler as a backstop, alongside
`run_log.txt` and the Event Log, if no email arrives and you're not sure why.

One real attachment-size caveat: `sendMail`'s inline (base64) attachments become unreliable
above a few MB combined. If the run's attachments (`summary.csv`, `detail.csv`, `sankey.html`,
etc.) together exceed `graph.max_inline_attachment_bytes` (default ~3MB), the lowest-priority
ones are dropped (logged as a warning in `run_log.txt`) rather than failing the whole send -
`sankey.html` is the most likely one to grow large enough to matter.

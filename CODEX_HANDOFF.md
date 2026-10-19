# Pac-Biz Dashboard v1 - Codex Handoff

Last updated: 2026-10-19

This file is the operating guide for Codex or another coding agent working on
the Pac-Biz Dashboard v1 project. Read `CLAUDE.md` and `PIPELINE_MONITOR.md`
as the detailed technical references; this file defines the expected working
behavior, commands, safeguards, and reporting format.

## Project Location

The project must live at:

```text
C:\Users\Mike Woo Cerna\Documents\PB\
```

Primary working directory:

```text
C:\Users\Mike Woo Cerna\Documents\PB\Masterlist
```

The root `PB` directory contains the supporting folders used by the pipeline:

- `Masterlist` - dashboard, monitor, pipeline controls, cache files, and Git repo.
- `Coaching` - Asana pull and Coaching output.
- `Quality` - one folder per QA account, source scripts, credentials, and RAW workbooks.

Use the same path on a replacement PC whenever possible. The batch files use
absolute `C:\Users\Mike Woo Cerna\Documents\PB\...` paths.

## Agent Behavior

- When Mike says **"update entire dashboard"**, execute the complete pipeline.
  Do not merely provide instructions or run a partial dashboard build.
- Be deliberate around email. Never send sample, retry, or duplicate movement
  notifications unless Mike explicitly requests it. The notification ledger is
  `movement_notified.json`.
- Do not alter the monitoring UI or dashboard UI unless Mike specifically asks.
- Do not alter or delete user-managed/untracked `.agents/` or `supabase/`.
- Do not use destructive Git commands (`reset --hard`, forced checkout, or
  deletion of local files) unless Mike explicitly asks.
- Never refer work to Claude or edit the external `.claude` directory unless
  Mike explicitly asks. This is a Codex-operated workflow.
- Commit and push only the files required by the requested change. Do not stage
  unrelated caches or untracked folders.

## Full Dashboard Update

Run this from `Masterlist`:

```powershell
& 'C:\Users\Mike Woo Cerna\Documents\PB\Masterlist\update_coaching_dashboard_auto.bat'
```

This is the definition of a full update:

1. Fetch Masterlist, History, and Movement data.
2. Check and send newly eligible movement notifications.
3. Pull Coaching from Asana.
4. Run all 22 QA sources in four controlled parallel lanes.
5. Enforce the QA validation barrier before building.
6. Run count-drop self-heal checks.
7. Build `masterlist_dashboard.html` from validated local workbooks.
8. Reconcile movement/cache data and overdue movement state.
9. Push dashboard and monitor updates to GitHub Pages.
10. Run the Pipeline Guardian final check.

The pipeline is successful only when its final monitor status is successful and
the validation barrier has passed. A GitHub Pages update normally appears about
one minute after the push.

### Partial Commands

Use only when Mike asks for a limited action:

```powershell
# Masterlist, History, and Movement cache fetch only
py -3 masterlist_fetch.py

# Newly processed movement notification check only
& 'C:\Users\Mike Woo Cerna\Documents\PB\Masterlist\update_movement_notifications_auto.bat'

# Build dashboard from existing cached/raw files; does not repull QA accounts
$env:SKIP_ACCOUNT_REFRESH='1'; py -3 dashboard.py
```

## Required Completion Report

After every full v1 Dashboard update, report all items below, using current
run data rather than copied historical values:

1. Pipeline: `24/24` sources passed, or exact failed/partial status.
2. QA rows: total count.
3. Masterlist: row count.
4. Latest History snapshot date.
5. History snapshot rows for that date.
6. History blank dates.
7. Coaching snapshot: row count.
8. Latest Coaching Date.
9. Latest Coaching Created/Modified Date.
10. Notifications: whether a newly processed movement was sent; name it when sent.
11. Whether there were significant count drops or overdue movements.

Mention material exceptions separately, such as PACE Coaching validation
alerts, a partial QA source, a GitHub publish failure, or an expected future-
dated Coaching entry.

### Snapshot Verification

The authoritative local inputs after a successful run are:

- `masterlist_cache.csv`
- `history_cache.csv`
- `movement_cache.csv`
- `C:\Users\Mike Woo Cerna\Documents\PB\Coaching\Output\coaching_logs.xlsx`
- `pipeline_status.json`

For History, count blank values in `Date Generated` and count rows matching
the latest valid date. For Coaching, report both the maximum `Coaching Date`
and the maximum `Created Date` / `Modified Date`; a Coaching Date can be a
valid future date and is not automatically a dashboard error.

## Movement Notifications

Notification code: `check_movement_notifications.py`.

- It reads `movement_cache.csv` and sends only unseen, already-processed rows.
- It records sent movement timestamps in `movement_notified.json` to prevent
  re-notification.
- Permanent CC recipients are:
  - `it-team@pac-biz.com`
  - `hr@pac-biz.com`
  - `gibson@pac-biz.com`
  - `carl.bandoles@pac-biz.com`
- BCC recipient: `reports@pac-biz.com`.
- Do not add `hadhassah@pac-biz.com`; it was intentionally removed.

Upstream movement processing is owned by the Masterlist Google Apps Script.
Python sends notification emails only after a Movement row is marked processed
in the published Movement data.

## Masterlist Google Apps Script and History Safety

Local source of truth for the deployed script:

```text
Masterlist\masterlist_appsscript_enhanced.gs
```

The Apps Script is cloud-hosted; editing the local `.gs` file alone does not
deploy it. When an Apps Script change is required, update the local source,
commit/push it, then have Mike paste it into the Google Apps Script editor,
save it, and run the requested function.

Important functions:

- `runMasterlistProcess()` - full Masterlist, Movement, and History routine.
- `runMovementProcessOnly()` - lightweight movement-only routine.
- `installMovementProcessingTrigger()` - creates the 15-minute movement trigger;
  run only once unless a trigger needs replacement.

History snapshots are now append-only by date. Never reintroduce a
delete-today-then-write pattern: a timeout after deletion can erase a daily
snapshot. The script first retains an existing snapshot for the date, and it
uses one History name lookup rather than rereading the entire History sheet
for every employee.

## Coaching Rules

- Coaching data is pulled by `C:\Users\Mike Woo Cerna\Documents\PB\Coaching\asana_pull.py`.
- PACE email validation notifications are deduplicated using their local state
  files. Do not resend them merely because the dashboard is rebuilt.
- Coaching report invalid-email entries are excluded until their source email is
  corrected. Report them as exceptions when newly detected.
- In reports about a future Coaching Date, state that it was created on the
  shown Created Date and request Operations to confirm the intended coaching
  date. Do not label it a dashboard error without evidence.

## QA Pipeline

There are 22 QA sources. The parallel runner is `parallel_qa_runner.py` and
uses four bounded lanes. Do not replace it with uncontrolled simultaneous
pulls.

The 22 accounts are:

```text
Skyline, Hamilton, VIP, Vermont, Trans Iowa,
Data Carz, Associated Cab, Ollies, Circle Taxi, YCOV,
M7, DMG, R4H, Parentis Health, Britelift, Britelift Chat, RideX,
C&H, Blueline, Reno Cab, Kelowna, YCDC
```

The validation barrier checks process success, workbook readability, nonzero
rows, standard columns, and protected row-count drops before the dashboard
can build. A source failure should fail closed and preserve a valid prior RAW
workbook when the safety rules allow it. Never overwrite a healthy RAW file
with partial, empty, unauthenticated, or schema-invalid data.

Several accounts now use direct QA API pulls rather than Google Apps Script,
including Hamilton, VIP, Reno Cab, Circle Taxi, Associated Cab, YCDC,
Vermont, Kelowna, Data Carz, Trans Iowa, and Skyline. Do not put API keys or
passwords into Markdown, Git, screenshots, or public files.

## Git, Monitor, and Publishing

- Repository: `https://github.com/MikeWooCerna/Pac-Biz.git`
- Branch: `main`
- Dashboard: `https://mikewoocerna.github.io/Pac-Biz/masterlist_dashboard.html`
- Monitor: `https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html`

The generated dashboard file is large and GitHub emits a 50 MB recommendation
warning. This warning alone is not a failed publish.

The Pipeline Guardian can warn because `.agents/` and `supabase/` are local
untracked folders. Do not delete them to silence the warning. Report the
pipeline's actual data and publish status separately from that housekeeping
warning.

## PC 2 Setup Checklist

1. Restore the full `PB` folder to `C:\Users\Mike Woo Cerna\Documents\PB\`.
2. Install Python 3 and Git, then run:

   ```powershell
   cd 'C:\Users\Mike Woo Cerna\Documents\PB\Masterlist'
   py -3 -m pip install -r requirements.txt
   ```

3. Ensure every copied `credentials.json` and `token.json` remains in its
   account folder. They are sensitive and must not be committed to Git.
4. Verify `py -3`, `git`, and GitHub authentication work from `Masterlist`.
5. Recreate Windows Task Scheduler jobs. They do not migrate with files.
6. Keep the Google Apps Script triggers; they are cloud-hosted and should not
   be duplicated on PC 2.
7. Run one full dashboard update and verify the required completion report.

If Git metadata was skipped during USB transfer, clone the GitHub repository
into the same `PB\Masterlist` location, then restore local-only assets,
credentials, tokens, and data caches from the USB copy.

## First-Run Checks on PC 2

Before the first production update:

```powershell
cd 'C:\Users\Mike Woo Cerna\Documents\PB\Masterlist'
py -3 -m py_compile masterlist_fetch.py check_movement_notifications.py dashboard.py
git status
```

Confirm that account source files and credentials exist, then run the full
batch. Do not send test movement emails; the notification ledger protects
against duplicates when copied from PC 1.

## Change Discipline

- Inspect surrounding code and existing behavior before changing logic.
- Keep changes scoped and add verification appropriate to the risk.
- Use `rg` for searches and `apply_patch` for manual edits.
- Before reporting a change as complete, run relevant validation.
- For code review, present defects and risks first, ordered by severity.
- When a request is ambiguous but could send email, modify production data, or
  erase data, confirm the intended scope before acting.

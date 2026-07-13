# Pipeline Monitor — Technical Reference

**Live URL:** https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html  
**Last updated:** 2026-07-11  
**Version signature:** `v26.06.22`

---

## What it is

A fully self-contained HTML dashboard that shows the real-time status of every step in the `update_coaching_dashboard_auto.bat` / `update_coaching_dashboard.bat` pipeline. Pushed to GitHub Pages automatically — both mid-run (live heartbeat after every step) and at pipeline finish.

---

## Files

| File | Purpose |
|------|---------|
| `generate_monitor.py` | Reads `pipeline_status.json` + Excel/CSV files → writes `pipeline_monitor.html`. Called by `log_step.py` after every step AND by both bat files at finish. |
| `log_step.py` | Called by bat files with `init` / `step` / `finish` commands. Writes to `pipeline_status.json`. After each `step`, calls `generate_monitor.py` and pushes to GitHub (live heartbeat). |
| `masterlist_fetch.py` | Fetches the 3 Masterlist Google Sheets CSVs and caches them locally as `masterlist_cache.csv`, `history_cache.csv`, `movement_cache.csv`. First step in both bat files. |
| `pipeline_status.json` | Runtime state written by `log_step.py`. Contains `run_id`, `started_at`, `finished_at`, `status`, `failed_at`, `steps[]`. Committed to git after every step. |
| `pipeline_log.json` | Persistent incident history. Appended to by `generate_monitor.py` when a run finishes. Survives across runs. Capped at 300 entries. |
| `pipeline_rowcount_baseline.json` | Last known row count per account (dict). Updated after every finished run. Used to detect drops between runs. |
| `pipeline_monitor.html` | Generated output. Self-contained, no external deps. Committed to git and served via GitHub Pages. Auto-reloads every 60s. |
| `masterlist_cache.csv` | Cached Masterlist sheet. Read by `dashboard.py` instead of fetching live. Falls back to live URL if file missing. Not committed to git. |
| `history_cache.csv` | Cached History sheet. Same pattern. Not committed to git. |
| `movement_cache.csv` | Cached Movement sheet. Same pattern. Not committed to git. |

---

## Data flow

```
bat file runs
  └─ log_step.py init
       └─ writes pipeline_status.json {status: "running"}

  for each account step:
    └─ py -3 script.py 2>step_err.tmp
    └─ log_step.py step "Account" "script.py" exit_code
         ├─ reads step_err.tmp for error message (on failure)
         ├─ appends step to pipeline_status.json
         └─ push_live():
              ├─ generate_monitor.py  →  pipeline_monitor.html
              ├─ git add pipeline_status.json pipeline_monitor.html
              ├─ git commit -m "[live] Account"
              └─ git push

  └─ log_step.py finish success|failed
       └─ sets pipeline_status.json {status: "success"|"failed", finished_at: ...}

  └─ generate_monitor.py
       ├─ reads pipeline_status.json + all Excel/CSV files
       ├─ compares row counts to pipeline_rowcount_baseline.json (drop detection)
       ├─ appends incidents to pipeline_log.json (failures, not-reached, count drops)
       ├─ updates pipeline_rowcount_baseline.json with current counts
       └─ writes pipeline_monitor.html

  └─ git add pipeline_status.json pipeline_monitor.html pipeline_log.json pipeline_rowcount_baseline.json
  └─ git commit + push
```

---

## Account order (ACCOUNTS list in generate_monitor.py)

Matches pipeline execution order. Masterlist is **first** (step 1) because it runs before all QA account pulls.

```
1.  Masterlist       masterlist_fetch.py   (CSV — Google Sheets)
2.  Coaching         asana_pull.py
3.  M7               m7_pull.py
4.  DMG              dmg_pull.py
5.  R4H              r4h_pull.py
6.  Parentis Health  parentis_pull.py
7.  Britelift        britelift_pull.py
8.  Britelift Chat   britelift_pull.py     (same script name, different dir)
9.  RideX            Ridex_pull.py
10. Hamilton         Hamilton_pull.py
11. Skyline          Skyline_pull.py
12. VIP              vip_pull.py
13. C&H              ch_pull.py
14. Reno Cab         rc_pull.py
15. Trans Iowa       ti_pull.py
16. Data Carz        dc_pull.py
17. Associated Cab   ac_pull.py
18. Ollies           ol_pull.py
19. Circle Taxi      ct_pull.py
20. YCOV             ycov_pull.py
21. Kelowna          kel_pull.py
22. Vermont          vt_pull.py
23. YCDC             ycdc_pull.py
24. Blueline         bl_pull.py
    Build            dashboard.py
    Git Push         git push
```

Left/right panel balance is dynamic. With the current 24 sources, the monitor renders Data Sources 1–12 on the left and 13–24 on the right. If accounts are added later, the split is recalculated from the source count instead of being hardcoded.

---

## Status states

| Status | Color | Dot | Meaning |
|--------|-------|-----|---------|
| `pass` | `#00e87a` green | Blinking green | Step completed successfully |
| `fail` | `#ff3d3d` red | Blinking red (fast) | Step exited non-zero; error captured from stderr |
| `running` | `#ffaa00` amber | Blinking amber | Step is currently executing |
| `blocked` | `#E84500` orange | Blinking orange | Pipeline stopped upstream; step never ran |
| `pending` | `#E84500` orange | Blinking orange | Account not in last run's data (new account) |

Both `blocked` and `pending` display as **"NOT REACHED"** in the UI. Internally distinct so logic works correctly.

### Special center-column nodes

- **Pipeline / Container Monitoring** (radar) — reflects overall run status: complete / running / stopped. **Sweep arm turns red** if any step failed (account, Build, or Git Push).
- **Data Engine / Processing** (agg node) — shows total rows, sources ok (blinking green/red), started, finished, next scheduled refresh
- **Dashboard** (browser card) — shows Live Data / Stale Data using same `schedTimes` logic as main dashboard
- **Git Repository** — `queued…` (amber) during run, `pushed` (green) on success, `not reached` (orange) if pipeline failed before git step
- **Header times** — the monitor header intentionally shows both `Dashboard:` and `Monitor:` timestamps. `Dashboard:` is sourced from the Build step/dashboard freshness timestamp; `Monitor:` is when `pipeline_monitor.html` itself was generated. This avoids a false mismatch when the monitor is regenerated after the dashboard build.

---

## Account card badges

Each account node card shows a small badge in the meta-pill row:

| Badge | Color | Meaning |
|-------|-------|---------|
| `✓` | Green | Account passed; row count same or higher than last run |
| `↓ N` | Red/orange | Account passed but row count dropped by N rows vs last run |
| *(none)* | — | Account failed or not reached (state already visible) |

---

## Radar sweep behavior

The radar sweep arm, trail, tip glow, and center dot change color based on run state:

| State | Color |
|-------|-------|
| Success / Running | Purple / violet (`rgba(200,130,255,...)`) |
| Any failure (account fail, Build fail, Git fail, pipeline failed) | Red (`rgba(255,70,70,...)`) |

---

## Incident log (`pipeline_log.json`)

- Appended only when `run_status` is `"success"` or `"failed"` — never mid-run
- Deduped by `run_id` — regenerating the monitor never double-logs
- Each entry: `{ run_id, date, account, script, status, error }`
- Capped at 300 entries; HTML table shows most recent 80, newest first
- Committed to git with every pipeline finish push so history persists

### Status types in the incident log

| Status value | Pill label | Color | Trigger |
|---|---|---|---|
| `fail` | FAILED | Red | Account, Build, or Git Push exited non-zero |
| `not_reached` | NOT REACHED | Orange | Step was blocked by an upstream failure |
| `count_drop` | COUNT DROP | Amber | Account passed but pulled fewer rows than the previous run |

Build (`dashboard.py`) and Git Push failures are logged to the incident table just like account failures.

---

## Row count drop detection

`pipeline_rowcount_baseline.json` stores the last known row count per account as a flat dict:
```json
{"Kelowna": 2261, "Associated Cab": 2191, ...}
```

After each finished run, `generate_monitor.py`:
1. Loads the baseline
2. For every account that passed (has a row count), compares current vs baseline
3. If current < baseline → appends a `count_drop` entry to the incident log and adds a `↓ N` badge to the account card
4. Shows an amber **Row count drop detected** warning strip listing all affected accounts with before → after numbers
5. Updates the baseline with the new counts

The baseline is committed to git so it survives machine restarts and re-clones.

---

## Live heartbeat

`log_step.py` calls `push_live(account)` after every step. This:
1. Runs `generate_monitor.py` (rebuilds HTML with current mid-run state)
2. `git add pipeline_status.json pipeline_monitor.html`
3. `git commit -m "[live] AccountName"` (only if there are staged changes)
4. `git pull --rebase --autostash`
5. `git push`

All wrapped in try/except — **never aborts the pipeline** if the monitor push fails. Adds ~5–15s per step overhead.

The page has a 60s auto-reload so viewers see updates approximately every 60s during a run.

---

## Pipeline Guardian

`pipeline_guardian.py` is a conservative safety check for the monitor layer. It does **not** re-pull account data and does **not** hide source-side count drops. Its job is to catch and repair stale monitor output, especially when `pipeline_status.json` says the run finished successfully but `pipeline_monitor.html` still shows blocked or not-reached steps.

Default report-only check:
```bat
py -3 pipeline_guardian.py --live
```

Repair stale local monitor output:
```bat
py -3 pipeline_guardian.py --fix
```

Repair, commit, push, and verify the live GitHub Pages monitor:
```bat
py -3 pipeline_guardian.py --fix --push --live
```

There is also a launcher:
```bat
pipeline_guardian.bat
```

Guardian behavior:
- Compares `pipeline_status.json` against the local generated monitor stats.
- Regenerates `pipeline_monitor.html` only when there is a mismatch and `--fix` is provided.
- Pushes only approved monitor artifacts when `--push` is provided.
- Blocks push if non-monitor files are dirty, so dashboard/code edits do not get mixed into monitor repairs.
- Reports live GitHub Pages mismatch as likely deployment/cache delay instead of overwriting source data.
- Treats cache files such as `masterlist_cache.csv`, `history_cache.csv`, `movement_cache.csv`, `movement_notified.json`, and `step_err.tmp` as local runtime files.

### Guardian incident logging (added 2026-07-04)

`log_guardian_event(issues, fixed, pushed)` in `pipeline_guardian.py` appends a Guardian entry to `pipeline_log.json` so a guardian run shows up in the same incident log as pipeline failures and count drops.

- Fires **only** when a monitor↔status mismatch is actually detected — never logged on a clean run.
- `guardian_fix` — a `--fix` attempt regenerated the monitor and fully resolved the mismatch.
- `guardian_warn` — a mismatch was found in report-only mode (no `--fix`), or a `--fix` attempt did not fully resolve it.
- Entry schema matches every other incident-log entry: `{ run_id, date, account: "Guardian", script: "pipeline_guardian.py", status, error }` — `error` is the mismatch list joined with `"; "`.
- `pipeline_log.json` is already in `APPROVED_PUSH_FILES`, so Guardian-authored entries survive the guardian's own `commit_and_push()`.

New pill statuses in `generate_monitor.py` → `render_log_table()`:

| Status value | Pill label | Color | Trigger |
|---|---|---|---|
| `guardian_fix` | GUARDIAN FIX | Teal `.lp-g` `#0e7490` | `--fix` attempt resolved a mismatch |
| `guardian_warn` | GUARDIAN WARN | Dark amber `.lp-g` `.lp-gw` `#92400e` | Mismatch detected and not resolved (report-only, or failed fix) |

The Git Repository stage node also surfaces the most recent Guardian run, sourced from the latest `account == "Guardian"` entry in the log:
- `✓ Guardian fixed · <date>` (teal) when the latest Guardian status is `guardian_fix`
- `⚠ Guardian warned · <date>` (amber) when the latest Guardian status is `guardian_warn`
- Nothing shown if no Guardian entry exists yet

---

## Employee Movement Notifications

`check_movement_notifications.py` runs immediately after `masterlist_fetch.py`, before the account QA pulls. It reads `movement_cache.csv`, compares each processed movement against `movement_notified.json`, and sends one HTML email per newly processed movement.

Notification eligibility:
- `Processed` must be `Yes`
- `Void` must not be `Yes`
- `Timestamp` must not already exist in `movement_notified.json`

Duplicate protection:
- `movement_notified.json` stores the timestamps that already generated movement emails.
- If `movement_notified.json` is missing, the script now seeds it from all existing processed/non-void movement rows and sends **no emails**. This prevents old processed movements from firing again after a local cleanup, clone, or tracker reset.

Employment status display in movement emails:
- The movement source can store class-style values such as `Regular`, `Probationary`, or `Option`.
- Emails must display only dashboard-style employment status values: `Active`, `Inactive`, or `No changes`.
- Mapping is handled by `normalize_employment_status_for_email()`:
  - `Regular`, `Probationary`, `Option`, full-time, and part-time style values → `Active`
  - Attrition rows → `Inactive`
  - Inactive/terminated/resigned/separated/end-of-contract style values → `Inactive`
  - Blank/N/A/dash values → `No changes`

The notification script is intentionally not logged as a formal `pipeline_status.json` step today; it prints its own `[movement_notify]` messages and the pipeline continues with a warning if the notification step exits non-zero.

`movement_reconcile.py` is the post-run safety net for this flow. It reads the same three local caches (`masterlist_cache.csv`, `history_cache.csv`, `movement_cache.csv`) and verifies that processed/non-void Movement rows are represented in both the notification ledger and the generated dashboard snapshot. It delegates missing emails back to `check_movement_notifications.py`, then refreshes the embedded `masterlist`, `historyData`, `movementData`, and `masterlistKpis` constants in `masterlist_dashboard.html` from cache so a processed movement cannot remain stale in the published dashboard after a successful build.

Where it runs:
- After `dashboard.py` in both `update_coaching_dashboard_auto.bat` and `update_coaching_dashboard.bat`, before Git publish.
- After `check_movement_notifications.py` in `update_movement_notifications_auto.bat`, for movement-only reconciliation. If this patches `masterlist_dashboard.html`, the movement-only batch commits and pushes that dashboard snapshot so the live Movement table catches up without waiting for the full QA pipeline.

---

## Adding a new account

1. Add the pull step to both bat files (same pattern as existing steps)
2. Add entry to `ACCOUNTS` list in `generate_monitor.py` — maintain pipeline execution order
3. Add short name to `SHORT_NAMES` dict in `generate_monitor.py`
4. Run `py -3 generate_monitor.py` to verify the radar and panels render correctly

---

## Layout — LOCKED design decisions (settled 2026-07-12, do not redesign)

These were settled with Mike through several preview rounds. Any future change to
monitor layout must respect these rules — do NOT re-litigate them:

1. **Three-column flanked layout IS the design.** Account cards in a LEFT column
   (Data Sources 1–12) and a RIGHT column (13–24) flanking the center hub
   (radar → Data Engine → Dashboard → Git Repository). Mike explicitly rejected
   a merged tile-wall layout with the hub promoted on top. Never re-stack on
   desktop.
2. **Side columns are capped at 330px** (`minmax(280px,330px)`), center column
   `minmax(328px,350px)`, `justify-content:center`. Mike rejected wide
   (`0.95fr`) stretchy cards.
3. **`.eco` container max-width is 1070px** — it must HUG the card block. Mike
   rejected empty grid-pattern background flanking the cards (the old 1390px).
4. **Narrow windows AUTO-SCALE, never re-stack.** `fitEco()` (bottom `<script>`)
   applies CSS `zoom` when viewport < `natural` px so quarter-screen FancyZones
   shows the full desktop layout, just smaller. Mike explicitly rejected the
   old stacking breakpoints as primary behavior — they survive only inside
   `@supports not (zoom:1)` as an old-browser fallback.
5. **`natural` (in fitEco) and `.eco` max-width are COUPLED — both 1070.** If
   column widths ever change, update BOTH or scaling starts at the wrong width.
6. **Readability compensation:** when fit zoom < 0.85, `.fit-sm` class bumps
   small text (node titles 15px, rows 13px, pills 13px, section headers 14px).
   Desktop (zoom 1) must stay unaffected.
7. **Manual browser zoom (Ctrl +/−) must be respected** — fitEco skips refits
   when devicePixelRatio changes (that's a user zoom, don't fight it).
8. **KPI/stat-card row is untouchable** — original flex layout. Mike explicitly
   said not to change it when it was accidentally restyled once.
9. **Header logo = `pacbiz_logo_dark.png`** (dark rendering, same as the
   Scheduler login), plain `<img>`, no blend/filter hacks. Falls back to the
   white `pacbiz_logo.png` if the dark file is missing (`logo_b64()`).

**Previewing safely:** never run `generate_monitor.py` just to preview CSS — it
writes state files and can send emails / trigger auto-heal. Instead copy the
live `pipeline_monitor.html`, patch the CSS strings in the copy, and open that
(`pipeline_monitor_preview.html` convention; delete after). Also remember the
scheduled pipeline runs `generate_monitor.py` from the LOCAL working copy — any
uncommitted local edit ships on the next run, so never leave experiments in the
file between runs.

## Color reference

| Token | Hex | Usage |
|-------|-----|-------|
| Pass green | `#00e87a` | Pass status, sources-ok dot when all pass |
| Fail red | `#ff3d3d` | Fail status |
| Running amber | `#ffaa00` | Running / in-queue status |
| Not reached orange | `#E84500` `rgb(232,69,0)` | Blocked / pending status |
| Count drop amber | `#ffaa00` | COUNT DROP pill in incident log |
| Brand purple | `#5000b4` / `#9050ff` | Section headers, gradients |
| Card text | `#c0a0ff` | Node titles, agg values |
| Muted | `#7060a0` | Labels, secondary text |
| Background | `#0a0018` / `#05001a` | Page / card backgrounds |

---

## Known incidents

### 2026-06-22 — VIP transform crash (Build failure)

**Error:**
```
File "dashboard.py", line 1864, in transform_vip_data
    df[vip_key] = [... for ai, mx in zip(ai_s, max_s)]
ValueError: Length of values (0) does not match length of index (1000)
```

**Root cause:** The VIP pull returned a sheet missing one of the criterion columns expected by `transform_vip_data`. Both the `VIP_CRIT_MAP` loop (line 1835) and `VIP_EXTRA_CRIT_MAP` loop (line 1861) used `df.get(col, pd.Series(dtype=float))` — the known footgun that returns a 0-length Series when the column is absent, causing the list comprehension to produce 0 items against a 1,000-row DataFrame.

**Fix (commit `478a497`):** Replaced both instances with the safe missing-column pattern:
```python
# Before (footgun):
ai_s = pd.to_numeric(df.get(ai_col, pd.Series(dtype=float)), errors="coerce").fillna(0)

# After (safe):
ai_s = pd.to_numeric(df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float), errors="coerce").fillna(0)
```

**Impact:** Pipeline Build step failed; dashboard was not auto-rebuilt. Manually fixed, rebuilt, and pushed. All 22 data source pulls completed successfully before the crash.

---

## Known limitations

- **Live push adds time** — ~5–15s per step for git operations. 22 steps = up to ~5min extra total pipeline time.
- **GitHub Pages delay** — ~1 min after each push before the page reflects changes. The 60s reload timer is tuned for this. Large HTML files (26MB) can take 3–5 min.
- **Row counts are snapshots** — baseline only stores last run. No full history trend. A drop is flagged once and then the baseline updates.
- **Mid-run Git Repository node** shows "queued…" because the Git Push step hasn't been logged yet. This resolves to "pushed" once the pipeline finishes.
- **Incident log dedup is per run_id** — if `generate_monitor.py` is re-run mid-session after manually patching `pipeline_status.json`, the log entry for that run won't be re-written. Add entries manually to `pipeline_log.json` if needed.

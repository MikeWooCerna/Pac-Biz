# Pipeline Monitor — Technical Reference

**Live URL:** https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html  
**Last updated:** 2026-06-22  
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
| `pipeline_log.json` | Persistent incident history. Appended to by `generate_monitor.py` when a run finishes with failures or not-reached accounts. Survives across runs. Capped at 300 entries. |
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
       ├─ appends incidents to pipeline_log.json (if run finished with failures)
       └─ writes pipeline_monitor.html

  └─ git add pipeline_status.json pipeline_monitor.html pipeline_log.json
  └─ git commit + push
```

---

## Account order (ACCOUNTS list in generate_monitor.py)

Matches pipeline execution order. Masterlist is **first** (step 1) because it runs before all QA account pulls.

```
1.  Masterlist       masterlist_fetch.py   (CSV — Google Sheets)
2.  Coaching         asana_pull.py
3.  M7               m7_pull.py
4.  Parentis Health  parentis_pull.py
5.  Britelift        britelift_pull.py
6.  Britelift Chat   britelift_pull.py     (same script name, different dir)
7.  RideX            Ridex_pull.py
8.  Hamilton         Hamilton_pull.py
9.  Skyline          Skyline_pull.py
10. VIP              vip_pull.py
11. C&H              ch_pull.py
12. Reno Cab         rc_pull.py
13. Trans Iowa       ti_pull.py
14. Data Carz        dc_pull.py
15. Associated Cab   ac_pull.py
16. Ollies           ol_pull.py
17. Circle Taxi      ct_pull.py
18. YCOV             ycov_pull.py
19. Kelowna          kel_pull.py
20. Vermont          vt_pull.py
21. YCDC             ycdc_pull.py
22. Blueline         bl_pull.py
    Build            dashboard.py
    Git Push         git push
```

Left panel = Data Sources 1–11. Right panel = Data Sources 12–22. Numbers are dynamic — auto-adjusts if accounts are added.

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

- **Pipeline / Container Monitoring** (radar) — reflects overall run status: complete / running / stopped
- **Data Engine / Processing** (agg node) — shows total rows, sources ok (blinking green/red), started, finished, next scheduled refresh
- **Dashboard** (browser card) — shows Live Data / Stale Data using same `schedTimes` logic as main dashboard
- **Git Repository** — `queued…` (amber) during run, `pushed` (green) on success, `not reached` (orange) if pipeline failed before git step

---

## Radar blips

22 blips evenly spaced around the radar. Alternating inner ring (`r=0.44`) and outer ring (`r=0.63`). Masterlist is at -90° (12 o'clock). Colors match status states above. Labels pulse for fail/running/not-reached.

Sweep animation: `ang += 0.012` per frame (~8.7s per rotation at 60fps). Trail arc = 1.35π.

---

## Incident log (`pipeline_log.json`)

- Appended only when `run_status` is `"success"` or `"failed"` — never mid-run
- Deduped by `run_id` — regenerating the monitor never double-logs
- Each entry: `{ run_id, date, account, script, status ("fail"|"not_reached"), error }`
- Capped at 300 entries; HTML table shows most recent 80, newest first
- Committed to git with every pipeline finish push so history persists

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

## Row counts

Row counts are read from the actual Excel/CSV files on disk immediately after each pull script completes. They reflect what was pulled from the source system for that run. Numbers change between runs — this is expected. No historical trending is stored (parked as future work).

---

## Adding a new account

1. Add the pull step to both bat files (same pattern as existing steps)
2. Add entry to `ACCOUNTS` list in `generate_monitor.py` — maintain pipeline execution order
3. Add short name to `SHORT_NAMES` dict in `generate_monitor.py`
4. Run `py -3 generate_monitor.py` to verify the radar and panels render correctly

---

## Color reference

| Token | Hex | Usage |
|-------|-----|-------|
| Pass green | `#00e87a` | Pass status, sources-ok dot when all pass |
| Fail red | `#ff3d3d` | Fail status |
| Running amber | `#ffaa00` | Running / in-queue status |
| Not reached orange | `#E84500` `rgb(232,69,0)` | Blocked / pending status |
| Brand purple | `#5000b4` / `#9050ff` | Section headers, gradients |
| Card text | `#c0a0ff` | Node titles, agg values |
| Muted | `#7060a0` | Labels, secondary text |
| Background | `#0a0018` / `#05001a` | Page / card backgrounds |

---

## Known limitations

- **Live push adds time** — ~5–15s per step for git operations. 22 steps = up to ~5min extra total pipeline time.
- **GitHub Pages delay** — ~1 min after each push before the page reflects changes. The 60s reload timer is tuned for this.
- **Row counts are snapshots** — no trend history. A drop in rows between runs is expected and not flagged automatically (future work).
- **Mid-run Git Repository node** shows "queued…" because the Git Push step hasn't been logged yet. This resolves to "pushed" once the pipeline finishes.

# PB Dashboard — project context
# This file is read by Claude Code (auto) and Codex (manual).
# Last updated: 2026-06-27

## What this is

Python dashboard builder for Pac-Biz QA and Coaching data.
`dashboard.py` (10,727 lines, ~508 KB) reads local Excel files produced by
per-account pull scripts, then writes `masterlist_dashboard.html` (~26 MB,
fully self-contained). The HTML is committed to Git and published via
GitHub Pages automatically by `update_coaching_dashboard_auto.bat`.

Coaching data comes from Asana. QA data comes from Google Sheets.
Masterlist KPI data is fetched live from a published Google Sheets CSV
at build time — it has no pull script and no local Excel file.

## Key files

| File | Purpose |
|------|---------|
| `dashboard.py` | Main builder — all HTML/CSS/JS generated here |
| `masterlist_dashboard.html` | Single-file HTML output served via GitHub Pages |
| `update_coaching_dashboard_auto.bat` | Full automated pipeline (Task Scheduler target) |
| `update_coaching_dashboard.bat` | Manual version of the same pipeline |
| `pull_coaching_update_publish.bat` | Coaching-only partial update + publish |
| `Masterlist_Pull.py` | Masterlist Google Sheets CSV pull (standalone) |
| `eval_dist_card.jsx` | Evaluation distribution card component (reference only) |
| `requirements.txt` | Python deps: pandas>=2.2.0, openpyxl>=3.1.0, requests>=2.32.0, xlsxwriter>=3.2.0 |
| `test_connection.py` | Google Sheets API connectivity test |
| `pacbiz_logo.png` / `pacbiz_favicon.png` | Branding assets embedded in the HTML at build time |
| `diagnose_drops.py` | Count drop investigation agent — runs after every build, sends email report, auto-triggers Apps Script heal |
| `appsscript_triggers.json` | Maps account name → Apps Script web app URL for auto-heal (currently: Kelowna) |
| `pipeline_drops_notified.json` | Tracks which drop events have already been reported — prevents duplicate emails |

Python executable: `C:\Users\Mike Woo Cerna\AppData\Local\Programs\Python\Python313\python.exe`

## Pipeline sequence (update_coaching_dashboard_auto.bat)

Each step runs `py -3 <script>.py` inside the account's directory.
If any step returns a non-zero exit code the entire pipeline aborts (`goto :fail`).

```
1.  Coaching          — %COACHING_DIR%\asana_pull.py
2.  M7                — %M7_DIR%\m7_pull.py
3.  DMG               — %DMG_DIR%\dmg_pull.py
4.  R4H               — %R4H_DIR%\r4h_pull.py
5.  Parentis Health   — %PARENTIS_DIR%\parentis_pull.py
6.  Britelift         — %BRITELIFT_DIR%\britelift_pull.py
7.  Britelift Chat    — %BLC_DIR%\britelift_pull.py        ← SAME script name, different dir
8.  RideX             — %RIDEX_DIR%\Ridex_pull.py
8.  Hamilton          — %HAMILTON_DIR%\Hamilton_pull.py
9.  Skyline           — %SKYLINE_DIR%\Skyline_pull.py
10. VIP               — %VIP_DIR%\vip_pull.py
11. C&H               — %CH_DIR%\ch_pull.py
12. Reno Cab          — %RC_DIR%\rc_pull.py
13. Trans Iowa        — %TI_DIR%\ti_pull.py
14. Data Carz         — %DC_DIR%\dc_pull.py
14. Associated Cab    — %AC_DIR%\ac_pull.py
15. Ollies            — %OL_DIR%\ol_pull.py
16. Circle Taxi       — %CT_DIR%\ct_pull.py
17. YCOV              — %YCOV_DIR%\ycov_pull.py
18. Kelowna           — %KEL_DIR%\kel_pull.py
19. Vermont           — %VT_DIR%\vt_pull.py
20. YCDC              — %YCDC_DIR%\ycdc_pull.py
21. Blueline          — %BL_DIR%\bl_pull.py
22. git pull --rebase --autostash   (sync before rebuild)
23. py -3 dashboard.py              (rebuild HTML)
24. git add → git commit → git pull --rebase → git push
```

## Architecture decisions already made

- **Per-account transform pattern:** every account has a matching pair of functions
  in `dashboard.py`: `transform_X_data(df)` and `refresh_X_output()`.
  Do not collapse or merge these — they are intentionally isolated.
- **Safe missing-column pattern:** use `pd.Series(0, index=df.index, dtype=float)`
  as the fallback (not `df.get(col, pd.Series(dtype=float))` which returns
  length-0 and causes a length mismatch crash).
- **acctPill ternary:** appears in TWO places — `qaRowHtml()` (uses `r._acct`)
  and the agent leaderboard (uses `a.acct`). Any new account must be added to both.
  Parentis is the hardcoded final fallback — always insert before it.
- **aivhAccts array** controls which accounts show the AI vs Human comparison card:
  `['hamilton','skyline','vip','ch','rc','ti','dc','ac','ol','ct','ycov','kel','vt','ycdc','bl']`
  M7, DMG, R4H, Parentis, Britelift, Britelift Chat, and RideX are NOT in this list.
- **Evaluation ID rules in the QA detail table:** AI accounts must display the
  raw `evaluation_id` value from the AI workbook. Non-AI accounts must display
  `QA_ID`/`qa_id`. Do not prefer generated `QA_ID` values for AI accounts
  (example bad Kelowna value: `KEL-...`; expected AI value format: `VABaKQC-...`).
  The JS helper set `QA_AI_ACCOUNTS` and every AI transform output must preserve
  the `evaluation_id` column.
- **sumStripMain** is hidden for all 15 accounts in `aivhAccts` — they use a
  custom criterion strip instead.
- **Cache-busting meta tags** added to `<head>` — do not remove them.
- **Auto-reload JS** added before `</body>` — fires at scheduled build times.
- **Refresh schedule (auto-reload):**
  `03:30, 06:30, 11:30, 15:30, 19:30, 22:30` — page reloads at these times (Task Scheduler fires 30 min earlier)
- **Freshness indicator schedule:** uses Task Scheduler fire times `03:00, 06:00, 11:00, 15:00, 19:00, 22:00`
  so a build that finishes at e.g. 3:11 PM is correctly flagged Live (build > 15:00 slot), not Stale.
  Stale only appears if a scheduled run was missed, failed, or the machine was off — never during normal operation.
- **Data freshness indicator** in the top header beneath "Refresh Time:" —
  green blinking dot = Live Data, yellow blinking dot = Stale Data.
  Both dots use `animation:qa-pulse 2s ease-in-out infinite` (reuses existing QA panel animation).
  Dot colors: Live = `#16A34A`, Stale = `#D97706`. Text colors: Live = `#15803D`, Stale = `#B45309`.
- **subprocess.run()** with `capture_output=True, timeout=120` for all subprocesses.
- All `{` and `}` inside JS/CSS strings in Python f-strings must be escaped as `{{` and `}}`.
- **Data freshness age ticks live:** the `updateFreshness()` function runs on page load then
  repeats every 60 seconds via `setInterval(updateFreshness, 60000)` — no refresh needed
  for the age string to stay accurate.
- **Build finish-time stamping:** `refresh_time` and `refresh_iso` are captured at build START
  in `dashboard.py` as placeholder strings. Just before `Path(OUTPUT_FILE).write_text()`, both are
  replaced with `datetime.now()` (finish time) so "Refresh Time:" and `PB_BUILD_TS` reflect when
  the build actually completed, not when it started. Do not remove this final patch step.

## GitHub

- **Remote:** https://github.com/MikeWooCerna/Pac-Biz.git
- **Branch:** `main`
- **Live URL:** https://mikewoocerna.github.io/Pac-Biz/masterlist_dashboard.html
- GitHub Pages takes ~1 minute to reflect a push.

## Active work / next steps

Pipeline monitoring system is fully live as of 2026-06-22. See `PIPELINE_MONITOR.md` for full technical reference.

### Changes made 2026-06-23 (batch 1)
- **Row count drop alerts** — `pipeline_rowcount_baseline.json` stores last known count per account. After each run, any account that pulled fewer rows than the baseline triggers: (1) an amber warning strip on the monitor, (2) a `↓ N` red badge on the account card, (3) a `COUNT DROP` entry in the incident log. Baseline auto-updates each run.
- **Account card badges** — every passing account card now shows a small `✓` green badge (no drop) or `↓ N` red badge (count dropped).
- **Build and Git Push failures in incident log** — previously only data-source account failures were logged. Now `dashboard.py` failures and `git push` failures also appear as `FAILED` rows in the incident log table.
- **Radar sweep turns red on failure** — sweep arm, trail, tip glow, and center dot switch from purple to red whenever any step fails (account, Build, or Git Push).
- **VIP transform crash fixed** — see Known issues below.

### Changes made 2026-06-23 (batch 2)
- **Attrition movement type fix** — the movement Google Sheet has two columns: `Type of Movement` ("Attrition" or "Internal Movement") and `Movement Type` (sub-type, blank for Attrition rows). `dashboard.py` now fills blank `Movement Type` from `Type of Movement` at build time so Attrition rows appear correctly in the Recent Employee Movements table.
- **"Initiated by" column in Recent Employee Movements** — the `Email Address` field from the movement data is looked up against `Company Email` in the masterlist to resolve the employee name. Displayed as the last column in the table. Blank when email is missing or has no match.
- **xlsxwriter engine for all pull scripts** — all 19 QA account pull scripts (in `Quality/` subdirectories) were switched from `openpyxl` to `xlsxwriter` as the pandas Excel engine. Root cause: Skyline's dataset reached 14,297 rows × 138 columns and triggered a `MemoryError` in openpyxl, which builds the entire workbook in RAM before writing. xlsxwriter writes rows sequentially to disk and is not affected by dataset size. **Any new pull script must use `engine="xlsxwriter"` in `df.to_excel()`.**
- **HIGH VOLUME warning system** — `generate_monitor.py` now detects accounts whose row count exceeds configured thresholds:
  - `HIGH_VOLUME_WARN = 10,000` rows → amber `HIGH VOL` badge on card + amber strip + one-time email (WARNING)
  - `HIGH_VOLUME_CRIT = 20,000` rows → red `CRITICAL VOL` badge + red strip + one-time email (CRITICAL)
  - Escalation from warn → crit sends a second email. Dropping back below 10,000 resets the account.
  - State persisted in `pipeline_highvol_notified.json` (tracked by git, auto-updated each run).
  - Email sent via `notify.notify_high_volume()`.
  - Incident log pill: `HIGH VOLUME` (amber, style `.lp-v`).

### Changes made 2026-06-23 (batch 3)
- **Continue-on-failure pipeline** — both `update_coaching_dashboard_auto.bat` and `update_coaching_dashboard.bat` now continue past individual account failures instead of aborting. Each failed account is logged and skipped; the dashboard still rebuilds using the last cached Excel file for any failed account. The pipeline monitor shows failed accounts in red so you know which data is stale.
  - **Fatal steps (still abort everything):** `git pull --rebase` before build, `dashboard.py` build, `git push`
  - **Non-fatal steps (skip and continue):** all 21 account pull steps (Masterlist fetch, Coaching, and all 19 QA accounts)
  - **New finish state:** `partial` — logged when the build succeeds but one or more accounts failed. Monitor shows "Done with warnings" in this case.

### Changes made 2026-06-26
- **Recent Employee Movements table cap fix** — `.slice(-10)` changed to `.slice(-20)` in `recentMovementsTable()` JS function in `dashboard.py`. Previously capped at 10 entries; now shows up to 20.
- **Count drop investigation agent** (`diagnose_drops.py`) — new file called automatically at the end of `generate_monitor.py` after every pipeline build. Reads `pipeline_log.json`, detects NEW count drops not yet reported, classifies each account's pattern (RECURRING / CLUSTER / ISOLATED / MINOR), and sends a styled investigation email to `reports@pac-biz.com` with findings + recommendations. Already-reported drops tracked in `pipeline_drops_notified.json` — no duplicate emails.
  - **RECURRING** = same account drops on multiple runs → likely a date-range filter in the Apps Script
  - **CLUSTER** = multiple accounts drop in the same run → likely an upsert overwrote the sheet
  - **ISOLATED** = single large one-time drop → transient data source issue
  - **MINOR** = single small drop (≤50 rows) → likely legitimate upstream deletion
- **Apps Script web app auto-heal for Kelowna** — when a Kelowna count drop is detected, `diagnose_drops.py` automatically runs the full heal sequence (see details in 2026-06-27 section below).
  - Kelowna trigger URL stored in `appsscript_triggers.json`
  - To add more accounts: deploy `doGet()` on their Apps Script, add URL to `appsscript_triggers.json`, add pull script path to `ACCOUNT_PULL_SCRIPTS` in `diagnose_drops.py`
  - Kelowna Apps Script updated with leftover-sheet cleanup (`KEL_Evaluations_new` guard) in both `pullKELEvaluationsToSheet()` and `pullKELLast30Days()`

### Changes made 2026-06-27
- **Auto-heal sequence corrected** — original `doGet()` called `pullKELLast30Days()` (rolling 30 days, ~731 rows) which wiped the full year dataset. Fixed: `doGet()` now calls `clearKELProgress()` + `pullKELEvaluationsToSheet()` (full year, ~2,300+ rows). Any future Apps Script web app for other accounts must follow the same pattern — never call the 30-day pull from `doGet()`.
- **`doGet()` JSON protocol** — now returns `{"status":"partial","rows_so_far":N,"last_week":"..."}` or `{"status":"complete","rows":N}` instead of plain text "OK". Includes `SpreadsheetApp.flush()` before `getLastRow()` to avoid stale sheet-rename timing. On first call clears progress; on resume calls skips `clearKELProgress()` so prior weekly progress is preserved.
- **Multi-call heal loop** (`diagnose_drops.py`) — `trigger_appsscript()` now loops up to `MAX_TRIGGER_CALLS = 12` times, calling the web app URL until it returns `"complete"`. Needed because `pullKELEvaluationsToSheet()` uses a 5-minute PropertiesService resume pattern — one HTTP call may only cover part of the year. Each partial response logs progress; loop breaks on `"complete"`. Falls back to plain-text "OK" for backward compatibility.
- **Pre-drop count verification** — `trigger_appsscript()` now accepts `prev_count` (the exact row count before the drop, read from the drop log event). After the Apps Script completes, it verifies the restored count is ≥ 80% of `prev_count`. Fails clearly if the count is still too low rather than proceeding to rebuild the dashboard with incomplete data. `run()` passes the highest `prev_count` from all new drops for that account.
- **Healed log entry enriched** — includes `prev_count`, `healed_count`, and percentage restored in the `error` detail field. Visible in the incident log.
- **Drop badge auto-clears after heal** (`generate_monitor.py`) — after building `drop_by_account`, checks if a `healed` log entry exists for that account with a later `run_id` (ISO string comparison). If yes, removes the account from `drop_by_account` so the "↓ N" badge is no longer shown. Previously the badge persisted across pipeline runs even after a successful heal.
- **Re-trigger loop bug fixed** (`diagnose_drops.py`) — drops are now marked as notified in `pipeline_drops_notified.json` BEFORE calling `trigger_appsscript()`, not after email success. The heal subprocess calls `generate_monitor.py` → `diagnose_drops.run()` again; if drops were still unnotified at that point, `run()` would re-trigger the heal and loop infinitely.
- **`pipeline_drops_notified.json` tracked in git** — added to the `git add` line in both `update_coaching_dashboard_auto.bat` and `update_coaching_dashboard.bat` (both `:publish_monitor` and `:fail` sections). Previously untracked; notification state would be lost on fresh clone.
- **Dashboard/monitor timestamp sync** — `dashboard.py` now stamps the actual build finish time (not start time) into the HTML just before writing to disk. `refresh_time` and `refresh_iso` are captured at build start as placeholders; a finish-time patch replaces both strings with `datetime.now()` values immediately before `Path(OUTPUT_FILE).write_text()`. `generate_monitor.py` syncs the Build step timestamp in `pipeline_status.json` to the HTML file's mtime after each run so both the monitor "Built:" display and the dashboard "Refresh Time:" header agree on when the build actually finished.
- **Auto-heal recommendation in drop emails** — when a count drop email is sent for an account that does NOT have an Apps Script trigger configured, `diagnose_drops.py` appends an HTML recommendation block: ✅ Recommended (RECURRING or large ISOLATED drop), ⚠ Investigate first (CLUSTER drop), or ❌ Not needed (MINOR drop or already recovered). Accounts that already have auto-heal configured are excluded from this block.
- **Coaching validation email alerts** — `dashboard.py` sends `PACE Validation Errors Detected in Coaching Entries` when coaching tasks are excluded because employee/supervisor emails do not match the Masterlist. Details are written to `coaching_validation_errors.csv`; duplicate emails are prevented with a separate `coaching_validation_notified.json` signature file. This is intentionally separate from `movement_notified.json` and must not trigger movement re-notifications.
- **Lightweight movement notifier** — `update_movement_notifications_auto.bat` refreshes Masterlist caches and runs only `check_movement_notifications.py`, so it can be scheduled more frequently than the full dashboard pipeline. Movement emails require `Processed = Yes`, not voided, plus nonblank `Timestamp`, `Employee Name`, `Type of Movement`, `Email Address`, `Processed Date`, and `Processed Note`. `movement_notified.json` still prevents duplicates.

### Changes made 2026-07-11
- **M7 upstream feed restored** (`Quality/M7/m7_appsscript_pull.gs`) — the Apps Script that fed the
  destination RAW tab was lost, freezing RAW at 29 rows since 2026-06-08. `m7_pull.py` was never the
  problem (it correctly reads dest spreadsheet `1Aq-IFsFS...` tab `RAW`). A new Apps Script
  `pullM7ToRaw()` recreates the transfer AND the enrichment:
  - **Source:** spreadsheet `1mAd86tsTts1xgPULyANMmgAYN0KUarLhrzJ-YWlTKNI`, tab selected by **gid 318886303** (never by name)
  - **Destination:** spreadsheet `1Aq-IFsFSCOUHkfH32vJ1v1QtyakzEYCQP2Y80pIlYEI`, tab `RAW` (full replace)
  - **Masterlist lookups:** spreadsheet `18hKmm2SmlWqB23osiV3JTF0aWn86vvZ-YJSC-Rr3JcY`, sheet `History`
  - **44-column RAW layout:** 4 generated ID cols (`QA_ID`, `EMPLOYEE_ID`, `QA_COACH_ID`, `SUPERVISOR_ID`)
    + 36 source cols verbatim + 4 generated trailing enrichment cols (`Emp Name`, `QA`,
    `Immediate Supervisor`, `LOB / Account`). Trailing cols reuse the SAME masterlist matches as the IDs.
  - **Coach resolution:** evaluator `Email Address` vs History `Company Email`, with a NAME fallback via
    the source `QA Coach:` column (needed for evaluators whose form email ≠ Company Email, e.g.
    Gina De Los Santos / ID 44, Nidalyn Mascardo / ID 574)
  - **Safety:** fail-closed — clearContents only after ALL reads/validations pass; header-shape check vs
    RAW's existing headers (this guard caught the missing 4 enrichment cols on the first live run);
    25-row safety floor after blank-row filtering; LockService concurrency guard; post-write row+column
    assertions; unmatched rows never dropped (`No match` + counters in the log summary)
  - **Trigger:** `installM7Trigger()` = every 2 hours, duplicate-guarded; `removeM7Triggers()` to clear
  - Any future `doGet()` web-app wrapper must call the FULL pull (Kelowna lesson — never a rolling-window variant)

- **QA_ID house pattern (documented for all non-AI accounts):**
  `QA_ID = "<ACCOUNT>-" + yyyyMMddHHmmss(evaluation Timestamp) + "-" + EMPLOYEE_ID`
  e.g. `M7-20260508040339-639`, `PARENTIS-20260507033119-487`, `BRITELIFT-20260409172331-488`.
  The numeric IDs (EMPLOYEE_ID / QA_COACH_ID / SUPERVISOR_ID) are masterlist History lookups
  (agent name match / evaluator email or QA Coach name match / employee's supervisor), NOT source data.

- **DMG non-AI QA account added** (`Quality/DMG/dmg_pull.py`, `Quality/DMG/dmg_appsscript_pull.gs`)
  follows the M7 pattern with a source feed into a destination RAW tab:
  - **Source:** spreadsheet `1i4CtBIXsdmpOwgbBLu0dQekSB4fxjGwUAgrUfiRexeU`, gid `919687516`, tab `Form Responses 1`
  - **Destination:** spreadsheet `1XGW2y3uz9sAEXkT-4qm1LvgcW5HolhttT28FyTsRYxc`, tab `RAW`
  - **Reviewer mapping:** source `Reviewer = Supervisor` becomes `Jamito, Frodel M.`;
    source `Reviewer = QA Evaluator` becomes `Burton, Gladys P`
  - **Local fallback:** `dmg_pull.py` reads destination `RAW` when present; if `RAW` is not present yet,
    it reads the source tab directly, applies Reviewer mapping, generates `QA_ID`, and writes `DMG_RAW.xlsx`
    so the scheduled pipeline does not fail while the Apps Script feed is being installed.
- **R4H non-AI QA account added** (`Quality/R4H/r4h_pull.py`, `Quality/R4H/r4h_appsscript_pull.gs`)
  is a non-AI QA account with its OWN source form and destination RAW tab — it does NOT share
  DMG's source or DMG's Reviewer-mapping logic. (An earlier draft of this note described R4H as
  sharing DMG's source spreadsheet and Reviewer mapping — that design was superseded; the values
  below reflect the actual live pipeline.)
  - **Source:** spreadsheet `13W-ij2u3PLVzt028Ih10fMFJ7bg4WgODs_UBRBHtUrk`, gid `62485887`
    (a distinct R4H form, NOT DMG's `1i4CtBIX...` sheet)
  - **Destination:** spreadsheet `1d3w61WWQgVDuFlfUnVwQyTPnNdyIrmc0M7xdf5C8cRw`, tab `RAW`
  - **No Reviewer mapping:** R4H's form carries its own `QA Coach:` / `Supervisor:` columns
    directly — there is no `Reviewer = Supervisor/QA Evaluator` remap step (that is DMG-only).
    `r4h_appsscript_pull.gs` even guards against consuming DMG's shape: it throws
    `"R4H pull aborted: source has Reviewer but no QA Coach column. This looks like a DMG source, not R4H."`
  - **Column map (`dashboard.py`):** R4H has its OWN dedicated `R4H_COLUMN_MAP` (~line 114) built
    from R4H's real headers. It must NOT be aliased to `M7_COLUMN_MAP` or `DMG_COLUMN_MAP` — R4H's
    criterion questions use different wording and point values (e.g. `Opening Spiel (Inbound Calls) - 2pts`
    vs M7's `- 1pt`), so an alias silently drops criterion data. See the 2026-07-11 fix note below.
  - **Safety rules:** blank `Employee Name` rows are excluded; generated bad IDs such as
    `R4H-INVALIDTS-NOEMPLOYEE` must never be written.
  - **Pipeline integration:** add R4H everywhere an account source is registered:
    `dashboard.py`, `update_coaching_dashboard_auto.bat`, `update_coaching_dashboard.bat`,
    `self_heal.py`, `generate_monitor.py`, `pipeline_rowcount_baseline.json`, and the QA account dropdown.

- **DMG + R4H live on the Quality tab as non-AI accounts (integration verified 2026-07-11)** —
  both accounts are fully wired into `dashboard.py` following the M7/Parentis non-AI pattern:
  NOT in `aivhAccts` (no AI vs Human card), NOT in `QA_AI_ACCOUNTS`, `sumStripMain` stays visible,
  QA detail table shows the house-pattern `QA_ID` (`DMG-yyyyMMddHHmmss-EMPID` / `R4H-...`). All 16
  new-account checklist sites present for both (dropdown, rawData const, `_acct` tag, `qaGetActiveData`
  case, all-accounts spread, both `acctPill` ternaries inserted before the Parentis fallback,
  KPI title/badge, eval-dist entry, chart labels/colors, `qaAcctsLoaded`).
  - **`R4H_COLUMN_MAP` defect found and fixed** — it was aliased as `R4H_COLUMN_MAP = M7_COLUMN_MAP`,
    which silently dropped every R4H criterion whose header/point-value differed from M7's form.
    Replaced with a dedicated 31-entry map built from `R4H_RAW.xlsx`'s actual columns
    (verified: no duplicate keys/values, every mapped value retained by `_transform_qa_source()`'s
    `keep` list, every key present verbatim in the RAW file). `DMG_COLUMN_MAP` verified correct — untouched.
  - **Known follow-up (not yet fixed):** R4H's real `Information Precision - 10pts` criterion column is
    currently UNMAPPED in `R4H_COLUMN_MAP` — it has no canonical key, so this criterion does not surface
    in the QA detail table or coaching breakdown. Adding it requires a new canonical key in both
    `R4H_COLUMN_MAP` and the `_transform_qa_source()` `keep` list plus the JS `critKeys` array.
  - **Cosmetic note:** DMG's donut color `#0E7490` is shared with Associated Cab / Circle Taxi, and
    R4H's `#0F766E` with Vermont — legend/tooltips still disambiguate by name; not a data issue.

### Pending / future work
- **Apps Script Monitoring** — Add a new section to `pipeline_monitor.html` (before the incident log) showing per-account Apps Script health: last run time, row count, duration, stale/fail badges. Implementation: `logRunToMasterlist_()` helper in each Apps Script → writes to `GAS_Heartbeat` tab in Masterlist spreadsheet → `check_gas_heartbeat.py` reads it → `pipeline_gas_status.json` → `generate_monitor.py` renders the section. Build when 5+ upsert functions are active and manually checking the Apps Script UI becomes painful.
- **Storage migration to Google Drive** — Move `PB` folder to Google Drive for Desktop. Only 3 bat files + Task Scheduler need path updates. See memory for details.

## Known issues / open questions

- **VIP/Reno Cab transform crash — fixed 2026-06-23** — All 74 `df.get(col, pd.Series(dtype=float))` instances across all AI account transforms replaced with the safe pattern via `fix_footgun.py`. `self_heal.py` runs this automatically before every build attempt, making it impossible for this class of error to recur.

- **Skyline MemoryError — fixed 2026-06-23** — Skyline's dataset reached 14,297 rows × 138 columns, causing `MemoryError` in openpyxl's `write_rows()`. openpyxl holds the entire workbook in RAM before writing. Fixed by switching all 19 pull scripts to `xlsxwriter` (writes row-by-row to disk). A HIGH VOLUME warning fires at 10,000 rows (amber) and 20,000 rows (red/critical) to give advance notice before any future threshold is approached.

- **Britelift Chat uses `britelift_pull.py`** — same script filename as Britelift.
  This is intentional: the bat file `cd`s into the BLC directory first,
  so the correct output file path is set inside the script. Do not rename.
- `df.get(col, pd.Series(dtype=float))` is a footgun — it returns an empty
  Series (length 0) when the column is missing, not a zero-filled Series.
  Any use of this pattern in new code should be replaced with the safe version.

---

## Layout

- **Font:** Arial, sans-serif (entire dashboard)
- **Body text:** 13px / 400 (tables, cards, detail rows)
- **Header h1:** ~20px / 700, white on green gradient
- **Subheads / card labels:** 13px / 600–700, `var(--blue)`
- **CSS variables (`:root`):**
  - `--blue: #004C97` — primary brand blue (card top border, headings, buttons)
  - `--green: #39B54A` — primary brand green (header bar, pass indicators)
  - `--bg: #F4F8F6` — page background
- **Card surface:** `#fff` (white), `border-top: 3px solid var(--blue)`;
  green-accent cards use `border-top-color: var(--green)`
- **Pass / green highlight:** `#DCFCE7` background
- **Warn / orange highlight:** `#FFEDD5` background
- **Info / blue highlight:** `#DBEAFE` background
- **Grid layout:** CSS Grid with fixed `repeat(N, minmax(0, 1fr))` columns
  (8-col for main KPI strip, 6-col for coaching, 9-col for QA summary, etc.)
  No sidebar. Navigation is tab-based across the top.
- **No max-width constraint** — full viewport width.

## Locked tiles / frames

- **Coaching summary tile** — do not change column order or row structure
  in `transform_coaching_logs()` or `refresh_coaching_output()`.
- **QA scorecard frame** — layout locked; only data and account-specific
  mappings should change. Do not restructure `qaRowHtml()`.
- **QA detail table column resizing** — the evaluation table uses a generated
  `colgroup`, always-visible header dividers, drag handles, and localStorage key
  `pacbiz.qa.detail.columnWidths.v1`. Preserve sorting click behavior by keeping
  resize-handle events stopped from propagating.
- **Auto-reload schedule array** — do not change the times in the JS
  `refreshTimes` array without also updating the data freshness
  `schedTimes` array and confirming the Task Scheduler entries match.
- No `// LOCKED` comments are currently in the source — treat the above
  functions as implicitly locked for structural changes.

## Things Claude Code must never change

- Do not modify the `transform_coaching_logs()` or `refresh_coaching_output()` structure.
- Do not change the auto-reload schedule array values without explicit instruction.
- Do not remove or alter the cache-busting meta tags in `<head>`.
- Do not add `df.get(col, pd.Series(dtype=float))` — use the safe missing-column pattern.
- Do not push to Git unless the user explicitly asks.
- Do not commit `.env` files, credential JSON files, or API tokens.
- Do not add `dashboard.py` back to the `git add` line in either bat file — it was intentionally
  removed so scheduled runs cannot overwrite manual fixes to `dashboard.py`.

---

## Deployment — pre-publish checklist

Claude Code must complete ALL of these steps in order
before running any git command. Do not skip any step.

### Step 1 — syntax check
- Run: `py -3 -m py_compile dashboard.py`
- Must return no errors before continuing
- If errors found: stop, fix, re-run check

### Step 2 — test build
- Run: `py -3 dashboard.py`
- Must complete with exit code 0
- Confirm output ends with: `Dashboard generated: masterlist_dashboard.html`
- If build fails: stop, do not proceed to git

### Step 3 — output validation
- Confirm `masterlist_dashboard.html` was updated (file timestamp newer than before the run)
- Search the file and confirm:
  - [ ] File is not suspiciously small (expect ~26 MB)
  - [ ] Contains expected sections: search for `Coaching`, `Skyline`, `Britelift`
  - [ ] Cache-busting meta tags present in `<head>`: `Cache-Control`, `Pragma`, `Expires`
  - [ ] Auto-reload script present before `</body>`: search for `scheduleNextReload`
  - [ ] No Python traceback text visible (search for `Traceback`)

### Step 4 — git hygiene
- Run: `git status`
- Confirm only expected files are modified:
  - `dashboard.py`
  - `masterlist_dashboard.html`
  - `update_coaching_dashboard.bat` (if changed)
  - `update_coaching_dashboard_auto.bat` (if changed)
- If unexpected files are staged: stop and ask before continuing

### Step 5 — pull before push
- Always run: `git pull --rebase --autostash`
- Resolve any conflicts before committing
- Never force push

### Step 6 — commit
- Commit message format: `Update dashboard — [brief reason]`
  - e.g. `Update dashboard — added Trans Iowa QA tile`
  - e.g. `Update dashboard — cache-busting fix`
- Run: `git add dashboard.py masterlist_dashboard.html`
- Run: `git commit -m "..."`
- Run: `git push`

### Step 7 — post-publish confirm
- Confirm push was successful (exit code 0, output shows `main -> main`)
- Note: GitHub Pages takes ~1 minute to reflect changes
- Report back: commit hash + what changed

## Deployment — hard rules

- Never run `git push` if any step above failed
- Never amend or rebase published commits
- Never commit credentials, API keys, service account JSON, or webhook URLs
- If unsure about any staged file — ask first, don't push
- Always report the outcome (pass or fail) for every step

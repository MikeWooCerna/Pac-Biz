# PB Dashboard — project context
# This file is read by Claude Code (auto) and Codex (manual).
# Last updated: 2026-06-21

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
| `requirements.txt` | Python deps: pandas>=2.2.0, openpyxl>=3.1.0, requests>=2.32.0 |
| `test_connection.py` | Google Sheets API connectivity test |
| `pacbiz_logo.png` / `pacbiz_favicon.png` | Branding assets embedded in the HTML at build time |

Python executable: `C:\Users\Mike Woo Cerna\AppData\Local\Programs\Python\Python313\python.exe`

## Pipeline sequence (update_coaching_dashboard_auto.bat)

Each step runs `py -3 <script>.py` inside the account's directory.
If any step returns a non-zero exit code the entire pipeline aborts (`goto :fail`).

```
1.  Coaching          — %COACHING_DIR%\asana_pull.py
2.  M7                — %M7_DIR%\m7_pull.py
3.  Parentis Health   — %PARENTIS_DIR%\parentis_pull.py
4.  Britelift         — %BRITELIFT_DIR%\britelift_pull.py
5.  Britelift Chat    — %BLC_DIR%\britelift_pull.py        ← SAME script name, different dir
6.  RideX             — %RIDEX_DIR%\Ridex_pull.py
7.  Hamilton          — %HAMILTON_DIR%\Hamilton_pull.py
8.  Skyline           — %SKYLINE_DIR%\Skyline_pull.py
9.  VIP               — %VIP_DIR%\vip_pull.py
10. C&H               — %CH_DIR%\ch_pull.py
11. Reno Cab          — %RC_DIR%\rc_pull.py
12. Trans Iowa        — %TI_DIR%\ti_pull.py
13. Data Carz         — %DC_DIR%\dc_pull.py
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
  M7, Parentis, Britelift, Britelift Chat, and RideX are NOT in this list.
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

## GitHub

- **Remote:** https://github.com/MikeWooCerna/Pac-Biz.git
- **Branch:** `main`
- **Live URL:** https://mikewoocerna.github.io/Pac-Biz/masterlist_dashboard.html
- GitHub Pages takes ~1 minute to reflect a push.

## Active work / next steps

Pipeline monitoring system is fully live as of 2026-06-22. See `PIPELINE_MONITOR.md` for full technical reference.

### Changes made 2026-06-23
- **Row count drop alerts** — `pipeline_rowcount_baseline.json` stores last known count per account. After each run, any account that pulled fewer rows than the baseline triggers: (1) an amber warning strip on the monitor, (2) a `↓ N` red badge on the account card, (3) a `COUNT DROP` entry in the incident log. Baseline auto-updates each run.
- **Account card badges** — every passing account card now shows a small `✓` green badge (no drop) or `↓ N` red badge (count dropped).
- **Build and Git Push failures in incident log** — previously only data-source account failures were logged. Now `dashboard.py` failures and `git push` failures also appear as `FAILED` rows in the incident log table.
- **Radar sweep turns red on failure** — sweep arm, trail, tip glow, and center dot switch from purple to red whenever any step fails (account, Build, or Git Push).
- **VIP transform crash fixed** — see Known issues below.

### Pending / future work
- No open work items.

## Known issues / open questions

- **VIP transform crash — fixed 2026-06-23 (commit `478a497`)** — `transform_vip_data` used `df.get(col, pd.Series(dtype=float))` in both the `VIP_CRIT_MAP` and `VIP_EXTRA_CRIT_MAP` loops. When VIP pulled a sheet missing a criterion column, the footgun returned a 0-length Series, causing a `ValueError: Length of values (0) does not match length of index (1000)` crash at Build time. Fixed by replacing with the safe pattern in both loops.

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

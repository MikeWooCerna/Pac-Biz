"""Generates pipeline_monitor.html from pipeline_status.json."""
import json, base64
from datetime import datetime
from pathlib import Path

BASE         = Path(__file__).parent
STATUS_FILE  = BASE / "pipeline_status.json"
LOG_FILE      = BASE / "pipeline_log.json"
BASELINE_FILE = BASE / "pipeline_rowcount_baseline.json"
HEAL_EVENTS           = BASE / "pipeline_heal_events.json"
HIGHVOL_NOTIFIED_FILE = BASE / "pipeline_highvol_notified.json"
OUTPUT_FILE  = BASE / "pipeline_monitor.html"

HIGH_VOLUME_WARN = 10_000   # amber badge + strip
HIGH_VOLUME_CRIT = 20_000   # red badge + strip + email
LOGO_FILE    = BASE / "pacbiz_logo.png"
FAVICON_FILE = BASE / "pacbiz_favicon.png"

DASHBOARD_URL = "https://mikewoocerna.github.io/Pac-Biz/masterlist_dashboard.html"

ACCOUNTS = [
    ("Masterlist",      "masterlist_fetch.py", r"C:\Users\Mike Woo Cerna\Documents\PB\Masterlist\masterlist_cache.csv"),
    ("Coaching",        "asana_pull.py",    r"C:\Users\Mike Woo Cerna\Documents\PB\Coaching\Output\coaching_logs.xlsx"),
    ("M7",              "m7_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\M7\M7_RAW.xlsx"),
    ("DMG",             "dmg_pull.py",      r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\DMG\DMG_RAW.xlsx"),
    ("R4H",             "r4h_pull.py",      r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\R4H\R4H_RAW.xlsx"),
    ("Parentis Health", "parentis_pull.py", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Parentis Health\PARENTIS_RAW.xlsx"),
    ("Britelift",       "britelift_pull.py",r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift\BRITELIFT_RAW.xlsx"),
    ("Britelift Chat",  "britelift_pull.py",r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift Chat\BLC_RAW.xlsx"),
    ("RideX",           "Ridex_pull.py",    r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\RideX\RIDEX_RAW.xlsx"),
    ("Hamilton",        "Hamilton_pull.py", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Hamilton\HAMILTON_RAW.xlsx"),
    ("Skyline",         "Skyline_pull.py",  r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Skyline\SKYLINE_RAW.xlsx"),
    ("VIP",             "vip_pull.py",      r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\VIP\VIP_RAW.xlsx"),
    ("C&H",             "ch_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\C&H\CH_RAW.xlsx"),
    ("Reno Cab",        "rc_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Reno Cab\RC_RAW.xlsx"),
    ("Trans Iowa",      "ti_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Trans Iowa\TI_RAW.xlsx"),
    ("Data Carz",       "dc_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Data Carz\DC_RAW.xlsx"),
    ("Associated Cab",  "ac_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Associated Cab\AC_RAW.xlsx"),
    ("Ollies",          "ol_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Ollies\OL_RAW.xlsx"),
    ("Circle Taxi",     "ct_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Circle Taxi\CT_RAW.xlsx"),
    ("YCOV",            "ycov_pull.py",     r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCOV\YCOV_RAW.xlsx"),
    ("Kelowna",         "kel_pull.py",      r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Kelowna\KEL_RAW.xlsx"),
    ("Vermont",         "vt_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Vermont\VT_RAW.xlsx"),
    ("YCDC",            "ycdc_pull.py",     r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCDC\YCDC_RAW.xlsx"),
    ("Blueline",        "bl_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Blueline\BL_RAW.xlsx"),
]

def get_row_count(file_path):
    try:
        p = Path(file_path)
        if not p.exists():
            return None
        if p.suffix.lower() == ".csv":
            lines = sum(1 for _ in p.open("r", encoding="utf-8", errors="replace"))
            return max(0, lines - 1)
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        count = max(0, (ws.max_row or 1) - 1)
        wb.close()
        return count
    except Exception:
        return None

def logo_b64():
    try:
        return base64.b64encode(LOGO_FILE.read_bytes()).decode()
    except Exception:
        return ""

def favicon_b64():
    try:
        return base64.b64encode(FAVICON_FILE.read_bytes()).decode()
    except Exception:
        return logo_b64()

def fmt_time(iso_str):
    if not iso_str:
        return "&mdash;"
    try:
        dt = datetime.fromisoformat(iso_str)
        return dt.strftime("%I:%M %p")
    except Exception:
        return str(iso_str)[:16]

def load_log():
    try:
        if LOG_FILE.exists():
            return json.loads(LOG_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return []

def save_log(entries):
    try:
        LOG_FILE.write_text(json.dumps(entries[-300:], indent=2, default=str), encoding="utf-8")
    except Exception:
        pass

def load_baseline():
    try:
        if BASELINE_FILE.exists():
            return json.loads(BASELINE_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def save_baseline(data):
    try:
        BASELINE_FILE.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
    except Exception:
        pass

def load_highvol_notified():
    try:
        if HIGHVOL_NOTIFIED_FILE.exists():
            return json.loads(HIGHVOL_NOTIFIED_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def save_highvol_notified(data):
    try:
        HIGHVOL_NOTIFIED_FILE.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
    except Exception:
        pass

def _hv_level(rows):
    """Return 'crit', 'warn', or None based on row count."""
    if rows is None:
        return None
    if rows >= HIGH_VOLUME_CRIT:
        return "crit"
    if rows >= HIGH_VOLUME_WARN:
        return "warn"
    return None

def fmt_log_date(iso_str):
    if not iso_str:
        return "&mdash;"
    try:
        dt = datetime.fromisoformat(iso_str)
        return dt.strftime("%b %d, %Y %I:%M %p").lstrip("0")
    except Exception:
        return str(iso_str)[:16]

def render_log_table(log_entries):
    if not log_entries:
        return '<div class="log-empty">No incidents recorded &mdash; all pipeline runs have been successful.</div>'
    rows = []
    for e in reversed(log_entries[-80:]):
        st  = e.get("status", "")
        pill = ('<span class="lp lp-f">FAILED</span>'          if st == "fail"
                else '<span class="lp lp-d">COUNT DROP</span>'  if st == "count_drop"
                else '<span class="lp lp-s">SOURCE CONFIRMED</span>' if st == "source_confirmed"
                else '<span class="lp lp-h">SELF HEALED</span>' if st == "healed"
                else '<span class="lp lp-v">HIGH VOLUME</span>' if st == "high_volume"
                else '<span class="lp lp-g">GUARDIAN FIX</span>'  if st == "guardian_fix"
                else '<span class="lp lp-g lp-gw">GUARDIAN WARN</span>'  if st == "guardian_warn"
                else '<span class="lp lp-n">NOT REACHED</span>')
        err  = (e.get("error") or "").strip()
        err_html = f'<div class="log-err">{err[:120]}</div>' if err else ""
        rows.append(
            f'<tr><td class="log-td-d">{e.get("date","&mdash;")}</td>'
            f'<td class="log-td-ds"><b>{e.get("account","")}</b>'
            f' <span class="log-sc">&middot; {e.get("script","")}</span>{err_html}</td>'
            f'<td class="log-td-st">{pill}</td></tr>'
        )
    return "\n".join(rows)

def calc_next_schedule():
    from datetime import timedelta
    slots = ['03:00','06:00','11:00','15:00','19:00','22:00']
    now = datetime.now()
    for t in slots:
        h, m = int(t[:2]), int(t[3:])
        candidate = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if candidate > now:
            return candidate.strftime("%I:%M %p").lstrip("0")
    h, m = int(slots[0][:2]), int(slots[0][3:])
    tmr = (now + timedelta(days=1)).replace(hour=h, minute=m, second=0, microsecond=0)
    return tmr.strftime("%I:%M %p").lstrip("0") + " +1d"

def dot(status):
    if status == "pass":    return '<div class="blink-g"></div>'
    if status == "fail":    return '<div class="blink-r"></div>'
    if status == "running": return '<div class="blink-a"></div>'
    return '<div class="blink-n"></div>'

def node_cls(status):
    return {"pass": "n-pass", "fail": "n-fail", "running": "n-warn"}.get(status, "n-blocked")

def title_color(status):
    return {"pass": "#00e87a", "fail": "#ff6060", "running": "#ffaa00",
            "blocked": "#E84500", "pending": "#E84500"}.get(status, "#4a3d7a")

def dot_cls(status):
    return {"pass": "b-g", "fail": "b-r"}.get(status, "b-n")

STATUS_LABELS = {"pass": "PASS", "fail": "FAIL", "running": "RUNNING",
                 "blocked": "NOT REACHED", "pending": "NOT REACHED"}

def render_node(name, script, status, ts, rows, error=None, drop_info=None, hv_lvl=None):
    pill_cls = {"pass": "g", "fail": "r", "blocked": "o", "pending": "o"}.get(status, "")
    rows_str = f"{rows:,}" if rows is not None else "&mdash;"
    status_lbl = STATUS_LABELS.get(status, status.upper())
    if status == "fail":
        if error:
            safe_err = error.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            err_html = f'<div class="node-err">{safe_err}</div>'
        else:
            err_html = '<div class="node-row err-row">Exit code 1 &middot; check logs</div>'
    else:
        err_html = ""
    if drop_info is not None:
        count_badge = f'<span class="meta-pill cnt-drop">&darr; {drop_info.get("drop", 0):,}</span>'
    elif status == "pass":
        count_badge = '<span class="meta-pill cnt-ok">&#10003;</span>'
    else:
        count_badge = ""
    if hv_lvl == "crit":
        vol_badge = '<span class="meta-pill cnt-vol-crit">&#9888; CRITICAL VOL</span>'
    elif hv_lvl == "warn":
        vol_badge = '<span class="meta-pill cnt-vol">&#9651; HIGH VOL</span>'
    else:
        vol_badge = ""
    return f"""<div class="node {node_cls(status)}">
  <div class="node-title" style="color:{title_color(status)};"><span>{name}</span>{dot(status)}</div>
  <div class="node-row"><b class="{dot_cls(status)}"></b>{script} &middot; {ts}</div>
  {err_html}
  <div class="node-meta">
    <span class="meta-pill {pill_cls}">{rows_str} rows</span>
    <span class="meta-pill {pill_cls}">{status_lbl}</span>
    {count_badge}{vol_badge}
  </div>
</div>"""

def generate():
    raw = None
    if STATUS_FILE.exists():
        try:
            raw = json.loads(STATUS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass

    # ── Sync Build timestamp to dashboard file mtime ─────────────────────────
    # If masterlist_dashboard.html is newer than the recorded Build step timestamp
    # (e.g. a heal rebuilt the dashboard after the scheduled run), update the
    # Build step timestamp so the freshness indicator matches the actual file age.
    if raw:
        dashboard_html = BASE / "masterlist_dashboard.html"
        build_step_sync = next((s for s in raw.get("steps", []) if s["account"] == "Build"), None)
        if dashboard_html.exists() and build_step_sync:
            try:
                html_mtime = datetime.fromtimestamp(dashboard_html.stat().st_mtime)
                recorded_ts = datetime.fromisoformat(build_step_sync.get("timestamp", ""))
                if html_mtime > recorded_ts:
                    build_step_sync["timestamp"] = html_mtime.isoformat()
                    if build_step_sync.get("status") == "pass":
                        raw["finished_at"] = html_mtime.isoformat()
                    STATUS_FILE.write_text(json.dumps(raw, indent=2, default=str), encoding="utf-8")
            except Exception:
                pass

    # ── Auto-correct stale Build failure ─────────────────────────────────────
    # If the last recorded status is "failed at Build" but masterlist_dashboard.html
    # was updated AFTER the recorded failure timestamp, a manual rebuild succeeded.
    # Patch the status in memory and write it back so the monitor clears itself.
    if raw and raw.get("status") == "failed" and raw.get("failed_at") == "Build":
        dashboard_html = BASE / "masterlist_dashboard.html"
        build_step = next((s for s in raw.get("steps", []) if s["account"] == "Build"), None)
        fail_ts = None
        if build_step:
            try:
                fail_ts = datetime.fromisoformat(build_step.get("timestamp", ""))
            except Exception:
                pass
        if dashboard_html.exists() and fail_ts:
            html_mtime = datetime.fromtimestamp(dashboard_html.stat().st_mtime)
            if html_mtime > fail_ts:
                # Dashboard was rebuilt after the failure — self-heal the status
                raw["status"]    = "success"
                raw["failed_at"] = None
                if build_step:
                    build_step["status"]    = "pass"
                    build_step["exit_code"] = 0
                    build_step["error"]     = None
                    build_step["timestamp"] = html_mtime.isoformat()
                # Add a Git Push step if not already present
                if not any(s["account"] == "Git Push" for s in raw.get("steps", [])):
                    raw.setdefault("steps", []).append({
                        "account": "Git Push", "script": "git push",
                        "exit_code": 0, "status": "pass",
                        "timestamp": html_mtime.isoformat(), "error": None
                    })
                raw["finished_at"] = html_mtime.isoformat()
                try:
                    STATUS_FILE.write_text(json.dumps(raw, indent=2, default=str), encoding="utf-8")
                except Exception:
                    pass

    now_str     = datetime.now().strftime("%b %d, %Y %I:%M %p")
    run_status  = raw.get("status", "unknown") if raw else "unknown"
    steps_map   = {s["account"]: s for s in raw.get("steps", [])} if raw else {}
    started_at  = fmt_time(raw.get("started_at"))  if raw else "&mdash;"
    finished_at = fmt_time(raw.get("finished_at")) if raw else "&mdash;"
    failed_at   = raw.get("failed_at")             if raw else None
    run_id      = (raw.get("run_id", "")[:16])     if raw else "&mdash;"

    # ISO timestamp for the freshness JS (use Build step timestamp if available)
    build_finish_iso = ""
    if raw:
        build_step = steps_map.get("Build")
        if build_step and build_step.get("status") == "pass":
            build_finish_iso = build_step.get("timestamp", raw.get("finished_at", "")) or ""
        elif run_status == "success":
            build_finish_iso = raw.get("finished_at", "") or ""
    dash_built_str = "&mdash;"
    if build_finish_iso:
        try:
            dash_built_str = datetime.fromisoformat(build_finish_iso).strftime("%b %d, %Y %I:%M %p")
        except Exception:
            dash_built_str = str(build_finish_iso)[:16]

    failed_idx = next((i for i, (n, _, _) in enumerate(ACCOUNTS) if n == failed_at), None)

    account_states = []
    for i, (name, script, xlsx) in enumerate(ACCOUNTS):
        step = steps_map.get(name)
        if step:
            s     = step["status"]
            ts    = fmt_time(step.get("timestamp"))
            error = step.get("error")
        else:
            if failed_idx is not None and i > failed_idx:
                s = "blocked"
            elif run_status == "success":
                s = "pending"
            else:
                s = "pending"
            ts    = "&mdash;"
            error = None
        rows = get_row_count(xlsx) if s == "pass" else None
        account_states.append({"name": name, "script": script, "status": s, "ts": ts, "rows": rows, "error": error})

    passed     = sum(1 for a in account_states if a["status"] == "pass")
    failed_ct  = sum(1 for a in account_states if a["status"] == "fail")
    blocked_ct = sum(1 for a in account_states if a["status"] in ("blocked", "pending"))
    total_rows = sum(a["rows"] for a in account_states if a["rows"] is not None)

    # Append finished-run incidents and count drops to persistent log
    log_entries = load_log()
    if run_status in ("success", "failed") and raw:
        run_id_full = raw.get("run_id", "")
        if run_id_full and not any(e.get("run_id") == run_id_full for e in log_entries):
            run_date = fmt_log_date(raw.get("started_at", ""))
            for a in account_states:
                if a["status"] == "fail":
                    log_entries.append({"run_id": run_id_full, "date": run_date,
                                        "account": a["name"], "script": a["script"],
                                        "status": "fail", "error": a.get("error")})
                elif a["status"] in ("blocked", "pending"):
                    log_entries.append({"run_id": run_id_full, "date": run_date,
                                        "account": a["name"], "script": a["script"],
                                        "status": "not_reached", "error": None})
            # Row count drop detection against baseline
            baseline    = load_baseline()
            new_baseline = dict(baseline)
            for a in account_states:
                if a["rows"] is not None:
                    prev = baseline.get(a["name"])
                    if prev is not None and a["rows"] < prev:
                        drop = prev - a["rows"]
                        log_entries.append({"run_id": run_id_full, "date": run_date,
                                            "account": a["name"], "script": a["script"],
                                            "status": "count_drop", "drop": drop,
                                            "prev_count": prev, "new_count": a["rows"],
                                            "error": f"Dropped {drop:,} rows ({prev:,} → {a['rows']:,})"})
                    new_baseline[a["name"]] = a["rows"]
            # Build and Git Push step failures
            build_s = steps_map.get("Build")
            if build_s and build_s.get("status") == "fail":
                log_entries.append({"run_id": run_id_full, "date": run_date,
                                    "account": "Build", "script": "dashboard.py",
                                    "status": "fail", "error": build_s.get("error")})
            git_s = steps_map.get("Git Push")
            if git_s and git_s.get("status") == "fail":
                log_entries.append({"run_id": run_id_full, "date": run_date,
                                    "account": "Git Push", "script": "git push",
                                    "status": "fail", "error": git_s.get("error")})
            # Incorporate self-heal events from staging file
            try:
                if HEAL_EVENTS.exists():
                    heal_entries = json.loads(HEAL_EVENTS.read_text(encoding="utf-8"))
                    for h in heal_entries:
                        if h.get("run_id") == run_id_full:
                            log_entries.append(h)
                    HEAL_EVENTS.write_text("[]", encoding="utf-8")
            except Exception:
                pass

            save_log(log_entries)
            save_baseline(new_baseline)

    # Collect count drops for the current run (for the alert strip and card badges)
    count_drops = []
    if raw:
        run_id_full = raw.get("run_id", "")
        count_drops = [e for e in log_entries
                       if e.get("run_id") == run_id_full and e.get("status") == "count_drop"]
    drop_by_account = {d["account"]: d for d in count_drops}

    # Remove drops that have already been healed or source-confirmed. Same-run
    # source confirmations are valid when a cleanup intentionally lowers rows.
    for e in log_entries:
        if e.get("status") in ("healed", "source_confirmed"):
            acct = e.get("account")
            if acct in drop_by_account:
                drop_run = drop_by_account[acct].get("run_id", "")
                heal_run = e.get("run_id", "")
                if heal_run >= drop_run:
                    drop_by_account.pop(acct)
    count_drops = [d for d in count_drops if d["account"] in drop_by_account]

    # ── High-volume detection ─────────────────────────────────────────────────
    hv_notify_map   = load_highvol_notified()
    hv_new_notices  = {}   # {account: (rows, level)} — newly crossed threshold this run
    high_vol_list   = []   # [(account_state, level)] — all currently above threshold

    for a in account_states:
        lvl = _hv_level(a["rows"])
        if lvl is None:
            hv_notify_map.pop(a["name"], None)   # dropped back below — clear record
            continue
        high_vol_list.append((a, lvl))
        prev_lvl = hv_notify_map.get(a["name"])
        # Notify on first crossing OR escalation from warn → crit
        if prev_lvl is None or (lvl == "crit" and prev_lvl == "warn"):
            hv_new_notices[a["name"]] = (a["rows"], lvl)
            hv_notify_map[a["name"]] = lvl

    save_highvol_notified(hv_notify_map)

    # Send email for newly-crossed thresholds
    if hv_new_notices:
        try:
            import notify as _notify_hv
            ACCOUNTS_SCRIPT_MAP = {name: script for name, script, _ in ACCOUNTS}
            for acct, (rows, lvl) in hv_new_notices.items():
                thresh = HIGH_VOLUME_CRIT if lvl == "crit" else HIGH_VOLUME_WARN
                _notify_hv.notify_high_volume(acct, rows, thresh, lvl)
        except Exception:
            pass

    # Log newly-crossed thresholds into the incident log (once per crossing)
    if run_status in ("success", "failed") and raw:
        run_id_full = raw.get("run_id", "")
        run_date    = fmt_log_date(raw.get("started_at", ""))
        ACCOUNTS_SCRIPT_MAP = {name: script for name, script, _ in ACCOUNTS}
        if run_id_full:
            for acct, (rows, lvl) in hv_new_notices.items():
                thresh = HIGH_VOLUME_CRIT if lvl == "crit" else HIGH_VOLUME_WARN
                log_entries.append({
                    "run_id":  run_id_full,
                    "date":    run_date,
                    "account": acct,
                    "script":  ACCOUNTS_SCRIPT_MAP.get(acct, ""),
                    "status":  "high_volume",
                    "error":   f"{'CRITICAL' if lvl == 'crit' else 'Warning'}: {rows:,} rows (threshold {thresh:,})",
                })
            if hv_new_notices:
                save_log(log_entries)

    high_vol_by_account = {a["name"]: lvl for a, lvl in high_vol_list}

    # Build HIGH VOLUME warning strip
    vol_html = ""
    if high_vol_list:
        crit_items = [(a, l) for a, l in high_vol_list if l == "crit"]
        warn_items = [(a, l) for a, l in high_vol_list if l == "warn"]
        parts = []
        for a, l in crit_items + warn_items:
            lbl = "CRITICAL" if l == "crit" else "HIGH VOL"
            parts.append(
                f"<b>{a['name']}</b> &#9651;&thinsp;{a['rows']:,} rows "
                f"<span style='opacity:0.65;font-size:11px;'>({lbl})</span>"
            )
        strip_cls = "vol-strip-crit" if crit_items else "vol-strip"
        icon      = "&#9888;" if crit_items else "&#9651;"
        vol_html  = (f'<div class="{strip_cls}">{icon} High volume datasets detected &mdash; '
                     + " &nbsp;&middot;&nbsp; ".join(parts) + "</div>")

    log_table_rows = render_log_table(log_entries)

    # Most recent Guardian run (if any) — surfaced as a small status line under
    # the Git Repository stage node.
    guardian_entries = [e for e in log_entries if e.get("account") == "Guardian"]
    last_guardian = guardian_entries[-1] if guardian_entries else None
    if last_guardian is None:
        guardian_line = ""
    elif last_guardian.get("status") == "guardian_fix":
        guardian_line = f'<div class="stage-detail" style="color:#0e7490;font-size:10px;">&#10003; Guardian fixed · {last_guardian["date"]}</div>'
    elif last_guardian.get("status") == "guardian_warn":
        guardian_line = f'<div class="stage-detail" style="color:#b45309;font-size:10px;">&#9888; Guardian warned · {last_guardian["date"]}</div>'
    else:
        guardian_line = ""

    build_step_data = steps_map.get("Build")
    git_step        = steps_map.get("Git Push")
    build_st  = build_step_data["status"] if build_step_data else ("blocked" if run_status in ("failed", "unknown") else "pending")
    git_st    = (git_step["status"] if git_step
                 else ("pass"    if run_status == "success"
                       else "running" if run_status == "running"
                       else "blocked"))

    if   run_status == "success": hub_st, hub_lbl, hub_color = "pass",    "complete", "#00e87a"
    elif run_status == "failed":  hub_st, hub_lbl, hub_color = "fail",    "stopped",  "#ff3d3d"
    elif run_status == "running": hub_st, hub_lbl, hub_color = "running", "running",  "#ffaa00"
    else:                         hub_st, hub_lbl, hub_color = "blocked", "no data",  "#4a3d7a"

    has_failure = (failed_ct > 0 or run_status == "failed"
                   or build_st == "fail" or git_st == "fail")
    radar_fail_js = "true" if has_failure else "false"

    n_total     = len(account_states)
    n_left      = (n_total + 1) // 2
    left_label  = f"Data Sources 1&ndash;{n_left}"
    right_label = f"Data Sources {n_left + 1}&ndash;{n_total}"
    left_html   = "\n".join(render_node(a["name"], a["script"], a["status"], a["ts"], a["rows"], a.get("error"), drop_by_account.get(a["name"]), high_vol_by_account.get(a["name"])) for a in account_states[:n_left])
    right_html  = "\n".join(render_node(a["name"], a["script"], a["status"], a["ts"], a["rows"], a.get("error"), drop_by_account.get(a["name"]), high_vol_by_account.get(a["name"])) for a in account_states[n_left:])

    warn_html = ""
    if run_status == "failed" and failed_at:
        warn_html = f"""<div class="warn-strip">&#9888; Pipeline stopped at <b>{failed_at}</b> &middot; Exit code 1 &middot; {blocked_ct} step{'s' if blocked_ct != 1 else ''} not reached</div>"""

    drop_html = ""
    if count_drops:
        items = " &nbsp;&middot;&nbsp; ".join(
            f"<b>{d['account']}</b> &darr;&thinsp;{d.get('drop', 0):,} rows "
            f"<span style='opacity:0.65;font-size:11px;'>({d.get('prev_count',0):,} &rarr; {d.get('new_count',0):,})</span>"
            for d in count_drops
        )
        drop_html = f'<div class="drop-strip">&#9660; Row count drop detected &mdash; {items}</div>'

    # Build radar blips JS array
    SHORT_NAMES = {
        "Coaching": "Coaching", "M7": "M7", "DMG": "DMG", "R4H": "R4H", "Parentis Health": "Parentis",
        "Britelift": "Britelift", "Britelift Chat": "BLC", "RideX": "RideX",
        "Hamilton": "Hamilton", "Skyline": "Skyline", "VIP": "VIP", "C&H": "C&H",
        "Reno Cab": "Reno Cab", "Trans Iowa": "Trans Iowa", "Data Carz": "Data Carz",
        "Associated Cab": "Assoc.Cab", "Ollies": "Ollies", "Circle Taxi": "Cir.Taxi",
        "YCOV": "YCOV", "Kelowna": "Kelowna", "Vermont": "Vermont",
        "YCDC": "YCDC", "Blueline": "Blueline", "Masterlist": "Masterlist",
    }
    blips_js_items = []
    for i, a in enumerate(account_states):
        angle = -90 + (i / len(account_states)) * 360
        r_frac = 0.63 if i % 2 == 0 else 0.44
        nm = SHORT_NAMES.get(a["name"], a["name"])
        blips_js_items.append(f'{{ang:{angle:.2f},r:{r_frac},st:"{a["status"]}",nm:"{nm}"}},')
    blips_js = "[\n    " + "\n    ".join(blips_js_items) + "\n  ]"

    logo = logo_b64()
    fav  = favicon_b64()
    favicon_tag = f'<link rel="icon" type="image/png" href="data:image/png;base64,{fav}">' if fav else ''
    logo_img = (f'<img src="data:image/png;base64,{logo}" class="logo-img" alt="PacBiz">'
                if logo else
                '<div class="logo-fallback">PB</div>')

    ok_color  = "#00e87a" if failed_ct == 0 else "#ff9090"
    sources_dot = "blink-g" if failed_ct == 0 else "blink-r"
    next_sched  = calc_next_schedule()
    run_icon  = "&#10003;" if run_status == "success" else ("&#10007;" if run_status == "failed" else "&mdash;")
    run_color = "#00e87a" if run_status == "success" else ("#ff3d3d" if run_status == "failed" else "#7060a0")
    git_detail = {"pass": "pushed", "fail": "failed", "running": "queued…",
                  "blocked": "not reached", "pending": "not reached"}.get(git_st, git_st)

    # Dashboard node inner content
    if build_st == "pass" and build_finish_iso:
        dash_inner = f"""<div class="stage-title">Dashboard</div>
          <div id="dash-freshness" style="margin:5px 0 3px;min-height:20px;display:flex;justify-content:center;align-items:center;font-size:12px;"></div>"""
    else:
        dash_inner = f"""<div class="stage-title">Dashboard</div>
          <div class="stage-status">{dot(build_st)}</div>
          <div class="stage-detail">{"not reached" if build_st in ("blocked","pending") else build_st}</div>"""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
{favicon_tag}
<title>PacBiz Pipeline Monitor</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{background:#0a0018;min-height:100vh;font-family:system-ui,-apple-system,sans-serif;padding:20px;}}
.eco{{background:#05001a;border-radius:12px;padding:1.25rem 1.25rem 1rem;color:#fff;position:relative;overflow:hidden;max-width:1390px;margin:0 auto;}}
.grid-bg{{position:absolute;inset:0;background-image:linear-gradient(rgba(80,0,180,0.07) 1px,transparent 1px),linear-gradient(90deg,rgba(80,0,180,0.07) 1px,transparent 1px);background-size:32px 32px;pointer-events:none;}}
.top-bar{{display:flex;align-items:center;justify-content:space-between;margin-bottom:0.3rem;position:relative;z-index:1;}}
.logo-box{{display:flex;align-items:center;gap:10px;}}
.logo-img{{width:46px;height:46px;border-radius:6px;object-fit:contain;mix-blend-mode:screen;filter:brightness(1.15) saturate(1.1);}}
.logo-fallback{{width:46px;height:46px;border-radius:6px;background:linear-gradient(135deg,#1a0050,#5000b4);display:flex;align-items:center;justify-content:center;font-size:17px;font-weight:700;color:#c0a0ff;flex-shrink:0;}}
.brand-name{{font-size:19px;font-weight:500;color:#c0a0ff;letter-spacing:0.04em;}}
.brand-tag{{font-size:13px;color:#7060a0;margin-top:2px;}}
.page-title{{text-align:center;margin:2px 0 9px;position:relative;z-index:1;}}
.page-title h1{{font-size:27px;font-weight:700;letter-spacing:0.05em;background:linear-gradient(90deg,#9050ff,#c090ff 35%,#00e87a 70%,#80ffcc);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;line-height:1.2;}}
.stat-cards{{display:flex;gap:8px;margin:0 0 10px;position:relative;z-index:1;}}
.stat-card{{flex:1;background:rgba(40,10,90,0.5);border:1px solid rgba(80,0,180,0.35);border-radius:10px;padding:10px 8px;text-align:center;}}
.sc-val{{font-size:22px;font-weight:600;line-height:1;}}
.sc-lbl{{font-size:12px;color:#7060a0;margin-top:5px;}}
.legend{{display:flex;justify-content:center;gap:16px;margin-bottom:9px;}}
.leg-item{{display:flex;align-items:center;gap:5px;font-size:13px;color:#9080c0;}}
.blink-g{{width:7px;height:7px;border-radius:50%;background:#00e87a;animation:blinkG 1.6s ease-in-out infinite;flex-shrink:0;}}
.blink-r{{width:7px;height:7px;border-radius:50%;background:#ff3d3d;animation:blinkR 0.9s ease-in-out infinite;flex-shrink:0;}}
.blink-n{{width:7px;height:7px;border-radius:50%;background:#E84500;animation:blinkO 1.4s ease-in-out infinite;flex-shrink:0;}}
.blink-a{{width:7px;height:7px;border-radius:50%;background:#ffaa00;animation:blinkA 1.2s ease-in-out infinite;flex-shrink:0;}}
@keyframes blinkG{{0%,100%{{opacity:1;box-shadow:0 0 5px #00e87a;}}50%{{opacity:0.3;box-shadow:none;}}}}
@keyframes blinkR{{0%,100%{{opacity:1;box-shadow:0 0 7px #ff3d3d;}}50%{{opacity:0.2;box-shadow:none;}}}}
@keyframes blinkA{{0%,100%{{opacity:1;box-shadow:0 0 5px #ffaa00;}}50%{{opacity:0.3;box-shadow:none;}}}}
@keyframes blinkO{{0%,100%{{opacity:1;box-shadow:0 0 6px #E84500;}}50%{{opacity:0.25;box-shadow:none;}}}}
.main-layout{{display:grid;grid-template-columns:minmax(280px,0.95fr) minmax(328px,350px) minmax(280px,0.95fr);gap:8px;align-items:start;position:relative;z-index:1;}}
.side-col{{display:flex;flex-direction:column;gap:4px;min-width:0;}}
.center-col{{display:flex;flex-direction:column;align-items:center;gap:4px;}}
/* --- Compact account nodes --- */
.node{{border-radius:7px;border:1px solid;padding:4px 7px;min-height:52px;}}
.node-title{{font-size:12px;font-weight:500;margin-bottom:3px;display:flex;align-items:center;justify-content:space-between;gap:4px;}}
.node-title span{{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.node-row{{display:flex;align-items:center;gap:4px;font-size:10px;color:#9080b8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.node-row b{{width:5px;height:5px;border-radius:50%;flex-shrink:0;}}
.err-row{{color:#ff9090;}}
.node-err{{font-size:10px;color:#ff9090;background:rgba(80,0,0,0.3);border-radius:4px;padding:3px 5px;margin-top:3px;font-family:monospace;line-height:1.4;white-space:pre-wrap;word-break:break-word;}}
.node-meta{{display:flex;gap:3px;margin-top:4px;flex-wrap:wrap;}}
.meta-pill{{font-size:10px;padding:1px 6px;border-radius:99px;background:rgba(80,0,180,0.2);color:#7060a0;border:1px solid rgba(80,0,180,0.18);white-space:nowrap;}}
.meta-pill.g{{background:rgba(0,232,122,0.07);color:#30b860;border-color:rgba(0,232,122,0.18);}}
.meta-pill.r{{background:rgba(255,61,61,0.1);color:#ff8080;border-color:rgba(255,61,61,0.2);}}
.meta-pill.o{{background:rgba(232,69,0,0.1);color:#E84500;border-color:rgba(232,69,0,0.3);}}
.b-g{{background:#00e87a;}}.b-r{{background:#ff3d3d;}}.b-n{{background:#E84500;}}
.n-pass{{background:rgba(0,50,25,0.2);border-color:rgba(0,232,122,0.18);}}
.n-fail{{background:rgba(70,0,0,0.35);border-color:rgba(255,61,61,0.38);}}
.n-blocked,.n-pending{{background:rgba(60,18,0,0.18);border-color:rgba(232,69,0,0.4);}}
.n-warn{{background:rgba(70,40,0,0.3);border-color:rgba(255,170,0,0.3);}}
/* --- Center column --- */
.radar-wrap{{position:relative;width:324px;height:324px;flex-shrink:0;}}
.radar-label{{text-align:center;margin-top:4px;}}
.radar-lbl-main{{font-size:13px;font-weight:500;color:#c0a0ff;}}
.radar-lbl-status{{font-size:12px;margin-top:2px;display:flex;align-items:center;justify-content:center;gap:4px;}}
.connector{{width:2px;height:10px;background:linear-gradient(rgba(80,0,180,0.4),rgba(80,0,180,0.1));margin:0 auto;}}
.stage-node{{border-radius:8px;border:1px solid rgba(80,0,180,0.4);background:rgba(40,10,90,0.4);padding:8px 10px;width:100%;text-align:center;}}
.stage-title{{font-size:13px;font-weight:500;color:#c0a0ff;margin-bottom:2px;}}
.stage-status{{display:flex;justify-content:center;margin:3px 0;}}
.stage-detail{{font-size:12px;color:#4a3d7a;}}
.agg-node{{border-radius:10px;border:1px solid rgba(80,0,180,0.47);background:rgba(40,10,90,0.5);padding:9px 11px;width:100%;margin:1px 0;}}
.agg-title{{font-size:13px;font-weight:500;color:#c0a0ff;text-align:center;margin-bottom:6px;}}
.agg-grid{{display:grid;grid-template-columns:1fr 1fr;gap:4px;}}
.agg-item{{text-align:center;background:rgba(80,0,180,0.15);border-radius:5px;padding:5px;}}
.agg-val{{font-size:14px;font-weight:500;}}
.agg-lbl{{font-size:11px;color:#7060a0;margin-top:2px;}}
.dash-node{{display:block;text-decoration:none;border-radius:8px;border:1px solid rgba(80,0,180,0.4);background:rgba(40,10,90,0.4);width:100%;overflow:hidden;transition:border-color 0.2s,background 0.2s;}}
.dash-node:hover{{border-color:rgba(0,232,122,0.5);background:rgba(0,50,25,0.2);}}
.browser-bar{{display:flex;align-items:center;gap:5px;background:rgba(80,0,180,0.25);padding:4px 8px;border-bottom:1px solid rgba(80,0,180,0.3);}}
.bd{{width:7px;height:7px;border-radius:50%;flex-shrink:0;}}
.bd-r{{background:#ff5f56;}}.bd-y{{background:#ffbd2e;}}.bd-g{{background:#27c93f;}}
.browser-url{{flex:1;font-size:10px;color:#7060a0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
.browser-arr{{font-size:11px;color:#5000b4;}}
.dash-body{{padding:7px 10px;text-align:center;}}
.warn-strip{{background:rgba(80,0,0,0.3);border:1px solid rgba(255,61,61,0.27);border-radius:7px;padding:7px 11px;font-size:13px;color:#ff9090;display:flex;align-items:flex-start;gap:6px;margin-top:7px;position:relative;z-index:1;}}
.section-hdr{{font-size:11px;color:#5000b4;letter-spacing:0.06em;text-transform:uppercase;margin-bottom:3px;}}
.divider{{border:none;border-top:1px solid rgba(80,0,180,0.2);margin:7px 0;position:relative;z-index:1;}}
.sig{{text-align:center;font-size:12px;font-weight:600;letter-spacing:0.08em;background:linear-gradient(90deg,#9050ff,#c090ff 35%,#00e87a 70%,#80ffcc);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;padding:6px 0 2px;position:relative;z-index:1;}}
.log-section{{margin-top:14px;position:relative;z-index:1;}}
.log-hdr{{font-size:11px;color:#5000b4;letter-spacing:0.07em;text-transform:uppercase;margin-bottom:7px;}}
.log-wrap{{overflow:auto;max-height:240px;border-radius:8px;border:1px solid rgba(80,0,180,0.22);}}
.log-table{{width:100%;border-collapse:collapse;font-size:11px;}}
.log-table thead tr{{background:rgba(40,10,90,0.6);}}
.log-table th{{position:sticky;top:0;z-index:2;text-align:left;color:#7060a0;font-weight:600;padding:6px 10px;border-bottom:1px solid rgba(80,0,180,0.25);white-space:nowrap;background:#12062a;}}
.log-table td{{padding:5px 10px;border-bottom:1px solid rgba(80,0,180,0.1);vertical-align:top;}}
.log-table tr:last-child td{{border-bottom:none;}}
.log-table tbody tr:hover td{{background:rgba(80,0,180,0.07);}}
.log-td-d{{color:#7060a0;white-space:nowrap;}}
.log-td-ds{{color:#c0a0ff;}}
.log-td-ds b{{font-weight:500;}}
.log-sc{{color:#5000b4;}}
.log-err{{font-size:10px;color:#ff9090;font-family:monospace;margin-top:2px;white-space:pre-wrap;word-break:break-word;}}
.log-td-st{{white-space:nowrap;}}
.lp{{font-size:10px;padding:1px 8px;border-radius:99px;white-space:nowrap;}}
.lp-f{{background:rgba(255,61,61,0.1);color:#ff8080;border:1px solid rgba(255,61,61,0.22);}}
.lp-n{{background:rgba(232,69,0,0.1);color:#E84500;border:1px solid rgba(232,69,0,0.28);}}
.log-empty{{text-align:center;color:#4a3d7a;font-size:12px;padding:14px;background:rgba(40,10,90,0.25);border-radius:8px;border:1px solid rgba(80,0,180,0.18);}}
.drop-strip{{background:rgba(50,30,0,0.35);border:1px solid rgba(255,170,0,0.32);border-radius:7px;padding:7px 12px;font-size:13px;color:#ffcc55;display:flex;align-items:flex-start;gap:6px;margin-top:7px;flex-wrap:wrap;position:relative;z-index:1;}}
.lp-d{{background:rgba(255,170,0,0.1);color:#ffaa00;border:1px solid rgba(255,170,0,0.28);}}
.lp-s{{background:rgba(14,165,233,0.1);color:#7dd3fc;border:1px solid rgba(14,165,233,0.28);}}
.lp-h{{background:rgba(0,232,122,0.1);color:#00e87a;border:1px solid rgba(0,232,122,0.28);}}
.cnt-drop{{background:rgba(232,69,0,0.15);color:#ff7040;border:1px solid rgba(232,69,0,0.38);font-weight:600;}}
.cnt-ok{{background:rgba(0,232,122,0.06);color:#20a060;border:1px solid rgba(0,232,122,0.18);}}
.cnt-vol{{background:rgba(255,170,0,0.12);color:#ffaa00;border:1px solid rgba(255,170,0,0.38);font-weight:600;}}
.cnt-vol-crit{{background:rgba(255,61,61,0.12);color:#ff8080;border:1px solid rgba(255,61,61,0.38);font-weight:600;}}
.vol-strip{{background:rgba(60,30,0,0.4);border:1px solid rgba(255,170,0,0.38);border-radius:7px;padding:7px 12px;font-size:13px;color:#ffcc55;display:flex;align-items:flex-start;gap:6px;margin-top:7px;flex-wrap:wrap;position:relative;z-index:1;}}
.vol-strip-crit{{background:rgba(60,0,0,0.4);border:1px solid rgba(255,61,61,0.38);border-radius:7px;padding:7px 12px;font-size:13px;color:#ff9090;display:flex;align-items:flex-start;gap:6px;margin-top:7px;flex-wrap:wrap;position:relative;z-index:1;}}
.lp-v{{background:rgba(255,170,0,0.1);color:#ffaa00;border:1px solid rgba(255,170,0,0.28);}}
.lp-g{{background:#0e7490;color:#fff;}}
.lp-gw{{background:#92400e;color:#fff;}}
@media(max-width:1180px){{
  body{{padding:12px;}}
  .main-layout{{display:grid;grid-template-columns:repeat(2,minmax(280px,390px));gap:8px;justify-content:center;}}
  .center-col{{grid-column:1 / -1;grid-row:1;width:100%;}}
  .radar-wrap{{width:240px;height:240px;flex-shrink:1;align-self:center;}}
  .radar-wrap canvas{{width:240px!important;height:240px!important;}}
  .side-col{{display:flex;flex-direction:column;gap:4px;width:100%;}}
  .side-col .section-hdr{{display:block;}}
  .stat-cards{{flex-wrap:wrap;}}
  .stat-card{{flex:1 1 80px;min-width:70px;}}
  .log-wrap{{max-height:220px;}}
}}
@media(max-width:720px){{
  .main-layout{{display:flex;flex-direction:column;gap:5px;}}
  .center-col{{order:-1;width:100%;}}
  .side-col{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:5px;}}
  .side-col .section-hdr{{display:none;}}
  .main-layout>.side-col:first-child::before{{content:'Data Sources';display:block;grid-column:span 2;font-size:11px;color:#5000b4;letter-spacing:0.06em;text-transform:uppercase;margin-bottom:3px;}}
}}
@media(max-width:480px){{
  body{{padding:8px;}}
  .eco{{padding:0.75rem 0.75rem 0.6rem;}}
  .page-title h1{{font-size:19px;}}
  .side-col{{grid-template-columns:1fr;}}
  .main-layout>.side-col:first-child::before{{grid-column:span 1;}}
  .radar-wrap{{width:200px;height:200px;}}
  .radar-wrap canvas{{width:200px!important;height:200px!important;}}
  .node-row,.meta-pill,.log-err,.browser-url{{font-size:11px;}}
  .agg-val{{font-size:12px;}}
  .brand-name{{font-size:16px;}}
}}
@media(prefers-reduced-motion:reduce){{
  .blink-g,.blink-r,.blink-a{{animation:none;}}
  canvas{{display:none;}}
}}
</style>
</head>
<body>
<div class="eco">
  <div class="grid-bg"></div>

  <div class="top-bar">
    <div class="logo-box">
      {logo_img}
      <div>
        <div class="brand-name">PacBiz</div>
        <div class="brand-tag">Pipeline Monitor</div>
      </div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:13px;color:#7060a0;">Run {run_id}</div>
      <div style="font-size:12px;color:#4a3d7a;margin-top:2px;">Dashboard: {dash_built_str}</div>
      <div style="font-size:12px;color:#4a3d7a;margin-top:2px;">Monitor: {now_str}</div>
      <div style="font-size:12px;color:#4a3d7a;margin-top:2px;" id="cd">Auto-refresh in 60s</div>
    </div>
  </div>

  <div class="page-title">
    <h1>Reports Pipeline Ecosystem</h1>
  </div>

  <div class="stat-cards">
    <div class="stat-card"><div class="sc-val" style="color:#00e87a;">{passed}</div><div class="sc-lbl">Passed</div></div>
    <div class="stat-card"><div class="sc-val" style="color:#ff3d3d;">{failed_ct}</div><div class="sc-lbl">Failed</div></div>
    <div class="stat-card"><div class="sc-val" style="color:#E84500;">{blocked_ct}</div><div class="sc-lbl">Not Reached</div></div>
    <div class="stat-card"><div class="sc-val" style="color:#c0a0ff;">{total_rows:,}</div><div class="sc-lbl">Rows Pulled</div></div>
    <div class="stat-card"><div class="sc-val" style="color:{run_color};">{run_icon}</div><div class="sc-lbl">{run_status}</div></div>
  </div>

  <div class="divider"></div>

  <div class="legend">
    <div class="leg-item"><div class="blink-g"></div> Pass</div>
    <div class="leg-item"><div class="blink-r"></div> Failed</div>
    <div class="leg-item"><div class="blink-a"></div> Running</div>
    <div class="leg-item"><div class="blink-n"></div> Not reached</div>
  </div>

  <div class="main-layout">
    <div class="side-col">
      <div class="section-hdr">{left_label}</div>
      {left_html}
    </div>

    <div class="center-col">
      <div class="radar-wrap">
        <canvas id="pb-radar" width="324" height="324" style="width:324px;height:324px;"></canvas>
      </div>
      <div class="radar-label">
        <div class="radar-lbl-main">Pipeline / Container Monitoring</div>
        <div class="radar-lbl-status">{dot(hub_st)}<span style="color:{hub_color};font-size:12px;">{hub_lbl}</span></div>
      </div>

      <div class="connector"></div>

      <div class="agg-node">
        <div class="agg-title">Data Engine / Processing</div>
        <div class="agg-grid">
          <div class="agg-item"><div class="agg-val" style="color:#00e87a;">{total_rows:,}</div><div class="agg-lbl">total rows</div></div>
          <div class="agg-item">
            <div class="agg-val" style="display:flex;align-items:center;justify-content:center;gap:5px;">
              <div class="{sources_dot}"></div>
              <span style="color:{ok_color};">{passed}/{passed+failed_ct+blocked_ct}</span>
            </div>
            <div class="agg-lbl">sources ok</div>
          </div>
          <div class="agg-item"><div class="agg-val" style="color:#c0a0ff;font-size:12px;">{started_at}</div><div class="agg-lbl">started</div></div>
          <div class="agg-item"><div class="agg-val" style="color:#c0a0ff;font-size:12px;">{finished_at}</div><div class="agg-lbl">finished</div></div>
          <div class="agg-item" style="grid-column:span 2;"><div class="agg-val" style="color:#ffaa00;font-size:12px;">{next_sched}</div><div class="agg-lbl">next schedule</div></div>
        </div>
      </div>

      <div class="connector"></div>

      <a href="{DASHBOARD_URL}" target="_blank" rel="noopener" class="dash-node">
        <div class="browser-bar">
          <span class="bd bd-r"></span><span class="bd bd-y"></span><span class="bd bd-g"></span>
          <span class="browser-url">PacBiz Dashboard</span>
          <span class="browser-arr">&#8599;</span>
        </div>
        <div class="dash-body">
          {dash_inner}
        </div>
      </a>

      <div class="connector"></div>

      <div class="stage-node">
        <div class="stage-title">Git Repository</div>
        <div class="stage-status">{dot(git_st)}</div>
        <div class="stage-detail">{git_detail}</div>
        {guardian_line}
      </div>
    </div>

    <div class="side-col">
      <div class="section-hdr">{right_label}</div>
      {right_html}
    </div>
  </div>

  {warn_html}
  {drop_html}
  {vol_html}

  <div class="divider" style="margin-top:12px;"></div>
  <div class="log-section">
    <div class="log-hdr">Incident Log</div>
    <div class="log-wrap">
      <table class="log-table">
        <thead><tr><th>Date</th><th>Dataset</th><th>Status</th></tr></thead>
        <tbody>{log_table_rows}</tbody>
      </table>
    </div>
  </div>

  <div class="divider" style="margin-top:12px;"></div>
  <div class="sig">Developed for Pac-Biz Reporting &nbsp;&middot;&nbsp; MCerna &nbsp;&middot;&nbsp; v26.06.22</div>
</div>

<script>
// ── Radar canvas ─────────────────────────────────────────────────────────────
(function() {{
  var c = document.getElementById('pb-radar');
  if (!c || !c.getContext) return;
  var ctx = c.getContext('2d');
  var W = c.width, H = c.height, cx = W/2, cy = H/2;
  var R = W/2 - 2;
  var wrapW = c.parentElement ? c.parentElement.offsetWidth : 0;
  if (wrapW > 0 && wrapW < W) {{ c.width = wrapW; c.height = wrapW; W = c.width; H = c.height; cx = W/2; cy = H/2; R = W/2 - 2; }}
  var ang = -Math.PI/2;
  var TRAIL = Math.PI * 1.35;
  var time = 0;
  var FAIL = {radar_fail_js};

  var blips = {blips_js};

  function draw() {{
    ctx.clearRect(0, 0, W, H);
    ctx.save();
    ctx.beginPath();
    ctx.arc(cx, cy, R, 0, Math.PI*2);
    ctx.clip();

    // background
    var bg = ctx.createRadialGradient(cx, cy, 0, cx, cy, R);
    bg.addColorStop(0, '#150035');
    bg.addColorStop(1, '#05001a');
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, W, H);

    // sweep trail
    var steps = 90;
    var trailRgb = FAIL ? '200,30,30' : '110,50,255';
    for (var i = 0; i < steps; i++) {{
      var frac = i / steps;
      var a0 = ang - TRAIL * (1 - frac);
      var a1 = ang - TRAIL * (1 - (i + 1) / steps);
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, R, a0, a1);
      ctx.closePath();
      ctx.fillStyle = 'rgba(' + trailRgb + ',' + (frac * frac * frac * 0.45) + ')';
      ctx.fill();
    }}

    // grid lines
    ctx.strokeStyle = 'rgba(110,50,255,0.18)';
    ctx.lineWidth = 0.5;
    [0, Math.PI/4, Math.PI/2, 3*Math.PI/4, Math.PI, 5*Math.PI/4, 3*Math.PI/2, 7*Math.PI/4].forEach(function(a) {{
      ctx.beginPath(); ctx.moveTo(cx, cy);
      ctx.lineTo(cx + Math.cos(a)*R, cy + Math.sin(a)*R); ctx.stroke();
    }});

    // rings
    [0.33, 0.66, 1.0].forEach(function(f) {{
      ctx.beginPath(); ctx.arc(cx, cy, R*f, 0, Math.PI*2);
      ctx.strokeStyle = f===1 ? 'rgba(130,60,255,0.65)' : 'rgba(110,50,255,0.28)';
      ctx.lineWidth = f===1 ? 1.5 : 0.75; ctx.stroke();
    }});

    // tick marks
    for (var t = 0; t < 48; t++) {{
      var ta = (t/48)*Math.PI*2, iLen = t%4===0 ? 7 : 4;
      ctx.beginPath();
      ctx.moveTo(cx+Math.cos(ta)*(R-iLen), cy+Math.sin(ta)*(R-iLen));
      ctx.lineTo(cx+Math.cos(ta)*R, cy+Math.sin(ta)*R);
      ctx.strokeStyle='rgba(130,60,255,0.45)'; ctx.lineWidth=1; ctx.stroke();
    }}

    // blips
    blips.forEach(function(b, idx) {{
      var ba = b.ang * Math.PI / 180;
      var bx = cx + Math.cos(ba) * R * b.r;
      var by = cy + Math.sin(ba) * R * b.r;
      var pulse = 0.5 + 0.5 * Math.sin(time * 2.2 + idx * 0.7);
      var col, sz, glow_r;
      if (b.st === 'pass')    {{ col=[0,232,122];   sz=3;   glow_r=10; }}
      else if (b.st==='fail') {{ col=[255,61,61];   sz=3.5; glow_r=11; }}
      else if (b.st==='running') {{ col=[255,170,0]; sz=3; glow_r=10; }}
      else if (b.st==='blocked'||b.st==='pending') {{ col=[232,69,0]; sz=3; glow_r=9; }}
      else                    {{ col=[74,61,122];   sz=2;   glow_r=0;  pulse=0.25; }}
      if (glow_r > 0) {{
        var g = ctx.createRadialGradient(bx,by,0,bx,by,glow_r);
        g.addColorStop(0,'rgba('+col+','+(pulse*0.55)+')');
        g.addColorStop(1,'rgba('+col+',0)');
        ctx.fillStyle=g; ctx.beginPath(); ctx.arc(bx,by,glow_r,0,Math.PI*2); ctx.fill();
      }}
      ctx.beginPath(); ctx.arc(bx,by,sz,0,Math.PI*2);
      ctx.fillStyle='rgba('+col+','+(0.45+pulse*0.55)+')'; ctx.fill();

      // label
      var isNotReached = b.st==='blocked'||b.st==='pending';
      var isFail = b.st==='fail', isRun = b.st==='running';
      var fontSize = (isFail||isRun||isNotReached) ? 8 : 7;
      var labelAlpha = isNotReached ? (0.5+pulse*0.3) : (isFail ? (0.7+pulse*0.3) : 0.55);
      ctx.font = 'bold '+fontSize+'px system-ui,sans-serif';
      ctx.fillStyle = 'rgba('+col+','+labelAlpha+')';
      var labelOffset = R * b.r + 14;
      var lx = cx + Math.cos(ba) * labelOffset;
      var ly = cy + Math.sin(ba) * labelOffset;
      var cosA = Math.cos(ba);
      ctx.textAlign = cosA > 0.2 ? 'left' : (cosA < -0.2 ? 'right' : 'center');
      ctx.textBaseline = 'middle';
      if (isFail || isRun || isNotReached) {{
        ctx.shadowBlur = 6; ctx.shadowColor = 'rgba('+col+',0.7)';
      }}
      ctx.fillText(b.nm, lx, ly);
      ctx.shadowBlur = 0;
    }});

    // sweep arm glow tip
    var tipX = cx + Math.cos(ang)*R*0.9, tipY = cy + Math.sin(ang)*R*0.9;
    var glow = ctx.createRadialGradient(tipX,tipY,0,tipX,tipY,22);
    glow.addColorStop(0, FAIL ? 'rgba(255,60,60,0.65)'   : 'rgba(180,100,255,0.65)');
    glow.addColorStop(1, FAIL ? 'rgba(180,20,20,0)'      : 'rgba(100,40,255,0)');
    ctx.fillStyle=glow; ctx.beginPath(); ctx.arc(tipX,tipY,22,0,Math.PI*2); ctx.fill();

    // sweep arm
    ctx.beginPath(); ctx.moveTo(cx,cy);
    ctx.lineTo(cx+Math.cos(ang)*R, cy+Math.sin(ang)*R);
    ctx.strokeStyle = FAIL ? 'rgba(255,70,70,0.95)'  : 'rgba(200,130,255,0.95)'; ctx.lineWidth=1.5;
    ctx.shadowBlur=10;
    ctx.shadowColor  = FAIL ? 'rgba(255,30,30,0.85)' : 'rgba(180,80,255,0.8)';
    ctx.stroke(); ctx.shadowBlur=0;

    // center dot
    ctx.beginPath(); ctx.arc(cx,cy,3.5,0,Math.PI*2);
    ctx.fillStyle = FAIL ? 'rgba(255,100,100,0.9)' : 'rgba(210,150,255,0.9)'; ctx.fill();

    ctx.restore();
    ang += 0.012; time += 0.022;
    requestAnimationFrame(draw);
  }}
  draw();
}})();

// ── Dashboard freshness ───────────────────────────────────────────────────────
(function() {{
  var BUILD_ISO = "{build_finish_iso}";
  var el = document.getElementById('dash-freshness');
  if (!el || !BUILD_ISO) return;
  var buildTs = new Date(BUILD_ISO);
  if (isNaN(buildTs)) return;
  var schedTimes = ['03:00','06:00','11:00','15:00','19:00','22:00'];
  function update() {{
    var now = new Date();
    var lastSched = null;
    schedTimes.forEach(function(t) {{
      var p = t.split(':'), d = new Date(now);
      d.setHours(+p[0], +p[1], 0, 0);
      if (d <= now && (!lastSched || d > lastSched)) lastSched = d;
    }});
    if (!lastSched) {{
      var p = schedTimes[schedTimes.length-1].split(':');
      lastSched = new Date(now);
      lastSched.setDate(lastSched.getDate()-1);
      lastSched.setHours(+p[0], +p[1], 0, 0);
    }}
    var isLive = buildTs >= lastSched;
    var mins = Math.round((now - buildTs) / 60000);
    var hrs = Math.floor(mins / 60);
    var age = hrs > 0 ? hrs+'h '+(mins%60)+'m' : mins+'m';
    if (isLive) {{
      el.innerHTML = '<span style="display:inline-flex;align-items:center;gap:5px;">'
        +'<span style="width:8px;height:8px;border-radius:50%;background:#16A34A;display:inline-block;animation:blinkG 2s ease-in-out infinite;"></span>'
        +'<span style="color:#16A34A;font-size:12px;font-weight:500;">Live Data</span>'
        +'<span style="color:#4a3d7a;font-size:11px;">&nbsp;&mdash;&nbsp;'+age+'</span>'
        +'</span>';
    }} else {{
      el.innerHTML = '<span style="display:inline-flex;align-items:center;gap:5px;">'
        +'<span style="width:8px;height:8px;border-radius:50%;background:#D97706;display:inline-block;animation:blinkA 2s ease-in-out infinite;"></span>'
        +'<span style="color:#D97706;font-size:12px;font-weight:500;">Stale Data</span>'
        +'<span style="color:#4a3d7a;font-size:11px;">&nbsp;&mdash;&nbsp;'+age+'</span>'
        +'</span>';
    }}
  }}
  update();
  setInterval(update, 60000);
}})();

// ── Auto-refresh countdown ────────────────────────────────────────────────────
(function() {{
  var s = 60, el = document.getElementById('cd');
  function tick() {{
    if (s <= 0) {{ location.reload(); return; }}
    if (el) el.textContent = 'Auto-refresh in ' + s + 's';
    s--; setTimeout(tick, 1000);
  }}
  tick();
}})();
</script>
</body>
</html>"""

    OUTPUT_FILE.write_text(html, encoding="utf-8")
    print(f"Pipeline monitor generated: {OUTPUT_FILE.name}")

generate()

# Run drop investigation — sends email only when new drops are found
try:
    import diagnose_drops
    diagnose_drops.run()
except Exception as _dd_err:
    print(f"[diagnose_drops] Error: {_dd_err}")

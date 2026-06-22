"""Generates pipeline_monitor.html from pipeline_status.json."""
import json, base64
from datetime import datetime
from pathlib import Path

BASE         = Path(__file__).parent
STATUS_FILE  = BASE / "pipeline_status.json"
OUTPUT_FILE  = BASE / "pipeline_monitor.html"
LOGO_FILE    = BASE / "pacbiz_logo.png"
FAVICON_FILE = BASE / "pacbiz_favicon.png"

DASHBOARD_URL = "https://mikewoocerna.github.io/Pac-Biz/masterlist_dashboard.html"

ACCOUNTS = [
    ("Coaching",        "asana_pull.py",    r"C:\Users\Mike Woo Cerna\Documents\PB\Coaching\Output\coaching_logs.xlsx"),
    ("M7",              "m7_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\M7\M7_RAW.xlsx"),
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
    ("Masterlist",      "masterlist_fetch.py", r"C:\Users\Mike Woo Cerna\Documents\PB\Masterlist\masterlist_cache.csv"),
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

def calc_next_schedule():
    from datetime import timedelta
    slots = ['03:30','06:30','11:30','15:30','19:30','22:30']
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
    return {"pass": "#00e87a", "fail": "#ff6060", "running": "#ffaa00"}.get(status, "#4a3d7a")

def dot_cls(status):
    return {"pass": "b-g", "fail": "b-r"}.get(status, "b-n")

def render_node(name, script, status, ts, rows, error=None):
    pill_cls = {"pass": "g", "fail": "r"}.get(status, "")
    rows_str = f"{rows:,}" if rows is not None else "&mdash;"
    if status == "fail":
        if error:
            safe_err = error.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            err_html = f'<div class="node-err">{safe_err}</div>'
        else:
            err_html = '<div class="node-row err-row">Exit code 1 &middot; check logs</div>'
    else:
        err_html = ""
    return f"""<div class="node {node_cls(status)}">
  <div class="node-title" style="color:{title_color(status)};"><span>{name}</span>{dot(status)}</div>
  <div class="node-row"><b class="{dot_cls(status)}"></b>{script} &middot; {ts}</div>
  {err_html}
  <div class="node-meta">
    <span class="meta-pill {pill_cls}">{rows_str} rows</span>
    <span class="meta-pill {pill_cls}">{status.upper()}</span>
  </div>
</div>"""

def generate():
    raw = None
    if STATUS_FILE.exists():
        try:
            raw = json.loads(STATUS_FILE.read_text(encoding="utf-8"))
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
                s = "blocked"
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

    build_step_data = steps_map.get("Build")
    git_step        = steps_map.get("Git Push")
    build_st  = build_step_data["status"] if build_step_data else ("blocked" if run_status in ("failed", "unknown") else "pending")
    git_st    = git_step["status"]        if git_step         else ("blocked" if run_status != "success"            else "pass")

    if   run_status == "success": hub_st, hub_lbl, hub_color = "pass",    "complete", "#00e87a"
    elif run_status == "failed":  hub_st, hub_lbl, hub_color = "fail",    "stopped",  "#ff3d3d"
    elif run_status == "running": hub_st, hub_lbl, hub_color = "running", "running",  "#ffaa00"
    else:                         hub_st, hub_lbl, hub_color = "blocked", "no data",  "#4a3d7a"

    n_left      = 11
    n_total     = len(account_states)
    left_label  = f"Data Sources 1&ndash;{n_left}"
    right_label = f"Data Sources {n_left + 1}&ndash;{n_total}"
    left_html   = "\n".join(render_node(a["name"], a["script"], a["status"], a["ts"], a["rows"], a.get("error")) for a in account_states[:n_left])
    right_html  = "\n".join(render_node(a["name"], a["script"], a["status"], a["ts"], a["rows"], a.get("error")) for a in account_states[n_left:])

    warn_html = ""
    if run_status == "failed" and failed_at:
        warn_html = f"""<div class="warn-strip">&#9888; Pipeline stopped at <b>{failed_at}</b> &middot; Exit code 1 &middot; {blocked_ct} step{'s' if blocked_ct != 1 else ''} not reached</div>"""

    # Build radar blips JS array
    SHORT_NAMES = {
        "Coaching": "Coaching", "M7": "M7", "Parentis Health": "Parentis",
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
    git_detail = "pushed" if git_st == "pass" else ("failed" if git_st == "fail" else "not reached")

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
.blink-n{{width:7px;height:7px;border-radius:50%;background:#4a3d7a;flex-shrink:0;}}
.blink-a{{width:7px;height:7px;border-radius:50%;background:#ffaa00;animation:blinkA 1.2s ease-in-out infinite;flex-shrink:0;}}
@keyframes blinkG{{0%,100%{{opacity:1;box-shadow:0 0 5px #00e87a;}}50%{{opacity:0.3;box-shadow:none;}}}}
@keyframes blinkR{{0%,100%{{opacity:1;box-shadow:0 0 7px #ff3d3d;}}50%{{opacity:0.2;box-shadow:none;}}}}
@keyframes blinkA{{0%,100%{{opacity:1;box-shadow:0 0 5px #ffaa00;}}50%{{opacity:0.3;box-shadow:none;}}}}
.main-layout{{display:grid;grid-template-columns:1fr 348px 1fr;gap:7px;align-items:start;position:relative;z-index:1;}}
.side-col{{display:flex;flex-direction:column;gap:3px;}}
.center-col{{display:flex;flex-direction:column;align-items:center;gap:4px;}}
/* --- Compact account nodes --- */
.node{{border-radius:7px;border:1px solid;padding:5px 8px;}}
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
.b-g{{background:#00e87a;}}.b-r{{background:#ff3d3d;}}.b-n{{background:#4a3d7a;}}
.n-pass{{background:rgba(0,50,25,0.2);border-color:rgba(0,232,122,0.18);}}
.n-fail{{background:rgba(70,0,0,0.35);border-color:rgba(255,61,61,0.38);}}
.n-blocked,.n-pending{{background:rgba(25,10,55,0.4);border-color:rgba(60,45,110,0.3);}}
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
      <div style="font-size:12px;color:#4a3d7a;margin-top:2px;">Built: {now_str}</div>
      <div style="font-size:12px;color:#4a3d7a;margin-top:2px;" id="cd">Auto-refresh in 60s</div>
    </div>
  </div>

  <div class="page-title">
    <h1>Reports Pipeline Ecosystem</h1>
  </div>

  <div class="stat-cards">
    <div class="stat-card"><div class="sc-val" style="color:#00e87a;">{passed}</div><div class="sc-lbl">Passed</div></div>
    <div class="stat-card"><div class="sc-val" style="color:#ff3d3d;">{failed_ct}</div><div class="sc-lbl">Failed</div></div>
    <div class="stat-card"><div class="sc-val" style="color:#4a3d7a;">{blocked_ct}</div><div class="sc-lbl">Not Reached</div></div>
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

      <a href="{DASHBOARD_URL}" target="_blank" class="dash-node">
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
      </div>
    </div>

    <div class="side-col">
      <div class="section-hdr">{right_label}</div>
      {right_html}
    </div>
  </div>

  {warn_html}
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
  var ang = -Math.PI/2;
  var TRAIL = Math.PI * 1.35;
  var time = 0;

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
    for (var i = 0; i < steps; i++) {{
      var frac = i / steps;
      var a0 = ang - TRAIL * (1 - frac);
      var a1 = ang - TRAIL * (1 - (i + 1) / steps);
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, R, a0, a1);
      ctx.closePath();
      ctx.fillStyle = 'rgba(110,50,255,' + (frac * frac * frac * 0.45) + ')';
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
      var isDim = b.st==='blocked'||b.st==='pending';
      var isFail = b.st==='fail', isRun = b.st==='running';
      var fontSize = (isFail||isRun) ? 8 : 7;
      var labelAlpha = isDim ? 0.2 : (isFail ? (0.7+pulse*0.3) : 0.55);
      ctx.font = 'bold '+fontSize+'px system-ui,sans-serif';
      ctx.fillStyle = 'rgba('+col+','+labelAlpha+')';
      var labelOffset = R * b.r + 14;
      var lx = cx + Math.cos(ba) * labelOffset;
      var ly = cy + Math.sin(ba) * labelOffset;
      var cosA = Math.cos(ba);
      ctx.textAlign = cosA > 0.2 ? 'left' : (cosA < -0.2 ? 'right' : 'center');
      ctx.textBaseline = 'middle';
      if (isFail || isRun) {{
        ctx.shadowBlur = 6; ctx.shadowColor = 'rgba('+col+',0.7)';
      }}
      ctx.fillText(b.nm, lx, ly);
      ctx.shadowBlur = 0;
    }});

    // sweep arm glow tip
    var tipX = cx + Math.cos(ang)*R*0.9, tipY = cy + Math.sin(ang)*R*0.9;
    var glow = ctx.createRadialGradient(tipX,tipY,0,tipX,tipY,22);
    glow.addColorStop(0,'rgba(180,100,255,0.65)'); glow.addColorStop(1,'rgba(100,40,255,0)');
    ctx.fillStyle=glow; ctx.beginPath(); ctx.arc(tipX,tipY,22,0,Math.PI*2); ctx.fill();

    // sweep arm
    ctx.beginPath(); ctx.moveTo(cx,cy);
    ctx.lineTo(cx+Math.cos(ang)*R, cy+Math.sin(ang)*R);
    ctx.strokeStyle='rgba(200,130,255,0.95)'; ctx.lineWidth=1.5;
    ctx.shadowBlur=10; ctx.shadowColor='rgba(180,80,255,0.8)'; ctx.stroke(); ctx.shadowBlur=0;

    // center dot
    ctx.beginPath(); ctx.arc(cx,cy,3.5,0,Math.PI*2);
    ctx.fillStyle='rgba(210,150,255,0.9)'; ctx.fill();

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

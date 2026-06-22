"""Generates pipeline_monitor.html from pipeline_status.json."""
import json, base64
from datetime import datetime
from pathlib import Path

BASE        = Path(__file__).parent
STATUS_FILE = BASE / "pipeline_status.json"
OUTPUT_FILE = BASE / "pipeline_monitor.html"
LOGO_FILE   = BASE / "pacbiz_logo.png"

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
]

def get_row_count(xlsx_path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
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

def fmt_time(iso_str):
    if not iso_str:
        return "&mdash;"
    try:
        dt = datetime.fromisoformat(iso_str)
        return dt.strftime("%I:%M %p")
    except Exception:
        return str(iso_str)[:16]

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

def render_node(name, script, status, ts, rows):
    pill_cls  = {"pass": "g", "fail": "r"}.get(status, "")
    rows_str  = f"{rows:,}" if rows is not None else "&mdash;"
    err_row   = '<div class="node-row" style="color:#ff9090;">Exit code 1 &middot; check logs</div>' if status == "fail" else ""
    return f"""<div class="node {node_cls(status)}">
  <div class="node-title" style="color:{title_color(status)};"><span>{name}</span>{dot(status)}</div>
  <div class="node-rows">
    <div class="node-row"><b class="{dot_cls(status)}"></b>{script} &middot; {ts}</div>
    {err_row}
  </div>
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

    # Determine status and blocked boundary
    failed_idx = next((i for i, (n, _, _) in enumerate(ACCOUNTS) if n == failed_at), None)

    account_states = []
    for i, (name, script, xlsx) in enumerate(ACCOUNTS):
        step = steps_map.get(name)
        if step:
            s  = step["status"]
            ts = fmt_time(step.get("timestamp"))
        else:
            if failed_idx is not None and i > failed_idx:
                s = "blocked"
            elif run_status == "success":
                s = "blocked"
            else:
                s = "pending"
            ts = "&mdash;"
        rows = get_row_count(xlsx) if s == "pass" else None
        account_states.append({"name": name, "script": script, "status": s, "ts": ts, "rows": rows})

    passed     = sum(1 for a in account_states if a["status"] == "pass")
    failed_ct  = sum(1 for a in account_states if a["status"] == "fail")
    blocked_ct = sum(1 for a in account_states if a["status"] in ("blocked", "pending"))
    total_rows = sum(a["rows"] for a in account_states if a["rows"] is not None)

    build_step = steps_map.get("Build")
    git_step   = steps_map.get("Git Push")
    build_st   = build_step["status"] if build_step else ("blocked" if run_status in ("failed", "unknown") else "pending")
    git_st     = git_step["status"]   if git_step   else ("blocked" if run_status != "success"             else "pass")

    if   run_status == "success": hub_st, hub_lbl, hub_color = "pass",    "complete", "#00e87a"
    elif run_status == "failed":  hub_st, hub_lbl, hub_color = "fail",    "stopped",  "#ff3d3d"
    elif run_status == "running": hub_st, hub_lbl, hub_color = "running", "running",  "#ffaa00"
    else:                         hub_st, hub_lbl, hub_color = "blocked", "no data",  "#4a3d7a"

    left_html  = "\n".join(render_node(a["name"], a["script"], a["status"], a["ts"], a["rows"]) for a in account_states[:11])
    right_html = "\n".join(render_node(a["name"], a["script"], a["status"], a["ts"], a["rows"]) for a in account_states[11:])

    warn_html = ""
    if run_status == "failed" and failed_at:
        warn_html = f"""<div class="warn-strip">&#9888; Pipeline stopped at <b>{failed_at}</b> &middot; Exit code 1 &middot; {blocked_ct} trigger{'s' if blocked_ct != 1 else ''} + dashboard.py + Git Repository not reached</div>"""

    logo = logo_b64()
    logo_img = (f'<img src="data:image/png;base64,{logo}" style="width:38px;height:38px;border-radius:6px;object-fit:contain;background:#fff;padding:2px;" alt="PacBiz">'
                if logo else
                '<div style="width:38px;height:38px;border-radius:6px;background:linear-gradient(135deg,#1a0050,#5000b4);display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:#c0a0ff;flex-shrink:0;">PB</div>')

    ok_color   = "#00e87a" if failed_ct == 0 else "#ff9090"
    run_icon   = "&#10003;" if run_status == "success" else ("&#10007;" if run_status == "failed" else "&mdash;")
    run_color  = "#00e87a" if run_status == "success" else ("#ff3d3d" if run_status == "failed" else "#7060a0")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<title>PacBiz Pipeline Monitor</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{background:#0a0018;min-height:100vh;font-family:system-ui,-apple-system,sans-serif;padding:20px;}}
.eco{{background:#05001a;border-radius:12px;padding:1.5rem 1rem;color:#fff;position:relative;overflow:hidden;max-width:980px;margin:0 auto;}}
.grid-bg{{position:absolute;inset:0;background-image:linear-gradient(rgba(80,0,180,0.08) 1px,transparent 1px),linear-gradient(90deg,rgba(80,0,180,0.08) 1px,transparent 1px);background-size:32px 32px;pointer-events:none;}}
.top-bar{{display:flex;align-items:center;justify-content:space-between;margin-bottom:0.25rem;position:relative;z-index:1;}}
.logo-box{{display:flex;align-items:center;gap:8px;}}
.brand-name{{font-size:15px;font-weight:500;color:#c0a0ff;letter-spacing:0.04em;}}
.brand-tag{{font-size:9px;color:#7060a0;margin-top:1px;}}
.legend{{display:flex;justify-content:center;gap:14px;margin-bottom:1rem;}}
.leg-item{{display:flex;align-items:center;gap:5px;font-size:10px;color:#9080c0;}}
.blink-g{{width:7px;height:7px;border-radius:50%;background:#00e87a;animation:blinkG 1.6s ease-in-out infinite;flex-shrink:0;}}
.blink-r{{width:7px;height:7px;border-radius:50%;background:#ff3d3d;animation:blinkR 0.9s ease-in-out infinite;flex-shrink:0;}}
.blink-n{{width:7px;height:7px;border-radius:50%;background:#4a3d7a;flex-shrink:0;}}
.blink-a{{width:7px;height:7px;border-radius:50%;background:#ffaa00;animation:blinkA 1.2s ease-in-out infinite;flex-shrink:0;}}
@keyframes blinkG{{0%,100%{{opacity:1;box-shadow:0 0 5px #00e87a;}}50%{{opacity:0.3;box-shadow:none;}}}}
@keyframes blinkR{{0%,100%{{opacity:1;box-shadow:0 0 7px #ff3d3d;}}50%{{opacity:0.2;box-shadow:none;}}}}
@keyframes blinkA{{0%,100%{{opacity:1;box-shadow:0 0 5px #ffaa00;}}50%{{opacity:0.3;box-shadow:none;}}}}
.main-layout{{display:grid;grid-template-columns:1fr 134px 1fr;gap:8px;align-items:start;position:relative;z-index:1;}}
.side-col{{display:flex;flex-direction:column;gap:5px;}}
.center-col{{display:flex;flex-direction:column;align-items:center;gap:4px;}}
.node{{border-radius:8px;border:1px solid;padding:7px 9px;}}
.node-title{{font-size:10px;font-weight:500;margin-bottom:4px;display:flex;align-items:center;justify-content:space-between;gap:4px;}}
.node-title span{{flex:1;}}
.node-rows{{display:flex;flex-direction:column;gap:2px;}}
.node-row{{display:flex;align-items:center;gap:4px;font-size:9px;color:#b0a0d0;}}
.node-row b{{width:5px;height:5px;border-radius:50%;flex-shrink:0;}}
.node-meta{{display:flex;gap:4px;margin-top:4px;flex-wrap:wrap;}}
.meta-pill{{font-size:8px;padding:1px 6px;border-radius:99px;background:rgba(80,0,180,0.2);color:#9080c0;border:1px solid rgba(80,0,180,0.2);white-space:nowrap;}}
.meta-pill.g{{background:rgba(0,232,122,0.08);color:#40c870;border-color:rgba(0,232,122,0.2);}}
.meta-pill.r{{background:rgba(255,61,61,0.1);color:#ff9090;border-color:rgba(255,61,61,0.2);}}
.b-g{{background:#00e87a;}}.b-r{{background:#ff3d3d;}}.b-n{{background:#4a3d7a;}}
.n-pass{{background:rgba(0,60,30,0.2);border-color:rgba(0,232,122,0.2);}}
.n-fail{{background:rgba(80,0,0,0.35);border-color:rgba(255,61,61,0.4);}}
.n-blocked,.n-pending{{background:rgba(30,15,60,0.4);border-color:rgba(74,61,122,0.33);}}
.n-warn{{background:rgba(80,50,0,0.3);border-color:rgba(255,170,0,0.3);}}
.center-hub{{width:94px;height:94px;border-radius:50%;border:2px solid #5000b4;background:radial-gradient(circle,#1a0050 60%,#0d0028);display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;position:relative;}}
.hub-ring{{position:absolute;inset:-7px;border-radius:50%;border:1px solid rgba(80,0,180,0.33);animation:spin 8s linear infinite;}}
.hub-ring2{{position:absolute;inset:-14px;border-radius:50%;border:1px solid rgba(80,0,180,0.13);animation:spin 14s linear infinite reverse;}}
@keyframes spin{{from{{transform:rotate(0deg);}}to{{transform:rotate(360deg);}}}}
.hub-label{{font-size:10px;font-weight:500;color:#c0a0ff;letter-spacing:0.02em;line-height:1.2;}}
.hub-status{{font-size:8px;margin-top:3px;display:flex;align-items:center;gap:2px;}}
.connector{{width:2px;height:14px;background:linear-gradient(rgba(80,0,180,0.33),rgba(80,0,180,0.13));margin:0 auto;}}
.stage-node{{border-radius:8px;border:1px solid rgba(80,0,180,0.4);background:rgba(40,10,90,0.4);padding:7px 8px;width:100%;text-align:center;}}
.stage-title{{font-size:9px;font-weight:500;color:#c0a0ff;margin-bottom:3px;}}
.stage-status{{display:flex;justify-content:center;margin:2px 0;}}
.stage-detail{{font-size:8px;color:#4a3d7a;}}
.agg-node{{border-radius:10px;border:1px solid rgba(80,0,180,0.47);background:rgba(40,10,90,0.5);padding:8px 10px;width:100%;margin:2px 0;}}
.agg-title{{font-size:9px;font-weight:500;color:#c0a0ff;text-align:center;margin-bottom:6px;}}
.agg-grid{{display:grid;grid-template-columns:1fr 1fr;gap:4px;}}
.agg-item{{text-align:center;background:rgba(80,0,180,0.15);border-radius:5px;padding:4px;}}
.agg-val{{font-size:11px;font-weight:500;}}
.agg-lbl{{font-size:8px;color:#7060a0;margin-top:1px;}}
.warn-strip{{background:rgba(80,0,0,0.3);border:1px solid rgba(255,61,61,0.27);border-radius:7px;padding:6px 9px;font-size:9px;color:#ff9090;display:flex;align-items:flex-start;gap:5px;margin-top:8px;position:relative;z-index:1;}}
.stat-bar{{display:flex;justify-content:center;gap:10px;padding:8px 0 0;flex-wrap:wrap;position:relative;z-index:1;}}
.stat{{text-align:center;}}
.stat-n{{font-size:16px;font-weight:500;}}
.stat-l{{font-size:9px;color:#7060a0;margin-top:1px;}}
.divline{{width:1px;height:30px;background:#2a1a4a;align-self:center;}}
.section-hdr{{font-size:8px;color:#5000b4;letter-spacing:0.06em;text-transform:uppercase;margin-bottom:3px;}}
.divider{{border:none;border-top:1px solid rgba(80,0,180,0.2);margin:8px 0;position:relative;z-index:1;}}
.meta-row{{display:flex;justify-content:space-between;align-items:center;position:relative;z-index:1;}}
@media(prefers-reduced-motion:reduce){{
  .blink-g,.blink-r,.blink-a,.hub-ring,.hub-ring2{{animation:none;}}
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
        <div class="brand-tag">Reports Pipeline Monitor</div>
      </div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:9px;color:#7060a0;">Live Pipeline Status &middot; Run {run_id}</div>
      <div style="font-size:8px;color:#4a3d7a;margin-top:1px;">Page built: {now_str}</div>
      <div style="font-size:8px;color:#4a3d7a;margin-top:1px;" id="cd">Auto-refresh in 60s</div>
    </div>
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
      <div class="section-hdr">Accounts 1&ndash;11</div>
      {left_html}
    </div>

    <div class="center-col">
      <div style="height:4px;"></div>
      <div class="center-hub">
        <div class="hub-ring"></div>
        <div class="hub-ring2"></div>
        <div class="hub-label">Pipeline<br>Trigger</div>
        <div class="hub-status">{dot(hub_st)}<span style="color:{hub_color};font-size:8px;">{hub_lbl}</span></div>
      </div>

      <div class="connector"></div>

      <div class="agg-node">
        <div class="agg-title">Data aggregate</div>
        <div class="agg-grid">
          <div class="agg-item"><div class="agg-val" style="color:#00e87a;">{total_rows:,}</div><div class="agg-lbl">total rows</div></div>
          <div class="agg-item"><div class="agg-val" style="color:{ok_color};">{passed}/{passed+failed_ct+blocked_ct}</div><div class="agg-lbl">sources ok</div></div>
          <div class="agg-item"><div class="agg-val" style="color:#c0a0ff;font-size:9px;">{started_at}</div><div class="agg-lbl">started</div></div>
          <div class="agg-item"><div class="agg-val" style="color:#c0a0ff;font-size:9px;">{finished_at}</div><div class="agg-lbl">finished</div></div>
        </div>
      </div>

      <div class="connector"></div>

      <div class="stage-node">
        <div class="stage-title">dashboard.py</div>
        <div class="stage-status">{dot(build_st)}</div>
        <div class="stage-detail">{"complete" if build_st == "pass" else ("failed" if build_st == "fail" else "not reached")}</div>
      </div>

      <div class="connector"></div>

      <div class="stage-node">
        <div class="stage-title">Git Repository</div>
        <div class="stage-status">{dot(git_st)}</div>
        <div class="stage-detail">{"pushed" if git_st == "pass" else ("failed" if git_st == "fail" else "not reached")}</div>
      </div>
    </div>

    <div class="side-col">
      <div class="section-hdr">Accounts 12&ndash;21</div>
      {right_html}
    </div>
  </div>

  {warn_html}

  <div class="stat-bar">
    <div class="stat"><div class="stat-n" style="color:#00e87a;">{passed}</div><div class="stat-l">passed</div></div>
    <div class="divline"></div>
    <div class="stat"><div class="stat-n" style="color:#ff3d3d;">{failed_ct}</div><div class="stat-l">failed</div></div>
    <div class="divline"></div>
    <div class="stat"><div class="stat-n" style="color:#4a3d7a;">{blocked_ct}</div><div class="stat-l">not reached</div></div>
    <div class="divline"></div>
    <div class="stat"><div class="stat-n" style="color:#c0a0ff;">{total_rows:,}</div><div class="stat-l">rows pulled</div></div>
    <div class="divline"></div>
    <div class="stat"><div class="stat-n" style="color:{run_color};">{run_icon}</div><div class="stat-l">{run_status}</div></div>
  </div>
</div>

<script>
(function() {{
  var s = 60, el = document.getElementById('cd');
  function tick() {{
    if (s <= 0) {{ location.reload(); return; }}
    if (el) el.textContent = 'Auto-refresh in ' + s + 's';
    s--;
    setTimeout(tick, 1000);
  }}
  tick();
}})();
</script>
</body>
</html>"""

    OUTPUT_FILE.write_text(html, encoding="utf-8")
    print(f"Pipeline monitor generated: {OUTPUT_FILE.name}")

generate()

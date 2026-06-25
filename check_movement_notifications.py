"""check_movement_notifications.py — Send per-movement HTML emails for newly processed rows.

Reads movement_cache.csv (written by masterlist_fetch.py), compares Timestamps
against movement_notified.json to find unseen rows, sends one email per movement,
then saves the updated seen list.

Filter criteria: Processed == "Yes" AND Void != "Yes" AND Timestamp not in seen list.
"""

import base64, json, sys
import pandas as pd
from datetime import datetime
from pathlib import Path

BASE_DIR         = Path(__file__).parent
MOVEMENT_CACHE   = BASE_DIR / "movement_cache.csv"
MASTERLIST_CACHE = BASE_DIR / "masterlist_cache.csv"
NOTIFIED_FILE    = BASE_DIR / "movement_notified.json"
LOGO_PATH        = BASE_DIR / "pacbiz_logo.png"

CC_ALWAYS = ["it-team@pac-biz.com", "hr@pac-biz.com"]


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def load_notified():
    if NOTIFIED_FILE.exists():
        try:
            return set(json.loads(NOTIFIED_FILE.read_text(encoding="utf-8")))
        except Exception:
            return set()
    return set()


def save_notified(seen):
    NOTIFIED_FILE.write_text(json.dumps(sorted(seen), indent=2), encoding="utf-8")


def na_or(val, fallback="No changes"):
    """Return fallback when val is blank, N/A, NA, None, NaN, or a dash."""
    if pd.isna(val):
        return fallback
    s = str(val).strip()
    if not s or s.upper() in ("N/A", "NA", "NONE", "-", "N/A.", "N.A."):
        return fallback
    return s


def format_date(val):
    if pd.isna(val) if hasattr(pd, "isna") else val != val:
        return ""
    s = str(val).strip()
    if not s or s.upper() in ("N/A", "NA", ""):
        return ""
    try:
        return pd.to_datetime(s).strftime("%d %b %Y")
    except Exception:
        return s


def make_ref(timestamp_str, seq):
    try:
        dt = pd.to_datetime(timestamp_str)
        return f"MOV-{dt.strftime('%Y-%m%d')}-{seq:03d}"
    except Exception:
        return f"MOV-{seq:03d}"


def logo_b64():
    if LOGO_PATH.exists():
        data = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        return f"data:image/jpeg;base64,{data}"
    return ""


# ---------------------------------------------------------------------------
# masterlist lookups
# ---------------------------------------------------------------------------

def build_lookups(ml_df):
    """Return (email_to_name dict, name_to_email dict) — both lower-keyed."""
    e2n, n2e = {}, {}
    for _, row in ml_df.iterrows():
        email = str(row.get("Company Email", "")).strip()
        name  = str(row.get("Emp Name", "")).strip()
        if email and name:
            e2n[email.lower()] = name
            n2e[name.lower()]  = email
    return e2n, n2e


def get_supervisor_email(row, ml_df, n2e):
    new_sup = na_or(row.get("New Supervisor", ""), "")
    if new_sup and new_sup != "No changes":
        email = n2e.get(new_sup.lower(), "")
        if not email:
            for k, v in n2e.items():
                if new_sup.lower() in k:
                    email = v
                    break
        return email

    # Fall back to employee's current supervisor in masterlist
    emp = str(row.get("Employee Name", "")).strip().lower()
    match = ml_df[ml_df["Emp Name"].str.strip().str.lower() == emp]
    if match.empty:
        match = ml_df[ml_df["Emp Name"].str.strip().str.lower().str.contains(emp, na=False, regex=False)]
    if not match.empty:
        current_sup = str(match.iloc[0].get("Immediate Supervisor", "")).strip()
        if current_sup:
            return n2e.get(current_sup.lower(), "")
    return ""


def get_employee_current(emp_name, ml_df):
    """Look up an employee's current info from the masterlist by name.
    Returns a dict with keys: dept, account, supervisor, job_title, emp_status.
    Falls back to empty string for any missing field.
    """
    emp_lower = emp_name.strip().lower()
    match = ml_df[ml_df["Emp Name"].str.strip().str.lower() == emp_lower]
    if match.empty:
        match = ml_df[ml_df["Emp Name"].str.strip().str.lower().str.contains(emp_lower, na=False, regex=False)]
    if match.empty:
        return {"dept": "", "account": "", "supervisor": "", "job_title": "", "emp_status": ""}
    r = match.iloc[0]
    return {
        "dept":       str(r.get("Department", "")).strip(),
        "account":    str(r.get("LOB / Account", "")).strip(),
        "supervisor": str(r.get("Immediate Supervisor", "")).strip(),
        "job_title":  str(r.get("Job Title", "")).strip(),
        "emp_status": str(r.get("Employment Status", "")).strip(),
    }


# ---------------------------------------------------------------------------
# HTML builder
# ---------------------------------------------------------------------------

LOGO_URI = logo_b64()


def _no_chg(text):
    """Render 'No changes' in muted italic, otherwise return text unchanged."""
    if text == "No changes":
        return '<span style="color:#94a3b8;font-style:italic;">No changes</span>'
    return text


def build_html(row, ref_num, submitted_by, processed_on, ml_df=None):
    is_attrition  = str(row.get("Type of Movement", "")).strip().lower() == "attrition"
    mov_type      = str(row.get("Type of Movement", "")).strip()
    emp_name      = str(row.get("Employee Name", "")).strip()
    effective     = format_date(row.get("Effective Date", ""))
    mov_subtype   = na_or(row.get("Movement Type", ""), "")

    if is_attrition and ml_df is not None:
        # For Attrition: pull current values from the masterlist (form fields are blank)
        cur       = get_employee_current(emp_name, ml_df)
        dept      = cur["dept"]      or "—"
        account   = cur["account"]   or "—"
        supervisor = cur["supervisor"] or "—"
        job_title  = cur["job_title"]  or "—"
        emp_status = cur["emp_status"] or "—"
    else:
        dept      = na_or(row.get("New Department", ""))
        account   = na_or(row.get("New Account", ""))
        supervisor = na_or(row.get("New Supervisor", ""))
        job_title  = na_or(row.get("New Job Title", ""))
        emp_status = na_or(row.get("New Employment Status", ""))
    remarks       = na_or(row.get("Remarks/Comments", ""), "")
    proc_note     = na_or(row.get("Processed Note", ""), "")

    # 4th summary cell
    if is_attrition:
        cell4_label = "Attrition Date"
        cell4_val   = effective or "—"
        cell4_color = "#dc2626"
        cell4_bg    = "#fee2e2"
    else:
        cell4_label = "Movement Type"
        cell4_val   = mov_subtype if mov_subtype and mov_subtype != "No changes" else "—"
        cell4_color = "#004C97"
        cell4_bg    = "#e8edf7"

    # Employment status pill
    sl = emp_status.lower()
    if sl == "inactive":
        pill = "background:#fee2e2;color:#991b1b;"
    elif sl in ("regular", "active", "probationary"):
        pill = "background:#dcfce7;color:#15803d;"
    elif emp_status == "No changes":
        pill = "background:#f1f5f9;color:#64748b;"
    else:
        pill = "background:#dbeafe;color:#1e40af;"

    # Changes Applied from Processed Note
    if proc_note and proc_note != "No changes":
        dot_color = "#dc2626" if is_attrition else "#004C97"
        items_html = "".join(
            f'<div style="display:flex;align-items:flex-start;gap:6px;margin-bottom:5px;font-size:11.5px;color:#374151;">'
            f'<div style="width:7px;height:7px;border-radius:50%;background:{dot_color};flex-shrink:0;margin-top:3px;"></div>'
            f'{item.strip()}</div>'
            for item in proc_note.split(",") if item.strip()
        )
    else:
        items_html = '<div style="font-size:11px;color:#94a3b8;font-style:italic;">No changes recorded.</div>'

    # Remarks (Internal Movement only)
    remarks_block = ""
    if not is_attrition and remarks and remarks != "No changes":
        remarks_block = (
            f'<div style="margin:0 24px 12px;padding:8px 12px;background:#f8fafc;'
            f'border-left:3px solid #cbd5e1;border-radius:0 4px 4px 0;'
            f'font-size:11px;color:#64748b;font-style:italic;line-height:1.5;">'
            f'{remarks}</div>'
        )

    # IT Note (Attrition only)
    it_note = ""
    if is_attrition:
        it_note = (
            '<div style="margin:0 24px 16px;padding:10px 16px;background:#fff5f5;'
            'border:1.5px solid #fca5a5;border-left:4px solid #dc2626;border-radius:4px;'
            'display:flex;align-items:flex-start;gap:10px;">'
            '<div style="font-size:12px;color:#991b1b;font-weight:700;">Note to IT: '
            '<span style="font-weight:400;">Kindly disable all credentials of the above employee.</span></div></div>'
        )

    logo_img = (
        f'<img src="{LOGO_URI}" alt="Pac-Biz" style="width:48px;height:48px;'
        f'object-fit:cover;border-radius:50%;flex-shrink:0;">'
        if LOGO_URI else
        '<div style="width:48px;height:48px;border-radius:50%;background:#004C97;'
        'display:flex;align-items:center;justify-content:center;flex-shrink:0;">'
        '<span style="color:#39B54A;font-weight:900;font-size:14px;">PB</span></div>'
    )

    attrition_icon_color = "#dc2626" if is_attrition else "#1e40af"
    attrition_icon_bg    = "#fee2e2" if is_attrition else "#dbeafe"
    changes_head_color   = "#dc2626" if is_attrition else "#004C97"

    # Employment status cell
    if emp_status == "No changes":
        status_cell = '<span style="color:#94a3b8;font-style:italic;">No changes</span>'
    else:
        status_cell = (
            f'<span style="display:inline-block;padding:3px 11px;border-radius:4px;'
            f'font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;{pill}">'
            f'{emp_status}</span>'
        )

    account_cell = (
        '<span style="color:#94a3b8;font-style:italic;">No changes</span>'
        if account == "No changes" else f"<strong>{account}</strong>"
    )
    supervisor_cell = (
        '<span style="color:#94a3b8;font-style:italic;">No changes</span>'
        if supervisor == "No changes" else f"<strong>{supervisor}</strong>"
    )

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#e8edf2;font-family:Arial,sans-serif;">
<div style="max-width:650px;margin:0 auto;background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 4px 28px rgba(0,0,0,0.13);">

  <!-- HEADER -->
  <div style="background:#0f2044;padding:18px 28px;display:flex;align-items:center;gap:0;">
    <div style="display:flex;align-items:center;gap:11px;flex:1;">
      {logo_img}
      <div>
        <div style="font-size:22px;font-weight:900;letter-spacing:.02em;line-height:1;color:#fff;">PAC <span style="color:#39B54A;">BIZ</span></div>
        <div style="font-size:9px;letter-spacing:.18em;color:rgba(255,255,255,0.55);text-transform:uppercase;margin-top:1px;">Outsourcing</div>
      </div>
    </div>
    <div style="width:1px;height:48px;background:rgba(255,255,255,0.15);margin:0 22px;flex-shrink:0;"></div>
    <div style="display:flex;align-items:center;gap:12px;">
      <div style="width:40px;height:40px;border-radius:50%;border:2px solid #39B54A;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#39B54A" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="7" r="4"/><path d="M4 21v-1a8 8 0 0 1 16 0v1"/></svg>
      </div>
      <div>
        <div style="font-size:10px;font-weight:700;letter-spacing:.12em;color:#39B54A;text-transform:uppercase;">System Notification</div>
        <div style="font-size:13px;color:#fff;font-weight:600;margin-top:2px;">Employee Movement ({mov_type})</div>
      </div>
    </div>
  </div>

  <!-- SUCCESS STRIP -->
  <div style="background:#f0fdf4;border-bottom:1px solid #bbf7d0;padding:13px 24px;display:flex;align-items:flex-start;gap:12px;">
    <div style="width:28px;height:28px;border-radius:50%;background:#39B54A;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px;">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
    </div>
    <div>
      <strong style="color:#15803d;font-size:13px;">An employee movement ({mov_type}) has been submitted and processed successfully.</strong>
      <p style="color:#374151;font-size:12px;margin-top:2px;margin-bottom:0;">The details of the movement are shown below.</p>
    </div>
  </div>

  <!-- SUMMARY ROW -->
  <div style="display:grid;grid-template-columns:repeat(4,1fr);border-bottom:1px solid #e5e7eb;">
    <div style="padding:16px 10px;text-align:center;border-right:1px solid #e5e7eb;">
      <div style="width:38px;height:38px;border-radius:50%;background:#e8edf7;margin:0 auto 8px;display:flex;align-items:center;justify-content:center;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#0f2044" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="7" r="4"/><path d="M4 21v-1a8 8 0 0 1 16 0v1"/></svg>
      </div>
      <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">Employee Name</div>
      <div style="font-size:12px;font-weight:700;color:#0f2044;line-height:1.3;">{emp_name}</div>
    </div>
    <div style="padding:16px 10px;text-align:center;border-right:1px solid #e5e7eb;">
      <div style="width:38px;height:38px;border-radius:50%;background:{attrition_icon_bg};margin:0 auto 8px;display:flex;align-items:center;justify-content:center;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="{attrition_icon_color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      </div>
      <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">Type of Movement</div>
      <div style="font-size:14px;font-weight:700;color:{attrition_icon_color};">{mov_type}</div>
    </div>
    <div style="padding:16px 10px;text-align:center;border-right:1px solid #e5e7eb;">
      <div style="width:38px;height:38px;border-radius:50%;background:#dcfce7;margin:0 auto 8px;display:flex;align-items:center;justify-content:center;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#15803d" stroke-width="2" stroke-linecap="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
      </div>
      <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">Effective Date</div>
      <div style="font-size:14px;font-weight:700;color:#15803d;">{effective or "—"}</div>
    </div>
    <div style="padding:16px 10px;text-align:center;">
      <div style="width:38px;height:38px;border-radius:50%;background:{cell4_bg};margin:0 auto 8px;display:flex;align-items:center;justify-content:center;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="{cell4_color}" stroke-width="2" stroke-linecap="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
      </div>
      <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">{cell4_label}</div>
      <div style="font-size:13px;font-weight:700;color:{cell4_color};line-height:1.3;">{cell4_val}</div>
    </div>
  </div>

  <!-- MOVEMENT DETAILS -->
  <div style="display:flex;align-items:center;gap:8px;padding:16px 24px 10px;">
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#39B54A" stroke-width="2" stroke-linecap="round"><path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="1"/><line x1="9" y1="12" x2="15" y2="12"/><line x1="9" y1="16" x2="13" y2="16"/></svg>
    <span style="font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#0f2044;">Movement Details</span>
  </div>

  <div style="padding:0 24px 0;">
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead>
        <tr style="background:#0f2044;">
          <th style="padding:9px 16px;text-align:left;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#fff;width:45%;">Updated Information</th>
          <th style="padding:9px 16px;text-align:left;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#39B54A;">Detail</th>
        </tr>
      </thead>
      <tbody>
        <tr style="border-bottom:1px solid #f1f5f9;">
          <td style="padding:9px 16px;color:#0f2044;font-weight:500;">Department</td>
          <td style="padding:9px 16px;color:#374151;">{_no_chg(dept)}</td>
        </tr>
        <tr style="border-bottom:1px solid #f1f5f9;background:#f8fafc;">
          <td style="padding:9px 16px;color:#0f2044;font-weight:500;">LOB / Account</td>
          <td style="padding:9px 16px;color:#374151;">{account_cell}</td>
        </tr>
        <tr style="border-bottom:1px solid #f1f5f9;">
          <td style="padding:9px 16px;color:#0f2044;font-weight:500;">Immediate Supervisor</td>
          <td style="padding:9px 16px;color:#374151;">{supervisor_cell}</td>
        </tr>
        <tr style="border-bottom:1px solid #f1f5f9;background:#f8fafc;">
          <td style="padding:9px 16px;color:#0f2044;font-weight:500;">Job Title</td>
          <td style="padding:9px 16px;color:#374151;">{_no_chg(job_title)}</td>
        </tr>
        <tr>
          <td style="padding:9px 16px;color:#0f2044;font-weight:500;">Employment Status</td>
          <td style="padding:9px 16px;">{status_cell}</td>
        </tr>
      </tbody>
    </table>
  </div>

  {remarks_block}

  <div style="height:16px;"></div>

  <!-- BOTTOM TWO-COL -->
  <div style="display:grid;grid-template-columns:1fr 1fr;margin:0 24px;border-radius:6px;border:1px solid #e5e7eb;overflow:hidden;">
    <div style="padding:16px 20px;border-right:1px solid #e5e7eb;">
      <div style="display:flex;align-items:center;gap:7px;margin-bottom:10px;">
        <span style="font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:{changes_head_color};">Changes Applied</span>
      </div>
      {items_html}
    </div>
    <div style="padding:16px 20px;">
      <div style="display:flex;align-items:center;gap:7px;margin-bottom:10px;">
        <span style="font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#004C97;">System Information</span>
      </div>
      <div style="display:flex;flex-direction:column;gap:6px;">
        <div style="display:flex;gap:8px;font-size:11.5px;">
          <span style="font-weight:700;color:#0f2044;min-width:90px;flex-shrink:0;">Processed By</span>
          <span style="color:#374151;">Pac-Biz Reports Automation</span>
        </div>
        <div style="display:flex;gap:8px;font-size:11.5px;">
          <span style="font-weight:700;color:#0f2044;min-width:90px;flex-shrink:0;">Submitted By</span>
          <span style="color:#374151;">{submitted_by}</span>
        </div>
        <div style="display:flex;gap:8px;font-size:11.5px;">
          <span style="font-weight:700;color:#0f2044;min-width:90px;flex-shrink:0;">Processed On</span>
          <span style="color:#374151;">{processed_on}</span>
        </div>
        <div style="display:flex;gap:8px;font-size:11.5px;">
          <span style="font-weight:700;color:#0f2044;min-width:90px;flex-shrink:0;">Reference No.</span>
          <span style="color:#0f2044;font-weight:700;">{ref_num}</span>
        </div>
      </div>
    </div>
  </div>

  <div style="height:16px;"></div>

  {it_note}

  <!-- FOOTER -->
  <div style="background:#0f2044;padding:8px 28px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;">
    <div style="display:flex;align-items:center;gap:10px;">
      {logo_img}
      <div>
        <div style="font-size:10px;font-weight:800;color:#fff;">PAC <span style="color:#39B54A;">BIZ</span> OUTSOURCING</div>
        <div style="font-size:10px;color:rgba(255,255,255,0.5);margin-top:4px;font-style:italic;">One Team. One Purpose. Endless Possibilities.</div>
      </div>
    </div>
    <a href="https://mikewoocerna.github.io/Pac-Biz/masterlist_dashboard.html"
       style="display:inline-block;background:#004C97;color:#fff;padding:9px 20px;border-radius:5px;text-decoration:none;font-size:12px;font-weight:700;letter-spacing:.02em;white-space:nowrap;flex-shrink:0;">
      View Dashboard
    </a>
  </div>
  <div style="height:3px;background:linear-gradient(90deg,#39B54A,#2d9e40);"></div>

</div>
</body>
</html>"""


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    if not MOVEMENT_CACHE.exists():
        print("[movement_notify] movement_cache.csv not found — skipping.")
        return 0

    if not MASTERLIST_CACHE.exists():
        print("[movement_notify] masterlist_cache.csv not found — skipping.")
        return 0

    try:
        mv_df = pd.read_csv(MOVEMENT_CACHE, dtype=str)
        ml_df = pd.read_csv(MASTERLIST_CACHE, dtype=str)
    except Exception as e:
        print(f"[movement_notify] Failed to read CSV: {e}")
        return 1

    seen = load_notified()
    e2n, n2e = build_lookups(ml_df)

    new_rows = [
        row for _, row in mv_df.iterrows()
        if str(row.get("Processed", "")).strip().lower() == "yes"
        and str(row.get("Void", "")).strip().lower() != "yes"
        and str(row.get("Timestamp", "")).strip() not in seen
    ]

    if not new_rows:
        print("[movement_notify] No new processed movements to notify.")
        return 0

    print(f"[movement_notify] Found {len(new_rows)} new movement(s) to notify.")

    from notify import notify_movement

    seq_start   = len(seen) + 1
    success_cnt = 0

    for i, row in enumerate(new_rows):
        emp_name        = str(row.get("Employee Name", "")).strip()
        mov_type        = str(row.get("Type of Movement", "")).strip()
        ts_str          = str(row.get("Timestamp", "")).strip()
        submitter_email = str(row.get("Email Address", "")).strip().lower()
        proc_date_raw   = na_or(row.get("Processed Date", ""), "")

        submitted_by  = e2n.get(submitter_email, submitter_email) if submitter_email else "—"
        processed_on  = format_date(proc_date_raw) if proc_date_raw and proc_date_raw != "No changes" else datetime.now().strftime("%d %b %Y")
        ref_num       = make_ref(ts_str, seq_start + i)

        to_email  = submitter_email or ""
        sup_email = get_supervisor_email(row, ml_df, n2e)
        cc_list   = [e for e in ([sup_email] + CC_ALWAYS) if e]

        subject   = f"Employee Movement {mov_type} - {emp_name}"
        html_body = build_html(row, ref_num, submitted_by, processed_on, ml_df=ml_df)

        ok = notify_movement(
            subject   = subject,
            body_html = html_body,
            to_email  = to_email,
            cc_list   = cc_list,
        )

        if ok:
            seen.add(ts_str)
            success_cnt += 1
            print(f"[movement_notify] Sent: {subject}")
        else:
            print(f"[movement_notify] Failed to send for: {emp_name}")

    save_notified(seen)
    print(f"[movement_notify] Done — {success_cnt}/{len(new_rows)} notifications sent.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

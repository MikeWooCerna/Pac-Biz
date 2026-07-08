"""check_movement_notifications.py — Send per-movement HTML emails for newly processed rows.

Reads movement_cache.csv (written by masterlist_fetch.py), compares Timestamps
against movement_notified.json to find unseen rows, sends one email per movement,
then saves the updated seen list.

Filter criteria: fully processed, not voided, and Timestamp not in seen list.
"""

import json, sys
import pandas as pd
from datetime import datetime
from pathlib import Path

BASE_DIR         = Path(__file__).parent
MOVEMENT_CACHE   = BASE_DIR / "movement_cache.csv"
MASTERLIST_CACHE = BASE_DIR / "masterlist_cache.csv"
NOTIFIED_FILE    = BASE_DIR / "movement_notified.json"
LOGO_PATH        = BASE_DIR / "pacbiz_logo.png"

CC_ALWAYS  = ["it-team@pac-biz.com", "hr@pac-biz.com"]
BCC_ALWAYS = ["reports@pac-biz.com"]


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def load_notified():
    if NOTIFIED_FILE.exists():
        try:
            return set(json.loads(NOTIFIED_FILE.read_text(encoding="utf-8")))
        except Exception:
            return set()
    return None


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


def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def has_value(val):
    text = clean_text(val)
    return bool(text) and text.upper() not in ("N/A", "NA", "NONE", "-", "N/A.", "N.A.")


def is_ready_to_notify(row):
    """Only notify after the movement has complete processing details."""
    required_fields = [
        "Timestamp",
        "Employee Name",
        "Type of Movement",
        "Email Address",
        "Processed Date",
        "Processed Note",
    ]
    return (
        clean_text(row.get("Processed")).lower() == "yes"
        and clean_text(row.get("Void")).lower() != "yes"
        and all(has_value(row.get(field)) for field in required_fields)
    )


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
    # Gmail strips data: URIs — use the hosted URL so the image loads in email clients
    return "https://mikewoocerna.github.io/Pac-Biz/pacbiz_logo.png"


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


def normalize_employment_status_for_email(value, is_attrition=False):
    """Movement forms may store class-like values; emails show only Active/Inactive."""
    if is_attrition:
        return "Inactive"

    raw = "" if pd.isna(value) else str(value).strip()
    key = raw.lower()
    if not key or key in ("n/a", "na", "none", "-", "n/a.", "n.a."):
        return "No changes"

    inactive_tokens = (
        "inactive",
        "terminated",
        "termination",
        "resigned",
        "attrition",
        "separated",
        "end of contract",
    )
    if any(token in key for token in inactive_tokens):
        return "Inactive"

    active_tokens = (
        "active",
        "regular",
        "probation",
        "probationary",
        "option",
        "full time",
        "full-time",
        "part time",
        "part-time",
    )
    if any(token in key for token in active_tokens):
        return "Active"

    return raw


def build_html(row, ref_num, submitted_by, processed_on, ml_df=None):
    import re as _re
    is_attrition  = str(row.get("Type of Movement", "")).strip().lower() == "attrition"
    mov_type      = str(row.get("Type of Movement", "")).strip()
    emp_name      = str(row.get("Employee Name", "")).strip()
    effective     = format_date(row.get("Effective Date", ""))
    mov_subtype   = na_or(row.get("Movement Type", ""), "")

    if is_attrition and ml_df is not None:
        cur        = get_employee_current(emp_name, ml_df)
        dept       = cur["dept"]       or "—"
        account    = cur["account"]    or "—"
        supervisor = cur["supervisor"] or "—"
        job_title  = cur["job_title"]  or "—"
        emp_status = normalize_employment_status_for_email(cur["emp_status"], is_attrition=True)
    else:
        dept       = na_or(row.get("New Department", ""))
        account    = na_or(row.get("New Account", ""))
        supervisor = na_or(row.get("New Supervisor", ""))
        job_title  = na_or(row.get("New Job Title", ""))
        emp_status = normalize_employment_status_for_email(row.get("New Employment Status", ""))
    remarks   = na_or(row.get("Remarks/Comments", ""), "")
    proc_note = na_or(row.get("Processed Note", ""), "")

    # 4th summary cell
    if is_attrition:
        cell4_label = "Attrition Date"
        cell4_val   = effective or "—"
        cell4_color = "#dc2626"
    else:
        cell4_label = "Movement Type"
        cell4_val   = mov_subtype if mov_subtype and mov_subtype != "No changes" else "—"
        cell4_color = "#004C97"

    type_color = "#dc2626" if is_attrition else "#1e40af"
    changes_color = "#dc2626" if is_attrition else "#004C97"

    # Employment status pill
    sl = emp_status.lower()
    if sl == "inactive":
        pill = "background:#fee2e2;color:#991b1b;"
    elif sl in ("regular", "active", "probationary"):
        pill = "background:#dcfce7;color:#15803d;"
    elif emp_status in ("No changes", "—"):
        pill = "background:#f1f5f9;color:#64748b;"
    else:
        pill = "background:#dbeafe;color:#1e40af;"

    status_cell = (
        f'<span style="display:inline-block;padding:3px 10px;border-radius:3px;'
        f'font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;{pill}">'
        f'{emp_status}</span>'
    )
    account_cell    = f"<strong>{account}</strong>" if account not in ("No changes", "—") else account
    supervisor_cell = f"<strong>{supervisor}</strong>" if supervisor not in ("No changes", "—") else supervisor

    # Changes Applied — split on | or , (proc_note may use either)
    dot = f'<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:{changes_color};margin-right:7px;vertical-align:middle;"></span>'
    if proc_note and proc_note != "No changes":
        change_items = [i.strip() for i in _re.split(r'[|,]', proc_note) if i.strip()]
        items_html = "".join(
            f'<p style="margin:0 0 5px 0;font-size:11.5px;color:#374151;">{dot}{item}</p>'
            for item in change_items
        )
    else:
        items_html = '<p style="margin:0;font-size:11px;color:#94a3b8;font-style:italic;">No changes recorded.</p>'

    # Remarks row (Internal only) — as a table row
    remarks_row = ""
    if not is_attrition and remarks and remarks != "No changes":
        remarks_row = (
            f'<tr><td style="padding:0 24px 12px;">'
            f'<div style="padding:8px 12px;background:#f8fafc;border-left:3px solid #cbd5e1;'
            f'font-size:11px;color:#64748b;font-style:italic;line-height:1.5;">{remarks}</div>'
            f'</td></tr>'
        )

    # IT Note (Attrition only) — as a table row
    it_note_row = ""
    if is_attrition:
        it_note_row = (
            '<tr><td style="padding:0 24px 16px;">'
            '<div style="padding:10px 16px;background:#fff5f5;border:1.5px solid #fca5a5;'
            'border-left:4px solid #dc2626;border-radius:4px;">'
            '<span style="font-size:12px;color:#991b1b;font-weight:700;">Note to IT:&nbsp;</span>'
            '<span style="font-size:12px;color:#991b1b;">Kindly disable all credentials of the above employee.</span>'
            '</div></td></tr>'
        )

    logo_img = (
        f'<img src="{LOGO_URI}" alt="PB" width="44" height="44" '
        f'style="width:44px;height:44px;border-radius:22px;display:block;">'
        if LOGO_URI else
        '<div style="width:44px;height:44px;border-radius:22px;background:#004C97;'
        'text-align:center;line-height:44px;color:#39B54A;font-weight:900;font-size:13px;">PB</div>'
    )

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:20px 10px;background:#e8edf2;font-family:Arial,Helvetica,sans-serif;">

<table width="620" cellpadding="0" cellspacing="0" border="0" align="center" style="background:#ffffff;border-radius:4px;">

  <!-- HEADER -->
  <tr>
    <td style="background:#0f2044;padding:16px 24px;border-radius:4px 4px 0 0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="52" style="vertical-align:middle;">{logo_img}</td>
          <td style="padding-left:10px;vertical-align:middle;">
            <div style="font-size:20px;font-weight:900;color:#ffffff;line-height:1;">PAC <span style="color:#39B54A;">BIZ</span></div>
            <div style="font-size:9px;letter-spacing:.15em;color:rgba(255,255,255,0.5);text-transform:uppercase;margin-top:2px;">Outsourcing</div>
            <div style="font-size:10px;color:rgba(255,255,255,0.5);margin-top:3px;font-style:italic;">One Team. One Purpose. Endless Possibilities.</div>
          </td>
          <td width="1" style="padding:0 18px;vertical-align:middle;">
            <div style="width:1px;height:40px;background:rgba(255,255,255,0.2);"></div>
          </td>
          <td style="vertical-align:middle;text-align:right;">
            <div style="font-size:10px;font-weight:700;letter-spacing:.1em;color:#39B54A;text-transform:uppercase;">System Notification</div>
            <div style="font-size:13px;color:#ffffff;font-weight:600;margin-top:3px;">Employee Movement ({mov_type})</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- SUCCESS STRIP -->
  <tr>
    <td style="background:#f0fdf4;border-bottom:1px solid #bbf7d0;padding:11px 24px;">
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="28" style="vertical-align:top;">
            <div style="width:24px;height:24px;border-radius:12px;background:#39B54A;text-align:center;line-height:24px;color:#ffffff;font-size:14px;font-weight:700;">&#10003;</div>
          </td>
          <td style="padding-left:10px;vertical-align:middle;">
            <div style="font-size:13px;font-weight:700;color:#15803d;">An employee movement ({mov_type}) has been submitted and processed successfully.</div>
            <div style="font-size:12px;color:#374151;margin-top:2px;">The details of the movement are shown below.</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- SUMMARY 4-CELL ROW -->
  <tr>
    <td style="border-bottom:1px solid #e5e7eb;padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="25%" style="padding:16px 8px;text-align:center;border-right:1px solid #e5e7eb;vertical-align:top;">
            <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">Employee Name</div>
            <div style="font-size:12px;font-weight:700;color:#0f2044;line-height:1.3;">{emp_name}</div>
          </td>
          <td width="25%" style="padding:16px 8px;text-align:center;border-right:1px solid #e5e7eb;vertical-align:top;">
            <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">Type of Movement</div>
            <div style="font-size:14px;font-weight:700;color:{type_color};">{mov_type}</div>
          </td>
          <td width="25%" style="padding:16px 8px;text-align:center;border-right:1px solid #e5e7eb;vertical-align:top;">
            <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">Effective Date</div>
            <div style="font-size:14px;font-weight:700;color:#15803d;">{effective or "—"}</div>
          </td>
          <td width="25%" style="padding:16px 8px;text-align:center;vertical-align:top;">
            <div style="font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin-bottom:5px;">{cell4_label}</div>
            <div style="font-size:13px;font-weight:700;color:{cell4_color};line-height:1.3;">{cell4_val}</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- MOVEMENT DETAILS heading -->
  <tr>
    <td style="padding:16px 24px 8px;">
      <div style="font-size:11px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#0f2044;">Movement Details</div>
    </td>
  </tr>

  <!-- MOVEMENT DETAILS table -->
  <tr>
    <td style="padding:0 24px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
        <tr>
          <th style="padding:8px 14px;background:#0f2044;text-align:left;font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#ffffff;width:45%;">Updated Information</th>
          <th style="padding:8px 14px;background:#0f2044;text-align:left;font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#39B54A;">Detail</th>
        </tr>
        <tr>
          <td style="padding:9px 14px;font-size:13px;color:#0f2044;font-weight:600;border-bottom:1px solid #f1f5f9;">Department</td>
          <td style="padding:9px 14px;font-size:13px;color:#374151;border-bottom:1px solid #f1f5f9;">{dept}</td>
        </tr>
        <tr style="background:#f8fafc;">
          <td style="padding:9px 14px;font-size:13px;color:#0f2044;font-weight:600;border-bottom:1px solid #f1f5f9;">LOB / Account</td>
          <td style="padding:9px 14px;font-size:13px;color:#374151;border-bottom:1px solid #f1f5f9;">{account_cell}</td>
        </tr>
        <tr>
          <td style="padding:9px 14px;font-size:13px;color:#0f2044;font-weight:600;border-bottom:1px solid #f1f5f9;">Immediate Supervisor</td>
          <td style="padding:9px 14px;font-size:13px;color:#374151;border-bottom:1px solid #f1f5f9;">{supervisor_cell}</td>
        </tr>
        <tr style="background:#f8fafc;">
          <td style="padding:9px 14px;font-size:13px;color:#0f2044;font-weight:600;border-bottom:1px solid #f1f5f9;">Job Title</td>
          <td style="padding:9px 14px;font-size:13px;color:#374151;border-bottom:1px solid #f1f5f9;">{job_title}</td>
        </tr>
        <tr>
          <td style="padding:9px 14px;font-size:13px;color:#0f2044;font-weight:600;">Employment Status</td>
          <td style="padding:9px 14px;">{status_cell}</td>
        </tr>
      </table>
    </td>
  </tr>

  {remarks_row}

  <!-- spacer -->
  <tr><td style="height:16px;font-size:0;">&nbsp;</td></tr>

  <!-- BOTTOM TWO-COL: Changes Applied + System Information -->
  <tr>
    <td style="padding:0 24px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #e5e7eb;border-collapse:collapse;">
        <tr>
          <td width="50%" style="padding:16px 18px;vertical-align:top;border-right:1px solid #e5e7eb;">
            <div style="font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:{changes_color};margin-bottom:10px;">Changes Applied</div>
            {items_html}
          </td>
          <td width="50%" style="padding:16px 18px;vertical-align:top;">
            <div style="font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#004C97;margin-bottom:10px;">System Information</div>
            <table cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="font-size:11.5px;font-weight:700;color:#0f2044;padding:3px 10px 3px 0;white-space:nowrap;vertical-align:top;">Processed By</td>
                <td style="font-size:11.5px;color:#374151;padding:3px 0;vertical-align:top;">Pac-Biz Reports Automation</td>
              </tr>
              <tr>
                <td style="font-size:11.5px;font-weight:700;color:#0f2044;padding:3px 10px 3px 0;white-space:nowrap;vertical-align:top;">Submitted By</td>
                <td style="font-size:11.5px;color:#374151;padding:3px 0;vertical-align:top;">{submitted_by}</td>
              </tr>
              <tr>
                <td style="font-size:11.5px;font-weight:700;color:#0f2044;padding:3px 10px 3px 0;white-space:nowrap;vertical-align:top;">Processed On</td>
                <td style="font-size:11.5px;color:#374151;padding:3px 0;vertical-align:top;">{processed_on}</td>
              </tr>
              <tr>
                <td style="font-size:11.5px;font-weight:700;color:#0f2044;padding:3px 10px 3px 0;white-space:nowrap;vertical-align:top;">Reference No.</td>
                <td style="font-size:11.5px;font-weight:700;color:#0f2044;padding:3px 0;vertical-align:top;">{ref_num}</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- spacer -->
  <tr><td style="height:16px;font-size:0;">&nbsp;</td></tr>

  {it_note_row}

  <!-- FOOTER -->
  <tr>
    <td style="background:#0f2044;padding:12px 24px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="vertical-align:middle;">
            <div style="font-size:10px;color:rgba(255,255,255,0.5);font-style:italic;">Developed for Pac-Biz Reporting &nbsp;&middot;&nbsp; MCerna &nbsp;&middot;&nbsp; v26.06.20</div>
          </td>
          <td align="right" style="vertical-align:middle;">
            <a href="https://mikewoocerna.github.io/Pac-Biz/masterlist_dashboard.html"
               style="display:inline-block;background:#004C97;color:#ffffff;padding:9px 20px;border-radius:4px;text-decoration:none;font-size:12px;font-weight:700;white-space:nowrap;">
              View Dashboard
            </a>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <!-- green bar -->
  <tr><td style="height:3px;background:#39B54A;font-size:0;border-radius:0 0 4px 4px;">&nbsp;</td></tr>

</table>

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

    eligible_timestamps = {
        str(row.get("Timestamp", "")).strip()
        for _, row in mv_df.iterrows()
        if is_ready_to_notify(row)
    }

    if seen is None:
        save_notified(eligible_timestamps)
        print(
            "[movement_notify] movement_notified.json missing — seeded "
            f"{len(eligible_timestamps)} existing processed movement(s); no emails sent."
        )
        return 0

    new_rows = [
        row for _, row in mv_df.iterrows()
        if is_ready_to_notify(row)
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
            bcc_list  = BCC_ALWAYS,
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

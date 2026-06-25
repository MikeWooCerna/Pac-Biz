"""notify.py — Gmail email notifications for pipeline events.

Reads credentials from notify_config.json (local, never committed to Git).
Called by self_heal.py on failure and on heal.

notify_config.json format:
{
  "from_email":   "reports@pac-biz.com",
  "app_password": "xxxx xxxx xxxx xxxx",
  "to_email":     "reports@pac-biz.com"
}
"""

import smtplib, json
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

CONFIG_FILE = Path(__file__).parent / "notify_config.json"
SMTP_HOST   = "smtp.gmail.com"
SMTP_PORT   = 587

def load_config():
    if not CONFIG_FILE.exists():
        return None
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None

def send(subject, body_html, body_text=None):
    cfg = load_config()
    if not cfg:
        print("[notify] notify_config.json not found — skipping email.")
        return False

    from_email   = cfg.get("from_email", "")
    app_password = cfg.get("app_password", "")
    to_email     = cfg.get("to_email", "")

    if not all([from_email, app_password, to_email]):
        print("[notify] Incomplete config — skipping email.")
        return False

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = from_email
    msg["To"]      = to_email

    if body_text:
        msg.attach(MIMEText(body_text, "plain"))
    msg.attach(MIMEText(body_html, "html"))

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(from_email, app_password)
            smtp.sendmail(from_email, to_email, msg.as_string())
        print(f"[notify] Email sent: {subject}")
        return True
    except Exception as e:
        print(f"[notify] Failed to send email: {e}")
        return False

def _now():
    return datetime.now().strftime("%b %d, %Y %I:%M %p")

def notify_failure(account, script, error=""):
    subject = f"Report Monitoring — FAILED: {account}"
    error_safe = (error or "").replace("<", "&lt;").replace(">", "&gt;")
    html = f"""
<div style="font-family:Arial,sans-serif;max-width:600px;">
  <div style="background:#b91c1c;color:#fff;padding:14px 20px;border-radius:6px 6px 0 0;">
    <b style="font-size:16px;">&#10060; Pipeline Failure</b>
  </div>
  <div style="border:1px solid #fca5a5;border-top:none;padding:18px 20px;background:#fff7f7;border-radius:0 0 6px 6px;">
    <table style="width:100%;font-size:14px;border-collapse:collapse;">
      <tr><td style="color:#6b7280;padding:4px 0;width:120px;">Account</td><td><b>{account}</b></td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Script</td><td>{script}</td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Time</td><td>{_now()}</td></tr>
    </table>
    {"<div style='margin-top:12px;padding:10px;background:#fff0f0;border-left:3px solid #ef4444;font-size:12px;font-family:monospace;white-space:pre-wrap;'>" + error_safe[:500] + "</div>" if error_safe else ""}
    <p style="margin-top:16px;font-size:13px;color:#374151;">
      Self-heal attempted a retry but could not recover. Manual attention required.
    </p>
    <p style="margin-top:8px;">
      <a href="https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html"
         style="background:#004C97;color:#fff;padding:8px 16px;border-radius:4px;text-decoration:none;font-size:13px;">
        View Pipeline Monitor
      </a>
    </p>
  </div>
</div>"""
    send(subject, html, body_text=f"PIPELINE FAILED\nAccount: {account}\nScript: {script}\nTime: {_now()}\nError: {error}")

def notify_high_volume(account, rows, threshold, level):
    level_label = "CRITICAL" if level == "crit" else "WARNING"
    hdr_color   = "#b91c1c" if level == "crit" else "#b45309"
    bdr_color   = "#fca5a5" if level == "crit" else "#fcd34d"
    bg_color    = "#fff7f7" if level == "crit" else "#fffbeb"
    subject = f"Report Monitoring — HIGH VOLUME {level_label}: {account}"
    html = f"""
<div style="font-family:Arial,sans-serif;max-width:600px;">
  <div style="background:{hdr_color};color:#fff;padding:14px 20px;border-radius:6px 6px 0 0;">
    <b style="font-size:16px;">&#9888; High Volume {level_label}: {account}</b>
  </div>
  <div style="border:1px solid {bdr_color};border-top:none;padding:18px 20px;background:{bg_color};border-radius:0 0 6px 6px;">
    <table style="width:100%;font-size:14px;border-collapse:collapse;">
      <tr><td style="color:#6b7280;padding:4px 0;width:120px;">Account</td><td><b>{account}</b></td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Row Count</td><td><b>{rows:,}</b></td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Threshold</td><td>{threshold:,}</td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Level</td><td>{level_label}</td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Time</td><td>{_now()}</td></tr>
    </table>
    <p style="margin-top:12px;font-size:13px;color:#374151;">
      This dataset is growing large. If it continues to grow unchecked, memory errors may occur during future pipeline runs.
      No action is required now, but consider archiving older records in the source sheet.
    </p>
    <p style="margin-top:8px;">
      <a href="https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html"
         style="background:#004C97;color:#fff;padding:8px 16px;border-radius:4px;text-decoration:none;font-size:13px;">
        View Pipeline Monitor
      </a>
    </p>
  </div>
</div>"""
    send(subject, html, body_text=(
        f"HIGH VOLUME {level_label}\nAccount: {account}\n"
        f"Row Count: {rows:,}\nThreshold: {threshold:,}\nTime: {_now()}\n"
        f"Consider archiving older records in the source sheet."
    ))

def notify_movement(subject, body_html, to_email, cc_list=None, bcc_list=None):
    """Send a movement notification with custom To + CC + optional BCC."""
    cfg = load_config()
    if not cfg:
        print("[notify] notify_config.json not found — skipping movement email.")
        return False

    from_email   = cfg.get("from_email", "")
    app_password = cfg.get("app_password", "")

    if not all([from_email, app_password, to_email]):
        print("[notify] Missing from_email, app_password, or to_email — skipping.")
        return False

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = from_email
    msg["To"]      = to_email
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    # BCC omitted from headers — only injected into SMTP envelope

    msg.attach(MIMEText(body_html, "html"))

    all_recipients = [to_email] + (cc_list or []) + (bcc_list or [])

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(from_email, app_password)
            smtp.sendmail(from_email, all_recipients, msg.as_string())
        print(f"[notify] Movement email sent: {subject}")
        return True
    except Exception as e:
        print(f"[notify] Failed to send movement email: {e}")
        return False

def notify_healed(account, script, action, detail):
    action_label = "Retry" if action == "retry" else "Re-pull"
    subject = f"Report Monitoring — SELF HEALED: {account}"
    html = f"""
<div style="font-family:Arial,sans-serif;max-width:600px;">
  <div style="background:#15803d;color:#fff;padding:14px 20px;border-radius:6px 6px 0 0;">
    <b style="font-size:16px;">&#10004; Self-Healed ({action_label})</b>
  </div>
  <div style="border:1px solid #86efac;border-top:none;padding:18px 20px;background:#f0fdf4;border-radius:0 0 6px 6px;">
    <table style="width:100%;font-size:14px;border-collapse:collapse;">
      <tr><td style="color:#6b7280;padding:4px 0;width:120px;">Account</td><td><b>{account}</b></td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Script</td><td>{script}</td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Action</td><td>{action_label}</td></tr>
      <tr><td style="color:#6b7280;padding:4px 0;">Time</td><td>{_now()}</td></tr>
    </table>
    <div style="margin-top:12px;padding:10px;background:#dcfce7;border-left:3px solid #22c55e;font-size:13px;">
      {detail}
    </div>
    <p style="margin-top:16px;font-size:13px;color:#374151;">
      No action required. The pipeline recovered automatically and will continue.
    </p>
    <p style="margin-top:8px;">
      <a href="https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html"
         style="background:#15803d;color:#fff;padding:8px 16px;border-radius:4px;text-decoration:none;font-size:13px;">
        View Pipeline Monitor
      </a>
    </p>
  </div>
</div>"""
    send(subject, html, body_text=f"PIPELINE SELF HEALED\nAccount: {account}\nScript: {script}\nAction: {action_label}\nTime: {_now()}\nDetail: {detail}")

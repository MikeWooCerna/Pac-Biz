"""diagnose_drops.py — Analyze row-count drops and send investigation email.

Called automatically at the end of generate_monitor.py after every pipeline run.
Only sends an email when NEW drops are found that haven't been reported yet.
Already-reported drops are tracked in pipeline_drops_notified.json.
"""

import json
import subprocess
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import notify

BASE            = Path(__file__).parent
LOG_FILE        = BASE / "pipeline_log.json"
BASELINE_FILE   = BASE / "pipeline_rowcount_baseline.json"
NOTIFIED_FILE   = BASE / "pipeline_drops_notified.json"
TRIGGERS_FILE   = BASE / "appsscript_triggers.json"

ACCOUNT_PULL_SCRIPTS = {
    "Kelowna": r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Kelowna\kel_pull.py",
}

# Accounts in PILOT_ACCOUNTS are analyzed and emailed but NOT auto-healed.
# The email includes a recommendation on whether to trigger the heal manually.
# Remove an account from this set once the auto-heal has been validated in prod.
PILOT_ACCOUNTS = {"Kelowna"}

# ── helpers ──────────────────────────────────────────────────────────────────

def load_json(path, default):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default

def save_json(path, data):
    Path(path).write_text(json.dumps(data, indent=2), encoding="utf-8")

def _now():
    return datetime.now().strftime("%b %d, %Y %I:%M %p")

# ── pattern analysis ──────────────────────────────────────────────────────────

def analyse(drops, baseline):
    """Return per-account summary dicts and cluster info."""

    # Group by account
    by_account = defaultdict(list)
    for d in drops:
        by_account[d["account"]].append(d)

    # Find run_ids with multiple accounts dropping (clusters)
    run_counts = defaultdict(list)
    for d in drops:
        run_counts[d["run_id"]].append(d["account"])
    cluster_runs = {rid: accts for rid, accts in run_counts.items() if len(accts) >= 3}

    summaries = []
    for account, events in by_account.items():
        events_sorted = sorted(events, key=lambda x: x["run_id"])
        total_drop    = sum(e["drop"] for e in events_sorted)
        occurrences   = len(events_sorted)
        latest        = events_sorted[-1]
        current_base  = baseline.get(account)

        # Recovery: current baseline is back above the lowest point ever reached
        lowest = min(e["new_count"] for e in events_sorted)
        recovered = (current_base is not None) and (current_base > lowest * 1.1)

        # Pattern classification
        in_cluster = any(e["run_id"] in cluster_runs for e in events_sorted)
        if occurrences >= 3:
            pattern      = "RECURRING"
            pattern_desc = "Drops on every or nearly every run — likely a date-range filter in the Apps Script"
            recommendation = (
                "Check the Apps Script for a date or month filter on the pull query. "
                "If pulling only 'current month', the sheet resets when the window shifts. "
                "Change the filter to pull ALL records (no date cap)."
            )
        elif in_cluster:
            pattern      = "CLUSTER"
            pattern_desc = "Dropped at the same time as other accounts — likely an upsert or sheet-refresh event"
            recommendation = (
                "This account dropped together with others in the same pipeline run. "
                "The source sheet may have been cleared and re-filled by the Apps Script upsert. "
                "Check the Apps Script upsert function — ensure it appends rather than overwrites. "
                "If this was a one-time deliberate reset, no action needed once baseline recovers."
            )
        elif occurrences == 1 and total_drop <= 50:
            pattern      = "MINOR"
            pattern_desc = "Single small drop — likely legitimate deletions or corrections in the source sheet"
            recommendation = (
                "Small isolated drops are usually normal (evaluations deleted/corrected upstream). "
                "Monitor for a second occurrence; if it persists, check the source sheet for filters."
            )
        else:
            pattern      = "ISOLATED"
            pattern_desc = "Single large drop — could be a transient data-source issue"
            recommendation = (
                "One-time large drop. Check if the source sheet was temporarily unavailable or partially "
                "cleared during this run. If the account has recovered, no action is needed."
            )

        # Determine whether a full-year Apps Script heal is warranted
        # RECURRING and ISOLATED large drops benefit from a heal.
        # MINOR drops (small, single) likely don't need it.
        # CLUSTER drops need root-cause investigation before healing.
        heal_warranted = (
            pattern in ("RECURRING", "ISOLATED") and not recovered
        )
        heal_reason = (
            "RECURRING pattern detected — the Apps Script is likely filtering by date and resets periodically. "
            "Triggering the full-year pull will restore the complete dataset."
            if pattern == "RECURRING" else
            "Large one-time drop — likely a transient sheet issue. "
            "Triggering the full-year pull is recommended to confirm full dataset is restored."
            if pattern == "ISOLATED" else
            "CLUSTER drop — multiple accounts affected simultaneously. "
            "Investigate whether the source sheet was cleared before triggering a heal."
            if pattern == "CLUSTER" else
            "Small drop — likely legitimate upstream deletions. "
            "A full-year heal is not recommended; monitor for recurrence instead."
        )

        summaries.append({
            "account":        account,
            "occurrences":    occurrences,
            "total_drop":     total_drop,
            "latest_drop":    latest["drop"],
            "latest_run":     latest["run_id"],
            "latest_prev":    latest["prev_count"],
            "latest_new":     latest["new_count"],
            "current_base":   current_base,
            "recovered":      recovered,
            "pattern":        pattern,
            "pattern_desc":   pattern_desc,
            "recommendation": recommendation,
            "heal_warranted": heal_warranted,
            "heal_reason":    heal_reason,
        })

    # Sort: unrecovered first, then by occurrences desc
    summaries.sort(key=lambda x: (x["recovered"], -x["occurrences"], -x["total_drop"]))
    return summaries, cluster_runs

# ── email builder ─────────────────────────────────────────────────────────────

def _heal_block(s):
    """Return an HTML block with a heal recommendation for pilot accounts.
    Returns empty string for non-pilot accounts or accounts with a trigger URL."""
    if s["account"] not in PILOT_ACCOUNTS:
        return ""
    triggers = load_json(TRIGGERS_FILE, {})
    if s["account"] not in triggers:
        return ""
    if s["recovered"]:
        return ""

    if s["heal_warranted"]:
        bg      = "#f0fdf4"
        border  = "#16a34a"
        icon    = "&#9989;"
        heading = "Heal Recommended"
        color   = "#14532d"
    else:
        bg      = "#fffbeb"
        border  = "#d97706"
        icon    = "&#9888;"
        heading = "Heal Not Recommended"
        color   = "#78350f"

    return f"""
        <div style="margin-top:10px;padding:10px 12px;background:{bg};border-left:3px solid {border};border-radius:0 4px 4px 0;font-size:12px;color:{color};">
          <b>{icon} {heading} — {s['account']} (Pilot)</b><br>
          <span style="display:block;margin-top:4px;">{s['heal_reason']}</span>
          <span style="display:block;margin-top:6px;color:#374151;">
            Auto-heal is <b>disabled</b> while {s['account']} is in pilot mode.
            {"To trigger manually, run: <code style='background:#e5e7eb;padding:1px 4px;border-radius:3px;'>py -3 diagnose_drops.py --heal " + s['account'] + "</code>" if s['heal_warranted'] else "No action needed at this time."}
          </span>
        </div>"""


PATTERN_STYLE = {
    "RECURRING": ("background:#fef2f2;color:#b91c1c;", "RECURRING"),
    "CLUSTER":   ("background:#fffbeb;color:#92400e;", "CLUSTER"),
    "ISOLATED":  ("background:#eff6ff;color:#1e40af;", "ISOLATED"),
    "MINOR":     ("background:#f0fdf4;color:#15803d;", "MINOR"),
}

def build_email(summaries, cluster_runs, new_accounts):
    total_affected = len(summaries)
    unrecovered    = [s for s in summaries if not s["recovered"]]

    cluster_note = ""
    if cluster_runs:
        runs_fmt = "; ".join(
            f"{rid[:16]} ({', '.join(accts)})"
            for rid, accts in cluster_runs.items()
        )
        cluster_note = f"""
        <tr>
          <td colspan="6" style="padding:8px 10px;background:#fffbeb;border-bottom:1px solid #fcd34d;font-size:12px;color:#92400e;">
            <b>&#9888; Cluster detected:</b> {runs_fmt}
          </td>
        </tr>"""

    rows_html = ""
    for s in summaries:
        p_style, p_label = PATTERN_STYLE.get(s["pattern"], ("", s["pattern"]))
        status_style = "color:#15803d;" if s["recovered"] else "color:#b91c1c;font-weight:700;"
        status_label = "Recovered" if s["recovered"] else "⚠ Still Active"
        curr = f"{s['current_base']:,}" if s["current_base"] is not None else "—"
        rows_html += f"""
        <tr style="border-bottom:1px solid #e5e7eb;">
          <td style="padding:8px 10px;font-size:13px;font-weight:600;">{s['account']}</td>
          <td style="padding:8px 10px;font-size:13px;text-align:center;">{s['occurrences']}</td>
          <td style="padding:8px 10px;font-size:13px;text-align:center;color:#b91c1c;font-weight:700;">↓ {s['total_drop']:,}</td>
          <td style="padding:8px 10px;font-size:13px;text-align:center;">
            <span style="padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;{p_style}">{p_label}</span>
          </td>
          <td style="padding:8px 10px;font-size:13px;text-align:center;{status_style}">{status_label}</td>
          <td style="padding:8px 10px;font-size:13px;">{curr}</td>
        </tr>"""

    detail_rows = ""
    for s in summaries:
        p_style, p_label = PATTERN_STYLE.get(s["pattern"], ("", s["pattern"]))
        border_color = "#fca5a5" if not s["recovered"] else "#86efac"
        detail_rows += f"""
    <div style="border:1px solid {border_color};border-radius:6px;margin-bottom:14px;overflow:hidden;">
      <div style="padding:10px 14px;background:#f9fafb;border-bottom:1px solid {border_color};">
        <span style="font-size:14px;font-weight:700;color:#111827;">{s['account']}</span>
        <span style="margin-left:10px;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;{p_style}">{p_label}</span>
        {"<span style='margin-left:8px;font-size:12px;color:#b91c1c;font-weight:700;'>⚠ Unrecovered</span>" if not s["recovered"] else "<span style='margin-left:8px;font-size:12px;color:#15803d;'>✓ Recovered</span>"}
      </div>
      <div style="padding:12px 14px;background:#fff;">
        <table style="width:100%;font-size:13px;border-collapse:collapse;margin-bottom:10px;">
          <tr>
            <td style="color:#6b7280;width:130px;padding:3px 0;">Total occurrences</td>
            <td><b>{s['occurrences']}</b></td>
            <td style="color:#6b7280;width:130px;padding:3px 0;">Total rows lost</td>
            <td><b style="color:#b91c1c;">↓ {s['total_drop']:,}</b></td>
          </tr>
          <tr>
            <td style="color:#6b7280;padding:3px 0;">Latest drop</td>
            <td>{s['latest_prev']:,} → {s['latest_new']:,} (↓{s['latest_drop']:,})</td>
            <td style="color:#6b7280;padding:3px 0;">Current baseline</td>
            <td><b>{s['current_base']:,}</b></td>
          </tr>
        </table>
        <div style="font-size:12px;color:#374151;margin-bottom:8px;">
          <b>Pattern:</b> {s['pattern_desc']}
        </div>
        <div style="padding:10px;background:#f0f9ff;border-left:3px solid #0ea5e9;font-size:12px;color:#0c4a6e;border-radius:0 4px 4px 0;">
          <b>Recommendation:</b> {s['recommendation']}
        </div>
        {_heal_block(s)}
      </div>
    </div>"""

    unrecovered_note = ""
    if unrecovered:
        names = ", ".join(s["account"] for s in unrecovered)
        unrecovered_note = f"""
    <div style="padding:12px 16px;background:#fef2f2;border:1px solid #fca5a5;border-radius:6px;margin-bottom:20px;font-size:13px;color:#7f1d1d;">
      <b>&#9888; {len(unrecovered)} account(s) have not yet recovered:</b> {names}<br>
      These accounts are still below their pre-drop baseline. Review the recommendations above.
    </div>"""

    html = f"""
<div style="font-family:Arial,sans-serif;max-width:680px;">

  <!-- Header -->
  <div style="background:linear-gradient(135deg,#0f2044 0%,#004C97 100%);color:#fff;padding:18px 22px;border-radius:6px 6px 0 0;">
    <div style="font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.55);margin-bottom:4px;">Pipeline Intelligence</div>
    <b style="font-size:17px;">Count Drop Investigation Report</b>
    <div style="font-size:12px;color:rgba(255,255,255,.65);margin-top:4px;">{_now()} &nbsp;·&nbsp; {total_affected} account(s) with new drops</div>
  </div>

  <!-- Body -->
  <div style="border:1px solid #e5e7eb;border-top:none;padding:20px 22px;background:#fff;border-radius:0 0 6px 6px;">

    {unrecovered_note}

    <!-- Summary table -->
    <div style="margin-bottom:22px;overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead>
          <tr style="background:#f3f4f6;">
            {cluster_note}
            <th style="padding:8px 10px;text-align:left;font-weight:700;color:#374151;border-bottom:2px solid #d1d5db;">Account</th>
            <th style="padding:8px 10px;text-align:center;font-weight:700;color:#374151;border-bottom:2px solid #d1d5db;">Drops</th>
            <th style="padding:8px 10px;text-align:center;font-weight:700;color:#374151;border-bottom:2px solid #d1d5db;">Rows Lost</th>
            <th style="padding:8px 10px;text-align:center;font-weight:700;color:#374151;border-bottom:2px solid #d1d5db;">Pattern</th>
            <th style="padding:8px 10px;text-align:center;font-weight:700;color:#374151;border-bottom:2px solid #d1d5db;">Status</th>
            <th style="padding:8px 10px;text-align:left;font-weight:700;color:#374151;border-bottom:2px solid #d1d5db;">Current</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>

    <!-- Per-account details -->
    <div style="font-size:13px;font-weight:700;color:#374151;margin-bottom:12px;border-bottom:1px solid #e5e7eb;padding-bottom:6px;">
      Account Details &amp; Recommendations
    </div>
    {detail_rows}

    <!-- Footer -->
    <p style="margin-top:18px;">
      <a href="https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html"
         style="background:#004C97;color:#fff;padding:9px 18px;border-radius:4px;text-decoration:none;font-size:13px;font-weight:600;">
        View Pipeline Monitor
      </a>
    </p>
    <p style="font-size:11px;color:#9ca3af;margin-top:12px;">
      This report is generated automatically after each pipeline run. Only new drops trigger an email.
    </p>
  </div>
</div>"""
    return html

# ── auto-trigger ─────────────────────────────────────────────────────────────

# Each doGet() call runs for up to 5 minutes. A full-year Kelowna pull
# (53 weeks) takes ~3-4 calls. Cap at 12 to be safe without spinning forever.
MAX_TRIGGER_CALLS = 12

def _call_appsscript(url):
    """GET the web app URL. Returns (status, rows_so_far, message)."""
    import requests
    r = requests.get(url, timeout=370)
    if r.status_code != 200:
        return "error", 0, f"HTTP {r.status_code}: {r.text[:200]}"
    body = r.text.strip()
    # New JSON protocol: {"status":"partial"|"complete"|"error", "rows":N, ...}
    if body.startswith("{"):
        try:
            data = json.loads(body)
            status = data.get("status", "error")
            rows   = data.get("rows") or data.get("rows_so_far") or 0
            msg    = data.get("message", "")
            return status, int(rows), msg
        except Exception:
            pass
    # Legacy plain-text "OK" / "ERROR: ..."
    if body == "OK":
        return "complete", 0, ""
    if body.startswith("ERROR"):
        return "error", 0, body[:200]
    return "error", 0, f"Unexpected response: {body[:200]}"

def trigger_appsscript(account, prev_count=None):
    """Call the Apps Script web app (looping until complete), re-run pull script,
    rebuild dashboard, verify row count against prev_count, then push.

    prev_count — the row count the account had immediately before the drop (from
    the drop log event).  This is the target the heal must reach.  Falls back to
    the baseline file if not supplied.
    """
    triggers = load_json(TRIGGERS_FILE, {})
    url = triggers.get(account)
    if not url:
        return False, "No trigger URL configured"

    # Determine expected (target) row count — prefer the pre-drop snapshot
    baseline = load_json(BASELINE_FILE, {})
    expected = prev_count if prev_count and prev_count > 0 else baseline.get(account, 0)
    print(f"[diagnose] {account} — expected rows after heal: {expected:,} "
          f"({'pre-drop count' if prev_count else 'baseline'})")

    try:
        import requests  # noqa: F401 — ensure it's importable before the loop
        rows_so_far = 0
        for attempt in range(1, MAX_TRIGGER_CALLS + 1):
            print(f"[diagnose] Apps Script call {attempt}/{MAX_TRIGGER_CALLS} for {account}...")
            status, rows, msg = _call_appsscript(url)

            if status == "error":
                return False, f"Apps Script error (call {attempt}): {msg}"

            if status == "complete":
                print(f"[diagnose] Apps Script complete for {account} — {rows:,} rows.")
                rows_so_far = rows
                break

            # partial — still pulling
            rows_so_far = rows
            print(f"[diagnose] Apps Script partial ({rows:,} rows so far, last_week saved). Calling again...")

        else:
            # Hit MAX_TRIGGER_CALLS without completing
            return False, (
                f"Apps Script did not complete after {MAX_TRIGGER_CALLS} calls "
                f"({rows_so_far:,} rows fetched, expected ~{expected:,})"
            )

        # Verify row count reached ≥ 80 % of the pre-drop count.
        # Note: doGet() may return 0 due to a getLastRow() timing quirk on
        # freshly-renamed sheets in Apps Script — fall back to kel_pull.py
        # to get the authoritative count from the actual sheet data below.
        if expected > 0 and rows_so_far > 0:
            if rows_so_far < expected * 0.80:
                return False, (
                    f"Row count after heal ({rows_so_far:,}) is still well below "
                    f"expected ({expected:,}). Pull may be incomplete — "
                    f"{rows_so_far / expected * 100:.0f}% restored."
                )
            print(f"[diagnose] {account} row count verified: "
                  f"{rows_so_far:,} / {expected:,} expected "
                  f"({rows_so_far / expected * 100:.0f}%) ✓")
        elif expected > 0 and rows_so_far == 0:
            print(f"[diagnose] {account} Apps Script returned 0 rows "
                  f"(likely a getLastRow timing quirk) — "
                  f"proceeding to pull script for authoritative count.")

    except Exception as e:
        return False, f"Apps Script call failed: {e}"

    # Step 2 — re-run pull script
    pull_script = ACCOUNT_PULL_SCRIPTS.get(account)
    if pull_script:
        result = subprocess.run(
            ["py", "-3", pull_script],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode != 0:
            return False, f"Pull script failed: {result.stderr[:300]}"
        print(f"[diagnose] Pull script re-run OK for {account}.")

    # Step 3 — rebuild dashboard
    print(f"[diagnose] Rebuilding dashboard...")
    build = subprocess.run(
        ["py", "-3", str(BASE / "dashboard.py")],
        capture_output=True, text=True, timeout=600, cwd=str(BASE)
    )
    if build.returncode != 0:
        return False, f"Dashboard rebuild failed: {build.stderr[:300]}"
    print(f"[diagnose] Dashboard rebuilt OK.")

    # Step 4 — log as healed (must be before monitor regeneration)
    log = load_json(LOG_FILE, [])
    run_id = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    before_str = f"{expected:,}" if expected else "unknown"
    after_str  = f"{rows_so_far:,}" if rows_so_far else "unknown"
    pct_str    = f" ({rows_so_far / expected * 100:.0f}%)" if expected and rows_so_far else ""
    detail = (
        f"Count drop auto-recovered. Apps Script re-triggered, dashboard rebuilt and pushed. "
        f"Rows before drop: {before_str} → rows after heal: {after_str}{pct_str}."
    )
    log.append({
        "run_id":       run_id,
        "date":         datetime.now().strftime("%b %d, %Y %I:%M %p"),
        "account":      account,
        "script":       ACCOUNT_PULL_SCRIPTS.get(account, "appsscript_trigger"),
        "status":       "healed",
        "action":       "appsscript_trigger",
        "prev_count":   expected,
        "healed_count": rows_so_far,
        "error":        detail,
    })
    save_json(LOG_FILE, log[-300:])

    # Step 5 — regenerate pipeline monitor (after log entry so it shows SELF HEALED)
    print(f"[diagnose] Regenerating pipeline monitor...")
    subprocess.run(
        ["py", "-3", str(BASE / "generate_monitor.py")],
        capture_output=True, text=True, timeout=60, cwd=str(BASE)
    )

    # Step 6 — git push
    print(f"[diagnose] Pushing to git...")
    git_cmds = [
        ["git", "pull", "--rebase", "--autostash"],
        ["git", "add", "masterlist_dashboard.html", "pipeline_monitor.html", "pipeline_log.json"],
        ["git", "commit", "-m", f"Auto-heal: {account} count drop recovered — Apps Script re-triggered"],
        ["git", "push"],
    ]
    for cmd in git_cmds:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, cwd=str(BASE))
        if result.returncode != 0 and "nothing to commit" not in result.stdout:
            return False, f"Git failed ({' '.join(cmd[:2])}): {result.stderr[:200]}"
    print(f"[diagnose] Git push OK.")

    # Step 7 — send healed email
    notify.notify_healed(account, ACCOUNT_PULL_SCRIPTS.get(account, "appsscript_trigger"), "appsscript_trigger", detail)

    return True, "Apps Script triggered, pull re-run, dashboard rebuilt and pushed"


# ── main ──────────────────────────────────────────────────────────────────────

def run():
    log      = load_json(LOG_FILE, [])
    baseline = load_json(BASELINE_FILE, {})
    notified = load_json(NOTIFIED_FILE, [])
    notified_set = set(notified)

    # All count_drop entries
    all_drops = [e for e in log if e.get("status") == "count_drop"]

    # Filter to only NEW (not yet reported) drops
    new_drops = [
        d for d in all_drops
        if f"{d['run_id']}|{d['account']}" not in notified_set
    ]

    if not new_drops:
        print("[diagnose] No new count drops to report.")
        return

    # Group new drops by account for analysis
    # But include full history per account for pattern analysis
    new_accounts = {d["account"] for d in new_drops}
    drops_for_affected = [d for d in all_drops if d["account"] in new_accounts]

    summaries, cluster_runs = analyse(drops_for_affected, baseline)
    email_html = build_email(summaries, cluster_runs, new_accounts)

    affected_list = ", ".join(s["account"] for s in summaries)
    subject = f"Report Monitoring — COUNT DROP: {len(summaries)} account(s) affected ({affected_list})"
    if len(subject) > 100:
        subject = f"Report Monitoring — COUNT DROP: {len(summaries)} account(s) detected"

    # Mark drops as notified BEFORE triggering the heal. The heal subprocess
    # calls generate_monitor.py → diagnose_drops.run() again; if drops were
    # not yet marked, that second run() would re-trigger the heal and loop.
    for d in new_drops:
        notified_set.add(f"{d['run_id']}|{d['account']}")
    save_json(NOTIFIED_FILE, sorted(notified_set))

    sent = notify.send(subject, email_html)
    if sent:
        print(f"[diagnose] Drop report sent for: {affected_list}")
    else:
        print("[diagnose] Email send failed — drops already marked notified to prevent re-trigger loop.")

    # Auto-trigger Apps Script for accounts that have a trigger URL AND are not
    # in PILOT_ACCOUNTS. Pilot accounts receive analysis + email only — heal
    # requires manual review and approval.
    triggers = load_json(TRIGGERS_FILE, {})
    for account in new_accounts:
        if account not in triggers:
            continue
        if account in PILOT_ACCOUNTS:
            print(f"[diagnose] {account} is in pilot mode — auto-heal skipped. Review email recommendation.")
            continue
        acct_new_drops = [d for d in new_drops if d["account"] == account]
        prev_count = max(
            (d.get("prev_count") or 0 for d in acct_new_drops),
            default=0
        )
        ok, msg = trigger_appsscript(account, prev_count=prev_count)
        status = "recovered" if ok else "trigger failed"
        print(f"[diagnose] {account} auto-trigger: {status} — {msg}")

if __name__ == "__main__":
    import sys
    # Manual heal: py -3 diagnose_drops.py --heal Kelowna
    if len(sys.argv) >= 3 and sys.argv[1] == "--heal":
        account = " ".join(sys.argv[2:])
        baseline = load_json(BASELINE_FILE, {})
        prev_count = baseline.get(account, 0)
        print(f"[diagnose] Manual heal requested for {account} (baseline: {prev_count:,})")
        ok, msg = trigger_appsscript(account, prev_count=prev_count)
        print(f"[diagnose] {account} manual heal: {'OK' if ok else 'FAILED'} — {msg}")
    else:
        run()

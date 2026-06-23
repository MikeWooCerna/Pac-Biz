"""self_heal.py — pipeline self-healing utilities.

Modes:
  run-step <script>   Run a script with 1 automatic retry (60s wait).
                      If script is dashboard.py, runs fix_footgun.py before each attempt.
                      Writes final stderr to step_err.tmp on permanent failure.
  heal-drops          Re-pulls any account whose row count dropped vs baseline.
                      Ignores drops under 5% (normal fluctuation).
"""

import subprocess, sys, time, json
from pathlib import Path

BASE      = Path(__file__).parent
STEP_ERR  = BASE / "step_err.tmp"
BASELINE  = BASE / "pipeline_rowcount_baseline.json"

RETRY_WAIT = 60   # seconds between pull retries
BUILD_WAIT = 30   # seconds between build retries

ACCOUNT_PULL_MAP = {
    "Masterlist":      ("masterlist_fetch.py", str(BASE)),
    "Coaching":        ("asana_pull.py",    r"C:\Users\Mike Woo Cerna\Documents\PB\Coaching"),
    "M7":              ("m7_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\M7"),
    "Parentis Health": ("parentis_pull.py", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Parentis Health"),
    "Britelift":       ("britelift_pull.py",r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift"),
    "Britelift Chat":  ("britelift_pull.py",r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift Chat"),
    "RideX":           ("Ridex_pull.py",    r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\RideX"),
    "Hamilton":        ("Hamilton_pull.py", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Hamilton"),
    "Skyline":         ("Skyline_pull.py",  r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Skyline"),
    "VIP":             ("vip_pull.py",      r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\VIP"),
    "C&H":             ("ch_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\C&H"),
    "Reno Cab":        ("rc_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Reno Cab"),
    "Trans Iowa":      ("ti_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Trans Iowa"),
    "Data Carz":       ("dc_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Data Carz"),
    "Associated Cab":  ("ac_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Associated Cab"),
    "Ollies":          ("ol_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Ollies"),
    "Circle Taxi":     ("ct_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Circle Taxi"),
    "YCOV":            ("ycov_pull.py",     r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCOV"),
    "Kelowna":         ("kel_pull.py",      r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Kelowna"),
    "Vermont":         ("vt_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Vermont"),
    "YCDC":            ("ycdc_pull.py",     r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCDC"),
    "Blueline":        ("bl_pull.py",       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Blueline"),
}

ACCOUNT_FILES = {
    "Masterlist":      str(BASE / "masterlist_cache.csv"),
    "Coaching":        r"C:\Users\Mike Woo Cerna\Documents\PB\Coaching\Output\coaching_logs.xlsx",
    "M7":              r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\M7\M7_RAW.xlsx",
    "Parentis Health": r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Parentis Health\PARENTIS_RAW.xlsx",
    "Britelift":       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift\BRITELIFT_RAW.xlsx",
    "Britelift Chat":  r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift Chat\BLC_RAW.xlsx",
    "RideX":           r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\RideX\RIDEX_RAW.xlsx",
    "Hamilton":        r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Hamilton\HAMILTON_RAW.xlsx",
    "Skyline":         r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Skyline\SKYLINE_RAW.xlsx",
    "VIP":             r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\VIP\VIP_RAW.xlsx",
    "C&H":             r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\C&H\CH_RAW.xlsx",
    "Reno Cab":        r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Reno Cab\RC_RAW.xlsx",
    "Trans Iowa":      r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Trans Iowa\TI_RAW.xlsx",
    "Data Carz":       r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Data Carz\DC_RAW.xlsx",
    "Associated Cab":  r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Associated Cab\AC_RAW.xlsx",
    "Ollies":          r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Ollies\OL_RAW.xlsx",
    "Circle Taxi":     r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Circle Taxi\CT_RAW.xlsx",
    "YCOV":            r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCOV\YCOV_RAW.xlsx",
    "Kelowna":         r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Kelowna\KEL_RAW.xlsx",
    "Vermont":         r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Vermont\VT_RAW.xlsx",
    "YCDC":            r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCDC\YCDC_RAW.xlsx",
    "Blueline":        r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Blueline\BL_RAW.xlsx",
}

def get_row_count(file_path):
    try:
        p = Path(file_path)
        if not p.exists():
            return None
        if p.suffix.lower() == ".csv":
            return max(0, sum(1 for _ in p.open("r", encoding="utf-8", errors="replace")) - 1)
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        count = max(0, (ws.max_row or 1) - 1)
        wb.close()
        return count
    except Exception:
        return None

def load_baseline():
    try:
        if BASELINE.exists():
            return json.loads(BASELINE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

# ── Mode: run-step ─────────────────────────────────────────────────────────────

def run_step(script):
    is_build = (script == "dashboard.py")
    wait     = BUILD_WAIT if is_build else RETRY_WAIT

    for attempt in range(2):
        if is_build:
            subprocess.run([sys.executable, str(BASE / "fix_footgun.py")])

        result = subprocess.run(
            [sys.executable, script],
            stderr=subprocess.PIPE,
            text=True,
            errors="replace"
        )

        if result.returncode == 0:
            sys.exit(0)

        if result.stderr:
            print(result.stderr, file=sys.stderr, flush=True)

        if attempt == 0:
            label = "Build" if is_build else script
            print(f"[self-heal] {label} failed — retrying in {wait}s...", flush=True)
            time.sleep(wait)
        else:
            STEP_ERR.write_text(result.stderr or "", encoding="utf-8", errors="replace")
            sys.exit(1)

# ── Mode: heal-drops ──────────────────────────────────────────────────────────

def heal_drops():
    baseline = load_baseline()
    if not baseline:
        print("[self-heal] No baseline — skipping drop check.")
        return

    any_drop = False
    for account, prev_count in baseline.items():
        file_path = ACCOUNT_FILES.get(account)
        if not file_path:
            continue
        current = get_row_count(file_path)
        if current is None or current >= prev_count:
            continue

        drop_pct = (prev_count - current) / prev_count * 100
        if drop_pct < 5:
            continue  # Ignore minor fluctuations

        any_drop = True
        pull_info = ACCOUNT_PULL_MAP.get(account)
        if not pull_info:
            print(f"[self-heal] {account}: drop detected but no pull map entry — skipping.")
            continue

        script, directory = pull_info
        print(f"[self-heal] {account}: dropped {prev_count:,} → {current:,} ({drop_pct:.1f}%). Re-pulling...", flush=True)

        result = subprocess.run(
            [sys.executable, script],
            cwd=directory,
            stderr=subprocess.PIPE,
            text=True,
            errors="replace"
        )

        if result.returncode != 0:
            print(f"[self-heal] {account}: re-pull failed. {(result.stderr or '')[:200]}", flush=True)
            continue

        new_count = get_row_count(file_path)
        if new_count is not None and new_count >= prev_count:
            print(f"[self-heal] {account}: recovered — {new_count:,} rows. ✓", flush=True)
        else:
            recovered = new_count if new_count is not None else current
            print(f"[self-heal] {account}: re-pulled but count still low ({recovered:,}). Likely a source-side issue.", flush=True)

    if not any_drop:
        print("[self-heal] No significant drops detected.", flush=True)

# ── Main ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    if mode == "run-step":
        if len(sys.argv) < 3:
            print("Usage: self_heal.py run-step <script>", file=sys.stderr)
            sys.exit(1)
        run_step(sys.argv[2])
    elif mode == "heal-drops":
        heal_drops()
    else:
        print(f"Unknown mode: '{mode}'", file=sys.stderr)
        sys.exit(1)

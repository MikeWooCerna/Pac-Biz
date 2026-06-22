"""Pipeline step logger — called by update_coaching_dashboard_auto.bat after each pull step."""
import sys, json, subprocess
from datetime import datetime
from pathlib import Path

def push_live(label):
    """Regenerate monitor HTML and push to GitHub Pages — best effort, never aborts pipeline."""
    base = Path(__file__).parent
    try:
        subprocess.run([sys.executable, str(base / "generate_monitor.py")],
                       cwd=str(base), capture_output=True, timeout=60)
        subprocess.run(["git", "add", "pipeline_status.json", "pipeline_monitor.html"],
                       cwd=str(base), capture_output=True, timeout=30)
        r = subprocess.run(["git", "diff", "--cached", "--quiet"],
                           cwd=str(base), capture_output=True, timeout=30)
        if r.returncode != 0:
            subprocess.run(["git", "commit", "-m", f"[live] {label}"],
                           cwd=str(base), capture_output=True, timeout=30)
            subprocess.run(["git", "pull", "--rebase", "--autostash"],
                           cwd=str(base), capture_output=True, timeout=60)
            subprocess.run(["git", "push"],
                           cwd=str(base), capture_output=True, timeout=60)
        print(f"  [monitor] live push ok: {label}")
    except Exception as e:
        print(f"  [monitor] live push skipped: {e}")

STATUS_FILE = Path(__file__).parent / "pipeline_status.json"
ERR_TMP     = Path(__file__).parent / "step_err.tmp"

def load():
    if STATUS_FILE.exists():
        try:
            return json.loads(STATUS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return None

def save(data):
    STATUS_FILE.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")

def read_error():
    try:
        if ERR_TMP.exists():
            text = ERR_TMP.read_text(encoding="utf-8", errors="replace").strip()
            if text:
                lines = [l for l in text.splitlines() if l.strip()]
                return "\n".join(lines[-4:]) if lines else None
    except Exception:
        pass
    return None

cmd = sys.argv[1] if len(sys.argv) > 1 else ""

if cmd == "init":
    now = datetime.now()
    data = {
        "run_id":      now.strftime("%Y-%m-%dT%H:%M:%S"),
        "started_at":  now.isoformat(),
        "finished_at": None,
        "status":      "running",
        "failed_at":   None,
        "steps":       []
    }
    save(data)
    # Clear leftover error from previous run
    try:
        ERR_TMP.write_text("", encoding="utf-8")
    except Exception:
        pass
    print(f"Pipeline log initialized: {data['run_id']}")

elif cmd == "step":
    account   = sys.argv[2] if len(sys.argv) > 2 else "Unknown"
    script    = sys.argv[3] if len(sys.argv) > 3 else "unknown.py"
    exit_code = int(sys.argv[4]) if len(sys.argv) > 4 else 1

    data = load()
    if data is None:
        now = datetime.now()
        data = {"run_id": now.strftime("%Y-%m-%dT%H:%M:%S"), "started_at": now.isoformat(),
                "finished_at": None, "status": "running", "failed_at": None, "steps": []}

    error_msg = read_error() if exit_code != 0 else None

    data["steps"].append({
        "account":   account,
        "script":    script,
        "exit_code": exit_code,
        "status":    "pass" if exit_code == 0 else "fail",
        "timestamp": datetime.now().isoformat(),
        "error":     error_msg
    })

    if exit_code != 0 and data.get("status") != "failed":
        data["status"]    = "failed"
        data["failed_at"] = account

    save(data)
    print(f"Logged: {account} ({script}) -> {'PASS' if exit_code == 0 else 'FAIL'}")
    if error_msg:
        print(f"  Error: {error_msg[:120]}")
    push_live(account)

elif cmd == "finish":
    outcome = sys.argv[2] if len(sys.argv) > 2 else "success"
    data = load()
    if data:
        data["finished_at"] = datetime.now().isoformat()
        if outcome == "success":
            data["status"] = "success"
        elif data.get("status") != "failed":
            data["status"] = "failed"
        save(data)
    print(f"Pipeline finished: {outcome}")

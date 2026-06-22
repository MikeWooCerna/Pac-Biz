"""Pipeline step logger — called by update_coaching_dashboard_auto.bat after each pull step."""
import sys, json
from datetime import datetime
from pathlib import Path

STATUS_FILE = Path(__file__).parent / "pipeline_status.json"

def load():
    if STATUS_FILE.exists():
        try:
            return json.loads(STATUS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return None

def save(data):
    STATUS_FILE.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")

cmd = sys.argv[1] if len(sys.argv) > 1 else ""

if cmd == "init":
    now = datetime.now()
    data = {
        "run_id":     now.strftime("%Y-%m-%dT%H:%M:%S"),
        "started_at": now.isoformat(),
        "finished_at": None,
        "status":     "running",
        "failed_at":  None,
        "steps":      []
    }
    save(data)
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

    data["steps"].append({
        "account":   account,
        "script":    script,
        "exit_code": exit_code,
        "status":    "pass" if exit_code == 0 else "fail",
        "timestamp": datetime.now().isoformat()
    })

    if exit_code != 0 and data.get("status") != "failed":
        data["status"]    = "failed"
        data["failed_at"] = account

    save(data)
    print(f"Logged: {account} ({script}) -> {'PASS' if exit_code == 0 else 'FAIL'}")

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

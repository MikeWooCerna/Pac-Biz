"""Append pipeline run status snapshots and build daily uptime datasets.

Inputs:
  - pipeline_status.json: latest pipeline run, overwritten each run
  - pipeline_log.json: incident log, used for incident counts by day

Outputs:
  - pipeline_run_history.json: one record per pipeline run
  - pipeline_step_history.csv: one record per pipeline step per run
  - pipeline_daily_uptime.csv/json: daily uptime summary for dashboarding

The files are intentionally simple so they can feed a future HTML dashboard,
Excel, Power BI, or another monitoring page without scraping pipeline_monitor.
"""

from __future__ import annotations

import csv
import json
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path


BASE = Path(__file__).parent
STATUS_FILE = BASE / "pipeline_status.json"
LOG_FILE = BASE / "pipeline_log.json"
RUN_HISTORY_FILE = BASE / "pipeline_run_history.json"
STEP_HISTORY_FILE = BASE / "pipeline_step_history.csv"
DAILY_UPTIME_CSV = BASE / "pipeline_daily_uptime.csv"
DAILY_UPTIME_JSON = BASE / "pipeline_daily_uptime.json"

MAX_RUN_HISTORY = 2500
STEP_FIELDS = [
    "run_id",
    "run_date",
    "started_at",
    "finished_at",
    "pipeline_status",
    "account",
    "script",
    "status",
    "exit_code",
    "timestamp",
    "error",
]
DAILY_FIELDS = [
    "date",
    "runs",
    "successful_runs",
    "partial_runs",
    "failed_runs",
    "uptime_pct",
    "steps",
    "passed_steps",
    "failed_steps",
    "step_uptime_pct",
    "incident_count",
    "count_drop_count",
    "self_healed_count",
    "high_volume_count",
]


def _read_json(path: Path, default):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _write_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False, default=str), encoding="utf-8")


def _date_from_iso(value: str | None) -> str:
    if not value:
        return ""
    try:
        return datetime.fromisoformat(value).date().isoformat()
    except Exception:
        return str(value)[:10]


def _duration_seconds(started_at: str | None, finished_at: str | None) -> float | None:
    if not started_at or not finished_at:
        return None
    try:
        started = datetime.fromisoformat(started_at)
        finished = datetime.fromisoformat(finished_at)
        return round((finished - started).total_seconds(), 3)
    except Exception:
        return None


def _status_rank(status: str) -> int:
    return {"success": 3, "partial": 2, "failed": 1, "running": 0}.get(status, 0)


def _normalize_pipeline_status(status: str | None) -> str:
    if status == "success":
        return "success"
    if status == "partial":
        return "partial"
    if status in {"failed", "fail"}:
        return "failed"
    return status or "unknown"


def append_current_run() -> tuple[list[dict], list[dict]]:
    status = _read_json(STATUS_FILE, {})
    if not status or not status.get("run_id"):
        raise ValueError("pipeline_status.json is missing run_id")

    run_id = str(status.get("run_id"))
    started_at = status.get("started_at") or ""
    finished_at = status.get("finished_at") or ""
    pipeline_status = _normalize_pipeline_status(status.get("status"))
    run_date = _date_from_iso(started_at or finished_at)
    steps = status.get("steps") or []

    run_record = {
        "run_id": run_id,
        "run_date": run_date,
        "started_at": started_at,
        "finished_at": finished_at,
        "duration_sec": _duration_seconds(started_at, finished_at),
        "status": pipeline_status,
        "step_count": len(steps),
        "passed_steps": sum(1 for s in steps if s.get("status") == "pass"),
        "failed_steps": sum(1 for s in steps if s.get("status") == "fail"),
        "not_reached_steps": sum(1 for s in steps if s.get("status") in {"blocked", "pending"}),
        "failed_at": status.get("failed_at"),
    }

    history = _read_json(RUN_HISTORY_FILE, [])
    by_run = {str(item.get("run_id")): item for item in history if item.get("run_id")}
    existing = by_run.get(run_id)
    if existing:
        if _status_rank(pipeline_status) >= _status_rank(existing.get("status")):
            by_run[run_id] = run_record
    else:
        by_run[run_id] = run_record

    history = sorted(by_run.values(), key=lambda r: (r.get("started_at") or "", r.get("run_id") or ""))
    history = history[-MAX_RUN_HISTORY:]
    _write_json(RUN_HISTORY_FILE, history)

    step_rows = []
    for step in steps:
        step_rows.append(
            {
                "run_id": run_id,
                "run_date": run_date,
                "started_at": started_at,
                "finished_at": finished_at,
                "pipeline_status": pipeline_status,
                "account": step.get("account") or "",
                "script": step.get("script") or "",
                "status": step.get("status") or "",
                "exit_code": step.get("exit_code"),
                "timestamp": step.get("timestamp") or "",
                "error": step.get("error") or "",
            }
        )
    return history, step_rows


def write_step_history(run_history: list[dict], current_step_rows: list[dict]) -> list[dict]:
    rows = []
    if STEP_HISTORY_FILE.exists():
        try:
            with STEP_HISTORY_FILE.open("r", encoding="utf-8", newline="") as handle:
                rows = list(csv.DictReader(handle))
        except Exception:
            rows = []

    current_run_id = current_step_rows[0]["run_id"] if current_step_rows else ""
    rows = [row for row in rows if row.get("run_id") != current_run_id]
    rows.extend(current_step_rows)

    allowed_run_ids = {str(item.get("run_id")) for item in run_history}
    rows = [row for row in rows if row.get("run_id") in allowed_run_ids]
    rows.sort(key=lambda r: (r.get("started_at") or "", r.get("account") or ""))

    with STEP_HISTORY_FILE.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=STEP_FIELDS)
        writer.writeheader()
        writer.writerows(rows)
    return rows


def _incident_counts_by_day() -> dict[str, Counter]:
    counters: dict[str, Counter] = defaultdict(Counter)
    for item in _read_json(LOG_FILE, []):
        run_id = str(item.get("run_id") or "")
        day = _date_from_iso(run_id) or _date_from_iso(item.get("date"))
        if not day:
            continue
        status = str(item.get("status") or "")
        counters[day]["incident_count"] += 1
        if status == "count_drop":
            counters[day]["count_drop_count"] += 1
        elif status == "healed":
            counters[day]["self_healed_count"] += 1
        elif status == "high_volume":
            counters[day]["high_volume_count"] += 1
    return counters


def build_daily_uptime(run_history: list[dict], step_rows: list[dict]) -> list[dict]:
    runs_by_day: dict[str, list[dict]] = defaultdict(list)
    steps_by_day: dict[str, list[dict]] = defaultdict(list)

    for run in run_history:
        day = str(run.get("run_date") or "")
        if day:
            runs_by_day[day].append(run)

    for step in step_rows:
        day = str(step.get("run_date") or "")
        if day:
            steps_by_day[day].append(step)

    incident_counts = _incident_counts_by_day()
    days = sorted(set(runs_by_day) | set(steps_by_day) | set(incident_counts))
    output = []

    for day in days:
        runs = runs_by_day.get(day, [])
        steps = steps_by_day.get(day, [])
        successful_runs = sum(1 for r in runs if r.get("status") == "success")
        partial_runs = sum(1 for r in runs if r.get("status") == "partial")
        failed_runs = sum(1 for r in runs if r.get("status") in {"failed", "fail"})
        passed_steps = sum(1 for s in steps if s.get("status") == "pass")
        failed_steps = sum(1 for s in steps if s.get("status") == "fail")

        run_count = len(runs)
        step_count = len(steps)
        uptime_pct = round(((successful_runs + partial_runs) / run_count * 100), 2) if run_count else None
        step_uptime_pct = round((passed_steps / step_count * 100), 2) if step_count else None
        incidents = incident_counts.get(day, Counter())

        output.append(
            {
                "date": day,
                "runs": run_count,
                "successful_runs": successful_runs,
                "partial_runs": partial_runs,
                "failed_runs": failed_runs,
                "uptime_pct": uptime_pct,
                "steps": step_count,
                "passed_steps": passed_steps,
                "failed_steps": failed_steps,
                "step_uptime_pct": step_uptime_pct,
                "incident_count": incidents.get("incident_count", 0),
                "count_drop_count": incidents.get("count_drop_count", 0),
                "self_healed_count": incidents.get("self_healed_count", 0),
                "high_volume_count": incidents.get("high_volume_count", 0),
            }
        )

    with DAILY_UPTIME_CSV.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=DAILY_FIELDS)
        writer.writeheader()
        writer.writerows(output)

    _write_json(DAILY_UPTIME_JSON, output)
    return output


def run() -> None:
    run_history, current_step_rows = append_current_run()
    step_rows = write_step_history(run_history, current_step_rows)
    daily = build_daily_uptime(run_history, step_rows)
    print(
        "[status_history] Updated "
        f"{RUN_HISTORY_FILE.name}, {STEP_HISTORY_FILE.name}, "
        f"{DAILY_UPTIME_CSV.name} ({len(daily)} day(s))."
    )


if __name__ == "__main__":
    run()

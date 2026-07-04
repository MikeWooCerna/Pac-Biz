"""Pipeline Guardian: conservative health checks for the Pac-Biz monitor.

Default mode is report-only. Use --fix to regenerate pipeline_monitor.html when
the finished pipeline status and generated monitor disagree. Use --push with
--fix to commit and push only approved monitor artifacts.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent
STATUS_FILE = BASE / "pipeline_status.json"
MONITOR_FILE = BASE / "pipeline_monitor.html"
GENERATE_MONITOR = BASE / "generate_monitor.py"
LOG_FILE = BASE / "pipeline_log.json"

LIVE_MONITOR_URL = "https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html"

APPROVED_PUSH_FILES = {
    "pipeline_status.json",
    "pipeline_monitor.html",
    "pipeline_log.json",
    "pipeline_rowcount_baseline.json",
    "pipeline_highvol_notified.json",
    "pipeline_drops_notified.json",
    "pipeline_heal_events.json",
}

IGNORABLE_UNTRACKED = {
    "history_cache.csv",
    "masterlist_cache.csv",
    "movement_cache.csv",
    "movement_notified.json",
    "step_err.tmp",
}

SPECIAL_STEPS = {"Build", "Git Push"}


@dataclass
class MonitorStats:
    passed: int | None = None
    failed: int | None = None
    not_reached: int | None = None
    rows_pulled: int | None = None
    run_label: str | None = None
    built_label: str | None = None


def run_cmd(args: list[str], check: bool = False, timeout: int = 120) -> subprocess.CompletedProcess:
    return subprocess.run(
        args,
        cwd=str(BASE),
        text=True,
        capture_output=True,
        timeout=timeout,
        check=check,
    )


def load_status() -> dict:
    if not STATUS_FILE.exists():
        return {}
    return json.loads(STATUS_FILE.read_text(encoding="utf-8"))


def short_run_id(run_id: str | None) -> str | None:
    if not run_id:
        return None
    return str(run_id)[:16]


def parse_int(text: str) -> int | None:
    cleaned = re.sub(r"[^0-9]", "", text or "")
    return int(cleaned) if cleaned else None


def parse_monitor(html: str) -> MonitorStats:
    stats = MonitorStats()

    for match in re.finditer(
        r'<div class="sc-val" style="color:[^"]+;">([^<]+)</div><div class="sc-lbl">([^<]+)</div>',
        html,
        flags=re.IGNORECASE,
    ):
        value = parse_int(match.group(1))
        label = match.group(2).strip().lower()
        if label == "passed":
            stats.passed = value
        elif label == "failed":
            stats.failed = value
        elif label == "not reached":
            stats.not_reached = value
        elif label == "rows pulled":
            stats.rows_pulled = value

    run_match = re.search(r">Run\s+([^<]+)</div>", html)
    if run_match:
        stats.run_label = run_match.group(1).strip()

    built_match = re.search(r">Built:\s+([^<]+)</div>", html)
    if built_match:
        stats.built_label = built_match.group(1).strip()

    return stats


def local_monitor_stats() -> MonitorStats:
    if not MONITOR_FILE.exists():
        return MonitorStats()
    return parse_monitor(MONITOR_FILE.read_text(encoding="utf-8", errors="replace"))


def live_monitor_stats() -> tuple[MonitorStats | None, str | None]:
    try:
        with urllib.request.urlopen(LIVE_MONITOR_URL + "?guardian=1", timeout=30) as response:
            html = response.read().decode("utf-8", errors="replace")
        return parse_monitor(html), None
    except Exception as exc:  # network checks should not block local health
        return None, str(exc)


def expected_counts(status: dict) -> tuple[int, int, int]:
    steps = [
        step
        for step in (status.get("steps") or [])
        if step.get("account") not in SPECIAL_STEPS
    ]
    passed = sum(1 for step in steps if step.get("status") == "pass")
    failed = sum(1 for step in steps if step.get("status") == "fail")
    if status.get("status") == "success":
        not_reached = 0
    else:
        # If a run failed before all expected steps were logged, the monitor may
        # show downstream steps as blocked. Keep this as a report-only signal.
        not_reached = max(0, 22 - len(steps))
    return passed, failed, not_reached


def find_monitor_mismatch(status: dict, monitor: MonitorStats) -> list[str]:
    issues: list[str] = []
    expected_passed, expected_failed, expected_not_reached = expected_counts(status)
    status_run = short_run_id(status.get("run_id"))

    if monitor.passed != expected_passed:
        issues.append(f"passed mismatch: status={expected_passed}, monitor={monitor.passed}")
    if monitor.failed != expected_failed:
        issues.append(f"failed mismatch: status={expected_failed}, monitor={monitor.failed}")
    if monitor.not_reached != expected_not_reached:
        issues.append(
            f"not reached mismatch: status={expected_not_reached}, monitor={monitor.not_reached}"
        )
    if status_run and monitor.run_label and status_run not in monitor.run_label:
        issues.append(f"run mismatch: status={status_run}, monitor={monitor.run_label}")
    return issues


def git_porcelain() -> list[str]:
    result = run_cmd(["git", "status", "--porcelain=v1"], timeout=30)
    if result.returncode != 0:
        return [f"!! git status failed: {result.stderr.strip() or result.stdout.strip()}"]
    return [line for line in result.stdout.splitlines() if line.strip()]


def parse_porcelain_path(line: str) -> str:
    path = line[3:]
    if " -> " in path:
        path = path.split(" -> ", 1)[1]
    return path.strip().strip('"')


def dirty_summary() -> tuple[list[str], list[str]]:
    blocker_lines: list[str] = []
    ignorable_lines: list[str] = []
    for line in git_porcelain():
        path = parse_porcelain_path(line)
        code = line[:2]
        if code == "??" and path in IGNORABLE_UNTRACKED:
            ignorable_lines.append(line)
        elif path not in APPROVED_PUSH_FILES:
            blocker_lines.append(line)
    return blocker_lines, ignorable_lines


def regenerate_monitor() -> bool:
    result = run_cmd([sys.executable, str(GENERATE_MONITOR)], timeout=180)
    if result.stdout.strip():
        print(result.stdout.strip())
    if result.stderr.strip():
        print(result.stderr.strip(), file=sys.stderr)
    return result.returncode == 0


def approved_monitor_changes() -> list[str]:
    changed: list[str] = []
    for line in git_porcelain():
        path = parse_porcelain_path(line)
        if path in APPROVED_PUSH_FILES:
            changed.append(path)
    return sorted(set(changed))


def commit_and_push(message: str) -> bool:
    blockers, _ = dirty_summary()
    if blockers:
        print("[guardian] Push blocked because non-monitor files are dirty:")
        for line in blockers:
            print(f"  {line}")
        return False

    changed = approved_monitor_changes()
    if not changed:
        print("[guardian] No approved monitor changes to commit.")
        return True

    add_result = run_cmd(["git", "add", *changed], timeout=30)
    if add_result.returncode != 0:
        print(add_result.stderr.strip() or add_result.stdout.strip(), file=sys.stderr)
        return False

    diff_result = run_cmd(["git", "diff", "--cached", "--quiet"], timeout=30)
    if diff_result.returncode == 0:
        print("[guardian] No staged monitor changes.")
        return True

    commit_result = run_cmd(["git", "commit", "-m", message], timeout=60)
    if commit_result.returncode != 0:
        print(commit_result.stderr.strip() or commit_result.stdout.strip(), file=sys.stderr)
        return False
    print(commit_result.stdout.strip())

    push_result = run_cmd(["git", "push"], timeout=120)
    if push_result.returncode != 0:
        print(push_result.stderr.strip() or push_result.stdout.strip(), file=sys.stderr)
        return False
    print(push_result.stdout.strip() or push_result.stderr.strip())
    return True


def print_stats(label: str, stats: MonitorStats) -> None:
    print(
        f"[guardian] {label}: "
        f"passed={stats.passed}, failed={stats.failed}, "
        f"not_reached={stats.not_reached}, rows={stats.rows_pulled}, "
        f"run={stats.run_label or '-'}, built={stats.built_label or '-'}"
    )


def log_guardian_event(issues: list[str], fixed: bool, pushed: bool) -> None:
    """Append a Guardian run to pipeline_log.json so the monitor's incident log
    can show when the guardian fixed or flagged a mismatch.

    Only called when a mismatch was actually detected — clean runs are never
    logged. `status` is "guardian_fix" when a --fix attempt resolved the
    mismatch, otherwise "guardian_warn" (report-only mode, or a --fix attempt
    that did not fully resolve the mismatch).
    """
    if not issues:
        return

    now = datetime.now()
    status = "guardian_fix" if fixed else "guardian_warn"
    entry = {
        "run_id": now.strftime("%Y-%m-%dT%H:%M:%S"),
        "date": now.strftime("%b %d, %Y %I:%M %p"),
        "account": "Guardian",
        "script": "pipeline_guardian.py",
        "status": status,
        "error": "; ".join(issues),
    }

    try:
        entries = json.loads(LOG_FILE.read_text(encoding="utf-8")) if LOG_FILE.exists() else []
    except Exception:
        entries = []
    if not isinstance(entries, list):
        entries = []

    entries.append(entry)

    try:
        LOG_FILE.write_text(json.dumps(entries[-300:], indent=2, default=str), encoding="utf-8")
    except Exception:
        pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Check and optionally repair the pipeline monitor.")
    parser.add_argument("--fix", action="store_true", help="Regenerate the monitor if it is stale.")
    parser.add_argument("--push", action="store_true", help="Commit and push approved monitor files after --fix.")
    parser.add_argument("--live", action="store_true", help="Also compare the live GitHub Pages monitor.")
    args = parser.parse_args()

    status = load_status()
    if not status:
        print("[guardian] ERROR: pipeline_status.json is missing or unreadable.")
        return 2

    print(
        f"[guardian] status: run={short_run_id(status.get('run_id'))}, "
        f"state={status.get('status')}, steps={len(status.get('steps') or [])}"
    )

    local_stats = local_monitor_stats()
    print_stats("local monitor", local_stats)
    issues = find_monitor_mismatch(status, local_stats)

    blockers, ignorable = dirty_summary()
    if ignorable:
        print("[guardian] Ignorable generated files present:")
        for line in ignorable:
            print(f"  {line}")
    if blockers:
        print("[guardian] Non-monitor local changes present:")
        for line in blockers:
            print(f"  {line}")

    # Preserve the originally-detected mismatch for the guardian log entry —
    # `issues` itself gets reassigned below once a --fix attempt is made.
    original_issues = list(issues)
    fixed = False

    if issues:
        print("[guardian] Local monitor mismatch detected:")
        for issue in issues:
            print(f"  - {issue}")
        if args.fix:
            print("[guardian] Regenerating monitor...")
            if not regenerate_monitor():
                print("[guardian] ERROR: monitor regeneration failed.")
                log_guardian_event(original_issues, fixed=False, pushed=False)
                return 1
            local_stats = local_monitor_stats()
            print_stats("local monitor after fix", local_stats)
            issues = find_monitor_mismatch(status, local_stats)
            if issues:
                print("[guardian] ERROR: monitor still mismatches after regeneration:")
                for issue in issues:
                    print(f"  - {issue}")
                log_guardian_event(original_issues, fixed=False, pushed=False)
                return 1
            fixed = True
        else:
            print("[guardian] Run with --fix to regenerate pipeline_monitor.html.")
    else:
        print("[guardian] Local monitor matches pipeline_status.json.")

    if args.push:
        if not args.fix:
            print("[guardian] ERROR: --push requires --fix.")
            return 2
        pushed = commit_and_push("Guardian: refresh pipeline monitor")
        if original_issues:
            log_guardian_event(original_issues, fixed=fixed, pushed=pushed)
        if not pushed:
            return 1
    elif original_issues:
        log_guardian_event(original_issues, fixed=fixed, pushed=False)

    if args.live:
        live_stats, live_error = live_monitor_stats()
        if live_error:
            print(f"[guardian] Live monitor check skipped: {live_error}")
        elif live_stats:
            print_stats("live monitor", live_stats)
            live_issues = find_monitor_mismatch(status, live_stats)
            if live_issues:
                print("[guardian] Live monitor mismatch detected:")
                for issue in live_issues:
                    print(f"  - {issue}")
                print("[guardian] This is often a GitHub Pages deployment/cache delay.")
            else:
                print("[guardian] Live monitor matches pipeline_status.json.")

    return 0 if not issues else 1


if __name__ == "__main__":
    raise SystemExit(main())

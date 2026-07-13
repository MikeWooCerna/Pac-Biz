"""movement_reconcile.py — reconcile processed movements after pipeline runs.

This is a safety net for the gap where Movement rows are already processed, but
the notification tracker or generated dashboard HTML has not caught up yet.

Source of truth:
  - movement_cache.csv
  - masterlist_cache.csv
  - history_cache.csv

The script is intentionally conservative:
  - It delegates email sending to check_movement_notifications.py.
  - It only patches generated dashboard data from local cache files.
  - It does not edit source Google Sheets or movement processing fields.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import pandas as pd

import check_movement_notifications as movement_notify


BASE_DIR = Path(__file__).parent
MASTERLIST_CACHE = BASE_DIR / "masterlist_cache.csv"
HISTORY_CACHE = BASE_DIR / "history_cache.csv"
MOVEMENT_CACHE = BASE_DIR / "movement_cache.csv"
DASHBOARD_HTML = BASE_DIR / "masterlist_dashboard.html"


def _read_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing required cache: {path.name}")
    return pd.read_csv(path, dtype=str).fillna("")


def _prepare_movement(movement: pd.DataFrame, masterlist: pd.DataFrame) -> pd.DataFrame:
    movement = movement.copy()

    if "Type of Movement" in movement.columns and "Movement Type" in movement.columns:
        movement["Movement Type"] = movement.apply(
            lambda r: (
                r["Movement Type"]
                if str(r.get("Movement Type", "")).strip()
                else str(r.get("Type of Movement", "")).strip()
            ),
            axis=1,
        )

    if "Company Email" in masterlist.columns and "Emp Name" in masterlist.columns:
        email_to_name = (
            masterlist.assign(_email=masterlist["Company Email"].str.strip().str.lower())
            .loc[lambda d: d["_email"].ne("")]
            .drop_duplicates("_email")
            .set_index("_email")["Emp Name"]
            .to_dict()
        )
    else:
        email_to_name = {}

    if "Email Address" in movement.columns:
        movement["Initiated by"] = movement["Email Address"].apply(
            lambda e: email_to_name.get(str(e).strip().lower(), "") if str(e).strip() else ""
        )
    else:
        movement["Initiated by"] = ""

    return movement


def _eligible_processed_rows(movement: pd.DataFrame) -> list[pd.Series]:
    return [
        row
        for _, row in movement.iterrows()
        if movement_notify.is_ready_to_notify(row)
    ]


def notify_missing_processed_movements(movement: pd.DataFrame) -> int:
    seen = movement_notify.load_notified()
    if seen is None:
        # Preserve the existing seeding behavior in check_movement_notifications.
        print("[movement_reconcile] movement_notified.json missing; delegating seed behavior.")
        return movement_notify.main()

    missing = [
        str(row.get("Timestamp", "")).strip()
        for row in _eligible_processed_rows(movement)
        if str(row.get("Timestamp", "")).strip() not in seen
    ]

    if not missing:
        print("[movement_reconcile] Notification ledger is current.")
        return 0

    print(
        "[movement_reconcile] Found "
        f"{len(missing)} processed movement(s) missing notification marker."
    )
    return movement_notify.main()


def _replace_const(text: str, name: str, value: object) -> str:
    prefix = f"const {name} = "
    start = text.index(prefix)
    value_start = start + len(prefix)
    end = text.index(";\nconst ", value_start)
    return text[:value_start] + json.dumps(value, ensure_ascii=False) + text[end:]


def patch_dashboard_from_cache(
    masterlist: pd.DataFrame,
    history: pd.DataFrame,
    movement: pd.DataFrame,
) -> bool:
    if not DASHBOARD_HTML.exists():
        print("[movement_reconcile] masterlist_dashboard.html missing; dashboard patch skipped.")
        return False

    original = DASHBOARD_HTML.read_text(encoding="utf-8")
    text = original

    text = _replace_const(text, "masterlist", masterlist.to_dict(orient="records"))
    text = _replace_const(text, "historyData", history.to_dict(orient="records"))
    text = _replace_const(text, "movementData", movement.to_dict(orient="records"))

    non_void = movement[
        ~movement.get("Void", pd.Series([""] * len(movement)))
        .str.strip()
        .str.upper()
        .eq("YES")
    ]
    for_processing = non_void[
        ~non_void.get("Processed", pd.Series([""] * len(non_void)))
        .str.strip()
        .str.lower()
        .eq("yes")
    ]

    kpi_match = re.search(r"const masterlistKpis = (.*?);", text)
    if not kpi_match:
        raise ValueError("masterlistKpis block not found in dashboard HTML")

    kpis = json.loads(kpi_match.group(1))
    kpis["movementsPending"] = int(len(non_void))
    kpis["movementsForProcessing"] = int(len(for_processing))
    kpis["historyRecords"] = int(len(history))
    text = (
        text[: kpi_match.start()]
        + "const masterlistKpis = "
        + json.dumps(kpis, ensure_ascii=False)
        + ";"
        + text[kpi_match.end() :]
    )

    if text == original:
        print("[movement_reconcile] Dashboard movement snapshot already current.")
        return False

    DASHBOARD_HTML.write_text(text, encoding="utf-8")
    print("[movement_reconcile] Dashboard movement snapshot refreshed from cache.")
    return True


def verify_processed_movements(masterlist: pd.DataFrame, history: pd.DataFrame, movement: pd.DataFrame) -> int:
    issues = 0
    master_by_name = {
        str(row.get("Emp Name", "")).strip().lower(): row
        for _, row in masterlist.iterrows()
        if str(row.get("Emp Name", "")).strip()
    }

    for row in _eligible_processed_rows(movement):
        emp_name = str(row.get("Employee Name", "")).strip()
        mov_type = str(row.get("Type of Movement", "")).strip().lower()
        effective = str(row.get("Effective Date", "")).strip()
        master_row = master_by_name.get(emp_name.lower())

        if master_row is None:
            print(f"[movement_reconcile][WARN] Processed movement employee not in masterlist: {emp_name}")
            issues += 1
            continue

        if mov_type == "attrition":
            status = str(master_row.get("Employment Status", "")).strip().lower()
            if status != "inactive":
                print(
                    "[movement_reconcile][WARN] Attrition not reflected as Inactive "
                    f"in masterlist cache: {emp_name} ({status or 'blank'})"
                )
                issues += 1

        if effective and {"Emp Name", "Date Generated", "Employment Status"}.issubset(history.columns):
            emp_history = history[history["Emp Name"].str.strip().str.lower().eq(emp_name.lower())]
            if emp_history.empty:
                print(f"[movement_reconcile][WARN] No history rows found for processed movement: {emp_name}")
                issues += 1

    if issues == 0:
        print("[movement_reconcile] Processed movement cache verification passed.")
    return issues


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--no-dashboard-patch",
        action="store_true",
        help="Only reconcile notifications and verify cache state.",
    )
    args = parser.parse_args(argv)

    try:
        masterlist = _read_csv(MASTERLIST_CACHE)
        history = _read_csv(HISTORY_CACHE)
        movement = _prepare_movement(_read_csv(MOVEMENT_CACHE), masterlist)
    except Exception as exc:
        print(f"[movement_reconcile][ERROR] {exc}")
        return 1

    notify_rc = notify_missing_processed_movements(movement)
    if notify_rc != 0:
        return notify_rc

    issues = verify_processed_movements(masterlist, history, movement)

    if not args.no_dashboard_patch:
        try:
            patch_dashboard_from_cache(masterlist, history, movement)
        except Exception as exc:
            print(f"[movement_reconcile][ERROR] Dashboard patch failed: {exc}")
            return 1

    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())

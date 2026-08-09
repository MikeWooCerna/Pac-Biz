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


def _text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


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


def _parse_date(value: object) -> pd.Timestamp | None:
    text = _text(value)
    if not text:
        return None
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.normalize()


def _date_series(values: pd.Series) -> pd.Series:
    return pd.to_datetime(values, errors="coerce").dt.normalize()


def _history_start_date(history: pd.DataFrame) -> pd.Timestamp | None:
    if "Date Generated" not in history.columns or history.empty:
        return None
    dates = _date_series(history["Date Generated"]).dropna()
    if dates.empty:
        return None
    return dates.min()


def _is_yes(value: object) -> bool:
    return _text(value).lower() == "yes"


def _change_type_from_processed_note(value: object) -> str:
    note = _text(value)
    if not note:
        return "Change of Employment Status"
    change_type = note.split("|", 1)[0].strip()
    return change_type or "Change of Employment Status"


def _week_start_text(date: pd.Timestamp) -> str:
    week_start = date - pd.Timedelta(days=int(date.weekday()))
    return week_start.strftime("%m/%d/%Y")


def repair_missing_effective_history_rows(
    masterlist: pd.DataFrame,
    history: pd.DataFrame,
    movement: pd.DataFrame,
) -> pd.DataFrame:
    """Append local History cache rows for processed movements missing their effective date.

    Apps Script is still the source-of-truth writer for Google Sheets. This local repair keeps
    the generated dashboard and pipeline validation coherent after a fetch when Apps Script
    marked a Movement row processed but the daily active History snapshot skipped the employee.
    """
    required_history = {"Emp Name", "Date Generated", "Employment Status"}
    if not required_history.issubset(history.columns):
        return history

    history_start = _history_start_date(history)
    if history_start is None:
        return history

    master_by_name = {
        str(row.get("Emp Name", "")).strip().lower(): row
        for _, row in masterlist.iterrows()
        if str(row.get("Emp Name", "")).strip()
    }

    repaired_rows: list[dict[str, object]] = []
    existing_keys = set()
    history_dates = _date_series(history["Date Generated"])
    for idx, row in history.iterrows():
        emp_name = _text(row.get("Emp Name")).lower()
        date = history_dates.loc[idx]
        if emp_name and not pd.isna(date):
            existing_keys.add((emp_name, date))

    for row in _eligible_processed_rows(movement):
        emp_name = _text(row.get("Employee Name"))
        effective_date = _parse_date(row.get("Effective Date"))
        if not emp_name or effective_date is None or effective_date < history_start:
            continue

        key = (emp_name.lower(), effective_date)
        if key in existing_keys:
            continue

        master_row = master_by_name.get(emp_name.lower())
        if master_row is None:
            continue

        repaired = {}
        for header in history.columns:
            if header == "Date Generated":
                repaired[header] = effective_date.strftime("%m/%d/%Y")
            elif header == "Change Type":
                repaired[header] = _change_type_from_processed_note(row.get("Processed Note"))
            elif header == "Week":
                repaired[header] = _week_start_text(effective_date)
            else:
                repaired[header] = master_row.get(header, "")

        repaired_rows.append(repaired)
        existing_keys.add(key)
        print(
            "[movement_reconcile] Repaired local History effective-date row: "
            f"{emp_name} ({effective_date.strftime('%m/%d/%Y')})"
        )

    if not repaired_rows:
        return history

    repaired_history = pd.concat(
        [history, pd.DataFrame(repaired_rows, columns=history.columns)],
        ignore_index=True,
    )
    repaired_history.to_csv(HISTORY_CACHE, index=False)
    print(f"[movement_reconcile] Added {len(repaired_rows)} local History repair row(s).")
    return repaired_history


def verify_no_overdue_unprocessed_movements(movement: pd.DataFrame) -> int:
    """Fail loudly when Apps Script left due Movement rows unprocessed.

    Movement emails intentionally require Processed = Yes. If an overdue row is
    still unprocessed, the email sender must not fire yet; this guard makes the
    pipeline surface the real upstream issue instead of passing silently.
    """
    today = pd.Timestamp.today().normalize()
    issues = 0

    for idx, row in movement.iterrows():
        if _is_yes(row.get("Void")) or _is_yes(row.get("Processed")):
            continue

        effective = _parse_date(row.get("Effective Date"))
        if effective is None or effective > today:
            continue

        emp_name = _text(row.get("Employee Name")) or "(blank employee)"
        mov_type = _text(row.get("Type of Movement")) or "(blank movement type)"
        timestamp = _text(row.get("Timestamp")) or "(blank timestamp)"
        note = _text(row.get("Processed Note")) or "no processed note"
        print(
            "[movement_reconcile][WARN] Overdue movement still unprocessed: "
            f"row {idx + 2}, {emp_name}, {mov_type}, effective "
            f"{effective.strftime('%m/%d/%Y')}, timestamp {timestamp}, {note}"
        )
        issues += 1

    if issues == 0:
        print("[movement_reconcile] No overdue unprocessed movements found.")
    return issues


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
    history_start = _history_start_date(history)
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
                continue

            effective_date = _parse_date(effective)
            if effective_date is not None:
                if history_start is not None and effective_date < history_start:
                    continue
                effective_history = emp_history[
                    _date_series(emp_history["Date Generated"]).eq(effective_date)
                ]
                if effective_history.empty:
                    print(
                        "[movement_reconcile][WARN] No effective-date history row "
                        f"for processed movement: {emp_name} ({effective_date.strftime('%m/%d/%Y')})"
                    )
                    issues += 1
                elif mov_type == "attrition":
                    inactive_rows = effective_history[
                        effective_history["Employment Status"].str.strip().str.lower().eq("inactive")
                    ]
                    if inactive_rows.empty:
                        print(
                            "[movement_reconcile][WARN] Attrition effective-date history row "
                            f"is not Inactive: {emp_name} ({effective_date.strftime('%m/%d/%Y')})"
                        )
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

    history = repair_missing_effective_history_rows(masterlist, history, movement)

    issues = verify_no_overdue_unprocessed_movements(movement)
    issues += verify_processed_movements(masterlist, history, movement)

    if not args.no_dashboard_patch:
        try:
            patch_dashboard_from_cache(masterlist, history, movement)
        except Exception as exc:
            print(f"[movement_reconcile][ERROR] Dashboard patch failed: {exc}")
            return 1

    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())

"""Fetches Masterlist Google Sheets CSVs and caches them locally for dashboard.py."""
import sys, requests
from pathlib import Path

BASE = Path(__file__).parent

MASTERLIST_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=0&single=true&output=csv"
HISTORY_CSV    = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=166777136&single=true&output=csv"
MOVEMENT_CSV   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=693236738&single=true&output=csv"

SHEETS = [
    ("Masterlist",  MASTERLIST_CSV, BASE / "masterlist_cache.csv"),
    ("History",     HISTORY_CSV,    BASE / "history_cache.csv"),
    ("Movement",    MOVEMENT_CSV,   BASE / "movement_cache.csv"),
]

for label, url, dest in SHEETS:
    print(f"Fetching {label} sheet ...")
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        dest.write_bytes(r.content)
        rows = max(0, r.text.count('\n') - 1)
        print(f"  -> {rows:,} rows cached to {dest.name}")
    except Exception as e:
        print(f"  ERROR fetching {label}: {e}", file=sys.stderr)
        sys.exit(1)

print("Masterlist fetch complete.")

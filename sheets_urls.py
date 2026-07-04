# Single source of truth for Masterlist Google Sheets published CSV URLs.
# When the sheet is re-published (URL rotation), update only _BASE here.
_BASE = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub"

MASTERLIST_CSV = f"{_BASE}?gid=0&single=true&output=csv"
HISTORY_CSV    = f"{_BASE}?gid=166777136&single=true&output=csv"
MOVEMENT_CSV   = f"{_BASE}?gid=693236738&single=true&output=csv"

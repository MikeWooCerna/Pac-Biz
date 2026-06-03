import pandas as pd

MASTERLIST_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=0&single=true&output=csv"

HISTORY_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=166777136&single=true&output=csv"

MOVEMENT_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=693236738&single=true&output=csv"

masterlist = pd.read_csv(MASTERLIST_CSV)
history = pd.read_csv(HISTORY_CSV)
movement = pd.read_csv(MOVEMENT_CSV)

print("=" * 50)
print("MASTERLIST")
print(masterlist.shape)

print("=" * 50)
print("HISTORY")
print(history.shape)

print("=" * 50)
print("MOVEMENT")
print(movement.shape)
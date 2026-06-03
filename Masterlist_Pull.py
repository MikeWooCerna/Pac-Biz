import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import datetime

INPUT_FILE = "Masterlist_Final.xlsx"
OUTPUT_FILE = "masterlist_dashboard.html"

MASTERLIST_SHEET = "Masterlist"
MOVEMENT_SHEET = "Movement"
HISTORY_SHEET = "History"


def clean_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_data():
    masterlist = pd.read_excel(INPUT_FILE, sheet_name=MASTERLIST_SHEET)
    movement = pd.read_excel(INPUT_FILE, sheet_name=MOVEMENT_SHEET)
    history = pd.read_excel(INPUT_FILE, sheet_name=HISTORY_SHEET)

    return (
        clean_columns(masterlist),
        clean_columns(movement),
        clean_columns(history),
    )


def safe_count(df, column, value=None):
    if column not in df.columns:
        return 0

    if value is None:
        return df[column].notna().sum()

    return df[df[column].astype(str).str.upper().str.strip() == value.upper()].shape[0]


def build_chart_html(fig):
    return fig.to_html(full_html=False, include_plotlyjs=False)


def main():
    masterlist, movement, history = load_data()

    today = datetime.now().strftime("%Y-%m-%d %I:%M %p")

    active_count = safe_count(masterlist, "Employment Status", "Active")
    inactive_count = safe_count(masterlist, "Employment Status", "Inactive")
    total_employees = masterlist.shape[0]

    movement_count = movement.shape[0]
    history_count = history.shape[0]

    latest_snapshot = ""
    if "Date Generated" in history.columns:
        latest_snapshot = pd.to_datetime(history["Date Generated"], errors="coerce").max()
        latest_snapshot = latest_snapshot.strftime("%m/%d/%Y") if pd.notna(latest_snapshot) else ""

    charts = []

    if "Department" in masterlist.columns:
        dept_counts = (
            masterlist["Department"]
            .fillna("Blank")
            .value_counts()
            .reset_index()
        )
        dept_counts.columns = ["Department", "Count"]

        fig = px.bar(
            dept_counts,
            x="Department",
            y="Count",
            title="Headcount by Department",
            text="Count",
        )
        charts.append(build_chart_html(fig))

    if "LOB / Account" in masterlist.columns:
        lob_counts = (
            masterlist["LOB / Account"]
            .fillna("Blank")
            .value_counts()
            .head(15)
            .reset_index()
        )
        lob_counts.columns = ["LOB / Account", "Count"]

        fig = px.bar(
            lob_counts,
            x="Count",
            y="LOB / Account",
            orientation="h",
            title="Top LOB / Account by Headcount",
            text="Count",
        )
        charts.append(build_chart_html(fig))

    if "Change Type" in history.columns:
        change_counts = (
            history["Change Type"]
            .fillna("Blank")
            .value_counts()
            .reset_index()
        )
        change_counts.columns = ["Change Type", "Count"]

        fig = px.pie(
            change_counts,
            names="Change Type",
            values="Count",
            title="Change Type Distribution",
            hole=0.45,
        )
        charts.append(build_chart_html(fig))

    if "Week" in history.columns:
        weekly = history.copy()
        weekly["Week"] = pd.to_datetime(weekly["Week"], errors="coerce")

        weekly_counts = (
            weekly.dropna(subset=["Week"])
            .groupby("Week")
            .size()
            .reset_index(name="Records")
        )

        fig = px.line(
            weekly_counts,
            x="Week",
            y="Records",
            title="Weekly Snapshot Records",
            markers=True,
        )
        charts.append(build_chart_html(fig))

    html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Masterlist Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
    <style>
        body {{
            font-family: Arial, sans-serif;
            background: #f4f6f8;
            margin: 0;
            color: #1f2937;
        }}

        .header {{
            background: linear-gradient(90deg, #4b0082, #6a00ff);
            color: white;
            padding: 24px 36px;
        }}

        .header h1 {{
            margin: 0;
            font-size: 28px;
        }}

        .header p {{
            margin: 6px 0 0;
            opacity: 0.9;
        }}

        .cards {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 16px;
            padding: 24px 36px;
        }}

        .card {{
            background: white;
            border-radius: 14px;
            padding: 18px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }}

        .card .label {{
            font-size: 12px;
            color: #64748b;
            text-transform: uppercase;
            font-weight: bold;
        }}

        .card .value {{
            font-size: 28px;
            font-weight: bold;
            margin-top: 8px;
        }}

        .grid {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 18px;
            padding: 0 36px 36px;
        }}

        .chart {{
            background: white;
            border-radius: 14px;
            padding: 14px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }}

        @media (max-width: 1200px) {{
            .cards {{
                grid-template-columns: repeat(2, 1fr);
            }}

            .grid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>

<div class="header">
    <h1>Masterlist Dashboard — v1.0</h1>
    <p>Refresh Time: {today} | Latest Snapshot: {latest_snapshot}</p>
</div>

<div class="cards">
    <div class="card">
        <div class="label">Total Employees</div>
        <div class="value">{total_employees:,}</div>
    </div>

    <div class="card">
        <div class="label">Active Employees</div>
        <div class="value">{active_count:,}</div>
    </div>

    <div class="card">
        <div class="label">Inactive Employees</div>
        <div class="value">{inactive_count:,}</div>
    </div>

    <div class="card">
        <div class="label">Movement Records</div>
        <div class="value">{movement_count:,}</div>
    </div>

    <div class="card">
        <div class="label">History Records</div>
        <div class="value">{history_count:,}</div>
    </div>
</div>

<div class="grid">
    {''.join(f'<div class="chart">{chart}</div>' for chart in charts)}
</div>

</body>
</html>
"""

    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")
    print(f"Dashboard generated: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
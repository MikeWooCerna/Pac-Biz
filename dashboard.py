import base64
import json
from pathlib import Path
from datetime import datetime

import pandas as pd

MASTERLIST_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=0&single=true&output=csv"
HISTORY_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=166777136&single=true&output=csv"
MOVEMENT_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=693236738&single=true&output=csv"

OUTPUT_FILE = "masterlist_dashboard.html"
LOGO_FILE = "pacbiz_logo.png"

PACBIZ_BLUE = "#004C97"
PACBIZ_GREEN = "#39B54A"


def clean_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df.fillna("")


def get_logo_data_uri():
    logo_path = Path(LOGO_FILE)
    if not logo_path.exists():
        return ""
    encoded = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
    return f"data:image/png;base64,{encoded}"


def to_records(df):
    return json.dumps(df.to_dict(orient="records"), ensure_ascii=False)


def main():
    masterlist = clean_columns(pd.read_csv(MASTERLIST_CSV))
    history = clean_columns(pd.read_csv(HISTORY_CSV))
    movement = clean_columns(pd.read_csv(MOVEMENT_CSV))

    refresh_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")

    latest_snapshot = ""
    if "Date Generated" in history.columns:
        latest_date = pd.to_datetime(history["Date Generated"], errors="coerce").max()
        latest_snapshot = latest_date.strftime("%m/%d/%Y") if pd.notna(latest_date) else ""

    logo_uri = get_logo_data_uri()

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Pac-Biz Dashboard</title>
{"<link rel='icon' type='image/png' href='" + logo_uri + "'>" if logo_uri else ""}
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>

<style>
    :root {{
        --blue: {PACBIZ_BLUE};
        --green: {PACBIZ_GREEN};
        --dark-blue: #002B5C;
        --bg: #F4F8F6;
        --card: #FFFFFF;
        --text: #172033;
        --muted: #64748B;
    }}

    * {{
        box-sizing: border-box;
    }}

    body {{
        margin: 0;
        font-family: Arial, sans-serif;
        background: var(--bg);
        color: var(--text);
    }}

    .topbar {{
        height: 8px;
        background: linear-gradient(90deg, var(--green), var(--blue));
    }}

    .sticky-dashboard-header {{
        position: sticky;
        top: 0;
        z-index: 1000;
        background: var(--bg);
        box-shadow: 0 3px 12px rgba(0,0,0,0.08);
    }}

    .header {{
        background: white;
        padding: 10px 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border-bottom: 3px solid var(--green);
    }}

    .brand {{
        display: flex;
        align-items: center;
        gap: 14px;
    }}

    .logo {{
        width: 95px;
        height: auto;
    }}

    .title h1 {{
        margin: 0;
        color: var(--dark-blue);
        font-size: 24px;
        font-weight: 900;
    }}

    .title p {{
        margin: 4px 0 0;
        color: #334155;
        font-size: 12px;
    }}

    .filters {{
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 12px;
        padding: 12px 18px 8px;
        background: var(--bg);
    }}

    .filter-box {{
        background: white;
        border: 1px solid var(--green);
        border-radius: 10px;
        padding: 8px 12px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.05);
    }}

    .filter-box label {{
        display: block;
        font-size: 11px;
        font-weight: 800;
        color: var(--dark-blue);
        margin-bottom: 4px;
    }}

    .filter-box select {{
        width: 100%;
        border: none;
        outline: none;
        font-size: 13px;
        background: white;
        color: var(--text);
    }}

    .cards {{
        display: grid;
        grid-template-columns: repeat(8, 1fr);
        gap: 10px;
        padding: 8px 18px 14px;
        background: var(--bg);
    }}

    .card {{
        background: white;
        border-radius: 10px;
        padding: 10px 12px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.06);
        border-top: 3px solid var(--blue);
        min-height: 76px;
    }}

    .card:nth-child(even) {{
        border-top-color: var(--green);
    }}

    .card .label {{
        font-size: 10px;
        color: var(--muted);
        text-transform: uppercase;
        font-weight: 800;
        letter-spacing: .04em;
    }}

    .card .value {{
        margin-top: 6px;
        font-size: 24px;
        color: var(--dark-blue);
        font-weight: 900;
    }}

    .grid {{
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 12px;
        padding: 0 18px 20px;
    }}

    .chart-card {{
        background: white;
        border-radius: 12px;
        padding: 10px;
        box-shadow: 0 1px 8px rgba(0,0,0,0.06);
        border-left: 4px solid var(--green);
        min-height: 320px;
    }}

    .bar-chart-row {{
        display: grid;
        grid-column: 1 / -1;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 12px;
    }}

    .donut-chart-row {{
        display: grid;
        grid-column: 1 / -1;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 12px;
    }}

    .table-scroll {{
        max-height: 430px;
        overflow: auto;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
    }}

    .full {{
        grid-column: 1 / -1;
    }}

    table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
    }}

    th {{
        background: var(--blue);
        color: white;
        padding: 8px;
        text-align: left;
        position: sticky;
        top: 0;
        z-index: 1;
    }}

    td {{
        padding: 7px 8px;
        border-bottom: 1px solid #E5E7EB;
    }}

    tr:nth-child(even) td {{
        background: #F8FAFC;
    }}

    .footer {{
        background: linear-gradient(90deg, var(--green), var(--blue));
        color: white;
        padding: 14px 28px;
        text-align: center;
        font-size: 13px;
    }}

    @media (max-width: 1300px) {{
        .cards {{
            grid-template-columns: repeat(4, 1fr);
        }}

        .grid {{
            grid-template-columns: 1fr;
        }}

        .filters {{
            grid-template-columns: repeat(2, 1fr);
        }}

        .bar-chart-row {{
            grid-template-columns: 1fr;
        }}

        .donut-chart-row {{
            grid-template-columns: 1fr;
        }}
    }}
</style>
</head>

<body>
<div class="sticky-dashboard-header">
<div class="topbar"></div>

<div class="header">
    <div class="brand">
        {"<img class='logo' src='" + logo_uri + "'>" if logo_uri else ""}
        <div class="title">
            <h1>Pac-Biz Dashboard</h1>
            <p>Refresh Time: {refresh_time} &nbsp;|&nbsp; Latest Snapshot: {latest_snapshot}</p>
        </div>
    </div>
</div>

<div class="filters">
    <div class="filter-box">
        <label>Department</label>
        <select id="departmentFilter"></select>
    </div>
    <div class="filter-box">
        <label>LOB / Account</label>
        <select id="accountFilter"></select>
    </div>
    <div class="filter-box">
        <label>Manager</label>
        <select id="managerFilter"></select>
    </div>
    <div class="filter-box">
        <label>Supervisor</label>
        <select id="supervisorFilter"></select>
    </div>
</div>

<div class="cards">
    <div class="card"><div class="label">Headcount</div><div class="value" id="headcount">0</div></div>
    <div class="card"><div class="label">Active</div><div class="value" id="active">0</div></div>
    <div class="card"><div class="label">Inactive</div><div class="value" id="inactive">0</div></div>
    <div class="card"><div class="label">Movements</div><div class="value" id="movements">0</div></div>
    <div class="card"><div class="label">History Records</div><div class="value" id="historyRecords">0</div></div>
    <div class="card"><div class="label">Departments</div><div class="value" id="departments">0</div></div>
    <div class="card"><div class="label">Account</div><div class="value" id="accounts">0</div></div>
    <div class="card"><div class="label">Managers</div><div class="value" id="managers">0</div></div>
</div>
</div>

<div class="grid">
    <div class="donut-chart-row">
        <div class="chart-card"><div id="deptDonut"></div></div>
        <div class="chart-card"><div id="activeDonut"></div></div>
        <div class="chart-card"><div id="employeeGroupDonut"></div></div>
    </div>
    <div class="bar-chart-row">
        <div class="chart-card"><div id="accountBar"></div></div>
        <div class="chart-card"><div id="managerBar"></div></div>
        <div class="chart-card"><div id="supervisorBar"></div></div>
    </div>
    <div class="chart-card"><div id="changeTypeDonut"></div></div>
    <div class="chart-card"><div id="weeklyLine"></div></div>
    <div class="chart-card full">
        <h3 style="color:#004C97;margin:4px 0 12px;">Master List</h3>
        <div id="masterlistTable"></div>
    </div>
    <div class="chart-card full">
        <h3 style="color:#004C97;margin:4px 0 12px;">Recent Employee Movements</h3>
        <div id="recentMovements"></div>
    </div>
</div>

<div class="footer">
    Developed for Pac-Biz Reporting MCerna | Data Source: Master List | Automation: Python 3.13.0
</div>

<script>
const masterlist = {to_records(masterlist)};
const historyData = {to_records(history)};
const movementData = {to_records(movement)};

const COLORS = ["#004C97", "#39B54A", "#002B5C", "#7AC943", "#00AEEF", "#94A3B8"];
const MASTERLIST_COLUMNS = [
    {{label: "Employee ID", field: "ID No."}},
    {{label: "Employee Name", field: "Emp Name"}},
    {{label: "Hire Date", field: "Hire Date"}},
    {{label: "Employement Class", field: "Employement Class"}},
    {{label: "Job Title", field: "Job Title"}},
    {{label: "Employee Group", field: "Employee Group"}},
    {{label: "Department", field: "Department"}},
    {{label: "LOB/Account", field: "LOB / Account"}},
    {{label: "Immediate Supervisor", field: "Immediate Supervisor"}},
    {{label: "Manager", field: "Manager"}},
    {{label: "Email", field: "Company Email"}},
];

function norm(v) {{
    return (v ?? "").toString().trim();
}}

function escapeHtml(v) {{
    return norm(v)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}}

function filteredMovementData() {{
    return movementData.filter(r => norm(r["Void Status"] || r["Void"]).toUpperCase() !== "YES");
}}

function uniqueValues(data, field) {{
    return [...new Set(data.map(r => norm(r[field])).filter(v => v !== ""))].sort();
}}

function populateFilter(id, field) {{
    const sel = document.getElementById(id);
    sel.innerHTML = "<option value=''>All</option>";
    uniqueValues(masterlist, field).forEach(v => {{
        const opt = document.createElement("option");
        opt.value = v;
        opt.textContent = v;
        sel.appendChild(opt);
    }});
    sel.addEventListener("change", render);
}}

function filteredMasterlist() {{
    const dept = document.getElementById("departmentFilter").value;
    const acc = document.getElementById("accountFilter").value;
    const mgr = document.getElementById("managerFilter").value;
    const sup = document.getElementById("supervisorFilter").value;

    return masterlist.filter(r =>
        (!dept || norm(r["Department"]) === dept) &&
        (!acc || norm(r["LOB / Account"]) === acc) &&
        (!mgr || norm(r["Manager"]) === mgr) &&
        (!sup || norm(r["Immediate Supervisor"]) === sup)
    );
}}

function countBy(data, field) {{
    const out = {{}};
    data.forEach(r => {{
        const key = norm(r[field]) || "Blank";
        out[key] = (out[key] || 0) + 1;
    }});
    return Object.entries(out).map(([name, count]) => ({{name, count}})).sort((a,b) => b.count - a.count);
}}

function setText(id, value) {{
    document.getElementById(id).textContent = Number(value).toLocaleString();
}}

function renderDataTable(id, rows, columns) {{
    let html = "<div class='table-scroll'><table><thead><tr>";
    columns.forEach(c => {{
        html += `<th>${{escapeHtml(c.label)}}</th>`;
    }});
    html += "</tr></thead><tbody>";

    if (rows.length === 0) {{
        html += `<tr><td colspan="${{columns.length}}">No records found</td></tr>`;
    }} else {{
        rows.forEach(r => {{
            html += "<tr>";
            columns.forEach(c => {{
                html += `<td>${{escapeHtml(r[c.field])}}</td>`;
            }});
            html += "</tr>";
        }});
    }}

    html += "</tbody></table></div>";
    document.getElementById(id).innerHTML = html;
}}

function donut(id, title, data, textInfo = "percent") {{
    Plotly.newPlot(id, [{{
        labels: data.map(d => d.name),
        values: data.map(d => d.count),
        type: "pie",
        hole: 0.58,
        marker: {{colors: COLORS}},
        textinfo: textInfo,
        hovertemplate: "%{{label}}<br>Headcount: %{{value}}<br>Percentage: %{{percent}}<extra></extra>",
    }}], {{
        title: {{text: title, font: {{color: "#004C97", size: 15}}}},
        height: 300,
        margin: {{l: 20, r: 20, t: 45, b: 20}},
        paper_bgcolor: "white",
        font: {{family: "Arial", size: 11}}
    }}, {{responsive: true}});
}}

function bar(id, title, data, yTitle) {{
    const top = data.slice(0, 10).reverse();
    Plotly.newPlot(id, [{{
        x: top.map(d => d.count),
        y: top.map(d => d.name),
        type: "bar",
        orientation: "h",
        text: top.map(d => d.count),
        textposition: "auto",
        marker: {{color: "#004C97"}}
    }}], {{
        title: {{text: title, font: {{color: "#004C97", size: 15}}}},
        height: 300,
        margin: {{l: 135, r: 20, t: 45, b: 35}},
        xaxis: {{title: "Headcount"}},
        yaxis: {{title: yTitle}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        font: {{family: "Arial", size: 10}}
    }}, {{responsive: true}});
}}

function weeklyChart() {{
    const grouped = {{}};
    historyData.forEach(r => {{
        const week = norm(r["Week"]);
        if (week) grouped[week] = (grouped[week] || 0) + 1;
    }});

    const rows = Object.entries(grouped)
        .map(([week, count]) => ({{week, count}}))
        .sort((a,b) => new Date(a.week) - new Date(b.week));

    Plotly.newPlot("weeklyLine", [{{
        x: rows.map(r => r.week),
        y: rows.map(r => r.count),
        type: "bar",
        text: rows.map(r => r.count),
        textposition: "auto",
        marker: {{color: "#39B54A"}}
    }}], {{
        title: {{text: "Weekly History Records", font: {{color: "#004C97", size: 15}}}},
        height: 300,
        margin: {{l: 45, r: 20, t: 45, b: 35}},
        yaxis: {{title: "Records"}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        font: {{family: "Arial", size: 10}}
    }}, {{responsive: true}});
}}

function masterlistTable(data) {{
    renderDataTable("masterlistTable", data, MASTERLIST_COLUMNS);
}}

function recentMovementsTable() {{
    const rows = filteredMovementData().slice(-10).reverse();
    renderDataTable("recentMovements", rows, [
        {{label: "Effective Date", field: "Effective Date"}},
        {{label: "Employee Name", field: "Employee Name"}},
        {{label: "Movement Type", field: "Movement Type"}},
        {{label: "New Department", field: "New Department"}},
        {{label: "New Account", field: "New Account"}},
        {{label: "New Supervisor", field: "New Supervisor"}},
        {{label: "New Job Title", field: "New Job Title"}},
    ]);
}}

function render() {{
    const data = filteredMasterlist();
    const movements = filteredMovementData();

    const active = data.filter(r => norm(r["Employment Status"]).toUpperCase() === "ACTIVE").length;
    const inactive = data.filter(r => norm(r["Employment Status"]).toUpperCase() === "INACTIVE").length;

    setText("headcount", data.length);
    setText("active", active);
    setText("inactive", inactive);
    setText("movements", movements.length);
    setText("historyRecords", historyData.length);
    setText("departments", uniqueValues(data, "Department").length);
    setText("accounts", uniqueValues(data, "LOB / Account").length);
    setText("managers", uniqueValues(data, "Manager").length);

    donut("deptDonut", "Headcount by Department", countBy(data, "Department"), "value");
    bar("accountBar", "Headcount by Account (Top 10)", countBy(data, "LOB / Account"), "Account");
    donut("activeDonut", "Active vs Inactive", [
        {{name: "Active", count: active}},
        {{name: "Inactive", count: inactive}}
    ]);
    bar("managerBar", "Headcount by Manager (Top 10)", countBy(data, "Manager"), "Manager");
    bar("supervisorBar", "Headcount by Supervisor (Top 10)", countBy(data, "Immediate Supervisor"), "Supervisor");
    donut("employeeGroupDonut", "Employee Group Distribution", countBy(data, "Employee Group"));
    donut("changeTypeDonut", "Change Type Distribution", countBy(historyData, "Change Type"));
    weeklyChart();
    masterlistTable(data);
    recentMovementsTable();
}}

populateFilter("departmentFilter", "Department");
populateFilter("accountFilter", "LOB / Account");
populateFilter("managerFilter", "Manager");
populateFilter("supervisorFilter", "Immediate Supervisor");
render();
</script>

</body>
</html>
"""

    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")
    print(f"Dashboard generated: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

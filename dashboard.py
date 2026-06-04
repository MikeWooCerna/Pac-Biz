import base64
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd

MASTERLIST_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=0&single=true&output=csv"
HISTORY_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=166777136&single=true&output=csv"
MOVEMENT_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS82OdHh0VFVA9K6P3b7Y8pjPmdeJSOQxj1KQ_4ts5HYL4YgUHwKjpFymrPCJdfMK0Rox6fnSTG3rKf/pub?gid=693236738&single=true&output=csv"

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/18hKmm2SmlWqB23osiV3JTF0aWn86vvZ-YJSC-Rr3JcY/edit#gid=0"

OUTPUT_FILE = "masterlist_dashboard.html"
LOGO_FILE = "pacbiz_logo.png"
FAVICON_FILE = "pacbiz_favicon.png"

COACHING_DIR = Path(os.getenv("COACHING_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Coaching"))
COACHING_SCRIPT = Path(os.getenv("COACHING_SCRIPT", str(COACHING_DIR / "asana_pull.py")))
COACHING_CONFIG = Path(os.getenv("COACHING_CONFIG", str(COACHING_DIR / "config.json")))
COACHING_OUTPUT_FILE = Path(os.getenv("COACHING_OUTPUT_FILE", str(COACHING_DIR / "Output" / "coaching_logs.xlsx")))
COACHING_TIMEZONE = ZoneInfo("Asia/Manila")
COACHING_TIMEZONE_NAME = "Asia/Manila"
EXCLUDED_COACHING_GIDS = {
    "1215293220797391",
    "1215291356597096",
    "1215276162401539",
    "1215275720566322",
    "1215275719365880",
    "1215275643385621",
    "1215275244346197",
    "1215275011155202",
    "1215275011105001",
    "1215274826340895",
}

PACBIZ_BLUE = "#004C97"
PACBIZ_GREEN = "#39B54A"


def clean_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df.fillna("")


def get_image_data_uri(filename):
    image_path = Path(filename)
    if not image_path.exists():
        return ""
    encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
    return f"data:image/png;base64,{encoded}"


def to_records(df):
    return json.dumps(df.to_dict(orient="records"), ensure_ascii=False, default=str)


def cell_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y %I:%M %p")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_email(value):
    return cell_text(value).lower()


def normalize_spaces(value):
    return " ".join(cell_text(value).split())


def format_asana_date_time(value):
    if not value:
        return "", ""

    text = cell_text(value)

    if len(text) == 10 and text[4] == "-" and text[7] == "-":
        dt = pd.to_datetime(text, errors="coerce")
        if pd.isna(dt):
            return text, ""
        return dt.strftime("%m/%d/%Y"), ""

    dt = pd.to_datetime(text, errors="coerce", utc=True)

    if pd.isna(dt):
        return text, ""

    dt = dt.tz_convert(COACHING_TIMEZONE_NAME)
    return dt.strftime("%m/%d/%Y"), dt.strftime("%I:%M %p")


def get_asana_custom_value(field):
    return (
        field.get("display_value")
        or field.get("text_value")
        or (field.get("date_value") or {}).get("date_time")
        or (field.get("date_value") or {}).get("date")
        or ((field.get("enum_value") or {}).get("name"))
        or ""
    )


def get_coaching_asana_config():
    token = os.getenv("ASANA_TOKEN", "").strip()
    project_id = (
        os.getenv("ASANA_PROJECT_ID", "").strip()
        or os.getenv("COACHING_PROJECT_ID", "").strip()
    )

    if token and project_id:
        return token, project_id

    if not COACHING_CONFIG.exists():
        return "", ""

    try:
        config = json.loads(COACHING_CONFIG.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"Skipping Coaching config load: {exc}")
        return "", ""

    return config.get("asana_token", ""), config.get("project_id", "")


def pull_coaching_from_asana():
    token, project_id = get_coaching_asana_config()
    if not token or not project_id:
        return pd.DataFrame()

    try:
        import requests
    except ImportError:
        print("Skipping Coaching Asana pull: requests is not installed.")
        return pd.DataFrame()

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    rows = []
    offset = None

    while True:
        params = {
            "project": project_id,
            "limit": 100,
            "opt_fields": ",".join([
                "gid",
                "name",
                "completed",
                "created_at",
                "modified_at",
                "created_by.name",
                "created_by.email",
                "custom_fields.name",
                "custom_fields.display_value",
                "custom_fields.text_value",
                "custom_fields.date_value",
                "custom_fields.enum_value.name",
            ]),
        }

        if offset:
            params["offset"] = offset

        try:
            response = requests.get(
                "https://app.asana.com/api/1.0/tasks",
                headers=headers,
                params=params,
                timeout=30,
            )
            response.raise_for_status()
        except requests.RequestException as exc:
            print(f"Skipping Coaching Asana pull: {exc}")
            return pd.DataFrame()

        data = response.json()

        for task in data.get("data", []):
            custom_fields = {}
            for field in task.get("custom_fields", []):
                custom_fields[field.get("name")] = get_asana_custom_value(field)

            coaching_date, coaching_time = format_asana_date_time(
                custom_fields.get("Date & Time of Coaching", "")
            )
            created_date, created_time = format_asana_date_time(task.get("created_at"))
            modified_date, modified_time = format_asana_date_time(task.get("modified_at"))

            rows.append({
                "Task GID": task.get("gid"),
                "Task Name": task.get("name"),
                "Completed": task.get("completed"),
                "Employee Name": custom_fields.get("Employee Name", ""),
                "Coaching Date": coaching_date,
                "Coaching Time": coaching_time,
                "Status": custom_fields.get("Status", ""),
                "Employee Email": custom_fields.get("Employee Email", ""),
                "Supervisor Email": (
                    custom_fields.get("Supervisors Email", "")
                    or custom_fields.get("Supervisor Email", "")
                ),
                "Agreed Action Steps": (
                    custom_fields.get("Agreed Action Steps:", "")
                    or custom_fields.get("Agreed Action Steps", "")
                ),
                "Improvement Timeline / Follow-Up Date": (
                    custom_fields.get("Improvement Timeline / Follow-Up Date:", "")
                    or custom_fields.get("Improvement Timeline / Follow-Up Date", "")
                    or custom_fields.get("Improvement Timeline/Follow-Up Date:", "")
                    or custom_fields.get("Improvement Timeline/Follow-Up Date", "")
                ),
                "Created By": (task.get("created_by") or {}).get("name", ""),
                "Created By Email": (task.get("created_by") or {}).get("email", ""),
                "Created Date": created_date,
                "Created Time": created_time,
                "Modified Date": modified_date,
                "Modified Time": modified_time,
                "Last Refreshed": datetime.now(COACHING_TIMEZONE).strftime("%m/%d/%Y %I:%M %p"),
            })

        next_page = data.get("next_page")
        if not next_page:
            break
        offset = next_page["offset"]

    print(f"Coaching rows pulled from Asana: {len(rows)}")
    return clean_columns(pd.DataFrame(rows))


def refresh_coaching_output():
    if not COACHING_SCRIPT.exists():
        return False

    try:
        result = subprocess.run(
            [sys.executable, str(COACHING_SCRIPT)],
            cwd=str(COACHING_SCRIPT.parent),
            capture_output=True,
            text=True,
            timeout=240,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Coaching script refresh: {exc}")
        return False

    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip() or "No details returned."
        print(f"Skipping Coaching script refresh: {message}")
        return False

    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_coaching_workbook():
    if not COACHING_OUTPUT_FILE.exists():
        return pd.DataFrame()

    try:
        return clean_columns(pd.read_excel(COACHING_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Coaching workbook load: {exc}")
        return pd.DataFrame()


def load_existing_records_from_html(variable_name):
    output_path = Path(OUTPUT_FILE)
    if not output_path.exists():
        return pd.DataFrame()

    try:
        html = output_path.read_text(encoding="utf-8")
    except OSError:
        return pd.DataFrame()

    marker = f"const {variable_name} = "
    start = html.find(marker)
    if start < 0:
        return pd.DataFrame()

    start += len(marker)
    end = html.find(";\n", start)
    if end < 0:
        return pd.DataFrame()

    try:
        return clean_columns(pd.DataFrame(json.loads(html[start:end])))
    except (json.JSONDecodeError, ValueError):
        return pd.DataFrame()


COACHING_CATEGORY_RULES = {
    "Behavioral": [
        ("professionalism", 3, "professionalism"),
        ("accountability", 3, "accountability"),
        ("attitude", 3, "attitude"),
        ("engagement", 2, "engagement"),
        ("engaged", 2, "engagement"),
        ("responsiveness", 3, "responsiveness"),
        ("ownership", 3, "ownership"),
        ("workplace conduct", 3, "workplace conduct"),
        ("conduct", 2, "conduct"),
        ("attentive", 2, "attentiveness"),
    ],
    "Communication": [
        ("notify", 3, "failure to notify"),
        ("notification", 2, "notification"),
        ("communicate", 3, "communication"),
        ("communication", 3, "communication"),
        ("message", 2, "messaging"),
        ("messaging", 2, "messaging"),
        ("unclear", 2, "unclear messaging"),
        ("clarify", 3, "clarification"),
        ("clarifying", 3, "clarification"),
        ("clarification", 3, "clarification"),
        ("confirm my understanding", 3, "confirmation gap"),
        ("channel", 2, "communication channel"),
    ],
    "Attendance & Adherence": [
        ("break", 3, "break adherence"),
        ("breaks", 3, "break adherence"),
        ("bio break", 4, "break adherence"),
        ("terminal break", 4, "break adherence"),
        ("tardy", 4, "tardiness"),
        ("tardiness", 4, "tardiness"),
        ("late", 2, "tardiness"),
        ("schedule adherence", 4, "schedule adherence"),
        ("adherence", 2, "adherence"),
        ("log in", 3, "login/logout"),
        ("login", 3, "login/logout"),
        ("log out", 3, "login/logout"),
        ("logout", 3, "login/logout"),
        ("attendance", 4, "attendance"),
        ("shift", 2, "shift compliance"),
    ],
    "Process Compliance": [
        ("sop", 4, "SOP"),
        ("workflow", 4, "workflow"),
        ("procedure", 3, "procedure"),
        ("procedures", 3, "procedure"),
        ("required steps", 3, "required steps"),
        ("documentation", 3, "documentation"),
        ("document", 2, "documentation"),
        ("account-specific", 4, "account instruction"),
        ("account specific", 4, "account instruction"),
        ("account instruction", 4, "account instruction"),
        ("account requirements", 3, "account requirement"),
        ("process", 2, "process"),
    ],
    "Performance & Quality": [
        ("qa score", 4, "QA score"),
        ("quality", 3, "quality standard"),
        ("accuracy", 4, "accuracy"),
        ("productivity", 4, "productivity"),
        ("call handling", 4, "call handling"),
        ("call", 2, "call handling"),
        ("score", 2, "score"),
        ("error", 3, "accuracy"),
        ("missed", 2, "missed standard"),
        ("standard", 2, "standard"),
        ("standards", 2, "standard"),
    ],
    "Knowledge & Training": [
        ("product knowledge", 4, "product knowledge"),
        ("lack of knowledge", 4, "knowledge gap"),
        ("knowledge", 2, "knowledge"),
        ("process understanding", 4, "process understanding"),
        ("understanding", 2, "understanding"),
        ("training", 4, "training need"),
        ("skill gap", 4, "skill gap"),
        ("coaching for skill", 4, "skill gap"),
    ],
    "Policy Compliance": [
        ("company policy", 4, "company policy"),
        ("company policies", 4, "company policy"),
        ("policy violation", 4, "policy violation"),
        ("policy", 3, "policy"),
        ("incident report", 4, "incident report"),
        ("hr", 3, "HR escalation"),
        ("non-compliance", 4, "non-compliance"),
        ("repeated non-compliance", 5, "repeated non-compliance"),
    ],
    "Customer Experience": [
        ("empathy", 4, "empathy"),
        ("tone", 4, "tone"),
        ("customer satisfaction", 4, "customer satisfaction"),
        ("customer handling", 4, "customer handling"),
        ("customer", 2, "customer handling"),
        ("customers", 2, "customer handling"),
        ("service delivery", 4, "service delivery"),
        ("client", 2, "client impact"),
        ("partner relations", 3, "partner relations"),
    ],
}

COACHING_OUTPUT_COLUMNS = [
    "Coaching ID",
    "Emp Name",
    "Coached by",
    "Coaching Details",
    "Remarks/Comment",
    "Coaching Category",
    "Confidence Level",
    "Reason",
    "Coaching Date",
    "Coaching Status",
    "Created Date",
    "Created Time",
]


def join_reason_labels(labels):
    unique_labels = []
    for label in labels:
        if label not in unique_labels:
            unique_labels.append(label)

    if "adherence" in unique_labels and any(
        label != "adherence" and "adherence" in label for label in unique_labels
    ):
        unique_labels = [label for label in unique_labels if label != "adherence"]

    if not unique_labels:
        return "matched indicators"
    if len(unique_labels) == 1:
        return unique_labels[0]
    if len(unique_labels) == 2:
        return f"{unique_labels[0]} and {unique_labels[1]}"
    return f"{unique_labels[0]}, {unique_labels[1]}, and {unique_labels[2]}"


def indicator_matches(text, needle):
    pattern = rf"(?<![A-Za-z0-9]){re.escape(needle)}(?![A-Za-z0-9])"
    return re.search(pattern, text) is not None


def classify_coaching_details(action_steps):
    text = normalize_spaces(action_steps)
    if len(text) < 12:
        return "Other / Review Required", "Low", "Text is too vague for reliable classification."

    lowered = text.lower()
    matches = []

    for category, indicators in COACHING_CATEGORY_RULES.items():
        score = 0
        labels = []
        for needle, weight, label in indicators:
            if indicator_matches(lowered, needle):
                score += weight
                labels.append(label)
        if score:
            matches.append((category, score, labels))

    if not matches:
        return "Other / Review Required", "Low", "No reliable category indicators found."

    matches.sort(key=lambda item: (-item[1], item[0]))
    top_category, top_score, top_labels = matches[0]
    second_category = matches[1][0] if len(matches) > 1 else ""
    second_score = matches[1][1] if len(matches) > 1 else 0

    if top_score <= 1:
        return "Other / Review Required", "Low", "Only weak indicators found; review manually."

    if (top_score >= 4 and top_score >= second_score + 2) or (top_score >= 3 and second_score == 0):
        confidence = "High"
        reason = f"Clear cues: {join_reason_labels(top_labels)}."
    else:
        confidence = "Medium"
        overlap = f"; overlaps with {second_category}" if second_category else ""
        reason = f"Primary cues: {join_reason_labels(top_labels)}{overlap}."

    return top_category, confidence, reason


def build_email_name_lookup(masterlist):
    lookup = {}
    if "Company Email" not in masterlist.columns or "Emp Name" not in masterlist.columns:
        return lookup

    for _, row in masterlist.iterrows():
        email = normalize_email(row.get("Company Email"))
        name = cell_text(row.get("Emp Name"))
        if email and name:
            lookup[email] = name

    return lookup


def resolve_name(email_lookup, email, fallback=""):
    normalized_email = normalize_email(email)
    if normalized_email in email_lookup:
        return email_lookup[normalized_email]
    return cell_text(fallback) or cell_text(email)


def transform_coaching_logs(source, masterlist):
    if source.empty:
        return pd.DataFrame(columns=COACHING_OUTPUT_COLUMNS)

    if all(column in source.columns for column in COACHING_OUTPUT_COLUMNS):
        return source[COACHING_OUTPUT_COLUMNS]

    email_lookup = build_email_name_lookup(masterlist)
    records = []

    for _, row in source.iterrows():
        coaching_id = cell_text(row.get("Task GID"))
        if coaching_id in EXCLUDED_COACHING_GIDS:
            continue

        details = cell_text(row.get("Agreed Action Steps"))
        category, confidence, reason = classify_coaching_details(details)

        records.append({
            "Coaching ID": coaching_id,
            "Emp Name": resolve_name(
                email_lookup,
                row.get("Employee Email"),
                row.get("Employee Name"),
            ),
            "Coached by": resolve_name(email_lookup, row.get("Supervisor Email")),
            "Coaching Details": details,
            "Remarks/Comment": cell_text(row.get("Improvement Timeline / Follow-Up Date")),
            "Coaching Category": category,
            "Confidence Level": confidence,
            "Reason": reason,
            "Coaching Date": cell_text(row.get("Coaching Date")),
            "Coaching Status": cell_text(row.get("Status")),
            "Created Date": cell_text(row.get("Created Date")),
            "Created Time": cell_text(row.get("Created Time")),
        })

    return clean_columns(pd.DataFrame(records, columns=COACHING_OUTPUT_COLUMNS))


def load_coaching_data(masterlist):
    refresh_coaching_output()

    source = read_coaching_workbook()
    if source.empty:
        source = pull_coaching_from_asana()

    if source.empty:
        existing = load_existing_records_from_html("coachingData")
        if not existing.empty:
            return transform_coaching_logs(existing, masterlist)

    return transform_coaching_logs(source, masterlist)


def main():
    masterlist = clean_columns(pd.read_csv(MASTERLIST_CSV))
    history = clean_columns(pd.read_csv(HISTORY_CSV))
    movement = clean_columns(pd.read_csv(MOVEMENT_CSV))
    coaching = load_coaching_data(masterlist)

    refresh_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")

    latest_snapshot = ""
    if "Date Generated" in history.columns:
        latest_date = pd.to_datetime(history["Date Generated"], errors="coerce").max()
        latest_snapshot = latest_date.strftime("%m/%d/%Y") if pd.notna(latest_date) else ""

    logo_uri = get_image_data_uri(LOGO_FILE)
    favicon_uri = get_image_data_uri(FAVICON_FILE) or logo_uri
    masterlist_source_url = GOOGLE_SHEET_URL or MASTERLIST_CSV
    masterlist_source_label = "Open Google Sheet" if GOOGLE_SHEET_URL else "Open Source CSV"
    masterlist_excel_url = MASTERLIST_CSV.split("/pub?")[0] + "/pub?output=xlsx"

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Pac-Biz Dashboard</title>
{"<link rel='icon' type='image/png' href='" + favicon_uri + "'>" if favicon_uri else ""}
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

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

    .tabs {{
        display: flex;
        gap: 8px;
        padding: 10px 18px 0;
        background: var(--bg);
    }}

    .tab-button {{
        --tab-accent: var(--blue);
        border: 1px solid #CBD5E1;
        border-bottom: 3px solid transparent;
        background: white;
        color: var(--tab-accent);
        border-radius: 8px 8px 0 0;
        padding: 10px 16px;
        font-size: 13px;
        font-weight: 800;
        cursor: pointer;
        transition: background .18s ease, color .18s ease, border-color .18s ease;
    }}

    .tab-button[data-tab="masterlist"] {{
        --tab-accent: var(--blue);
    }}

    .tab-button[data-tab="coaching"] {{
        --tab-accent: var(--green);
    }}

    .tab-button[data-tab="quality"] {{
        --tab-accent: var(--dark-blue);
    }}

    .tab-button:hover {{
        border-color: var(--tab-accent);
        background: #F8FAFC;
    }}

    .tab-button.active {{
        border-color: var(--tab-accent);
        background: var(--tab-accent);
        color: white;
    }}

    .masterlist-controls.hidden,
    .coaching-controls.hidden {{
        display: none;
    }}

    .tab-panel {{
        display: none;
    }}

    .tab-panel.active {{
        display: block;
    }}

    .under-construction {{
        min-height: calc(100vh - 180px);
        display: grid;
        place-items: center;
        color: var(--dark-blue);
        font-size: 28px;
        font-weight: 900;
        text-align: center;
        padding: 36px;
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

    .coaching-filters {{
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 12px;
        padding: 14px 18px 8px;
    }}

    .multi-filter {{
        position: relative;
    }}

    .multi-filter summary {{
        list-style: none;
        cursor: pointer;
        color: var(--text);
        font-size: 13px;
        min-height: 22px;
    }}

    .multi-filter summary::-webkit-details-marker {{
        display: none;
    }}

    .multi-filter summary::after {{
        content: "v";
        float: right;
        color: var(--blue);
        font-size: 11px;
        margin-top: 2px;
    }}

    .multi-filter[open] summary::after {{
        content: "^";
    }}

    .multi-options {{
        position: absolute;
        left: 0;
        right: 0;
        top: calc(100% + 8px);
        z-index: 20;
        max-height: 230px;
        overflow: auto;
        background: white;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.18);
        padding: 8px;
    }}

    .multi-option {{
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 6px 8px;
        border-radius: 6px;
        font-size: 12px;
        color: var(--text);
    }}

    .multi-option:hover {{
        background: #F1F5F9;
    }}

    .multi-option.select-all {{
        border-bottom: 1px solid #E5E7EB;
        border-radius: 0;
        color: var(--blue);
        font-weight: 900;
        margin-bottom: 4px;
        padding-bottom: 8px;
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

    .coaching-cards {{
        grid-template-columns: repeat(3, minmax(0, 1fr));
    }}

    .card.completed-card {{
        border-top-color: var(--green);
    }}

    .card.completed-card .value {{
        color: var(--green);
    }}

    .card.pending-card {{
        border-top-color: #FFC000;
    }}

    .card.pending-card .value {{
        color: #FFC000;
    }}

    .grid {{
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 12px;
        padding: 0 18px 20px;
    }}

    .coaching-grid {{
        padding-top: 8px;
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
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 12px;
    }}

    .donut-chart-row .chart-card {{
        min-height: 295px;
    }}

    .coaching-chart-row {{
        display: grid;
        grid-column: 1 / -1;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 12px;
    }}

    .coaching-summary-card {{
        grid-column: 1;
    }}

    .empty-widget-space {{
        min-height: 1px;
    }}

    .chart-stack {{
        display: grid;
        gap: 12px;
    }}

    .table-scroll {{
        max-height: 430px;
        overflow: auto;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
    }}

    .table-heading {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin: 4px 0 12px;
    }}

    .table-heading h3 {{
        color: var(--blue);
        margin: 0;
    }}

    .table-actions {{
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
    }}

    .table-meta {{
        display: inline-flex;
        align-items: center;
        min-height: 32px;
        color: var(--muted);
        font-size: 12px;
        font-weight: 800;
    }}

    .table-action {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 32px;
        border: 1px solid var(--green);
        border-radius: 8px;
        padding: 7px 12px;
        background: var(--green);
        color: white;
        font-size: 12px;
        font-weight: 800;
        font-family: inherit;
        text-decoration: none;
        cursor: pointer;
    }}

    .table-action.secondary {{
        border-color: var(--blue);
        background: white;
        color: var(--blue);
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

    th.sortable {{
        cursor: pointer;
        user-select: none;
    }}

    .sort-indicator {{
        display: inline-block;
        min-width: 12px;
        margin-left: 4px;
        opacity: .9;
    }}

    td {{
        padding: 7px 8px;
        border-bottom: 1px solid #E5E7EB;
        vertical-align: top;
    }}

    .long-text {{
        min-width: 280px;
        max-width: 520px;
        white-space: pre-wrap;
        line-height: 1.35;
    }}

    .nowrap {{
        white-space: nowrap;
    }}

    .disclaimer {{
        grid-column: 1 / -1;
        color: var(--muted);
        font-size: 12px;
        font-style: italic;
        padding: 0 2px 4px;
    }}

    .disclaimer .completed-word {{
        color: var(--green);
        font-weight: 800;
    }}

    .score-gauge {{
        height: 280px;
        display: grid;
        align-items: center;
        justify-items: center;
        padding: 4px 0 0;
    }}

    .score-gauge svg {{
        width: min(100%, 560px);
        height: 260px;
        overflow: visible;
    }}

    .gauge-label {{
        font-family: Arial, sans-serif;
        font-size: 15px;
        font-weight: 900;
        fill: #111827;
        dominant-baseline: middle;
        text-anchor: middle;
    }}

    .gauge-value {{
        font-size: 28px;
        font-weight: 900;
        fill: var(--dark-blue);
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

        .coaching-filters,
        .coaching-cards,
        .coaching-chart-row {{
            grid-template-columns: 1fr;
        }}

        .bar-chart-row {{
            grid-template-columns: 1fr;
        }}

        .donut-chart-row {{
            grid-template-columns: 1fr;
        }}

        .table-heading {{
            align-items: flex-start;
            flex-direction: column;
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

<div class="tabs" role="tablist" aria-label="Dashboard sections">
    <button class="tab-button active" type="button" data-tab="masterlist" role="tab" aria-selected="true">Masterlist</button>
    <button class="tab-button" type="button" data-tab="coaching" role="tab" aria-selected="false">Coaching</button>
    <button class="tab-button" type="button" data-tab="quality" role="tab" aria-selected="false">Quality</button>
</div>

<div class="masterlist-controls" id="masterlistControls">
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

<div class="coaching-controls hidden" id="coachingControls">
<div class="coaching-filters">
    <div class="filter-box">
        <label>Emp Name</label>
        <details class="multi-filter" id="coachingEmpFilter">
            <summary id="coachingEmpFilterSummary">All</summary>
            <div class="multi-options" id="coachingEmpOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Coached by</label>
        <details class="multi-filter" id="coachingLeaderFilter">
            <summary id="coachingLeaderFilterSummary">All</summary>
            <div class="multi-options" id="coachingLeaderOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Month Yr</label>
        <details class="multi-filter" id="coachingMonthFilter">
            <summary id="coachingMonthFilterSummary">All</summary>
            <div class="multi-options" id="coachingMonthOptions"></div>
        </details>
    </div>
</div>

<div class="cards coaching-cards">
    <div class="card"><div class="label">Coaching Count</div><div class="value" id="coachingCount">0</div></div>
    <div class="card completed-card"><div class="label">Completed</div><div class="value" id="coachingCompleted">0</div></div>
    <div class="card pending-card"><div class="label">Pending</div><div class="value" id="coachingPending">0</div></div>
</div>
</div>
</div>

<div class="tab-panel active" id="masterlistPanel" data-tab="masterlist" role="tabpanel">
<div class="grid">
    <div class="donut-chart-row">
        <div class="chart-card"><div id="deptDonut"></div></div>
        <div class="chart-card"><div id="activeDonut"></div></div>
        <div class="chart-card"><div id="employeeGroupDonut"></div></div>
        <div class="chart-card"><div id="employmentClassDonut"></div></div>
    </div>
    <div class="bar-chart-row">
        <div class="chart-stack">
            <div class="chart-card"><div id="accountBar"></div></div>
            <div class="chart-card"><div id="accountTenureStack"></div></div>
        </div>
        <div class="chart-card"><div id="managerBar"></div></div>
        <div class="chart-card"><div id="supervisorBar"></div></div>
    </div>
    <div class="chart-card"><div id="tenureSegmentation"></div></div>
    <div class="chart-card"><div id="ageGroupBar"></div></div>
    <div class="chart-card"><div id="weeklyLine"></div></div>
    <div class="chart-card full">
        <div class="table-heading">
            <h3>Master List</h3>
            <div class="table-actions">
                <a class="table-action" href="{masterlist_source_url}" target="_blank" rel="noopener">{masterlist_source_label}</a>
                <a class="table-action secondary" href="{masterlist_excel_url}" target="_blank" rel="noopener">Download Excel</a>
            </div>
        </div>
        <div id="masterlistTable"></div>
    </div>
    <div class="chart-card full">
        <h3 style="color:#004C97;margin:4px 0 12px;">Recent Employee Movements</h3>
        <div id="recentMovements"></div>
    </div>
</div>
</div>

<div class="tab-panel" id="coachingPanel" data-tab="coaching" role="tabpanel">
<div class="grid coaching-grid">
    <div class="coaching-chart-row">
        <div class="chart-card"><div id="coachingCategoryDonut"></div></div>
        <div class="chart-card"><div id="coachingConfidenceGauge"></div></div>
    </div>
    <div class="chart-card coaching-summary-card">
        <div class="table-heading">
            <h3>Summary</h3>
            <div class="table-actions">
                <span class="table-meta" id="coachingSummaryMeta">0 completed coaching sessions</span>
            </div>
        </div>
        <div id="coachingSummaryTable"></div>
        <div class="disclaimer">Coaching will be only be counted once status is <span class="completed-word">Completed</span></div>
    </div>
    <div class="empty-widget-space"></div>
    <div class="chart-card full">
        <div class="table-heading">
            <h3>Coaching Logs</h3>
            <div class="table-actions">
                <button class="table-action" type="button" id="downloadCoachingExcel">Download Excel</button>
                <button class="table-action secondary" type="button" id="openCoachingSheets">Google Sheets</button>
                <span class="table-meta" id="coachingLoadedMeta">{len(coaching):,} records loaded</span>
            </div>
        </div>
        <div id="coachingTable"></div>
    </div>
</div>
</div>

<div class="tab-panel" id="qualityPanel" data-tab="quality" role="tabpanel">
    <div class="under-construction">Under Construction</div>
</div>

<div class="footer">
    Developed for Pac-Biz Reporting MCerna | Data Source: Master List | Automation: Python 3.13.0
</div>

<script>
const masterlist = {to_records(masterlist)};
const historyData = {to_records(history)};
const movementData = {to_records(movement)};
const coachingData = {to_records(coaching)};

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
const COACHING_COLUMNS = [
    {{label: "Coaching ID", field: "Coaching ID", className: "nowrap"}},
    {{label: "Emp Name", field: "Emp Name", className: "nowrap", sortable: true}},
    {{label: "Coached by", field: "Coached by", className: "nowrap", sortable: true}},
    {{label: "Coaching Details", field: "Coaching Details", className: "long-text"}},
    {{label: "Remarks/Comment", field: "Remarks/Comment", className: "long-text"}},
    {{label: "Coaching Category", field: "Coaching Category", className: "nowrap", sortable: true}},
    {{label: "Confidence Level", field: "Confidence Level", className: "nowrap"}},
    {{label: "Reason", field: "Reason", className: "long-text"}},
    {{label: "Coaching Date", field: "Coaching Date", className: "nowrap", sortable: true, sortType: "date"}},
    {{label: "Coaching Status", field: "Coaching Status", className: "nowrap"}},
    {{label: "Created Date", field: "Created Date", className: "nowrap", sortable: true, sortType: "date"}},
    {{label: "Created Time", field: "Created Time", className: "nowrap"}},
];
const COACHING_SUMMARY_COLUMNS = [
    {{label: "Team Leader", field: "Team Leader", className: "nowrap"}},
];
const TENURE_GROUPS = [
    {{name: "0-30 Days", maxDays: 30}},
    {{name: "31-60 Days", maxDays: 60}},
    {{name: "61-90 Days", maxDays: 90}},
    {{name: "91-180 Days", maxDays: 180}},
    {{name: "181-365 Days", maxDays: 365}},
    {{name: "1-2 Years", maxDays: 730}},
    {{name: "2-5 Years", maxDays: 1825}},
    {{name: "5+ Years", maxDays: Infinity}},
];
const AGE_GROUPS = [
    {{name: "Under 20", min: 0, max: 19}},
    {{name: "20-24", min: 20, max: 24}},
    {{name: "25-29", min: 25, max: 29}},
    {{name: "30-34", min: 30, max: 34}},
    {{name: "35-39", min: 35, max: 39}},
    {{name: "40-49", min: 40, max: 49}},
    {{name: "50+", min: 50, max: Infinity}},
];
const MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const COACHING_FILTERS = {{
    emp: new Set(),
    leader: new Set(),
    month: new Set(),
}};
const coachingSortState = {{
    field: "Coaching Date",
    direction: "desc",
}};

function norm(v) {{
    return (v ?? "").toString().trim();
}}

function parseDateValue(v) {{
    const raw = norm(v);
    if (!raw) return null;

    const monthMap = {{
        jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
        jul: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11
    }};

    const textDate = raw.match(/^(\\d{{1,2}})[-\\s]([A-Za-z]{{3,}})[-\\s](\\d{{2,4}})$/);
    if (textDate) {{
        let year = Number(textDate[3]);
        if (year < 100) year += year < 50 ? 2000 : 1900;
        const month = monthMap[textDate[2].toLowerCase()];
        if (month !== undefined) return new Date(year, month, Number(textDate[1]));
    }}

    const slashDate = raw.match(/^(\\d{{1,2}})\\/(\\d{{1,2}})\\/(\\d{{2,4}})$/);
    if (slashDate) {{
        let year = Number(slashDate[3]);
        if (year < 100) year += year < 50 ? 2000 : 1900;
        return new Date(year, Number(slashDate[1]) - 1, Number(slashDate[2]));
    }}

    const parsed = new Date(raw);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
}}

function wholeDayDiff(startDate, endDate) {{
    const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
    return Math.floor((end - start) / 86400000);
}}

function tenureGroupName(hireDate, asOfDate) {{
    if (!hireDate || !asOfDate) return "";
    const days = wholeDayDiff(hireDate, asOfDate);
    if (days < 0) return "";
    return TENURE_GROUPS.find(group => days <= group.maxDays)?.name || "";
}}

function tenureCounts(data, asOfField = "") {{
    const counts = Object.fromEntries(TENURE_GROUPS.map(group => [group.name, 0]));
    data.forEach(r => {{
        const hireDate = parseDateValue(r["Hire Date"]);
        const asOfDate = asOfField ? parseDateValue(r[asOfField]) : new Date();
        const groupName = tenureGroupName(hireDate, asOfDate);
        if (groupName) counts[groupName] += 1;
    }});
    return TENURE_GROUPS.map(group => ({{name: group.name, count: counts[group.name]}}));
}}

function isInvalidDob(v) {{
    return ["1/0/00", "01/00/00", "1/0/2000", "01/00/2000"].includes(norm(v));
}}

function ageGroupName(age) {{
    const n = Number(age);
    if (!Number.isFinite(n) || n < 0) return "";
    return AGE_GROUPS.find(group => n >= group.min && n <= group.max)?.name || "";
}}

function ageGroupCounts(data) {{
    const counts = Object.fromEntries(AGE_GROUPS.map(group => [group.name, 0]));
    data.forEach(r => {{
        if (isInvalidDob(r["DOB"])) return;
        const groupName = ageGroupName(norm(r["Age"]).replace(/,/g, ""));
        if (groupName) counts[groupName] += 1;
    }});
    return AGE_GROUPS.map(group => ({{name: group.name, count: counts[group.name]}}));
}}

function formatDateOnly(v) {{
    const raw = norm(v);
    if (!raw) return "";
    const firstPart = raw.split(" ")[0];
    const parsed = parseDateValue(firstPart);
    if (!parsed) return firstPart;
    return parsed.toLocaleDateString("en-US");
}}

function coachingMonthKey(v) {{
    const parsed = parseDateValue(v);
    if (!parsed) return "";
    return `${{parsed.getFullYear()}}-${{String(parsed.getMonth() + 1).padStart(2, "0")}}`;
}}

function coachingMonthLabelFromKey(key) {{
    const match = norm(key).match(/^(\\d{{4}})-(\\d{{2}})$/);
    if (!match) return "";
    const year = Number(match[1]);
    const month = Number(match[2]) - 1;
    return `${{MONTH_LABELS[month]}} '${{String(year).slice(-2)}}`;
}}

function statusKey(v) {{
    return norm(v).toUpperCase().replace(/\\s+/g, " ");
}}

function isCompletedStatus(v) {{
    return statusKey(v) === "COMPLETED";
}}

function isPendingStatus(v) {{
    const value = statusKey(v);
    return value === "FOR ACKNOWLEDGEMENT" || value === "FOR ACKNOWLEDGMENT";
}}

function weekRangeLabel(v) {{
    const parsed = parseDateValue(v);
    if (!parsed) return "";
    const start = weekStartDate(parsed);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    return `${{formatShortDate(start)}} - ${{formatShortDate(end)}}`;
}}

function weekStartDate(date) {{
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const day = start.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    start.setDate(start.getDate() + mondayOffset);
    return start;
}}

function weekStartKey(v) {{
    const parsed = parseDateValue(v);
    if (!parsed) return "";
    const start = weekStartDate(parsed);
    return `${{start.getFullYear()}}-${{String(start.getMonth() + 1).padStart(2, "0")}}-${{String(start.getDate()).padStart(2, "0")}}`;
}}

function weekStartLabel(key) {{
    const parsed = parseDateValue(key);
    if (!parsed) return key;
    return parsed.toLocaleDateString("en-US");
}}

function formatShortDate(date) {{
    return `${{MONTH_LABELS[date.getMonth()]}} ${{date.getDate()}}, '${{String(date.getFullYear()).slice(-2)}}`;
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

function setMultiSummary(summaryId, selectedSet) {{
    const summary = document.getElementById(summaryId);
    if (!summary) return;
    if (selectedSet.size === 0) {{
        summary.textContent = "All";
    }} else if (selectedSet.size === 1) {{
        const value = [...selectedSet][0];
        summary.textContent = summaryId === "coachingMonthFilterSummary" ? coachingMonthLabelFromKey(value) : value;
    }} else {{
        summary.textContent = `${{selectedSet.size}} selected`;
    }}
}}

function syncMultiFilterOptions(optionsId, selectedSet, allValues) {{
    const box = document.getElementById(optionsId);
    const selectAll = box?.querySelector("input[data-select-all='true']");
    if (!selectAll) return;

    if (selectedSet.size === allValues.length) {{
        selectedSet.clear();
    }}

    const allSelected = selectedSet.size === 0;
    selectAll.checked = allSelected;
    selectAll.indeterminate = !allSelected && selectedSet.size < allValues.length;
    box.querySelectorAll("input[type='checkbox']:not([data-select-all='true'])").forEach(checkbox => {{
        checkbox.checked = allSelected || selectedSet.has(checkbox.value);
    }});
}}

function populateMultiFilter(optionsId, summaryId, values, selectedSet) {{
    const box = document.getElementById(optionsId);
    if (!box) return;
    box.innerHTML = "";
    const normalizedValues = values.map(item => ({{
        value: typeof item === "string" ? item : item.value,
        label: typeof item === "string" ? item : item.label,
    }}));
    const allValues = normalizedValues.map(item => item.value);

    const selectAllOption = document.createElement("label");
    selectAllOption.className = "multi-option select-all";

    const selectAllCheckbox = document.createElement("input");
    selectAllCheckbox.type = "checkbox";
    selectAllCheckbox.dataset.selectAll = "true";
    selectAllCheckbox.addEventListener("change", () => {{
        selectedSet.clear();
        syncMultiFilterOptions(optionsId, selectedSet, allValues);
        setMultiSummary(summaryId, selectedSet);
        renderCoaching();
    }});

    const selectAllText = document.createElement("span");
    selectAllText.textContent = "Select All";
    selectAllOption.appendChild(selectAllCheckbox);
    selectAllOption.appendChild(selectAllText);
    box.appendChild(selectAllOption);

    normalizedValues.forEach(item => {{
        const value = item.value;
        const label = item.label;
        const option = document.createElement("label");
        option.className = "multi-option";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = value;
        checkbox.checked = selectedSet.size === 0 || selectedSet.has(value);
        checkbox.addEventListener("change", () => {{
            if (selectedSet.size === 0) {{
                allValues.forEach(v => selectedSet.add(v));
            }}
            if (checkbox.checked) {{
                selectedSet.add(value);
            }} else {{
                selectedSet.delete(value);
            }}
            syncMultiFilterOptions(optionsId, selectedSet, allValues);
            setMultiSummary(summaryId, selectedSet);
            renderCoaching();
        }});

        const text = document.createElement("span");
        text.textContent = label;

        option.appendChild(checkbox);
        option.appendChild(text);
        box.appendChild(option);
    }});

    syncMultiFilterOptions(optionsId, selectedSet, allValues);
    setMultiSummary(summaryId, selectedSet);
}}

function populateCoachingFilters() {{
    populateMultiFilter(
        "coachingEmpOptions",
        "coachingEmpFilterSummary",
        uniqueValues(coachingData, "Emp Name"),
        COACHING_FILTERS.emp
    );
    populateMultiFilter(
        "coachingLeaderOptions",
        "coachingLeaderFilterSummary",
        uniqueValues(coachingData, "Coached by"),
        COACHING_FILTERS.leader
    );

    const monthValues = [...new Set(coachingData.map(r => coachingMonthKey(r["Coaching Date"])).filter(Boolean))]
        .sort()
        .map(value => ({{value, label: coachingMonthLabelFromKey(value)}}));
    populateMultiFilter(
        "coachingMonthOptions",
        "coachingMonthFilterSummary",
        monthValues,
        COACHING_FILTERS.month
    );
}}

function closeMultiFilters(exceptFilter = null) {{
    document.querySelectorAll(".multi-filter[open]").forEach(filter => {{
        if (filter !== exceptFilter) {{
            filter.removeAttribute("open");
        }}
    }});
}}

function initMultiFilterBehavior() {{
    document.querySelectorAll(".multi-filter").forEach(filter => {{
        filter.addEventListener("toggle", () => {{
            if (filter.open) {{
                closeMultiFilters(filter);
            }}
        }});
    }});

    document.querySelectorAll(".multi-options").forEach(options => {{
        options.addEventListener("click", event => {{
            event.stopPropagation();
        }});
    }});

    document.addEventListener("click", event => {{
        if (!event.target.closest(".multi-filter")) {{
            closeMultiFilters();
        }}
    }});

    document.addEventListener("keydown", event => {{
        if (event.key === "Escape") {{
            closeMultiFilters();
        }}
    }});
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

function filteredCoachingData() {{
    return coachingData.filter(r => {{
        const monthKey = coachingMonthKey(r["Coaching Date"]);
        return (
            (COACHING_FILTERS.emp.size === 0 || COACHING_FILTERS.emp.has(norm(r["Emp Name"]))) &&
            (COACHING_FILTERS.leader.size === 0 || COACHING_FILTERS.leader.has(norm(r["Coached by"]))) &&
            (COACHING_FILTERS.month.size === 0 || COACHING_FILTERS.month.has(monthKey))
        );
    }});
}}

function countBy(data, field) {{
    const out = {{}};
    data.forEach(r => {{
        const key = norm(r[field]) || "Blank";
        out[key] = (out[key] || 0) + 1;
    }});
    return Object.entries(out).map(([name, count]) => ({{name, count}})).sort((a,b) => b.count - a.count);
}}

function sortedCoachingRows(data) {{
    const column = COACHING_COLUMNS.find(c => c.field === coachingSortState.field) || {{}};
    const direction = coachingSortState.direction === "asc" ? 1 : -1;
    return [...data].sort((a, b) => {{
        let av = a[coachingSortState.field];
        let bv = b[coachingSortState.field];

        if (column.sortType === "date") {{
            av = parseDateValue(av)?.getTime() || 0;
            bv = parseDateValue(bv)?.getTime() || 0;
            return (av - bv) * direction;
        }}

        return norm(av).localeCompare(norm(bv), undefined, {{sensitivity: "base"}}) * direction;
    }});
}}

function confidenceScore(level) {{
    const value = norm(level).toUpperCase();
    if (value === "HIGH") return 100;
    if (value === "MEDIUM") return 65;
    if (value === "LOW") return 35;
    return 0;
}}

function coachingConfidenceAverage(data) {{
    if (data.length === 0) return 0;
    const total = data.reduce((sum, row) => sum + confidenceScore(row["Confidence Level"]), 0);
    return Math.round(total / data.length);
}}

function coachingSummaryPivot(data) {{
    const completedRows = data.filter(r => isCompletedStatus(r["Coaching Status"]));
    const weekKeys = [...new Set(completedRows.map(r => weekStartKey(r["Coaching Date"])).filter(Boolean))]
        .sort();
    const leaders = [...new Set(completedRows.map(r => norm(r["Coached by"]) || "Blank"))]
        .sort((a, b) => a.localeCompare(b, undefined, {{sensitivity: "base"}}));
    const counts = {{}};

    completedRows.forEach(r => {{
        const leader = norm(r["Coached by"]) || "Blank";
        const week = weekStartKey(r["Coaching Date"]);
        if (!week) return;
        const key = `${{leader}}|${{week}}`;
        counts[key] = (counts[key] || 0) + 1;
    }});

    const columns = [
        ...COACHING_SUMMARY_COLUMNS,
        ...weekKeys.map(week => ({{label: weekStartLabel(week), field: week, className: "nowrap"}})),
    ];
    const rows = leaders.map(leader => {{
        const row = {{"Team Leader": leader}};
        weekKeys.forEach(week => {{
            row[week] = counts[`${{leader}}|${{week}}`] || 0;
        }});
        return row;
    }});

    return {{columns, rows, completedCount: completedRows.length}};
}}

function setText(id, value) {{
    document.getElementById(id).textContent = Number(value).toLocaleString();
}}

function renderDataTable(id, rows, columns, sortState = null) {{
    let html = "<div class='table-scroll'><table><thead><tr>";
    columns.forEach(c => {{
        const sortableClass = c.sortable ? " class='sortable'" : "";
        const sortField = c.sortable ? ` data-sort-field="${{escapeHtml(c.field)}}"` : "";
        const active = sortState && sortState.field === c.field;
        const indicator = c.sortable ? `<span class="sort-indicator">${{active ? (sortState.direction === "asc" ? "^" : "v") : ""}}</span>` : "";
        html += `<th${{sortableClass}}${{sortField}}>${{escapeHtml(c.label)}}${{indicator}}</th>`;
    }});
    html += "</tr></thead><tbody>";

    if (rows.length === 0) {{
        html += `<tr><td colspan="${{columns.length}}">No records found</td></tr>`;
    }} else {{
        rows.forEach(r => {{
            html += "<tr>";
            columns.forEach(c => {{
                const className = c.className ? ` class="${{escapeHtml(c.className)}}"` : "";
                html += `<td${{className}}>${{escapeHtml(r[c.field])}}</td>`;
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
        height: 280,
        margin: {{l: 10, r: 10, t: 45, b: 10}},
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

function segmentBar(id, title, data, yTitle) {{
    const rows = [...data].reverse();
    Plotly.newPlot(id, [{{
        x: rows.map(d => d.count),
        y: rows.map(d => d.name),
        type: "bar",
        orientation: "h",
        text: rows.map(d => d.count),
        textposition: "auto",
        marker: {{color: rows.map((_, i) => COLORS[i % COLORS.length])}},
        hovertemplate: "%{{y}}<br>Headcount: %{{x}}<extra></extra>",
    }}], {{
        title: {{text: title, font: {{color: "#004C97", size: 15}}}},
        height: 300,
        margin: {{l: 120, r: 20, t: 45, b: 35}},
        xaxis: {{title: "Headcount"}},
        yaxis: {{title: yTitle}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        font: {{family: "Arial", size: 10}}
    }}, {{responsive: true}});
}}

function accountTenureStack(data) {{
    const accounts = countBy(data, "LOB / Account").slice(0, 10).map(d => d.name);
    const rows = [...accounts].reverse();
    const counts = {{}};

    rows.forEach(account => {{
        counts[account] = Object.fromEntries(TENURE_GROUPS.map(group => [group.name, 0]));
    }});

    data.forEach(r => {{
        const account = norm(r["LOB / Account"]) || "Blank";
        if (!counts[account]) return;
        const groupName = tenureGroupName(parseDateValue(r["Hire Date"]), new Date());
        if (groupName) counts[account][groupName] += 1;
    }});

    const traces = TENURE_GROUPS.map((group, i) => ({{
        name: group.name,
        x: rows.map(account => counts[account][group.name]),
        y: rows,
        type: "bar",
        orientation: "h",
        marker: {{color: COLORS[i % COLORS.length]}},
        hovertemplate: "%{{y}}<br>" + group.name + ": %{{x}}<extra></extra>",
    }}));

    Plotly.newPlot("accountTenureStack", traces, {{
        title: {{text: "Tenure by Account", font: {{color: "#004C97", size: 15}}}},
        height: 340,
        margin: {{l: 135, r: 20, t: 45, b: 35}},
        barmode: "stack",
        xaxis: {{title: "Headcount"}},
        yaxis: {{title: "Account"}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        legend: {{orientation: "h", y: -0.25, font: {{size: 9}}}},
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

function coachingCategoryChart(data) {{
    donut("coachingCategoryDonut", "Coaching Category", countBy(data, "Coaching Category"));
}}

function gaugePoint(cx, cy, radius, angle) {{
    const rad = angle * Math.PI / 180;
    return {{
        x: cx + radius * Math.cos(rad),
        y: cy - radius * Math.sin(rad),
    }};
}}

function gaugeSegmentPath(cx, cy, outerRadius, innerRadius, startAngle, endAngle) {{
    const steps = 18;
    const points = [];
    for (let i = 0; i <= steps; i += 1) {{
        const angle = startAngle + ((endAngle - startAngle) * i / steps);
        points.push(gaugePoint(cx, cy, outerRadius, angle));
    }}
    for (let i = steps; i >= 0; i -= 1) {{
        const angle = startAngle + ((endAngle - startAngle) * i / steps);
        points.push(gaugePoint(cx, cy, innerRadius, angle));
    }}
    return `M ${{points.map(p => `${{p.x.toFixed(1)}} ${{p.y.toFixed(1)}}`).join(" L ")}} Z`;
}}

function gaugeNeedlePath(cx, cy, angle, length, baseWidth) {{
    const tip = gaugePoint(cx, cy, length, angle);
    const left = gaugePoint(cx, cy, baseWidth, angle - 90);
    const right = gaugePoint(cx, cy, baseWidth, angle + 90);
    return `M ${{tip.x.toFixed(1)}} ${{tip.y.toFixed(1)}} L ${{left.x.toFixed(1)}} ${{left.y.toFixed(1)}} L ${{right.x.toFixed(1)}} ${{right.y.toFixed(1)}} Z`;
}}

function confidenceBand(value) {{
    if (value >= 80) return "EXCELLENT";
    if (value >= 60) return "GOOD";
    if (value >= 40) return "FAIR";
    if (value >= 20) return "POOR";
    return "VERY POOR";
}}

function coachingConfidenceGauge(data) {{
    const value = coachingConfidenceAverage(data);
    const cx = 300;
    const cy = 260;
    const segments = [
        {{label: "VERY POOR", start: 180, end: 144, color: "#F4511E"}},
        {{label: "POOR", start: 144, end: 108, color: "#FDBA3B"}},
        {{label: "FAIR", start: 108, end: 72, color: "#DDE817"}},
        {{label: "GOOD", start: 72, end: 36, color: "#39B54A"}},
        {{label: "EXCELLENT", start: 36, end: 0, color: "#007A3D"}},
    ];
    const needleAngle = 180 - Math.max(0, Math.min(value, 100)) * 1.8;
    const segmentMarkup = segments.map(segment => {{
        const labelAngle = (segment.start + segment.end) / 2;
        const labelPoint = gaugePoint(cx, cy, 194, labelAngle);
        return `
            <path d="${{gaugeSegmentPath(cx, cy, 220, 178, segment.start, segment.end)}}" fill="#B8B8B8" stroke="white" stroke-width="1" />
            <path d="${{gaugeSegmentPath(cx, cy, 178, 105, segment.start, segment.end)}}" fill="${{segment.color}}" stroke="white" stroke-width="1" />
            <text class="gauge-label" x="${{labelPoint.x.toFixed(1)}}" y="${{labelPoint.y.toFixed(1)}}" transform="rotate(${{(90 - labelAngle).toFixed(1)}} ${{labelPoint.x.toFixed(1)}} ${{labelPoint.y.toFixed(1)}})">${{segment.label}}</text>
        `;
    }}).join("");

    document.getElementById("coachingConfidenceGauge").innerHTML = `
        <div class="score-gauge">
            <svg viewBox="0 0 600 320" role="img" aria-label="AI Confidence Level Detection ${{value}} percent">
                <text x="300" y="24" text-anchor="middle" style="font: 700 15px Arial; fill: #004C97;">AI Confidence Level Detection</text>
                ${{segmentMarkup}}
                <path d="${{gaugeNeedlePath(cx, cy, needleAngle, 150, 18)}}" fill="#050505" />
                <circle cx="${{cx}}" cy="${{cy}}" r="32" fill="#050505" />
                <circle cx="${{cx}}" cy="${{cy}}" r="18" fill="white" />
                <circle cx="${{cx}}" cy="${{cy}}" r="6" fill="#050505" />
                <text class="gauge-value" x="300" y="302" text-anchor="middle">${{value}}% - ${{confidenceBand(value)}}</text>
            </svg>
        </div>
    `;
}}

function masterlistTable(data) {{
    renderDataTable("masterlistTable", data, MASTERLIST_COLUMNS);
}}

function coachingTable(data) {{
    renderDataTable("coachingTable", sortedCoachingRows(data), COACHING_COLUMNS, coachingSortState);
    document.querySelectorAll("#coachingTable th.sortable").forEach(th => {{
        th.addEventListener("click", () => {{
            const field = th.dataset.sortField;
            if (coachingSortState.field === field) {{
                coachingSortState.direction = coachingSortState.direction === "asc" ? "desc" : "asc";
            }} else {{
                coachingSortState.field = field;
                coachingSortState.direction = field.includes("Date") ? "desc" : "asc";
            }}
            renderCoaching();
        }});
    }});
}}

function coachingExportRows() {{
    return sortedCoachingRows(filteredCoachingData()).map(row => {{
        const out = {{}};
        COACHING_COLUMNS.forEach(column => {{
            out[column.label] = norm(row[column.field]);
        }});
        return out;
    }});
}}

function coachingExportText(delimiter = "\\t") {{
    const headers = COACHING_COLUMNS.map(column => column.label);
    const rows = coachingExportRows();
    const cleanCell = value => norm(value).replace(/[\\t\\r\\n]+/g, " ");
    const formatCell = value => {{
        const cleaned = cleanCell(value);
        if (delimiter === "," && /[",]/.test(cleaned)) {{
            return `"${{cleaned.replaceAll('"', '""')}}"`;
        }}
        return cleaned;
    }};
    return [
        headers.join(delimiter),
        ...rows.map(row => headers.map(header => formatCell(row[header])).join(delimiter)),
    ].join("\\n");
}}

function downloadBlob(filename, content, type) {{
    const blob = new Blob([content], {{type}});
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}}

function coachingExportHtmlTable() {{
    const headers = COACHING_COLUMNS.map(column => column.label);
    const rows = coachingExportRows();
    const headerHtml = headers.map(header => `<th>${{escapeHtml(header)}}</th>`).join("");
    const rowHtml = rows.map(row => (
        `<tr>${{headers.map(header => `<td>${{escapeHtml(row[header])}}</td>`).join("")}}</tr>`
    )).join("");
    return `
        <html>
        <head><meta charset="UTF-8"></head>
        <body>
            <table>
                <thead><tr>${{headerHtml}}</tr></thead>
                <tbody>${{rowHtml}}</tbody>
            </table>
        </body>
        </html>
    `;
}}

function downloadCoachingExcel() {{
    const rows = coachingExportRows();
    const dateTag = new Date().toISOString().slice(0, 10);

    if (window.XLSX) {{
        const worksheet = XLSX.utils.json_to_sheet(rows);
        worksheet["!cols"] = COACHING_COLUMNS.map(column => ({{
            wch: column.className === "long-text" ? 42 : Math.max(14, column.label.length + 2),
        }}));
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Coaching Logs");
        XLSX.writeFile(workbook, `coaching_logs_${{dateTag}}.xlsx`);
    }} else {{
        downloadBlob(
            `coaching_logs_${{dateTag}}.xls`,
            coachingExportHtmlTable(),
            "application/vnd.ms-excel;charset=utf-8"
        );
    }}
}}

async function openCoachingSheets() {{
    const button = document.getElementById("openCoachingSheets");
    const originalText = button?.textContent || "Google Sheets";
    window.open("https://sheets.new", "_blank", "noopener");

    try {{
        await navigator.clipboard.writeText(coachingExportText("\\t"));
        if (button) {{
            button.textContent = "Copied for Sheets";
            setTimeout(() => {{
                button.textContent = originalText;
            }}, 1800);
        }}
    }} catch (error) {{
        downloadBlob(
            `coaching_logs_${{new Date().toISOString().slice(0, 10)}}.tsv`,
            coachingExportText("\\t"),
            "text/tab-separated-values;charset=utf-8"
        );
    }}
}}

function renderCoaching() {{
    const data = filteredCoachingData();
    const completed = data.filter(r => isCompletedStatus(r["Coaching Status"])).length;
    const pending = data.filter(r => isPendingStatus(r["Coaching Status"])).length;

    setText("coachingCount", data.length);
    setText("coachingCompleted", completed);
    setText("coachingPending", pending);

    const meta = document.getElementById("coachingLoadedMeta");
    if (meta) {{
        meta.textContent = `${{data.length.toLocaleString()}} of ${{coachingData.length.toLocaleString()}} records shown`;
    }}

    coachingCategoryChart(data);
    coachingConfidenceGauge(data);
    const summary = coachingSummaryPivot(data);
    const summaryMeta = document.getElementById("coachingSummaryMeta");
    if (summaryMeta) {{
        summaryMeta.textContent = `${{summary.completedCount.toLocaleString()}} completed coaching sessions`;
    }}
    renderDataTable("coachingSummaryTable", summary.rows, summary.columns);
    coachingTable(data);
}}

function recentMovementsTable() {{
    const rows = filteredMovementData().slice(-10).reverse().map(r => ({{
        ...r,
        "Date Initiated": formatDateOnly(r["Timestamp"]),
        "Process Status": norm(r["Processed"]),
        "Processed Date Only": formatDateOnly(r["Processed Date"]),
    }}));
    renderDataTable("recentMovements", rows, [
        {{label: "Employee Name", field: "Employee Name"}},
        {{label: "Movement Type", field: "Movement Type"}},
        {{label: "New Department", field: "New Department"}},
        {{label: "New Account", field: "New Account"}},
        {{label: "New Supervisor", field: "New Supervisor"}},
        {{label: "New Job Title", field: "New Job Title"}},
        {{label: "Date Initiated", field: "Date Initiated"}},
        {{label: "Effective Date", field: "Effective Date"}},
        {{label: "Process Status", field: "Process Status"}},
        {{label: "Processed Date", field: "Processed Date Only"}},
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
    donut("employmentClassDonut", "Employement Class", countBy(data, "Employement Class"));
    accountTenureStack(data);
    segmentBar("tenureSegmentation", "Tenure Segmentation", tenureCounts(data), "Tenure Group");
    segmentBar("ageGroupBar", "Age Group", ageGroupCounts(data), "Age Group");
    weeklyChart();
    masterlistTable(data);
    recentMovementsTable();
}}

function switchTab(tabName) {{
    const masterlistControls = document.getElementById("masterlistControls");
    const coachingControls = document.getElementById("coachingControls");
    const isMasterlist = tabName === "masterlist";
    const isCoaching = tabName === "coaching";

    document.querySelectorAll(".tab-button").forEach(button => {{
        const active = button.dataset.tab === tabName;
        button.classList.toggle("active", active);
        button.setAttribute("aria-selected", active ? "true" : "false");
    }});

    document.querySelectorAll(".tab-panel").forEach(panel => {{
        panel.classList.toggle("active", panel.dataset.tab === tabName);
    }});

    masterlistControls.classList.toggle("hidden", !isMasterlist);
    coachingControls.classList.toggle("hidden", !isCoaching);

    if (isMasterlist || isCoaching) {{
        setTimeout(() => window.dispatchEvent(new Event("resize")), 0);
    }}
}}

document.querySelectorAll(".tab-button").forEach(button => {{
    button.addEventListener("click", () => switchTab(button.dataset.tab));
}});

populateFilter("departmentFilter", "Department");
populateFilter("accountFilter", "LOB / Account");
populateFilter("managerFilter", "Manager");
populateFilter("supervisorFilter", "Immediate Supervisor");
populateCoachingFilters();
initMultiFilterBehavior();
document.getElementById("downloadCoachingExcel")?.addEventListener("click", downloadCoachingExcel);
document.getElementById("openCoachingSheets")?.addEventListener("click", openCoachingSheets);
renderCoaching();
render();
</script>

</body>
</html>
"""

    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")
    print(f"Dashboard generated: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

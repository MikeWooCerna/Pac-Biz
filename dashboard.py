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
COACHING_VALIDATION_ERRORS_FILE = COACHING_OUTPUT_FILE.parent / "coaching_validation_errors.csv"
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
        ("personal responsibility", 3, "personal responsibility"),
        ("attitude", 3, "attitude"),
        ("engagement", 2, "engagement"),
        ("engaged", 2, "engagement"),
        ("initiative", 3, "initiative"),
        ("proactive", 2, "proactive behavior"),
        ("coachability", 3, "coachability"),
        ("applying feedback", 3, "applying feedback"),
        ("apply the feedback", 3, "applying feedback"),
        ("responsiveness", 3, "responsiveness"),
        ("ownership", 3, "ownership"),
        ("workplace conduct", 3, "workplace conduct"),
        ("conduct", 2, "conduct"),
        ("attentive", 2, "attentiveness"),
        ("attentiveness", 2, "attentiveness"),
    ],
    "Communication": [
        ("notify", 3, "failure to notify"),
        ("notification", 2, "notification"),
        ("escalate", 3, "escalation"),
        ("escalation", 3, "escalation"),
        ("listening", 3, "listening skills"),
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
        ("on time", 4, "punctuality"),
        ("schedule change", 4, "schedule change"),
        ("schedule changes", 4, "schedule change"),
        ("schedule adherence", 4, "schedule adherence"),
        ("scheduled shift", 3, "shift compliance"),
        ("scheduled shifts", 3, "shift compliance"),
        ("work-hour", 4, "work-hour adherence"),
        ("work hour", 4, "work-hour adherence"),
        ("time management", 4, "time management"),
        ("availability", 4, "availability"),
        ("reliability", 4, "reliability"),
        ("reliable", 3, "reliability"),
        ("time adjustment", 4, "time adjustment"),
        ("time adjustments", 4, "time adjustment"),
        ("backup alarm", 4, "backup alarm"),
        ("backup alarms", 4, "backup alarm"),
        ("alarm", 3, "alarm"),
        ("alarms", 3, "alarm"),
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
        ("booking", 4, "booking procedure"),
        ("verification", 4, "verification step"),
        ("verify", 3, "verification step"),
        ("required process", 4, "required process"),
        ("operational process", 4, "operational process"),
        ("process", 2, "process"),
    ],
    "Performance & Quality": [
        ("qa score", 4, "QA score"),
        ("quality", 3, "quality standard"),
        ("accuracy", 4, "accuracy"),
        ("productivity", 4, "productivity"),
        ("efficiency", 4, "efficiency"),
        ("error rate", 4, "error rate"),
        ("output quality", 4, "output quality"),
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
        ("skill development", 4, "skill development"),
        ("deficiency", 3, "knowledge deficiency"),
        ("deficiencies", 3, "knowledge deficiency"),
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
        ("disciplinary", 4, "disciplinary concern"),
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
    "Category Status",
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
        return "Other / Review Required", "Low", "Clear cues: insufficient information."

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
        return "Other / Review Required", "Low", "Clear cues: no reliable match."

    matches.sort(key=lambda item: (-item[1], item[0]))
    top_category, top_score, top_labels = matches[0]
    second_category = matches[1][0] if len(matches) > 1 else ""
    second_score = matches[1][1] if len(matches) > 1 else 0

    if top_score <= 1:
        return "Other / Review Required", "Low", "Clear cues: weak indicators."

    total_score = sum(item[1] for item in matches)
    dominant_share = top_score / total_score if total_score else 0

    if dominant_share >= 0.60 or (top_score >= 4 and top_score >= second_score + 2) or (top_score >= 3 and second_score == 0):
        confidence = "High"
        reason = f"Clear cues: {join_reason_labels(top_labels)}."
    else:
        confidence = "Medium"
        reason = f"Clear cues: {join_reason_labels(top_labels)}."

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


def coaching_email_error(label, email, email_lookup):
    email_text = cell_text(email)
    normalized_email = normalize_email(email)

    if not normalized_email:
        return f"{label} email missing"
    if normalized_email not in email_lookup:
        return f"{label} not found for email: {email_text}"
    return ""


def write_coaching_validation_errors(errors):
    try:
        if errors:
            COACHING_VALIDATION_ERRORS_FILE.parent.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(errors).to_csv(COACHING_VALIDATION_ERRORS_FILE, index=False)
        elif COACHING_VALIDATION_ERRORS_FILE.exists():
            COACHING_VALIDATION_ERRORS_FILE.unlink()
    except OSError as exc:
        print(f"Unable to update Coaching validation error file: {exc}")


def print_coaching_validation_alerts(errors):
    if not errors:
        print("Coaching email validation: No errors found.")
        return

    print("")
    print("========================================")
    print("COACHING EMAIL VALIDATION ALERTS")
    print("========================================")
    print(f"{len(errors)} task(s) excluded from the Coaching report until corrected.")
    print("")
    for item in errors:
        print(f"Task GID: {item['Task GID']}")
        print(f"Error: {item['Error']}")
        print(f"Employee Email: {item['Employee Email'] or '(blank)'}")
        print(f"Supervisor Email: {item['Supervisor Email'] or '(blank)'}")
        print("")
    print(f"Saved validation details: {COACHING_VALIDATION_ERRORS_FILE}")
    print("========================================")


def parse_coaching_date_for_status(value):
    text = cell_text(value)
    if not text:
        return pd.NaT
    first_part = text.split(" ")[0]
    return pd.to_datetime(first_part, errors="coerce")


def assign_category_status(coaching):
    if coaching.empty:
        return coaching

    result = coaching.copy()
    if "Category Status" not in result.columns:
        result["Category Status"] = ""

    required = {"Emp Name", "Coaching Category", "Coaching Date"}
    if not required.issubset(result.columns):
        return result

    work = result[["Emp Name", "Coaching Category", "Coaching Date"]].copy()
    work["_date"] = work["Coaching Date"].apply(parse_coaching_date_for_status)
    work["_index"] = work.index
    work["_month"] = work["_date"].dt.to_period("M")
    work["_emp_key"] = work["Emp Name"].map(lambda value: cell_text(value).casefold())
    work["_category_key"] = work["Coaching Category"].map(lambda value: cell_text(value).casefold())
    work = work.sort_values(
        ["_emp_key", "_category_key", "_date", "_index"],
        kind="mergesort",
        na_position="last",
    )

    seen_any = set()
    seen_month = set()
    statuses = {}

    for _, row in work.iterrows():
        key = (row["_emp_key"], row["_category_key"])
        month_key = (row["_emp_key"], row["_category_key"], str(row["_month"]) if pd.notna(row["_month"]) else "")

        if key in seen_any:
            statuses[row["_index"]] = "Recurring This Month" if month_key in seen_month else "Historical Repeat"
        else:
            statuses[row["_index"]] = "New"

        seen_any.add(key)
        if pd.notna(row["_month"]):
            seen_month.add(month_key)

    result["Category Status"] = result.index.map(lambda idx: statuses.get(idx, "New"))
    return result


def transform_coaching_logs(source, masterlist):
    if source.empty:
        return pd.DataFrame(columns=COACHING_OUTPUT_COLUMNS)

    if all(column in source.columns for column in COACHING_OUTPUT_COLUMNS):
        return assign_category_status(source)[COACHING_OUTPUT_COLUMNS]

    email_lookup = build_email_name_lookup(masterlist)
    records = []
    validation_errors = []

    for _, row in source.iterrows():
        coaching_id = cell_text(row.get("Task GID"))
        if coaching_id in EXCLUDED_COACHING_GIDS:
            continue

        employee_email = row.get("Employee Email")
        supervisor_email = row.get("Supervisor Email")
        row_errors = [
            error for error in [
                coaching_email_error("Employee", employee_email, email_lookup),
                coaching_email_error("Supervisor", supervisor_email, email_lookup),
            ]
            if error
        ]

        if row_errors:
            validation_errors.append({
                "Task GID": coaching_id or "(blank)",
                "Error": "; ".join(row_errors),
                "Employee Email": cell_text(employee_email),
                "Supervisor Email": cell_text(supervisor_email),
                "Task Name": cell_text(row.get("Task Name")),
            })
            continue

        details = cell_text(row.get("Agreed Action Steps"))
        category, confidence, reason = classify_coaching_details(details)

        records.append({
            "Coaching ID": coaching_id,
            "Emp Name": resolve_name(
                email_lookup,
                employee_email,
                row.get("Employee Name"),
            ),
            "Coached by": resolve_name(email_lookup, supervisor_email),
            "Coaching Details": details,
            "Remarks/Comment": cell_text(row.get("Improvement Timeline / Follow-Up Date")),
            "Coaching Category": category,
            "Category Status": "",
            "Confidence Level": confidence,
            "Reason": reason,
            "Coaching Date": cell_text(row.get("Coaching Date")),
            "Coaching Status": cell_text(row.get("Status")),
            "Created Date": cell_text(row.get("Created Date")),
            "Created Time": cell_text(row.get("Created Time")),
        })

    write_coaching_validation_errors(validation_errors)
    print_coaching_validation_alerts(validation_errors)
    return clean_columns(assign_category_status(pd.DataFrame(records, columns=COACHING_OUTPUT_COLUMNS)))


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
        grid-template-columns: repeat(8, minmax(0, 1fr));
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
        grid-template-columns: repeat(6, minmax(0, 1fr));
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
        grid-template-columns: repeat(3, minmax(0, 1fr));
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
        scrollbar-gutter: stable;
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
        align-items: center;
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

    .table-action.icon-action {{
        width: 32px;
        min-width: 32px;
        padding: 0;
        background: white;
        color: var(--blue);
        border-color: var(--blue);
    }}

    .logs-focus-icon {{
        position: relative;
        display: inline-block;
        width: 15px;
        height: 15px;
    }}

    .logs-focus-icon::before,
    .logs-focus-icon::after {{
        content: "";
        position: absolute;
        width: 6px;
        height: 6px;
        border-color: currentColor;
        border-style: solid;
    }}

    .logs-focus-icon::before {{
        top: 0;
        right: 0;
        border-width: 2px 2px 0 0;
    }}

    .logs-focus-icon::after {{
        left: 0;
        bottom: 0;
        border-width: 0 0 2px 2px;
    }}

    .logs-focus-mode .logs-focus-icon::before {{
        top: 2px;
        right: 2px;
        border-width: 0 0 2px 2px;
    }}

    .logs-focus-mode .logs-focus-icon::after {{
        left: 2px;
        bottom: 2px;
        border-width: 2px 2px 0 0;
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
        padding: 6px 8px;
        text-align: left;
        position: sticky;
        top: 0;
        z-index: 1;
        vertical-align: middle;
        line-height: 1.2;
    }}

    th.sortable {{
        cursor: pointer;
        user-select: none;
    }}

    .sort-indicator {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 12px;
        width: 12px;
        margin-left: 5px;
        opacity: .9;
        flex: 0 0 12px;
        line-height: 1;
    }}

    .th-content {{
        display: inline-flex;
        align-items: center;
        gap: 4px;
        max-width: 100%;
        white-space: nowrap;
    }}

    .th-label {{
        min-width: 0;
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

    #coachingTable table {{
        table-layout: fixed;
        min-width: 1680px;
    }}

    #coachingTable td {{
        line-height: 1.35;
    }}

    #coachingTable .long-text {{
        min-width: 0;
        max-width: none;
    }}

    #coachingTable .coaching-detail-col {{
        width: 300px;
    }}

    #coachingTable .remarks-col {{
        width: 260px;
    }}

    #coachingTable .reason-col {{
        width: 210px;
    }}

    #coachingTable .compact-col {{
        width: 116px;
        white-space: nowrap;
    }}

    #coachingTable .status-col {{
        width: 142px;
        white-space: nowrap;
    }}

    #coachingTable .name-col {{
        width: 170px;
        white-space: nowrap;
    }}

    #coachingTable .category-col {{
        width: 170px;
        white-space: nowrap;
    }}

    #coachingTable .category-status-col {{
        width: 170px;
        white-space: nowrap;
    }}

    #coachingTable td.category-status-new {{
        background: #DCFCE7;
        color: #166534;
        font-weight: 900;
    }}

    #coachingTable td.category-status-recurring {{
        background: #FFEDD5;
        color: #9A3412;
        font-weight: 900;
    }}

    #coachingTable td.category-status-historical {{
        background: #DBEAFE;
        color: #1E3A8A;
        font-weight: 900;
    }}

    #masterlistTable td.emp-status-active {{
        color: var(--green);
        font-weight: 700;
    }}

    #masterlistTable td.emp-status-inactive {{
        color: #E53935;
        font-weight: 700;
    }}

    #coachingTable .id-col {{
        width: 135px;
        white-space: nowrap;
    }}

    #coachingLogsCard {{
        display: flex;
        flex-direction: column;
    }}

    #coachingLogsCard #coachingTable {{
        min-height: 0;
    }}

    body.logs-focus-mode #summarySection,
    body.logs-focus-mode #summarySpacer {{
        display: none;
    }}

    body.logs-focus-mode #coachingLogsCard {{
        min-height: 0;
        height: calc(100vh - 330px);
    }}

    body.logs-focus-mode #coachingLogsCard .table-scroll {{
        max-height: none;
        height: 100%;
    }}

    body.logs-focus-mode #coachingTable {{
        flex: 1 1 auto;
        min-height: 0;
    }}

    body.logs-focus-mode #coachingPanel .coaching-grid {{
        grid-template-rows: auto 1fr;
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
        height: 320px;
        display: grid;
        align-items: center;
        justify-items: center;
        padding: 0;
    }}

    .score-gauge svg {{
        width: min(100%, 600px);
        height: 320px;
        overflow: visible;
    }}

    .gauge-label {{
        font-family: Arial, sans-serif;
        font-size: 11.5px;
        font-weight: 600;
        fill: #111827;
        dominant-baseline: middle;
        text-anchor: middle;
    }}

    .gauge-value {{
        font-size: 24px;
        font-weight: 900;
        fill: var(--dark-blue);
    }}

    .chart-summary {{
        display: grid;
        gap: 6px;
        margin: 0 4px 4px;
        padding: 8px 10px 2px;
        font-size: 12px;
        color: var(--text);
    }}

    .coaching-donut-widget {{
        display: flex;
        flex-direction: column;
        min-height: 410px;
    }}

    .coaching-donut-plot {{
        flex: 0 0 280px;
    }}

    .chart-summary-rows {{
        display: grid;
        align-content: start;
        gap: 6px;
        min-height: 112px;
    }}

    .chart-summary-row {{
        display: grid;
        grid-template-columns: 12px minmax(0, 1fr) auto;
        align-items: center;
        gap: 8px;
    }}

    .chart-summary-swatch {{
        width: 10px;
        height: 10px;
        border-radius: 999px;
    }}

    .chart-summary-name {{
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        font-weight: 700;
    }}

    .chart-summary-value {{
        color: var(--dark-blue);
        font-weight: 900;
        white-space: nowrap;
    }}

    .chart-summary-total {{
        margin-top: 2px;
        padding-top: 6px;
        border-top: 1px solid #E5E7EB;
        color: var(--muted);
        font-weight: 900;
        min-height: 28px;
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
        <label>Employee Name</label>
        <details class="multi-filter" id="empNameFilter">
            <summary id="empNameFilterSummary">All</summary>
            <div class="multi-options" id="empNameOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Department</label>
        <details class="multi-filter" id="departmentFilter">
            <summary id="departmentFilterSummary">All</summary>
            <div class="multi-options" id="departmentOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>LOB / Account</label>
        <details class="multi-filter" id="accountFilter">
            <summary id="accountFilterSummary">All</summary>
            <div class="multi-options" id="accountOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Supervisor</label>
        <details class="multi-filter" id="supervisorFilter">
            <summary id="supervisorFilterSummary">All</summary>
            <div class="multi-options" id="supervisorOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Manager</label>
        <details class="multi-filter" id="managerFilter">
            <summary id="managerFilterSummary">All</summary>
            <div class="multi-options" id="managerOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Tenure</label>
        <details class="multi-filter" id="tenureFilter">
            <summary id="tenureFilterSummary">All</summary>
            <div class="multi-options" id="tenureOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Employee Group</label>
        <details class="multi-filter" id="employeeGroupFilter">
            <summary id="employeeGroupFilterSummary">All</summary>
            <div class="multi-options" id="employeeGroupOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Employment Status</label>
        <details class="multi-filter" id="employmentStatusFilter">
            <summary id="employmentStatusFilterSummary">All</summary>
            <div class="multi-options" id="employmentStatusOptions"></div>
        </details>
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
        <label>Coaching Status</label>
        <details class="multi-filter" id="coachingStatusFilter">
            <summary id="coachingStatusFilterSummary">All</summary>
            <div class="multi-options" id="coachingStatusOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Coaching Category</label>
        <details class="multi-filter" id="coachingCategoryFilter">
            <summary id="coachingCategoryFilterSummary">All</summary>
            <div class="multi-options" id="coachingCategoryOptions"></div>
        </details>
    </div>
    <div class="filter-box">
        <label>Category Status</label>
        <details class="multi-filter" id="coachingCategoryStatusFilter">
            <summary id="coachingCategoryStatusFilterSummary">All</summary>
            <div class="multi-options" id="coachingCategoryStatusOptions"></div>
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
        <div class="chart-card"><div id="accountBar"></div></div>
        <div class="chart-card"><div id="managerBar"></div></div>
        <div class="chart-card"><div id="supervisorBar"></div></div>
    </div>
    <div class="chart-card"><div id="accountTenureStack"></div></div>
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
        <div class="chart-card"><div id="coachingStatusDonut"></div></div>
        <div class="chart-card"><div id="coachingConfidenceGauge"></div></div>
    </div>
    <div class="chart-card coaching-summary-card" id="summarySection">
        <div class="table-heading">
            <h3>Summary</h3>
            <div class="table-actions">
                <span class="table-meta" id="coachingSummaryMeta">0 completed coaching sessions</span>
            </div>
        </div>
        <div id="coachingSummaryTable"></div>
        <div class="disclaimer">Coaching will be only be counted once status is <span class="completed-word">Completed</span></div>
    </div>
    <div class="empty-widget-space" id="summarySpacer"></div>
    <div class="chart-card full" id="coachingLogsCard">
        <div class="table-heading">
            <h3>Coaching Logs</h3>
            <div class="table-actions">
                <button class="table-action" type="button" id="downloadCoachingExcel">Download Excel</button>
                <button class="table-action secondary" type="button" id="openCoachingSheets">Copy to Sheets</button>
                <button class="table-action icon-action" type="button" id="logsFocusToggle" aria-label="Expand Coaching Logs" title="Expand Coaching Logs">
                    <span class="logs-focus-icon" aria-hidden="true"></span>
                </button>
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
const COACHING_CATEGORY_COLORS = {{
    "Attendance & Adherence": "#0057B8",
    "Process Compliance": "#43A047",
    "Policy Compliance": "#FB8C00",
    "Communication": "#8E24AA",
    "Behavioral": "#E53935",
}};
const COACHING_STATUS_COLORS = {{
    "Pending": "#FB8C00",
    "Completed": "#43A047",
}};
const MASTERLIST_COLUMNS = [
    {{label: "Employee ID", field: "ID No.", sortable: true, sortType: "number"}},
    {{label: "Employee Name", field: "Emp Name", sortable: true}},
    {{label: "Employment Status", field: "Employment Status", sortable: true}},
    {{label: "Hire Date", field: "Hire Date", sortable: true, sortType: "date"}},
    {{label: "Employement Class", field: "Employement Class", sortable: true}},
    {{label: "Tenure", field: "Tenure", sortable: true, sortField: "__TenureDays", sortType: "number"}},
    {{label: "Job Title", field: "Job Title", sortable: true}},
    {{label: "Employee Group", field: "Employee Group", sortable: true}},
    {{label: "Department", field: "Department", sortable: true}},
    {{label: "LOB/Account", field: "LOB / Account", sortable: true}},
    {{label: "Immediate Supervisor", field: "Immediate Supervisor", sortable: true}},
    {{label: "Manager", field: "Manager", sortable: true}},
    {{label: "Email", field: "Company Email", sortable: true}},
];
const COACHING_COLUMNS = [
    {{label: "Coaching ID", field: "Coaching ID", className: "id-col nowrap"}},
    {{label: "Emp Name", field: "Emp Name", className: "name-col nowrap", sortable: true}},
    {{label: "Coached by", field: "Coached by", className: "name-col nowrap", sortable: true}},
    {{label: "Coaching Details", field: "Coaching Details", className: "long-text coaching-detail-col"}},
    {{label: "Remarks/Comment", field: "Remarks/Comment", className: "long-text remarks-col"}},
    {{label: "Coaching Category", field: "Coaching Category", className: "category-col nowrap", sortable: true}},
    {{label: "Category Status", field: "Category Status", className: "category-status-col nowrap", sortable: true}},
    {{label: "Confidence Level", field: "Confidence Level", className: "compact-col nowrap"}},
    {{label: "Reason", field: "Reason", className: "long-text reason-col"}},
    {{label: "Coaching Date", field: "Coaching Date", className: "compact-col nowrap", sortable: true, sortType: "date"}},
    {{label: "Coaching Status", field: "Coaching Status", className: "status-col nowrap"}},
    {{label: "Created Date", field: "Created Date", className: "compact-col nowrap", sortable: true, sortType: "date"}},
    {{label: "Created Time", field: "Created Time", className: "compact-col nowrap"}},
];
const COACHING_SUMMARY_COLUMNS = [
    {{label: "Team Leader", field: "Team Leader", className: "nowrap"}},
];
const RECENT_MOVEMENT_COLUMNS = [
    {{label: "Employee Name", field: "Employee Name", sortable: true}},
    {{label: "Movement Type", field: "Movement Type", sortable: true}},
    {{label: "New Department", field: "New Department", sortable: true}},
    {{label: "New Account", field: "New Account", sortable: true}},
    {{label: "New Supervisor", field: "New Supervisor", sortable: true}},
    {{label: "New Job Title", field: "New Job Title", sortable: true}},
    {{label: "Date Initiated", field: "Date Initiated", sortable: true, sortType: "date"}},
    {{label: "Effective Date", field: "Effective Date", sortable: true, sortType: "date"}},
    {{label: "Process Status", field: "Process Status", sortable: true}},
    {{label: "Processed Date", field: "Processed Date Only", sortable: true, sortType: "date"}},
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
const MASTERLIST_FILTERS = {{
    empName: new Set(),
    department: new Set(),
    account: new Set(),
    supervisor: new Set(),
    manager: new Set(),
    tenure: new Set(),
    employeeGroup: new Set(),
    employmentStatus: new Set(),
}};
const COACHING_FILTERS = {{
    emp: new Set(),
    leader: new Set(),
    status: new Set(),
    category: new Set(),
    categoryStatus: new Set(),
    month: new Set(),
}};
const masterlistSortState = {{
    field: "Emp Name",
    direction: "asc",
}};
const movementSortState = {{
    field: "Date Initiated",
    direction: "desc",
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

function formatTenureDisplay(hireDate) {{
    if (!hireDate) return "";
    const now = new Date();
    const days = wholeDayDiff(hireDate, now);
    if (days < 0) return "";
    if (days < 31) return `${{days}} Day${{days !== 1 ? "s" : ""}}`;
    let years = now.getFullYear() - hireDate.getFullYear();
    let months = now.getMonth() - hireDate.getMonth();
    if (months < 0) {{ years--; months += 12; }}
    if (years === 0) return `${{months}} Month${{months !== 1 ? "s" : ""}}`;
    if (months === 0) return `${{years}} Year${{years !== 1 ? "s" : ""}}`;
    return `${{years}} Year${{years !== 1 ? "s" : ""}}, ${{months}} Month${{months !== 1 ? "s" : ""}}`;
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

const NONE_SELECTED = "__NONE_SELECTED__";

function setMultiSummary(summaryId, selectedSet, labelFormatter = value => value) {{
    const summary = document.getElementById(summaryId);
    if (!summary) return;
    if (selectedSet.has(NONE_SELECTED)) {{
        summary.textContent = "None";
    }} else if (selectedSet.size === 0) {{
        summary.textContent = "All";
    }} else if (selectedSet.size === 1) {{
        const value = [...selectedSet][0];
        summary.textContent = labelFormatter(value);
    }} else {{
        summary.textContent = `${{selectedSet.size}} selected`;
    }}
}}

function syncMultiFilterOptions(optionsId, selectedSet, allValues) {{
    const box = document.getElementById(optionsId);
    const selectAll = box?.querySelector("input[data-select-all='true']");
    if (!selectAll) return;

    if (!selectedSet.has(NONE_SELECTED) && selectedSet.size === allValues.length) {{
        selectedSet.clear();
    }}

    const noneSelected = selectedSet.has(NONE_SELECTED);
    const allSelected = !noneSelected && selectedSet.size === 0;
    selectAll.checked = allSelected;
    selectAll.indeterminate = !noneSelected && !allSelected && selectedSet.size < allValues.length;
    box.querySelectorAll("input[type='checkbox']:not([data-select-all='true'])").forEach(checkbox => {{
        checkbox.checked = !noneSelected && (allSelected || selectedSet.has(checkbox.value));
    }});
}}

function populateMultiFilter(optionsId, summaryId, values, selectedSet, onChange, labelFormatter = value => value) {{
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
        if (!selectAllCheckbox.checked) {{
            selectedSet.add(NONE_SELECTED);
        }}
        syncMultiFilterOptions(optionsId, selectedSet, allValues);
        setMultiSummary(summaryId, selectedSet, labelFormatter);
        onChange();
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
        checkbox.checked = !selectedSet.has(NONE_SELECTED) && (selectedSet.size === 0 || selectedSet.has(value));
        checkbox.addEventListener("change", () => {{
            const wasAllSelected = !selectedSet.has(NONE_SELECTED) && selectedSet.size === 0;
            if (selectedSet.has(NONE_SELECTED)) {{
                selectedSet.clear();
            }}
            if (wasAllSelected && !checkbox.checked) {{
                allValues.forEach(v => selectedSet.add(v));
            }}
            if (checkbox.checked) {{
                selectedSet.add(value);
            }} else {{
                selectedSet.delete(value);
                if (selectedSet.size === 0) {{
                    selectedSet.add(NONE_SELECTED);
                }}
            }}
            syncMultiFilterOptions(optionsId, selectedSet, allValues);
            setMultiSummary(summaryId, selectedSet, labelFormatter);
            onChange();
        }});

        const text = document.createElement("span");
        text.textContent = label;

        option.appendChild(checkbox);
        option.appendChild(text);
        box.appendChild(option);
    }});

    syncMultiFilterOptions(optionsId, selectedSet, allValues);
    setMultiSummary(summaryId, selectedSet, labelFormatter);
}}

function populateCoachingFilters() {{
    populateMultiFilter(
        "coachingEmpOptions",
        "coachingEmpFilterSummary",
        uniqueValues(coachingData, "Emp Name"),
        COACHING_FILTERS.emp,
        renderCoaching
    );
    populateMultiFilter(
        "coachingLeaderOptions",
        "coachingLeaderFilterSummary",
        uniqueValues(coachingData, "Coached by"),
        COACHING_FILTERS.leader,
        renderCoaching
    );
    populateMultiFilter(
        "coachingStatusOptions",
        "coachingStatusFilterSummary",
        uniqueValues(coachingData, "Coaching Status"),
        COACHING_FILTERS.status,
        renderCoaching
    );
    populateMultiFilter(
        "coachingCategoryOptions",
        "coachingCategoryFilterSummary",
        uniqueValues(coachingData, "Coaching Category"),
        COACHING_FILTERS.category,
        renderCoaching
    );
    populateMultiFilter(
        "coachingCategoryStatusOptions",
        "coachingCategoryStatusFilterSummary",
        ["New", "Recurring This Month", "Historical Repeat"].map(v => ({{value: v, label: v}})),
        COACHING_FILTERS.categoryStatus,
        renderCoaching
    );

    const monthValues = [...new Set(coachingData.map(r => coachingMonthKey(r["Coaching Date"])).filter(Boolean))]
        .sort()
        .map(value => ({{value, label: coachingMonthLabelFromKey(value)}}));
    populateMultiFilter(
        "coachingMonthOptions",
        "coachingMonthFilterSummary",
        monthValues,
        COACHING_FILTERS.month,
        renderCoaching,
        coachingMonthLabelFromKey
    );
}}

function populateMasterlistFilters() {{
    populateMultiFilter(
        "empNameOptions",
        "empNameFilterSummary",
        uniqueValues(masterlist, "Emp Name"),
        MASTERLIST_FILTERS.empName,
        render
    );
    populateMultiFilter(
        "departmentOptions",
        "departmentFilterSummary",
        uniqueValues(masterlist, "Department"),
        MASTERLIST_FILTERS.department,
        render
    );
    populateMultiFilter(
        "accountOptions",
        "accountFilterSummary",
        uniqueValues(masterlist, "LOB / Account"),
        MASTERLIST_FILTERS.account,
        render
    );
    populateMultiFilter(
        "supervisorOptions",
        "supervisorFilterSummary",
        uniqueValues(masterlist, "Immediate Supervisor"),
        MASTERLIST_FILTERS.supervisor,
        render
    );
    populateMultiFilter(
        "managerOptions",
        "managerFilterSummary",
        uniqueValues(masterlist, "Manager"),
        MASTERLIST_FILTERS.manager,
        render
    );
    populateMultiFilter(
        "tenureOptions",
        "tenureFilterSummary",
        TENURE_GROUPS.map(g => ({{value: g.name, label: g.name}})),
        MASTERLIST_FILTERS.tenure,
        render
    );
    populateMultiFilter(
        "employeeGroupOptions",
        "employeeGroupFilterSummary",
        uniqueValues(masterlist, "Employee Group"),
        MASTERLIST_FILTERS.employeeGroup,
        render
    );
    populateMultiFilter(
        "employmentStatusOptions",
        "employmentStatusFilterSummary",
        ["Active", "Inactive"].map(v => ({{value: v, label: v}})),
        MASTERLIST_FILTERS.employmentStatus,
        render
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
    return masterlist.filter(r => {{
        const tenureGroup = tenureGroupName(parseDateValue(r["Hire Date"]), new Date());
        return (
            filterMatches(MASTERLIST_FILTERS.department, norm(r["Department"])) &&
            filterMatches(MASTERLIST_FILTERS.account, norm(r["LOB / Account"])) &&
            filterMatches(MASTERLIST_FILTERS.supervisor, norm(r["Immediate Supervisor"])) &&
            filterMatches(MASTERLIST_FILTERS.manager, norm(r["Manager"])) &&
            filterMatches(MASTERLIST_FILTERS.tenure, tenureGroup) &&
            filterMatches(MASTERLIST_FILTERS.employeeGroup, norm(r["Employee Group"])) &&
            filterMatches(MASTERLIST_FILTERS.employmentStatus, norm(r["Employment Status"]))
        );
    }});
}}

function filteredCoachingData() {{
    return coachingData.filter(r => {{
        const monthKey = coachingMonthKey(r["Coaching Date"]);
        return (
            filterMatches(COACHING_FILTERS.emp, norm(r["Emp Name"])) &&
            filterMatches(COACHING_FILTERS.leader, norm(r["Coached by"])) &&
            filterMatches(COACHING_FILTERS.status, norm(r["Coaching Status"])) &&
            filterMatches(COACHING_FILTERS.category, norm(r["Coaching Category"])) &&
            filterMatches(COACHING_FILTERS.categoryStatus, norm(r["Category Status"])) &&
            filterMatches(COACHING_FILTERS.month, monthKey)
        );
    }});
}}

function filterMatches(selectedSet, value) {{
    if (selectedSet.has(NONE_SELECTED)) return false;
    return selectedSet.size === 0 || selectedSet.has(value);
}}

function countBy(data, field) {{
    const out = {{}};
    data.forEach(r => {{
        const key = norm(r[field]) || "Blank";
        out[key] = (out[key] || 0) + 1;
    }});
    return Object.entries(out).map(([name, count]) => ({{name, count}})).sort((a,b) => b.count - a.count);
}}

function sortedTableRows(data, columns, sortState) {{
    const column = columns.find(c => c.field === sortState.field) || {{}};
    const direction = sortState.direction === "asc" ? 1 : -1;
    return [...data].sort((a, b) => {{
        let av = a[sortState.field];
        let bv = b[sortState.field];

        if (column.sortType === "date") {{
            av = parseDateValue(av)?.getTime() || 0;
            bv = parseDateValue(bv)?.getTime() || 0;
            return (av - bv) * direction;
        }}

        if (column.sortType === "number") {{
            av = Number(norm(av).replace(/,/g, ""));
            bv = Number(norm(bv).replace(/,/g, ""));
            av = Number.isFinite(av) ? av : 0;
            bv = Number.isFinite(bv) ? bv : 0;
            return (av - bv) * direction;
        }}

        return norm(av).localeCompare(norm(bv), undefined, {{sensitivity: "base"}}) * direction;
    }});
}}

function updateSortState(sortState, columns, field) {{
    const column = columns.find(c => c.field === field) || {{}};
    if (sortState.field === field) {{
        sortState.direction = sortState.direction === "asc" ? "desc" : "asc";
    }} else {{
        sortState.field = field;
        sortState.direction = column.sortType === "date" || column.sortType === "number" ? "desc" : "asc";
    }}
}}

function bindTableSorting(tableId, columns, sortState, renderFn) {{
    document.querySelectorAll(`#${{tableId}} th.sortable`).forEach(th => {{
        th.addEventListener("click", () => {{
            updateSortState(sortState, columns, th.dataset.sortField);
            renderFn();
        }});
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

function employmentStatusClass(value) {{
    const v = norm(value).toUpperCase();
    if (v === "ACTIVE") return "emp-status-active";
    if (v === "INACTIVE") return "emp-status-inactive";
    return "";
}}

function categoryStatusClass(value) {{
    const status = norm(value).toLowerCase();
    if (status === "new") return "category-status-new";
    if (status === "recurring this month") return "category-status-recurring";
    if (status === "historical repeat") return "category-status-historical";
    return "";
}}

function renderDataTable(id, rows, columns, sortState = null) {{
    let html = "<div class='table-scroll'><table><thead><tr>";
    columns.forEach(c => {{
        const thClasses = [];
        if (c.sortable) thClasses.push("sortable");
        if (c.className) thClasses.push(c.className);
        const classAttr = thClasses.length ? ` class="${{escapeHtml(thClasses.join(" "))}}"` : "";
        const sortField = c.sortable ? ` data-sort-field="${{escapeHtml(c.sortField || c.field)}}"` : "";
        const active = sortState && sortState.field === (c.sortField || c.field);
        const indicator = c.sortable ? `<span class="sort-indicator">${{active ? (sortState.direction === "asc" ? "^" : "v") : ""}}</span>` : "";
        html += `<th${{classAttr}}${{sortField}}><span class="th-content"><span class="th-label">${{escapeHtml(c.label)}}</span>${{indicator}}</span></th>`;
    }});
    html += "</tr></thead><tbody>";

    if (rows.length === 0) {{
        html += `<tr><td colspan="${{columns.length}}">No records found</td></tr>`;
    }} else {{
        rows.forEach(r => {{
            html += "<tr>";
            columns.forEach(c => {{
                const classes = [];
                if (c.className) classes.push(c.className);
                if (c.field === "Category Status") classes.push(categoryStatusClass(r[c.field]));
                if (c.field === "Employment Status") classes.push(employmentStatusClass(r[c.field]));
                const className = classes.filter(Boolean).length ? ` class="${{escapeHtml(classes.filter(Boolean).join(" "))}}"` : "";
                html += `<td${{className}}>${{escapeHtml(r[c.field])}}</td>`;
            }});
            html += "</tr>";
        }});
    }}

    html += "</tbody></table></div>";
    document.getElementById(id).innerHTML = html;
}}

function donut(id, title, data, textInfo = "percent", colors = COLORS) {{
    Plotly.newPlot(id, [{{
        labels: data.map(d => d.name),
        values: data.map(d => d.count),
        type: "pie",
        hole: 0.58,
        marker: {{colors}},
        textinfo: textInfo,
        textposition: "inside",
        insidetextorientation: "horizontal",
        textfont: {{color: "white", size: 12, family: "Arial", weight: 800}},
        hovertemplate: "%{{label}}<br>Count: %{{value}}<br>Percentage: %{{percent}}<extra></extra>",
    }}], {{
        title: {{text: title, font: {{color: "#004C97", size: 15}}}},
        height: 280,
        margin: {{l: 10, r: 10, t: 45, b: 10}},
        paper_bgcolor: "white",
        font: {{family: "Arial", size: 11}}
    }}, {{responsive: true}});
}}

function chartSummaryMarkup(data, colors, totalLabel, totalSuffix = "") {{
    const total = data.reduce((sum, item) => sum + Number(item.count || 0), 0);
    const rows = data.map((item, index) => {{
        const pct = total ? Math.round((Number(item.count || 0) / total) * 100) : 0;
        return `
            <div class="chart-summary-row">
                <span class="chart-summary-swatch" style="background:${{escapeHtml(colors[index] || COLORS[index % COLORS.length])}}"></span>
                <span class="chart-summary-name">${{escapeHtml(item.name)}}</span>
                <span class="chart-summary-value">${{Number(item.count || 0).toLocaleString()}} (${{pct}}%)</span>
            </div>
        `;
    }}).join("");
    return `<div class="chart-summary"><div class="chart-summary-rows">${{rows}}</div><div class="chart-summary-total">${{escapeHtml(totalLabel)}}: ${{total.toLocaleString()}} ${{escapeHtml(totalSuffix)}}</div></div>`;
}}

function renderDonutWithSummary(id, title, data, colors, totalLabel, textInfo = "none", totalSuffix = "") {{
    const container = document.getElementById(id);
    if (!container) return;
    container.innerHTML = `<div class="coaching-donut-widget"><div class="coaching-donut-plot" id="${{id}}Plot"></div><div id="${{id}}Summary"></div></div>`;
    donut(`${{id}}Plot`, title, data, textInfo, colors);
    document.getElementById(`${{id}}Summary`).innerHTML = chartSummaryMarkup(data, colors, totalLabel, totalSuffix);
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

    const totals = rows.map(account =>
        TENURE_GROUPS.reduce((sum, group) => sum + counts[account][group.name], 0)
    );
    const maxTotal = Math.max(...totals, 1);
    const minimumReadableSegment = Math.max(3, Math.ceil(maxTotal * 0.05));
    const shouldShowSegmentLabel = (value, total) =>
        value >= minimumReadableSegment && (value / Math.max(total, 1)) >= 0.16;

    const traces = TENURE_GROUPS.map((group, i) => ({{
        name: group.name,
        x: rows.map(account => counts[account][group.name]),
        y: rows,
        type: "bar",
        orientation: "h",
        text: rows.map((account, index) => {{
            const value = counts[account][group.name];
            return shouldShowSegmentLabel(value, totals[index]) ? value : "";
        }}),
        textposition: "inside",
        texttemplate: "%{{text}}",
        textangle: 0,
        textfont: {{family: "Arial", size: 10}},
        insidetextanchor: "middle",
        constraintext: "inside",
        cliponaxis: true,
        marker: {{color: COLORS[i % COLORS.length]}},
        hovertemplate: "%{{y}}<br>" + group.name + ": %{{x}}<extra></extra>",
    }}));

    const totalAnnotations = rows.map((account, index) => ({{
        x: totals[index],
        y: account,
        text: `<b>${{totals[index]}}</b>`,
        xref: "x",
        yref: "y",
        showarrow: false,
        xanchor: "left",
        yanchor: "middle",
        xshift: 8,
        font: {{family: "Arial", size: 11, color: "#002B5C"}},
    }}));

    Plotly.newPlot("accountTenureStack", traces, {{
        title: {{text: "Tenure by Account", font: {{color: "#004C97", size: 15}}}},
        height: 340,
        margin: {{l: 135, r: 58, t: 45, b: 35}},
        barmode: "stack",
        xaxis: {{title: "Headcount", range: [0, Math.ceil(maxTotal * 1.12)]}},
        yaxis: {{title: "Account"}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        legend: {{orientation: "h", y: -0.25, font: {{size: 9}}}},
        annotations: totalAnnotations,
        uniformtext: {{mode: "hide", minsize: 9}},
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
    const rows = countBy(data, "Coaching Category");
    const colors = rows.map((row, index) => COACHING_CATEGORY_COLORS[row.name] || COLORS[index % COLORS.length]);
    renderDonutWithSummary("coachingCategoryDonut", "Coaching Category", rows, colors, "Total", "percent", " Sessions");
}}

function coachingStatusChart(data) {{
    const statusCounts = {{"Completed": 0, "Pending": 0}};
    data.forEach(row => {{
        if (isCompletedStatus(row["Coaching Status"])) {{
            statusCounts.Completed += 1;
        }} else if (isPendingStatus(row["Coaching Status"])) {{
            statusCounts.Pending += 1;
        }}
    }});
    const rows = [
        {{name: "Pending", count: statusCounts.Pending}},
        {{name: "Completed", count: statusCounts.Completed}},
    ];
    renderDonutWithSummary(
        "coachingStatusDonut",
        "Coaching Status",
        rows,
        rows.map(row => COACHING_STATUS_COLORS[row.name]),
        "Total",
        "percent",
        " Sessions"
    );
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
    const cx = 280;
    const cy = 315;
    const segments = [
        {{label: "VERY POOR", start: 180, end: 144, color: "#F4511E"}},
        {{label: "POOR", start: 144, end: 108, color: "#FB8C00"}},
        {{label: "FAIR", start: 108, end: 72, color: "#DCE800"}},
        {{label: "GOOD", start: 72, end: 36, color: "#35B84B"}},
        {{label: "EXCELLENT", start: 36, end: 0, color: "#00843D"}},
    ];
    const needleAngle = 180 - Math.max(0, Math.min(value, 100)) * 1.8;
    const grayArc = gaugeSegmentPath(cx, cy, 193, 184, 180, 0);
    const segmentMarkup = segments.map(segment => {{
        const labelAngle = (segment.start + segment.end) / 2;
        const labelPoint = gaugePoint(cx, cy, 222, labelAngle);
        return `
            <path d="${{gaugeSegmentPath(cx, cy, 180, 130, segment.start, segment.end)}}" fill="${{segment.color}}" stroke="white" stroke-width="2" />
            <text class="gauge-label" x="${{labelPoint.x.toFixed(1)}}" y="${{labelPoint.y.toFixed(1)}}">${{segment.label}}</text>
        `;
    }}).join("");

    document.getElementById("coachingConfidenceGauge").innerHTML = `
        <div class="score-gauge">
            <svg viewBox="0 0 560 390" role="img" aria-label="AI Confidence Level Detection ${{value}} percent">
                <text x="280" y="24" text-anchor="middle" style="font: 700 17px Arial; fill: #004C97;">AI Confidence Level Detection</text>
                <path d="${{grayArc}}" fill="#ECEFF1" />
                ${{segmentMarkup}}
                <path d="${{gaugeNeedlePath(cx, cy, needleAngle, 152, 10)}}" fill="#050505" />
                <circle cx="${{cx}}" cy="${{cy}}" r="15" fill="#050505" />
                <circle cx="${{cx}}" cy="${{cy}}" r="7" fill="white" />
                <circle cx="${{cx}}" cy="${{cy}}" r="2.5" fill="#050505" />
                <text class="gauge-value" x="280" y="375" text-anchor="middle">${{value}}% - ${{confidenceBand(value)}}</text>
            </svg>
        </div>
    `;
}}

function masterlistTable(data) {{
    const tableData = data
        .filter(r => filterMatches(MASTERLIST_FILTERS.empName, norm(r["Emp Name"])))
        .map(r => {{
            const hireDate = parseDateValue(r["Hire Date"]);
            return Object.assign({{}}, r, {{
                Tenure: tenureGroupName(hireDate, new Date()),
                __TenureDays: hireDate ? wholeDayDiff(hireDate, new Date()) : -1,
            }});
        }});
    renderDataTable("masterlistTable", sortedTableRows(tableData, MASTERLIST_COLUMNS, masterlistSortState), MASTERLIST_COLUMNS, masterlistSortState);
    bindTableSorting("masterlistTable", MASTERLIST_COLUMNS, masterlistSortState, render);
}}

function coachingTable(data) {{
    renderDataTable("coachingTable", sortedTableRows(data, COACHING_COLUMNS, coachingSortState), COACHING_COLUMNS, coachingSortState);
    bindTableSorting("coachingTable", COACHING_COLUMNS, coachingSortState, renderCoaching);
}}

function coachingExportRows() {{
    return sortedTableRows(filteredCoachingData(), COACHING_COLUMNS, coachingSortState).map(row => {{
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
            wch: (column.className || "").includes("long-text") ? 42 : Math.max(14, column.label.length + 2),
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

async function copyCoachingExportToClipboard(exportText) {{
    if (navigator.clipboard?.write && window.ClipboardItem) {{
        try {{
            const item = new ClipboardItem({{
                "text/plain": new Blob([exportText], {{type: "text/plain"}}),
                "text/html": new Blob([coachingExportHtmlTable()], {{type: "text/html"}}),
            }});
            await navigator.clipboard.write([item]);
            return true;
        }} catch (error) {{
            // Fall through to plain-text copy methods.
        }}
    }}

    if (navigator.clipboard?.writeText) {{
        try {{
            await navigator.clipboard.writeText(exportText);
            return true;
        }} catch (error) {{
            // Fall through to the legacy selection copy method.
        }}
    }}

    const textarea = document.createElement("textarea");
    textarea.value = exportText;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    textarea.style.top = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    let copied = false;
    try {{
        copied = document.execCommand("copy");
    }} catch (error) {{
        copied = false;
    }}
    textarea.remove();
    return copied;
}}

async function openCoachingSheets() {{
    const button = document.getElementById("openCoachingSheets");
    const originalText = button?.textContent || "Copy to Sheets";
    const exportText = coachingExportText("\\t");
    const rowCount = coachingExportRows().length;

    if (rowCount === 0) {{
        if (button) {{
            button.textContent = "No Rows to Copy";
            setTimeout(() => {{
                button.textContent = originalText;
            }}, 2600);
        }}
        return;
    }}

    const copied = await copyCoachingExportToClipboard(exportText);

    if (copied) {{
        window.open("https://sheets.new", "_blank", "noopener");
        if (button) {{
            button.textContent = "Copied - Paste in Sheets";
            setTimeout(() => {{
                button.textContent = originalText;
            }}, 2600);
        }}
        window.setTimeout(() => {{
            alert("Coaching Logs copied. In the new Google Sheet, click cell A1 and press Ctrl+V.");
        }}, 120);
        return;
    }}

    downloadBlob(
        `coaching_logs_${{new Date().toISOString().slice(0, 10)}}.tsv`,
        exportText,
        "text/tab-separated-values;charset=utf-8"
    );
    if (button) {{
        button.textContent = "Downloaded TSV";
        setTimeout(() => {{
            button.textContent = originalText;
        }}, 2600);
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
    coachingStatusChart(data);
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
    renderDataTable("recentMovements", sortedTableRows(rows, RECENT_MOVEMENT_COLUMNS, movementSortState), RECENT_MOVEMENT_COLUMNS, movementSortState);
    bindTableSorting("recentMovements", RECENT_MOVEMENT_COLUMNS, movementSortState, recentMovementsTable);
}}

function reflowDashboard() {{
    window.dispatchEvent(new Event("resize"));
    if (window.Plotly) {{
        document.querySelectorAll(".js-plotly-plot").forEach(plot => Plotly.Plots.resize(plot));
    }}
}}

function setLogsFocusMode(active) {{
    document.body.classList.toggle("logs-focus-mode", active);
    const button = document.getElementById("logsFocusToggle");
    if (button) {{
        button.setAttribute("aria-label", active ? "Collapse Coaching Logs" : "Expand Coaching Logs");
        button.setAttribute("title", active ? "Collapse Coaching Logs" : "Expand Coaching Logs");
    }}
    setTimeout(reflowDashboard, 0);
}}

function toggleLogsFocusMode() {{
    setLogsFocusMode(!document.body.classList.contains("logs-focus-mode"));
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

    if (!isCoaching) {{
        setLogsFocusMode(false);
    }}

    if (isMasterlist || isCoaching) {{
        setTimeout(reflowDashboard, 0);
    }}
}}

document.querySelectorAll(".tab-button").forEach(button => {{
    button.addEventListener("click", () => switchTab(button.dataset.tab));
}});

populateMasterlistFilters();
populateCoachingFilters();
initMultiFilterBehavior();
document.getElementById("downloadCoachingExcel")?.addEventListener("click", downloadCoachingExcel);
document.getElementById("openCoachingSheets")?.addEventListener("click", openCoachingSheets);
document.getElementById("logsFocusToggle")?.addEventListener("click", toggleLogsFocusMode);
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

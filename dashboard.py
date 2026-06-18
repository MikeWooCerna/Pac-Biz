import base64
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from datetime import datetime, timedelta
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

M7_DIR = Path(os.getenv("M7_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\M7"))
M7_SCRIPT = M7_DIR / "m7_pull.py"
M7_OUTPUT_FILE = M7_DIR / "M7_RAW.xlsx"

M7_COLUMN_MAP = {
    "Timestamp":                                              "ts",
    "Call Date:":                                            "date",
    "Emp Name":                                              "agent",
    "Score":                                                 "score",
    "Evaluation Type:":                                      "type",
    "QA":                                                    "coach",
    "Immediate Supervisor":                                  "supervisor",
    "Conducted Thorough Investigation:":                     "invest",
    "Feedback Summary:":                                     "feedback",
    "Opening Spiel (Inbound Calls) - 1pt":                   "os_in",
    "Opening Spiel (Outbound Calls) - 1pt":                  "os_out",
    "Closing Spiel - 1pt":                                   "closing",
    "Appropriate Response - 2pts":                           "approp",
    "No Response - 2pts":                                    "no_resp",
    "Fillers / Slang Words - 2pts":                          "fillers",
    "Acknowledgement / Ownsership - 1pt":                    "ack",
    "Proper Handling of Pauses or Hold Requests - 2pts":     "hold",
    "Acknowledges and Thank the Customer for Waiting - 1pt": "ack_hold",
    "Response Efficiency - 2pts":                            "resp_eff",
    "Empathy / Sympathy - 3pts":                             "empathy",
    "Adjust to Customer's Level - 3pts":                     "adjust",
    "Mute Button Usage - 1pt":                               "mute",
    "Active Listening - 5pts":                               "active",
    "Answered Customer's Questions - 4pts":                  "answered",
    "Probing Questions - 4pts":                              "probing",
    "Customer Verification - 10pts":                         "verif",
    "Clarification When Information is Missed - 4pts":       "clarif",
    "Lost Item SOP - 6pts":                                  "lost_sop",
    "Rudeness - 20pts":                                      "rude",
    "Transaction Completion - 20pts":                        "trans",
    "Speech Clarify - 5pts":                                 "speech",
    "QA_ID":                                                 "qa_id",
    "EMPLOYEE_ID":                                           "emp_id",
}

PARENTIS_DIR = Path(os.getenv("PARENTIS_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Parentis Health"))
PARENTIS_SCRIPT = PARENTIS_DIR / "parentis_pull.py"
PARENTIS_OUTPUT_FILE = PARENTIS_DIR / "PARENTIS_RAW.xlsx"

PARENTIS_COLUMN_MAP = {
    "Timestamp":                                              "ts",
    "Call Date:":                                            "date",
    "Emp Name":                                              "agent",
    "Score":                                                 "score",
    "Evaluation Type:":                                      "type",
    "QA":                                                    "coach",
    "Immediate Supervisor":                                  "supervisor",
    "Conducted Thorough Investigation:":                     "invest",
    "Feedback Summary:":                                     "feedback",
    "Opening Spiel (Inbound Calls) - 1pt":                   "os_in",
    "Opening Spiel (Outbound Calls) - 1pt":                  "os_out",
    "Closing Spiel - 1pt":                                   "closing",
    "Appropriate Response - 2pts":                           "approp",
    "No Response - 2pts":                                    "no_resp",
    "Fillers / Slang Words - 2pts":                          "fillers",
    "Acknowledgement / Ownsership - 1pt":                    "ack",
    "Proper Handling of Pauses or Hold Requests - 2pts":     "hold",
    "Acknowledges and Thank the Customer for Waiting - 1pt": "ack_hold",
    "Response Efficiency - 2pts":                            "resp_eff",
    "Empathy / Sympathy - 3pts":                             "empathy",
    "Adjust to Customer's Level - 3pts":                     "adjust",
    "Mute Button Usage - 1pt":                               "mute",
    "Active Listening - 5pts":                               "active",
    "Answered Customer's Questions - 4pts":                  "answered",
    "Probing Questions - 4pts":                              "probing",
    "Customer Verification - 10pts":                         "verif",
    "Clarification When Information is Missed - 4pts":       "clarif",
    "Lost Item SOP - 6pts":                                  "lost_sop",
    "Rudeness - 20pts":                                      "rude",
    "Transaction Completion - 20pts":                        "trans",
    "Speech Clarify - 5pts":                                 "speech",
    "QA_ID":                                                 "qa_id",
    "EMPLOYEE_ID":                                           "emp_id",
}

BRITELIFT_DIR = Path(os.getenv("BRITELIFT_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift"))
BRITELIFT_SCRIPT = BRITELIFT_DIR / "britelift_pull.py"
BRITELIFT_OUTPUT_FILE = BRITELIFT_DIR / "BRITELIFT_RAW.xlsx"

BRITELIFT_COLUMN_MAP = {
    "Timestamp":                                              "ts",
    "Call Date":                                             "date",
    "Emp Name":                                              "agent",
    "Score":                                                 "score",
    "Evaluation Type":                                       "type",
    "QA":                                                    "coach",
    "Immediate Supervisor":                                  "supervisor",
    "Conducted Thorough Investigation":                      "invest",
    "Feedback Summary ":                                     "feedback",
    "Opening Spiel (Inbound Calls) - 1pt":                   "os_in",
    "Opening Spiel (Outbound Calls) - 1pt":                  "os_out",
    "Closing Spiel - 1pt":                                   "closing",
    "Appropriate Response - 2pts":                           "approp",
    "No Response - 2pts":                                    "no_resp",
    "Fillers / Slang Words - 2pts":                          "fillers",
    "Acknowledgement / Ownsership - 1pt":                    "ack",
    "Proper Handling of Pauses or Hold Requests - 2pts":     "hold",
    "Acknowledges and Thank the Customer for Waiting - 1pt": "ack_hold",
    "Response Efficiency - 2pts":                            "resp_eff",
    "Empathy / Sympathy - 3pts":                             "empathy",
    "Adjust to Customer's Level - 3pts":                     "adjust",
    "Mute Button Usage - 1pt":                               "mute",
    "Active Listening - 5pts":                               "active",
    "Answered Customer's Questions - 4pts":                  "answered",
    "Probing Questions - 4pts":                              "probing",
    "Customer Verification - 10pts":                         "verif",
    "Clarification When Information is Missed - 4pts":       "clarif",
    "Lost Item SOP - 6pts":                                  "lost_sop",
    "Rudeness - 20pts":                                      "rude",
    "Transaction Completion - 20pts":                        "trans",
    "Speech Clarify - 5pts":                                 "speech",
    "QA_ID":                                                 "qa_id",
    "EMPLOYEE_ID":                                           "emp_id",
}

RIDEX_DIR = Path(os.getenv("RIDEX_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\RideX"))
RIDEX_SCRIPT = RIDEX_DIR / "Ridex_pull.py"
RIDEX_OUTPUT_FILE = RIDEX_DIR / "RIDEX_RAW.xlsx"

RIDEX_COLUMN_MAP = {
    "Timestamp":                                              "ts",
    "Call Date":                                             "date",
    "Emp Name":                                              "agent",
    "Score":                                                 "score",
    "Evaluation Type":                                       "type",
    "QA":                                                    "coach",
    "Immediate Supervisor":                                  "supervisor",
    "Conducted Thorough Investigation":                      "invest",
    "Feedback Summary ":                                     "feedback",
    "Opening Spiel (Inbound Calls) - 1pt":                   "os_in",
    "Opening Spiel (Outbound Calls) - 1pt":                  "os_out",
    "Closing Spiel - 1pt":                                   "closing",
    "Appropriate Response - 2pts":                           "approp",
    "No Response - 2pts":                                    "no_resp",
    "Fillers / Slang Words - 2pts":                          "fillers",
    "Acknowledgement / Ownsership - 1pt":                    "ack",
    "Proper Handling of Pauses or Hold Requests - 2pts":     "hold",
    "Acknowledges and Thank the Customer for Waiting - 1pt": "ack_hold",
    "Response Efficiency - 2pts":                            "resp_eff",
    "Empathy / Sympathy - 3pts":                             "empathy",
    "Adjust to Customer's Level - 3pts":                     "adjust",
    "Mute Button Usage - 1pt":                               "mute",
    "Active Listening - 5pts":                               "active",
    "Answered Customer's Questions - 4pts":                  "answered",
    "Probing Questions - 4pts":                              "probing",
    "Customer Verification - 10pts":                         "verif",
    "Clarification When Information is Missed - 4pts":       "clarif",
    "Lost Item SOP - 6pts":                                  "lost_sop",
    "Rudeness - 20pts":                                      "rude",
    "Transaction Completion - 20pts":                        "trans",
    "Speech Clarify - 5pts":                                 "speech",
    "QA_ID":                                                 "qa_id",
    "EMPLOYEE_ID":                                           "emp_id",
}

HAMILTON_DIR = Path(os.getenv("HAMILTON_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Hamilton"))
HAMILTON_SCRIPT = HAMILTON_DIR / "Hamilton_pull.py"
HAMILTON_OUTPUT_FILE = HAMILTON_DIR / "HAMILTON_RAW.xlsx"

SKYLINE_DIR = Path(os.getenv("SKYLINE_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Skyline"))
SKYLINE_SCRIPT = SKYLINE_DIR / "Skyline_pull.py"
SKYLINE_OUTPUT_FILE = SKYLINE_DIR / "SKYLINE_RAW.xlsx"

# Maps standard detail-table criterion keys → Hamilton base column name (without _AI/_Max suffix)
HAMILTON_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_Response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Hold",
    "ack_hold": "Thank_You_After_Hold",
    "resp_eff": "Resolution_Etiquette",
    "empathy":  "Concern_Handled_Properly",
    "adjust":   "Professionalism",
    "mute":     "Mute_Button_Usage",
    "active":   "Listening_Attentively",
    "answered": "Answers_Customer_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification_Measures",
    "clarif":   "Clarifies_When_Needed",
    "lost_sop": "Follow_up_SOP_Lost_Item_etc",
    "rude":     "Tone_Rudeness",
    "trans":    "Transfer_Escalation",
    "speech":   "Communication_Quality",
    "gen_q":    "General_Questions",
}

# Maps standard detail-table criterion keys → Skyline base column name (without _AI/_Max suffix)
SKYLINE_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_Response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgment_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thanks_the_Customer_for_Waiting",
    "resp_eff": "Resolution_Etiquette",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Professionalism",
    "mute":     "Mute_Button_Usage",
    "active":   "Listening_Attentively",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification_Measures",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Successfully_Processed",
    "speech":   "Communication_Quality",
    "gen_q":    "General_Questions",
}

# Skyline-unique criteria not in the standard QA_CRIT_META
SKYLINE_EXTRA_CRIT_MAP = {
    "sl_avoid_int":  "Avoiding_Interruptions_Verbal_Collision",
    "sl_no_vern":    "No_Vernacular_Local_Language",
    "sl_dead_air_l": "Dead_Air_Length_of_Pause",
    "sl_dead_air_nr":"Dead_Air_Spiel_No_Response_from_Customer",
    "sl_verif_sub":  "Customer_Verification",
    "sl_ride_cncl":  "Ride_Cancellation_SOP",
    "sl_timeliness": "Timeliness_in_Handling",
    "sl_stutter":    "Stuttering",
    "sl_grammar":    "Grammar",
    "sl_pronunc":    "Pronunciation",
    "sl_tone":       "Tone_of_Voice",
    "sl_prolong":    "Prolonging_the_Call",
}

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


def refresh_m7_output():
    if not M7_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(M7_SCRIPT)],
            cwd=str(M7_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping M7 pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping M7 pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_m7_workbook():
    if not M7_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(M7_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping M7 workbook load: {exc}")
        return pd.DataFrame()


# Canonical agent/supervisor name aliases (key = lowercase stripped variant)
NAME_ALIASES = {
    "espeleta, kenneth b": "Espeleta, Kenneth B",
    "kenneth espeleta":    "Espeleta, Kenneth B",
}

def _apply_name_aliases(series):
    return series.apply(
        lambda v: NAME_ALIASES.get(str(v).strip().lower(), str(v).strip()) if v else v
    )


def _transform_qa_source(source, column_map):
    df = source.rename(columns={k: v for k, v in column_map.items() if k in source.columns})

    def fmt_ts(v):
        try:
            dt = pd.to_datetime(str(v), errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        s = str(v).strip()
        return "" if s.lower() in ("nan", "na", "") else s

    def fmt_date(v):
        try:
            dt = pd.to_datetime(str(v), errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d")
        except Exception:
            pass
        s = str(v).strip()
        return "" if s.lower() in ("nan", "na", "") else s

    def clean_val(v):
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        s = str(v).strip()
        return "" if s.lower() in ("nan", "na") else s

    if "ts" in df.columns:
        df["ts"] = df["ts"].apply(fmt_ts)
    if "date" in df.columns:
        df["date"] = df["date"].apply(fmt_date)
    if "score" in df.columns:
        df["score"] = pd.to_numeric(df["score"], errors="coerce").fillna(0).astype(int)
    if "feedback" in df.columns:
        df["feedback"] = df["feedback"].apply(clean_val)
    if "agent" in df.columns:
        df["agent"] = _apply_name_aliases(df["agent"])
    if "supervisor" in df.columns:
        df["supervisor"] = _apply_name_aliases(df["supervisor"])
    if "coach" in df.columns:
        df["coach"] = _apply_name_aliases(df["coach"])

    if "emp_id" in df.columns and "ts" in df.columns:
        df["eval_key"] = (
            df["emp_id"].astype(str)
            + "_"
            + df["ts"].str.replace(r"[-: ]", "", regex=True)
        )

    def get_week_start(ts_val):
        try:
            s = str(ts_val).split()[0]
            dt = pd.to_datetime(s, errors="coerce")
            if pd.isna(dt):
                return ""
            monday = dt - timedelta(days=dt.weekday())
            return monday.strftime("%Y-%m-%d")
        except Exception:
            return ""

    if "ts" in df.columns:
        df["week_start"] = df["ts"].apply(get_week_start)

    keep = [
        "qa_id", "eval_key", "emp_id", "ts", "week_start", "date",
        "agent", "score", "type", "coach", "supervisor", "invest", "feedback",
        "os_in", "os_out", "closing", "approp", "no_resp", "fillers",
        "ack", "hold", "ack_hold", "resp_eff", "empathy", "adjust",
        "mute", "active", "answered", "probing", "verif", "clarif",
        "lost_sop", "rude", "trans", "speech",
    ]
    return df[[c for c in keep if c in df.columns]]


def transform_m7_data(source):
    return _transform_qa_source(source, M7_COLUMN_MAP)


def transform_parentis_data(source):
    return _transform_qa_source(source, PARENTIS_COLUMN_MAP)


def load_m7_data():
    refresh_m7_output()
    source = read_m7_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_m7_data(source)
    print(f"M7 QA rows: {len(result)}")
    return result


def refresh_parentis_output():
    if not PARENTIS_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(PARENTIS_SCRIPT)],
            cwd=str(PARENTIS_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Parentis pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Parentis pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_parentis_workbook():
    if not PARENTIS_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(PARENTIS_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Parentis workbook load: {exc}")
        return pd.DataFrame()


def load_parentis_data():
    refresh_parentis_output()
    source = read_parentis_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_parentis_data(source)
    print(f"Parentis QA rows: {len(result)}")
    return result


def transform_britelift_data(source):
    return _transform_qa_source(source, BRITELIFT_COLUMN_MAP)


def refresh_britelift_output():
    if not BRITELIFT_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(BRITELIFT_SCRIPT)],
            cwd=str(BRITELIFT_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Britelift pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Britelift pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_britelift_workbook():
    if not BRITELIFT_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(BRITELIFT_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Britelift workbook load: {exc}")
        return pd.DataFrame()


def load_britelift_data():
    refresh_britelift_output()
    source = read_britelift_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_britelift_data(source)
    print(f"Britelift QA rows: {len(result)}")
    return result


def transform_ridex_data(source):
    return _transform_qa_source(source, RIDEX_COLUMN_MAP)


def refresh_ridex_output():
    if not RIDEX_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(RIDEX_SCRIPT)],
            cwd=str(RIDEX_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping RideX pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping RideX pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ridex_workbook():
    if not RIDEX_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(RIDEX_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping RideX workbook load: {exc}")
        return pd.DataFrame()


def load_ridex_data():
    refresh_ridex_output()
    source = read_ridex_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_ridex_data(source)
    print(f"RideX QA rows: {len(result)}")
    return result


def transform_hamilton_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":    "ts",
        "Emp Name":           "agent",
        "QA":                 "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":  "status",
        "overall_score_ai":   "score_ai",
        "overall_score_human": "score_human",
        "QA_ID":              "qa_id",
        "EMPLOYEE_ID":        "emp_id",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    def fmt_ts(v):
        try:
            dt = pd.to_datetime(str(v), errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        return ""

    if "ts" in df.columns:
        df["ts"] = df["ts"].apply(fmt_ts)

    if "status" in df.columns:
        df["status"] = df["status"].apply(
            lambda v: str(v).strip().lower() if str(v).strip() not in ("", "nan") else "rated"
        )

    score_ai_s = pd.to_numeric(df.get("score_ai", pd.Series(dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric(df.get("score_human", pd.Series(dtype=float)), errors="coerce")
    df["score_human"] = [round(float(v), 1) if pd.notna(v) else None for v in score_h_s]

    def effective_score(row):
        if str(row.get("status", "")).lower() == "corrected":
            v = row.get("score_human")
            try:
                if v is not None:
                    return int(round(float(v)))
            except Exception:
                pass
        v = row.get("score_ai", 0)
        try:
            return int(round(float(v)))
        except Exception:
            return 0

    df["score"] = df.apply(effective_score, axis=1)

    def get_week_start(ts_val):
        try:
            s = str(ts_val).split()[0]
            dt = pd.to_datetime(s, errors="coerce")
            if pd.isna(dt):
                return ""
            monday = dt - timedelta(days=dt.weekday())
            return monday.strftime("%Y-%m-%d")
        except Exception:
            return ""

    if "ts" in df.columns:
        df["week_start"] = df["ts"].apply(get_week_start)

    if "emp_id" in df.columns and "ts" in df.columns:
        df["eval_key"] = (
            df["emp_id"].astype(str)
            + "_"
            + df["ts"].str.replace(r"[-: ]", "", regex=True)
        )

    # Map Hamilton criterion columns to standard detail-table keys
    for std_key, ham_base in HAMILTON_CRIT_MAP.items():
        ai_col  = f"{ham_base}_AI"
        max_col = f"{ham_base}_Max"
        ai_s  = pd.to_numeric(df.get(ai_col,  pd.Series(dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric(df.get(max_col, pd.Series(dtype=float)), errors="coerce").fillna(0)
        if std_key == "rude":
            df[std_key] = [
                None if mx == 0 else ("No" if ai >= mx else "Yes")
                for ai, mx in zip(ai_s, max_s)
            ]
        elif std_key == "lost_sop":
            df[std_key] = [
                "Not Applicable" if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        elif std_key in {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}:
            # Main Hamilton attributes — pass requires full marks (ai >= max)
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    crit_keys = list(HAMILTON_CRIT_MAP.keys())
    keep = [
        "qa_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_hamilton_output():
    if not HAMILTON_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(HAMILTON_SCRIPT)],
            cwd=str(HAMILTON_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Hamilton pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Hamilton pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_hamilton_workbook():
    if not HAMILTON_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(HAMILTON_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Hamilton workbook load: {exc}")
        return pd.DataFrame()


def load_hamilton_data():
    refresh_hamilton_output()
    source = read_hamilton_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_hamilton_data(source)
    print(f"Hamilton QA rows: {len(result)}")
    return result


def transform_skyline_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":     "ts",
        "Emp Name":            "agent",
        "QA":                  "coach",
        "Immediate Supervisor":"supervisor",
        "evaluation_status":   "status",
        "overall_score_ai":    "score_ai",
        "overall_score_human": "score_human",
        "QA_ID":               "qa_id",
        "EMPLOYEE_ID":         "emp_id",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    def fmt_ts(v):
        try:
            dt = pd.to_datetime(str(v), errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        return ""

    if "ts" in df.columns:
        df["ts"] = df["ts"].apply(fmt_ts)

    if "status" in df.columns:
        df["status"] = df["status"].apply(
            lambda v: str(v).strip().lower() if str(v).strip() not in ("", "nan") else "rated"
        )

    score_ai_s = pd.to_numeric(df.get("score_ai", pd.Series(dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric(df.get("score_human", pd.Series(dtype=float)), errors="coerce")
    df["score_human"] = [round(float(v), 1) if pd.notna(v) else None for v in score_h_s]

    def effective_score(row):
        if str(row.get("status", "")).lower() == "corrected":
            v = row.get("score_human")
            try:
                if v is not None:
                    return int(round(float(v)))
            except Exception:
                pass
        v = row.get("score_ai", 0)
        try:
            return int(round(float(v)))
        except Exception:
            return 0

    df["score"] = df.apply(effective_score, axis=1)

    def get_week_start(ts_val):
        try:
            s = str(ts_val).split()[0]
            dt = pd.to_datetime(s, errors="coerce")
            if pd.isna(dt):
                return ""
            monday = dt - timedelta(days=dt.weekday())
            return monday.strftime("%Y-%m-%d")
        except Exception:
            return ""

    if "ts" in df.columns:
        df["week_start"] = df["ts"].apply(get_week_start)

    if "emp_id" in df.columns and "ts" in df.columns:
        df["eval_key"] = (
            df["emp_id"].astype(str)
            + "_"
            + df["ts"].str.replace(r"[-: ]", "", regex=True)
        )

    # Map standard criterion keys to Skyline columns
    SKYLINE_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, sl_base in SKYLINE_CRIT_MAP.items():
        ai_col  = f"{sl_base}_AI"
        max_col = f"{sl_base}_Max"
        ai_s  = pd.to_numeric(df.get(ai_col,  pd.Series(dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric(df.get(max_col, pd.Series(dtype=float)), errors="coerce").fillna(0)
        if std_key == "rude":
            df[std_key] = [
                None if mx == 0 else ("No" if ai >= mx else "Yes")
                for ai, mx in zip(ai_s, max_s)
            ]
        elif std_key == "lost_sop":
            df[std_key] = [
                "Not Applicable" if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        elif std_key in SKYLINE_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    # Map Skyline-unique criteria
    SKYLINE_NA_KEYS = {"sl_ride_cncl", "sl_timeliness"}
    for sl_key, sl_base in SKYLINE_EXTRA_CRIT_MAP.items():
        ai_col  = f"{sl_base}_AI"
        max_col = f"{sl_base}_Max"
        ai_s  = pd.to_numeric(df.get(ai_col,  pd.Series(dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric(df.get(max_col, pd.Series(dtype=float)), errors="coerce").fillna(0)
        if sl_key in SKYLINE_NA_KEYS:
            df[sl_key] = [
                "Not Applicable" if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[sl_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    crit_keys = list(SKYLINE_CRIT_MAP.keys()) + list(SKYLINE_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_skyline_output():
    if not SKYLINE_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(SKYLINE_SCRIPT)],
            cwd=str(SKYLINE_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Skyline pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Skyline pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_skyline_workbook():
    if not SKYLINE_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(SKYLINE_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Skyline workbook load: {exc}")
        return pd.DataFrame()


def load_skyline_data():
    refresh_skyline_output()
    source = read_skyline_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_skyline_data(source)
    print(f"Skyline QA rows: {len(result)}")
    return result


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


def _parse_created_dt(date_val, time_val):
    date_str = cell_text(date_val)
    time_str = cell_text(time_val)
    combined = f"{date_str} {time_str}".strip()
    if not combined:
        return pd.NaT
    for fmt in ("%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M", "%m/%d/%Y"):
        try:
            return pd.to_datetime(combined if "%I:%M %p" in fmt or "%H:%M" in fmt else date_str, format=fmt)
        except (ValueError, Exception):
            continue
    return pd.to_datetime(date_str, errors="coerce")


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

    created_date = result["Created Date"] if "Created Date" in result.columns else pd.Series("", index=result.index)
    created_time = result["Created Time"] if "Created Time" in result.columns else pd.Series("", index=result.index)
    work["_created_dt"] = [_parse_created_dt(d, t) for d, t in zip(created_date, created_time)]

    coaching_id = result["Coaching ID"] if "Coaching ID" in result.columns else pd.Series("", index=result.index)
    work["_coaching_id_key"] = coaching_id.map(lambda v: cell_text(v).casefold())

    work = work.sort_values(
        ["_emp_key", "_category_key", "_date", "_created_dt", "_coaching_id_key", "_index"],
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


def _qa_block_html(aid, display_name, live_banner_name, badge_label, badge_cls, threshold, n_evals, expanded=True):
    body_display = "" if expanded else ' style="display:none"'
    chevron = "&#9650;" if expanded else "&#9660;"
    expanded_cls = " expanded" if expanded else ""
    return f"""
<div class="qa-acct-block{expanded_cls}" id="qa-block-{aid}">
  <div class="qa-acct-hdr" onclick="qaToggleBlock('{aid}',event)">
    <div class="qa-acct-hdr-left">
      <div class="qa-lb-left" style="margin-bottom:3px">
        <span class="qa-lb-dot"></span>
        <span style="font-size:12px;font-weight:700;color:#15803D">Live data &nbsp;&middot;&nbsp; {live_banner_name} &nbsp;&middot;&nbsp; <span id="{aid}-banner-count">{n_evals}</span> evaluations</span>
      </div>
      <div style="display:flex;gap:7px;flex-wrap:wrap;margin-top:3px">
        <span class="qa-badge {badge_cls}">{badge_label}</span>
        <span class="qa-badge qa-b-amber">Pass threshold: {threshold}%</span>
        <span class="qa-badge qa-b-green" id="{aid}-hdr-agents">&mdash; Agents</span>
        <span class="qa-badge qa-b-teal" id="{aid}-hdr-evals">&mdash; Evaluations</span>
      </div>
    </div>
    <div class="qa-acct-hdr-stats" id="{aid}-hdr-stats">
      <div class="qa-acct-stat"><span class="qa-acct-stat-l">Avg Score</span><span class="qa-acct-stat-v" id="{aid}-hdr-avg">&mdash;</span></div>
      <div class="qa-acct-stat"><span class="qa-acct-stat-l">Compliance Score</span><span class="qa-acct-stat-v" id="{aid}-hdr-pass">&mdash;</span></div>
      <div class="qa-acct-stat"><span class="qa-acct-stat-l">Below 85%</span><span class="qa-acct-stat-v" id="{aid}-hdr-below" style="color:#FCA5A5">&mdash;</span></div>
    </div>
    <button class="qa-acct-chevron" type="button" id="{aid}-chevron">{chevron}</button>
  </div>
  <div class="qa-acct-body" id="qa-body-{aid}"{body_display}>
    <div class="qa-page">
      <div class="qa-section-head">
        <div>
          <div class="qa-sh-title">{display_name} &mdash; Quality Assurance</div>
          <div class="qa-sh-sub" id="{aid}-view-sub">&mdash;</div>
        </div>
        <div class="qa-badges">
          <span class="qa-badge {badge_cls}">{badge_label}</span>
          <span class="qa-badge qa-b-green" id="{aid}-badge-agents">&mdash; Agents</span>
          <span class="qa-badge qa-b-teal" id="{aid}-badge-evals">&mdash; Evaluations</span>
          <span class="qa-badge qa-b-amber">Pass threshold: {threshold}%</span>
        </div>
      </div>
      <div class="qa-kpi-row">
        <div class="qa-kpi qa-knavy"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#1D4ED8" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg></div><div class="qa-kpi-lbl">Avg QA Score</div></div><div class="qa-kpi-val" id="{aid}-kpi-avg">&mdash;</div><div class="qa-kpi-d qa-dn" id="{aid}-kpi-avg-sub">&mdash;</div></div>
        <div class="qa-kpi qa-kgreen"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#15803D" stroke-width="2"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg></div><div class="qa-kpi-lbl">Compliance Score</div></div><div class="qa-kpi-val" id="{aid}-kpi-pass">&mdash;</div><div class="qa-kpi-d qa-du" id="{aid}-kpi-pass-sub">&mdash;</div></div>
        <div class="qa-kpi qa-kteal"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#0F766E" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg></div><div class="qa-kpi-lbl">Total Evaluations</div></div><div class="qa-kpi-val" id="{aid}-kpi-evals">&mdash;</div><div class="qa-kpi-d qa-dn" id="{aid}-kpi-evals-sub">&mdash;</div></div>
        <div class="qa-kpi qa-kpurple"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#6D28D9" stroke-width="2"><circle cx="12" cy="8" r="4"/><path d="M6 20v-2a4 4 0 014-4h4a4 4 0 014 4v2"/></svg></div><div class="qa-kpi-lbl">Total Agents</div></div><div class="qa-kpi-val" id="{aid}-kpi-agents">&mdash;</div><div class="qa-kpi-d qa-dn">Evaluated this period</div></div>
        <div class="qa-kpi qa-kamber"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#B45309" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg></div><div class="qa-kpi-lbl">Lowest Score</div></div><div class="qa-kpi-val" id="{aid}-kpi-low">&mdash;</div><div class="qa-kpi-d qa-dd" id="{aid}-kpi-low-sub">&mdash;</div></div>
        <div class="qa-kpi qa-kcoral"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#DC2626" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg></div><div class="qa-kpi-lbl">Below 85%</div></div><div class="qa-kpi-val" id="{aid}-kpi-below">&mdash;</div><div class="qa-kpi-d qa-dd" id="{aid}-kpi-below-sub">&mdash;</div></div>
      </div>
      <div class="qa-sum-strip">
        <div class="qa-sum-card" style="border-top:3px solid #0F9B58"><div class="qa-sum-lbl">Avg QA score</div><div class="qa-sum-val" id="{aid}-sum-avg" style="color:#0F9B58">&mdash;</div><div class="qa-sum-sub">All evaluations in range</div><div class="qa-sum-score" id="{aid}-sum-avg-note" style="color:#0F9B58">&mdash;</div></div>
        <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">Compliance Score</div><div class="qa-sum-val" id="{aid}-sum-pass" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="{aid}-sum-pass-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F9B58">Pass threshold: {threshold}%</div></div>
        <div class="qa-sum-card" style="border-top:3px solid #0D3B6E"><div class="qa-sum-lbl">Top performer</div><div class="qa-sum-val" id="{aid}-sum-top" style="color:#0D3B6E;font-size:15px">&mdash;</div><div class="qa-sum-sub" id="{aid}-sum-top-sub">&mdash;</div><div class="qa-sum-score" id="{aid}-sum-top-pass" style="color:#0F9B58">&mdash;</div></div>
        <div class="qa-sum-card" style="border-top:3px solid #F59E0B"><div class="qa-sum-lbl">Needs attention</div><div class="qa-sum-val" id="{aid}-sum-attn" style="color:#0D3B6E;font-size:15px">&mdash;</div><div class="qa-sum-sub" id="{aid}-sum-attn-sub">&mdash;</div><div class="qa-sum-score" style="color:#E85D3F">Coaching recommended</div></div>
      </div>
      <div class="qa-g3">
        <div class="qa-card">
          <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22,12 18,12 15,21 9,3 6,12 2,12"/></svg>QA Score Trend</div><div class="qa-cs" id="{aid}-trend-sub">Weekly avg (Mon&ndash;Sun) &middot; Target: {threshold}%</div></div><span class="qa-cb qa-cbg" id="{aid}-trend-badge">&mdash;</span></div>
          <div class="qa-cbody" style="padding:6px 10px;display:flex;flex-direction:column"><div style="position:relative;flex:1;min-height:90px"><canvas id="{aid}-trend-chart"></canvas></div></div>
        </div>
        <div class="qa-card">
          <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18M15 3v18M3 9h18M3 15h18"/></svg>Criteria pass rates</div><div class="qa-cs" id="{aid}-crit-sub">All 22 criteria &middot; sorted by pass rate</div></div><span class="qa-cb qa-cba">Account specific</span></div>
          <div style="padding:0 14px">
            <div style="max-height:120px;overflow-y:auto;padding:4px 0;border-bottom:1px solid #E2E8F0" id="{aid}-criteria-bars"></div>
            <div style="padding:3px 0 4px;font-size:10px;color:#CBD5E1">&#9679; Scroll to see all 22 criteria</div>
          </div>
        </div>
        <div class="qa-card">
          <div class="qa-ch"><div><div class="qa-ct">Score distribution</div><div class="qa-cs" id="{aid}-donut-sub">All evaluations</div></div></div>
          <div class="qa-cbody" style="padding:6px 10px">
            <div style="position:relative;height:130px"><canvas id="{aid}-donut-chart"></canvas></div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:2px;margin-top:5px" id="{aid}-donut-legend"></div>
          </div>
        </div>
      </div>
      <div class="qa-g2">
        <div class="qa-card">
          <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12,2 15.09,8.26 22,9.27 17,14.14 18.18,21.02 12,17.77 5.82,21.02 7,14.14 2,9.27 8.91,8.26"/></svg>Agent leaderboard</div><div class="qa-cs">Ranked by avg score</div></div><span class="qa-cb qa-cbb" id="{aid}-lb-badge">&mdash;</span></div>
          <div class="qa-cbody" style="padding:0 16px 8px">
            <table class="qa-lbt"><thead><tr><th>#</th><th>Agent</th><th>Evals</th><th>Avg</th><th>Min</th><th>Max</th><th>Comp Score</th></tr></thead><tbody id="{aid}-leaderboard"></tbody></table>
          </div>
        </div>
        <div class="qa-card">
          <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg>Coaching opportunities</div><div class="qa-cs">Criteria below 95% pass rate</div></div><span class="qa-cb qa-cbr" id="{aid}-coaching-count">&mdash;</span></div>
          <div class="qa-cbody" style="padding:10px 16px">
            <div id="{aid}-coaching-bars"></div>
            <div style="height:1px;background:#F1F5F9;margin:12px 0"></div>
            <div style="font-size:10px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.04em;margin-bottom:7px">QA coach breakdown</div>
            <div style="display:flex;gap:8px" id="{aid}-coach-breakdown"></div>
          </div>
        </div>
      </div>
      <div class="qa-card" id="{aid}-detail-card">
        <div class="qa-ch">
          <div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14,2 14,8 20,8"/></svg>All evaluations &mdash; detail table</div></div>
          <div style="display:flex;align-items:center;gap:8px">
            <span class="qa-cb qa-cbb" id="{aid}-tbl-count">&mdash;</span>
            <button onclick="downloadQAExcel('{aid}')" title="Download Excel" style="font-size:11px;padding:3px 10px;border:1px solid #CBD5E1;border-radius:5px;background:#fff;color:#475569;cursor:pointer;display:flex;align-items:center;gap:4px;white-space:nowrap">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7,10 12,15 17,10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>Excel
            </button>
            <button id="{aid}-focus-toggle" onclick="toggleQAFocusMode('{aid}')" title="Expand table" aria-label="Expand table" style="width:28px;height:28px;border:1px solid #CBD5E1;border-radius:5px;background:#fff;color:#475569;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0">
              <span class="qa-focus-icon" aria-hidden="true"></span>
            </button>
          </div>
        </div>
        <div class="qa-tbl-scroll">
          <table class="qa-dtbl">
            <thead><tr>
              <th style="min-width:130px">Evaluation Date</th>
              <th style="min-width:155px">Emp Name</th>
              <th style="min-width:110px">Immediate Head</th>
              <th style="min-width:110px">QA Coach</th>
              <th style="min-width:55px">Score</th>
              <th style="min-width:75px">Opening In</th><th style="min-width:75px">Opening Out</th>
              <th style="min-width:65px">Closing</th><th style="min-width:75px">Approp. Resp</th>
              <th style="min-width:65px">No Resp</th><th style="min-width:60px">Fillers</th>
              <th style="min-width:80px">Acknowledge</th><th style="min-width:70px">Hold Proc.</th>
              <th style="min-width:70px">Ack. Hold</th><th style="min-width:70px">Resp. Eff.</th>
              <th style="min-width:60px">Empathy</th><th style="min-width:60px">Adjust</th>
              <th style="min-width:55px">Mute</th><th style="min-width:70px">Active List.</th>
              <th style="min-width:70px">Answered Q</th><th style="min-width:65px">Probing Q</th>
              <th style="min-width:70px">Cust. Verif.</th><th style="min-width:65px">Clarif.</th>
              <th style="min-width:65px">Lost SOP</th><th style="min-width:65px">Rudeness</th>
              <th style="min-width:75px">Transaction</th><th style="min-width:65px">Speech</th>
              <th style="min-width:75px">Investigation</th>
              <th style="min-width:240px">Feedback Summary</th>
            </tr></thead>
            <tbody id="{aid}-detail-table"></tbody>
          </table>
        </div>
        <div style="padding:6px 16px;font-size:10px;color:#94A3B8;border-top:1px solid #F1F5F9" id="{aid}-tbl-footer">&mdash;</div>
      </div>
    </div>
  </div>
</div>"""


def main():
    masterlist = clean_columns(pd.read_csv(MASTERLIST_CSV))
    history = clean_columns(pd.read_csv(HISTORY_CSV))
    movement = clean_columns(pd.read_csv(MOVEMENT_CSV))
    coaching = load_coaching_data(masterlist)
    m7 = load_m7_data()
    parentis = load_parentis_data()
    britelift = load_britelift_data()
    ridex = load_ridex_data()
    hamilton = load_hamilton_data()
    skyline = load_skyline_data()

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
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>

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
        background: linear-gradient(90deg, #1B8A2E, var(--green));
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
        grid-template-columns: repeat(9, minmax(0, 1fr));
        gap: 8px;
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

    .chart-card-scrollable {{
        height: 420px;
        min-height: 0;
        overflow: hidden;
        display: flex;
        flex-direction: column;
    }}
    .chart-card-scrollable > div {{
        flex: 1 1 0;
        min-height: 0;
    }}
    .chart-scroll-container {{
        display: flex;
        flex-direction: column;
        height: 100%;
        min-height: 0;
    }}
    .chart-scroll-title {{
        color: #004C97;
        font-size: 15px;
        font-weight: bold;
        text-align: center;
        padding: 8px 0 4px;
        flex-shrink: 0;
    }}
    .chart-scroll-area {{
        flex: 1 1 0;
        min-height: 0;
        overflow-y: auto;
        overflow-x: hidden;
        scrollbar-width: thin;
        scrollbar-color: #004C97 #f0f0f0;
    }}
    .chart-scroll-area::-webkit-scrollbar {{
        width: 6px;
    }}
    .chart-scroll-area::-webkit-scrollbar-track {{
        background: #f0f0f0;
        border-radius: 3px;
    }}
    .chart-scroll-area::-webkit-scrollbar-thumb {{
        background: #004C97;
        border-radius: 3px;
    }}
    .chart-scroll-footer {{
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 4px 10px;
        padding: 6px 4px 4px;
        border-top: 1px solid #f0f0f0;
        flex-shrink: 0;
    }}
    .scroll-legend-item {{
        display: inline-flex;
        align-items: center;
        font-size: 10px;
        color: #444;
        white-space: nowrap;
        font-family: Arial, sans-serif;
    }}
    .scroll-legend-swatch {{
        display: inline-block;
        width: 10px;
        height: 10px;
        border-radius: 2px;
        margin-right: 4px;
        flex-shrink: 0;
    }}

    .coaching-chart-row {{
        display: grid;
        grid-column: 1 / -1;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 12px;
    }}
    .coaching-chart-row .chart-card {{
        height: 340px;
        max-height: 340px;
        overflow: hidden;
        display: flex;
        flex-direction: column;
        position: relative;
    }}
    .coaching-expand-btn {{
        position: absolute;
        top: 7px;
        right: 7px;
        z-index: 10;
        background: rgba(255,255,255,0.88);
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        width: 26px;
        height: 26px;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #64748B;
        padding: 0;
        transition: background .15s, color .15s, border-color .15s;
    }}
    .coaching-expand-btn:hover {{
        background: #EFF6FF;
        color: #0D3B6E;
        border-color: #0D3B6E;
    }}
    /* Expanded card state */
    .coaching-chart-row.has-expanded .chart-card:not(.expanded) {{
        display: none;
    }}
    .chart-card.expanded {{
        grid-column: 1 / -1;
        height: 560px !important;
        max-height: 560px !important;
    }}
    .chart-card.expanded .coaching-donut-plot {{
        flex: 0 0 390px;
        height: 390px;
    }}
    .chart-card.expanded .chart-summary-rows {{
        gap: 8px 28px;
    }}
    .chart-card.expanded .chart-summary-row {{
        font-size: 12px;
    }}
    .chart-card.expanded .score-gauge {{
        height: 510px;
    }}
    .chart-card.expanded .score-gauge svg {{
        height: 510px;
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

    #masterlistCard {{
        display: flex;
        flex-direction: column;
    }}

    #masterlistCard #masterlistTable {{
        min-height: 0;
    }}

    .masterlist-focus-icon {{
        position: relative;
        display: inline-block;
        width: 15px;
        height: 15px;
    }}

    .masterlist-focus-icon::before,
    .masterlist-focus-icon::after {{
        content: "";
        position: absolute;
        width: 6px;
        height: 6px;
        border-color: currentColor;
        border-style: solid;
    }}

    .masterlist-focus-icon::before {{
        top: 0;
        right: 0;
        border-width: 2px 2px 0 0;
    }}

    .masterlist-focus-icon::after {{
        left: 0;
        bottom: 0;
        border-width: 0 0 2px 2px;
    }}

    .masterlist-focus-mode .masterlist-focus-icon::before {{
        top: 2px;
        right: 2px;
        border-width: 0 0 2px 2px;
    }}

    .masterlist-focus-mode .masterlist-focus-icon::after {{
        left: 2px;
        bottom: 2px;
        border-width: 2px 2px 0 0;
    }}

    body.masterlist-focus-mode #masterlistControls .cards {{
        display: none;
    }}

    body.masterlist-focus-mode #masterlistPanel .grid > :not(#masterlistCard) {{
        display: none;
    }}

    body.masterlist-focus-mode #masterlistCard {{
        min-height: 0;
        height: calc(100vh - 300px);
    }}

    body.masterlist-focus-mode #masterlistCard .table-scroll {{
        max-height: none;
        height: 100%;
    }}

    body.masterlist-focus-mode #masterlistTable {{
        flex: 1 1 auto;
        min-height: 0;
    }}

    /* Quality detail table focus mode */
    .qa-focus-icon {{
        position: relative;
        display: inline-block;
        width: 15px;
        height: 15px;
    }}
    .qa-focus-icon::before,
    .qa-focus-icon::after {{
        content: "";
        position: absolute;
        width: 6px;
        height: 6px;
        border-color: currentColor;
        border-style: solid;
    }}
    .qa-focus-icon::before {{ top: 0; right: 0; border-width: 2px 2px 0 0; }}
    .qa-focus-icon::after  {{ left: 0; bottom: 0; border-width: 0 0 2px 2px; }}
    .qa-focus-mode .qa-focus-icon::before {{ top: 2px; right: 2px; border-width: 0 0 2px 2px; }}
    .qa-focus-mode .qa-focus-icon::after  {{ left: 2px; bottom: 2px; border-width: 2px 2px 0 0; }}
    body.qa-focus-mode #qualityPanel .qa-kpi-strip,
    body.qa-focus-mode #qualityPanel .qa-kpi-row,
    body.qa-focus-mode #qualityPanel .qa-sum-strip,
    body.qa-focus-mode #qualityPanel .qa-g3,
    body.qa-focus-mode #qualityPanel .qa-g2 {{ display: none; }}
    body.qa-focus-mode #qa-detail-card {{ min-height: 0; height: calc(100vh - 190px); }}
    body.qa-focus-mode #qa-detail-card .qa-tbl-scroll {{ max-height: none; height: calc(100% - 56px); }}

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
        height: 100%;
        overflow: hidden;
    }}

    .coaching-donut-plot {{
        flex: 0 0 210px;
        height: 210px;
    }}

    .coaching-donut-summary {{
        flex: 1;
        min-height: 0;
        overflow: hidden;
        display: flex;
        flex-direction: column;
    }}

    .chart-summary {{
        display: flex;
        flex-direction: column;
        height: 100%;
    }}

    .chart-summary-rows {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 5px 18px;
        overflow: hidden;
    }}

    .chart-summary-row {{
        display: grid;
        grid-template-columns: 10px minmax(0, 1fr) auto;
        align-items: center;
        gap: 6px;
        font-size: 11px;
        line-height: 1.2;
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
        flex-shrink: 0;
        margin-top: 6px;
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

    /* ── Quality tab ── */
    #qualityPanel {{ background:#F1F5F9;padding:0 }}
    #qualityPanel .qa-sticky-ctrl {{ position:sticky;top:var(--qa-stick-top,112px);z-index:90;background:#F1F5F9;border-bottom:2px solid transparent;transition:box-shadow .2s,border-color .2s }}
    #qualityPanel .qa-sticky-ctrl.scrolled {{ box-shadow:0 4px 16px rgba(0,0,0,.10);border-bottom-color:#CBD5E1 }}
    #qualityPanel .qa-sticky-ctrl .qa-section-head {{ padding:8px 20px 6px;background:#fff;border-top:1px solid #E2E8F0 }}
    #qualityPanel .qa-sticky-ctrl .qa-kpi-row {{ padding:8px 20px 0 }}
    #qualityPanel .qa-sticky-ctrl .qa-sum-strip {{ padding:8px 20px 10px }}
    .qa-live-pill {{ display:inline-flex;align-items:center;gap:5px;background:#F0FDF4;color:#166534;border:1px solid #BBF7D0;border-radius:999px;padding:3px 10px;font-size:11px;font-weight:600;white-space:nowrap }}
    .qa-live-dot {{ width:8px;height:8px;border-radius:50%;background:#00A651;animation:qa-pulse 2s ease-in-out infinite;flex-shrink:0 }}
    @keyframes qa-pulse {{ 0%,100%{{opacity:1;transform:scale(1)}} 50%{{opacity:.4;transform:scale(0.8)}} }}
    #qualityPanel .qa-filter-bar {{ background:#fff;border-bottom:1px solid #E2E8F0;padding:8px 20px;display:flex;align-items:flex-end;gap:10px;flex-wrap:wrap }}
    #qualityPanel .qa-fg {{ display:flex;flex-direction:column;gap:3px }}
    #qualityPanel .qa-fg label {{ font-size:10px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.06em }}
    #qualityPanel .qa-fsel {{ height:34px;padding:0 28px 0 10px;font-size:12px;border:1px solid #CBD5E1;border-radius:7px;background:#fff;color:#374151;cursor:pointer;font-weight:500;min-width:150px;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2394A3B8' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 8px center }}
    #qualityPanel .qa-fsel.hl {{ border-color:#0D3B6E;background-color:#EFF6FF;color:#0D3B6E;font-weight:700 }}
    #qualityPanel .qa-fdiv {{ width:1px;height:34px;background:#E2E8F0;align-self:flex-end }}
    #qualityPanel .qa-btn-clear {{ height:34px;padding:0 16px;font-size:12px;border:1px solid #E85D3F;border-radius:7px;background:#fff;color:#E85D3F;cursor:pointer;font-weight:600;align-self:flex-end }}
    #qualityPanel .qa-filter-bar .qa-bar-info {{ margin-left:auto;display:flex;align-items:center;gap:6px;flex-shrink:0 }}
    #qualityPanel .qa-info-pill {{ font-size:10px;font-weight:700;color:#0F9B58;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:5px;padding:3px 9px;white-space:nowrap }}
    #qualityPanel .qa-agent-lbt thead tr th {{ background:#0D3B6E;color:#fff;font-weight:700 }}
    #qualityPanel .qa-tl-lbt thead tr th {{ background:#7C3AED;color:#fff;font-weight:700 }}
    #qualityPanel .qa-coach-lbt thead tr th {{ background:#00A651;color:#fff;font-weight:700 }}
    /* Date range picker */
    #qualityPanel .qa-date-range-wrap {{ position:relative }}
    #qualityPanel .qa-date-range-btn {{ height:34px;padding:0 12px;font-size:12px;border:1px solid #CBD5E1;border-radius:7px;background:#fff;color:#374151;cursor:pointer;font-weight:500;display:flex;align-items:center;gap:7px;white-space:nowrap;min-width:210px }}
    #qualityPanel .qa-date-range-btn.picking {{ border-color:#0D3B6E;border-style:dashed;background:#F8FBFF }}
    #qualityPanel .qa-date-range-btn svg {{ width:13px;height:13px;color:#94A3B8;flex-shrink:0 }}
    #qualityPanel .qa-date-range-btn:hover {{ border-color:#0D3B6E }}
    #qualityPanel .qa-drp {{ position:absolute;top:calc(100% + 5px);left:0;background:#fff;border:1px solid #E2E8F0;border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,.12);z-index:999;display:none;padding:16px }}
    #qualityPanel .qa-drp.open {{ display:flex;flex-direction:column;gap:12px }}
    #qualityPanel .qa-drp-cals {{ display:flex;gap:20px }}
    #qualityPanel .qa-cal {{ width:210px }}
    #qualityPanel .qa-cal-header {{ display:flex;align-items:center;justify-content:space-between;margin-bottom:10px }}
    #qualityPanel .qa-cal-nav {{ background:none;border:none;cursor:pointer;color:#64748B;font-size:16px;width:24px;height:24px;display:flex;align-items:center;justify-content:center;border-radius:4px }}
    #qualityPanel .qa-cal-nav:hover {{ background:#F1F5F9 }}
    #qualityPanel .qa-cal-title {{ font-size:12px;font-weight:700;color:#1E293B }}
    #qualityPanel .qa-cal-grid {{ display:grid;grid-template-columns:repeat(7,1fr);gap:2px }}
    #qualityPanel .qa-cal-dh {{ font-size:9px;font-weight:700;color:#94A3B8;text-align:center;padding:3px 0;text-transform:uppercase }}
    #qualityPanel .qa-cal-d {{ width:100%;aspect-ratio:1;border:none;background:none;cursor:pointer;border-radius:6px;font-size:11px;color:#374151;display:flex;align-items:center;justify-content:center }}
    #qualityPanel .qa-cal-d:hover:not(:disabled) {{ background:#F1F5F9 }}
    #qualityPanel .qa-cal-d.today {{ border:1.5px solid #0891B2;color:#0891B2;font-weight:700 }}
    #qualityPanel .qa-cal-d.sel-start, #qualityPanel .qa-cal-d.sel-end {{ background:#0D3B6E!important;color:#fff!important;font-weight:700;border-radius:6px }}
    #qualityPanel .qa-cal-d.sel-preview {{ background:#1E5FA8!important;color:#fff!important;border-radius:6px;opacity:.85 }}
    #qualityPanel .qa-cal-d.in-range {{ background:#EFF6FF;border-radius:0;color:#1D4ED8 }}
    #qualityPanel .qa-cal-d:disabled {{ color:#CBD5E1;cursor:default }}
    #qualityPanel .qa-cal-d.other-month {{ color:#CBD5E1 }}
    #qualityPanel .qa-drp-footer {{ display:flex;align-items:center;justify-content:space-between;gap:8px;border-top:1px solid #F1F5F9;padding-top:10px }}
    #qualityPanel .qa-drp-inputs {{ display:flex;align-items:center;gap:8px }}
    #qualityPanel .qa-drp-input {{ height:30px;padding:0 8px;font-size:11px;border:1px solid #E2E8F0;border-radius:6px;color:#374151;width:100px;text-align:center }}
    #qualityPanel .qa-drp-dash {{ font-size:11px;color:#94A3B8 }}
    #qualityPanel .qa-drp-btns {{ display:flex;gap:6px }}
    #qualityPanel .qa-drp-apply {{ height:30px;padding:0 14px;font-size:11px;border:none;border-radius:6px;background:#0D3B6E;color:#fff;cursor:pointer;font-weight:700 }}
    #qualityPanel .qa-drp-reset {{ height:30px;padding:0 12px;font-size:11px;border:1px solid #E2E8F0;border-radius:6px;background:#fff;color:#64748B;cursor:pointer;font-weight:600 }}
    /* KPI strip */
    #qualityPanel .qa-kpi-strip {{ background:#0D3B6E;padding:0 20px;height:0;overflow:hidden;display:flex;align-items:center;transition:height .2s ease }}
    #qualityPanel .qa-kpi-strip.visible {{ height:32px }}
    #qualityPanel .qa-kpi-strip-items {{ display:flex;align-items:center;gap:20px;font-size:11px;white-space:nowrap }}
    #qualityPanel .qa-kpi-strip-item {{ display:flex;align-items:center;gap:5px }}
    #qualityPanel .qa-kpi-strip-label {{ font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:rgba(255,255,255,.5) }}
    #qualityPanel .qa-kpi-strip-val {{ font-size:12px;font-weight:800;color:#fff }}
    #qualityPanel .qa-kpi-strip-sep {{ width:1px;height:14px;background:rgba(255,255,255,.15) }}
    /* Live banner */
    #qualityPanel .qa-live-banner {{ background:#F0FDF4;border-bottom:1px solid #BBF7D0;padding:5px 20px;display:flex;align-items:center;justify-content:space-between }}
    #qualityPanel .qa-lb-left {{ display:flex;align-items:center;gap:8px;font-size:11px;color:#15803D;font-weight:600 }}
    #qualityPanel .qa-lb-dot {{ width:7px;height:7px;border-radius:50%;background:#0F9B58;animation:qaPulse 2s infinite }}
    @keyframes qaPulse {{ 0%,100%{{opacity:1}}50%{{opacity:.4}} }}
    /* Page */
    #qualityPanel .qa-page {{ padding:12px 20px;display:flex;flex-direction:column;gap:12px }}
    #qualityPanel .qa-section-head {{ display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px }}
    #qualityPanel .qa-sh-title {{ font-size:15px;font-weight:700;color:#0D3B6E }}
    #qualityPanel .qa-sh-sub {{ font-size:11px;color:#94A3B8;margin-top:2px }}
    #qualityPanel .qa-badges {{ display:flex;gap:7px;flex-wrap:wrap }}
    #qualityPanel .qa-badge {{ padding:3px 9px;border-radius:20px;font-size:10px;font-weight:600;border:1px solid }}
    #qualityPanel .qa-b-blue {{ background:#EFF6FF;color:#1D4ED8;border-color:#BFDBFE }}
    #qualityPanel .qa-b-green {{ background:#F0FDF4;color:#15803D;border-color:#BBF7D0 }}
    #qualityPanel .qa-b-teal {{ background:#F0FDFA;color:#0F766E;border-color:#99F6E4 }}
    #qualityPanel .qa-b-amber {{ background:#FFFBEB;color:#B45309;border-color:#FDE68A }}
    #qualityPanel .qa-b-skyline {{ background:#F0F9FF;color:#0369A1;border-color:#BAE6FD }}
    /* KPI cards */
    #qualityPanel .qa-kpi-row {{ display:grid;grid-template-columns:repeat(6,1fr);gap:8px }}
    #qualityPanel .qa-kpi {{ background:#fff;border-radius:8px;padding:10px 14px;border:1px solid #E2E8F0;position:relative;overflow:hidden }}
    #qualityPanel .qa-kpi::before {{ content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:8px 8px 0 0 }}
    #qualityPanel .qa-knavy::before {{ background:#0D3B6E }} #qualityPanel .qa-kgreen::before {{ background:#0F9B58 }}
    #qualityPanel .qa-kteal::before {{ background:#0891B2 }} #qualityPanel .qa-kpurple::before {{ background:#7C3AED }}
    #qualityPanel .qa-kamber::before {{ background:#F59E0B }} #qualityPanel .qa-kcoral::before {{ background:#E85D3F }}
    #qualityPanel .qa-kpi-icon {{ width:22px;height:22px;border-radius:6px;display:inline-flex;align-items:center;justify-content:center;margin-right:6px;vertical-align:middle;flex-shrink:0 }}
    #qualityPanel .qa-kpi-icon svg {{ width:11px;height:11px }}
    #qualityPanel .qa-knavy .qa-kpi-icon {{ background:#EFF6FF }} #qualityPanel .qa-kgreen .qa-kpi-icon {{ background:#F0FDF4 }}
    #qualityPanel .qa-kteal .qa-kpi-icon {{ background:#F0FDFA }} #qualityPanel .qa-kpurple .qa-kpi-icon {{ background:#F5F3FF }}
    #qualityPanel .qa-kamber .qa-kpi-icon {{ background:#FFFBEB }} #qualityPanel .qa-kcoral .qa-kpi-icon {{ background:#FEF2F2 }}
    #qualityPanel .qa-kpi-head {{ display:flex;align-items:center;margin-bottom:5px }}
    #qualityPanel .qa-kpi-lbl {{ font-size:9px;color:#64748B;font-weight:700;text-transform:uppercase;letter-spacing:.06em }}
    #qualityPanel .qa-kpi-val {{ font-size:20px;font-weight:800;color:#1E293B;line-height:1;margin-bottom:3px }}
    #qualityPanel .qa-kpi-d {{ font-size:10px;font-weight:600 }}
    #qualityPanel .qa-du {{ color:#0F9B58 }} #qualityPanel .qa-dd {{ color:#E85D3F }} #qualityPanel .qa-dn {{ color:#64748B }}
    /* Summary strip */
    #qualityPanel .qa-sum-strip {{ display:grid;grid-template-columns:repeat(4,1fr);gap:8px }}
    #qualityPanel .qa-sum-card {{ background:#fff;border-radius:8px;border:1px solid #E2E8F0;padding:11px 14px }}
    #qualityPanel .qa-sum-lbl {{ font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#94A3B8;margin-bottom:5px }}
    #qualityPanel .qa-sum-val {{ font-size:18px;font-weight:800;line-height:1 }}
    #qualityPanel .qa-sum-sub {{ font-size:10px;color:#64748B;margin-top:3px }}
    #qualityPanel .qa-sum-score {{ font-size:11px;font-weight:700;margin-top:3px }}
    /* Grids & cards */
    #qualityPanel .qa-g3 {{ display:grid;grid-template-columns:1.6fr 1fr 1fr 0.85fr;gap:12px }}
    #qualityPanel .qa-g2 {{ display:grid;grid-template-columns:repeat(3,1fr);gap:12px;grid-auto-rows:300px;align-items:stretch }}
    #qualityPanel .qa-card {{ background:#fff;border-radius:10px;border:1px solid #E2E8F0;overflow:hidden;display:flex;flex-direction:column }}
    #qualityPanel .qa-ch {{ padding:7px 12px;border-bottom:1px solid #F1F5F9;display:flex;align-items:center;justify-content:space-between;gap:8px }}
    #qualityPanel .qa-ct {{ font-size:12px;font-weight:700;color:#1E293B;display:flex;align-items:center;gap:6px }}
    #qualityPanel .qa-ct svg {{ width:13px;height:13px;color:#94A3B8 }}
    #qualityPanel .qa-cs {{ font-size:10px;color:#94A3B8;margin-top:1px }}
    #qualityPanel .qa-cbody {{ padding:8px 12px;flex:1 }}
    #qualityPanel .qa-cb {{ font-size:10px;padding:2px 8px;border-radius:10px;font-weight:600;white-space:nowrap }}
    #qualityPanel .qa-cbg {{ background:#F0FDF4;color:#15803D }} #qualityPanel .qa-cbb {{ background:#EFF6FF;color:#1D4ED8 }}
    #qualityPanel .qa-cba {{ background:#FFFBEB;color:#B45309 }} #qualityPanel .qa-cbr {{ background:#FEF2F2;color:#DC2626 }}
    /* Criteria bars */
    #qualityPanel .qa-cr-row {{ display:flex;align-items:center;gap:6px;margin-bottom:4px }}
    #qualityPanel .qa-cr-row:last-child {{ margin-bottom:0 }}
    #qualityPanel .qa-cr-lbl {{ font-size:11px;color:#374151;font-weight:500;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0 }}
    #qualityPanel .qa-cr-elig {{ font-size:9px;color:#CBD5E1;min-width:24px;text-align:center;flex-shrink:0 }}
    #qualityPanel .qa-cr-bg {{ width:80px;height:6px;background:#F1F5F9;border-radius:3px;overflow:hidden;flex-shrink:0 }}
    #qualityPanel .qa-cr-fill {{ height:100%;border-radius:3px }}
    #qualityPanel .qa-cr-val {{ font-size:11px;font-weight:700;min-width:38px;text-align:right;flex-shrink:0 }}
    /* Leaderboard */
    #qualityPanel .qa-lbt {{ width:100%;border-collapse:collapse;font-size:11px;table-layout:fixed }}
    #qualityPanel .qa-lbt th {{ padding:6px 8px;text-align:left;font-size:9px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid #F1F5F9;white-space:nowrap;overflow:hidden;position:sticky;top:0;z-index:2;background:#fff }}
    #qualityPanel .qa-lbt td {{ padding:6px 8px;border-bottom:1px solid #F8FAFC;vertical-align:middle }}
    #qualityPanel .qa-lbt tbody tr {{ height:36px }}
    #qualityPanel .qa-lbt td:nth-child(2) {{ overflow:hidden;white-space:nowrap;text-overflow:ellipsis }}
    #qualityPanel .qa-lbt tr:last-child td {{ border-bottom:none }}
    #qualityPanel .qa-lbt tr:hover td {{ background:#F8FAFC }}
    /* Fixed-height leaderboard cards — all three identical */
    #qualityPanel .qa-g2 .qa-card {{ height:100%;display:flex;flex-direction:column }}
    #qualityPanel .qa-g2 .qa-card .qa-cbody {{ flex:1;min-height:0;overflow-y:auto }}
    #qualityPanel .qa-rk {{ font-size:11px;font-weight:800;color:#94A3B8;width:18px;text-align:center }}
    #qualityPanel .qa-rk.gold {{ color:#F59E0B }} #qualityPanel .qa-rk.silver {{ color:#94A3B8 }} #qualityPanel .qa-rk.bronze {{ color:#B45309 }}
    #qualityPanel .qa-chip {{ display:inline-block;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:700 }}
    #qualityPanel .qa-cg {{ background:#F0FDF4;color:#15803D }} #qualityPanel .qa-cam {{ background:#FFFBEB;color:#B45309 }} #qualityPanel .qa-crr {{ background:#FEF2F2;color:#DC2626 }}
    /* Coaching bars */
    #qualityPanel .qa-cch-row {{ display:flex;align-items:center;gap:6px;margin-bottom:4px }}
    #qualityPanel .qa-cch-row:last-child {{ margin-bottom:0 }}
    #qualityPanel .qa-cch-lbl {{ font-size:11px;font-weight:500;color:#374151;width:158px;flex-shrink:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis }}
    #qualityPanel .qa-cch-bg {{ flex:1;height:7px;background:#F1F5F9;border-radius:4px;overflow:hidden }}
    #qualityPanel .qa-cch-fill {{ height:100%;border-radius:4px }}
    #qualityPanel .qa-cch-ct {{ font-size:11px;font-weight:700;min-width:60px;text-align:right }}
    /* Detail table */
    #qualityPanel .qa-tbl-scroll {{ overflow-x:auto;max-height:420px;overflow-y:auto }}
    #qualityPanel .qa-dtbl {{ width:100%;border-collapse:collapse;font-size:11px }}
    #qualityPanel .qa-dtbl th {{ padding:7px 10px;text-align:left;font-size:9px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid #E2E8F0;background:#F8FAFC;white-space:nowrap;position:sticky;top:0;z-index:1 }}
    #qualityPanel .qa-dtbl td {{ padding:6px 10px;border-bottom:1px solid #F8FAFC;vertical-align:middle }}
    #qualityPanel .qa-dtbl tbody tr {{ height:34px;box-sizing:border-box }}
    #qualityPanel .qa-dtbl tr:hover td {{ background:#F8FAFC }}
    #qualityPanel .qa-dtbl tr:last-child td {{ border-bottom:none }}
    /* Sticky first 5 columns */
    #qualityPanel .qa-dtbl th:nth-child(1),#qualityPanel .qa-dtbl td:nth-child(1){{position:sticky;left:0}}
    #qualityPanel .qa-dtbl th:nth-child(2),#qualityPanel .qa-dtbl td:nth-child(2){{position:sticky;left:130px}}
    #qualityPanel .qa-dtbl th:nth-child(3),#qualityPanel .qa-dtbl td:nth-child(3){{position:sticky;left:285px}}
    #qualityPanel .qa-dtbl th:nth-child(4),#qualityPanel .qa-dtbl td:nth-child(4){{position:sticky;left:395px}}
    #qualityPanel .qa-dtbl th:nth-child(5),#qualityPanel .qa-dtbl td:nth-child(5){{position:sticky;left:505px;box-shadow:2px 0 4px rgba(0,0,0,.06)}}
    #qualityPanel .qa-dtbl thead th:nth-child(-n+5){{z-index:3;background:#F8FAFC}}
    #qualityPanel .qa-dtbl tbody td:nth-child(-n+5){{background:#fff}}
    #qualityPanel .qa-dtbl tr:hover td:nth-child(-n+5){{background:#F8FAFC}}
    #qualityPanel .qa-av {{ width:24px;height:24px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;flex-shrink:0;vertical-align:middle;margin-right:4px }}
    #qualityPanel .qa-yn-y {{ color:#0F9B58;font-weight:700;font-size:11px }}
    #qualityPanel .qa-yn-n {{ color:#E85D3F;font-weight:700;font-size:11px }}
    #qualityPanel .qa-yn-na {{ color:#CBD5E1;font-size:10px }}
    #qualityPanel .qa-fb-cell {{ font-size:10px;color:#64748B;line-height:1.4;max-width:260px;white-space:normal }}
    @media (max-width:768px) {{
        #qualityPanel .qa-kpi-row {{ grid-template-columns:repeat(3,1fr) }}
        #qualityPanel .qa-sum-strip {{ grid-template-columns:repeat(2,1fr) }}
        #qualityPanel .qa-g3 {{ grid-template-columns:1fr }}
        #qualityPanel .qa-g2 {{ grid-template-columns:1fr }}
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
    <div class="card"><div class="label">Departments</div><div class="value" id="departments">0</div></div>
    <div class="card"><div class="label">Sub Departments</div><div class="value" id="accounts">0</div></div>
    <div class="card"><div class="label">Accounts</div><div class="value" id="approvedAccounts">0</div></div>
    <div class="card"><div class="label">Managers</div><div class="value" id="managers">0</div></div>
    <div class="card"><div class="label">History Records</div><div class="value" id="historyRecords">0</div></div>
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
        <div class="chart-card chart-card-scrollable"><div id="accountBar"></div></div>
        <div class="chart-card"><div id="managerBar"></div></div>
        <div class="chart-card"><div id="supervisorBar"></div></div>
    </div>
    <div class="chart-card chart-card-scrollable"><div id="accountTenureStack"></div></div>
    <div class="chart-card"><div id="tenureSegmentation"></div></div>
    <div class="chart-card"><div id="ageGroupBar"></div></div>
    <div class="chart-card"><div id="weeklyLine"></div></div>
    <div class="chart-card full" id="masterlistCard">
        <div class="table-heading">
            <h3>Master List</h3>
            <div class="table-actions">
                <a class="table-action" href="{masterlist_source_url}" target="_blank" rel="noopener">{masterlist_source_label}</a>
                <a class="table-action secondary" href="{masterlist_excel_url}" target="_blank" rel="noopener">Download Excel</a>
                <button class="table-action icon-action" type="button" id="masterlistFocusToggle" aria-label="Expand Master List" title="Expand Master List">
                    <span class="masterlist-focus-icon" aria-hidden="true"></span>
                </button>
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
    <div class="coaching-chart-row" id="coaching-chart-row">
        <div class="chart-card" id="coaching-cat-card">
            <button class="coaching-expand-btn" onclick="toggleCoachingCardExpand('coaching-cat-card')" title="Expand"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button>
            <div id="coachingCategoryDonut"></div>
        </div>
        <div class="chart-card" id="coaching-status-card">
            <button class="coaching-expand-btn" onclick="toggleCoachingCardExpand('coaching-status-card')" title="Expand"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button>
            <div id="coachingStatusDonut"></div>
        </div>
        <div class="chart-card" id="coaching-cov-card">
            <button class="coaching-expand-btn" onclick="toggleCoachingCardExpand('coaching-cov-card')" title="Expand"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button>
            <div id="coachingCoverageDonut"></div>
        </div>
        <div class="chart-card" id="coaching-conf-card">
            <button class="coaching-expand-btn" onclick="toggleCoachingCardExpand('coaching-conf-card')" title="Expand"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button>
            <div id="coachingConfidenceGauge"></div>
        </div>
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

<div class="qa-sticky-ctrl" id="qa-sticky-ctrl">
<!-- Filter bar -->
<div class="qa-filter-bar">
  <div class="qa-fg">
    <label>Date range</label>
    <div class="qa-date-range-wrap" id="qa-drp-wrap">
      <button class="qa-date-range-btn" id="qa-drp-trigger" onclick="qaToggleDRP(event)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
        <span id="qa-drp-label">May 8 &ndash; Jun 8, 2026</span>
      </button>
      <div class="qa-drp" id="qa-drp-panel">
        <div class="qa-drp-cals">
          <div class="qa-cal" id="qa-cal-left"></div>
          <div class="qa-cal" id="qa-cal-right"></div>
        </div>
        <div class="qa-drp-footer">
          <div class="qa-drp-inputs">
            <input class="qa-drp-input" id="qa-inp-start" type="text" placeholder="Start date" readonly>
            <span class="qa-drp-dash">&ndash;</span>
            <input class="qa-drp-input" id="qa-inp-end" type="text" placeholder="End date" readonly>
          </div>
          <div class="qa-drp-btns">
            <button class="qa-drp-reset" onclick="qaResetDRP()">Reset</button>
            <button class="qa-drp-apply" onclick="qaApplyDRP()">Apply</button>
          </div>
        </div>
      </div>
    </div>
  </div>
  <div class="qa-fdiv"></div>
  <div class="qa-fg">
    <label>Account</label>
    <select class="qa-fsel" id="qa-sel-account" onchange="qaOnAccountChange()">
      <option value="">All Accounts</option>
      <option value="m7">M7 &ndash; Ride-hailing support</option>
      <option value="parentis">Parentis Health</option>
      <option value="britelift">Britelift</option>
      <option value="ridex">RideX</option>
      <option value="hamilton">Hamilton</option>
      <option value="skyline">Skyline</option>
    </select>
  </div>
  <div class="qa-fg">
    <label>QA Coach</label>
    <select class="qa-fsel" id="qa-sel-coach" onchange="qaApplyFilters()">
      <option value="">All QA Coaches</option>
    </select>
  </div>
  <div class="qa-fg">
    <label>Agent</label>
    <select class="qa-fsel" id="qa-sel-agent" onchange="qaApplyFilters()">
      <option value="">All Agents</option>
    </select>
  </div>
  <div class="qa-fg">
    <label>Immediate Head</label>
    <select class="qa-fsel" id="qa-sel-head" onchange="qaApplyFilters()">
      <option value="">All Heads</option>
    </select>
  </div>
  <div class="qa-fdiv"></div>
  <button class="qa-btn-clear" onclick="qaClearFilters()">&times; Clear</button>
  <div class="qa-bar-info">
    <span class="qa-live-pill"><span class="qa-live-dot"></span>Live Data</span>
    <span class="qa-info-pill" id="qa-pill-qa-accounts">2 QA Accounts Loaded</span>
    <span class="qa-info-pill" id="qa-pill-total-accounts">&mdash; Accounts</span>
  </div>
</div>

<div class="qa-section-head">
  <div>
    <div class="qa-sh-title" id="qa-sh-title">All Accounts &mdash; Quality Assurance</div>
    <div class="qa-sh-sub" id="qa-view-sub">&mdash;</div>
  </div>
  <div class="qa-badges">
    <span class="qa-badge qa-b-amber" id="qa-badge-account">All Accounts</span>
    <span class="qa-badge qa-b-green" id="qa-badge-agents">&mdash; Agents</span>
    <span class="qa-badge qa-b-teal" id="qa-badge-evals">&mdash; Evaluations</span>
    <span class="qa-badge qa-b-amber">Pass threshold: 85%</span>
  </div>
</div>
  <div class="qa-kpi-row">
    <div class="qa-kpi qa-knavy"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#1D4ED8" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg></div><div class="qa-kpi-lbl">Avg QA Score</div></div><div class="qa-kpi-val" id="qa-kpi-avg">&mdash;</div><div class="qa-kpi-d qa-dn" id="qa-kpi-avg-sub">&mdash;</div></div>
    <div class="qa-kpi qa-kgreen"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#15803D" stroke-width="2"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg></div><div class="qa-kpi-lbl">Compliance Score</div></div><div class="qa-kpi-val" id="qa-kpi-pass">&mdash;</div><div class="qa-kpi-d qa-du" id="qa-kpi-pass-sub">&mdash;</div></div>
    <div class="qa-kpi qa-kteal"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#0F766E" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg></div><div class="qa-kpi-lbl">Total Evaluations</div></div><div class="qa-kpi-val" id="qa-kpi-evals">&mdash;</div><div class="qa-kpi-d qa-dn" id="qa-kpi-evals-sub">&mdash;</div></div>
    <div class="qa-kpi qa-kpurple"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#6D28D9" stroke-width="2"><circle cx="12" cy="8" r="4"/><path d="M6 20v-2a4 4 0 014-4h4a4 4 0 014 4v2"/></svg></div><div class="qa-kpi-lbl">Total Agents</div></div><div class="qa-kpi-val" id="qa-kpi-agents">&mdash;</div><div class="qa-kpi-d qa-dn">Evaluated this period</div></div>
    <div class="qa-kpi qa-kamber"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#B45309" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg></div><div class="qa-kpi-lbl">Lowest Score</div></div><div class="qa-kpi-val" id="qa-kpi-low">&mdash;</div><div class="qa-kpi-d qa-dd" id="qa-kpi-low-sub">&mdash;</div></div>
    <div class="qa-kpi qa-kcoral"><div class="qa-kpi-head"><div class="qa-kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="#DC2626" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg></div><div class="qa-kpi-lbl">Below 85%</div></div><div class="qa-kpi-val" id="qa-kpi-below">&mdash;</div><div class="qa-kpi-d qa-dd" id="qa-kpi-below-sub">&mdash;</div></div>
  </div>
  <!-- Hamilton-only key criteria strip (shown when Hamilton account selected) -->
  <div id="qa-hamilton-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #065F46"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-hcrit-greet-val" style="color:#065F46">&mdash;</div><div class="qa-sum-sub" id="qa-hcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#065F46">Hamilton criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-hcrit-prof-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-hcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">Hamilton criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-hcrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-hcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">Hamilton criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-hcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-hcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">Hamilton criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Resolution Etiquette</div><div class="qa-sum-val" id="qa-hcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-hcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">Hamilton criterion</div></div>
    </div>
  </div>
  <!-- Skyline-only key criteria strip (shown when Skyline account selected) -->
  <div id="qa-skyline-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #0369A1"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-scrit-greet-val" style="color:#0369A1">&mdash;</div><div class="qa-sum-sub" id="qa-scrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#0369A1">Skyline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-scrit-prof-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-scrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">Skyline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-scrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-scrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">Skyline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-scrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-scrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">Skyline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Resolution Etiquette</div><div class="qa-sum-val" id="qa-scrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-scrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">Skyline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Communication Qual.</div><div class="qa-sum-val" id="qa-scrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-scrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">Skyline criterion</div></div>
    </div>
  </div>
  <div class="qa-sum-strip" id="qa-sum-strip-main">
    <div class="qa-sum-card" style="border-top:3px solid #0F9B58"><div class="qa-sum-lbl">Avg QA score</div><div class="qa-sum-val" id="qa-sum-avg" style="color:#0F9B58">&mdash;</div><div class="qa-sum-sub">All evaluations in range</div><div class="qa-sum-score" id="qa-sum-avg-note" style="color:#0F9B58">&mdash;</div></div>
    <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">Compliance Score</div><div class="qa-sum-val" id="qa-sum-pass" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-sum-pass-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F9B58">Pass threshold: 85%</div></div>
    <div class="qa-sum-card" style="border-top:3px solid #0D3B6E"><div class="qa-sum-lbl">Top performer</div><div class="qa-sum-val" id="qa-sum-top" style="color:#0D3B6E;font-size:15px">&mdash;</div><div class="qa-sum-sub" id="qa-sum-top-sub">&mdash;</div><div class="qa-sum-score" id="qa-sum-top-pass" style="color:#0F9B58">&mdash;</div></div>
    <div class="qa-sum-card" style="border-top:3px solid #F59E0B"><div class="qa-sum-lbl">Needs attention</div><div class="qa-sum-val" id="qa-sum-attn" style="color:#0D3B6E;font-size:15px">&mdash;</div><div class="qa-sum-sub" id="qa-sum-attn-sub">&mdash;</div><div class="qa-sum-score" style="color:#E85D3F">Coaching recommended</div></div>
  </div>
</div><!-- /qa-sticky-ctrl -->

<!-- Compact KPI strip (appears on scroll) -->
<div class="qa-kpi-strip" id="qa-kpi-strip">
  <div class="qa-kpi-strip-items">
    <div class="qa-kpi-strip-item"><span class="qa-kpi-strip-label">Avg score</span><span class="qa-kpi-strip-val" id="qa-strip-avg">&mdash;</span></div>
    <div class="qa-kpi-strip-sep"></div>
    <div class="qa-kpi-strip-item"><span class="qa-kpi-strip-label">Compliance Score</span><span class="qa-kpi-strip-val" id="qa-strip-pass">&mdash;</span></div>
    <div class="qa-kpi-strip-sep"></div>
    <div class="qa-kpi-strip-item"><span class="qa-kpi-strip-label">Evals</span><span class="qa-kpi-strip-val" id="qa-strip-evals">&mdash;</span></div>
    <div class="qa-kpi-strip-sep"></div>
    <div class="qa-kpi-strip-item"><span class="qa-kpi-strip-label">Agents</span><span class="qa-kpi-strip-val" id="qa-strip-agents">&mdash;</span></div>
    <div class="qa-kpi-strip-sep"></div>
    <div class="qa-kpi-strip-item"><span class="qa-kpi-strip-label">Below 85%</span><span class="qa-kpi-strip-val" id="qa-strip-below" style="color:#FCA5A5">&mdash;</span></div>
    <div class="qa-kpi-strip-sep"></div>
    <div class="qa-kpi-strip-item"><span class="qa-kpi-strip-label">Lowest</span><span class="qa-kpi-strip-val" id="qa-strip-low" style="color:#FCD34D">&mdash;</span></div>
  </div>
</div>

<div id="qa-kpi-sentinel"></div>

<div class="qa-page" id="qa-shared-page">
  <div class="qa-g3">
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22,12 18,12 15,21 9,3 6,12 2,12"/></svg>QA Score Trend</div><div class="qa-cs" id="qa-trend-sub">Weekly avg (Mon&ndash;Sun) &middot; Target: 85%</div></div><span class="qa-cb qa-cbg" id="qa-trend-badge">&mdash;</span></div>
      <div class="qa-cbody" style="padding:6px 10px;display:flex;flex-direction:column"><div style="position:relative;flex:1;min-height:90px"><canvas id="qa-trend-chart"></canvas></div></div>
    </div>
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18M15 3v18M3 9h18M3 15h18"/></svg>Criteria pass rates</div><div class="qa-cs" id="qa-crit-sub">All 22 criteria &middot; sorted by pass rate</div></div><span class="qa-cb qa-cba">Score breakdown</span></div>
      <div style="padding:0 14px;flex:1;display:flex;flex-direction:column;justify-content:center">
        <div style="max-height:120px;overflow-y:auto;padding:4px 0;border-bottom:1px solid #E2E8F0" id="qa-criteria-bars"></div>
        <div style="padding:3px 0 4px;font-size:10px;color:#CBD5E1">&#9679; Scroll to see all 22 criteria</div>
      </div>
    </div>
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg>Coaching opportunities</div><div class="qa-cs">Criteria below 95% pass rate</div></div><span class="qa-cb qa-cbr" id="qa-coaching-count">&mdash;</span></div>
      <div class="qa-cbody" style="padding:6px 10px">
        <div id="qa-coaching-bars" style="max-height:110px;overflow-y:auto"></div>
        <div style="height:1px;background:#F1F5F9;margin:6px 0"></div>
        <div style="font-size:10px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.04em;margin-bottom:5px">QA coach breakdown</div>
        <div style="display:flex;gap:8px" id="qa-coach-breakdown"></div>
      </div>
    </div>
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct" id="qa-dist-title">Score distribution</div><div class="qa-cs" id="qa-donut-sub">All evaluations</div></div></div>
      <div class="qa-cbody" style="padding:6px 10px;display:flex;flex-direction:column;justify-content:center">
        <div id="qa-score-dist-wrap">
          <div style="position:relative;height:130px;flex-shrink:0"><canvas id="qa-donut-chart"></canvas></div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:2px;margin-top:5px" id="qa-donut-legend"></div>
        </div>
        <div id="qa-eval-dist-wrap" style="display:none">
          <div style="position:relative;height:150px;flex-shrink:0"><canvas id="qa-eval-dist-chart"></canvas></div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:2px;margin-top:5px" id="qa-eval-dist-legend"></div>
        </div>
      </div>
    </div>
  </div>
  <div id="qa-aivh-card" style="display:none;margin-bottom:14px">
    <!-- KPI tile strip -->
    <div id="qa-aivh-kpi-strip" style="display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:10px"></div>
    <!-- Gap distribution + summary row -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
      <!-- Gap Distribution card -->
      <div class="qa-card" style="margin-bottom:0">
        <div class="qa-ch"><div><div class="qa-ct">Gap Distribution</div><div class="qa-cs">H&minus;AI score difference</div></div></div>
        <div class="qa-cbody" style="padding:10px 14px;display:flex;gap:14px;align-items:center">
          <div style="position:relative;width:120px;height:120px;flex-shrink:0"><canvas id="qa-aivh-gap-donut"></canvas></div>
          <div style="flex:1">
            <div id="qa-aivh-gap-list" style="display:flex;flex-direction:column;gap:5px;font-size:11px"></div>
            <div id="qa-aivh-gap-range" style="margin-top:8px;font-size:10px;color:#94A3B8"></div>
          </div>
        </div>
      </div>
      <!-- AI vs Human Summary card -->
      <div style="background:#0F172A;border-radius:10px;padding:16px 18px;display:flex;flex-direction:column;gap:10px">
        <div>
          <div style="font-size:13px;font-weight:700;color:#F8FAFC">AI vs Human Score Comparison</div>
          <div style="font-size:11px;color:#94A3B8;margin-top:2px">Based on corrected evaluations only</div>
        </div>
        <div style="display:flex;gap:8px" id="qa-aivh-chips"></div>
        <div id="qa-aivh-insight" style="font-size:11px;color:#CBD5E1;line-height:1.5;border-top:1px solid #1E293B;padding-top:8px"></div>
      </div>
    </div>
  </div>
  <div class="qa-g2">
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12,2 15.09,8.26 22,9.27 17,14.14 18.18,21.02 12,17.77 5.82,21.02 7,14.14 2,9.27 8.91,8.26"/></svg>Agent leaderboard</div><div class="qa-cs">Ranked by avg score</div></div><span class="qa-cb qa-cbb" id="qa-lb-badge">&mdash;</span></div>
      <div class="qa-cbody" style="padding:0 16px 8px">
        <table class="qa-lbt qa-agent-lbt"><colgroup><col style="width:7%"><col><col style="width:11%"><col style="width:11%"><col style="width:10%"><col style="width:10%"><col style="width:13%"><col style="width:13%"></colgroup><thead><tr><th>#</th><th>Agent</th><th>Evals</th><th>Avg</th><th>Min</th><th>Max</th><th>Comp Score</th><th>Account</th></tr></thead><tbody id="qa-leaderboard"></tbody></table>
      </div>
    </div>
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>Team Leader leaderboard</div><div class="qa-cs">Ranked by avg score</div></div><span class="qa-cb qa-cbb" id="qa-tl-lb-badge">&mdash;</span></div>
      <div class="qa-cbody" style="padding:0 16px 8px">
        <table class="qa-lbt qa-tl-lbt"><colgroup><col style="width:8%"><col><col style="width:12%"><col style="width:13%"><col style="width:12%"><col style="width:12%"><col style="width:16%"></colgroup><thead><tr><th>#</th><th>Team Leader</th><th>Evals</th><th>Avg</th><th>Min</th><th>Max</th><th>Comp Score</th></tr></thead><tbody id="qa-tl-leaderboard"></tbody></table>
      </div>
    </div>
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg>QA Coach leaderboard</div><div class="qa-cs">Ranked by avg score</div></div><span class="qa-cb qa-cbb" id="qa-coach-lb-badge">&mdash;</span></div>
      <div class="qa-cbody" style="padding:0 16px 8px">
        <table class="qa-lbt qa-coach-lbt"><colgroup><col style="width:8%"><col><col style="width:12%"><col style="width:13%"><col style="width:12%"><col style="width:12%"><col style="width:16%"></colgroup><thead><tr><th>#</th><th>QA Coach</th><th>Evals</th><th>Avg</th><th>Min</th><th>Max</th><th>Comp Score</th></tr></thead><tbody id="qa-coach-leaderboard"></tbody></table>
      </div>
    </div>
  </div>
  <div class="qa-card" id="qa-detail-card">
    <div class="qa-ch">
      <div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14,2 14,8 20,8"/></svg>All evaluations &mdash; detail table</div></div>
      <div style="display:flex;align-items:center;gap:8px">
        <span class="qa-cb qa-cbb" id="qa-tbl-count">&mdash;</span>
        <button onclick="downloadQAExcel()" title="Download Excel" style="font-size:11px;padding:3px 10px;border:1px solid #CBD5E1;border-radius:5px;background:#fff;color:#475569;cursor:pointer;display:flex;align-items:center;gap:4px;white-space:nowrap">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7,10 12,15 17,10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>Excel
        </button>
        <button id="qa-focus-toggle" onclick="toggleQAFocusMode()" title="Expand table" aria-label="Expand table" style="width:28px;height:28px;border:1px solid #CBD5E1;border-radius:5px;background:#fff;color:#475569;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0">
          <span class="qa-focus-icon" aria-hidden="true"></span>
        </button>
      </div>
    </div>
    <div class="qa-tbl-scroll" id="qa-tbl-scroll-main">
      <table class="qa-dtbl">
        <thead><tr>
          <th style="min-width:130px;cursor:pointer;user-select:none" data-qa-sort="ts" onclick="qaSortTable('ts')">Evaluation Date <span class="sort-indicator">▼</span></th>
          <th style="min-width:200px;cursor:pointer;user-select:none" data-qa-sort="agent" onclick="qaSortTable('agent')">Emp Name <span class="sort-indicator"></span></th>
          <th style="min-width:160px;cursor:pointer;user-select:none" data-qa-sort="supervisor" onclick="qaSortTable('supervisor')">Immediate Head <span class="sort-indicator"></span></th>
          <th style="min-width:160px;cursor:pointer;user-select:none" data-qa-sort="coach" onclick="qaSortTable('coach')">QA Coach <span class="sort-indicator"></span></th>
          <th style="min-width:90px">Account</th>
          <th style="min-width:55px;cursor:pointer;user-select:none" data-qa-sort="score" onclick="qaSortTable('score')">Score <span class="sort-indicator"></span></th>
          <th style="min-width:75px">Opening In</th><th style="min-width:75px">Opening Out</th>
          <th style="min-width:65px">Closing</th><th style="min-width:75px">Approp. Resp</th>
          <th style="min-width:65px">No Resp</th><th style="min-width:60px">Fillers</th>
          <th style="min-width:80px">Acknowledge</th><th style="min-width:70px">Hold Proc.</th>
          <th style="min-width:70px">Ack. Hold</th><th style="min-width:70px">Resp. Eff.</th>
          <th style="min-width:60px">Empathy</th><th style="min-width:60px">Adjust</th>
          <th style="min-width:55px">Mute</th><th style="min-width:70px">Active List.</th>
          <th style="min-width:70px">Answered Q</th><th style="min-width:65px">Probing Q</th>
          <th style="min-width:70px">Cust. Verif.</th><th style="min-width:65px">Clarif.</th>
          <th style="min-width:65px">Lost SOP</th><th style="min-width:65px">Rudeness</th>
          <th style="min-width:75px">Transaction</th><th style="min-width:65px">Speech</th>
          <th style="min-width:75px">Investigation</th>
          <th style="min-width:240px">Feedback Summary</th>
        </tr></thead>
        <tbody id="qa-detail-table"></tbody>
      </table>
    </div>
    <div style="padding:6px 16px;font-size:10px;color:#94A3B8;border-top:1px solid #F1F5F9" id="qa-tbl-footer">&mdash;</div>
  </div>
</div>

</div>

<div class="footer">
    Developed for Pac-Biz Reporting MCerna | Data Source: Master List | Automation: Python 3.13.0
</div>

<script>
const masterlist = {to_records(masterlist)};
const historyData = {to_records(history)};
const movementData = {to_records(movement)};
const coachingData = {to_records(coaching)};
const qaRawData = {to_records(m7)};
const parentisRawData = {to_records(parentis)};
const briteliftRawData = {to_records(britelift)};
const ridexRawData = {to_records(ridex)};
const hamiltonRawData = {to_records(hamilton)};
const skylineRawData = {to_records(skyline)};

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
    {{label: "Hire Date", field: "Hire Date", sortable: true, sortType: "date"}},
    {{label: "Employement Class", field: "Employement Class", sortable: true}},
    {{label: "Tenure", field: "Tenure", sortable: true, sortField: "__TenureDays", sortType: "number"}},
    {{label: "Job Title", field: "Job Title", sortable: true}},
    {{label: "Employee Group", field: "Employee Group", sortable: true}},
    {{label: "Department", field: "Department", sortable: true}},
    {{label: "LOB/Account", field: "LOB / Account", sortable: true}},
    {{label: "Immediate Supervisor", field: "Immediate Supervisor", sortable: true}},
    {{label: "Manager", field: "Manager", sortable: true}},
    {{label: "Employment Status", field: "Employment Status", sortable: true}},
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
    {{label: "Created Time", field: "Created Time", className: "compact-col nowrap", sortable: true}},
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
const APPROVED_ACCOUNTS = new Set([
    "Alpha Tax", "Associate", "Blueline", "Brite Lift", "Buffalo", "C&H",
    "Circle Taxi", "Data Carz", "DMG", "Hamilton", "Kaizen", "Kelowna",
    "Keys Please", "M7 Ride", "Mediroute", "Monsoon", "Ollies", "Parentis Health",
    "R4H", "Reno Nevada", "Skyline", "Trans IOWA", "Victoria YC", "VIP", "VRN", "YCDC",
].map(v => v.toLowerCase()));
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
const qaTableSortState = {{field:'ts', dir:'desc'}};

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
    const baseFiltered = coachingData.filter(r => {{
        const monthKey = coachingMonthKey(r["Coaching Date"]);
        return (
            filterMatches(COACHING_FILTERS.emp, norm(r["Emp Name"])) &&
            filterMatches(COACHING_FILTERS.leader, norm(r["Coached by"])) &&
            filterMatches(COACHING_FILTERS.status, norm(r["Coaching Status"])) &&
            filterMatches(COACHING_FILTERS.category, norm(r["Coaching Category"])) &&
            filterMatches(COACHING_FILTERS.month, monthKey)
        );
    }});
    if (COACHING_FILTERS.categoryStatus.has(NONE_SELECTED)) return [];
    if (COACHING_FILTERS.categoryStatus.size === 0) return baseFiltered;
    const matchingEmps = new Set(
        baseFiltered
            .filter(r => filterMatches(COACHING_FILTERS.categoryStatus, norm(r["Category Status"])))
            .map(r => norm(r["Emp Name"]))
    );
    return baseFiltered.filter(r => matchingEmps.has(norm(r["Emp Name"])));
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

function donut(id, title, data, textInfo = "percent", colors = COLORS, showLegend = true) {{
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
        font: {{family: "Arial", size: 11}},
        showlegend: showLegend,
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
    container.innerHTML = `<div class="coaching-donut-widget"><div class="coaching-donut-plot" id="${{id}}Plot"></div><div class="coaching-donut-summary" id="${{id}}Summary"></div></div>`;
    donut(`${{id}}Plot`, title, data, textInfo, colors, false);
    document.getElementById(`${{id}}Summary`).innerHTML = chartSummaryMarkup(data, colors, totalLabel, totalSuffix);
}}

function toggleCoachingCardExpand(cardId) {{
    const card = document.getElementById(cardId);
    if(!card) return;
    const row = document.getElementById('coaching-chart-row');
    const isExpanded = card.classList.toggle('expanded');
    if(row) row.classList.toggle('has-expanded', isExpanded);
    const btn = card.querySelector('.coaching-expand-btn');
    if(btn) {{
        btn.title = isExpanded ? 'Collapse' : 'Expand';
        btn.innerHTML = isExpanded
            ? '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/><line x1="10" y1="14" x2="3" y2="21"/><line x1="21" y1="3" x2="14" y2="10"/></svg>'
            : '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>';
    }}
    if(window.Plotly) setTimeout(()=>window.dispatchEvent(new Event('resize')), 80);
}}

function bar(id, title, data, yTitle, maxItems = 10) {{
    const top = (maxItems === null ? data : data.slice(0, maxItems)).reverse();
    const chartHeight = Math.max(300, top.length * 30 + 80);
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
        height: chartHeight,
        margin: {{l: 135, r: 20, t: 45, b: 35}},
        xaxis: {{title: "Headcount"}},
        yaxis: {{title: yTitle}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        font: {{family: "Arial", size: 10}}
    }}, {{responsive: true}});
}}

function scrollableBar(id, title, data, yTitle) {{
    const container = document.getElementById(id);
    if (!container) return;
    const reversed = [...data].reverse();
    const rowH = 28, topM = 10, botM = 25;
    const chartH = Math.max(100, reversed.length * rowH + topM + botM);
    const pw = Math.max(200, container.clientWidth - 24);
    const footerHtml = `<div style="width:100%;text-align:center;font-size:10px;color:#555;font-family:Arial;padding-bottom:2px;">Headcount</div>`;
    container.innerHTML = `<div class="chart-scroll-container"><div class="chart-scroll-title">${{escapeHtml(title)}}</div><div class="chart-scroll-area"><div id="${{id}}Plot"></div></div><div class="chart-scroll-footer">${{footerHtml}}</div></div>`;
    Plotly.newPlot(`${{id}}Plot`, [{{
        x: reversed.map(d => d.count),
        y: reversed.map(d => d.name),
        type: "bar", orientation: "h",
        text: reversed.map(d => d.count), textposition: "auto",
        marker: {{color: "#004C97"}}
    }}], {{
        height: chartH, width: pw,
        margin: {{l: 135, r: 20, t: topM, b: botM}},
        xaxis: {{side: "bottom"}},
        yaxis: {{title: yTitle}},
        paper_bgcolor: "white", plot_bgcolor: "white",
        font: {{family: "Arial", size: 10}}
    }}, {{responsive: false}});
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
        height: 390,
        margin: {{l: 120, r: 20, t: 45, b: 35}},
        xaxis: {{title: "Headcount"}},
        yaxis: {{title: yTitle}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        font: {{family: "Arial", size: 10}}
    }}, {{responsive: true}});
}}

function accountTenureStack(data) {{
    const accounts = countBy(data, "LOB / Account").filter(d => APPROVED_ACCOUNTS.has(d.name.toLowerCase())).map(d => d.name);
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

    const atsContainer = document.getElementById("accountTenureStack");
    const rowH = 28, topM = 10, botM = 25;
    const chartH = Math.max(100, rows.length * rowH + topM + botM);
    const pw = Math.max(200, atsContainer.clientWidth - 24);
    const legendHtml = `<div style="width:100%;text-align:center;font-size:10px;color:#555;font-family:Arial;padding-bottom:2px;">Headcount</div>` +
        TENURE_GROUPS.map((g, i) =>
            `<span class="scroll-legend-item"><span class="scroll-legend-swatch" style="background:${{COLORS[i % COLORS.length]}}"></span>${{escapeHtml(g.name)}}</span>`
        ).join("");
    atsContainer.innerHTML = `<div class="chart-scroll-container"><div class="chart-scroll-title">Tenure by Account</div><div class="chart-scroll-area"><div id="accountTenureStackPlot"></div></div><div class="chart-scroll-footer">${{legendHtml}}</div></div>`;
    Plotly.newPlot("accountTenureStackPlot", traces, {{
        height: chartH, width: pw,
        margin: {{l: 135, r: 58, t: topM, b: botM}},
        barmode: "stack",
        showlegend: false,
        xaxis: {{range: [0, Math.ceil(maxTotal * 1.12)]}},
        yaxis: {{title: "Account"}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        annotations: totalAnnotations,
        uniformtext: {{mode: "hide", minsize: 9}},
        font: {{family: "Arial", size: 10}}
    }}, {{responsive: false}});
}}

function weeklyChart() {{
    const seen = new Set();
    const deduped = [];
    historyData.forEach(r => {{
        const week = norm(r["Week"]);
        const empId = norm(r["ID No."]);
        if (week && empId && !seen.has(`${{week}}|${{empId}}`)) {{
            seen.add(`${{week}}|${{empId}}`);
            deduped.push(r);
        }}
    }});

    const weeks = [...new Set(deduped.map(r => norm(r["Week"])))].filter(Boolean)
        .sort((a, b) => new Date(a) - new Date(b));

    const countMap = {{}};
    deduped.forEach(r => {{
        const week = norm(r["Week"]);
        const cls = norm(r["Employement Class"]) || "Other";
        const key = `${{week}}|${{cls}}`;
        countMap[key] = (countMap[key] || 0) + 1;
    }});

    // Sort classes by total headcount descending: largest at bottom (first trace), smallest at top (last trace)
    const classTotals = {{}};
    deduped.forEach(r => {{
        const cls = norm(r["Employement Class"]) || "Other";
        classTotals[cls] = (classTotals[cls] || 0) + 1;
    }});
    const allClasses = Object.keys(classTotals).sort((a, b) => classTotals[b] - classTotals[a]);

    const totals = {{}};
    weeks.forEach(week => {{
        totals[week] = allClasses.reduce((sum, cls) => sum + (countMap[`${{week}}|${{cls}}`] || 0), 0);
    }});
    const maxTotal = Math.max(...Object.values(totals), 1);

    const traces = allClasses.map((cls, i) => {{
        const yValues = weeks.map(week => countMap[`${{week}}|${{cls}}`] || 0);
        return {{
            name: cls,
            x: weeks,
            y: yValues,
            type: "bar",
            text: yValues.map(v => v > 0 ? v : ""),
            textposition: "inside",
            insidetextanchor: "middle",
            textfont: {{family: "Arial", size: 10, color: "white"}},
            constraintext: "inside",
            hovertemplate: `${{cls}}<br>Week: %{{x}}<br>Headcount: %{{y}}<extra></extra>`,
            marker: {{color: COLORS[i % COLORS.length]}},
        }};
    }});

    const annotations = weeks.map(week => ({{
        x: week,
        y: totals[week],
        text: `<b>${{totals[week]}}</b>`,
        xref: "x",
        yref: "y",
        showarrow: false,
        yanchor: "bottom",
        xanchor: "center",
        yshift: 5,
        font: {{family: "Arial", size: 13, color: "#002B5C"}},
    }}));

    const legendHtml = allClasses.map((cls, i) =>
        `<span style="display:inline-flex;align-items:center;font-size:10px;color:#444;font-family:Arial;white-space:nowrap;margin:0 4px;">` +
        `<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${{COLORS[i % COLORS.length]}};margin-right:4px;"></span>` +
        `${{escapeHtml(cls)}}</span>`
    ).join("");
    const wContainer = document.getElementById("weeklyLine");
    if (!wContainer) return;
    wContainer.innerHTML =
        `<div id="weeklyLinePlot"></div>` +
        `<div style="text-align:center;font-size:11px;color:#555;font-family:Arial;margin:24px 0 0;">Week</div>` +
        `<div style="display:flex;flex-wrap:wrap;justify-content:center;gap:4px 10px;padding:24px 4px 24px;">${{legendHtml}}</div>`;

    Plotly.newPlot("weeklyLinePlot", traces, {{
        title: {{text: "Weekly Headcount", font: {{color: "#004C97", size: 15}}}},
        height: 310,
        margin: {{l: 45, r: 20, t: 45, b: 30}},
        barmode: "stack",
        xaxis: {{title: ""}},
        yaxis: {{title: "Headcount", range: [0, Math.ceil(maxTotal * 1.15)]}},
        paper_bgcolor: "white",
        plot_bgcolor: "white",
        showlegend: false,
        annotations: annotations,
        uniformtext: {{mode: "hide", minsize: 8}},
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

function getCoverageStatus(pct) {{
    if(pct >= 75) return '🟢 Full Coverage';
    if(pct >= 50) return '🔵 High Coverage';
    if(pct >= 25) return '🟠 Moderate Coverage';
    return '🔴 Low Coverage';
}}

function computeCoachingCoverageMetrics(data) {{
    // Sessions and coached agents = completed records only
    const completedData = data.filter(r => isCompletedStatus(r["Coaching Status"]));
    const completedCount = completedData.length;
    const coachedSet = new Set(completedData.map(r => norm(r["Emp Name"])).filter(Boolean));
    const coachedCount = coachedSet.size;
    // Build masterlist: supervisor -> Set of active agents
    const mlBySup = {{}};
    const hasStatus = masterlist.some(r => (r["Employment Status"] || "").trim() !== "");
    masterlist.forEach(r => {{
        const sup = norm(r["Immediate Supervisor"]);
        const emp = norm(r["Emp Name"]);
        if(!sup || !emp) return;
        if(hasStatus) {{
            const st = (r["Employment Status"] || "").toLowerCase().trim();
            if(st && st !== "active") return;
        }}
        if(!mlBySup[sup]) mlBySup[sup] = new Set();
        mlBySup[sup].add(emp);
    }});
    // Denominator: ALL historic TLs from coachingData, filtered by leader filter if active
    const allHistoricLeaders = new Set(coachingData.map(r => norm(r["Coached by"])).filter(Boolean));
    const targetLeaders = [...allHistoricLeaders].filter(l => filterMatches(COACHING_FILTERS.leader, l));
    const leadersInML = targetLeaders.filter(l => mlBySup[l]);
    let totalDR = null, coveragePct = null, coverageStatus;
    if(targetLeaders.length === 0) {{
        totalDR = 0; coveragePct = 0; coverageStatus = getCoverageStatus(0);
    }} else if(leadersInML.length === 0) {{
        coverageStatus = '⚪ Direct Reports Not Configured';
    }} else {{
        const allDR = new Set();
        leadersInML.forEach(l => mlBySup[l].forEach(e => allDR.add(e)));
        totalDR = allDR.size;
        coveragePct = totalDR > 0 ? Math.round((coachedCount / totalDR) * 100) : 0;
        coverageStatus = getCoverageStatus(coveragePct);
        if(leadersInML.length < targetLeaders.length) coverageStatus += ' (partial data)';
    }}
    return {{ completedCount, coachedCount, totalDR, coveragePct, coverageStatus }};
}}

function _renderCoverageDonut(elementId, metrics) {{
    const totalDR = metrics.totalDR !== null ? metrics.totalDR : metrics.coachedCount;
    const notCoachedCount = Math.max(0, totalDR - metrics.coachedCount);
    const coverageRows = metrics.totalDR !== null
        ? [{{name: "Coached", count: metrics.coachedCount}}, {{name: "Not Coached", count: notCoachedCount}}]
        : [{{name: "Coached", count: metrics.coachedCount}}];
    const colors = metrics.totalDR !== null ? ["#00A651", "#CBD5E1"] : ["#00A651"];
    const totalLabel = metrics.totalDR !== null ? "Total Direct Reports" : "Coached Agents";
    renderDonutWithSummary(elementId, "Coaching Coverage", coverageRows, colors, totalLabel, "percent", " Agents");
}}

function coachingCoverageChart(metrics) {{
    _renderCoverageDonut("coachingCoverageDonut", metrics);
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

    const loadedMeta = document.getElementById("coachingLoadedMeta");
    if(loadedMeta) {{
        loadedMeta.textContent = `${{data.length.toLocaleString()}} of ${{coachingData.length.toLocaleString()}} records shown`;
    }}

    const covMetrics = computeCoachingCoverageMetrics(data);

    coachingCategoryChart(data);
    coachingStatusChart(data);
    coachingCoverageChart(covMetrics);
    coachingConfidenceGauge(data);

    const summary = coachingSummaryPivot(data);
    const summaryMeta = document.getElementById("coachingSummaryMeta");
    if(summaryMeta) {{
        const drStr = covMetrics.totalDR !== null ? covMetrics.totalDR : "Unknown";
        const covStr = covMetrics.coveragePct !== null ? `${{covMetrics.coveragePct}}%` : "N/A";
        summaryMeta.textContent = `${{covMetrics.completedCount.toLocaleString()}} Sessions | ${{covMetrics.coachedCount}}/${{drStr}} Agents Coached | ${{covStr}} Coverage | ${{covMetrics.coverageStatus}}`;
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

function setMasterlistFocusMode(active) {{
    document.body.classList.toggle("masterlist-focus-mode", active);
    const button = document.getElementById("masterlistFocusToggle");
    if (button) {{
        button.setAttribute("aria-label", active ? "Collapse Master List" : "Expand Master List");
        button.setAttribute("title", active ? "Collapse Master List" : "Expand Master List");
    }}
    setTimeout(reflowDashboard, 0);
}}

function toggleMasterlistFocusMode() {{
    setMasterlistFocusMode(!document.body.classList.contains("masterlist-focus-mode"));
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
    setText("approvedAccounts", new Set(data.map(r => norm(r["LOB / Account"])).filter(v => v && APPROVED_ACCOUNTS.has(v.toLowerCase()))).size);
    setText("managers", uniqueValues(data, "Manager").length);

    donut("deptDonut", "Headcount by Department", countBy(data, "Department"), "value");
    const approvedAccountData = countBy(data, "LOB / Account").filter(d => APPROVED_ACCOUNTS.has(d.name.toLowerCase()));
    scrollableBar("accountBar", "Headcount by Account", approvedAccountData, "Account");
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

// ─── QA QUALITY TAB ──────────────────────────────────────────────────────────
let qaChartsInitialized = false;
let qaTrendChart = null;
let qaDonutChart = null;
let qaEvalDistChart = null;
let qaAivhGapDonut = null;
let qaCurrentFiltered = [];

// Tag rows at init time
qaRawData.forEach(r => r._acct = 'M7');
parentisRawData.forEach(r => r._acct = 'Parentis');
briteliftRawData.forEach(r => r._acct = 'Britelift');
ridexRawData.forEach(r => r._acct = 'RideX');
hamiltonRawData.forEach(r => r._acct = 'Hamilton');
skylineRawData.forEach(r => r._acct = 'Skyline');

// Date range picker state — default: last 30 days
const _qaToday = new Date(); _qaToday.setHours(0,0,0,0);
let qaDrpStart = new Date(_qaToday.getTime() - 29 * 86400000);
let qaDrpEnd   = new Date(_qaToday);
let qaDrpPhase    = 0;
let qaDrpOpen     = false;
let qaDrpHoverDate = null;
let qaCalLeftYear  = _qaToday.getFullYear();
let qaCalLeftMonth = _qaToday.getMonth();

const QA_CRIT_META = [
    {{key:"os_in",   name:"Opening Spiel (Inbound)",     pts:"1pt",   inverse:false}},
    {{key:"os_out",  name:"Opening Spiel (Outbound)",    pts:"1pt",   inverse:false}},
    {{key:"closing", name:"Closing Spiel",               pts:"1pt",   inverse:false}},
    {{key:"approp",  name:"Appropriate Response",        pts:"2pts",  inverse:false}},
    {{key:"no_resp", name:"No Response",                 pts:"2pts",  inverse:false}},
    {{key:"fillers", name:"Fillers / Slang Words",       pts:"2pts",  inverse:false}},
    {{key:"ack",     name:"Acknowledgement / Ownership", pts:"1pt",   inverse:false}},
    {{key:"hold",    name:"Hold Requests",               pts:"2pts",  inverse:false}},
    {{key:"ack_hold",name:"Ack. for Waiting",            pts:"1pt",   inverse:false}},
    {{key:"resp_eff",name:"Response Efficiency",         pts:"2pts",  inverse:false}},
    {{key:"empathy", name:"Empathy / Sympathy",          pts:"3pts",  inverse:false}},
    {{key:"adjust",  name:"Adjust to Customer Level",    pts:"3pts",  inverse:false}},
    {{key:"mute",    name:"Mute Button Usage",           pts:"1pt",   inverse:false}},
    {{key:"active",  name:"Active Listening",            pts:"5pts",  inverse:false}},
    {{key:"answered",name:"Answered Questions",          pts:"4pts",  inverse:false}},
    {{key:"probing", name:"Probing Questions",           pts:"4pts",  inverse:false}},
    {{key:"verif",   name:"Customer Verification",       pts:"10pts", inverse:false}},
    {{key:"clarif",  name:"Clarification",               pts:"4pts",  inverse:false}},
    {{key:"lost_sop",name:"Lost Item SOP",               pts:"6pts",  inverse:false}},
    {{key:"rude",    name:"Rudeness",                    pts:"20pts", inverse:true}},
    {{key:"trans",   name:"Transaction Completion",      pts:"20pts", inverse:false}},
    {{key:"speech",  name:"Speech Clarity",              pts:"5pts",  inverse:false}},
];

const QA_AV = {{
    "Kenneth Vailoces":       {{bg:"#EFF6FF",tc:"#1D4ED8",ini:"KV"}},
    "Ruben John Cardama":     {{bg:"#FAECE7",tc:"#712B13",ini:"RC"}},
    "Alvin Lauga":            {{bg:"#F0FDF4",tc:"#166534",ini:"AL"}},
    "Raf Faustino":           {{bg:"#FFF7ED",tc:"#9A3412",ini:"RF"}},
    "Sheryl Lastimosa":       {{bg:"#FDF4FF",tc:"#86198F",ini:"SL"}},
    "Jherard Daclan":         {{bg:"#F0F9FF",tc:"#075985",ini:"JD"}},
    "Jannel Jade Alam":       {{bg:"#E0F2FE",tc:"#0369A1",ini:"JA"}},
    "Jonnarah Velasco":       {{bg:"#FDF4FF",tc:"#7E22CE",ini:"JV"}},
    "Lori Jane Faustorilla":  {{bg:"#FFF7ED",tc:"#C2410C",ini:"LF"}},
}};

const QA_BANDS = [
    {{label:"100%",     min:100, max:100, color:"#0F9B58"}},
    {{label:"95–99%",  min:95,  max:99,  color:"#0891B2"}},
    {{label:"85–94%",  min:85,  max:94,  color:"#F59E0B"}},
    {{label:"Below 85%",min:0,   max:84,  color:"#E85D3F"}},
];

const QA_MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const QA_MONTHS_S = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function qaEscapeHtml(s) {{
    return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}}
function qaAvg(arr) {{
    return arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : 0;
}}
function qaFmtDate(d) {{
    if (!d) return '';
    return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}}
function qaFmtDisplay(d) {{
    return QA_MONTHS_S[d.getMonth()]+' '+d.getDate()+', '+d.getFullYear();
}}
function qaChipCls(s) {{
    return s>=95?'qa-cg':s>=85?'qa-cam':'qa-crr';
}}
function qaYN(v) {{
    if (v==='Yes') return '<span style="color:#0F9B58;font-weight:600">✓</span>';
    if (v==='No')  return '<span style="color:#E85D3F;font-weight:600">✗</span>';
    return '<span style="color:#94A3B8">&mdash;</span>';
}}

function qaGetActiveData() {{
    const acct = (document.getElementById('qa-sel-account')?.value||'').trim();
    if (acct === 'm7') return qaRawData;
    if (acct === 'parentis') return parentisRawData;
    if (acct === 'britelift') return briteliftRawData;
    if (acct === 'ridex') return ridexRawData;
    if (acct === 'hamilton') return hamiltonRawData;
    if (acct === 'skyline') return skylineRawData;
    return [...qaRawData, ...parentisRawData, ...briteliftRawData, ...ridexRawData, ...hamiltonRawData, ...skylineRawData];
}}

// ─── Date Range Picker ────────────────────────────────────────────────────────
function qaUpdateDRPLabel() {{
    const el = document.getElementById('qa-drp-label');
    if (el) el.textContent = qaDrpEnd
        ? qaFmtDisplay(qaDrpStart)+' – '+qaFmtDisplay(qaDrpEnd)
        : qaFmtDisplay(qaDrpStart)+' – …';
    const si = document.getElementById('qa-inp-start');
    const ei = document.getElementById('qa-inp-end');
    if (si) si.value = qaFmtDate(qaDrpStart);
    if (ei) ei.value = qaDrpEnd ? qaFmtDate(qaDrpEnd) : '';
}}

function qaToggleDRP(e) {{
    e && e.stopPropagation();
    if (qaDrpOpen) {{ qaCloseDatePicker(); return; }}
    qaOpenDatePicker();
}}

function qaOpenDatePicker() {{
    qaDrpOpen = true; qaDrpPhase = 0; qaDrpHoverDate = null;
    qaCalLeftYear  = qaDrpStart.getFullYear();
    qaCalLeftMonth = qaDrpStart.getMonth();
    const panel = document.getElementById('qa-drp-panel');
    if (panel) panel.classList.add('open');
    qaRenderCals();
    qaUpdateDRPLabel();
}}

function qaCloseDatePicker() {{
    qaDrpOpen = false; qaDrpPhase = 0; qaDrpHoverDate = null;
    const panel = document.getElementById('qa-drp-panel');
    if (panel) panel.classList.remove('open');
    const trig = document.getElementById('qa-drp-trigger');
    if (trig) trig.classList.remove('picking');
}}

function qaNavCal(side, dir, e) {{
    if (e) {{ e.stopPropagation(); e.preventDefault(); }}
    void side;
    qaCalLeftMonth += dir;
    if (qaCalLeftMonth > 11) {{ qaCalLeftMonth=0; qaCalLeftYear++; }}
    if (qaCalLeftMonth < 0)  {{ qaCalLeftMonth=11; qaCalLeftYear--; }}
    qaRenderCals();
}}

function qaRenderCals() {{
    qaRenderOneCal('left', qaCalLeftYear, qaCalLeftMonth);
    let ry=qaCalLeftYear, rm=qaCalLeftMonth+1;
    if(rm>11){{rm=0;ry++;}}
    qaRenderOneCal('right', ry, rm);
}}

function qaRenderOneCal(side, year, month) {{
    const el = document.getElementById('qa-cal-'+side); if (!el) return;
    const daysInMonth=(y,m)=>new Date(y,m+1,0).getDate();
    const firstDow=new Date(year,month,1).getDay();
    const today=new Date(); today.setHours(0,0,0,0);
    const mName=QA_MONTHS[month]+' '+year;
    let html=`<div class="qa-cal-header">
        <button class="qa-cal-nav" type="button" onclick="qaNavCal('${{side}}',-1,event)">&#8249;</button>
        <span class="qa-cal-title">${{mName}}</span>
        <button class="qa-cal-nav" type="button" onclick="qaNavCal('${{side}}',1,event)">&#8250;</button>
    </div><div class="qa-cal-grid">`;
    ['Su','Mo','Tu','We','Th','Fr','Sa'].forEach(d=>html+=`<div class="qa-cal-dh">${{d}}</div>`);
    const prevDays=daysInMonth(year,month-1<0?11:month-1);
    for(let i=0;i<firstDow;i++) {{
        const d=prevDays-firstDow+1+i;
        const pm=month-1<0?11:month-1, py=month-1<0?year-1:year;
        html+=`<button class="qa-cal-d other-month" type="button" onclick="qaPickDay('${{py}}-${{String(pm+1).padStart(2,'0')}}-${{String(d).padStart(2,'0')}}')">${{d}}</button>`;
    }}
    const dim=daysInMonth(year,month);
    for(let d=1;d<=dim;d++) {{
        const dStr=year+'-'+String(month+1).padStart(2,'0')+'-'+String(d).padStart(2,'0');
        const dt=new Date(dStr+'T00:00:00');
        let cls='qa-cal-d';
        if(dt.getTime()===today.getTime()) cls+=' today';
        const start=qaDrpStart, end=qaDrpEnd;
        if(start&&dt.getTime()===start.getTime()) cls+=' sel-start';
        else if(end&&dt.getTime()===end.getTime()) cls+=' sel-end';
        else if(start&&end&&dt>start&&dt<end) cls+=' in-range';
        html+=`<button class="${{cls}}" type="button" onmouseenter="qaHoverDay('${{dStr}}')" onclick="qaPickDay('${{dStr}}')">${{d}}</button>`;
    }}
    const totalCells=firstDow+dim, rows=Math.ceil(totalCells/7), remaining=rows*7-totalCells;
    for(let i=1;i<=remaining;i++) {{
        const nm=month+1>11?0:month+1, ny=month+1>11?year+1:year;
        html+=`<button class="qa-cal-d other-month" type="button" onclick="qaPickDay('${{ny}}-${{String(nm+1).padStart(2,'0')}}-${{String(i).padStart(2,'0')}}')">${{i}}</button>`;
    }}
    html+='</div>';
    el.innerHTML=html;
}}

function qaHoverDay(dStr) {{
    if(qaDrpPhase!==1) return;
    const d=new Date(dStr+'T00:00:00');
    if(d.getTime()===(qaDrpHoverDate?.getTime()||0)) return;
    qaDrpHoverDate=d;
    ['left','right'].forEach(side=>{{
        const cal=document.getElementById('qa-cal-'+side); if(!cal) return;
        cal.querySelectorAll('.qa-cal-d').forEach(btn=>{{
            const ds=btn.getAttribute('onclick')?.match(/qaPickDay\\('([^']+)'\\)/)?.[1]; if(!ds) return;
            const dt=new Date(ds+'T00:00:00');
            btn.classList.remove('sel-start','sel-end','sel-preview','in-range');
            if(qaDrpStart&&dt.getTime()===qaDrpStart.getTime()) btn.classList.add('sel-start');
            else if(dt.getTime()===d.getTime()) btn.classList.add('sel-end');
            else if(qaDrpStart&&dt>qaDrpStart&&dt<d) btn.classList.add('sel-preview');
        }});
    }});
}}

function qaPickDay(dStr) {{
    const d=new Date(dStr+'T00:00:00'), trig=document.getElementById('qa-drp-trigger');
    if (qaDrpPhase===0) {{
        qaDrpStart=d; qaDrpEnd=null; qaDrpHoverDate=null; qaDrpPhase=1;
        if (trig) trig.classList.add('picking');
        const lbl=document.getElementById('qa-drp-label');
        if (lbl) lbl.textContent='Pick end date…';
        const si=document.getElementById('qa-inp-start'), ei=document.getElementById('qa-inp-end');
        if (si) si.value=dStr;
        if (ei) ei.value='';
        qaRenderCals();
    }} else {{
        if (d<qaDrpStart) {{ qaDrpEnd=qaDrpStart; qaDrpStart=d; }}
        else qaDrpEnd=d;
        qaDrpOpen=false; qaDrpPhase=0; qaDrpHoverDate=null;
        if (trig) trig.classList.remove('picking');
        qaUpdateDRPLabel();
        qaRenderCals();
        const panel=document.getElementById('qa-drp-panel');
        if (panel) panel.classList.remove('open');
        qaApplyFilters();
    }}
}}

function qaApplyDRP() {{ if(qaDrpEnd) {{ qaCloseDatePicker(); qaApplyFilters(); }} }}

function qaResetDRP() {{
    const _t=new Date(); _t.setHours(0,0,0,0);
    qaDrpStart=new Date(_t.getTime()-29*86400000);
    qaDrpEnd=new Date(_t);
    qaDrpPhase=0; qaUpdateDRPLabel(); qaCloseDatePicker(); qaApplyFilters();
}}

// ─── Account change ───────────────────────────────────────────────────────────
function qaOnAccountChange() {{
    ['qa-sel-coach','qa-sel-agent','qa-sel-head'].forEach(id=>{{
        const el=document.getElementById(id); if(el) el.value='';
    }});
    qaPopulateSelects();
    qaApplyFilters();
}}

// ─── Selects ──────────────────────────────────────────────────────────────────
function qaPopulateSelects() {{
    const pool = qaGetActiveData();
    [['qa-sel-coach','All QA Coaches',pool.map(r=>r.coach)],
     ['qa-sel-agent','All Agents',    pool.map(r=>r.agent)],
     ['qa-sel-head', 'All Immediate Heads', pool.map(r=>r.supervisor)]].forEach(([id,dflt,vals])=>{{
        const el=document.getElementById(id); if(!el) return;
        const cur=el.value;
        const arr=[...new Set(vals.filter(Boolean))].sort();
        el.innerHTML=`<option value="">${{dflt}}</option>`+arr.map(v=>`<option value="${{qaEscapeHtml(v)}}">${{qaEscapeHtml(v)}}</option>`).join('');
        el.value=arr.includes(cur)?cur:'';
    }});
}}

function qaComputeCriteria(data) {{
    return QA_CRIT_META.map(c=>{{
        const rows=data.map(r=>r[c.key]);
        const allNA=rows.every(v=>!v||v==='Not Applicable');
        if(allNA&&c.key==='lost_sop') return {{...c,yes:0,no:0,elig:0,pct:null,na:true}};
        const yes=c.inverse?rows.filter(v=>v==='No').length:rows.filter(v=>v==='Yes').length;
        const no=c.inverse?rows.filter(v=>v==='Yes').length:rows.filter(v=>v==='No').length;
        const elig=yes+no;
        return {{...c,yes,no,elig,pct:elig>0?Math.round(yes/elig*100):null,na:false}};
    }});
}}

function qaBuildTrend(data) {{
    const weeks={{}};
    data.forEach(r=>{{
        const k=r.week_start; if(!k) return;
        if(!weeks[k]) weeks[k]={{scores:[]}};
        const s=Number(r.score);
        if(!isNaN(s)&&s>0) weeks[k].scores.push(s);
    }});
    return Object.keys(weeks).sort().map(k=>{{
        const d=new Date(k+'T00:00:00');
        const lbl=QA_MONTHS[d.getMonth()]+' '+String(d.getDate()).padStart(2,'0')+', '+d.getFullYear();
        return {{week:lbl,avg:parseFloat(qaAvg(weeks[k].scores).toFixed(1)),n:weeks[k].scores.length}};
    }});
}}

// ─── Render functions ─────────────────────────────────────────────────────────
function qaUpdateKPIs(data) {{
    const acct=(document.getElementById('qa-sel-account')?.value||'').trim();
    const titleEl=document.getElementById('qa-sh-title');
    const badgeEl=document.getElementById('qa-badge-account');
    if(titleEl) {{
        if(acct==='m7') titleEl.textContent='M7 — Quality Assurance';
        else if(acct==='parentis') titleEl.textContent='Parentis Health — Quality Assurance';
        else if(acct==='britelift') titleEl.textContent='Britelift — Quality Assurance';
        else if(acct==='ridex') titleEl.textContent='RideX — Quality Assurance';
        else if(acct==='hamilton') titleEl.textContent='Hamilton — Quality Assurance';
        else if(acct==='skyline') titleEl.textContent='Skyline — Quality Assurance';
        else titleEl.textContent='All Accounts — Quality Assurance';
    }}
    if(badgeEl) {{
        if(acct==='m7') {{ badgeEl.textContent='M7 Account'; badgeEl.className='qa-badge qa-b-blue'; }}
        else if(acct==='parentis') {{ badgeEl.textContent='Parentis Health'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='britelift') {{ badgeEl.textContent='Britelift'; badgeEl.className='qa-badge qa-b-amber'; }}
        else if(acct==='ridex') {{ badgeEl.textContent='RideX'; badgeEl.className='qa-badge qa-b-purple'; }}
        else if(acct==='hamilton') {{ badgeEl.textContent='Hamilton'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='skyline') {{ badgeEl.textContent='Skyline'; badgeEl.className='qa-badge qa-b-skyline'; }}
        else {{ badgeEl.textContent='All Accounts'; badgeEl.className='qa-badge qa-b-amber'; }}
    }}
    const scores=data.map(r=>Number(r.score)).filter(v=>!isNaN(v)&&v>0);
    const n=data.length;
    const agents=new Set(data.map(r=>r.agent).filter(Boolean));
    const avg=scores.length?qaAvg(scores):null;
    const passCount=scores.filter(s=>s>=85).length;
    const passRate=scores.length?passCount/scores.length*100:null;
    const belowCount=scores.filter(s=>s<85).length;
    const minScore=scores.length?Math.min(...scores):null;
    const set=(id,v)=>{{const el=document.getElementById(id);if(el)el.textContent=v;}};
    set('qa-kpi-avg',    avg?avg.toFixed(1)+'%':'—');
    set('qa-kpi-pass',   passRate!==null?passRate.toFixed(1)+'%':'—');
    set('qa-kpi-evals',  n);
    set('qa-kpi-agents', agents.size);
    set('qa-kpi-low',    minScore!==null?minScore.toFixed(0)+'%':'—');
    set('qa-kpi-below',  belowCount);
    set('qa-kpi-avg-sub',  avg!==null?(avg>=85?'✔ Above threshold':'⚠ Below target'):'—');
    set('qa-kpi-pass-sub', passCount+' of '+n+' compliant');
    set('qa-kpi-evals-sub',agents.size+' agent'+(agents.size===1?'':'s'));
    set('qa-kpi-low-sub',  minScore!==null?(minScore<85?'⚠ Below threshold':'Within range'):'—');
    set('qa-kpi-below-sub',belowCount?belowCount+' eval'+(belowCount===1?'':'s')+' below 85%':'None below 85%');
    const agBadge=document.getElementById('qa-badge-agents');
    if(agBadge) agBadge.textContent=agents.size+' Agents';
    const evBadge=document.getElementById('qa-badge-evals');
    if(evBadge) evBadge.textContent=n+' Evaluations';
    const viewSub=document.getElementById('qa-view-sub');
    if(viewSub) {{
        const s=qaFmtDisplay(qaDrpStart), e=qaDrpEnd?qaFmtDisplay(qaDrpEnd):'…';
        viewSub.textContent=`${{s}} – ${{e}} · ${{agents.size}} agent${{agents.size===1?'':'s'}} · ${{n}} evaluation${{n===1?'':'s'}}`;
    }}
    const byAgent={{}};
    data.forEach(r=>{{
        if(!r.agent) return;
        const s=Number(r.score);
        if(!isNaN(s)&&s>0){{if(!byAgent[r.agent])byAgent[r.agent]={{scores:[]}};byAgent[r.agent].scores.push(s);}}
    }});
    const agentArr=Object.entries(byAgent).map(([name,d])=>{{return {{name,avg:qaAvg(d.scores),n:d.scores.length}};}})
        .filter(a=>a.n>0).sort((a,b)=>b.avg-a.avg);
    set('qa-sum-avg',      avg?avg.toFixed(1)+'%':'—');
    set('qa-sum-avg-note', avg!==null?(avg>=85?'✓ Above 85% threshold':'⚠ Below target'):'—');
    set('qa-sum-pass',     passRate!==null?passRate.toFixed(1)+'%':'—');
    set('qa-sum-pass-sub', passCount+' of '+n+' evaluations');
    set('qa-sum-top',      agentArr[0]?.name||'—');
    set('qa-sum-top-sub',  agentArr[0]?'Avg '+agentArr[0].avg.toFixed(1)+'% · '+agentArr[0].n+' eval'+(agentArr[0].n===1?'':'s'):'');
    set('qa-sum-top-pass', agentArr[0]?agentArr[0].avg.toFixed(1)+'%':'—');
    const critStats=qaComputeCriteria(data).filter(c=>c.pct!==null).sort((a,b)=>a.pct-b.pct);
    const worst=critStats[0];
    set('qa-sum-attn',     worst?worst.name:'—');
    set('qa-sum-attn-sub', worst?worst.pct+'% pass rate':'');
}}

function qaUpdateTrend(data) {{
    if(!qaTrendChart) return;
    const trend=qaBuildTrend(data);
    const labels=trend.map(t=>t.week), avgs=trend.map(t=>t.avg), counts=trend.map(t=>t.n), target=labels.map(()=>85);
    qaTrendChart.data.labels=labels;
    qaTrendChart.data.datasets[0].data=avgs;
    qaTrendChart.data.datasets[1].data=target;
    qaTrendChart.data.datasets[2].data=counts;
    qaTrendChart.options.scales.y1.max=Math.ceil(Math.max(...counts,1)/0.4);
    qaTrendChart.update();
    const last2=avgs.slice(-2), tv=last2.length===2?last2[1]-last2[0]:0;
    const badge=document.getElementById('qa-trend-badge');
    if(badge){{
        badge.textContent=avgs.length?(tv>=0?'▲ '+Math.abs(tv).toFixed(1)+'%':'▼ '+Math.abs(tv).toFixed(1)+'%'):'—';
        badge.className='qa-cb '+(tv>=0?'qa-cbg':'qa-cbr');
    }}
}}

function qaRenderCriteria(data) {{
    const el=document.getElementById('qa-criteria-bars'); if(!el) return;
    if(!data.length){{el.innerHTML="<div style='padding:12px;color:#94A3B8;font-size:12px'>No data</div>";return;}}
    const stats=qaComputeCriteria(data).filter(c=>c.pct!==null||c.na).sort((a,b)=>{{
        if(a.na&&!b.na)return 1;if(!a.na&&b.na)return -1;return(a.pct||0)-(b.pct||0);
    }});
    el.innerHTML=stats.map(c=>{{
        if(c.na)return`<div style="margin-bottom:7px"><div style="display:flex;justify-content:space-between;font-size:10px;color:#475569;margin-bottom:2px"><span>${{qaEscapeHtml(c.name)}}</span><span style="color:#94A3B8">N/A</span></div><div style="height:5px;background:#F1F5F9;border-radius:3px"></div></div>`;
        const color=c.pct>=95?'#0F9B58':c.pct>=85?'#F59E0B':'#E85D3F';
        return`<div style="margin-bottom:7px"><div style="display:flex;justify-content:space-between;font-size:10px;color:#475569;margin-bottom:2px"><span>${{qaEscapeHtml(c.name)}}</span><span style="font-weight:700;color:${{color}}">${{c.pct}}%</span></div><div style="height:5px;background:#F1F5F9;border-radius:3px;overflow:hidden"><div style="height:100%;width:${{c.pct}}%;background:${{color}};border-radius:3px;transition:width .4s"></div></div></div>`;
    }}).join('');
    const sub=document.getElementById('qa-crit-sub');
    if(sub)sub.textContent=stats.length+' criteria · sorted by pass rate';
}}

function qaRenderCoaching(data) {{
    const barsEl=document.getElementById('qa-coaching-bars');
    const bdEl=document.getElementById('qa-coach-breakdown');
    const cntEl=document.getElementById('qa-coaching-count');
    if(!barsEl) return;
    const stats=qaComputeCriteria(data).filter(c=>c.pct!==null&&c.pct<95).sort((a,b)=>a.pct-b.pct);
    if(cntEl)cntEl.textContent=stats.length?stats.length+' criteria':'All clear';
    barsEl.innerHTML=stats.length?stats.map(c=>{{
        const color=c.pct>=85?'#F59E0B':'#E85D3F';
        return`<div style="margin-bottom:8px"><div style="display:flex;justify-content:space-between;font-size:11px;color:#1E293B;margin-bottom:3px"><span>${{qaEscapeHtml(c.name)}}</span><span style="font-weight:700;color:${{color}}">${{c.pct}}%</span></div><div style="height:6px;background:#F1F5F9;border-radius:3px;overflow:hidden"><div style="height:100%;width:${{c.pct}}%;background:${{color}};border-radius:3px"></div></div></div>`;
    }}).join(''):`<div style="text-align:center;padding:16px;color:#0F9B58;font-size:12px">✓ All criteria above 95% — great job!</div>`;
    if(bdEl){{
        const byCoach={{}};
        data.forEach(r=>{{if(!r.coach)return;if(!byCoach[r.coach])byCoach[r.coach]={{n:0}};byCoach[r.coach].n++;}});
        bdEl.innerHTML=Object.entries(byCoach).sort((a,b)=>b[1].n-a[1].n).map(([name,d])=>
            `<div style="flex:1;background:#F8FAFC;border-radius:8px;padding:8px;text-align:center"><div style="font-size:10px;color:#64748B;margin-bottom:2px">${{qaEscapeHtml(name)}}</div><div style="font-size:16px;font-weight:800;color:#0D3B6E">${{d.n}}</div><div style="font-size:9px;color:#94A3B8">${{d.n}} eval${{d.n===1?'':'s'}}</div></div>`
        ).join('')||`<div style="color:#94A3B8;font-size:11px">No data</div>`;
    }}
}}

function qaRenderLeaderboard(data) {{
    const el=document.getElementById('qa-leaderboard');
    const badge=document.getElementById('qa-lb-badge');
    if(!el) return;
    const byAgent={{}};
    data.forEach(r=>{{
        if(!r.agent) return;
        const s=Number(r.score);if(isNaN(s)||s<=0)return;
        if(!byAgent[r.agent])byAgent[r.agent]={{scores:[],acct:r._acct||''}};
        byAgent[r.agent].scores.push(s);
    }});
    const agents=Object.entries(byAgent).map(([name,d])=>{{
        const avg=qaAvg(d.scores), pass=d.scores.filter(s=>s>=85).length;
        const words=name.split(' ').slice(0,2).join(' ');
        const av=QA_AV[name]||{{bg:'#F1F5F9',tc:'#475569',ini:name.split(' ').map(w=>w[0]||'').join('').slice(0,2).toUpperCase()}};
        return{{name,words,av,avg,min:Math.min(...d.scores),max:Math.max(...d.scores),pass,passRate:d.scores.length?pass/d.scores.length*100:0,n:d.scores.length,acct:d.acct}};
    }}).sort((a,b)=>b.avg-a.avg);
    if(badge)badge.textContent=agents.length+' agents';
    el.innerHTML=agents.map((a,i)=>{{
        const chipCls=qaChipCls(a.avg);
        const acctPill=a.acct==='M7'
            ?`<span style="background:#EFF6FF;color:#1D4ED8;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">M7</span>`
            :a.acct==='Britelift'
            ?`<span style="background:#FFF7ED;color:#C2410C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Britelift</span>`
            :a.acct==='RideX'
            ?`<span style="background:#F5F3FF;color:#6D28D9;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">RideX</span>`
            :a.acct==='Hamilton'
            ?`<span style="background:#ECFDF5;color:#065F46;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Hamilton</span>`
            :a.acct==='Skyline'
            ?`<span style="background:#F0F9FF;color:#0369A1;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Skyline</span>`
            :`<span style="background:#FFF0F3;color:#9F1239;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Parentis</span>`;
        return`<tr><td style="font-size:11px;color:#94A3B8">${{i+1}}</td><td><div style="display:flex;align-items:center;gap:6px"><span style="width:24px;height:24px;border-radius:50%;background:${{a.av.bg}};color:${{a.av.tc}};font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${{a.av.ini}}</span><span style="font-weight:600;font-size:11px">${{qaEscapeHtml(a.words)}}</span></div></td><td style="text-align:center">${{a.n}}</td><td><span class="qa-chip ${{chipCls}}">${{a.avg.toFixed(1)}}%</span></td><td style="text-align:center;font-size:11px">${{a.min.toFixed(1)}}%</td><td style="text-align:center;font-size:11px">${{a.max.toFixed(1)}}%</td><td style="text-align:center;color:${{a.passRate>=85?'#0F9B58':'#E85D3F'}};font-size:11px">${{a.passRate.toFixed(0)}}%</td><td>${{acctPill}}</td></tr>`;
    }}).join('')||`<tr><td colspan="8" style="text-align:center;color:#94A3B8;padding:16px">No data</td></tr>`;
}}

function qaRenderTLLeaderboard(data) {{
    const el=document.getElementById('qa-tl-leaderboard');
    const badge=document.getElementById('qa-tl-lb-badge');
    if(!el) return;
    const byTL={{}};
    data.forEach(r=>{{
        if(!r.supervisor) return;
        const s=Number(r.score);if(isNaN(s)||s<=0)return;
        if(!byTL[r.supervisor])byTL[r.supervisor]={{scores:[]}};
        byTL[r.supervisor].scores.push(s);
    }});
    const tls=Object.entries(byTL).map(([name,d])=>{{
        const avg=qaAvg(d.scores),pass=d.scores.filter(s=>s>=85).length;
        return{{name,avg,min:Math.min(...d.scores),max:Math.max(...d.scores),pass,passRate:d.scores.length?pass/d.scores.length*100:0,n:d.scores.length}};
    }}).sort((a,b)=>b.avg-a.avg);
    if(badge)badge.textContent=tls.length+' team leaders';
    el.innerHTML=tls.map((t,i)=>{{
        const chipCls=qaChipCls(t.avg);
        return`<tr><td style="font-size:11px;color:#94A3B8">${{i+1}}</td><td style="font-weight:600;font-size:11px">${{qaEscapeHtml(t.name)}}</td><td style="text-align:center">${{t.n}}</td><td><span class="qa-chip ${{chipCls}}">${{t.avg.toFixed(1)}}%</span></td><td style="text-align:center;font-size:11px">${{t.min.toFixed(1)}}%</td><td style="text-align:center;font-size:11px">${{t.max.toFixed(1)}}%</td><td style="text-align:center;color:${{t.passRate>=85?'#0F9B58':'#E85D3F'}};font-size:11px">${{t.passRate.toFixed(0)}}%</td></tr>`;
    }}).join('')||`<tr><td colspan="7" style="text-align:center;color:#94A3B8;padding:16px">No data</td></tr>`;
}}

function qaRenderCoachLeaderboard(data) {{
    const el=document.getElementById('qa-coach-leaderboard');
    const badge=document.getElementById('qa-coach-lb-badge');
    if(!el) return;
    const byCoach={{}};
    data.forEach(r=>{{
        if(!r.coach) return;
        const s=Number(r.score);if(isNaN(s)||s<=0)return;
        if(!byCoach[r.coach])byCoach[r.coach]={{scores:[]}};
        byCoach[r.coach].scores.push(s);
    }});
    const coaches=Object.entries(byCoach).map(([name,d])=>{{
        const avg=qaAvg(d.scores), pass=d.scores.filter(s=>s>=85).length;
        return{{name,avg,min:Math.min(...d.scores),max:Math.max(...d.scores),pass,passRate:d.scores.length?pass/d.scores.length*100:0,n:d.scores.length}};
    }}).sort((a,b)=>b.avg-a.avg);
    if(badge)badge.textContent=coaches.length+' coaches';
    el.innerHTML=coaches.map((c,i)=>{{
        const chipCls=qaChipCls(c.avg);
        return`<tr><td style="font-size:11px;color:#94A3B8">${{i+1}}</td><td style="font-weight:600;font-size:11px">${{qaEscapeHtml(c.name)}}</td><td style="text-align:center">${{c.n}}</td><td><span class="qa-chip ${{chipCls}}">${{c.avg.toFixed(1)}}%</span></td><td style="text-align:center;font-size:11px">${{c.min.toFixed(1)}}%</td><td style="text-align:center;font-size:11px">${{c.max.toFixed(1)}}%</td><td style="text-align:center;color:${{c.passRate>=85?'#0F9B58':'#E85D3F'}};font-size:11px">${{c.passRate.toFixed(0)}}%</td></tr>`;
    }}).join('')||`<tr><td colspan="7" style="text-align:center;color:#94A3B8;padding:16px">No data</td></tr>`;
}}

function qaSortTable(field) {{
    if (qaTableSortState.field === field) {{
        qaTableSortState.dir = qaTableSortState.dir === 'asc' ? 'desc' : 'asc';
    }} else {{
        qaTableSortState.field = field;
        qaTableSortState.dir = (field === 'ts' || field === 'score') ? 'desc' : 'asc';
    }}
    qaRenderTable(qaCurrentFiltered);
}}

// ─── Virtual scroll state for QA detail table ─────────────────────────────────
let qaVsRows=[];
const QA_VS_ROW_H=34;   // px — must match CSS tbody tr height
const QA_VS_BUFFER=12;  // extra rows rendered above/below viewport
let qaVsRaf=null;

function qaRowHtml(r){{
    const score=Number(r.score),sc=!isNaN(score)&&score>0?score:null;
    const chipCls=sc!==null?qaChipCls(sc):'',disp=sc!==null?sc.toFixed(1)+'%':'—';
    const av=QA_AV[r.agent]||{{bg:'#F1F5F9',tc:'#475569',ini:(r.agent||'?').split(' ').map(w=>w[0]||'').join('').slice(0,2).toUpperCase()}};
    const acctPill=r._acct==='M7'
        ?`<span style="background:#EFF6FF;color:#1D4ED8;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">M7</span>`
        :r._acct==='Britelift'
        ?`<span style="background:#FFF7ED;color:#C2410C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Britelift</span>`
        :r._acct==='RideX'
        ?`<span style="background:#F5F3FF;color:#6D28D9;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">RideX</span>`
        :r._acct==='Hamilton'
        ?`<span style="background:#ECFDF5;color:#065F46;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Hamilton</span>`
        :r._acct==='Skyline'
        ?`<span style="background:#F0F9FF;color:#0369A1;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Skyline</span>`
        :`<span style="background:#FFF0F3;color:#9F1239;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Parentis</span>`;
    const critKeys=['os_in','os_out','closing','approp','no_resp','fillers','ack','hold','ack_hold',
                    'resp_eff','empathy','adjust','mute','active','answered','probing','verif',
                    'clarif','lost_sop','rude','trans','speech'];
    const critCells=critKeys.map(k=>{{
        const v=r[k];
        if(k==='rude'){{
            if(v==='No') return'<td style="text-align:center"><span style="color:#0F9B58;font-weight:600">✓</span></td>';
            if(v==='Yes')return'<td style="text-align:center"><span style="color:#E85D3F;font-weight:600">✗</span></td>';
            return'<td style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        }}
        if(k==='lost_sop'&&(!v||v==='Not Applicable'))return'<td style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const fbTxt=qaEscapeHtml(r.feedback||'—');
    return`<tr><td style="white-space:nowrap;font-size:11px">${{qaEscapeHtml((r.ts||'—').slice(0,10))}}</td><td style="max-width:200px;overflow:hidden" title="${{qaEscapeHtml(r.agent||'')}}"><div style="display:flex;align-items:center;gap:5px;overflow:hidden"><span style="width:22px;height:22px;border-radius:50%;background:${{av.bg}};color:${{av.tc}};font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${{av.ini}}</span><span style="font-weight:600;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0">${{qaEscapeHtml(r.agent||'—')}}</span></div></td><td style="font-size:11px;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${{qaEscapeHtml(r.supervisor||'')}}">${{qaEscapeHtml(r.supervisor||'—')}}</td><td style="font-size:11px;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${{qaEscapeHtml(r.coach||'')}}">${{qaEscapeHtml(r.coach||'—')}}</td><td>${{acctPill}}</td><td><span class="qa-chip ${{chipCls}}">${{disp}}</span></td>${{critCells.join('')}}<td style="font-size:10px;color:#475569;white-space:nowrap">${{qaEscapeHtml(r.invest||'—')}}</td><td style="max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:10px;color:#475569" title="${{fbTxt}}">${{fbTxt}}</td></tr>`;
}}

function qaRenderVisibleRows(){{
    const el=document.getElementById('qa-detail-table');
    const scroller=document.getElementById('qa-tbl-scroll-main');
    if(!el||!scroller) return;
    if(!qaVsRows.length){{
        el.innerHTML=`<tr><td colspan="30" style="text-align:center;color:#94A3B8;padding:20px">No data</td></tr>`;
        return;
    }}
    const scrollTop=scroller.scrollTop;
    const viewH=scroller.clientHeight;
    const total=qaVsRows.length;
    const startIdx=Math.max(0,Math.floor(scrollTop/QA_VS_ROW_H)-QA_VS_BUFFER);
    const endIdx=Math.min(total,Math.ceil((scrollTop+viewH)/QA_VS_ROW_H)+QA_VS_BUFFER);
    const topH=startIdx*QA_VS_ROW_H;
    const botH=(total-endIdx)*QA_VS_ROW_H;
    let html='';
    if(topH>0) html+=`<tr style="height:${{topH}}px;pointer-events:none"><td colspan="30" style="padding:0;border:none"></td></tr>`;
    html+=qaVsRows.slice(startIdx,endIdx).map(qaRowHtml).join('');
    if(botH>0) html+=`<tr style="height:${{botH}}px;pointer-events:none"><td colspan="30" style="padding:0;border:none"></td></tr>`;
    el.innerHTML=html;
}}

function qaRenderTable(data) {{
    const el=document.getElementById('qa-detail-table');
    const count=document.getElementById('qa-tbl-count');
    const foot=document.getElementById('qa-tbl-footer');
    if(!el) return;
    const {{field:sf,dir:sd}}=qaTableSortState;
    const smul=sd==='asc'?1:-1;
    qaVsRows=data.slice().sort((a,b)=>{{
        if(sf==='score') return smul*((Number(a.score)||0)-(Number(b.score)||0));
        return smul*(a[sf]||'').localeCompare(b[sf]||'');
    }});
    if(count)count.textContent=qaVsRows.length+' records';
    const total=qaGetActiveData().length;
    if(foot)foot.textContent='Showing '+qaVsRows.length+' of '+total+' evaluations';
    const thead=el.closest('table')?.querySelector('thead');
    if(thead){{
        thead.querySelectorAll('th[data-qa-sort]').forEach(th=>{{
            const ind=th.querySelector('.sort-indicator');
            if(!ind) return;
            if(th.dataset.qaSort===sf){{
                ind.textContent=sd==='asc'?'▲':'▼';
                th.style.color='#0D3B6E';
            }} else {{
                ind.textContent='';
                th.style.color='';
            }}
        }});
    }}
    const scroller=document.getElementById('qa-tbl-scroll-main');
    if(scroller) scroller.scrollTop=0;
    qaRenderVisibleRows();
}}

function qaUpdateDonut(data) {{
    if(!qaDonutChart) return;
    const scores=data.map(r=>Number(r.score)).filter(v=>!isNaN(v)&&v>0);
    const buckets=QA_BANDS.map(b=>scores.filter(s=>s>=b.min&&s<=b.max).length);
    qaDonutChart.data.datasets[0].data=buckets;
    qaDonutChart.update();
    const sub=document.getElementById('qa-donut-sub');
    if(sub)sub.textContent=data.length+' evaluation'+(data.length===1?'':'s');
    const legEl=document.getElementById('qa-donut-legend');
    if(legEl){{
        const total=scores.length;
        legEl.innerHTML=QA_BANDS.map((b,i)=>{{
            const pct=total?(buckets[i]/total*100).toFixed(1)+'%':'0%';
            return`<div style="display:flex;align-items:center;gap:4px;font-size:10px;color:#475569"><span style="width:8px;height:8px;border-radius:50%;background:${{b.color}};flex-shrink:0"></span><span>${{b.label}}</span><span style="margin-left:auto;font-weight:700;color:${{b.color}}">${{pct}}</span></div>`;
        }}).join('');
    }}
}}

function qaUpdateEvalDist(data) {{
    if(!qaEvalDistChart) return;
    const accts=[
        {{key:'M7',color:'#4F81BD'}},
        {{key:'Parentis',color:'#2C3E8C'}},
        {{key:'Britelift',color:'#C0392B'}},
        {{key:'RideX',color:'#8E44AD'}},
        {{key:'Hamilton',color:'#065F46'}},
        {{key:'Skyline',color:'#0EA5E9'}}
    ];
    const counts=accts.map(a=>data.filter(r=>r._acct===a.key).length);
    const total=counts.reduce((s,v)=>s+v,0);
    qaEvalDistChart.data.datasets[0].data=counts;
    qaEvalDistChart.update();
    const sub=document.getElementById('qa-donut-sub');
    if(sub) sub.textContent=total+' evaluation'+(total===1?'':'s')+' across all accounts';
    const legEl=document.getElementById('qa-eval-dist-legend');
    if(legEl){{
        legEl.innerHTML=accts.map((a,i)=>{{
            const pct=total?(counts[i]/total*100).toFixed(1)+'%':'0%';
            return`<div style="display:flex;align-items:center;gap:4px;font-size:10px;color:#475569"><span style="width:8px;height:8px;border-radius:50%;background:${{a.color}};flex-shrink:0"></span><span>${{a.key}}</span><span style="margin-left:auto;font-weight:700;color:${{a.color}}">${{pct}}</span></div>`;
        }}).join('');
    }}
}}

function qaUpdateHamiltonCrit(data) {{
    const crits=[
        {{key:'os_out',  valId:'qa-hcrit-greet-val',subId:'qa-hcrit-greet-sub'}},
        {{key:'adjust',  valId:'qa-hcrit-prof-val', subId:'qa-hcrit-prof-sub'}},
        {{key:'gen_q',   valId:'qa-hcrit-genq-val', subId:'qa-hcrit-genq-sub'}},
        {{key:'verif',   valId:'qa-hcrit-verif-val',subId:'qa-hcrit-verif-sub'}},
        {{key:'resp_eff',valId:'qa-hcrit-resol-val',subId:'qa-hcrit-resol-sub'}},
    ];
    crits.forEach(c=>{{
        const rel=data.filter(r=>r[c.key]!=null&&r[c.key]!=='');
        const passed=rel.filter(r=>r[c.key]==='Yes').length;
        const pct=rel.length?(passed/rel.length*100).toFixed(1)+'%':'—';
        const vEl=document.getElementById(c.valId);
        const sEl=document.getElementById(c.subId);
        if(vEl)vEl.textContent=pct;
        if(sEl)sEl.textContent=rel.length?passed+' of '+rel.length+' passed':'No data';
    }});
}}

function qaUpdateSkylineCrit(data) {{
    const crits=[
        {{key:'os_out',  valId:'qa-scrit-greet-val',subId:'qa-scrit-greet-sub'}},
        {{key:'adjust',  valId:'qa-scrit-prof-val', subId:'qa-scrit-prof-sub'}},
        {{key:'gen_q',   valId:'qa-scrit-genq-val', subId:'qa-scrit-genq-sub'}},
        {{key:'verif',   valId:'qa-scrit-verif-val',subId:'qa-scrit-verif-sub'}},
        {{key:'resp_eff',valId:'qa-scrit-resol-val',subId:'qa-scrit-resol-sub'}},
        {{key:'speech',  valId:'qa-scrit-comm-val', subId:'qa-scrit-comm-sub'}},
    ];
    crits.forEach(c=>{{
        const rel=data.filter(r=>r[c.key]!=null&&r[c.key]!=='');
        const passed=rel.filter(r=>r[c.key]==='Yes').length;
        const pct=rel.length?(passed/rel.length*100).toFixed(1)+'%':'—';
        const vEl=document.getElementById(c.valId);
        const sEl=document.getElementById(c.subId);
        if(vEl)vEl.textContent=pct;
        if(sEl)sEl.textContent=rel.length?passed+' of '+rel.length+' passed':'No data';
    }});
}}

function qaUpdateAivh(data) {{
    const corrected=data.filter(r=>r.status==='corrected'&&r.score_ai!=null&&r.score_human!=null&&r.score_human!=='');
    const total=data.length;
    const nCorr=corrected.length;
    const allAiVals=data.filter(r=>r.score_ai!=null).map(r=>Number(r.score_ai));
    const avgAi=allAiVals.length?allAiVals.reduce((s,v)=>s+v,0)/allAiVals.length:0;
    const avgHu=nCorr?corrected.reduce((s,r)=>s+Number(r.score_human),0)/nCorr:0;
    const avgGap=avgHu-avgAi;

    // Gap categories
    const gaps=corrected.map(r=>Number(r.score_human)-Number(r.score_ai));
    const nHgtA=gaps.filter(g=>g>0).length;
    const nEq=gaps.filter(g=>g===0).length;
    const nAgtH=gaps.filter(g=>g<0).length;
    const maxGap=gaps.length?Math.max(...gaps):0;
    const minGap=gaps.length?Math.min(...gaps):0;
    const pHgtA=nCorr?Math.round(nHgtA/nCorr*100):0;
    const pEq=nCorr?Math.round(nEq/nCorr*100):0;
    const pAgtH=nCorr?Math.round(nAgtH/nCorr*100):0;

    // --- KPI tile strip ---
    const strip=document.getElementById('qa-aivh-kpi-strip');
    if(strip){{
        const tiles=[
            {{label:'AI AVG SCORE',val:avgAi>0?avgAi.toFixed(1)+'%':'—',sub:total+' total evals',bg:'#1E3A5F',fg:'#BFDBFE',vfg:'#EFF6FF'}},
            {{label:'HUMAN AVG SCORE',val:nCorr?avgHu.toFixed(1)+'%':'—',sub:(avgGap>=0?'+':'')+avgGap.toFixed(1)+'% vs AI',bg:'#064E3B',fg:'#6EE7B7',vfg:'#ECFDF5'}},
            {{label:'AVG GAP (H-AI)',val:nCorr?(avgGap>=0?'+':'')+avgGap.toFixed(1)+'%':'—',sub:'Human scored higher',bg:'#134E4A',fg:'#99F6E4',vfg:'#F0FDFA'}},
            {{label:'HUMAN > AI',val:nHgtA,sub:pHgtA+'% of corrected',bg:'#3B0764',fg:'#D8B4FE',vfg:'#FAF5FF'}},
            {{label:'EXACT MATCH',val:nEq,sub:pEq+'% aligned',bg:'#78350F',fg:'#FDE68A',vfg:'#FFFBEB'}},
            {{label:'AI > HUMAN',val:nAgtH,sub:pAgtH+'% — AI scored higher',bg:'#7F1D1D',fg:'#FCA5A5',vfg:'#FEF2F2'}},
        ];
        strip.innerHTML=tiles.map(t=>`
            <div style="background:${{t.bg}};border-radius:10px;padding:12px 14px;display:flex;flex-direction:column;gap:4px">
                <span style="font-size:9px;font-weight:700;color:${{t.fg}};text-transform:uppercase;letter-spacing:.05em">${{t.label}}</span>
                <span style="font-size:20px;font-weight:800;color:${{t.vfg}};line-height:1.1">${{t.val}}</span>
                <span style="font-size:9px;color:${{t.fg}};opacity:.8">${{t.sub}}</span>
            </div>`).join('');
    }}

    // --- Gap Distribution donut ---
    const gapList=document.getElementById('qa-aivh-gap-list');
    if(gapList){{
        const cats=[
            {{label:'Human > AI',n:nHgtA,pct:pHgtA,color:'#1D4ED8'}},
            {{label:'Exact Match',n:nEq,pct:pEq,color:'#94A3B8'}},
            {{label:'AI > Human',n:nAgtH,pct:pAgtH,color:'#EF4444'}},
        ];
        gapList.innerHTML=cats.map(c=>`
            <div style="display:flex;align-items:center;gap:6px">
                <span style="width:8px;height:8px;border-radius:50%;background:${{c.color}};flex-shrink:0"></span>
                <span style="color:#475569;flex:1">${{c.label}}</span>
                <span style="font-weight:700;color:#1E293B">${{c.n}}</span>
                <span style="color:#94A3B8">(${{c.pct}}%)</span>
            </div>`).join('');
    }}
    const gapRange=document.getElementById('qa-aivh-gap-range');
    if(gapRange){{
        gapRange.innerHTML=nCorr
            ?`Max gap: <strong>+${{maxGap.toFixed(1)}}%</strong> &nbsp; Min: <strong>${{(minGap>=0?'+':'')+minGap.toFixed(1)}}%</strong>`
            :'No corrected evals';
    }}
    if(qaAivhGapDonut){{
        qaAivhGapDonut.data.datasets[0].data=[nHgtA,nEq,nAgtH];
        qaAivhGapDonut.update();
    }}

    // --- Summary card chips + insight ---
    const chips=document.getElementById('qa-aivh-chips');
    if(chips){{
        const chipData=[
            {{label:'AI avg',val:nCorr?avgAi.toFixed(1)+'%':'—',bg:'#1E3A8A',fg:'#BFDBFE'}},
            {{label:'Human avg',val:nCorr?avgHu.toFixed(1)+'%':'—',bg:'#065F46',fg:'#6EE7B7'}},
            {{label:'Avg gap',val:nCorr?(avgGap>=0?'+':'')+avgGap.toFixed(1)+'%':'—',bg:'#134E4A',fg:'#99F6E4'}},
        ];
        chips.innerHTML=chipData.map(c=>`<div style="background:${{c.bg}};border-radius:6px;padding:6px 10px;display:flex;flex-direction:column;gap:1px"><span style="font-size:9px;color:${{c.fg}};font-weight:600;text-transform:uppercase;letter-spacing:.04em">${{c.label}}</span><span style="font-size:14px;font-weight:800;color:#F8FAFC">${{c.val}}</span></div>`).join('');
    }}
    const insight=document.getElementById('qa-aivh-insight');
    if(insight){{
        if(nCorr){{
            const topDir=nHgtA>nAgtH?'Human QA scored higher than AI':'AI scored higher than Human QA';
            const pctTop=nHgtA>nAgtH?pHgtA:pAgtH;
            insight.innerHTML=`<span style="color:#FBBF24;margin-right:6px">&#8226;</span>In <strong style="color:#F8FAFC">${{pctTop}}%</strong> of corrected evals, ${{topDir}}. The average correction added <strong style="color:#F8FAFC">${{(avgGap>=0?'+':'')+avgGap.toFixed(1)}} pts</strong>, indicating QA reviewers tend to ${{avgGap>=0?'raise':'lower'}} scores after human review.`;
        }} else {{
            insight.textContent='No corrected evaluations in the selected period.';
        }}
    }}
}}

function qaToggleDistWidget(isAllAccounts) {{
    const title=document.getElementById('qa-dist-title');
    const scoreWrap=document.getElementById('qa-score-dist-wrap');
    const evalWrap=document.getElementById('qa-eval-dist-wrap');
    if(!title||!scoreWrap||!evalWrap) return;
    if(isAllAccounts){{
        title.textContent='Evaluation distribution';
        scoreWrap.style.display='none';
        evalWrap.style.display='';
    }} else {{
        title.textContent='Score distribution';
        scoreWrap.style.display='';
        evalWrap.style.display='none';
    }}
}}

// ─── Main filter apply ────────────────────────────────────────────────────────
function qaApplyFilters() {{
    const coach=(document.getElementById('qa-sel-coach')?.value||'').trim();
    const agent=(document.getElementById('qa-sel-agent')?.value||'').trim();
    const head=(document.getElementById('qa-sel-head')?.value||'').trim();
    const startStr=qaFmtDate(qaDrpStart), endStr=qaFmtDate(qaDrpEnd);
    const pool=qaGetActiveData();
    const filtered=pool.filter(r=>{{
        const rDate=(r.ts||'').slice(0,10);
        return rDate>=startStr&&rDate<=endStr&&
            (!coach||r.coach===coach)&&
            (!agent||r.agent===agent)&&
            (!head||r.supervisor===head);
    }});
    qaCurrentFiltered=filtered;
    qaUpdateKPIs(filtered);
    qaUpdateTrend(filtered);
    qaRenderCriteria(filtered);
    qaRenderCoaching(filtered);
    qaRenderLeaderboard(filtered);
    qaRenderTLLeaderboard(filtered);
    qaRenderCoachLeaderboard(filtered);
    qaRenderTable(filtered);
    const acct=(document.getElementById('qa-sel-account')?.value||'').trim();
    qaToggleDistWidget(!acct);
    if(acct) qaUpdateDonut(filtered); else qaUpdateEvalDist(filtered);
    const aivhCard=document.getElementById('qa-aivh-card');
    if(aivhCard)aivhCard.style.display=(acct==='hamilton'||acct==='skyline')?'':'none';
    if(acct==='hamilton'||acct==='skyline')qaUpdateAivh(filtered);
    const hCritEl=document.getElementById('qa-hamilton-crit');
    if(hCritEl)hCritEl.style.display=(acct==='hamilton')?'':'none';
    if(acct==='hamilton')qaUpdateHamiltonCrit(filtered);
    const sCritEl=document.getElementById('qa-skyline-crit');
    if(sCritEl)sCritEl.style.display=(acct==='skyline')?'':'none';
    if(acct==='skyline')qaUpdateSkylineCrit(filtered);
    const sumStripMain=document.getElementById('qa-sum-strip-main');
    if(sumStripMain)sumStripMain.style.display=(acct==='hamilton'||acct==='skyline')?'none':'';
    const scores=filtered.map(r=>Number(r.score)).filter(v=>!isNaN(v)&&v>0);
    const agents=new Set(filtered.map(r=>r.agent).filter(Boolean));
    const avg=scores.length?qaAvg(scores):null;
    const pass=scores.length?scores.filter(s=>s>=85).length/scores.length*100:null;
    const below=scores.filter(s=>s<85).length;
    const low=scores.filter(s=>s>=85&&s<95).length;
    const sset=(id,v)=>{{const el=document.getElementById(id);if(el)el.textContent=v;}};
    sset('qa-strip-avg',  avg?avg.toFixed(1)+'%':'—');
    sset('qa-strip-pass', pass!==null?pass.toFixed(1)+'%':'—');
    sset('qa-strip-evals',filtered.length);
    sset('qa-strip-agents',agents.size);
    sset('qa-strip-below',below);
    sset('qa-strip-low',  low);
}}

function qaClearFilters() {{
    const acctEl=document.getElementById('qa-sel-account');
    if(acctEl) acctEl.value='';
    ['qa-sel-coach','qa-sel-agent','qa-sel-head'].forEach(id=>{{
        const el=document.getElementById(id);if(el)el.value='';
    }});
    qaPopulateSelects();
    qaResetDRP();
}}

// ─── Chart initialization ─────────────────────────────────────────────────────
function initQualityCharts() {{
    if(qaChartsInitialized) return;
    qaChartsInitialized=true;

    document.addEventListener('click',function(e){{
        if(!qaDrpOpen) return;
        if(qaDrpPhase===1) return;
        if(!e.target.isConnected) return;
        const wrap=document.getElementById('qa-drp-wrap');
        if(wrap&&!wrap.contains(e.target)) qaCloseDatePicker();
    }});

    // measure main header and set sticky top offset
    const mainHdr=document.querySelector('.sticky-dashboard-header');
    if(mainHdr){{
        const panel=document.getElementById('qualityPanel');
        if(panel) panel.style.setProperty('--qa-stick-top', mainHdr.offsetHeight+'px');
    }}

    const sentinel=document.getElementById('qa-kpi-sentinel');
    if(sentinel){{
        const strip=document.getElementById('qa-kpi-strip');
        const stickyCtrl=document.getElementById('qa-sticky-ctrl');
        const obs=new IntersectionObserver(entries=>{{
            entries.forEach(en=>{{
                const off=!en.isIntersecting;
                if(strip) strip.classList.toggle('visible',off);
                if(stickyCtrl) stickyCtrl.classList.toggle('scrolled',off);
            }});
        }},{{threshold:0}});
        obs.observe(sentinel);
    }}

    const qaTblScroller=document.getElementById('qa-tbl-scroll-main');
    if(qaTblScroller){{
        qaTblScroller.addEventListener('scroll',()=>{{
            if(qaVsRaf) cancelAnimationFrame(qaVsRaf);
            qaVsRaf=requestAnimationFrame(qaRenderVisibleRows);
        }},{{passive:true}});
    }}

    const trendLabelPlugin={{
        id:'qaTrendLabels',
        afterDatasetsDraw(chart){{
            const ds=chart.data.datasets[0],meta=chart.getDatasetMeta(0),ctx2=chart.ctx;
            ctx2.save();ctx2.font='700 9px sans-serif';ctx2.fillStyle='#0D3B6E';ctx2.textAlign='center';
            meta.data.forEach((pt,i)=>{{const v=ds.data[i];if(v!=null)ctx2.fillText(v.toFixed(1)+'%',pt.x,pt.y-8);}});
            ctx2.restore();
        }}
    }};
    const barLabelPlugin={{
        id:'qaBarLabels',
        afterDatasetsDraw(chart){{
            const ds=chart.data.datasets[2],meta=chart.getDatasetMeta(2),ctx2=chart.ctx;
            if(!meta||!ds) return;
            ctx2.save();ctx2.font='bold 9px sans-serif';ctx2.fillStyle='#166534';ctx2.textAlign='center';ctx2.textBaseline='top';
            meta.data.forEach((bar,i)=>{{
                const v=ds.data[i];if(v==null||v===0)return;
                ctx2.fillText(v,bar.x,bar.y+4);
            }});
            ctx2.restore();
        }}
    }};
    const donutLabelPlugin={{
        id:'qaDonutLabels',
        afterDatasetsDraw(chart){{
            const ds=chart.data.datasets[0],meta=chart.getDatasetMeta(0),ctx2=chart.ctx;
            const total=ds.data.reduce((a,b)=>a+(b||0),0);
            ctx2.save();ctx2.font='700 9px sans-serif';ctx2.textAlign='center';
            meta.data.forEach((arc,i)=>{{
                const v=ds.data[i];if(!v)return;
                const angle=(arc.startAngle+arc.endAngle)/2,r=(arc.outerRadius+arc.innerRadius)/2;
                const x=arc.x+Math.cos(angle)*r,y=arc.y+Math.sin(angle)*r;
                ctx2.fillStyle='#fff';
                const pct=total?Math.round(v/total*100):0;
                if(pct>=5)ctx2.fillText(pct+'%',x,y+3);
            }});
            ctx2.restore();
        }}
    }};

    qaPopulateSelects();
    qaUpdateDRPLabel();

    // Info pills
    const qaAcctsLoaded=[qaRawData,parentisRawData,briteliftRawData,ridexRawData,hamiltonRawData,skylineRawData].filter(d=>d.length>0).length;
    const qaPillAccts=document.getElementById('qa-pill-qa-accounts');
    if(qaPillAccts)qaPillAccts.textContent=qaAcctsLoaded+' QA Account'+(qaAcctsLoaded===1?'':'s')+' Loaded';
    const qaPillTotal=document.getElementById('qa-pill-total-accounts');
    if(qaPillTotal){{
        const n=Number(document.getElementById('approvedAccounts')?.textContent)||0;
        if(n>0)qaPillTotal.textContent=n+' Accounts';
    }}

    try {{

    const trendCtx=document.getElementById('qa-trend-chart');
    if(trendCtx){{
        qaTrendChart=new Chart(trendCtx,{{
            type:'line',
            plugins:[trendLabelPlugin,barLabelPlugin],
            data:{{labels:[],datasets:[
                {{label:'Avg QA Score',data:[],borderColor:'#0D3B6E',backgroundColor:'rgba(13,59,110,0.08)',tension:0.3,fill:true,pointRadius:4,pointHoverRadius:6,pointBackgroundColor:'#0D3B6E',yAxisID:'y',order:0}},
                {{label:'Target (85%)',data:[],borderColor:'#E85D3F',borderDash:[5,4],borderWidth:1.5,pointRadius:0,fill:false,yAxisID:'y',order:0}},
                {{type:'bar',label:'Total Evaluations',data:[],backgroundColor:'rgba(57,181,74,0.55)',borderColor:'#39B54A',borderWidth:1,yAxisID:'y1',barPercentage:0.6,categoryPercentage:0.7,order:1}}
            ]}},
            options:{{
                responsive:true,maintainAspectRatio:false,
                layout:{{padding:{{top:24,bottom:0}}}},
                plugins:{{legend:{{display:true,position:'bottom',labels:{{font:{{size:10}},boxWidth:12}}}},tooltip:{{callbacks:{{label:ctx=>ctx.datasetIndex===2?ctx.parsed.y+' evals':ctx.parsed.y.toFixed(1)+'%'}}}}}},
                scales:{{
                    y:{{min:80,max:100,ticks:{{display:false}},border:{{display:false}},grid:{{color:'#F1F5F9'}}}},
                    y1:{{display:false,position:'right',beginAtZero:true,grid:{{display:false}}}},
                    x:{{ticks:{{font:{{size:9}},maxRotation:30}},grid:{{display:false}}}}
                }}
            }}
        }});
    }}
    const donutCtx=document.getElementById('qa-donut-chart');
    if(donutCtx){{
        qaDonutChart=new Chart(donutCtx,{{
            type:'doughnut',
            plugins:[donutLabelPlugin],
            data:{{labels:QA_BANDS.map(b=>b.label),datasets:[{{data:[0,0,0,0],backgroundColor:QA_BANDS.map(b=>b.color),borderWidth:2,borderColor:'#fff'}}]}},
            options:{{responsive:true,maintainAspectRatio:false,cutout:'50%',layout:{{padding:8}},plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>`${{ctx.label}}: ${{ctx.parsed}}`}}}}}}}}
        }});
    }}
    // Slices below 2 % of total get no outside label (still shown in tooltip/legend).
    // All visible labels go into two fixed columns — one left, one right — so they are
    // always outside the arc ring regardless of slice angle.  Collisions are resolved
    // by spreading labels vertically within each column.
    const evalDistLabelPlugin={{
        id:'qaEvalDistLabels',
        afterDatasetsDraw(chart){{
            const ds=chart.data.datasets[0],meta=chart.getDatasetMeta(0),ctx2=chart.ctx;
            const total=ds.data.reduce((a,b)=>a+(b||0),0);
            if(!total)return;
            const validArcs=meta.data.filter(a=>a.outerRadius>0);
            if(!validArcs.length)return;
            const CX=validArcs[0].x, CY=validArcs[0].y, OR=validArcs[0].outerRadius;
            const area=chart.chartArea;
            ctx2.save();

            // Pass 1 — count inside arc (skip arcs too thin to read)
            meta.data.forEach((arc,i)=>{{
                const v=ds.data[i];if(!v)return;
                if(arc.endAngle-arc.startAngle<0.3)return;
                const mid=(arc.startAngle+arc.endAngle)/2;
                const r=(arc.innerRadius+arc.outerRadius)/2;
                ctx2.font='bold 13px sans-serif';ctx2.fillStyle='#fff';
                ctx2.textAlign='center';ctx2.textBaseline='middle';
                ctx2.fillText(v,CX+Math.cos(mid)*r,CY+Math.sin(mid)*r);
            }});

            // Pass 2 — collect outside-label candidates (>= 2 % of total)
            const lbs=[];
            meta.data.forEach((arc,i)=>{{
                const v=ds.data[i];if(!v||v/total<0.02)return;
                const mid=(arc.startAngle+arc.endAngle)/2;
                const cm=Math.cos(mid),sm=Math.sin(mid);
                lbs.push({{i,v,mid,cm,sm,
                    label:chart.data.labels[i],color:ds.backgroundColor[i],
                    right:cm>=0,y:CY+sm*(OR+18)}});
            }});

            // Fixed columns — always clear of the arc ring
            const COL_OFF=24, TICK=9;
            const rColX=CX+OR+COL_OFF, lColX=CX-OR-COL_OFF;

            // Pass 3 — deterministic spread: centre labels around their idealY centroid,
            // evenly spaced by MIN_GAP, then clamp the whole group inside the chart area.
            const MIN_GAP=17;
            for(const isRight of [true,false]){{
                const grp=lbs.filter(l=>l.right===isRight).sort((a,b)=>a.y-b.y);
                if(!grp.length)continue;
                const n=grp.length;
                const span=(n-1)*MIN_GAP;
                const centY=grp.reduce((s,l)=>s+l.y,0)/n;
                let startY=centY-span/2;
                startY=Math.max(area.top+6,Math.min(area.bottom-6-span,startY));
                grp.forEach((lb,i)=>{{lb.y=startY+i*MIN_GAP;}});
            }}

            // Pass 4 — draw leader line and label
            lbs.forEach(lb=>{{
                const dir=lb.right?1:-1;
                const colX=lb.right?rColX:lColX;
                const tickEndX=colX+dir*TICK;
                const ax=CX+lb.cm*OR, ay=CY+lb.sm*OR;   // arc surface
                const sx=CX+lb.cm*(OR+5), sy=CY+lb.sm*(OR+5); // 5 px radial stub

                ctx2.beginPath();
                ctx2.moveTo(ax,ay);
                ctx2.lineTo(sx,sy);         // radial stub
                ctx2.lineTo(colX,lb.y);     // diagonal to column at resolved y
                ctx2.lineTo(tickEndX,lb.y); // horizontal tick
                ctx2.strokeStyle=lb.color;ctx2.lineWidth=1.5;ctx2.stroke();

                ctx2.font='bold 10px sans-serif';ctx2.fillStyle=lb.color;
                ctx2.textAlign=lb.right?'left':'right';ctx2.textBaseline='middle';
                ctx2.fillText(lb.label,tickEndX+dir*3,lb.y);
            }});

            ctx2.restore();
        }}
    }};
    const evalDistCtx=document.getElementById('qa-eval-dist-chart');
    if(evalDistCtx){{
        qaEvalDistChart=new Chart(evalDistCtx,{{
            type:'doughnut',
            plugins:[evalDistLabelPlugin],
            data:{{labels:['M7','Parentis','Britelift','RideX','Hamilton','Skyline'],datasets:[{{data:[0,0,0,0,0,0],backgroundColor:['#4F81BD','#2C3E8C','#C0392B','#8E44AD','#065F46','#0EA5E9'],borderWidth:2,borderColor:'#fff'}}]}},
            options:{{
                responsive:true,maintainAspectRatio:false,cutout:'50%',
                layout:{{padding:{{top:28,bottom:28,left:28,right:28}}}},
                plugins:{{
                    legend:{{display:false}},
                    tooltip:{{callbacks:{{label:ctx=>{{
                        const tot=ctx.dataset.data.reduce((a,b)=>a+(b||0),0);
                        const pct=tot?(ctx.dataset.data[ctx.dataIndex]/tot*100).toFixed(1)+'%':'0%';
                        return`${{ctx.label}}: ${{ctx.dataset.data[ctx.dataIndex]}} evaluation${{ctx.dataset.data[ctx.dataIndex]===1?'':'s'}} (${{pct}})`;
                    }}}}}}
                }}
            }}
        }});
    }}

    const aivhDonutCtx=document.getElementById('qa-aivh-gap-donut');
    if(aivhDonutCtx){{
        qaAivhGapDonut=new Chart(aivhDonutCtx,{{
            type:'doughnut',
            data:{{
                labels:['Human > AI','Equal','AI > Human'],
                datasets:[{{data:[0,0,0],backgroundColor:['#1D4ED8','#94A3B8','#EF4444'],borderWidth:2,borderColor:'#fff'}}]
            }},
            options:{{
                responsive:true,maintainAspectRatio:false,cutout:'68%',
                plugins:{{
                    legend:{{display:false}},
                    tooltip:{{callbacks:{{label:ctx=>`${{ctx.label}}: ${{ctx.parsed}} (${{ctx.dataset.data.reduce((a,b)=>a+b,0)?Math.round(ctx.parsed/ctx.dataset.data.reduce((a,b)=>a+b,0)*100):0}}%)`}}}}
                }}
            }}
        }});
    }}

    }} catch(e) {{
        console.warn('Quality chart init failed:', e);
    }}

    qaApplyFilters();
}}

function setQAFocusMode(active) {{
    document.body.classList.toggle('qa-focus-mode',active);
    const btn=document.getElementById('qa-focus-toggle');
    if(btn){{
        btn.setAttribute('aria-label',active?'Collapse table':'Expand table');
        btn.setAttribute('title',active?'Collapse table':'Expand table');
    }}
}}

function toggleQAFocusMode() {{
    setQAFocusMode(!document.body.classList.contains('qa-focus-mode'));
}}

function downloadQAExcel() {{
    const rows=qaCurrentFiltered.map(r=>{{
        const obj={{'Evaluation Date':(r.ts||'').slice(0,10),'Emp Name':r.agent||'','Account':r._acct||'','Immediate Head':r.supervisor||'','QA Coach':r.coach||'','Score':r.score||''}};
        QA_CRIT_META.forEach(c=>{{obj[c.name]=r[c.key]||'';}});
        obj['Investigation']=r.invest||'';obj['Feedback Summary']=r.feedback||'';
        return obj;
    }});
    const dateTag=new Date().toISOString().slice(0,10);
    const acct=(document.getElementById('qa-sel-account')?.value||'all').trim()||'all';
    const fname=`qa_${{acct}}_evaluations_${{dateTag}}`;
    if(window.XLSX){{
        const ws=XLSX.utils.json_to_sheet(rows);
        ws['!cols']=Object.keys(rows[0]||{{}}).map(k=>({{wch:k==='Feedback Summary'?50:k==='Emp Name'||k==='Immediate Head'?22:Math.max(14,k.length+2)}}));
        const wb=XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb,ws,'QA Evaluations');
        XLSX.writeFile(wb,fname+'.xlsx');
    }} else {{
        const headers=Object.keys(rows[0]||{{}});
        const headerHtml=headers.map(h=>`<th>${{h}}</th>`).join('');
        const rowHtml=rows.map(r=>`<tr>${{headers.map(h=>`<td>${{String(r[h]||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}}</td>`).join('')}}</tr>`).join('');
        const blob=new Blob([`<html><head><meta charset="UTF-8"></head><body><table><thead><tr>${{headerHtml}}</tr></thead><tbody>${{rowHtml}}</tbody></table></body></html>`],{{type:'application/vnd.ms-excel'}});
        const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=fname+'.xls';a.click();
    }}
}}

// ─── END QA QUALITY TAB ───────────────────────────────────────────────────────

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
    if (!isMasterlist) {{
        setMasterlistFocusMode(false);
    }}
    if (tabName !== "quality") {{
        setQAFocusMode(false);
    }}

    if (isMasterlist || isCoaching) {{
        setTimeout(reflowDashboard, 0);
    }}

    if (tabName === "quality") {{
        setTimeout(initQualityCharts, 0);
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
document.getElementById("masterlistFocusToggle")?.addEventListener("click", toggleMasterlistFocusMode);
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

import base64
import hashlib
import html
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd

import notify
from sheets_urls import MASTERLIST_CSV, HISTORY_CSV, MOVEMENT_CSV

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

DMG_DIR = Path(os.getenv("DMG_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\DMG"))
DMG_SCRIPT = DMG_DIR / "dmg_pull.py"
DMG_OUTPUT_FILE = DMG_DIR / "DMG_RAW.xlsx"

R4H_DIR = Path(os.getenv("R4H_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\R4H"))
R4H_SCRIPT = R4H_DIR / "r4h_pull.py"
R4H_OUTPUT_FILE = R4H_DIR / "R4H_RAW.xlsx"

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

DMG_COLUMN_MAP = {
    "Timestamp": "ts",
    "Emp Name": "agent",
    "Employee Name": "agent",
    "Score": "score",
    "Evaluation Type": "type",
    "Reviewer": "coach",
    "QA": "coach",
    "Immediate Supervisor": "supervisor",
    "Ticket ID": "invest",
    "OVERALL FEEDBACK": "feedback",
    "QA_ID": "qa_id",
    "EMPLOYEE_ID": "emp_id",
    "1.1 Ticket Claimed & Assigned (3 pts)\nThe ticket is claimed and the correct organization is assigned in the CRM.": "invest",
    "1.2 Subject Line Edited (3 pts)\nThe subject clearly reflects the issue; chat labels preserved.": "clarif",
    "1.3 Internal Edits (Live Chat) (4 pts)\nAgent removes their email and inputs org name on closure.": "hold",
    "2.1 Professional Yet Personable Tone (5 pts)\nFriendly, natural tone that avoids robotic or curt replies.": "approp",
    "2.2 Grammar & Clarity (5 pts)\nFree from grammar errors; easy to read.": "speech",
    "2.3 Greeting & Sign-off Present (5 pts)\nOpening and polite closing included in reply.": "os_in",
    "3.1 Clarifying Questions When Needed (5 pts)\nAgent seeks into when something is unclear.": "probing",
    "3.2 No Speculation (5 pts)\nThe agent avoids overpromising or guessing.": "no_resp",
    "3.3 Accurate Information (10 pts)\nThe agent provides correct answers per SOP/FYI.": "answered",
    "4. Full Ticket Context (5 pts)\nThe customer knows the current step, next step, and flow.": "active",
    "4.2 Follow-through Statements (5 pts)\nIndicates intent to follow up or update customers.": "resp_eff",
    "4.3 Realistic Expectations (5 pts)\nNo unverified deadlines or pricing promises.": "adjust",
    "5.1 Proper Escalation (5 pts)\nClear and complete internal escalation that includes supporting evidence, and replication steps.": "lost_sop",
    "5.2 Offer of Further Help (5 pts)\nIf replies aren’t resolving the issue, offer alternate support.": "closing",
    "5.3 Ticket Summary (5 pts)\nClear internal summary of the issue, actions taken, and final resolution.": "ack_hold",
    "5.4 Resolution Clarity (5 pts)\nThe agent clearly communicates to the customer what was done and confirms the final outcome.": "trans",
    "6.1 CRM Notes (10 pts)\nThe agent adds internal comments reflecting actions": "ack",
    "6.2 Ticket Structure & Dupes (10 pts)\nRelated tickets merged; unrelated issues split.": "verif",
}

# R4H pulls from its own source form (not M7's) — several criterion headers share
# the same wording as M7's but with different point values ("- 1pt" vs "- 2pts"),
# so reusing M7_COLUMN_MAP silently dropped those criteria for R4H. Dedicated map
# built from R4H_RAW.xlsx's real headers (2026-07-11).
R4H_COLUMN_MAP = {
    "Timestamp":                                              "ts",
    "Call Date:":                                            "date",
    "Emp Name":                                              "agent",
    "Score":                                                 "score",
    "Evaluation Type:":                                      "type",
    "QA":                                                    "coach",
    "Immediate Supervisor":                                  "supervisor",
    "Conducted Thorough Investigation:":                     "invest",
    "Feedback Summary:":                                     "feedback",
    "Opening Spiel (Inbound Calls) - 2pts":                  "os_in",
    "Opening Spiel (Outbound Calls) - 1pt":                  "os_out",
    "Closing Spiel - 1pt":                                   "closing",
    "Appropriate Response - 2pts":                           "approp",
    "No Response - 2pts":                                    "no_resp",
    "Fillers / Slang Words - 1pt":                           "fillers",
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
    "Rudeness - 20pts":                                      "rude",
    "Transaction Completion - 20pts":                        "trans",
    "Speech Clarify - 5pts":                                 "speech",
    "Information Precision - 10pts":                         "info_prec",
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

BLC_DIR = Path(os.getenv("BLC_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Britelift Chat"))
BLC_SCRIPT = BLC_DIR / "britelift_pull.py"
BLC_OUTPUT_FILE = BLC_DIR / "BLC_RAW.xlsx"

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

BLC_COLUMN_MAP = {
    "QA_ID":                                                                                                               "qa_id",
    "EMPLOYEE_ID":                                                                                                         "emp_id",
    "Timestamp":                                                                                                           "ts",
    "Chat Date":                                                                                                           "date",
    "Emp Name":                                                                                                            "agent",
    "Score":                                                                                                               "score",
    "Evaluation Type":                                                                                                     "type",
    "QA":                                                                                                                  "coach",
    "Immediate Supervisor":                                                                                                "supervisor",
    "Feedback Summary ":                                                                                                   "feedback",
    "Did the agent follow greeting and any other chat scripts?  (10%)":                                                     "os_in",
    "Did the agent maintain professionalism?  (10%)":                                                                       "approp",
    "Did the agent follow in providing customer basic information such as ETA and fixed rate?  (15%)":                      "answered",
    "Did the agent ask for the customer's name, phone number, address and email address?  (5%)":                           "verif",
    "Did the agent respond to the customer promptly?  (5%)":                                                               "resp_eff",
    "Proper usage of Grammar & Punctuation Marks.  (15%)":                                                                 "speech",
    "Rudeness  (20%)":                                                                                                     "rude",
    "Successfully processed customer's request / was able to address the customer's concern and handle the live chat conversation properly.  (20%)": "trans",
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

VIP_DIR = Path(os.getenv("VIP_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\VIP"))
VIP_SCRIPT = VIP_DIR / "vip_pull.py"
VIP_OUTPUT_FILE = VIP_DIR / "VIP_RAW.xlsx"

CH_DIR = Path(os.getenv("CH_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\C&H"))
CH_SCRIPT = CH_DIR / "ch_pull.py"
CH_OUTPUT_FILE = CH_DIR / "CH_RAW.xlsx"

RC_DIR = Path(os.getenv("RC_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Reno Cab"))
RC_SCRIPT = RC_DIR / "rc_pull.py"
RC_OUTPUT_FILE = RC_DIR / "RC_RAW.xlsx"

TI_DIR = Path(os.getenv("TI_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Trans Iowa"))
TI_SCRIPT = TI_DIR / "ti_pull.py"
TI_OUTPUT_FILE = TI_DIR / "TI_RAW.xlsx"

DC_DIR = Path(os.getenv("DC_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Data Carz"))
DC_SCRIPT = DC_DIR / "dc_pull.py"
DC_OUTPUT_FILE = DC_DIR / "DC_RAW.xlsx"

AC_DIR = Path(os.getenv("AC_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Associated Cab"))
AC_SCRIPT = AC_DIR / "ac_pull.py"
AC_OUTPUT_FILE = AC_DIR / "AC_RAW.xlsx"

OL_DIR = Path(os.getenv("OL_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Ollies"))
OL_SCRIPT = OL_DIR / "ol_pull.py"
OL_OUTPUT_FILE = OL_DIR / "OL_RAW.xlsx"

CT_DIR = Path(os.getenv("CT_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Circle Taxi"))
CT_SCRIPT = CT_DIR / "ct_pull.py"
CT_OUTPUT_FILE = CT_DIR / "CT_RAW.xlsx"

YCOV_DIR = Path(os.getenv("YCOV_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCOV"))
YCOV_SCRIPT = YCOV_DIR / "ycov_pull.py"
YCOV_OUTPUT_FILE = YCOV_DIR / "YCOV_RAW.xlsx"

KEL_DIR = Path(os.getenv("KEL_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Kelowna"))
KEL_SCRIPT = KEL_DIR / "kel_pull.py"
KEL_OUTPUT_FILE = KEL_DIR / "KEL_RAW.xlsx"

VT_DIR = Path(os.getenv("VT_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Vermont"))
VT_SCRIPT = VT_DIR / "vt_pull.py"
VT_OUTPUT_FILE = VT_DIR / "VT_RAW.xlsx"

YCDC_DIR = Path(os.getenv("YCDC_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\YCDC"))
YCDC_SCRIPT = YCDC_DIR / "ycdc_pull.py"
YCDC_OUTPUT_FILE = YCDC_DIR / "YCDC_RAW.xlsx"

BL_DIR = Path(os.getenv("BL_DIR", r"C:\Users\Mike Woo Cerna\Documents\PB\Quality\Blueline"))
BL_SCRIPT = BL_DIR / "bl_pull.py"
BL_OUTPUT_FILE = BL_DIR / "BL_RAW.xlsx"

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

# Maps standard detail-table criterion keys → VIP base column name (without _AI/_Max suffix)
VIP_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

# Maps standard detail-table criterion keys → C&H base column name (without _AI/_Max suffix)
CH_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

# C&H-unique criteria not in the standard QA_CRIT_META
CH_EXTRA_CRIT_MAP = {
    "ch_os_out2":         "Opening_Spiel_Outbound_Call",
    "ch_profess":         "Professionalism",
    "ch_verif_other":     "Verification_Other_Measures",
    "ch_comm_quality":    "Communication_Quality",
    "ch_cust_verif_meas": "Customer_Verification_Measures",
    "ch_res_etiq":        "Resolution_Etiquette",
    "ch_subjective":      "Subjective",
}

# VIP-unique criteria not in the standard QA_CRIT_META
VIP_EXTRA_CRIT_MAP = {
    "vip_os_out2":        "Opening_Spiel_Outbound_Call",
    "vip_profess":        "Professionalism",
    "vip_verif_other":    "Verification_Other_Measures",
    "vip_comm_quality":   "Communication_Quality",
    "vip_cust_verif_oth": "Customer_Verification_Other_Measures",
    "vip_res_etiq":       "Resolution_Etiquette",
    "vip_sop":            "VIP_Standard_Operating_Procedure",
}

RC_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

RC_EXTRA_CRIT_MAP = {
    "rc_profess":         "Professionalism",
    "rc_verif_other":     "Verification_Other_Measures",
    "rc_res_etiq":        "Resolution_Etiquette",
    "rc_comm_quality":    "Communication_Quality",
    "rc_cust_verif_meas": "Customer_Verification_Measures",
}

TI_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_Response",
    "fillers":  "Filler_Slang_Words",
    "ack":      "Acknowledgment_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thanks_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Successfully_Processed",
    "speech":   "Speech_Clarify",
}

TI_EXTRA_CRIT_MAP = {
    "ti_os_out2":      "Opening_Spiel_Outbound_Call",
    "ti_profess":      "Professionalism",
    "ti_verif_other":  "Customer_Verification_Other_Measures",
    "ti_res_etiq":     "Resolution_Etiquette",
    "ti_comm_quality": "Communication_Quality",
}

DC_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_Response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thanks_the_Customer_for_Waiting",
    "resp_eff": "Prompt_Response_to_Customer",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Listening_Attentively",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Successfully_Processed",
    "speech":   "Speech_Clarify",
}

DC_EXTRA_CRIT_MAP = {
    "dc_os_out2":      "Opening_Spiel_Outbound_Call",
    "dc_profess":      "Professionalism",
    "dc_res_etiq":     "Resolution_Etiquette",
    "dc_comm_quality": "Communication_Quality",
}

AC_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

AC_EXTRA_CRIT_MAP = {
    "ac_os_out2":         "Opening_Spiel_Outbound_Call",
    "ac_profess":         "Professionalism",
    "ac_verif_other":     "Verification_Other_Measures",
    "ac_comm_quality":    "Communication_Quality",
    "ac_cust_verif_meas": "Customer_Verification_Measures",
    "ac_res_etiq":        "Resolution_Etiquette",
}

OL_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_Response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgment_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thanks_the_Customer_for_Waiting",
    "resp_eff": "Prompt_Response_to_Customer",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Listening_Attentively",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Successfully_Processed",
    "speech":   "Communication_Quality",
}

OL_EXTRA_CRIT_MAP = {
    "ol_prompt":          "Opening_Promptness",
    "ol_profess":         "Professionalism",
    "ol_dead_air":        "Dead_Air_Length_of_Pause",
    "ol_no_vern":         "No_Vernacular",
    "ol_avoid_int":       "Avoiding_Interruptions_Verbal_Collision",
    "ol_cust_verif_meas": "Customer_Verification_Measures",
    "ol_ride_cancel":     "Ride_Cancellation_SOP",
    "ol_timeliness":      "Timeliness_in_Handling",
    "ol_res_etiq":        "Resolution_Etiquette",
    "ol_comm_quality":    "Communication_Quality",
    "ol_stutter":         "Stuttering",
    "ol_grammar":         "Grammar",
    "ol_pronunc":         "Pronunciation",
    "ol_tone":            "Tone_of_Voice",
    "ol_prolong":         "Prolonging_the_Call",
}

CT_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

CT_EXTRA_CRIT_MAP = {
    "ct_os_out2":         "Opening_Spiel_Outbound_Call",
    "ct_profess":         "Professionalism",
    "ct_verif_other":     "Verification_Other_Measures",
    "ct_comm_quality":    "Communication_Quality",
    "ct_cust_verif_meas": "Customer_Verification_Measures",
    "ct_res_etiq":        "Resolution_Etiquette",
    "ct_subjective":      "Subjective",
}

YCOV_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

YCOV_EXTRA_CRIT_MAP = {
    "ycov_os_out2":         "Opening_Spiel_Outbound_Call",
    "ycov_profess":         "Professionalism",
    "ycov_verif_other":     "Verification_Other_Measures",
    "ycov_comm_quality":    "Communication_Quality",
    "ycov_cust_verif_meas": "Customer_Verification_Measures",
    "ycov_res_etiq":        "Resolution_Etiquette",
    "ycov_subjective":      "Subjective",
}

KEL_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

KEL_EXTRA_CRIT_MAP = {
    "kel_os_out2":         "Opening_Spiel_Outbound_Call",
    "kel_profess":         "Professionalism",
    "kel_verif_other":     "Verification_Other_Measures",
    "kel_comm_quality":    "Communication_Quality",
    "kel_cust_verif_meas": "Customer_Verification_Measures",
    "kel_res_etiq":        "Resolution_Etiquette",
}

VT_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

VT_EXTRA_CRIT_MAP = {
    "vt_os_out2":      "Opening_Spiel_Outbound_Call",
    "vt_profess":      "Professionalism",
    "vt_verif_other":  "Verification_Other_Measures",
    "vt_comm_quality": "Communication_Quality",
    "vt_res_etiq":     "Resolution_Etiquette",
}

YCDC_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_Listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

YCDC_EXTRA_CRIT_MAP = {
    "ycdc_os_out2":      "Opening_Spiel_Outbound_Call",
    "ycdc_profess":      "Professionalism",
    "ycdc_verif_other":  "Verification_Other_Measures",
    "ycdc_comm_quality": "Communication_Quality",
    "ycdc_res_etiq":     "Resolution_Etiquette",
}

BL_CRIT_MAP = {
    "os_in":    "Opening_Spiel_Inbound_Call",
    "os_out":   "Greetings_Call_Script",
    "closing":  "Closing_Spiel",
    "approp":   "Appropriate_Response",
    "no_resp":  "No_response",
    "fillers":  "Fillers_Slang_Words",
    "ack":      "Acknowledgement_Ownership",
    "hold":     "Proper_Handling_of_Pauses_or_Hold_Requests",
    "ack_hold": "Acknowledges_and_Thank_the_Customer_for_Waiting",
    "resp_eff": "Response_Efficiency",
    "empathy":  "Empathy_Sympathy",
    "adjust":   "Adjusts_to_Customer_s_Level",
    "mute":     "Mute_Button_Usage",
    "active":   "Active_listening",
    "gen_q":    "General_Questions",
    "answered": "Answered_Customer_s_Questions",
    "probing":  "Probing_Questions",
    "verif":    "Customer_Verification_Other_Measures",
    "clarif":   "Clarification_When_Information_is_Missed",
    "lost_sop": "Lost_Item_SOP",
    "rude":     "Rudeness",
    "trans":    "Transaction_Completion",
    "speech":   "Speech_Clarify",
}

BL_EXTRA_CRIT_MAP = {
    "bl_os_out2":      "Opening_Spiel_Outbound_Call",
    "bl_profess":      "Professionalism",
    "bl_verif_other":  "Verification_Other_Measures",
    "bl_comm_quality": "Communication_Quality",
    "bl_res_etiq":     "Resolution_Etiquette",
}

COACHING_VALIDATION_ERRORS_FILE = COACHING_OUTPUT_FILE.parent / "coaching_validation_errors.csv"
COACHING_VALIDATION_NOTIFY_FILE = COACHING_OUTPUT_FILE.parent / "coaching_validation_notified.json"
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
    "1216430783913500",
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


def refresh_dmg_output():
    if not DMG_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(DMG_SCRIPT)],
            cwd=str(DMG_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping DMG pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping DMG pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_dmg_workbook():
    if not DMG_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(DMG_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping DMG workbook load: {exc}")
        return pd.DataFrame()


def refresh_r4h_output():
    if not R4H_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(R4H_SCRIPT)],
            cwd=str(R4H_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping R4H pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping R4H pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_r4h_workbook():
    if not R4H_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(R4H_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping R4H workbook load: {exc}")
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
    df = df.loc[:, ~df.columns.duplicated(keep="last")]

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
        def parse_score(v):
            s = clean_val(v)
            if not s:
                return 0
            if "/" in s:
                left, right = s.split("/", 1)
                try:
                    numerator = float(left.strip())
                    denominator = float(right.strip())
                    if denominator and denominator != 100:
                        return numerator / denominator * 100
                    return numerator
                except ValueError:
                    pass
            match = re.search(r"-?\d+(?:\.\d+)?", s)
            return float(match.group(0)) if match else 0

        df["score"] = df["score"].apply(parse_score).round().astype(int)
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
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start", "date",
        "agent", "score", "type", "coach", "supervisor", "invest", "feedback",
        "os_in", "os_out", "closing", "approp", "no_resp", "fillers",
        "ack", "hold", "ack_hold", "resp_eff", "empathy", "adjust",
        "mute", "active", "answered", "probing", "verif", "clarif",
        "lost_sop", "rude", "trans", "speech", "info_prec",
    ]
    return df[[c for c in keep if c in df.columns]]


def transform_m7_data(source):
    return _transform_qa_source(source, M7_COLUMN_MAP)


def transform_dmg_data(source):
    return _transform_qa_source(source, DMG_COLUMN_MAP)


def transform_r4h_data(source):
    return _transform_qa_source(source, R4H_COLUMN_MAP)


def transform_parentis_data(source):
    return _transform_qa_source(source, PARENTIS_COLUMN_MAP)


def load_m7_data():
    refresh_m7_output()
    source = read_m7_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_m7_data(source)
    print(f"M7 QA rows: {len(result)}")
    return result


def load_dmg_data():
    refresh_dmg_output()
    source = read_dmg_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_dmg_data(source)
    print(f"DMG QA rows: {len(result)}")
    return result


def load_r4h_data():
    refresh_r4h_output()
    source = read_r4h_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_r4h_data(source)
    print(f"R4H QA rows: {len(result)}")
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
        return pd.DataFrame(columns=["qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "date",
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
        return pd.DataFrame(columns=["qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_britelift_data(source)
    print(f"Britelift QA rows: {len(result)}")
    return result


def transform_blc_data(source):
    return _transform_qa_source(source, BLC_COLUMN_MAP)


def refresh_blc_output():
    if not BLC_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(BLC_SCRIPT)],
            cwd=str(BLC_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Britelift Chat pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Britelift Chat pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_blc_workbook():
    if not BLC_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(BLC_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Britelift Chat workbook load: {exc}")
        return pd.DataFrame()


def load_blc_data():
    refresh_blc_output()
    source = read_blc_workbook()
    if source.empty:
        return pd.DataFrame(columns=["qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "date",
                                     "agent", "score", "type", "coach", "supervisor",
                                     "invest", "feedback"])
    result = transform_blc_data(source)
    print(f"Britelift Chat QA rows: {len(result)}")
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
        return pd.DataFrame(columns=["qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "date",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
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
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
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
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_skyline_data(source)
    print(f"Skyline QA rows: {len(result)}")
    return result


def transform_vip_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    VIP_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, vip_base in VIP_CRIT_MAP.items():
        ai_col  = f"{vip_base}_AI"
        max_col = f"{vip_base}_Max"
        ai_s  = pd.to_numeric(df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float), errors="coerce").fillna(0)
        max_s = pd.to_numeric(df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float), errors="coerce").fillna(0)
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
        elif std_key in VIP_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for vip_key, vip_base in VIP_EXTRA_CRIT_MAP.items():
        ai_col  = f"{vip_base}_AI"
        max_col = f"{vip_base}_Max"
        ai_s  = pd.to_numeric(df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float), errors="coerce").fillna(0)
        max_s = pd.to_numeric(df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float), errors="coerce").fillna(0)
        if vip_key == "vip_sop":
            df[vip_key] = [
                "Not Applicable" if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[vip_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    crit_keys = list(VIP_CRIT_MAP.keys()) + list(VIP_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_vip_output():
    if not VIP_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(VIP_SCRIPT)],
            cwd=str(VIP_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping VIP pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping VIP pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_vip_workbook():
    if not VIP_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(VIP_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping VIP workbook load: {exc}")
        return pd.DataFrame()


def load_vip_data():
    refresh_vip_output()
    source = read_vip_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_vip_data(source)
    print(f"VIP QA rows: {len(result)}")
    return result


def transform_ch_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    CH_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, ch_base in CH_CRIT_MAP.items():
        ai_col  = f"{ch_base}_AI"
        max_col = f"{ch_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in CH_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for ch_key, ch_base in CH_EXTRA_CRIT_MAP.items():
        ai_col  = f"{ch_base}_AI"
        max_col = f"{ch_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[ch_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(CH_CRIT_MAP.keys()) + list(CH_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_ch_output():
    if not CH_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(CH_SCRIPT)],
            cwd=str(CH_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping C&H pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping C&H pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ch_workbook():
    if not CH_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(CH_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping C&H workbook load: {exc}")
        return pd.DataFrame()


def load_ch_data():
    refresh_ch_output()
    source = read_ch_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_ch_data(source)
    print(f"C&H QA rows: {len(result)}")
    return result


def transform_rc_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    RC_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, rc_base in RC_CRIT_MAP.items():
        ai_col  = f"{rc_base}_AI"
        max_col = f"{rc_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in RC_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for rc_key, rc_base in RC_EXTRA_CRIT_MAP.items():
        ai_col  = f"{rc_base}_AI"
        max_col = f"{rc_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[rc_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(RC_CRIT_MAP.keys()) + list(RC_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_rc_output():
    if not RC_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(RC_SCRIPT)],
            cwd=str(RC_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Reno Cab pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Reno Cab pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_rc_workbook():
    if not RC_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(RC_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Reno Cab workbook load: {exc}")
        return pd.DataFrame()


def load_rc_data():
    refresh_rc_output()
    source = read_rc_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_rc_data(source)
    print(f"Reno Cab QA rows: {len(result)}")
    return result


def transform_ti_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    TI_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, ti_base in TI_CRIT_MAP.items():
        ai_col  = f"{ti_base}_AI"
        max_col = f"{ti_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in TI_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for ti_key, ti_base in TI_EXTRA_CRIT_MAP.items():
        ai_col  = f"{ti_base}_AI"
        max_col = f"{ti_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[ti_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(TI_CRIT_MAP.keys()) + list(TI_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_ti_output():
    if not TI_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(TI_SCRIPT)],
            cwd=str(TI_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Trans Iowa pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Trans Iowa pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ti_workbook():
    if not TI_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(TI_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Trans Iowa workbook load: {exc}")
        return pd.DataFrame()


def load_ti_data():
    refresh_ti_output()
    source = read_ti_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_ti_data(source)
    print(f"Trans Iowa QA rows: {len(result)}")
    return result


def transform_dc_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    DC_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    _zero_s = pd.Series(0, index=df.index, dtype=float)
    for std_key, dc_base in DC_CRIT_MAP.items():
        ai_col  = f"{dc_base}_AI"
        max_col = f"{dc_base}_Max"
        ai_s  = pd.to_numeric(df[ai_col]  if ai_col  in df.columns else _zero_s, errors="coerce").fillna(0)
        max_s = pd.to_numeric(df[max_col] if max_col in df.columns else _zero_s, errors="coerce").fillna(0)
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
        elif std_key in DC_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for dc_key, dc_base in DC_EXTRA_CRIT_MAP.items():
        ai_col  = f"{dc_base}_AI"
        max_col = f"{dc_base}_Max"
        ai_s  = pd.to_numeric(df[ai_col]  if ai_col  in df.columns else _zero_s, errors="coerce").fillna(0)
        max_s = pd.to_numeric(df[max_col] if max_col in df.columns else _zero_s, errors="coerce").fillna(0)
        df[dc_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(DC_CRIT_MAP.keys()) + list(DC_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_dc_output():
    if not DC_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(DC_SCRIPT)],
            cwd=str(DC_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Data Carz pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Data Carz pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_dc_workbook():
    if not DC_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(DC_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Data Carz workbook load: {exc}")
        return pd.DataFrame()


def load_dc_data():
    refresh_dc_output()
    source = read_dc_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_dc_data(source)
    print(f"Data Carz QA rows: {len(result)}")
    return result


def transform_ac_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    AC_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, ac_base in AC_CRIT_MAP.items():
        ai_col  = f"{ac_base}_AI"
        max_col = f"{ac_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in AC_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for ac_key, ac_base in AC_EXTRA_CRIT_MAP.items():
        ai_col  = f"{ac_base}_AI"
        max_col = f"{ac_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[ac_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(AC_CRIT_MAP.keys()) + list(AC_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_ac_output():
    if not AC_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(AC_SCRIPT)],
            cwd=str(AC_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Associated Cab pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Associated Cab pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ac_workbook():
    if not AC_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(AC_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Associated Cab workbook load: {exc}")
        return pd.DataFrame()


def load_ac_data():
    refresh_ac_output()
    source = read_ac_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_ac_data(source)
    print(f"Associated Cab QA rows: {len(result)}")
    return result


def transform_ol_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    OL_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, ol_base in OL_CRIT_MAP.items():
        ai_col  = f"{ol_base}_AI"
        max_col = f"{ol_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in OL_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    OL_NA_KEYS = {"ol_ride_cancel", "ol_timeliness"}
    for ol_key, ol_base in OL_EXTRA_CRIT_MAP.items():
        ai_col  = f"{ol_base}_AI"
        max_col = f"{ol_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        if ol_key in OL_NA_KEYS:
            df[ol_key] = [
                "Not Applicable" if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[ol_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    crit_keys = list(OL_CRIT_MAP.keys()) + list(OL_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_ol_output():
    if not OL_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(OL_SCRIPT)],
            cwd=str(OL_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Ollies pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Ollies pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ol_workbook():
    if not OL_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(OL_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Ollies workbook load: {exc}")
        return pd.DataFrame()


def load_ol_data():
    refresh_ol_output()
    source = read_ol_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_ol_data(source)
    print(f"Ollies QA rows: {len(result)}")
    return result


def transform_ct_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    CT_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, ct_base in CT_CRIT_MAP.items():
        ai_col  = f"{ct_base}_AI"
        max_col = f"{ct_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in CT_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for ct_key, ct_base in CT_EXTRA_CRIT_MAP.items():
        ai_col  = f"{ct_base}_AI"
        max_col = f"{ct_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[ct_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(CT_CRIT_MAP.keys()) + list(CT_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_ct_output():
    if not CT_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(CT_SCRIPT)],
            cwd=str(CT_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Circle Taxi pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Circle Taxi pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ct_workbook():
    if not CT_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(CT_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Circle Taxi workbook load: {exc}")
        return pd.DataFrame()


def load_ct_data():
    refresh_ct_output()
    source = read_ct_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_ct_data(source)
    print(f"Circle Taxi QA rows: {len(result)}")
    return result


def transform_ycov_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    YCOV_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, ycov_base in YCOV_CRIT_MAP.items():
        ai_col  = f"{ycov_base}_AI"
        max_col = f"{ycov_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in YCOV_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for ycov_key, ycov_base in YCOV_EXTRA_CRIT_MAP.items():
        ai_col  = f"{ycov_base}_AI"
        max_col = f"{ycov_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[ycov_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(YCOV_CRIT_MAP.keys()) + list(YCOV_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_ycov_output():
    if not YCOV_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(YCOV_SCRIPT)],
            cwd=str(YCOV_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping YCOV pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping YCOV pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ycov_workbook():
    if not YCOV_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(YCOV_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping YCOV workbook load: {exc}")
        return pd.DataFrame()


def load_ycov_data():
    refresh_ycov_output()
    source = read_ycov_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_ycov_data(source)
    print(f"YCOV QA rows: {len(result)}")
    return result


def transform_kel_data(source):
    df = source.copy()
    rename = {
        "evaluation_date":      "ts",
        "Emp Name":             "agent",
        "QA":                   "coach",
        "Immediate Supervisor": "supervisor",
        "evaluation_status":    "status",
        "overall_score_ai":     "score_ai",
        "overall_score_human":  "score_human",
        "QA_ID":                "qa_id",
        "evaluation_id":        "evaluation_id",
        "EMPLOYEE_ID":          "emp_id",
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

    score_ai_s = pd.to_numeric((df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
    df["score_ai"] = score_ai_s.round(1)

    score_h_s = pd.to_numeric((df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce")
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

    KEL_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, kel_base in KEL_CRIT_MAP.items():
        ai_col  = f"{kel_base}_AI"
        max_col = f"{kel_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in KEL_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for kel_key, kel_base in KEL_EXTRA_CRIT_MAP.items():
        ai_col  = f"{kel_base}_AI"
        max_col = f"{kel_base}_Max"
        ai_s  = pd.to_numeric((df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[kel_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(KEL_CRIT_MAP.keys()) + list(KEL_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_kel_output():
    if not KEL_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(KEL_SCRIPT)],
            cwd=str(KEL_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Kelowna pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Kelowna pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_kel_workbook():
    if not KEL_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(KEL_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Kelowna workbook load: {exc}")
        return pd.DataFrame()


def load_kel_data():
    refresh_kel_output()
    source = read_kel_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_kel_data(source)
    print(f"Kelowna QA rows: {len(result)}")
    return result


def transform_vt_data(df):
    df = clean_columns(df.copy())

    rename_map = {
        "QA_ID": "qa_id",
        "EMPLOYEE_ID": "emp_id",
        "Emp Name": "agent",
        "agent_name": "agent",
        "QA": "coach",
        "quality_evaluator": "coach",
        "Immediate Supervisor": "supervisor",
        "recorded_date": "ts",
        "evaluation_date": "date",
        "evaluation_status": "status",
        "overall_score_percentage": "score",
        "overall_score_ai": "score_ai",
        "overall_score_human": "score_human",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    df = df.loc[:, ~df.columns.duplicated()]
    for col in ("qa_id", "emp_id", "agent", "coach", "supervisor", "ts", "status", "score"):
        if col not in df.columns:
            df[col] = ""

    def fmt_ts(v):
        try:
            dt = pd.to_datetime(str(v), errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        return ""

    df["ts"] = df["ts"].apply(fmt_ts)
    df["status"] = df["status"].apply(
        lambda v: str(v).strip().lower() if str(v).strip() not in ("", "nan") else "rated"
    )

    for col in ("score", "score_ai", "score_human"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "score_ai" not in df.columns:
        df["score_ai"] = None
    if "score_human" not in df.columns:
        df["score_human"] = None

    def effective_score(row):
        human = row.get("score_human")
        ai = row.get("score_ai")
        score = row.get("score")
        if pd.notna(human):
            return human
        if pd.notna(ai):
            return ai
        return score

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

    VT_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, vt_base in VT_CRIT_MAP.items():
        ai_col = f"{vt_base}_AI"
        max_col = f"{vt_base}_Max"
        ai_s = pd.to_numeric((df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in VT_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for vt_key, vt_base in VT_EXTRA_CRIT_MAP.items():
        ai_col = f"{vt_base}_AI"
        max_col = f"{vt_base}_Max"
        ai_s = pd.to_numeric((df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[vt_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(VT_CRIT_MAP.keys()) + list(VT_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_vt_output():
    if not VT_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(VT_SCRIPT)],
            cwd=str(VT_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Vermont pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Vermont pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_vt_workbook():
    if not VT_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(VT_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Vermont workbook load: {exc}")
        return pd.DataFrame()


def load_vt_data():
    refresh_vt_output()
    source = read_vt_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_vt_data(source)
    print(f"Vermont QA rows: {len(result)}")
    return result


def transform_ycdc_data(df):
    df = clean_columns(df.copy())

    rename_map = {
        "QA_ID": "qa_id",
        "EMPLOYEE_ID": "emp_id",
        "Emp Name": "agent",
        "agent_name": "agent",
        "QA": "coach",
        "quality_evaluator": "coach",
        "Immediate Supervisor": "supervisor",
        "recorded_date": "ts",
        "evaluation_date": "date",
        "evaluation_status": "status",
        "overall_score_percentage": "score",
        "overall_score_ai": "score_ai",
        "overall_score_human": "score_human",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    df = df.loc[:, ~df.columns.duplicated()]
    for col in ("qa_id", "emp_id", "agent", "coach", "supervisor", "ts", "status", "score"):
        if col not in df.columns:
            df[col] = ""

    def fmt_ts(v):
        try:
            dt = pd.to_datetime(str(v), errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        return ""

    df["ts"] = df["ts"].apply(fmt_ts)
    df["status"] = df["status"].apply(
        lambda v: str(v).strip().lower() if str(v).strip() not in ("", "nan") else "rated"
    )

    for col in ("score", "score_ai", "score_human"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "score_ai" not in df.columns:
        df["score_ai"] = None
    if "score_human" not in df.columns:
        df["score_human"] = None

    def effective_score(row):
        human = row.get("score_human")
        ai = row.get("score_ai")
        score = row.get("score")
        if pd.notna(human):
            return human
        if pd.notna(ai):
            return ai
        return score

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

    YCDC_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, ycdc_base in YCDC_CRIT_MAP.items():
        ai_col = f"{ycdc_base}_AI"
        max_col = f"{ycdc_base}_Max"
        ai_s = pd.to_numeric((df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in YCDC_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for ycdc_key, ycdc_base in YCDC_EXTRA_CRIT_MAP.items():
        ai_col = f"{ycdc_base}_AI"
        max_col = f"{ycdc_base}_Max"
        ai_s = pd.to_numeric((df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[ycdc_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(YCDC_CRIT_MAP.keys()) + list(YCDC_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_ycdc_output():
    if not YCDC_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(YCDC_SCRIPT)],
            cwd=str(YCDC_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping YCDC pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping YCDC pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_ycdc_workbook():
    if not YCDC_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(YCDC_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping YCDC workbook load: {exc}")
        return pd.DataFrame()


def load_ycdc_data():
    refresh_ycdc_output()
    source = read_ycdc_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_ycdc_data(source)
    print(f"YCDC QA rows: {len(result)}")
    return result


def transform_bl_data(df):
    df = clean_columns(df.copy())

    rename_map = {
        "QA_ID": "qa_id",
        "EMPLOYEE_ID": "emp_id",
        "Emp Name": "agent",
        "agent_name": "agent",
        "QA": "coach",
        "quality_evaluator": "coach",
        "Immediate Supervisor": "supervisor",
        "recorded_date": "ts",
        "evaluation_date": "date",
        "evaluation_status": "status",
        "overall_score_percentage": "score",
        "overall_score_ai": "score_ai",
        "overall_score_human": "score_human",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    df = df.loc[:, ~df.columns.duplicated()]
    for col in ("qa_id", "emp_id", "agent", "coach", "supervisor", "ts", "status", "score"):
        if col not in df.columns:
            df[col] = ""

    def fmt_ts(v):
        try:
            dt = pd.to_datetime(str(v), errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        return ""

    df["ts"] = df["ts"].apply(fmt_ts)
    df["status"] = df["status"].apply(
        lambda v: str(v).strip().lower() if str(v).strip() not in ("", "nan") else "rated"
    )

    for col in ("score", "score_ai", "score_human"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "score_ai" not in df.columns:
        df["score_ai"] = None
    if "score_human" not in df.columns:
        df["score_human"] = None

    def effective_score(row):
        human = row.get("score_human")
        ai = row.get("score_ai")
        score = row.get("score")
        if pd.notna(human):
            return human
        if pd.notna(ai):
            return ai
        return score

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

    BL_MAIN_ATTRS = {"os_out", "adjust", "gen_q", "verif", "resp_eff", "speech"}
    for std_key, bl_base in BL_CRIT_MAP.items():
        ai_col = f"{bl_base}_AI"
        max_col = f"{bl_base}_Max"
        ai_s = pd.to_numeric((df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
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
        elif std_key in BL_MAIN_ATTRS:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai >= mx else "No")
                for ai, mx in zip(ai_s, max_s)
            ]
        else:
            df[std_key] = [
                None if mx == 0 else ("Yes" if ai > 0 else "No")
                for ai, mx in zip(ai_s, max_s)
            ]

    for bl_key, bl_base in BL_EXTRA_CRIT_MAP.items():
        ai_col = f"{bl_base}_AI"
        max_col = f"{bl_base}_Max"
        ai_s = pd.to_numeric((df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        max_s = pd.to_numeric((df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float)), errors="coerce").fillna(0)
        df[bl_key] = [
            None if mx == 0 else ("Yes" if ai > 0 else "No")
            for ai, mx in zip(ai_s, max_s)
        ]

    crit_keys = list(BL_CRIT_MAP.keys()) + list(BL_EXTRA_CRIT_MAP.keys())
    keep = [
        "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
        "agent", "score", "score_ai", "score_human", "status",
        "coach", "supervisor",
    ] + crit_keys
    for col in ("agent", "supervisor", "coach"):
        if col in df.columns:
            df[col] = _apply_name_aliases(df[col])
    return df[[c for c in keep if c in df.columns]]


def refresh_bl_output():
    if not BL_SCRIPT.exists():
        return False
    try:
        result = subprocess.run(
            [sys.executable, str(BL_SCRIPT)],
            cwd=str(BL_DIR),
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"Skipping Blueline pull: {exc}")
        return False
    if result.returncode != 0:
        msg = result.stderr.strip() or result.stdout.strip() or "No details."
        print(f"Skipping Blueline pull: {msg}")
        return False
    if result.stdout.strip():
        print(result.stdout.strip())
    return True


def read_bl_workbook():
    if not BL_OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        return clean_columns(pd.read_excel(BL_OUTPUT_FILE))
    except Exception as exc:
        print(f"Skipping Blueline workbook load: {exc}")
        return pd.DataFrame()


def load_bl_data():
    refresh_bl_output()
    source = read_bl_workbook()
    if source.empty:
        return pd.DataFrame(columns=[
            "qa_id", "evaluation_id", "eval_key", "emp_id", "ts", "week_start",
            "agent", "score", "score_ai", "score_human", "status",
            "coach", "supervisor",
        ])
    result = transform_bl_data(source)
    print(f"Blueline QA rows: {len(result)}")
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


def coaching_validation_signature(errors):
    normalized = [
        {
            "task_gid": cell_text(item.get("Task GID")),
            "error": cell_text(item.get("Error")),
            "employee_email": normalize_email(item.get("Employee Email")),
            "supervisor_email": normalize_email(item.get("Supervisor Email")),
        }
        for item in errors
    ]
    payload = json.dumps(sorted(normalized, key=lambda item: item["task_gid"]), sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def load_coaching_validation_notification_state():
    try:
        if COACHING_VALIDATION_NOTIFY_FILE.exists():
            return json.loads(COACHING_VALIDATION_NOTIFY_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        pass
    return {}


def save_coaching_validation_notification_state(state):
    try:
        COACHING_VALIDATION_NOTIFY_FILE.parent.mkdir(parents=True, exist_ok=True)
        COACHING_VALIDATION_NOTIFY_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")
    except OSError as exc:
        print(f"Unable to update Coaching validation notification state: {exc}")


def notify_coaching_validation_alerts(errors):
    state = load_coaching_validation_notification_state()

    if not errors:
        if state:
            save_coaching_validation_notification_state({})
        return

    signature = coaching_validation_signature(errors)
    if state.get("signature") == signature:
        print("[notify] Coaching validation alert already sent for current error set.")
        return

    rows_html = []
    rows_text = []
    for item in errors:
        task_gid = html.escape(cell_text(item.get("Task GID")) or "(blank)")
        error = html.escape(cell_text(item.get("Error")))
        employee_email = html.escape(cell_text(item.get("Employee Email")) or "(blank)")
        supervisor_email = html.escape(cell_text(item.get("Supervisor Email")) or "(blank)")
        task_name = html.escape(cell_text(item.get("Task Name")) or "(blank)")
        rows_html.append(
            "<tr>"
            f"<td style='padding:8px;border-bottom:1px solid #e5e7eb;font-family:monospace;font-size:12px;'>{task_gid}</td>"
            f"<td style='padding:8px;border-bottom:1px solid #e5e7eb;'>{error}</td>"
            f"<td style='padding:8px;border-bottom:1px solid #e5e7eb;'>{employee_email}</td>"
            f"<td style='padding:8px;border-bottom:1px solid #e5e7eb;'>{supervisor_email}</td>"
            "</tr>"
        )
        rows_text.append(
            f"Coaching ID: {cell_text(item.get('Task GID')) or '(blank)'}\n"
            f"Task Name: {cell_text(item.get('Task Name')) or '(blank)'}\n"
            f"Error: {cell_text(item.get('Error'))}\n"
            f"Employee Email: {cell_text(item.get('Employee Email')) or '(blank)'}\n"
            f"Supervisor Email: {cell_text(item.get('Supervisor Email')) or '(blank)'}"
        )

    session_word = "session" if len(errors) == 1 else "sessions"
    subject = "PACE Validation Errors Detected in Coaching Entries"
    html_body = f"""
<div style="font-family:Arial,sans-serif;max-width:760px;">
  <div style="background:#b45309;color:#fff;padding:14px 20px;border-radius:6px 6px 0 0;">
    <b style="font-size:16px;">PACE Validation ALERT!!!</b>
  </div>
  <div style="border:1px solid #fcd34d;border-top:none;padding:18px 20px;background:#fffbeb;border-radius:0 0 6px 6px;">
    <p style="margin:0 0 12px;font-size:14px;color:#374151;">Hi Mike,</p>
    <p style="margin:0 0 12px;font-size:14px;color:#374151;">
      PACE (Pac-Biz Automated Coaching Engine) has detected {len(errors)} coaching {session_word} that were excluded from the Coaching Report due to invalid or incorrect employee email addresses entered in Asana.
    </p>
    <table style="width:100%;font-size:13px;border-collapse:collapse;background:#fff;border:1px solid #e5e7eb;">
      <thead>
        <tr style="background:#004C97;color:#fff;text-align:left;">
          <th style="padding:8px;">Coaching ID</th>
          <th style="padding:8px;">Issue</th>
          <th style="padding:8px;">Employee Email</th>
          <th style="padding:8px;">Supervisor Email</th>
        </tr>
      </thead>
      <tbody>
        {''.join(rows_html)}
      </tbody>
    </table>
    <p style="margin-top:14px;font-size:13px;color:#374151;">
      Correct the email address in Asana or add the employee/supervisor email to the Masterlist, then rerun the dashboard pipeline.
    </p>
    <p style="margin-top:8px;font-size:12px;color:#6b7280;">
      Validation details were saved locally to: {html.escape(str(COACHING_VALIDATION_ERRORS_FILE))}
    </p>
    <p style="margin-top:8px;font-size:12px;color:#6b7280;">
      PACE AI Email notification Activated as Watchers.. Report to Mike
    </p>
    <p style="margin-top:12px;">
      <a href="https://mikewoocerna.github.io/Pac-Biz/pipeline_monitor.html"
         style="background:#004C97;color:#fff;padding:8px 16px;border-radius:4px;text-decoration:none;font-size:13px;">
        View Pipeline Monitor
      </a>
    </p>
  </div>
</div>"""
    text_body = (
        "PACE Validation ALERT!!!\n\n"
        "Hi Mike,\n\n"
        f"PACE (Pac-Biz Automated Coaching Engine) has detected {len(errors)} coaching {session_word} "
        "that were excluded from the Coaching Report due to invalid or incorrect employee email addresses entered in Asana.\n\n"
        + "\n\n".join(rows_text)
        + f"\n\nValidation details: {COACHING_VALIDATION_ERRORS_FILE}"
        + "\nPACE AI Email notification Activated as Watchers.. Report to Mike"
    )

    if notify.send(subject, html_body, body_text=text_body):
        save_coaching_validation_notification_state({
            "signature": signature,
            "count": len(errors),
            "sent_at": datetime.now(COACHING_TIMEZONE).isoformat(timespec="seconds"),
        })


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
    notify_coaching_validation_alerts(validation_errors)
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
          <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22,12 18,12 15,21 9,3 6,12 2,12"/></svg>Overall QA Score Trend</div><div class="qa-cs" id="{aid}-trend-sub">Weekly avg (Mon&ndash;Sun) &middot; Target: {threshold}%</div></div><span class="qa-cb qa-cbg" id="{aid}-trend-badge">&mdash;</span></div>
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
              <th style="min-width:120px">Evaluation ID</th>
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
    _ml_cache   = Path("masterlist_cache.csv")
    _hist_cache = Path("history_cache.csv")
    _move_cache = Path("movement_cache.csv")
    masterlist = clean_columns(pd.read_csv(_ml_cache   if _ml_cache.exists()   else MASTERLIST_CSV))
    history    = clean_columns(pd.read_csv(_hist_cache if _hist_cache.exists() else HISTORY_CSV))
    movement   = clean_columns(pd.read_csv(_move_cache if _move_cache.exists() else MOVEMENT_CSV))
    coaching = load_coaching_data(masterlist)
    # NOTE: the "Emp Name" rebuild below is intentionally deferred until AFTER
    # load_coaching_data() so Coaching sees the ORIGINAL untouched "Emp Name"
    # (transform_coaching_logs -> build_email_name_lookup reads masterlist["Emp Name"]).
    # Everything below only feeds the Masterlist tab / Movements table, which
    # run later, so mutating masterlist here is safe.
    if "Last Name" in masterlist.columns and "First Name" in masterlist.columns:
        # Rebuild "Emp Name" as strict two-part "Lastname, Firstname" (middle name dropped).
        _ln = masterlist["Last Name"].astype(str).str.strip()
        _fn = masterlist["First Name"].astype(str).str.strip()
        _ln = _ln.where(_ln.str.lower() != "nan", "")
        _fn = _fn.where(_fn.str.lower() != "nan", "")
        masterlist["Emp Name"] = [
            f"{ln}, {fn}" if ln and fn else (ln or fn)
            for ln, fn in zip(_ln, _fn)
        ]
    if "Type of Movement" in movement.columns and "Movement Type" in movement.columns:
        movement["Movement Type"] = movement.apply(
            lambda r: r["Movement Type"] if (pd.notna(r["Movement Type"]) and str(r["Movement Type"]).strip())
                      else str(r.get("Type of Movement", "") or "").strip(),
            axis=1
        )
    if "Company Email" in masterlist.columns and "Emp Name" in masterlist.columns:
        email_to_name = (
            masterlist[["Company Email", "Emp Name"]]
            .dropna(subset=["Company Email"])
            .assign(**{"Company Email": lambda d: d["Company Email"].str.strip().str.lower()})
            .drop_duplicates("Company Email")
            .set_index("Company Email")["Emp Name"]
            .to_dict()
        )
    else:
        email_to_name = {}
    if "Email Address" in movement.columns:
        movement["Initiated by"] = movement["Email Address"].apply(
            lambda e: email_to_name.get(str(e).strip().lower(), "") if pd.notna(e) and str(e).strip() else ""
        )
    else:
        movement["Initiated by"] = ""
    m7 = load_m7_data()
    dmg = load_dmg_data()
    r4h = load_r4h_data()
    parentis = load_parentis_data()
    britelift = load_britelift_data()
    blc = load_blc_data()
    ridex = load_ridex_data()
    hamilton = load_hamilton_data()
    skyline = load_skyline_data()
    vip = load_vip_data()
    ch = load_ch_data()
    rc = load_rc_data()
    ti = load_ti_data()
    dc = load_dc_data()
    ac = load_ac_data()
    ol = load_ol_data()
    ct = load_ct_data()
    ycov = load_ycov_data()
    kel = load_kel_data()
    vt = load_vt_data()
    ycdc = load_ycdc_data()
    bl = load_bl_data()

    refresh_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    refresh_iso  = datetime.now().strftime("%Y-%m-%dT%H:%M:00")

    latest_snapshot = ""
    if "Date Generated" in history.columns:
        latest_date = pd.to_datetime(history["Date Generated"], errors="coerce").max()
        latest_snapshot = latest_date.strftime("%m/%d/%Y") if pd.notna(latest_date) else ""

    # Masterlist KPI strip scalar (avg tenure of active staff).
    # Computed here in Python; JS does not currently calculate this.
    masterlist_kpis = {"avgTenure": 0}
    if "Hire Date" in masterlist.columns:
        _hire_dt = pd.to_datetime(masterlist["Hire Date"], errors="coerce")  # temp only, does not overwrite display column
        _now = datetime.now()
        if "Employment Status" in masterlist.columns:
            _active_mask = masterlist["Employment Status"].astype(str).str.strip().str.upper() == "ACTIVE"
        else:
            _active_mask = pd.Series(True, index=masterlist.index)
        _active_tenure_days = (_now - _hire_dt[_active_mask]).dt.days.dropna()
        if len(_active_tenure_days):
            masterlist_kpis["avgTenure"] = round(float(_active_tenure_days.mean()) / 365.25, 1)

    # Movement / history KPI strip scalars for the redesigned 7-tile KPI strip.
    # movement sheet has no literal "Status" column, so processing status is
    # read from "Processed" (Yes/blank), which is column R (18th column,
    # position index 17) as the sheet is currently laid out. "Void" marks a
    # row as cancelled regardless of its processed state.
    # "Movements" (main tile value) = total logged movements, excluding voided
    # rows entirely — a voided row was never a real movement.
    # "for processing" (sub-value) = the subset of those non-void movements
    # whose "Processed" column is still blank (i.e. genuinely pending action).
    masterlist_kpis["movementsPending"] = 0
    masterlist_kpis["movementsForProcessing"] = 0
    masterlist_kpis["historyRecords"] = int(len(history))
    if not movement.empty and len(movement.columns) > 17:
        _colR_name = None
        for _c in movement.columns:
            if str(_c).strip().lower() == "processed":
                _colR_name = _c
                break
        if _colR_name is None:
            _colR_name = movement.columns[17]  # positional fallback; currently "Processed"
            print("WARN: 'Processed' column not found by name; falling back to column index 17:", _colR_name)
        _colR_blank = movement[_colR_name].fillna("").astype(str).str.strip() == ""
        if "Void" in movement.columns:
            _not_void = movement["Void"].fillna("").astype(str).str.strip().str.upper() != "YES"
        else:
            _not_void = pd.Series(True, index=movement.index)
        # Movements: total non-void movement rows logged (voided rows excluded
        # entirely — they never happened).
        masterlist_kpis["movementsPending"] = int(_not_void.sum())
        # For processing: of those non-void movements, the ones whose
        # "Processed" column is still blank and therefore need action.
        masterlist_kpis["movementsForProcessing"] = int((_colR_blank & _not_void).sum())

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
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
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

    .tab-button[data-tab="schedule"] {{ --tab-accent: #39B54A; }}

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
        /* Change 14: the footer is now sticky-bottom (global, all tabs) — reserve
           clearance so it never overlaps the last row of whichever tab is active. */
        padding-bottom: 40px;
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

    /* Change 3: Recent Employee Movements header — explicit white-bold-on-blue,
       scoped to the #recentMovements container only (generic renderDataTable()
       output has no id/class of its own to hook, so this is the one place we
       target the container id directly) so the rule can never leak onto the
       other tables that share the same generic table/th markup (Coaching
       summary/logs tables, QA leaderboards, etc.). */
    #recentMovements table thead th {{
        background: var(--blue);
        color: #fff;
        font-weight: 700;
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
    body.qa-dist-focus-mode #qa-dist-focus-toggle .qa-focus-icon::before {{ top: 2px; right: 2px; border-width: 0 0 2px 2px; }}
    body.qa-dist-focus-mode #qa-dist-focus-toggle .qa-focus-icon::after  {{ left: 2px; bottom: 2px; border-width: 2px 2px 0 0; }}
    body.qa-dist-focus-mode #qualityPanel .qa-kpi-strip,
    body.qa-dist-focus-mode #qualityPanel .qa-kpi-row,
    body.qa-dist-focus-mode #qualityPanel .qa-sum-strip,
    body.qa-dist-focus-mode #qualityPanel .qa-g2,
    body.qa-dist-focus-mode #qa-aivh-card,
    body.qa-dist-focus-mode #qa-detail-card {{ display: none; }}
    body.qa-dist-focus-mode #qualityPanel .qa-g3 {{ display: flex; justify-content: center; align-items: flex-start; grid-template-columns: 1fr; }}
    body.qa-dist-focus-mode #qualityPanel .qa-g3 > .qa-card:not(#qa-dist-card) {{ display: none; }}
    body.qa-dist-focus-mode #qa-dist-card {{ min-height: 0; width: 95vw; max-width: 1900px; height: 75vh; max-height: 900px; }}
    body.qa-dist-focus-mode #qa-dist-card .qa-ch {{ position: relative; z-index: 3; }}
    body.qa-dist-focus-mode #qa-dist-focus-toggle {{ position: relative; z-index: 4; }}
    body.qa-dist-focus-mode #qa-dist-card .qa-cbody {{ height: calc(100% - 56px); padding: 14px 18px !important; }}
    body.qa-dist-focus-mode #qa-score-dist-wrap,
    body.qa-dist-focus-mode #qa-eval-dist-wrap {{ height: 100%; }}
    body.qa-dist-focus-mode #qa-eval-dist-wrap {{ display: grid; grid-template-columns: minmax(0, 1.7fr) minmax(320px, 0.85fr); align-items: center; gap: 28px; min-height: 0; }}
    body.qa-dist-focus-mode #qa-score-chart-host,
    body.qa-dist-focus-mode #qa-eval-chart-host {{ height: min(64vh, 780px) !important; }}
    body.qa-dist-focus-mode #qa-eval-chart-host {{ height: 100% !important; min-height: 0; }}
    body.qa-dist-focus-mode #qa-donut-legend,
    body.qa-dist-focus-mode #qa-eval-dist-legend {{ grid-template-columns: repeat(3, minmax(0, 1fr)) !important; gap: 8px !important; margin-top: 12px !important; }}
    body.qa-dist-focus-mode #qa-eval-dist-legend {{ display: block !important; align-self: center; max-height: 100%; overflow-y: auto; margin-top: 0 !important; background: #fff; border: 1px solid #D8E1EC; border-radius: 8px; box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08); }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: #0F172A; }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-table th {{ padding: 10px 12px; text-align: left; font-weight: 800; color: #fff; background: var(--blue); border-bottom: 1px solid #D8E1EC; }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-table th:nth-child(2),
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-table th:nth-child(3) {{ text-align: right; }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-table td {{ padding: 7px 12px; border-bottom: 1px solid #E5EAF1; vertical-align: middle; }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-table td:nth-child(2),
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-table td:nth-child(3) {{ text-align: right; font-variant-numeric: tabular-nums; }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-account {{ display: flex; align-items: center; gap: 8px; min-width: 0; }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-dot {{ width: 12px; height: 12px; border-radius: 50%; flex: 0 0 auto; box-shadow: inset 0 0 0 1px rgba(15, 23, 42, 0.08); }}
    body.qa-dist-focus-mode #qa-eval-dist-legend .qa-dist-total td {{ font-weight: 900; color: #fff; background: var(--green); border-bottom: 0; border-top: 1px solid #D8E1EC; }}
    #qa-eval-dist-other-note {{ display: none; }}
    body.qa-dist-focus-mode #qa-eval-dist-other-note {{ display: block; position: absolute; left: 10px; right: 10px; bottom: 4px; text-align: center; font-size: 11px; line-height: 1.25; font-style: italic; color: #64748B; }}
    #qa-eval-dist-legend.qa-legend-compact {{ display: block !important; max-height: 82px; overflow-y: auto; border: 1px solid #D8E1EC; border-radius: 7px; background: #fff; box-shadow: inset 0 1px 0 rgba(255,255,255,0.65); }}
    #qa-eval-dist-legend .qa-legend-head,
    #qa-eval-dist-legend .qa-legend-row {{ display: grid; grid-template-columns: minmax(0,1fr) 46px; align-items: center; column-gap: 8px; }}
    #qa-eval-dist-legend .qa-legend-head {{ position: sticky; top: 0; z-index: 1; padding: 4px 7px; background: var(--green); border-bottom: 1px solid rgba(15,23,42,.12); color: #fff; font-size: 8px; font-weight: 800; text-transform: uppercase; letter-spacing: .04em; }}
    #qa-eval-dist-legend .qa-legend-row {{ padding: 4px 7px; border-bottom: 1px solid #EEF2F7; font-size: 9px; color: #334155; }}
    #qa-eval-dist-legend .qa-legend-row:last-child {{ border-bottom: 0; }}
    #qa-eval-dist-legend .qa-legend-account {{ display: flex; align-items: center; gap: 5px; min-width: 0; }}
    #qa-eval-dist-legend .qa-legend-dot {{ width: 7px; height: 7px; border-radius: 50%; flex: 0 0 auto; box-shadow: inset 0 0 0 1px rgba(15,23,42,.08); }}
    #qa-eval-dist-legend .qa-legend-name {{ overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    #qa-eval-dist-legend .qa-legend-pct {{ text-align: right; font-weight: 800; font-variant-numeric: tabular-nums; }}
    body.qa-aiqe-focus-mode #qa-aiqe-focus-toggle .qa-focus-icon::before {{ top: 2px; right: 2px; border-width: 0 0 2px 2px; }}
    body.qa-aiqe-focus-mode #qa-aiqe-focus-toggle .qa-focus-icon::after  {{ left: 2px; bottom: 2px; border-width: 2px 2px 0 0; }}
    body.qa-aiqe-focus-mode #qualityPanel .qa-kpi-strip,
    body.qa-aiqe-focus-mode #qualityPanel .qa-kpi-row,
    body.qa-aiqe-focus-mode #qualityPanel .qa-sum-strip,
    body.qa-aiqe-focus-mode #qualityPanel .qa-g2,
    body.qa-aiqe-focus-mode #qa-aivh-card,
    body.qa-aiqe-focus-mode #qa-detail-card {{ display: none; }}
    body.qa-aiqe-focus-mode #qualityPanel .qa-g3 {{ display: flex; justify-content: center; align-items: flex-start; }}
    body.qa-aiqe-focus-mode #qualityPanel .qa-g3 > .qa-card {{ display: none; }}
    body.qa-aiqe-focus-mode #qualityPanel .qa-g3 > #qa-aiqe-card {{ display: flex; }}
    body.qa-aiqe-focus-mode #qa-aiqe-card {{ min-height: 0; width: 95vw; max-width: 1900px; height: 75vh; max-height: 900px; }}
    body.qa-aiqe-focus-mode #qa-aiqe-card .qa-cbody {{ height: calc(100% - 56px); padding: 14px 18px !important; }}
    body.qa-aiqe-focus-mode #qa-aiqe-chart-host {{ height: min(64vh, 780px); }}

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

    /* Host div: propagates card height through to the widget/summary/footer chain */
    .coaching-donut-host {{
        flex: 1;
        min-height: 0;
        display: flex;
        flex-direction: column;
    }}
    .coaching-donut-host > .coaching-donut-widget {{
        flex: 1;
        min-height: 0;
    }}
    /* Footer anchored to bottom via margin-top: auto */
    .coaching-donut-host .chart-summary-total {{
        margin-top: auto;
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
        padding: 6px 28px;
        text-align: center;
        font-size: 11px;
        position: sticky;
        bottom: 0;
        z-index: 500; /* below the expand modal (9999), above the sticky KPI strip (150) */
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
    #qualityPanel .qa-b-indigo {{ background:#EEF2FF;color:#4338CA;border-color:#C7D2FE }}
    #qualityPanel .qa-b-orange {{ background:#FFF7ED;color:#C2410C;border-color:#FED7AA }}
    #qualityPanel .qa-b-rose {{ background:#FFF1F2;color:#BE123C;border-color:#FECDD3 }}
    #qualityPanel .qa-b-cyan {{ background:#ECFEFF;color:#0E7490;border-color:#A5F3FC }}
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
    #qualityPanel .qa-g3 {{ display:grid;grid-template-columns:1.12fr 1.12fr 0.92fr 0.92fr 0.82fr;gap:10px;grid-auto-rows:300px;align-items:stretch }}
    #qualityPanel .qa-g2 {{ display:grid;grid-template-columns:repeat(3,1fr);gap:12px;grid-auto-rows:300px;align-items:stretch }}
    #qualityPanel .qa-card {{ background:#fff;border-radius:10px;border:1px solid #E2E8F0;overflow:hidden;display:flex;flex-direction:column }}
    #qualityPanel .qa-g3 .qa-card {{ height:100%;min-height:0 }}
    #qualityPanel .qa-g3 .qa-card .qa-cbody {{ min-height:0 }}
    #qualityPanel .qa-ch {{ padding:7px 10px;border-bottom:1px solid #F1F5F9;display:flex;align-items:center;justify-content:space-between;gap:8px }}
    #qualityPanel .qa-ct {{ font-size:11px;font-weight:700;color:#1E293B;display:flex;align-items:center;gap:5px }}
    #qualityPanel .qa-ct svg {{ width:13px;height:13px;color:#94A3B8 }}
    #qualityPanel .qa-cs {{ font-size:10px;color:#94A3B8;margin-top:1px }}
    #qualityPanel .qa-cbody {{ padding:8px 10px;flex:1 }}
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
    #qualityPanel .qa-dtbl {{ width:max-content;min-width:100%;border-collapse:collapse;font-size:11px;table-layout:fixed }}
    #qualityPanel .qa-dtbl th {{ padding:7px 10px;text-align:left;font-size:9px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.05em;border-right:1px solid #D1D5DB;border-bottom:1px solid #E2E8F0;background:#F8FAFC;white-space:nowrap;position:sticky;top:0;z-index:1;overflow:hidden;text-overflow:ellipsis }}
    #qualityPanel .qa-dtbl th:last-child {{ border-right:none }}
    #qualityPanel .qa-dtbl th.qa-resizable {{ padding-right:16px }}
    #qualityPanel .qa-col-resizer {{ position:absolute;top:0;right:-4px;width:8px;height:100%;cursor:col-resize;z-index:7;opacity:1 }}
    #qualityPanel .qa-col-resizer::after {{ content:"";position:absolute;top:0;bottom:0;right:3px;width:1px;background:#9CA3AF;border-radius:1px }}
    #qualityPanel .qa-col-resizer::before {{ content:"⋮";position:absolute;top:50%;right:0;transform:translateY(-50%);font-size:9px;line-height:1;color:#94A3B8;opacity:.7;pointer-events:none }}
    #qualityPanel .qa-dtbl th:hover,
    body.qa-col-resizing #qualityPanel .qa-dtbl th {{ border-right-color:#94A3B8 }}
    #qualityPanel .qa-col-resizer:hover::after,
    body.qa-col-resizing #qualityPanel .qa-col-resizer::after {{ width:2px;background:#64748B }}
    #qualityPanel .qa-col-resizer:hover::before,
    body.qa-col-resizing #qualityPanel .qa-col-resizer::before {{ color:#64748B;opacity:1 }}
    body.qa-col-resizing {{ cursor:col-resize;user-select:none }}
    body.qa-col-resizing * {{ cursor:col-resize!important }}
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
    #qualityPanel .qa-dtbl td:nth-child(5){{min-width:125px;white-space:nowrap}}
    #qualityPanel .qa-dtbl td:nth-child(5) span,
    #qualityPanel .qa-lbt td:last-child span{{display:inline-flex;align-items:center;white-space:nowrap;line-height:1.15}}
    #qualityPanel .qa-dtbl thead th:nth-child(-n+5){{z-index:3;background:#F8FAFC}}
    #qualityPanel .qa-dtbl tbody td:nth-child(-n+5){{background:#fff}}
    #qualityPanel .qa-dtbl tr:hover td:nth-child(-n+5){{background:#F8FAFC}}
    .vip-extra-col{{display:none}}
    .vip-mode .vip-extra-col{{display:table-cell}}
    .ch-extra-col{{display:none}}
    .ch-mode .ch-extra-col{{display:table-cell}}
    .rc-extra-col{{display:none}}
    .rc-mode .rc-extra-col{{display:table-cell}}
    .ti-extra-col{{display:none}}
    .ti-mode .ti-extra-col{{display:table-cell}}
    .dc-extra-col{{display:none}}
    .dc-mode .dc-extra-col{{display:table-cell}}
    .ac-extra-col{{display:none}}
    .ac-mode .ac-extra-col{{display:table-cell}}
    .ol-extra-col{{display:none}}
    .ol-mode .ol-extra-col{{display:table-cell}}
    .ct-extra-col{{display:none}}
    .ct-mode .ct-extra-col{{display:table-cell}}
    .ycov-extra-col{{display:none}}
    .ycov-mode .ycov-extra-col{{display:table-cell}}
    .kel-extra-col{{display:none}}
    .kel-mode .kel-extra-col{{display:table-cell}}
    .vt-extra-col{{display:none}}
    .vt-mode .vt-extra-col{{display:table-cell}}
    .ycdc-extra-col{{display:none}}
    .ycdc-mode .ycdc-extra-col{{display:table-cell}}
    .bl-extra-col{{display:none}}
    .bl-mode .bl-extra-col{{display:table-cell}}
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

    /* ── Masterlist tab redesign (Phase 2 port of masterlist-poc.html) ──
       All new selectors are scoped under #masterlistPanel / #masterlistControls
       and every new class is ml- prefixed so nothing here can leak into or
       collide with Coaching's .card/.cards/.chart-card or Quality's .qa- rules.
       --ml-surface/--ml-border/--ml-shadow/--ml-r are NEW local tokens (not a
       redefinition of the shared :root --blue/--green/--bg/--text/--muted).
       --ml-pad-x/--ml-pad-y/--ml-gap (item D) are the shared spacing rhythm
       for every chart card's header/body padding and inter-card gaps — on
       the 4/8/12/16/24/32 scale, used instead of each card having its own
       ad hoc padding numbers. */
    #masterlistPanel, #masterlistControls {{
        --ml-surface: #fff;
        --ml-border: #DDE6EE;
        --ml-shadow: 0 1px 3px rgba(0,76,151,.07), 0 4px 14px rgba(0,76,151,.05);
        --ml-r: 8px;
        --ml-pad-x: 16px;
        --ml-pad-y: 12px;
        --ml-gap: 16px;
    }}

    /* Filter bar — same #masterlistControls markup/IDs/handlers, restyled via
       scoped override only (NOT a rename of the shared .filters/.filter-box/
       .multi-filter/.multi-options classes, which Coaching also uses).
       Change 1: compact pill-style dropdowns (POC look) + z-index/overflow fix
       so an open dropdown never renders behind the sticky KPI strip below it. */
    #masterlistControls .filters {{
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: flex-end;
        grid-template-columns: none;
        background: var(--ml-surface);
        border-radius: var(--ml-r);
        box-shadow: var(--ml-shadow);
        border-top: 3px solid var(--blue);
        padding: 10px 15px;
        margin: 0 18px 14px;
        overflow: visible; /* never clip the open dropdown menu below */
    }}
    #masterlistControls .filter-box {{
        background: none;
        border: none;
        box-shadow: none;
        padding: 0;
        display: flex;
        flex-direction: column;
        gap: 4px;
        min-width: 132px;
        overflow: visible;
    }}
    #masterlistControls .filter-box label {{
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: .06em;
        color: var(--muted);
        margin-bottom: 0;
    }}
    #masterlistControls .multi-filter summary {{
        padding: 6px 14px;
        min-height: 32px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        border: 1.5px solid var(--ml-border);
        border-radius: 999px;
        font-size: 12px;
        background: var(--ml-surface);
        color: var(--text);
        transition: border-color .15s, box-shadow .15s;
    }}
    #masterlistControls .multi-filter summary:hover {{ border-color: var(--blue); }}
    #masterlistControls .multi-filter[open] summary {{ border-color: var(--blue); box-shadow: 0 0 0 3px rgba(0,76,151,.08); }}
    /* Bug fix (change 1): base .multi-options is z-index:20 — lower than the
       sticky KPI strip's z-index:150 below it, so an opened dropdown painted
       behind the strip. Raise it well above (300) without touching the base
       class (Coaching's own filter bar keeps its original stacking). */
    /* Fix (item 1): base .multi-options is left:0;right:0 — same width as the
       ~132px filter-box, too narrow for longer option labels and forcing
       mid-word wrapping. Let it size to its own content (clamped) so it opens
       wide enough for a clean single line per option. */
    #masterlistControls .multi-options {{
        left: 0;
        right: auto;
        min-width: 220px;
        max-width: 320px;
        border-radius: 8px;
        box-shadow: 0 12px 28px rgba(15,23,42,.15);
        z-index: 300;
        padding: 6px;
    }}
    /* Fix (item 1): consistent per-row height, checkbox + label on one line,
       adequate padding, subtle hover — scoped so Coaching's own multi-select
       filter (same base .multi-option class) is untouched. */
    #masterlistControls .multi-option {{
        display: flex;
        align-items: center;
        flex-wrap: nowrap;
        gap: 10px;
        padding: 9px 10px;
        min-height: 34px;
        border-radius: 6px;
        font-size: 12px;
    }}
    #masterlistControls .multi-option span {{
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}
    #masterlistControls .multi-option input[type="checkbox"] {{
        flex-shrink: 0;
        width: 14px;
        height: 14px;
        margin: 0;
        accent-color: var(--blue);
    }}
    #masterlistControls .multi-option:hover {{ background: var(--bg); }}
    #masterlistControls .multi-option.select-all {{
        padding-bottom: 12px;
        margin-bottom: 6px;
    }}

    /* Clear filters button (item 6) — reuses the same .ml-btn/.ml-sec visual
       pattern already used by the chart-card action buttons (defined under
       #masterlistPanel further below); duplicated here under #masterlistControls
       because the filter bar lives outside #masterlistPanel, so it can't inherit
       that scoped rule. Pushed to the end of the filter row via margin-left:auto,
       matching the POC's f-clear-btn placement/behavior. */
    #masterlistControls .ml-btn {{
        padding: 5px 12px; border-radius: 4px; font-size: 12px; font-weight: 600; cursor: pointer; border: 1.5px solid var(--blue);
        background: var(--blue); color: #fff; font-family: Arial, sans-serif; transition: opacity .15s; text-decoration: none;
        display: inline-flex; align-items: center; justify-content: center; min-height: 32px;
    }}
    #masterlistControls .ml-btn:hover {{ opacity: .85; }}
    #masterlistControls .ml-btn:focus-visible {{ outline: 2px solid var(--blue); outline-offset: 1px; }}
    #masterlistControls .ml-btn.ml-sec {{ background: none; color: var(--blue); }}
    #masterlistControls .ml-clear-btn {{ margin-left: auto; align-self: flex-end; }}

    /* KPI strip — brand-new 5-tile sticky strip. Deliberately NOT the shared
       .cards/.card classes (those are global — used by Coaching's own KPI
       tiles) so this cannot restyle anything outside the Masterlist tab. */
    #masterlistControls .ml-kpi-strip {{
        display: grid;
        grid-template-columns: repeat(7, minmax(0,1fr));
        gap: 12px;
        padding: 0 18px 14px;
        position: sticky;
        top: var(--ml-stick-top, 0px);
        z-index: 150;
        background: var(--bg);
    }}
    #masterlistControls .ml-kpi {{
        background: var(--ml-surface);
        border-radius: var(--ml-r);
        padding: 13px 15px 11px;
        box-shadow: var(--ml-shadow);
        border-left: 3px solid var(--ml-border);
    }}
    #masterlistControls .ml-kpi.g {{ border-left-color: var(--green); }}
    #masterlistControls .ml-kpi.b {{ border-left-color: var(--blue); }}
    #masterlistControls .ml-kpi-lbl {{
        font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .08em;
        color: var(--muted); margin-bottom: 2px;
    }}
    #masterlistControls .ml-kpi-val {{
        font-size: 26px; font-weight: 700; line-height: 1.1; font-variant-numeric: tabular-nums; color: var(--text);
    }}
    #masterlistControls .ml-kpi.g .ml-kpi-val {{ color: var(--green); }}
    #masterlistControls .ml-kpi.b .ml-kpi-val {{ color: var(--blue); }}
    #masterlistControls .ml-kpi-sub {{ font-size: 11px; color: var(--muted); margin-top: 2px; }}

    /* Chart cards — new ml-card pattern, distinct from the shared .card/.chart-card */
    #masterlistPanel .ml-grid {{ display: flex; flex-direction: column; gap: var(--ml-gap); }}
    /* Item A fix: align-items:start (was the grid default "stretch") so each
       card in a row keeps ITS OWN natural content height instead of being
       stretched to match the tallest card in the row (e.g. "By Department"'s
       long legend). Stretching was the real cause of inconsistent legend
       anchoring — a short donut (Active Status) would get stretched tall by
       its row-mate, then its canvas centered inside that leftover space,
       so the legend never actually sat at the visual bottom of the card. */
    #masterlistPanel .ml-row {{ display: grid; gap: var(--ml-gap); align-items: start; }}
    /* The five donut cards use one shared canvas height, so stretching this
       row keeps the card shells aligned without letting legends float. */
    #masterlistPanel .ml-r5 {{ align-items: stretch; }}
    #masterlistPanel .ml-r5 {{ grid-template-columns: repeat(5, minmax(0,1fr)); }}
    #masterlistPanel .ml-r4 {{ grid-template-columns: repeat(4, minmax(0,1fr)); }}
    /* Change 2: this row only (Tenure by Account + Weekly Headcount Trend)
       overrides the row default above back to align-items:stretch so both
       cards render at the SAME height (the row's height is still driven by
       the taller card's natural content, same as before — stretch just makes
       the shorter card's box grow to match it instead of sitting shorter).
       Scoped to .ml-r2asym specifically (same specificity as .ml-row, but
       wins on source order) so the other rows keep the donut-legend
       bottom-anchoring fix above untouched. */
    #masterlistPanel .ml-r2asym {{ grid-template-columns: 3fr 2fr; align-items: stretch; }}
    #masterlistPanel .ml-card {{
        background: var(--ml-surface);
        border-radius: var(--ml-r);
        box-shadow: var(--ml-shadow);
        border-top: 3px solid var(--blue);
        display: flex;
        flex-direction: column;
        overflow: hidden;
        transition: box-shadow .15s;
    }}
    #masterlistPanel .ml-card:hover {{ box-shadow: 0 2px 8px rgba(0,76,151,.12), 0 8px 24px rgba(0,76,151,.08); }}
    #masterlistPanel .ml-card.gc {{ border-top-color: var(--green); }}
    /* Item D: header padding uses the shared --ml-pad-x/--ml-pad-y rhythm
       (was the ad hoc 11px 13px 9px) so every card's header aligns exactly
       with its body below. */
    #masterlistPanel .ml-card-hd {{
        display: flex; align-items: flex-start; justify-content: space-between; padding: var(--ml-pad-y) var(--ml-pad-x) 8px;
        border-bottom: 1px solid var(--ml-border); gap: 8px;
    }}
    #masterlistPanel .ml-card-ttl {{ font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .07em; color: var(--blue); margin-bottom: 1px; }}
    #masterlistPanel .ml-card.gc .ml-card-ttl {{ color: #1a7a2e; }}
    #masterlistPanel .ml-card-sub {{ font-size: 11px; color: var(--muted); }}
    /* Item D: body padding uses the same shared rhythm, symmetric top/bottom
       (was 10px 12px 13px) so the left/right edge lines up with the header
       above. min-height stays — it's the floor for a very short donut
       (e.g. Employment Class) before the card grows to fit real content. */
    #masterlistPanel .ml-card-bd {{ padding: var(--ml-pad-y) var(--ml-pad-x); flex: 1; display: flex; align-items: center; justify-content: center; min-height: 170px; }}
    #masterlistPanel .ml-card-bd canvas {{ display: block; max-width: 100%; box-sizing: border-box; }}
    /* Fix (item 5): the max-width above is a safety net so a bad first-paint
       measurement can never push a canvas wider than its card and cause a
       page-level horizontal scrollbar. Exempt the two charts that deliberately
       draw WIDER than the visible card and rely on their own .ml-hscroll-x
       wrapper (overflow-x:auto) to scroll — clamping them here would silently
       break that intentional natural-width/horizontal-scroll behavior. */
    #masterlistPanel .ml-hscroll-x canvas {{ max-width: none; }}
    /* Active Status donut — blinking legend dots (cosmetic only, nothing
       moves). Canvas can't animate, so a small absolutely-positioned DOM
       dot is overlaid exactly on top of each canvas-drawn legend swatch for
       THIS card only (see mlSyncActiveLegendDots/mlDonut, id "c-ml-active").
       position:relative on this one card's body gives the dots a local
       positioning context without affecting any other card's layout. */
    #masterlistPanel .ml-card[data-cid="active"] .ml-card-bd {{ position: relative; }}
    .ml-legend-pulse-dot {{
        position: absolute; border-radius: 50%; pointer-events: none;
        animation: qa-pulse 2s ease-in-out infinite;
    }}
    @media (prefers-reduced-motion: reduce) {{
        .ml-legend-pulse-dot {{ animation: none; }}
    }}
    /* Item D: left padding matches --ml-pad-x (16px, was 12px) so this card's
       content lines up with every other card's left edge; top/bottom match
       --ml-pad-y too. Right stays a literal 6px — NOT the shared token — by
       design: it's deliberately thin to leave room for .ml-hbar-scroll's own
       6px scrollbar gutter (12px combined, matching the left edge), not an
       ad hoc leftover value. */
    #masterlistPanel .ml-card-bd.ml-hbar-bd {{ align-items: stretch; justify-content: flex-start; padding: var(--ml-pad-y) 6px var(--ml-pad-y) var(--ml-pad-x); }}
    #masterlistPanel .ml-hbar-scroll {{ width: 100%; max-height: 242px; overflow-y: auto; padding-right: 6px; }}
    #masterlistPanel .ml-hbar-scroll::-webkit-scrollbar {{ width: 8px; }}
    #masterlistPanel .ml-hbar-scroll::-webkit-scrollbar-track {{ background: transparent; }}
    #masterlistPanel .ml-hbar-scroll::-webkit-scrollbar-thumb {{ background: var(--ml-border); border-radius: 4px; }}
    #masterlistPanel .ml-hbar-scroll:hover::-webkit-scrollbar-thumb {{ background: var(--muted); }}

    /* Horizontal-scroll pattern (robustness fixes B1/B2) — analogous to
       .ml-hbar-scroll above but scrolls sideways. Used by "Tenure by Account"
       (22+ accounts) and "Weekly Headcount Trend" (grows every week) so bars
       get a real minimum width instead of being silently squeezed/clipped;
       the canvas takes its natural computed width and this wrapper scrolls
       within the card's fixed-width body. Same item-D padding rationale as
       .ml-hbar-bd above — 16px left/12px top-bottom from the shared tokens,
       6px right reserved for .ml-hscroll-x's own scrollbar gutter. */
    #masterlistPanel .ml-card-bd.ml-hscroll-bd {{ align-items: stretch; justify-content: flex-start; padding: var(--ml-pad-y) 6px var(--ml-pad-y) var(--ml-pad-x); }}
    #masterlistPanel .ml-hscroll-x {{ width: 100%; overflow-x: auto; overflow-y: hidden; padding-bottom: 6px; }}
    #masterlistPanel .ml-hscroll-x::-webkit-scrollbar {{ height: 8px; }}
    #masterlistPanel .ml-hscroll-x::-webkit-scrollbar-track {{ background: transparent; }}
    #masterlistPanel .ml-hscroll-x::-webkit-scrollbar-thumb {{ background: var(--ml-border); border-radius: 4px; }}
    #masterlistPanel .ml-hscroll-x:hover::-webkit-scrollbar-thumb {{ background: var(--muted); }}

    /* Expand button — hover-only, keyboard-visible. Own ml-xbtn class — NOT
       the Coaching tab's .coaching-expand-btn. */
    #masterlistPanel .ml-xbtn {{
        background: none; border: 1px solid var(--ml-border); cursor: pointer; padding: 5px 6px; border-radius: 4px;
        color: var(--muted); display: flex; align-items: center; justify-content: center;
        transition: background .15s, color .15s, border-color .15s, opacity .15s; flex-shrink: 0; opacity: 0;
    }}
    #masterlistPanel .ml-card:hover .ml-xbtn, #masterlistPanel .ml-card:focus-within .ml-xbtn,
    #masterlistPanel .ml-tcrd:hover .ml-xbtn, #masterlistPanel .ml-tcrd:focus-within .ml-xbtn,
    #masterlistPanel .ml-xbtn:focus {{ opacity: 1; }}
    #masterlistPanel .ml-xbtn:hover {{ background: var(--bg); color: var(--blue); border-color: var(--blue); }}
    #masterlistPanel .ml-xbtn:focus-visible {{ outline: 2px solid var(--blue); outline-offset: 1px; }}

    /* Master List table card */
    #masterlistPanel .ml-tcrd {{
        background: var(--ml-surface); border-radius: var(--ml-r); box-shadow: var(--ml-shadow);
        border-top: 3px solid var(--blue); overflow: hidden;
    }}
    #masterlistPanel .ml-thd {{ display: flex; align-items: center; justify-content: space-between; padding: 11px 15px; border-bottom: 1px solid var(--ml-border); flex-wrap: wrap; gap: 8px; }}
    #masterlistPanel .ml-ttl {{ font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .07em; color: var(--blue); }}
    #masterlistPanel .ml-tacts {{ display: flex; gap: 7px; align-items: center; flex-wrap: wrap; }}
    /* Item 1 perf fix: hint the browser to isolate this scroller as its own
       paint/composite layer so horizontal-scrollbar drag doesn't force a
       repaint of ancestor/sibling content each frame (the likely cause of
       the reported drag lag — see report). `contain: paint` clips/isolates
       painting to this box; `will-change: scroll-position` is the
       purpose-built hint for scrollable containers (cheaper than promoting
       via `transform`). Deliberately NOT applied to the sticky th/td cells
       themselves or to any element between them and this scroller, per Item
       3's compatibility note — sticky positioning still resolves against
       this element as the nearest scrolling ancestor. */
    /* min-height matches max-height so the card keeps its full size even when
       a filter returns few or zero rows — the table area never collapses. */
    #masterlistPanel .ml-twrap {{ overflow-x: auto; min-height: 420px; max-height: 420px; overflow-y: auto; contain: paint; will-change: scroll-position; }}
    #masterlistPanel table.ml-dt {{ width: 100%; border-collapse: collapse; font-size: 12px; font-variant-numeric: tabular-nums; }}
    /* Change 3: header restyled to white bold on Pac-Biz blue (matches the QA
       detail table header convention) — was gray text (var(--muted)) on a
       light var(--bg) row, low contrast. Sort/hover states now darken the
       blue background instead of changing text color, so the white text and
       its sort-indicator glyph (.ml-sort-ic, inherits color) stay legible in
       every state — changing the text color itself would have gone
       white-on-blue -> blue-on-blue (invisible) on hover/sorted. */
    #masterlistPanel .ml-dt thead tr {{ background: var(--blue); }}
    #masterlistPanel .ml-dt th {{
        padding: 7px 12px; text-align: left; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .07em;
        color: #fff; border-bottom: 1px solid var(--ml-border); white-space: nowrap; cursor: default; user-select: none;
    }}
    /* Item 3: freeze columns through "Employee Name" (cols 1-2: ID No., Emp
       Name) on horizontal scroll — Excel freeze-panes behavior. Fixed
       width+box-sizing:border-box makes the left offsets below deterministic
       even though the table itself stays table-layout:auto. Placed BEFORE the
       .ml-sorted/:hover rules further down so those still win the header
       background on tie (equal specificity, source-order tiebreak) — sort/
       hover feedback on the frozen header cells is preserved. Backgrounds are
       flat/opaque (no gradients) and no `contain`/`will-change`/`transform`
       is applied to these cells themselves — only to the .ml-twrap scroll
       ancestor below — so sticky positioning keeps working (see Item 1 fix). */
    #masterlistPanel .ml-dt th:nth-child(1), #masterlistPanel .ml-dt td:nth-child(1),
    #masterlistPanel .ml-dt th:nth-child(2), #masterlistPanel .ml-dt td:nth-child(2) {{
        position: sticky; box-sizing: border-box;
    }}
    #masterlistPanel .ml-dt th:nth-child(1), #masterlistPanel .ml-dt td:nth-child(1) {{
        left: 0; width: 116px; min-width: 116px; max-width: 116px;
    }}
    #masterlistPanel .ml-dt th:nth-child(2), #masterlistPanel .ml-dt td:nth-child(2) {{
        left: 116px; width: 220px; min-width: 220px; max-width: 220px;
        box-shadow: 2px 0 4px rgba(0,0,0,.06); /* freeze boundary divider */
    }}
    #masterlistPanel .ml-dt th:nth-child(1), #masterlistPanel .ml-dt th:nth-child(2) {{ z-index: 4; background: var(--blue); }}
    #masterlistPanel .ml-dt td:nth-child(1), #masterlistPanel .ml-dt td:nth-child(2) {{ z-index: 2; background: #fff; }}
    #masterlistPanel #ml-thead th[data-key] {{ cursor: pointer; }}
    #masterlistPanel #ml-thead th[data-key]:hover {{ background: #003B73; }}
    #masterlistPanel .ml-dt th.ml-sorted {{ background: #003B73; }}
    #masterlistPanel .ml-dt th.ml-th-static {{ cursor: default; }}
    #masterlistPanel .ml-dt th.ml-th-static:hover {{ background: var(--blue); }}
    #masterlistPanel .ml-dt th .ml-sort-ic {{ margin-left: 2px; color: #fff; }}
    #masterlistPanel .ml-dt th[data-key]:focus-visible {{ outline: 2px solid #fff; outline-offset: -2px; }}
    #masterlistPanel .ml-dt td {{ padding: 7px 12px; border-bottom: 1px solid var(--ml-border); color: var(--text); white-space: nowrap; }}
    #masterlistPanel .ml-dt tbody tr:last-child td {{ border-bottom: none; }}
    #masterlistPanel .ml-dt tbody tr:hover td {{ background: rgba(0,76,151,.025); }}
    #masterlistPanel .ml-dt .ml-col-id {{ min-width: 76px; }}
    #masterlistPanel .ml-dt .ml-col-name {{ min-width: 180px; }}
    #masterlistPanel .ml-dt .ml-col-date {{ min-width: 96px; }}
    #masterlistPanel .ml-dt .ml-col-class {{ min-width: 132px; }}
    #masterlistPanel .ml-dt .ml-col-tenure {{ min-width: 132px; }}
    #masterlistPanel .ml-dt .ml-col-title {{ min-width: 210px; }}
    #masterlistPanel .ml-dt .ml-col-group {{ min-width: 145px; }}
    #masterlistPanel .ml-dt .ml-col-dept {{ min-width: 130px; }}
    #masterlistPanel .ml-dt .ml-col-account {{ min-width: 150px; }}
    #masterlistPanel .ml-dt .ml-col-supervisor,
    #masterlistPanel .ml-dt .ml-col-manager {{ min-width: 180px; }}
    #masterlistPanel .ml-dt .ml-col-status {{ min-width: 140px; }}
    #masterlistPanel .ml-dt .ml-col-email {{ min-width: 240px; max-width: 320px; }}
    #masterlistPanel .ml-dt .ml-clip {{ display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    #masterlistPanel .ml-pill {{ display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 700; letter-spacing: .04em; }}
    #masterlistPanel .ml-pa {{ background: #DCFCE7; color: #15803D; }}
    #masterlistPanel .ml-pi {{ background: #FEE2E2; color: #B91C1C; }}
    #masterlistPanel .ml-pp {{ background: #FEF9C3; color: #854D0E; }}

    /* Pager */
    #masterlistPanel .ml-tpager {{ display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 9px 15px; border-top: 1px solid var(--ml-border); flex-wrap: wrap; }}
    #masterlistPanel .ml-tpager-range {{ font-size: 11px; color: var(--muted); }}
    #masterlistPanel .ml-tpager-controls {{ display: flex; align-items: center; gap: 8px; }}
    #masterlistPanel .ml-tpager-page {{ font-size: 11px; color: var(--muted); font-variant-numeric: tabular-nums; min-width: 88px; text-align: center; }}
    #masterlistPanel .ml-btn {{
        padding: 5px 12px; border-radius: 4px; font-size: 12px; font-weight: 600; cursor: pointer; border: 1.5px solid var(--blue);
        background: var(--blue); color: #fff; font-family: Arial, sans-serif; transition: opacity .15s; text-decoration: none;
        display: inline-flex; align-items: center; justify-content: center; min-height: 32px;
    }}
    #masterlistPanel .ml-btn:hover {{ opacity: .85; }}
    #masterlistPanel .ml-btn.ml-sec {{ background: none; color: var(--blue); }}
    #masterlistPanel .ml-btn.ml-ico {{ padding: 5px 7px; color: var(--muted); background: none; border-color: var(--ml-border); }}
    #masterlistPanel .ml-btn.ml-ico:hover {{ color: var(--blue); border-color: var(--blue); }}
    #masterlistPanel .ml-btn.ml-ico[disabled] {{ opacity: .4; cursor: not-allowed; }}
    #masterlistPanel .ml-btn.ml-ico[disabled]:hover {{ color: var(--muted); border-color: var(--ml-border); }}

    /* Expand overlay + hover tooltip — top-level IDs, safe to leave unscoped (unique, no collisions) */
    #ml-ovl {{ position: fixed; inset: 0; background: rgba(8,20,45,.58); z-index: 9999; display: flex; align-items: center; justify-content: center; padding: 20px; opacity: 0; pointer-events: none; transition: opacity .2s ease; }}
    #ml-ovl.ml-open {{ opacity: 1; pointer-events: all; }}
    #ml-ovl .ml-xcard {{
        background: #fff; border-radius: 8px; box-shadow: 0 28px 70px rgba(0,0,0,.28); width: min(96vw, 1760px); max-width: none; height: 90vh; max-height: 92vh;
        display: flex; flex-direction: column; overflow: hidden; border-top: 3px solid var(--blue);
        transform: scale(.97) translateY(6px); transition: transform .2s ease;
    }}
    #ml-ovl.ml-open .ml-xcard {{ transform: scale(1) translateY(0); }}
    #ml-ovl .ml-xcard-hd {{ display: flex; align-items: center; justify-content: space-between; padding: 13px 16px; border-bottom: 1px solid #DDE6EE; }}
    #ml-ovl .ml-xcard-ttl {{ font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .07em; color: var(--blue); }}
    #ml-ovl .ml-xcard-sub {{ font-size: 11px; color: var(--muted); margin-top: 1px; }}
    #ml-ovl .ml-closebtn {{ background: none; border: none; cursor: pointer; color: var(--muted); padding: 4px 6px; border-radius: 4px; font-family: Arial, sans-serif; font-size: 18px; line-height: 1; transition: background .15s, color .15s; }}
    #ml-ovl .ml-closebtn:hover {{ background: var(--bg); color: var(--text); }}
    #ml-xbd {{ flex: 1; min-height: 0; overflow: auto; padding: 24px; display: flex; align-items: center; justify-content: center; }}
    #ml-xbd.ml-xbd-table {{ display: block; padding: 0; overflow: hidden; min-height: 0; }}
    /* Item 1 perf fix (modal): same rationale as the #masterlistPanel .ml-twrap
       rule — isolate this scroller's paint/composite layer so dragging the
       horizontal scrollbar over the full (unpaginated) row set doesn't force
       a repaint of the sticky header/frozen columns on every scroll tick. */
    #ml-xbd.ml-xbd-table .ml-twrap {{ height: 100%; max-height: none; width: 100%; overflow: auto; contain: paint; will-change: scroll-position; }}
    #ml-xbd.ml-xbd-table .ml-twrap table.ml-dt {{ width: max-content; min-width: 100%; table-layout: auto; border-collapse: separate; border-spacing: 0; font-size: 12px; font-variant-numeric: tabular-nums; }}
    #ml-xbd.ml-xbd-table .ml-dt thead tr {{ background: var(--blue); }}
    #ml-xbd.ml-xbd-table .ml-dt th,
    #ml-xbd.ml-xbd-table .ml-dt td {{ padding: 9px 14px; vertical-align: middle; line-height: 1.25; white-space: nowrap; border-bottom: 1px solid var(--ml-border); color: var(--text); }}
    #ml-xbd.ml-xbd-table .ml-dt th {{ position: sticky; top: 0; z-index: 5; padding: 9px 14px; text-align: left; font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: .06em; color: #fff; background: var(--blue); border-bottom: 1px solid #003B73; box-shadow: 0 1px 0 rgba(15,23,42,.18); white-space: nowrap; }}
    #ml-xbd.ml-xbd-table .ml-dt tbody tr:hover td {{ background: rgba(0,76,151,.025); }}
    /* Item 3: freeze columns through "Employee Name" in the expand modal too
       (same rationale as the #masterlistPanel block above). th here already
       has top:0 sticky — adding left:0/left:<offset> makes these two header
       cells sticky on BOTH axes (valid, standard "frozen corner" pattern), so
       z-index must clear the plain top-sticky header's z-index:5 above. The
       two frozen columns' .ml-clip spans are force-clipped (ellipsis) below,
       overriding the modal's normal no-clip behavior, so the fixed width
       holds regardless of content length — same tradeoff as the small table
       view, and full text is still available via the existing title tooltip. */
    #ml-xbd.ml-xbd-table .ml-dt th:nth-child(1), #ml-xbd.ml-xbd-table .ml-dt td:nth-child(1),
    #ml-xbd.ml-xbd-table .ml-dt th:nth-child(2), #ml-xbd.ml-xbd-table .ml-dt td:nth-child(2) {{
        position: sticky; box-sizing: border-box;
    }}
    #ml-xbd.ml-xbd-table .ml-dt th:nth-child(1), #ml-xbd.ml-xbd-table .ml-dt td:nth-child(1) {{
        left: 0; width: 128px; min-width: 128px; max-width: 128px;
    }}
    #ml-xbd.ml-xbd-table .ml-dt th:nth-child(2), #ml-xbd.ml-xbd-table .ml-dt td:nth-child(2) {{
        left: 128px; width: 248px; min-width: 248px; max-width: 248px;
        box-shadow: 2px 0 4px rgba(0,0,0,.06); /* freeze boundary divider */
    }}
    #ml-xbd.ml-xbd-table .ml-dt th:nth-child(1), #ml-xbd.ml-xbd-table .ml-dt th:nth-child(2) {{ z-index: 8; background: var(--blue); }}
    #ml-xbd.ml-xbd-table .ml-dt td:nth-child(1), #ml-xbd.ml-xbd-table .ml-dt td:nth-child(2) {{ z-index: 6; background: #fff; }}
    #ml-xbd.ml-xbd-table .ml-dt td:nth-child(1) .ml-clip, #ml-xbd.ml-xbd-table .ml-dt td:nth-child(2) .ml-clip {{
        overflow: hidden; text-overflow: ellipsis;
    }}
    /* Item 2 (modal pill support): #ml-ovl is a top-level overlay outside
       #masterlistPanel, so the .ml-pill/.ml-pa/.ml-pi/.ml-pp rules scoped to
       #masterlistPanel never reached it — Employment Status pills were
       already silently unstyled in the expand modal before this change. Add
       the same rules here (verbatim colors) so both Employment Status AND
       the new Tenure conditional-formatting pills render correctly. */
    #ml-ovl .ml-pill {{ display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 700; letter-spacing: .04em; }}
    #ml-ovl .ml-pa {{ background: #DCFCE7; color: #15803D; }}
    #ml-ovl .ml-pi {{ background: #FEE2E2; color: #B91C1C; }}
    #ml-ovl .ml-pp {{ background: #FEF9C3; color: #854D0E; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-id {{ min-width: 88px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-name {{ min-width: 220px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-date {{ min-width: 116px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-class {{ min-width: 156px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-tenure {{ min-width: 150px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-title {{ min-width: 260px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-group {{ min-width: 170px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-dept {{ min-width: 150px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-account {{ min-width: 180px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-supervisor,
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-manager {{ min-width: 220px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-status {{ min-width: 150px; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-col-email {{ min-width: 310px; max-width: none; }}
    #ml-xbd.ml-xbd-table .ml-dt .ml-clip {{ display: block; overflow: visible; text-overflow: clip; white-space: nowrap; }}
    #ml-xbd.ml-xbd-scroll {{ align-items: flex-start; }}
    #ml-cv-tt {{ position: fixed; pointer-events: none; z-index: 10000; background: #0F2240; color: #fff; font-size: 11px; font-family: Arial, sans-serif; padding: 5px 9px; border-radius: 5px; box-shadow: 0 4px 14px rgba(0,0,0,.25); white-space: nowrap; opacity: 0; transition: opacity .1s; top: 0; left: 0; }}
    #ml-cv-tt.ml-show {{ opacity: 1; }}

    @media (prefers-reduced-motion: reduce) {{
        #masterlistPanel .ml-card, #masterlistPanel .ml-xbtn, #ml-ovl, #ml-ovl .ml-xcard {{ transition: none; }}
    }}

    @media (max-width: 1300px) {{
        #masterlistPanel .ml-r5, #masterlistPanel .ml-r4, #masterlistPanel .ml-r2asym {{ grid-template-columns: 1fr; }}
        #masterlistControls .ml-kpi-strip {{ grid-template-columns: repeat(2, 1fr); position: static; }}
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
            <p style="margin:2px 0 0 0"><span id="pb-data-freshness" style="font-size:11px;font-weight:600"></span></p>
        </div>
    </div>
</div>

<div class="tabs" role="tablist" aria-label="Dashboard sections">
    <button class="tab-button active" type="button" data-tab="masterlist" role="tab" aria-selected="true">Masterlist</button>
    <button class="tab-button" type="button" data-tab="coaching" role="tab" aria-selected="false">Coaching</button>
    <button class="tab-button" type="button" data-tab="quality" role="tab" aria-selected="false">Quality</button>
    <button class="tab-button" type="button" data-tab="schedule" role="tab" aria-selected="false">Schedule</button>
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
    <div class="filter-box">
        <label>Employment Class</label>
        <details class="multi-filter" id="employmentClassFilter">
            <summary id="employmentClassFilterSummary">All</summary>
            <div class="multi-options" id="employmentClassOptions"></div>
        </details>
    </div>
    <button class="ml-btn ml-sec ml-clear-btn" id="mlClearFiltersBtn" type="button">Clear filters</button>
</div>

<div class="ml-kpi-strip" id="masterlistKpiStrip">
    <div class="ml-kpi g">
        <div class="ml-kpi-lbl">Active Employees</div>
        <div class="ml-kpi-val" id="ml-kpi-active">0</div>
        <div class="ml-kpi-sub">Currently active</div>
    </div>
    <div class="ml-kpi b">
        <div class="ml-kpi-lbl">Total Headcount</div>
        <div class="ml-kpi-val" id="ml-kpi-total">0</div>
        <div class="ml-kpi-sub" id="ml-kpi-total-sub">Incl. <span style="color:#B91C1C;font-weight:700;font-size:13px">0</span> inactive</div>
    </div>
    <div class="ml-kpi">
        <div class="ml-kpi-lbl">Departments</div>
        <div class="ml-kpi-val" id="ml-kpi-departments">0</div>
        <div class="ml-kpi-sub">Across filtered results</div>
    </div>
    <div class="ml-kpi">
        <div class="ml-kpi-lbl">Accounts</div>
        <div class="ml-kpi-val" id="ml-kpi-accounts">22</div>
        <div class="ml-kpi-sub">Active accounts</div>
    </div>
    <div class="ml-kpi">
        <div class="ml-kpi-lbl">Avg Tenure</div>
        <div class="ml-kpi-val" id="ml-kpi-avgtenure">0<span style="font-size:14px;font-weight:400"> yrs</span></div>
        <div class="ml-kpi-sub">Org-wide, active staff</div>
    </div>
    <div class="ml-kpi">
        <div class="ml-kpi-lbl">Movements</div>
        <div class="ml-kpi-val" id="ml-kpi-movements">0</div>
        <div class="ml-kpi-sub" id="ml-kpi-movements-sub">0 for processing</div>
    </div>
    <div class="ml-kpi">
        <div class="ml-kpi-lbl">History Records</div>
        <div class="ml-kpi-val" id="ml-kpi-history">0</div>
        <div class="ml-kpi-sub">Snapshot log entries</div>
    </div>
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
<div class="grid ml-grid">

  <div class="ml-row ml-r5">
    <div class="ml-card" data-cid="active" data-ttl="Active Status" data-sub="Active vs inactive employees">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">Active Status</div><div class="ml-card-sub">Active vs inactive employees</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand Active Status" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd"><canvas id="c-ml-active"></canvas></div>
    </div>
    <div class="ml-card" data-cid="dept" data-ttl="By Department" data-sub="Headcount distribution">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">By Department</div><div class="ml-card-sub">Headcount distribution</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand By Department" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd"><canvas id="c-ml-dept"></canvas></div>
    </div>
    <div class="ml-card" data-cid="empgrp" data-ttl="Employee Group" data-sub="Classification breakdown">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">Employee Group</div><div class="ml-card-sub">Classification breakdown</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand Employee Group" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd"><canvas id="c-ml-empgrp"></canvas></div>
    </div>
    <div class="ml-card" data-cid="empclass" data-ttl="Employment Class" data-sub="Full-time · Part-time · Casual">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">Employment Class</div><div class="ml-card-sub">Full-time · Part-time · Casual</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand Employment Class" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd"><canvas id="c-ml-empclass"></canvas></div>
    </div>
    <div class="ml-card" data-cid="tenureseg" data-ttl="Tenure Segmentation" data-sub="Headcount by tenure band">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">Tenure Segmentation</div><div class="ml-card-sub">Headcount by tenure band</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand Tenure Segmentation" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd"><canvas id="c-ml-tenureseg"></canvas></div>
    </div>
  </div>

  <div class="ml-row ml-r4">
    <div class="ml-card" data-cid="account" data-ttl="By Account" data-sub="Headcount per account">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">By Account</div><div class="ml-card-sub">Headcount per account</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand By Account" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd ml-hbar-bd"><div class="ml-hbar-scroll"><canvas id="c-ml-account"></canvas></div></div>
    </div>
    <div class="ml-card" data-cid="manager" data-ttl="By Manager" data-sub="Direct reports per manager">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">By Manager</div><div class="ml-card-sub">Direct reports per manager</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand By Manager" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd ml-hbar-bd"><div class="ml-hbar-scroll"><canvas id="c-ml-manager"></canvas></div></div>
    </div>
    <div class="ml-card" data-cid="supervisor" data-ttl="By Supervisor" data-sub="Span of control">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">By Supervisor</div><div class="ml-card-sub">Span of control</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand By Supervisor" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd ml-hbar-bd"><div class="ml-hbar-scroll"><canvas id="c-ml-supervisor"></canvas></div></div>
    </div>
    <div class="ml-card" data-cid="age" data-ttl="Age Group" data-sub="Workforce demographics">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">Age Group</div><div class="ml-card-sub">Workforce demographics</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand Age Group" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd ml-hbar-bd"><div class="ml-hbar-scroll"><canvas id="c-ml-age"></canvas></div></div>
    </div>
  </div>

  <div class="ml-row ml-r2asym">
    <div class="ml-card" data-cid="tenurestack" data-ttl="Tenure by Account" data-sub="Stacked tenure bands per account · scroll for all">
      <div class="ml-card-hd"><div><div class="ml-card-ttl">Tenure by Account</div><div class="ml-card-sub">Stacked tenure bands per account · scroll for all</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand Tenure by Account" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd ml-hscroll-bd"><div class="ml-hscroll-x"><canvas id="c-ml-tenurestack"></canvas></div></div>
    </div>
    <div class="ml-card gc" data-cid="weekly" data-ttl="Weekly Headcount Trend" data-sub="Headcount by week, by Employment Class">
      <div class="ml-card-hd"><div><div class="ml-card-ttl" style="color:#1a7a2e">Weekly Headcount Trend</div><div class="ml-card-sub">Headcount by week, by Employment Class</div></div>
      <button class="ml-xbtn" type="button" aria-label="Expand Weekly Trend" onclick="mlOpenExpand(this.closest('.ml-card'), this)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button></div>
      <div class="ml-card-bd ml-hscroll-bd"><div class="ml-hscroll-x"><canvas id="c-ml-weekly"></canvas></div></div>
    </div>
  </div>

  <div class="ml-tcrd" data-cid="masterlist" data-ttl="Master List" data-sub="Full employee roster" id="masterlistCard">
    <div class="ml-thd">
      <div class="ml-ttl">Master List</div>
      <div class="ml-tacts">
        <a class="ml-btn" href="{masterlist_source_url}" target="_blank" rel="noopener">{masterlist_source_label}</a>
        <a class="ml-btn ml-sec" href="{masterlist_excel_url}" target="_blank" rel="noopener">Download Excel</a>
        <button class="ml-btn ml-sec ml-ico" type="button" id="masterlistFocusToggle" aria-label="Expand Master List" title="Expand Master List">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>
        </button>
      </div>
    </div>
    <div class="ml-twrap">
      <table class="ml-dt" id="ml-table">
        <thead id="ml-thead"></thead>
        <tbody id="ml-tbody"></tbody>
      </table>
    </div>
    <div class="ml-tpager">
      <span class="ml-tpager-range" id="ml-pager-range">Showing 0 of 0 employees</span>
      <div class="ml-tpager-controls">
        <button class="ml-btn ml-ico" type="button" id="ml-prev" aria-label="Previous page"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 18 9 12 15 6"/></svg></button>
        <span class="ml-tpager-page" id="ml-pager-page">Page 1 of 1</span>
        <button class="ml-btn ml-ico" type="button" id="ml-next" aria-label="Next page"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="9 18 15 12 9 6"/></svg></button>
      </div>
    </div>
  </div>

  <div class="chart-card full">
    <h3 style="color:#004C97;margin:4px 0 12px;">Recent Employee Movements</h3>
    <div id="recentMovements"></div>
  </div>

</div>
</div>

<!-- Masterlist expand modal + hover tooltip (top-level; ml-open/ml-ovl/ml-xbd IDs are unique dashboard-wide) -->
<div id="ml-ovl" role="dialog" aria-modal="true" aria-labelledby="ml-xttl">
  <div class="ml-xcard">
    <div class="ml-xcard-hd">
      <div><div class="ml-xcard-ttl" id="ml-xttl"></div><div class="ml-xcard-sub" id="ml-xsub"></div></div>
      <button class="ml-closebtn" id="ml-xcls" type="button" aria-label="Close">&#10005;</button>
    </div>
    <div class="ml-xcard-bd" id="ml-xbd"></div>
  </div>
</div>
<div id="ml-cv-tt" aria-hidden="true"></div>

<div class="tab-panel" id="coachingPanel" data-tab="coaching" role="tabpanel">
<div class="grid coaching-grid">
    <div class="coaching-chart-row" id="coaching-chart-row">
        <div class="chart-card" id="coaching-cat-card">
            <button class="coaching-expand-btn" onclick="toggleCoachingCardExpand('coaching-cat-card')" title="Expand"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button>
            <div id="coachingCategoryDonut" class="coaching-donut-host"></div>
        </div>
        <div class="chart-card" id="coaching-status-card">
            <button class="coaching-expand-btn" onclick="toggleCoachingCardExpand('coaching-status-card')" title="Expand"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button>
            <div id="coachingStatusDonut" class="coaching-donut-host"></div>
        </div>
        <div class="chart-card" id="coaching-cov-card">
            <button class="coaching-expand-btn" onclick="toggleCoachingCardExpand('coaching-cov-card')" title="Expand"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg></button>
            <div id="coachingCoverageDonut" class="coaching-donut-host"></div>
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
    <label>Evaluation Date Range</label>
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
      <option value="ac">Associated Cab</option>
      <option value="bl">Blueline</option>
      <option value="britelift">Britelift</option>
      <option value="blc">Britelift Chat</option>
      <option value="ch">C&amp;H</option>
      <option value="ct">Circle Taxi</option>
      <option value="dc">Data Carz</option>
      <option value="dmg">DMG</option>
      <option value="hamilton">Hamilton</option>
      <option value="kel">Kelowna</option>
      <option value="m7">M7 &ndash; Ride-hailing support</option>
      <option value="ol">Ollies</option>
      <option value="parentis">Parentis Health</option>
      <option value="r4h">R4H</option>
      <option value="rc">Reno Cab</option>
      <option value="ridex">RideX</option>
      <option value="skyline">Skyline</option>
      <option value="ti">Trans Iowa</option>
      <option value="vt">Vermont</option>
      <option value="vip">VIP</option>
      <option value="ycdc">YCDC</option>
      <option value="ycov">YCOV</option>
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
  <!-- VIP-only key criteria strip (shown when VIP account selected) -->
  <div id="qa-vip-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #B45309"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-vcrit-greet-val" style="color:#B45309">&mdash;</div><div class="qa-sum-sub" id="qa-vcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#B45309">VIP criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #D97706"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-vcrit-prof-val" style="color:#D97706">&mdash;</div><div class="qa-sum-sub" id="qa-vcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#D97706">VIP criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-vcrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-vcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">VIP criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-vcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-vcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">VIP criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-vcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-vcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">VIP criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-vcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-vcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">VIP criterion</div></div>
    </div>
  </div>
  <!-- C&H-only key criteria strip (shown when C&H account selected) -->
  <div id="qa-ch-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-chcrit-greet-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-chcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">C&amp;H criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-chcrit-prof-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-chcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">C&amp;H criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0369A1"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-chcrit-genq-val" style="color:#0369A1">&mdash;</div><div class="qa-sum-sub" id="qa-chcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0369A1">C&amp;H criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-chcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-chcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">C&amp;H criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-chcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-chcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">C&amp;H criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-chcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-chcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">C&amp;H criterion</div></div>
    </div>
  </div>
  <!-- Reno Cab-only key criteria strip (shown when Reno Cab account selected) -->
  <div id="qa-rc-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #15803D"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-rccrit-greet-val" style="color:#15803D">&mdash;</div><div class="qa-sum-sub" id="qa-rccrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#15803D">RC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #059669"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-rccrit-prof-val" style="color:#059669">&mdash;</div><div class="qa-sum-sub" id="qa-rccrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#059669">RC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-rccrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-rccrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">RC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-rccrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-rccrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">RC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-rccrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-rccrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">RC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-rccrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-rccrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">RC criterion</div></div>
    </div>
  </div>
  <!-- Data Carz-only key criteria strip (shown when Data Carz account selected) -->
  <div id="qa-dc-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #C2410C"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-dccrit-greet-val" style="color:#C2410C">&mdash;</div><div class="qa-sum-sub" id="qa-dccrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#C2410C">DC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #EA580C"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-dccrit-prof-val" style="color:#EA580C">&mdash;</div><div class="qa-sum-sub" id="qa-dccrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#EA580C">DC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-dccrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-dccrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">DC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-dccrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-dccrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">DC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-dccrit-resol-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-dccrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">DC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #065F46"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-dccrit-comm-val" style="color:#065F46">&mdash;</div><div class="qa-sum-sub" id="qa-dccrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#065F46">DC criterion</div></div>
    </div>
  </div>
  <!-- Trans Iowa-only key criteria strip (shown when Trans Iowa account selected) -->
  <div id="qa-ti-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #4338CA"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-ticrit-greet-val" style="color:#4338CA">&mdash;</div><div class="qa-sum-sub" id="qa-ticrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#4338CA">TI criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-ticrit-prof-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-ticrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">TI criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-ticrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-ticrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">TI criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-ticrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-ticrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">TI criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #7C3AED"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-ticrit-resol-val" style="color:#7C3AED">&mdash;</div><div class="qa-sum-sub" id="qa-ticrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#7C3AED">TI criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-ticrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-ticrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">TI criterion</div></div>
    </div>
  </div>
  <!-- Associated Cab-only key criteria strip (shown when Associated Cab account selected) -->
  <div id="qa-ac-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #0E7490"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-accrit-greet-val" style="color:#0E7490">&mdash;</div><div class="qa-sum-sub" id="qa-accrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#0E7490">AC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-accrit-prof-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-accrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">AC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-accrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-accrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">AC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-accrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-accrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">AC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #7C3AED"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-accrit-resol-val" style="color:#7C3AED">&mdash;</div><div class="qa-sum-sub" id="qa-accrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#7C3AED">AC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-accrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-accrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">AC criterion</div></div>
    </div>
  </div>
  <!-- Ollies-only key criteria strip (shown when Ollies account selected) -->
  <div id="qa-ol-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #BE123C"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-olcrit-greet-val" style="color:#BE123C">&mdash;</div><div class="qa-sum-sub" id="qa-olcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#BE123C">OL criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #E11D48"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-olcrit-prof-val" style="color:#E11D48">&mdash;</div><div class="qa-sum-sub" id="qa-olcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#E11D48">OL criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-olcrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-olcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">OL criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-olcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-olcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">OL criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #7C3AED"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-olcrit-resol-val" style="color:#7C3AED">&mdash;</div><div class="qa-sum-sub" id="qa-olcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#7C3AED">OL criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-olcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-olcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">OL criterion</div></div>
    </div>
  </div>
  <!-- Circle Taxi-only key criteria strip (shown when Circle Taxi account selected) -->
  <div id="qa-ct-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-ctcrit-greet-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-ctcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">CT criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-ctcrit-prof-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-ctcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">CT criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0369A1"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-ctcrit-genq-val" style="color:#0369A1">&mdash;</div><div class="qa-sum-sub" id="qa-ctcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0369A1">CT criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-ctcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-ctcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">CT criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-ctcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-ctcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">CT criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-ctcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-ctcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">CT criterion</div></div>
    </div>
  </div>
  <!-- YCOV-only key criteria strip (shown when YCOV account selected) -->
  <div id="qa-ycov-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #047857"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-ycovcrit-greet-val" style="color:#047857">&mdash;</div><div class="qa-sum-sub" id="qa-ycovcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#047857">YCOV criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-ycovcrit-prof-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-ycovcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">YCOV criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-ycovcrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-ycovcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">YCOV criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-ycovcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-ycovcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">YCOV criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-ycovcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-ycovcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">YCOV criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-ycovcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-ycovcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">YCOV criterion</div></div>
    </div>
  </div>
  <!-- Kelowna-only key criteria strip (shown when Kelowna account selected) -->
  <div id="qa-kel-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #166534"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-kelcrit-greet-val" style="color:#166534">&mdash;</div><div class="qa-sum-sub" id="qa-kelcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#166534">Kelowna criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #15803D"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-kelcrit-prof-val" style="color:#15803D">&mdash;</div><div class="qa-sum-sub" id="qa-kelcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#15803D">Kelowna criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-kelcrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-kelcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">Kelowna criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-kelcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-kelcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">Kelowna criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-kelcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-kelcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">Kelowna criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-kelcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-kelcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">Kelowna criterion</div></div>
    </div>
  </div>
  <!-- Vermont-only key criteria strip (shown when Vermont account selected) -->
  <div id="qa-vt-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-vtcrit-greet-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-vtcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">Vermont criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #047857"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-vtcrit-prof-val" style="color:#047857">&mdash;</div><div class="qa-sum-sub" id="qa-vtcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#047857">Vermont criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-vtcrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-vtcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">Vermont criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-vtcrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-vtcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">Vermont criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-vtcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-vtcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">Vermont criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-vtcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-vtcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">Vermont criterion</div></div>
    </div>
  </div>
  <!-- YCDC-only key criteria strip (shown when YCDC account selected) -->
  <div id="qa-ycdc-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #155E75"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-ycdccrit-greet-val" style="color:#155E75">&mdash;</div><div class="qa-sum-sub" id="qa-ycdccrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#155E75">YCDC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-ycdccrit-prof-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-ycdccrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">YCDC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-ycdccrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-ycdccrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">YCDC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-ycdccrit-verif-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-ycdccrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">YCDC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-ycdccrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-ycdccrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">YCDC criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-ycdccrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-ycdccrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">YCDC criterion</div></div>
    </div>
  </div>
  <!-- Blueline-only key criteria strip (shown when Blueline account selected) -->
  <div id="qa-bl-crit" style="display:none;padding:0 20px 8px">
    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:8px">
      <div class="qa-sum-card" style="border-top:3px solid #1D4ED8"><div class="qa-sum-lbl">Greetings Script</div><div class="qa-sum-val" id="qa-blcrit-greet-val" style="color:#1D4ED8">&mdash;</div><div class="qa-sum-sub" id="qa-blcrit-greet-sub">&mdash;</div><div class="qa-sum-score" style="color:#1D4ED8">Blueline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #2563EB"><div class="qa-sum-lbl">Professionalism</div><div class="qa-sum-val" id="qa-blcrit-prof-val" style="color:#2563EB">&mdash;</div><div class="qa-sum-sub" id="qa-blcrit-prof-sub">&mdash;</div><div class="qa-sum-score" style="color:#2563EB">Blueline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0891B2"><div class="qa-sum-lbl">General Questions</div><div class="qa-sum-val" id="qa-blcrit-genq-val" style="color:#0891B2">&mdash;</div><div class="qa-sum-sub" id="qa-blcrit-genq-sub">&mdash;</div><div class="qa-sum-score" style="color:#0891B2">Blueline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #0F766E"><div class="qa-sum-lbl">Cust. Verification</div><div class="qa-sum-val" id="qa-blcrit-verif-val" style="color:#0F766E">&mdash;</div><div class="qa-sum-sub" id="qa-blcrit-verif-sub">&mdash;</div><div class="qa-sum-score" style="color:#0F766E">Blueline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #6D28D9"><div class="qa-sum-lbl">Res. Etiquette</div><div class="qa-sum-val" id="qa-blcrit-resol-val" style="color:#6D28D9">&mdash;</div><div class="qa-sum-sub" id="qa-blcrit-resol-sub">&mdash;</div><div class="qa-sum-score" style="color:#6D28D9">Blueline criterion</div></div>
      <div class="qa-sum-card" style="border-top:3px solid #9D174D"><div class="qa-sum-lbl">Comm. Quality</div><div class="qa-sum-val" id="qa-blcrit-comm-val" style="color:#9D174D">&mdash;</div><div class="qa-sum-sub" id="qa-blcrit-comm-sub">&mdash;</div><div class="qa-sum-score" style="color:#9D174D">Blueline criterion</div></div>
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
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22,12 18,12 15,21 9,3 6,12 2,12"/></svg>Overall QA Score Trend</div><div class="qa-cs" id="qa-trend-sub">Weekly avg (Mon&ndash;Sun) &middot; Target: 85%</div></div><span class="qa-cb qa-cbg" id="qa-trend-badge">&mdash;</span></div>
      <div class="qa-cbody" style="padding:6px 10px;display:flex;flex-direction:column"><div style="position:relative;flex:1;min-height:90px"><canvas id="qa-trend-chart"></canvas></div></div>
    </div>
    <div class="qa-card" id="qa-aiqe-card">
      <div class="qa-ch">
        <div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M7 15l4-4 3 3 5-7"/></svg>AI x QE Score Trend</div><div class="qa-cs" id="qa-aiqe-trend-sub">Date range &amp; account only</div></div>
        <div style="display:flex;align-items:center;gap:6px;flex-shrink:0">
          <span class="qa-cb qa-cbb" id="qa-aiqe-trend-badge">&mdash;</span>
          <button id="qa-aiqe-focus-toggle" type="button" title="Expand AI x QE trend" aria-label="Expand AI x QE trend" style="width:28px;height:28px;border:1px solid #CBD5E1;border-radius:5px;background:#fff;color:#475569;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0">
            <span class="qa-focus-icon" aria-hidden="true"></span>
          </button>
        </div>
      </div>
      <div class="qa-cbody" style="padding:6px 10px;display:flex;flex-direction:column"><div id="qa-aiqe-chart-host" style="position:relative;flex:1;min-height:90px"><canvas id="qa-aiqe-trend-chart"></canvas></div></div>
    </div>
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18M15 3v18M3 9h18M3 15h18"/></svg>Criteria pass rates</div><div class="qa-cs" id="qa-crit-sub">All 23 criteria &middot; sorted by pass rate</div></div><span class="qa-cb qa-cba">Score breakdown</span></div>
      <div style="padding:6px 10px;flex:1;display:flex;flex-direction:column;min-height:0">
        <div style="overflow-y:auto;flex:1 1 auto;min-height:0;padding-right:2px;border-bottom:1px solid #E2E8F0" id="qa-criteria-bars"></div>
        <div style="padding:3px 0 4px;font-size:10px;color:#CBD5E1">&#9679; Scroll to see all 23 criteria</div>
      </div>
    </div>
    <div class="qa-card">
      <div class="qa-ch"><div><div class="qa-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4z"/></svg>Coaching opportunities</div><div class="qa-cs">Criteria below 95% pass rate</div></div><span class="qa-cb qa-cbr" id="qa-coaching-count">&mdash;</span></div>
      <div class="qa-cbody" style="padding:6px 10px;display:flex;flex-direction:column;min-height:0">
        <div id="qa-coaching-bars" style="overflow-y:auto;flex:1 1 auto;min-height:0;padding-right:2px"></div>
        <div style="margin-top:auto;padding-top:6px">
          <div style="height:1px;background:#F1F5F9;margin:0 0 6px"></div>
          <div style="font-size:10px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.04em;margin-bottom:5px">QA coach breakdown</div>
          <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(62px,1fr));gap:8px;align-items:stretch" id="qa-coach-breakdown"></div>
        </div>
      </div>
    </div>
    <div class="qa-card" id="qa-dist-card">
      <div class="qa-ch">
        <div><div class="qa-ct" id="qa-dist-title">Score distribution</div><div class="qa-cs" id="qa-donut-sub">All evaluations</div></div>
        <button id="qa-dist-focus-toggle" onclick="toggleQADistFocusMode()" title="Expand distribution" aria-label="Expand distribution" style="width:28px;height:28px;border:1px solid #CBD5E1;border-radius:5px;background:#fff;color:#475569;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0">
          <span class="qa-focus-icon" aria-hidden="true"></span>
        </button>
      </div>
      <div class="qa-cbody" style="padding:6px 10px;display:flex;flex-direction:column;justify-content:center">
        <div id="qa-score-dist-wrap">
          <div id="qa-score-chart-host" style="position:relative;height:130px;flex-shrink:0"><canvas id="qa-donut-chart"></canvas></div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:2px;margin-top:5px" id="qa-donut-legend"></div>
        </div>
        <div id="qa-eval-dist-wrap" style="display:none">
          <div id="qa-eval-chart-host" style="position:relative;height:150px;flex-shrink:0">
            <canvas id="qa-eval-dist-chart"></canvas>
            <div id="qa-eval-dist-other-note">Disclaimer: Accounts contributing less than 1% of total evaluations are grouped under Other for readability.</div>
          </div>
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
          <th style="min-width:120px">Evaluation ID</th>
          <th style="min-width:130px;cursor:pointer;user-select:none" data-qa-sort="ts" onclick="qaSortTable('ts')">Evaluation Date <span class="sort-indicator">▼</span></th>
          <th style="min-width:200px;cursor:pointer;user-select:none" data-qa-sort="agent" onclick="qaSortTable('agent')">Emp Name <span class="sort-indicator"></span></th>
          <th style="min-width:160px;cursor:pointer;user-select:none" data-qa-sort="supervisor" onclick="qaSortTable('supervisor')">Immediate Head <span class="sort-indicator"></span></th>
          <th style="min-width:160px;cursor:pointer;user-select:none" data-qa-sort="coach" onclick="qaSortTable('coach')">QA Coach <span class="sort-indicator"></span></th>
          <th style="min-width:125px">Account</th>
          <th style="min-width:55px;cursor:pointer;user-select:none" data-qa-sort="score" onclick="qaSortTable('score')">Score <span class="sort-indicator"></span></th>
          <th style="min-width:75px">Opening In</th><th style="min-width:75px">Opening Out</th>
          <th style="min-width:65px">Closing</th><th style="min-width:75px">Approp. Resp</th>
          <th style="min-width:65px">No Resp</th><th style="min-width:60px">Fillers</th>
          <th style="min-width:80px">Acknowledge</th><th style="min-width:70px">Hold Proc.</th>
          <th style="min-width:70px">Ack. Hold</th><th style="min-width:70px">Resp. Eff.</th>
          <th style="min-width:60px">Empathy</th><th style="min-width:60px">Adjust</th>
          <th style="min-width:55px">Mute</th><th style="min-width:70px">Active List.</th>
          <th style="min-width:65px">General Q</th><th style="min-width:70px">Answered Q</th><th style="min-width:65px">Probing Q</th>
          <th style="min-width:70px">Cust. Verif.</th><th style="min-width:65px">Clarif.</th>
          <th style="min-width:65px">Lost SOP</th><th style="min-width:65px">Rudeness</th>
          <th style="min-width:75px">Transaction</th><th style="min-width:65px">Speech</th>
          <th style="min-width:100px">Information Precision</th>
          <th class="vip-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="vip-extra-col" style="min-width:80px">Professionalism</th>
          <th class="vip-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="vip-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="vip-extra-col" style="min-width:80px">CV Other</th>
          <th class="vip-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="vip-extra-col" style="min-width:65px">VIP SOP</th>
          <th class="ch-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="ch-extra-col" style="min-width:80px">Professionalism</th>
          <th class="ch-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="ch-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="ch-extra-col" style="min-width:90px">CV Measures</th>
          <th class="ch-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="ch-extra-col" style="min-width:75px">Subjective</th>
          <th class="rc-extra-col" style="min-width:80px">Professionalism</th>
          <th class="rc-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="rc-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="rc-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="rc-extra-col" style="min-width:90px">CV Measures</th>
          <th class="ti-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="ti-extra-col" style="min-width:80px">Professionalism</th>
          <th class="ti-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="ti-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="ti-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="dc-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="dc-extra-col" style="min-width:80px">Professionalism</th>
          <th class="dc-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="dc-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="ac-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="ac-extra-col" style="min-width:80px">Professionalism</th>
          <th class="ac-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="ac-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="ac-extra-col" style="min-width:90px">CV Measures</th>
          <th class="ac-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="ol-extra-col" style="min-width:80px">Promptness</th>
          <th class="ol-extra-col" style="min-width:80px">Professionalism</th>
          <th class="ol-extra-col" style="min-width:80px">Dead Air</th>
          <th class="ol-extra-col" style="min-width:80px">No Vern.</th>
          <th class="ol-extra-col" style="min-width:80px">Avoid Int.</th>
          <th class="ol-extra-col" style="min-width:90px">CV Measures</th>
          <th class="ol-extra-col" style="min-width:80px">Ride Cancel</th>
          <th class="ol-extra-col" style="min-width:80px">Timeliness</th>
          <th class="ol-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="ol-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="ct-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="ct-extra-col" style="min-width:80px">Professionalism</th>
          <th class="ct-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="ct-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="ct-extra-col" style="min-width:90px">CV Measures</th>
          <th class="ct-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="ct-extra-col" style="min-width:75px">Subjective</th>
          <th class="ycov-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="ycov-extra-col" style="min-width:80px">Professionalism</th>
          <th class="ycov-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="ycov-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="ycov-extra-col" style="min-width:90px">CV Measures</th>
          <th class="ycov-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="ycov-extra-col" style="min-width:75px">Subjective</th>
          <th class="kel-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="kel-extra-col" style="min-width:80px">Professionalism</th>
          <th class="kel-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="kel-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="kel-extra-col" style="min-width:90px">CV Measures</th>
          <th class="kel-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="vt-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="vt-extra-col" style="min-width:80px">Professionalism</th>
          <th class="vt-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="vt-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="vt-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="ycdc-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="ycdc-extra-col" style="min-width:80px">Professionalism</th>
          <th class="ycdc-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="ycdc-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="ycdc-extra-col" style="min-width:80px">Res. Etiquette</th>
          <th class="bl-extra-col" style="min-width:75px">Opening Out 2</th>
          <th class="bl-extra-col" style="min-width:80px">Professionalism</th>
          <th class="bl-extra-col" style="min-width:70px">Verif. Other</th>
          <th class="bl-extra-col" style="min-width:80px">Comm. Quality</th>
          <th class="bl-extra-col" style="min-width:80px">Res. Etiquette</th>
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

<div class="tab-panel" id="schedulePanel" data-tab="schedule" role="tabpanel">
    <iframe
        src="scheduler-arch.html"
        style="width:100%;height:calc(100vh - 120px);border:none;display:block;"
        title="Scheduler System Architecture"
        loading="lazy">
    </iframe>
</div>

<div class="footer">
    Developed for Pac-Biz Reporting MCerna | Data Source: Master List | Automation: Python 3.13.0
</div>

<script>
const masterlist = {to_records(masterlist)};
const masterlistKpis = {json.dumps(masterlist_kpis, default=str)};
const historyData = {to_records(history)};
const movementData = {to_records(movement)};
const coachingData = {to_records(coaching)};
const qaRawData = {to_records(m7)};
const dmgRawData = {to_records(dmg)};
const r4hRawData = {to_records(r4h)};
const parentisRawData = {to_records(parentis)};
const briteliftRawData = {to_records(britelift)};
const blcRawData = {to_records(blc)};
const ridexRawData = {to_records(ridex)};
const hamiltonRawData = {to_records(hamilton)};
const skylineRawData = {to_records(skyline)};
const vipRawData = {to_records(vip)};
const chRawData = {to_records(ch)};
const rcRawData = {to_records(rc)};
const tiRawData = {to_records(ti)};
const dcRawData = {to_records(dc)};
const acRawData = {to_records(ac)};
const olRawData = {to_records(ol)};
const ctRawData = {to_records(ct)};
const ycovRawData = {to_records(ycov)};
const kelRawData = {to_records(kel)};
const vtRawData = {to_records(vt)};
const ycdcRawData = {to_records(ycdc)};
const blRawData = {to_records(bl)};

document.addEventListener("click", (event) => {{
    const btn = event.target?.closest?.("#qa-aiqe-focus-toggle");
    if (!btn) return;
    event.preventDefault();
    event.stopPropagation();
    if (typeof setQAAiQeFocusMode === "function") {{
        setQAAiQeFocusMode(!document.body.classList.contains("qa-aiqe-focus-mode"));
    }}
}}, true);

const PB_BUILD_TS = new Date("{refresh_iso}");
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
    {{label: "Initiated by", field: "Initiated by", sortable: true}},
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
    employmentClass: new Set(),
}};
const APPROVED_ACCOUNTS = new Set([
    "Alpha Tax", "Associate", "Associated Cab", "Brite Lift", "Buffalo", "C&H",
    "Circle Taxi", "Data Carz", "DMG", "Hamilton", "Kaizen", "Kelowna", "Vermont", "YCDC",
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

// Single source of truth for the calendar-based (day-of-month agnostic)
// years/months breakdown used by the displayed tenure label. asOfDate is
// optional so callers that need to reuse one "now" across multiple
// computations for the same row (see mlPrepareRows) can pass it in.
function tenureBreakdown(hireDate, asOfDate) {{
    const now = asOfDate || new Date();
    const days = wholeDayDiff(hireDate, now);
    if (days < 0) return null;
    let years = now.getFullYear() - hireDate.getFullYear();
    let months = now.getMonth() - hireDate.getMonth();
    if (months < 0) {{ years--; months += 12; }}
    return {{ days, years, months, totalMonths: years * 12 + months }};
}}
function formatTenureDisplay(hireDate, asOfDate) {{
    if (!hireDate) return "";
    const b = tenureBreakdown(hireDate, asOfDate);
    if (!b) return "";
    const {{ days, years, months }} = b;
    if (days < 31) return `${{days}} Day${{days !== 1 ? "s" : ""}}`;
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

function coachingSummaryDate(row) {{
    return norm(row["Coaching Date"]) || norm(row["Created Date"]);
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
    populateMultiFilter(
        "employmentClassOptions",
        "employmentClassFilterSummary",
        uniqueValues(masterlist, "Employement Class"),
        MASTERLIST_FILTERS.employmentClass,
        render
    );
}}

// Item 6 — Clear filters (POC parity: masterlist-poc.html's #f-clear resets
// FILTER_STATE to its defaults, resets each control, then re-runs drawAll()).
// Here "default" for every MASTERLIST_FILTERS Set is empty == "all selected"
// (see populateMultiFilter/syncMultiFilterOptions), so clearing every Set and
// rebuilding the dropdowns puts every filter back to "All" with no search —
// then render() re-draws charts/table/KPIs off the now-unfiltered data.
function mlClearFilters() {{
    Object.values(MASTERLIST_FILTERS).forEach(set => set.clear());
    closeMultiFilters();
    populateMasterlistFilters();
    render();
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
            filterMatches(MASTERLIST_FILTERS.employmentStatus, norm(r["Employment Status"])) &&
            filterMatches(MASTERLIST_FILTERS.employmentClass, norm(r["Employement Class"]))
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
    const weekKeys = [...new Set(completedRows.map(r => weekStartKey(coachingSummaryDate(r))).filter(Boolean))]
        .sort();
    const leaders = [...new Set(completedRows.map(r => norm(r["Coached by"]) || "Blank"))]
        .sort((a, b) => a.localeCompare(b, undefined, {{sensitivity: "base"}}));
    const counts = {{}};

    completedRows.forEach(r => {{
        const leader = norm(r["Coached by"]) || "Blank";
        const week = weekStartKey(coachingSummaryDate(r));
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
    const colors = metrics.totalDR !== null ? ["#00A651", "#DC2626"] : ["#00A651"];
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
    const rows = filteredMovementData().slice(-20).reverse().map(r => ({{
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
    // Masterlist canvases report 0 width/height while their tab is display:none —
    // recompute the sticky-header offset and redraw only when the tab is visible.
    const mainHdr = document.querySelector(".sticky-dashboard-header");
    if (mainHdr) {{
        document.documentElement.style.setProperty("--ml-stick-top", mainHdr.offsetHeight + "px");
    }}
    if (document.getElementById("masterlistPanel")?.classList.contains("active")) {{
        mlDrawAll();
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

// ─── MASTERLIST TAB — CANVAS CHART SYSTEM ──────────────────────────────────
// Ported from masterlist-poc.html (Phase 2, Mike-approved). Every new global
// here is ml-/MC_/ML_-prefixed and every new DOM id is ml- prefixed so none
// of this can collide with the Coaching tab's donut()/bar()/scrollableBar()/
// accountTenureStack()/weeklyChart()/renderDonutWithSummary()/COLORS or the
// Quality tab's qa-* helpers. Data always comes from the SAME filtered
// dataset render() already computes via filteredMasterlist() — there is no
// second/parallel filter system here.
const MC_COLORS = ["#004C97", "#39B54A", "#1E7EE6", "#F59E0B", "#0891B2", "#8B5CF6", "#EF4444", "#F97316", "#0D9488"];
const ML_STATUS_COLORS = {{ ACTIVE: "#39B54A", PROBATION: "#F59E0B", INACTIVE: "#EF4444" }};

function mlGetVar(name) {{
    return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}}

function mlMkCanvas(id, w, h) {{
    const el = document.getElementById(id);
    if (!el) return null;
    const dpr = window.devicePixelRatio || 1;
    el.width = w * dpr; el.height = h * dpr;
    el.style.width = w + "px"; el.style.height = h + "px";
    const ctx = el.getContext("2d");
    ctx.scale(dpr, dpr);
    return {{el, ctx, w, h}};
}}

function mlEmptyState(ctx, w, h) {{
    ctx.clearRect(0, 0, w, h);
    ctx.textAlign = "center";
    ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
    ctx.font = "12px Arial";
    ctx.fillText("No data for selected filters", w / 2, h / 2);
}}

// Fix (item 4): measure the actual card-body INNER width at draw time via
// getBoundingClientRect() (not a hardcoded number) — subtract the immediate
// parent's REAL computed left/right padding instead of the old fixed "-24"
// magic number, which was only correct for the plain .ml-card-bd padding
// (10px 12px 13px = 24) and slightly off for cards like the hbar wrapper.
function mlCardWidth(canvasId, fallback) {{
    const el = document.getElementById(canvasId);
    if (!el || !el.parentElement) return fallback;
    const parent = el.parentElement;
    const rect = parent.getBoundingClientRect();
    const cs = window.getComputedStyle(parent);
    const padX = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0);
    const w = rect.width - padX;
    return w > 80 ? Math.floor(w) : fallback;
}}

function mlShowTip(x, y, text) {{
    const tt = document.getElementById("ml-cv-tt");
    if (!tt) return;
    tt.textContent = text;
    tt.style.left = (x + 14) + "px";
    tt.style.top = (y + 14) + "px";
    tt.classList.add("ml-show");
}}
// HTML variant (change 13) — only used where the tooltip needs a second line
// (e.g. Tenure by Account's per-segment "Total:" line). Callers are
// responsible for escaping any data-derived text before interpolating.
function mlShowTipHtml(x, y, html) {{
    const tt = document.getElementById("ml-cv-tt");
    if (!tt) return;
    tt.innerHTML = html;
    tt.style.left = (x + 14) + "px";
    tt.style.top = (y + 14) + "px";
    tt.classList.add("ml-show");
}}
function mlMoveTip(x, y) {{
    const tt = document.getElementById("ml-cv-tt");
    if (!tt) return;
    tt.style.left = (x + 14) + "px";
    tt.style.top = (y + 14) + "px";
}}
function mlHideTip() {{
    document.getElementById("ml-cv-tt")?.classList.remove("ml-show");
}}
// fmt (optional): function(rect) -> tooltip text. Defaults to "label — N employees".
// htmlMode (optional): when true, fmt's return value is treated as pre-escaped HTML
// (via mlShowTipHtml) instead of plain text — used only by the Tenure-by-Account
// stacked-segment tooltip (change 13), which needs a second "Total:" line.
//
// Perf fix (change 12): the hit-test used to run — and, worse, could have driven
// a full chart redraw — on every raw mousemove event. It now only (a) coalesces
// rapid events via requestAnimationFrame so the hit-test runs at most once per
// frame, and (b) skips re-building the tooltip text/DOM write entirely while the
// cursor stays over the SAME rect, only repositioning it. The underlying chart
// canvas is never redrawn here — only the floating #ml-cv-tt tooltip moves.
function mlAttachBarTooltip(el, rects, fmt, htmlMode) {{
    fmt = fmt || (r => `${{r.label}} — ${{r.value}} employee${{r.value === 1 ? "" : "s"}}`);
    let pendingEvt = null, rafId = null, lastHitIdx = -1;
    function process() {{
        rafId = null;
        const e = pendingEvt;
        pendingEvt = null;
        if (!e) return;
        const x = e.offsetX, y = e.offsetY;
        let hit = null, hitIdx = -1;
        for (let i = 0; i < rects.length; i++) {{
            const r = rects[i];
            if (x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h) {{ hit = r; hitIdx = i; break; }}
        }}
        if (hit) {{
            if (hitIdx !== lastHitIdx) {{
                if (htmlMode) mlShowTipHtml(e.clientX, e.clientY, fmt(hit));
                else mlShowTip(e.clientX, e.clientY, fmt(hit));
                lastHitIdx = hitIdx;
            }} else {{
                mlMoveTip(e.clientX, e.clientY);
            }}
            el.style.cursor = "pointer";
        }} else {{
            if (lastHitIdx !== -1) {{ mlHideTip(); lastHitIdx = -1; }}
            el.style.cursor = "default";
        }}
    }}
    el.onmousemove = (e) => {{
        pendingEvt = e;
        if (rafId == null) rafId = requestAnimationFrame(process);
    }};
    el.onmouseleave = () => {{
        if (rafId != null) {{ cancelAnimationFrame(rafId); rafId = null; }}
        pendingEvt = null;
        lastHitIdx = -1;
        mlHideTip();
        el.style.cursor = "default";
    }};
}}

// ── Legend layout helper (item A — standardized donut legends) — STRICT
//    column grid (default 2, parameterized via `cols` — Tweak 3: the Tenure
//    Segmentation donut requests 3 columns to fit its 8 bands into 3 rows;
//    every other donut keeps the default 2), not the old dynamic flow-wrap
//    (which could pack 3-4 short entries per row on one donut and only 1 on
//    another). Every donut calls this through mlDonut with the SAME
//    font/rowH/dot size, so the grid reads identically card-to-card; only
//    the column width adapts to that card's own longest label (so a
//    3-entry list like Active Status doesn't get stretched into a
//    half-empty column). mlDonut anchors this grid to the BOTTOM of the
//    canvas (Tweak 2) and reserves exactly
//    `rows * rowH` of height for it. ──────────────────────────────────────
function mlLegendLayout(ctx, items, w, font, rowH, startX, cols) {{
    startX = startX == null ? 8 : startX;
    cols = cols || 2; // default 2-column grid; Tenure Segmentation passes 3
    ctx.font = font;
    const colGap = 16; // gap between adjacent legend columns
    const dotTextGap = 14; // dot + gap before label text — matches the +14 draw offset in mlDonut
    const availW = w - startX * 2;
    // Change 1: columns are placed in N equal slots spanning the FULL
    // available legend width edge-to-edge (was startX + i*(colW+colGap),
    // which packed columns tightly from the left using only as much width
    // as the longest label needed — on a short-label donut like the 3-column
    // Tenure Segmentation legend that left all 3 columns clustered in the
    // left/middle of the card instead of reaching the right edge). Each
    // column's own text-wrap width is still capped by its longest label
    // (or the slot minus colGap, whichever is smaller) so labels never
    // collide across a slot boundary — only the DOT/column starting X now
    // comes from the even slot division.
    const slotW = availW / cols;
    let maxTextW = 0;
    items.forEach(it => {{
        const tw = ctx.measureText(it.name).width;
        if (tw > maxTextW) maxTextW = tw;
    }});
    const colW = Math.max(20, Math.min(slotW - colGap, dotTextGap + maxTextW));
    const colX = [];
    for (let i = 0; i < cols; i++) colX.push(startX + i * slotW);
    const positions = items.map((it, i) => ({{
        item: it,
        x: colX[i % cols],
        y: Math.floor(i / cols) * rowH,
        colW,
    }}));
    const rows = Math.max(1, Math.ceil(items.length / cols));
    return {{positions, rows, colW}};
}}
// Truncates text with an ellipsis so a legend label can never overflow past
// its column into the neighboring one (item A — strict column grid).
function mlEllipsize(ctx, text, maxW) {{
    if (ctx.measureText(text).width <= maxW) return text;
    let out = text;
    while (out.length > 1 && ctx.measureText(out + "…").width > maxW) out = out.slice(0, -1);
    return out + "…";
}}

function mlDonutNaturalHeight(segs, w, h, opts) {{
    opts = opts || {{}};
    w = w || 220; h = h || 219;
    const total = (segs || []).reduce((s, x) => s + x.count, 0);
    if (opts.mini || !total) return h;
    const legendCols = opts.legendCols || 2;
    const legendFont = "10px Arial", legendRowH = 18, legendGap = 22;
    const meas = document.createElement("canvas").getContext("2d");
    const legendLayout = mlLegendLayout(meas, segs, w, legendFont, legendRowH, undefined, legendCols);
    const legendH = legendGap + legendLayout.rows * legendRowH + 6;
    const minArcH = 150;
    return Math.max(h, minArcH + legendH);
}}

// ── Active Status donut — blinking legend dot overlay (cosmetic only) ──────
// Canvas can't animate a fill, so these two small DOM dots sit exactly on
// top of the canvas-drawn "Active"/"Inactive" legend swatches for the
// "c-ml-active" card only — every other donut legend is untouched. Called
// from inside mlDonut (below) every time that card redraws, so it tracks
// filter-change re-renders and window resizes automatically without any
// separate polling loop. Coordinates are converted from the canvas's own
// logical drawing space (w x drawH, the same space legendBaseY/p.x/p.y are
// computed in) into actual on-screen CSS pixels via getBoundingClientRect(),
// which correctly accounts for devicePixelRatio (the canvas's internal
// width/height are DPR-scaled, but style.width/height and therefore the
// rendered box are not) AND any CSS max-width squeeze of the canvas element
// relative to its authored w/h.
function mlClearActiveLegendDots(canvasEl) {{
    const wrap = canvasEl && canvasEl.parentElement;
    if (!wrap) return;
    wrap.querySelectorAll(".ml-legend-pulse-dot").forEach(d => d.remove());
}}
function mlSyncActiveLegendDots(canvasEl, legendLayout, legendBaseY, w, drawH) {{
    const wrap = canvasEl && canvasEl.parentElement;
    if (!wrap || !legendLayout || !legendLayout.positions.length) {{ mlClearActiveLegendDots(canvasEl); return; }}
    let dots = Array.from(wrap.querySelectorAll(".ml-legend-pulse-dot"));
    if (dots.length !== legendLayout.positions.length) {{
        dots.forEach(d => d.remove());
        dots = legendLayout.positions.map(() => {{
            const d = document.createElement("div");
            d.className = "ml-legend-pulse-dot";
            d.setAttribute("aria-hidden", "true");
            wrap.appendChild(d);
            return d;
        }});
    }}
    const canvasRect = canvasEl.getBoundingClientRect();
    const wrapRect = wrap.getBoundingClientRect();
    if (!canvasRect.width || !canvasRect.height) return; // hidden tab, etc.
    const scaleX = canvasRect.width / w;
    const scaleY = canvasRect.height / drawH;
    const offLeft = canvasRect.left - wrapRect.left;
    const offTop = canvasRect.top - wrapRect.top;
    const dotD = 8 * Math.min(scaleX, scaleY); // matches the r=4 canvas arc (diameter 8)
    legendLayout.positions.forEach((p, i) => {{
        const dot = dots[i];
        if (!dot) return;
        dot.style.background = p.item.color;
        dot.style.width = dotD + "px";
        dot.style.height = dotD + "px";
        dot.style.left = (offLeft + (p.x + 4) * scaleX - dotD / 2) + "px";
        dot.style.top = (offTop + (legendBaseY + p.y + 4) * scaleY - dotD / 2) + "px";
    }});
}}

// ── Donut chart — segs: [{{name, count, color}}] ────────────────────────────
function mlDonut(id, segs, w, h, opts) {{
    opts = opts || {{}};
    w = w || 220; h = h || 219;
    const total = (segs || []).reduce((s, x) => s + x.count, 0);
    // Fix (item 3): measure the legend BEFORE sizing the canvas, using a
    // detached scratch context (same pattern as mlHbar's label measuring
    // below), so the card can grow a little to fit every entry — proper
    // row/column spacing, no squashing — instead of being locked to the old
    // fixed height regardless of how many legend rows a big list (e.g. "By
    // Department" with 11 items) actually needs.
    // Tweak 3: legend column count is parameterized (default 2); only the
    // Tenure Segmentation donut (8 bands) requests 3 via opts.legendCols —
    // every other donut keeps the default 2-column grid.
    const legendCols = opts.legendCols || 2;
    const legendFont = "10px Arial", legendRowH = 18, legendGap = opts.mini ? 0 : 22;
    const meas = document.createElement("canvas").getContext("2d");
    const legendLayout = (opts.mini || !total) ? {{positions: [], rows: 0}} : mlLegendLayout(meas, segs, w, legendFont, legendRowH, undefined, legendCols);
    const legendH = opts.mini ? 0 : (legendGap + legendLayout.rows * legendRowH + 6);
    const minArcH = 150; // room for the arc + on-chart labels regardless of legend height
    const naturalDrawH = opts.mini ? h : Math.max(h, minArcH + legendH);
    const drawH = opts.fixedDrawH ? Math.max(opts.fixedDrawH, naturalDrawH) : naturalDrawH;
    const c = mlMkCanvas(id, w, drawH);
    if (!c) return;
    const ctx = c.ctx;
    ctx.clearRect(0, 0, w, drawH);
    if (!total) {{ mlEmptyState(ctx, w, drawH); if (id === "c-ml-active") mlClearActiveLegendDots(c.el); c.el.onmousemove = null; c.el.onmouseleave = null; c.el.style.cursor = "default"; return; }}
    // Change 7/8: legend now shows EVERY segment (not just the first 4) and
    // wraps to as many rows as it needs — pre-measured above at the smaller
    // 10px legend font so the arc is sized around the ACTUAL legend height,
    // giving a clear, non-touching gap between the arc/on-chart labels and
    // the legend block below, whether it's a 3-segment or an 8-segment donut.
    const cx = w / 2, cy = (drawH - legendH) / 2 + (opts.mini ? 0 : 4);
    const r = Math.min(cx, cy) * (opts.mini ? 0.9 : 0.76), ir = r * 0.58;
    let a = -Math.PI / 2;
    const arcs = [];
    segs.forEach(s => {{
        const sw = (s.count / total) * Math.PI * 2;
        if (s.count > 0) {{
            ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, a, a + sw); ctx.closePath();
            ctx.fillStyle = s.color; ctx.fill();
            arcs.push({{start: a, end: a + sw, seg: s}});
        }}
        a += sw;
    }});
    ctx.beginPath(); ctx.arc(cx, cy, ir, 0, Math.PI * 2);
    ctx.fillStyle = "#fff";
    ctx.fill();

    if (!opts.mini) {{
        arcs.forEach(ar => {{
            const pct = ar.seg.count / total;
            if (pct < 0.04) return;
            const mid = (ar.start + ar.end) / 2;
            const ax = cx + Math.cos(mid) * (r + 2), ay = cy + Math.sin(mid) * (r + 2);
            const lx = cx + Math.cos(mid) * (r + 15), ly = cy + Math.sin(mid) * (r + 15);
            ctx.strokeStyle = ar.seg.color; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(ax, ay); ctx.lineTo(lx, ly); ctx.stroke();
            ctx.font = "bold 13px Arial"; ctx.textBaseline = "middle";
            ctx.textAlign = Math.cos(mid) >= 0 ? "left" : "right";
            ctx.fillStyle = mlGetVar("--text") || "#0F2240";
            ctx.fillText(ar.seg.count, lx + (Math.cos(mid) >= 0 ? 4 : -4), ly);
        }});
        ctx.textBaseline = "alphabetic";
    }}

    if (opts.mini) {{ if (id === "c-ml-active") mlClearActiveLegendDots(c.el); c.el.onmousemove = null; c.el.onmouseleave = null; return; }}

    const big = segs.reduce((m, s) => s.count > m.count ? s : m, segs[0]);
    ctx.textAlign = "center";
    ctx.fillStyle = mlGetVar("--text") || "#0F2240";
    ctx.font = "bold 22px Arial";
    ctx.fillText(big.count, cx, cy + 2);
    ctx.font = "11px Arial";
    ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
    ctx.fillText(big.name, cx, cy + 17);

    // Item A: draw every legend entry (including 0-count ones, e.g. Tenure
    // Segmentation bands with no members) from the pre-measured strict
    // column grid above (2, or 3 for Tenure Segmentation) — same dot
    // size/row height/font on every donut, wrapped rows land exactly where
    // mlLegendLayout() predicted (legendH already reserved it).
    // Tweak 2: anchor the legend block to the BOTTOM of the canvas (fixed
    // ~22px gap from the card-body bottom edge to the LAST row's text
    // baseline) instead of floating directly beneath the arc at a variable
    // position — matches the weekly-headcount chart's fixed bottom legend.
    // drawH already reserves legendH of room above this point (the arc's
    // own sizing/centering via cy, above, is untouched), so a taller legend
    // (more rows) only pushes legendBaseY — and the arc above it — further
    // up; it can never collide with or run past the bottom edge.
    const legendBottomPad = opts.legendBottomPad == null ? 12 : opts.legendBottomPad;
    const legendBaseY = drawH - legendBottomPad - 8 - (legendLayout.rows - 1) * legendRowH;
    ctx.font = legendFont; ctx.textAlign = "left";
    legendLayout.positions.forEach(p => {{
        // Active Status card: the swatch dot is NOT drawn on canvas — the
        // pulsing DOM overlay dot (mlSyncActiveLegendDots) is the only dot,
        // otherwise the static canvas dot shows through the fade and the
        // blink is invisible.
        if (id !== "c-ml-active") {{
            ctx.fillStyle = p.item.color;
            ctx.beginPath(); ctx.arc(p.x + 4, legendBaseY + p.y + 4, 4, 0, Math.PI * 2); ctx.fill();
        }}
        ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
        // +14 gap between the swatch dot and its label text; label is
        // ellipsized to the column width so long names (e.g. a consolidated
        // "HR Group" department slice) never overlap the next column.
        const maxTextW = Math.max(4, p.colW - 14);
        ctx.fillText(mlEllipsize(ctx, p.item.name, maxTextW), p.x + 14, legendBaseY + p.y + 8);
    }});

    // Blinking legend dot overlay — Active Status card only (see
    // mlSyncActiveLegendDots above). Runs on every call to mlDonut for this
    // card, so it self-corrects on filter-change re-renders, the expand
    // modal close redraw (mlRedrawCard), and the debounced window resize
    // handler (mlRedrawIfActive -> mlDrawAll) — no separate polling needed.
    if (id === "c-ml-active") {{
        mlSyncActiveLegendDots(c.el, legendLayout, legendBaseY, w, drawH);
    }}

    c.el.onmousemove = (e) => {{
        const x = e.offsetX, y = e.offsetY, dx = x - cx, dy = y - cy;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < ir || dist > r) {{ mlHideTip(); c.el.style.cursor = "default"; return; }}
        let ang = Math.atan2(dy, dx);
        if (ang < -Math.PI / 2) ang += Math.PI * 2;
        let hit = null;
        for (let i = 0; i < arcs.length; i++) {{
            if (ang >= arcs[i].start && ang <= arcs[i].end) {{ hit = arcs[i]; break; }}
        }}
        if (hit) {{
            const pct = Math.round((hit.seg.count / total) * 100);
            mlShowTip(e.clientX, e.clientY, `${{hit.seg.name}} — ${{hit.seg.count}} employee${{hit.seg.count === 1 ? "" : "s"}} (${{pct}}%)`);
            c.el.style.cursor = "pointer";
        }} else {{ mlHideTip(); c.el.style.cursor = "default"; }}
    }};
    c.el.onmouseleave = () => {{ mlHideTip(); c.el.style.cursor = "default"; }};
}}

// ── Horizontal bar — shows EVERY row (data must already be full, no slice) ──
function mlHbar(id, data, w, opts) {{
    opts = opts || {{}};
    w = w || 280;
    const barH = opts.barH || 22, gap = opts.gap || 10;
    if (!data || !data.length || data.reduce((s, d) => s + d.count, 0) === 0) {{
        const cEmpty = mlMkCanvas(id, w, 120);
        if (!cEmpty) return;
        mlEmptyState(cEmpty.ctx, w, 120);
        return;
    }}
    const meas = document.createElement("canvas").getContext("2d");
    meas.font = "11px Arial";
    let labelW = 0; data.forEach(d => {{ const mw = meas.measureText(d.name).width; if (mw > labelW) labelW = mw; }});
    meas.font = "bold 11px Arial";
    let valueW = 0; data.forEach(d => {{ const mw = meas.measureText(String(d.count)).width; if (mw > valueW) valueW = mw; }});
    const pad = {{ t: 8, r: Math.max(28, valueW + 16), b: 8, l: Math.min(220, Math.max(72, labelW + 16)) }};
    const h = pad.t + pad.b + data.length * (barH + gap) - gap;
    const c = mlMkCanvas(id, w, h);
    if (!c) return;
    const ctx = c.ctx;
    ctx.clearRect(0, 0, w, h);
    const maxV = Math.max.apply(null, data.map(d => d.count)) || 1;
    const bw = w - pad.l - pad.r;
    const rects = [];
    data.forEach((d, i) => {{
        const y = pad.t + (barH + gap) * i;
        const blen = (d.count / maxV) * bw;
        ctx.fillStyle = "rgba(0,76,151,.06)";
        ctx.beginPath(); ctx.roundRect(pad.l, y, bw, barH, 3); ctx.fill();
        const grad = ctx.createLinearGradient(pad.l, 0, pad.l + Math.max(blen, 2), 0);
        grad.addColorStop(0, "#004C97"); grad.addColorStop(1, "#1E7EE6");
        ctx.fillStyle = grad;
        ctx.beginPath(); ctx.roundRect(pad.l, y, Math.max(blen, 2), barH, 3); ctx.fill();
        ctx.textAlign = "right"; ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
        ctx.font = "11px Arial";
        ctx.fillText(d.name, pad.l - 8, y + barH / 2 + 4);
        ctx.textAlign = "left"; ctx.fillStyle = mlGetVar("--text") || "#0F2240";
        ctx.font = "bold 11px Arial";
        ctx.fillText(d.count, pad.l + blen + 6, y + barH / 2 + 4);
        rects.push({{x: pad.l, y, w: bw, h: barH, label: d.name, value: d.count}});
    }});
    mlAttachBarTooltip(c.el, rects);
}}

// ── Vertical bar (Weekly Headcount) ─────────────────────────────────────────
function mlVbar(id, data, w, h) {{
    w = w || 320; h = h || 219;
    const c = mlMkCanvas(id, w, h);
    if (!c) return;
    const ctx = c.ctx;
    ctx.clearRect(0, 0, w, h);
    if (!data || !data.length || data.reduce((s, d) => s + d.count, 0) === 0) {{ mlEmptyState(ctx, w, h); return; }}
    const pad = {{ t: 20, r: 10, b: 26, l: 8 }};
    const cw = w - pad.l - pad.r, ch = h - pad.t - pad.b;
    const n = data.length, gap = 10;
    const barW = Math.min(40, (cw - gap * (n - 1)) / n);
    const totalW = barW * n + gap * (n - 1);
    const startX = pad.l + (cw - totalW) / 2;
    const maxV = Math.max.apply(null, data.map(d => d.count)) || 1;
    const rects = [];
    data.forEach((d, i) => {{
        const x = startX + i * (barW + gap);
        const bh = (d.count / maxV) * ch;
        const y = pad.t + ch - bh;
        ctx.fillStyle = "rgba(57,181,74,.08)";
        ctx.beginPath(); ctx.roundRect(x, pad.t, barW, ch, 3); ctx.fill();
        const grad = ctx.createLinearGradient(0, y, 0, pad.t + ch);
        grad.addColorStop(0, "#39B54A"); grad.addColorStop(1, "#8fd89b");
        ctx.fillStyle = grad;
        ctx.beginPath(); ctx.roundRect(x, y, barW, Math.max(bh, 2), 3); ctx.fill();
        ctx.textAlign = "center"; ctx.fillStyle = mlGetVar("--text") || "#0F2240";
        ctx.font = "bold 10px Arial";
        ctx.fillText(d.count, x + barW / 2, Math.max(y - 6, 10));
        ctx.fillStyle = mlGetVar("--muted") || "#5A6B80"; ctx.font = "9px Arial";
        ctx.fillText(d.name, x + barW / 2, h - pad.b + 12);
        rects.push({{x, y: pad.t, w: barW, h: ch, label: d.name, value: d.count}});
    }});
    mlAttachBarTooltip(c.el, rects);
}}

// ── Stacked bar (Tenure by Account) ─────────────────────────────────────────
// WCAG-based readable text-contrast helper for in-segment value labels.
function mlContrastTextColor(hex) {{
    const r = parseInt(hex.slice(1, 3), 16) / 255, g = parseInt(hex.slice(3, 5), 16) / 255, b = parseInt(hex.slice(5, 7), 16) / 255;
    const lin = ch => ch <= 0.03928 ? ch / 12.92 : Math.pow((ch + 0.055) / 1.055, 2.4);
    const L = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
    const contrastWhite = 1.05 / (L + 0.05);
    const contrastDark = (L + 0.05) / 0.05;
    return contrastWhite >= contrastDark ? "#FFFFFF" : "#0F2240";
}}
function mlStackbar(id, accounts, groups, colors, w, h) {{
    w = w || 900; h = h || 299;
    const c = mlMkCanvas(id, w, h);
    if (!c) return;
    const ctx = c.ctx;
    ctx.clearRect(0, 0, w, h);
    if (!accounts || !accounts.length) {{ mlEmptyState(ctx, w, h); c.el.onmousemove = null; c.el.onmouseleave = null; return; }}
    const pad = {{ t: 8, r: 16, b: 40, l: 14 }};
    const bw = Math.max((w - pad.l - pad.r) / accounts.length - 10, 18);
    const maxV = Math.max.apply(null, accounts.map(a => a.vals.reduce((s, v) => s + v, 0))) || 1;
    const ch = h - pad.t - pad.b;
    const segRects = [];
    const segLabelFont = "bold 11px Arial";
    accounts.forEach((acc, i) => {{
        const x = pad.l + i * (bw + 10);
        let y = h - pad.b;
        const accountTotal = acc.vals.reduce((s, v) => s + v, 0); // change 13
        acc.vals.forEach((v, j) => {{
            const bh = (v / maxV) * ch;
            ctx.fillStyle = colors[j];
            ctx.beginPath();
            if (j === acc.vals.length - 1) ctx.roundRect(x, y - bh, bw, bh, 3);
            else ctx.rect(x, y - bh, bw, bh);
            ctx.fill();
            segRects.push({{x, y: y - bh, w: bw, h: bh, account: acc.name, band: groups[j], value: v, total: accountTotal}});
            if (v > 0) {{
                ctx.font = segLabelFont;
                const txt = String(v);
                const tw = ctx.measureText(txt).width;
                if (bh >= 15 && bw >= tw + 6) {{
                    ctx.fillStyle = mlContrastTextColor(colors[j]);
                    ctx.textAlign = "center"; ctx.textBaseline = "middle";
                    ctx.fillText(txt, x + bw / 2, y - bh / 2 + 1);
                }}
            }}
            y -= bh;
        }});
        ctx.textBaseline = "alphabetic";
        ctx.textAlign = "center"; ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
        ctx.font = "10px Arial";
        ctx.fillText(acc.name, x + bw / 2, h - pad.b + 14);
    }});
    let lx = pad.l, ly = h - 10;
    groups.forEach((g, i) => {{
        ctx.fillStyle = colors[i];
        ctx.beginPath(); ctx.arc(lx + 4, ly, 4, 0, Math.PI * 2); ctx.fill();
        ctx.textAlign = "left"; ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
        ctx.font = "9px Arial";
        ctx.fillText(g, lx + 11, ly + 4);
        lx += ctx.measureText(g).width + 24;
        if (lx > w - 110) {{ lx = pad.l; ly += 13; }}
    }});
    // Change 13: tooltip keeps the existing "<Account> — <band>: N employees" line
    // and adds a second line with the account's TOTAL across all bands.
    mlAttachBarTooltip(c.el, segRects, r =>
        `${{escapeHtml(r.account)}} — ${{escapeHtml(r.band)}}: ${{r.value}} employee${{r.value === 1 ? "" : "s"}}` +
        `<br>Total: ${{r.total}} employee${{r.total === 1 ? "" : "s"}}`,
        true);
}}

// ── "Nice" axis-step helper (item C) — picks a round gridline interval
// (1/2/5 × a power of ten) for a given max value and target tick count, e.g.
// a max of 187 with 4 ticks -> step 50 (gridlines at 0/50/100/150/200), never
// an arbitrary in-between number. ───────────────────────────────────────────
function mlNiceStep(maxV, ticks) {{
    ticks = ticks || 4;
    if (!maxV || maxV <= 0) return 1;
    const raw = maxV / ticks;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const norm = raw / mag;
    const nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10;
    return nice * mag;
}}

// ── Vertical stacked bar (Weekly Headcount Trend, by Employment Class) —
// weekly: {{weeks: [{{name, vals: [...]}}], classes: [...], colors: [...]}}.
// Item C rework to match the reference design:
//   • stack order fixed by mlBuildWeeklyStacked (Regular bottom/green,
//     Probationary top/blue) — this function just draws whatever order/
//     colors it's given, bottom segment first;
//   • bold total-count label above each bar;
//   • in-segment count labels, drawn only when a segment is tall/wide
//     enough to hold the text without overlapping;
//   • "nice"-step y-axis gridlines (faint, dotted) with small value labels;
//   • slim bars (~30% of each slot) with generous (~70%) gaps, real
//     left/right chart margins;
//   • two-line x-axis labels (week date, muted year below);
//   • bottom, left-aligned legend.
function mlVStackbar(id, weekly, w, h) {{
    w = w || 320; h = h || 260;
    const weeks = (weekly && weekly.weeks) || [];
    const classes = (weekly && weekly.classes) || [];
    const colors = (weekly && weekly.colors) || [];
    const c = mlMkCanvas(id, w, h);
    if (!c) return;
    const ctx = c.ctx;
    ctx.clearRect(0, 0, w, h);
    const grandTotal = weeks.reduce((s, wk) => s + wk.vals.reduce((s2, v) => s2 + v, 0), 0);
    if (!weeks.length || !grandTotal) {{ mlEmptyState(ctx, w, h); c.el.onmousemove = null; c.el.onmouseleave = null; return; }}
    // Fixed chart margins (item C) — pad.l leaves room for the y-axis value
    // labels, pad.b leaves room for the axis gap + 2-line x-labels + legend
    // row + bottom breathing room. Kept in sync with mlWeeklyNaturalWidth's
    // own pad.l/pad.r + minSlotW below, so the .ml-hscroll-x fallback only
    // ever has to scroll — it never has to squeeze a bar narrower than
    // minBarW.
    // Keep the same card/canvas size, but lower the x-axis and tighten the
    // date label gap so the weekly labels read as attached to the axis instead
    // of floating in the lower whitespace.
    const pad = {{t: 30, r: 16, b: 68, l: 36}};
    const cw = w - pad.l - pad.r, ch = h - pad.t - pad.b;
    const n = weeks.length;
    // Change 4: barRatio 0.3 -> 0.4 — bars now ~40% of each slot, ~60% gap
    // (was ~30%/~70%), i.e. a 60% gap width as requested.
    const minBarW = 16, barRatio = 0.4; // bars ~40% of each slot, ~60% gap
    let barW, gap;
    if (n <= 1) {{
        barW = Math.max(minBarW, cw * barRatio); gap = 0;
    }} else {{
        const slotW = cw / n;
        barW = Math.max(minBarW, slotW * barRatio);
        gap = slotW - barW;
    }}
    const maxV = Math.max.apply(null, weeks.map(wk => wk.vals.reduce((s, v) => s + v, 0))) || 1;
    const axisY = pad.t + ch;

    // Y-axis — "nice" step gridlines (faint dotted, zero line solid) + small
    // muted value labels. Bars below are scaled to the SAME topTick (not
    // maxV) so the tallest bar lines up exactly with its gridline instead of
    // always touching the very top of the chart regardless of scale.
    const step = mlNiceStep(maxV, 4);
    const topTick = Math.max(step, Math.ceil(maxV / step) * step);
    ctx.textAlign = "right"; ctx.textBaseline = "middle"; ctx.font = "10px Arial";
    for (let tick = 0; tick <= topTick + 0.0001; tick += step) {{
        const y = axisY - (tick / topTick) * ch;
        ctx.strokeStyle = tick === 0 ? "rgba(15,32,64,.18)" : "rgba(15,32,64,.10)";
        ctx.setLineDash(tick === 0 ? [] : [2, 3]);
        ctx.lineWidth = 1;
        ctx.beginPath(); ctx.moveTo(pad.l, y); ctx.lineTo(pad.l + cw, y); ctx.stroke();
        ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
        ctx.fillText(String(Math.round(tick)), pad.l - 8, y);
    }}
    ctx.setLineDash([]);
    ctx.textBaseline = "alphabetic";

    // Bars — drawn bottom segment first (index 0) upward; only the topmost
    // segment gets rounded TOP corners (bottom corners stay square so it
    // joins the segment below with no visual seam).
    const segRects = [];
    const segLabelFont = "bold 11px Arial";
    weeks.forEach((wk, i) => {{
        const x = pad.l + i * (barW + gap);
        let y = axisY;
        const total = wk.vals.reduce((s, v) => s + v, 0);
        wk.vals.forEach((v, j) => {{
            const bh = (v / topTick) * ch;
            const isTop = j === wk.vals.length - 1;
            if (v > 0) {{
                ctx.fillStyle = colors[j];
                ctx.beginPath();
                if (isTop) ctx.roundRect(x, y - bh, barW, bh, [3, 3, 0, 0]);
                else ctx.rect(x, y - bh, barW, bh);
                ctx.fill();
                const txt = String(v);
                ctx.font = segLabelFont;
                const tw = ctx.measureText(txt).width;
                // Only skip a genuinely tiny sliver or a segment narrower than
                // the count text. Counts are intentionally used without the
                // percentage suffix to keep the labels readable in compact bars.
                // Height threshold scaled from 9 -> 10 to stay proportional
                // after segLabelFont grew from 10px to 11px (keeps the same
                // suppression behavior relative to the new glyph height).
                if (bh >= 10 && barW >= tw + 4) {{
                    // WCAG-based contrast pick (item C accessibility note):
                    // Mike's reference calls for white text on both segments,
                    // but white-on-green (#39B54A) is only ~2.7:1 — below the
                    // 4.5:1 body-text minimum at this label size. Reusing the
                    // same mlContrastTextColor() helper already used by the
                    // Tenure-by-Account stacked bar keeps white on the blue
                    // (Probationary) segment and switches to dark navy on the
                    // green (Regular) segment automatically.
                    ctx.fillStyle = mlContrastTextColor(colors[j]);
                    ctx.textAlign = "center"; ctx.textBaseline = "middle";
                    ctx.lineWidth = 2;
                    ctx.strokeStyle = ctx.fillStyle === "#fff" ? "rgba(15,32,64,.25)" : "rgba(255,255,255,.45)";
                    ctx.strokeText(txt, x + barW / 2, y - bh / 2 + 1);
                    ctx.fillText(txt, x + barW / 2, y - bh / 2 + 1);
                    ctx.textBaseline = "alphabetic";
                }}
            }}
            segRects.push({{x, y: y - bh, w: barW, h: Math.max(bh, v > 0 ? 1 : 0), week: wk.name, cls: classes[j] || "Unspecified", value: v, total}});
            y -= bh;
        }});
        // Bold total-count label above the bar.
        ctx.font = "bold 13px Arial"; ctx.textAlign = "center";
        ctx.fillStyle = mlGetVar("--text") || "#0F2240";
        const topY = axisY - (total / topTick) * ch;
        ctx.fillText(String(total), x + barW / 2, Math.max(topY - 8, 12));

        // Two-line x-axis label: week date on top, muted 4-digit year below.
        const wkDate = parseDateValue(wk.name);
        const line1 = wkDate ? `${{MONTH_LABELS[wkDate.getMonth()]}} ${{wkDate.getDate()}}` : wk.name;
        ctx.font = "bold 11px Arial"; ctx.fillStyle = mlGetVar("--text") || "#0F2240";
        ctx.fillText(line1, x + barW / 2, axisY + 16);
        if (wkDate) {{
            ctx.font = "10px Arial"; ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
            ctx.fillText(String(wkDate.getFullYear()), x + barW / 2, axisY + 27);
        }}
    }});

    // Legend — bottom, left-aligned, listed top-segment-first (Probationary,
    // then Regular) to match the stack's visual top-to-bottom reading order.
    const colorByClass = Object.fromEntries(classes.map((cl, i) => [cl, colors[i]]));
    const legendOrder = [...classes].reverse();
    ctx.font = "11px Arial"; ctx.textAlign = "left";
    let lx = pad.l;
    const legendY = h - 6;
    legendOrder.forEach(cls => {{
        ctx.fillStyle = colorByClass[cls];
        ctx.beginPath(); ctx.arc(lx + 4, legendY, 4, 0, Math.PI * 2); ctx.fill();
        ctx.fillStyle = mlGetVar("--muted") || "#5A6B80";
        ctx.fillText(cls, lx + 12, legendY + 4);
        lx += ctx.measureText(cls).width + 12 + 22;
    }});

    mlAttachBarTooltip(c.el, segRects, r => `${{r.week}} — ${{r.cls}}: ${{r.value}} employee${{r.value === 1 ? "" : "s"}} (Total: ${{r.total}})`);
}}

// ── Aggregation helpers — reuse countBy()/tenureCounts()/ageGroupCounts()/
//    TENURE_GROUPS/AGE_GROUPS/APPROVED_ACCOUNTS/tenureGroupName() (all defined
//    above); do not duplicate their bucketing logic. ──────────────────────────
function mlWithColors(list) {{
    return list.map((d, i) => ({{name: d.name, count: d.count, color: MC_COLORS[i % MC_COLORS.length]}}));
}}
function mlStatusSegs(data) {{
    const order = ["Active", "Probation", "Inactive"];
    const counts = Object.fromEntries(order.map(s => [s, 0]));
    let other = 0;
    data.forEach(r => {{
        const st = norm(r["Employment Status"]);
        const key = order.find(o => o.toLowerCase() === st.toLowerCase());
        if (key) counts[key] += 1; else if (st) other += 1;
    }});
    const segs = order.filter(k => counts[k] > 0).map(k => ({{name: k, count: counts[k], color: ML_STATUS_COLORS[k.toUpperCase()]}}));
    if (other > 0) segs.push({{name: "Other", count: other, color: "#94A3B8"}});
    return segs;
}}
function mlTenureSegs(data) {{
    // Item A: show EVERY tenure band in the legend, even one with 0 members
    // right now (was filtered out here, silently dropping bands like
    // "31-60 Days" from the legend when nobody currently falls in them).
    return tenureCounts(data).map((d, i) => ({{name: d.name, count: d.count, color: MC_COLORS[i % MC_COLORS.length]}}));
}}
// By-Department DONUT ONLY: collapse every "HR"-prefixed sub-department
// (case-insensitive after trim, e.g. "HR", "HR - Payroll", "HR-Finance",
// "hr ops") into a single "HR Group" slice. This is a donut-only view of the
// raw "Department" field — it does NOT mutate `data`/`masterlist`, so the
// Master List table (mlRenderTable), the Department filter dropdown
// (uniqueValues(masterlist, "Department")), and the "Departments" KPI tile
// (uniqueValues(data, "Department").length) all keep reading the untouched
// per-employee Department values elsewhere and are unaffected by this
// grouping. Blank/missing values fall back to "Blank" (same as countBy())
// and are never folded into "HR Group".
function mlDeptDonutCounts(data) {{
    const out = {{}};
    data.forEach(r => {{
        const raw = norm(r["Department"]);
        const key = raw ? (raw.toLowerCase().startsWith("hr") ? "HR Group" : raw) : "Blank";
        out[key] = (out[key] || 0) + 1;
    }});
    return Object.entries(out).map(([name, count]) => ({{name, count}})).sort((a, b) => b.count - a.count);
}}
function mlAccountTenureStack(data) {{
    const accountNames = countBy(data, "LOB / Account")
        .filter(d => APPROVED_ACCOUNTS.has(d.name.toLowerCase()))
        .map(d => d.name);
    const counts = {{}};
    accountNames.forEach(account => {{ counts[account] = Object.fromEntries(TENURE_GROUPS.map(g => [g.name, 0])); }});
    data.forEach(r => {{
        const account = norm(r["LOB / Account"]) || "Blank";
        if (!counts[account]) return;
        const groupName = tenureGroupName(parseDateValue(r["Hire Date"]), new Date());
        if (groupName) counts[account][groupName] += 1;
    }});
    const accounts = accountNames.map(name => ({{name, vals: TENURE_GROUPS.map(g => counts[name][g.name])}}));
    return {{
        accounts,
        groups: TENURE_GROUPS.map(g => g.name),
        colors: TENURE_GROUPS.map((g, i) => MC_COLORS[i % MC_COLORS.length]),
    }};
}}
// Change 11: Weekly Headcount Trend is now a STACKED bar by Employment Class
// (was a single-series total). Same source/dedupe logic as the old
// mlBuildWeekly() (History sheet "Week" column, one row per employee per
// Monday-start week, deduped by week+ID) — now also buckets by
// "Employement Class" (typo preserved to match the masterlist/history field).
function mlBuildWeeklyStacked() {{
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
    const weekKeys = [...new Set(deduped.map(r => norm(r["Week"])))].filter(Boolean).sort((a, b) => new Date(a) - new Date(b));
    const classSet = new Set();
    deduped.forEach(r => {{ classSet.add(norm(r["Employement Class"]) || "Unspecified"); }});
    // Item C: fixed stack order — Regular always first (drawn at the BOTTOM
    // of the stack, per mlVStackbar's bottom-up draw loop) then Probationary
    // (drawn on TOP), matching the reference design. Any other/unexpected
    // class value the data happens to contain is appended afterward
    // (alphabetically) so the chart never silently drops a class it doesn't
    // recognize.
    const ML_CLASS_ORDER = ["Regular", "Probationary"];
    const classes = [
        ...ML_CLASS_ORDER.filter(cls => classSet.has(cls)),
        ...[...classSet].filter(cls => !ML_CLASS_ORDER.includes(cls)).sort(),
    ];
    // Fixed brand colors so Regular is always green / Probationary always
    // blue regardless of draw order; any extra class falls back to the
    // shared MC_COLORS palette.
    const ML_CLASS_COLORS = {{Regular: "#39B54A", Probationary: "#004C97"}};
    let extraColorIdx = 0;
    const colors = classes.map(cls => ML_CLASS_COLORS[cls] || MC_COLORS[(extraColorIdx++) % MC_COLORS.length]);
    const counts = {{}};
    weekKeys.forEach(w => {{ counts[w] = Object.fromEntries(classes.map(c => [c, 0])); }});
    deduped.forEach(r => {{
        const week = norm(r["Week"]);
        if (!counts[week]) return;
        const cls = norm(r["Employement Class"]) || "Unspecified";
        counts[week][cls] = (counts[week][cls] || 0) + 1;
    }});
    const weeks = weekKeys.map(w => ({{name: w, vals: classes.map(c => counts[w][c])}}));
    return {{ weeks, classes, colors }};
}}

// ── Natural-width helpers (robustness B1/B2 + change 11) — the minimum bar
// width each chart needs to stay legible, independent of the card's visible
// width. Callers take Math.max(cardWidth, naturalWidth) and let the
// .ml-hscroll-x wrapper (or the modal's own overflow:auto body) scroll
// horizontally whenever the natural width wins. ────────────────────────────
function mlStackNaturalWidth(accountCount) {{
    const minBarW = 34, gap = 10, padLR = 14 + 16;
    return padLR + Math.max(accountCount, 1) * (minBarW + gap) - gap;
}}
function mlWeeklyNaturalWidth(weekCount) {{
    // Item C: kept in sync with mlVStackbar's own pad.l/pad.r (36/16) and
    // bar-gap geometry. Change 4 note: minBarW/barRatio moved from 16/0.3 to
    // 16/0.4 (minSlotW 54 -> 40) in mlVStackbar — updated here too so this
    // stays in sync; the .ml-hscroll-x wrapper only ever needs to scroll, it
    // never has to squeeze a bar narrower than minBarW.
    const minSlotW = 40, padLR = 36 + 16;
    return padLR + Math.max(weekCount, 1) * minSlotW;
}}

// ── Chart registry (used by the expand modal) + last-rendered args ─────────
const ML_CHART_FN = {{
    dept: "donut", active: "donut", empgrp: "donut", empclass: "donut", tenureseg: "donut",
    account: "hbar", manager: "hbar", supervisor: "hbar", age: "hbar",
    tenurestack: "stackbar", weekly: "vstackbar",
}};
let mlCurrentArgs = {{}};
// Item C: ML_VSTACK_H raised 219 -> 264 to fit the new total-count label
// above the bars, the two-line x-axis (date + year), and the bottom legend
// row without cramping — see mlVStackbar's own pad.t/pad.b for the exact
// vertical budget this height needs to satisfy.
const ML_DONUT_H = 219, ML_STACKBAR_H = 299, ML_VBAR_H = 219, ML_VSTACK_H = 264;
let mlCurrentDonutRowH = ML_DONUT_H;

function mlRenderCharts(data) {{
    mlCurrentArgs.dept = mlWithColors(mlDeptDonutCounts(data));
    mlCurrentArgs.active = mlStatusSegs(data);
    mlCurrentArgs.empgrp = mlWithColors(countBy(data, "Employee Group"));
    mlCurrentArgs.empclass = mlWithColors(countBy(data, "Employement Class"));
    mlCurrentArgs.tenureseg = mlTenureSegs(data);
    // Tweak 3: 3-column legend for Tenure Segmentation's 8 bands (all other
    // donuts stay the default 2-column grid — see mlDonut/mlLegendLayout).
    const rowDonutH = Math.max(
        mlDonutNaturalHeight(mlCurrentArgs.dept, 220, ML_DONUT_H),
        mlDonutNaturalHeight(mlCurrentArgs.active, 220, ML_DONUT_H),
        mlDonutNaturalHeight(mlCurrentArgs.empgrp, 220, ML_DONUT_H),
        mlDonutNaturalHeight(mlCurrentArgs.empclass, 220, ML_DONUT_H),
        mlDonutNaturalHeight(mlCurrentArgs.tenureseg, 220, ML_DONUT_H, {{legendCols: 3}})
    );
    mlCurrentDonutRowH = rowDonutH;
    mlDonut("c-ml-dept", mlCurrentArgs.dept, 220, ML_DONUT_H, {{fixedDrawH: rowDonutH}});
    mlDonut("c-ml-active", mlCurrentArgs.active, 220, ML_DONUT_H, {{fixedDrawH: rowDonutH}});
    mlDonut("c-ml-empgrp", mlCurrentArgs.empgrp, 220, ML_DONUT_H, {{fixedDrawH: rowDonutH}});
    mlDonut("c-ml-empclass", mlCurrentArgs.empclass, 220, ML_DONUT_H, {{fixedDrawH: rowDonutH}});
    mlDonut("c-ml-tenureseg", mlCurrentArgs.tenureseg, 220, ML_DONUT_H, {{legendCols: 3, fixedDrawH: rowDonutH}});

    mlCurrentArgs.account = countBy(data, "LOB / Account").filter(d => APPROVED_ACCOUNTS.has(d.name.toLowerCase()));
    mlHbar("c-ml-account", mlCurrentArgs.account, mlCardWidth("c-ml-account", 280));

    mlCurrentArgs.manager = countBy(data, "Manager");
    mlHbar("c-ml-manager", mlCurrentArgs.manager, mlCardWidth("c-ml-manager", 280));

    mlCurrentArgs.supervisor = countBy(data, "Immediate Supervisor");
    mlHbar("c-ml-supervisor", mlCurrentArgs.supervisor, mlCardWidth("c-ml-supervisor", 280));

    mlCurrentArgs.age = ageGroupCounts(data);
    mlHbar("c-ml-age", mlCurrentArgs.age, mlCardWidth("c-ml-age", 280));

    // Change 13/B2: canvas width is Math.max(visible card width, the natural
    // width the account count actually needs at a legible minimum bar width) —
    // the .ml-hscroll-x wrapper (added to this card's markup) scrolls whenever
    // the natural width wins, instead of the old behavior of squeezing every
    // bar to fit and eventually clipping content off-canvas.
    mlCurrentArgs.tenurestack = mlAccountTenureStack(data);
    {{
        const stackW = Math.max(mlCardWidth("c-ml-tenurestack", 900), mlStackNaturalWidth(mlCurrentArgs.tenurestack.accounts.length));
        mlStackbar("c-ml-tenurestack", mlCurrentArgs.tenurestack.accounts, mlCurrentArgs.tenurestack.groups, mlCurrentArgs.tenurestack.colors, stackW, ML_STACKBAR_H);
    }}

    // Change 11/B1: stacked-by-Employment-Class weekly trend, full card width,
    // with the same natural-width + horizontal-scroll robustness as above so
    // it never degrades into unreadable slivers as more weeks accumulate.
    mlCurrentArgs.weekly = mlBuildWeeklyStacked();
    {{
        const weeklyW = Math.max(mlCardWidth("c-ml-weekly", 320), mlWeeklyNaturalWidth(mlCurrentArgs.weekly.weeks.length));
        mlVStackbar("c-ml-weekly", mlCurrentArgs.weekly, weeklyW, ML_VSTACK_H);
    }}

    mlRenderTable(data);
}}

function mlDrawAll() {{
    mlRenderCharts(filteredMasterlist());
}}

// ── Master List table — click-to-sort + pagination ──────────────────────────
// NOTE: preserves the pre-existing quirk of the old masterlistTable(): the
// Employee Name multi-filter (MASTERLIST_FILTERS.empName) is applied ONLY to
// the table, not to the charts above — same as before this port.
// Change A (fixes change #10): restored to the full 13 columns of the dead
// masterlistTable()/MASTERLIST_COLUMNS reference — same field keys, same
// labels (incl. the "Employement Class" typo), same order. The 7-column trim
// is what caused header labels to render wrong (e.g. "Class" instead of
// "Employement Class", "Account" instead of "LOB/Account").
const ML_COLUMNS = [
    {{key: "ID No.", label: "Employee ID", sortType: "number", cls: "ml-col-id"}},
    {{key: "Emp Name", label: "Employee Name", cls: "ml-col-name"}},
    {{key: "Hire Date", label: "Hire Date", sortType: "date", cls: "ml-col-date"}},
    {{key: "Employement Class", label: "Employement Class", cls: "ml-col-class"}},
    {{key: "__mlTenureDays", label: "Tenure", cls: "ml-col-tenure"}},
    {{key: "Job Title", label: "Job Title", cls: "ml-col-title"}},
    {{key: "Employee Group", label: "Employee Group", cls: "ml-col-group"}},
    {{key: "Department", label: "Department", cls: "ml-col-dept"}},
    {{key: "LOB / Account", label: "LOB/Account", cls: "ml-col-account"}},
    {{key: "Immediate Supervisor", label: "Immediate Supervisor", cls: "ml-col-supervisor"}},
    {{key: "Manager", label: "Manager", cls: "ml-col-manager"}},
    {{key: "Employment Status", label: "Employment Status", cls: "ml-col-status"}},
    {{key: "Company Email", label: "Email", cls: "ml-col-email"}},
];
// Default sort mirrors the pre-port default (masterlistSortState: Emp Name asc)
// so the first paint isn't a jarring switch to unsorted/insertion order.
const ML_TABLE_STATE = {{page: 1, pageSize: 20, sortKey: "Emp Name", sortDir: "asc"}};

function mlSortGlyph(key) {{
    if (ML_TABLE_STATE.sortKey !== key) return "\\u2195";
    return ML_TABLE_STATE.sortDir === "asc" ? "\\u25B2" : "\\u25BC";
}}
function mlBuildTheadHtml(interactive) {{
    return "<tr>" + ML_COLUMNS.map(c => {{
        const active = ML_TABLE_STATE.sortKey === c.key;
        const classes = [];
        if (c.cls) classes.push(c.cls);
        if (active) classes.push("ml-sorted");
        if (!interactive) classes.push("ml-th-static");
        const clsAttr = classes.length ? ` class="${{classes.join(" ")}}"` : "";
        let attrs = "";
        if (interactive) {{
            const ariaSort = active ? (ML_TABLE_STATE.sortDir === "asc" ? "ascending" : "descending") : "none";
            attrs = ` data-key="${{c.key}}" tabindex="0" aria-sort="${{ariaSort}}"`;
        }}
        return `<th${{clsAttr}}${{attrs}}>${{escapeHtml(c.label)}} <span class="ml-sort-ic" aria-hidden="true">${{mlSortGlyph(c.key)}}</span></th>`;
    }}).join("") + "</tr>";
}}
function mlRenderThead() {{
    const thead = document.getElementById("ml-thead");
    if (thead) thead.innerHTML = mlBuildTheadHtml(true);
}}
function mlCycleSort(key) {{
    if (ML_TABLE_STATE.sortKey !== key) {{ ML_TABLE_STATE.sortKey = key; ML_TABLE_STATE.sortDir = "asc"; }}
    else if (ML_TABLE_STATE.sortDir === "asc") {{ ML_TABLE_STATE.sortDir = "desc"; }}
    else {{ ML_TABLE_STATE.sortKey = null; ML_TABLE_STATE.sortDir = null; }}
    ML_TABLE_STATE.page = 1;
    mlRenderTable(filteredMasterlist());
}}
// Change A: sort now respects each column's declared sortType (number/date),
// falling back to the existing tenure-days special case and plain string
// compare — matching the original MASTERLIST_COLUMNS sort behavior instead of
// treating every restored column (e.g. "ID No.", "Hire Date") as a string.
function mlSortRows(rows) {{
    if (!ML_TABLE_STATE.sortKey) return rows.slice();
    const key = ML_TABLE_STATE.sortKey, dir = ML_TABLE_STATE.sortDir;
    const col = ML_COLUMNS.find(c => c.key === key) || {{}};
    return rows.slice().sort((a, b) => {{
        let cmp;
        if (key === "__mlTenureDays") {{
            cmp = (a.__mlTenureDays || 0) - (b.__mlTenureDays || 0);
        }} else if (col.sortType === "date") {{
            const av = parseDateValue(a[key])?.getTime() || 0;
            const bv = parseDateValue(b[key])?.getTime() || 0;
            cmp = av - bv;
        }} else if (col.sortType === "number") {{
            const av = Number(norm(a[key]).replace(/,/g, "")) || 0;
            const bv = Number(norm(b[key]).replace(/,/g, "")) || 0;
            cmp = av - bv;
        }} else {{
            cmp = norm(a[key]).localeCompare(norm(b[key]), undefined, {{sensitivity: "base"}});
        }}
        return dir === "desc" ? -cmp : cmp;
    }});
}}
function mlPillClass(status) {{
    const v = norm(status).toUpperCase();
    if (v === "ACTIVE") return "ml-pa";
    if (v === "INACTIVE") return "ml-pi";
    if (v) return "ml-pp";
    return "";
}}
// Change A: cell rendering is now driven generically off ML_COLUMNS so all 13
// columns stay in sync with the thead with no per-column hardcoding to drift.
function mlCellHtml(r, col) {{
    if (col.key === "__mlTenureDays") {{
        const label = escapeHtml(r.__mlTenureLabel || "\\u2014");
        return r.__mlTenureFlag ? `<span class="ml-pill ${{r.__mlTenureFlag}}">${{label}}</span>` : label;
    }}
    if (col.key === "Employment Status") {{
        const pillClass = mlPillClass(r[col.key]);
        return pillClass ? `<span class="ml-pill ${{pillClass}}">${{escapeHtml(r[col.key])}}</span>` : "";
    }}
    return escapeHtml(r[col.key]);
}}
function mlBuildRowsHtml(rows) {{
    if (!rows.length) {{
        return `<tr><td colspan="${{ML_COLUMNS.length}}" style="text-align:center;color:var(--muted);padding:170px 0;border-bottom:none">No employees match the selected filters.</td></tr>`;
    }}
    return rows.map(r => "<tr>" + ML_COLUMNS.map(c => `<td class="${{c.cls || ""}}"><span class="ml-clip" title="${{escapeHtml(norm(c.key === "__mlTenureDays" ? r.__mlTenureLabel : r[c.key]))}}">${{mlCellHtml(r, c)}}</span></td>`).join("") + "</tr>").join("");
}}
function mlRenderPager(total, totalPages, start, end) {{
    const rangeEl = document.getElementById("ml-pager-range");
    const pageEl = document.getElementById("ml-pager-page");
    const prevBtn = document.getElementById("ml-prev");
    const nextBtn = document.getElementById("ml-next");
    if (!rangeEl || !pageEl || !prevBtn || !nextBtn) return;
    rangeEl.textContent = total === 0 ? "Showing 0 of 0 employees" : `Showing ${{start + 1}}\\u2013${{end}} of ${{total}} employees`;
    pageEl.textContent = `Page ${{ML_TABLE_STATE.page}} of ${{totalPages}}`;
    prevBtn.disabled = ML_TABLE_STATE.page <= 1;
    nextBtn.disabled = ML_TABLE_STATE.page >= totalPages;
}}
// Item 2 thresholds (Probationary rows only): >=6 months -> red (ml-pi, same
// convention as the Inactive pill), >=5 and <6 months -> yellow (ml-pp, same
// convention as the Probation pill). Non-Probationary rows and rows with no
// parseable Hire Date never get a flag. tenureMonths MUST come from the same
// tenureBreakdown() call that produced the displayed label so the color can
// never contradict the visible number (see mlPrepareRows).
function mlTenureFlag(empClass, tenureMonths) {{
    if (norm(empClass) !== "Probationary" || tenureMonths < 0) return "";
    if (tenureMonths >= 6) return "ml-pi";
    if (tenureMonths >= 5) return "ml-pp";
    return "";
}}
function mlPrepareRows(data) {{
    const now = new Date();
    return data.map(r => {{
        const hireDate = parseDateValue(r["Hire Date"]);
        const breakdown = hireDate ? tenureBreakdown(hireDate, now) : null;
        return Object.assign({{}}, r, {{
            __mlTenureLabel: (hireDate ? formatTenureDisplay(hireDate, now) : "") || "\\u2014",
            __mlTenureDays: breakdown ? breakdown.days : -1,
            __mlTenureFlag: mlTenureFlag(r["Employement Class"], breakdown ? breakdown.totalMonths : -1),
        }});
    }});
}}
function mlTableData(data) {{
    return data.filter(r => filterMatches(MASTERLIST_FILTERS.empName, norm(r["Emp Name"])));
}}
function mlRenderTable(data) {{
    mlRenderThead();
    const prepared = mlPrepareRows(mlTableData(data));
    const sorted = mlSortRows(prepared);
    const total = sorted.length;
    const totalPages = Math.max(1, Math.ceil(total / ML_TABLE_STATE.pageSize));
    if (ML_TABLE_STATE.page > totalPages) ML_TABLE_STATE.page = totalPages;
    if (ML_TABLE_STATE.page < 1) ML_TABLE_STATE.page = 1;
    const start = total ? (ML_TABLE_STATE.page - 1) * ML_TABLE_STATE.pageSize : 0;
    const end = Math.min(start + ML_TABLE_STATE.pageSize, total);
    const tbody = document.getElementById("ml-tbody");
    if (tbody) tbody.innerHTML = mlBuildRowsHtml(sorted.slice(start, end));
    mlRenderPager(total, totalPages, start, end);
}}
function mlWireTable() {{
    const thead = document.getElementById("ml-thead");
    if (thead) {{
        thead.addEventListener("click", (e) => {{
            const th = e.target.closest("th[data-key]");
            if (!th || !thead.contains(th)) return;
            mlCycleSort(th.getAttribute("data-key"));
        }});
        thead.addEventListener("keydown", (e) => {{
            if (e.key !== "Enter" && e.key !== " ") return;
            const th = e.target.closest("th[data-key]");
            if (!th || !thead.contains(th)) return;
            e.preventDefault();
            mlCycleSort(th.getAttribute("data-key"));
        }});
    }}
    const prevBtn = document.getElementById("ml-prev");
    const nextBtn = document.getElementById("ml-next");
    if (prevBtn) prevBtn.addEventListener("click", () => {{
        if (!prevBtn.disabled) {{ ML_TABLE_STATE.page--; mlRenderTable(filteredMasterlist()); }}
    }});
    if (nextBtn) nextBtn.addEventListener("click", () => {{
        if (!nextBtn.disabled) {{ ML_TABLE_STATE.page++; mlRenderTable(filteredMasterlist()); }}
    }});
}}

// ── Expand modal ─────────────────────────────────────────────────────────
let mlActiveCid = null;
let mlLastTrigger = null;

function mlOpenExpand(card, triggerEl) {{
    const ovl = document.getElementById("ml-ovl");
    const xbd = document.getElementById("ml-xbd");
    const xttl = document.getElementById("ml-xttl");
    const xsub = document.getElementById("ml-xsub");
    const xcls = document.getElementById("ml-xcls");
    if (!ovl || !xbd || !card) return;
    mlLastTrigger = triggerEl || null;
    const cid = card.getAttribute("data-cid");
    if (xttl) xttl.textContent = card.getAttribute("data-ttl") || "";
    if (xsub) xsub.textContent = card.getAttribute("data-sub") || "";
    xbd.innerHTML = "";
    xbd.classList.remove("ml-xbd-table", "ml-xbd-scroll");
    mlActiveCid = cid;
    const isTableCard = card.classList.contains("ml-tcrd");
    if (isTableCard) {{
        xbd.classList.add("ml-xbd-table");
        if (cid === "masterlist") {{
            const sortedAll = mlSortRows(mlPrepareRows(mlTableData(filteredMasterlist())));
            const wrap = document.createElement("div");
            wrap.className = "ml-twrap";
            const tbl = document.createElement("table");
            tbl.className = "ml-dt";
            const theadEl = document.createElement("thead");
            theadEl.innerHTML = mlBuildTheadHtml(false);
            const tbodyEl = document.createElement("tbody");
            tbodyEl.innerHTML = mlBuildRowsHtml(sortedAll);
            tbl.appendChild(theadEl); tbl.appendChild(tbodyEl);
            wrap.appendChild(tbl);
            xbd.appendChild(wrap);
        }}
    }} else if (cid && ML_CHART_FN[cid]) {{
        const fn = ML_CHART_FN[cid];
        const cw = Math.min(1000, window.innerWidth - 100);
        const el = document.createElement("canvas");
        el.id = "xc-ml-" + cid;
        xbd.appendChild(el);
        if (fn === "donut") {{
            const ch = Math.min(480, window.innerHeight - 200);
            // Tweak 3: preserve the 3-column legend for Tenure Segmentation
            // in the expand modal too; every other donut stays default 2.
            mlDonut(el.id, mlCurrentArgs[cid], cw > 600 ? 420 : 300, ch > 360 ? 360 : 280, cid === "tenureseg" ? {{legendCols: 3}} : undefined);
        }} else if (fn === "hbar") {{
            xbd.classList.add("ml-xbd-scroll");
            mlHbar(el.id, mlCurrentArgs[cid], cw, {{barH: 26, gap: 14}});
        }} else if (fn === "stackbar") {{
            // B2/change 13: same natural-width floor as the card view — #ml-xbd
            // already has overflow:auto, so it scrolls horizontally on its own.
            xbd.classList.add("ml-xbd-scroll");
            const stackW = Math.max(cw, mlStackNaturalWidth((mlCurrentArgs[cid].accounts || []).length));
            mlStackbar(el.id, mlCurrentArgs[cid].accounts, mlCurrentArgs[cid].groups, mlCurrentArgs[cid].colors, stackW, 340);
        }} else if (fn === "vstackbar") {{
            xbd.classList.add("ml-xbd-scroll");
            const weeklyW = Math.max(cw, mlWeeklyNaturalWidth(((mlCurrentArgs[cid] && mlCurrentArgs[cid].weeks) || []).length));
            const ch2 = Math.min(420, window.innerHeight - 220);
            mlVStackbar(el.id, mlCurrentArgs[cid], weeklyW, ch2);
        }}
    }}
    ovl.classList.add("ml-open");
    document.body.style.overflow = "hidden";
    if (xcls) xcls.focus();
}}
// Change 12 (expand-close scrollbar bug): the expanded canvas is a wholly
// separate DOM node (id "xc-ml-<cid>") appended into #ml-xbd, never the
// card's own "c-ml-<cid>" canvas — but to guard against any leftover inline
// sizing/oversized node lingering after close, this now (1) explicitly
// removes every canvas under #ml-xbd before clearing it, and (2) re-runs the
// ORIGINATING card's own chart at its normal card width so the small card
// canvas is guaranteed freshly sized/redrawn, never left showing a stray
// horizontal scrollbar from whatever was being displayed in the modal.
function mlRedrawCard(cid) {{
    if (!cid || !ML_CHART_FN[cid] || !mlCurrentArgs[cid]) return;
    const fn = ML_CHART_FN[cid];
    const canvasId = "c-ml-" + cid;
    if (fn === "donut") {{
        // Tweak 3: preserve the 3-column legend for Tenure Segmentation on
        // card redraw too; every other donut stays default 2.
        mlDonut(canvasId, mlCurrentArgs[cid], 220, ML_DONUT_H, cid === "tenureseg" ? {{legendCols: 3, fixedDrawH: mlCurrentDonutRowH}} : {{fixedDrawH: mlCurrentDonutRowH}});
    }} else if (fn === "hbar") {{
        mlHbar(canvasId, mlCurrentArgs[cid], mlCardWidth(canvasId, 280));
    }} else if (fn === "stackbar") {{
        const stackW = Math.max(mlCardWidth(canvasId, 900), mlStackNaturalWidth((mlCurrentArgs[cid].accounts || []).length));
        mlStackbar(canvasId, mlCurrentArgs[cid].accounts, mlCurrentArgs[cid].groups, mlCurrentArgs[cid].colors, stackW, ML_STACKBAR_H);
    }} else if (fn === "vstackbar") {{
        const weeklyW = Math.max(mlCardWidth(canvasId, 320), mlWeeklyNaturalWidth(((mlCurrentArgs[cid] && mlCurrentArgs[cid].weeks) || []).length));
        mlVStackbar(canvasId, mlCurrentArgs[cid], weeklyW, ML_VSTACK_H);
    }}
}}
function mlCloseExpand() {{
    const ovl = document.getElementById("ml-ovl");
    const xbd = document.getElementById("ml-xbd");
    if (!ovl || !xbd) return;
    if (!ovl.classList.contains("ml-open")) return;
    const closingCid = mlActiveCid;
    ovl.classList.remove("ml-open");
    document.body.style.overflow = "";
    mlActiveCid = null;
    xbd.querySelectorAll("canvas").forEach(cv => cv.remove()); // fully tear down the expanded canvas node
    xbd.innerHTML = "";
    xbd.classList.remove("ml-xbd-table", "ml-xbd-scroll");
    mlHideTip();
    mlRedrawCard(closingCid); // restore the originating card's own canvas at its proper width
    if (mlLastTrigger) {{ mlLastTrigger.focus(); mlLastTrigger = null; }}
}}
function mlWireExpandModal() {{
    const ovl = document.getElementById("ml-ovl");
    const xcls = document.getElementById("ml-xcls");
    if (xcls) xcls.addEventListener("click", mlCloseExpand);
    if (ovl) ovl.addEventListener("click", (e) => {{ if (e.target === ovl) mlCloseExpand(); }});
    document.addEventListener("keydown", (e) => {{
        if (e.key === "Escape" && ovl && ovl.classList.contains("ml-open")) mlCloseExpand();
    }});
}}
// ─── END MASTERLIST TAB — CANVAS CHART SYSTEM ──────────────────────────────

function render() {{
    const data = filteredMasterlist();

    const active = data.filter(r => norm(r["Employment Status"]).toUpperCase() === "ACTIVE").length;
    const inactive = data.filter(r => norm(r["Employment Status"]).toUpperCase() === "INACTIVE").length;

    // KPI strip (change 6 — 7 tiles: Active Employees | Total Headcount |
    // Departments | Accounts | Avg Tenure | Movements | History Records).
    // Active/Total/Departments react to the current filters (same "data" used
    // by every chart below); Avg Tenure/Movements/History Records are org-wide
    // scalars computed once in Python (masterlistKpis); Accounts is a fixed,
    // hardcoded 22 baked into the markup and is never touched here.
    setText("ml-kpi-active", active);
    setText("ml-kpi-total", data.length);
    const mlTotalSubEl = document.getElementById("ml-kpi-total-sub"); // change 2
    if (mlTotalSubEl) mlTotalSubEl.innerHTML = `Incl. <span style="color:#B91C1C;font-weight:700;font-size:13px">${{Number(inactive).toLocaleString()}}</span> inactive`;
    setText("ml-kpi-departments", uniqueValues(data, "Department").length);
    const mlAvgTenureEl = document.getElementById("ml-kpi-avgtenure");
    if (mlAvgTenureEl) mlAvgTenureEl.innerHTML = `${{Number(masterlistKpis.avgTenure || 0).toLocaleString()}}<span style="font-size:14px;font-weight:400"> yrs</span>`;
    setText("ml-kpi-movements", masterlistKpis.movementsPending || 0); // change 4
    const mlMovementsSubEl = document.getElementById("ml-kpi-movements-sub");
    if (mlMovementsSubEl) mlMovementsSubEl.textContent = `${{Number(masterlistKpis.movementsForProcessing || 0).toLocaleString()}} for processing`;
    setText("ml-kpi-history", masterlistKpis.historyRecords || 0); // change 5

    // Any filter change resets the Master List table back to page 1 (render()
    // only fires on a filter change or the initial load in this codebase —
    // sort-cycle/pager clicks call mlRenderTable() directly and do not hit this
    // reset, matching the POC's applyFilters() behavior).
    ML_TABLE_STATE.page = 1;

    mlRenderCharts(data);
    recentMovementsTable();
}}

// ─── QA QUALITY TAB ──────────────────────────────────────────────────────────
let qaChartsInitialized = false;
let qaTrendChart = null;
let qaAiQeTrendChart = null;
let qaDonutChart = null;
let qaEvalDistChart = null;
let qaAivhGapDonut = null;
let qaCurrentFiltered = [];

// Tag rows at init time
qaRawData.forEach(r => r._acct = 'M7');
dmgRawData.forEach(r => r._acct = 'DMG');
r4hRawData.forEach(r => r._acct = 'R4H');
parentisRawData.forEach(r => r._acct = 'Parentis');
briteliftRawData.forEach(r => r._acct = 'Britelift');
blcRawData.forEach(r => r._acct = 'Britelift Chat');
ridexRawData.forEach(r => r._acct = 'RideX');
hamiltonRawData.forEach(r => r._acct = 'Hamilton');
skylineRawData.forEach(r => r._acct = 'Skyline');
vipRawData.forEach(r => r._acct = 'VIP');
chRawData.forEach(r => r._acct = 'C&H');
rcRawData.forEach(r => r._acct = 'Reno Cab');
tiRawData.forEach(r => r._acct = 'Trans Iowa');
dcRawData.forEach(r => r._acct = 'Data Carz');
acRawData.forEach(r => r._acct = 'Associated Cab');
olRawData.forEach(r => r._acct = 'Ollies');
ctRawData.forEach(r => r._acct = 'Circle Taxi');
ycovRawData.forEach(r => r._acct = 'YCOV');
kelRawData.forEach(r => r._acct = 'Kelowna');
vtRawData.forEach(r => r._acct = 'Vermont');
ycdcRawData.forEach(r => r._acct = 'YCDC');
blRawData.forEach(r => r._acct = 'Blueline');

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
    {{key:"info_prec",name:"Information Precision",      pts:"10pts", inverse:false}},
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
    if (acct === 'dmg') return dmgRawData;
    if (acct === 'r4h') return r4hRawData;
    if (acct === 'parentis') return parentisRawData;
    if (acct === 'britelift') return briteliftRawData;
    if (acct === 'blc') return blcRawData;
    if (acct === 'ridex') return ridexRawData;
    if (acct === 'hamilton') return hamiltonRawData;
    if (acct === 'skyline') return skylineRawData;
    if (acct === 'vip') return vipRawData;
    if (acct === 'ch') return chRawData;
    if (acct === 'rc') return rcRawData;
    if (acct === 'ti') return tiRawData;
    if (acct === 'dc') return dcRawData;
    if (acct === 'ac') return acRawData;
    if (acct === 'ol') return olRawData;
    if (acct === 'ct') return ctRawData;
    if (acct === 'ycov') return ycovRawData;
    if (acct === 'kel') return kelRawData;
    if (acct === 'vt') return vtRawData;
    if (acct === 'ycdc') return ycdcRawData;
    if (acct === 'bl') return blRawData;
    return [...qaRawData, ...dmgRawData, ...r4hRawData, ...parentisRawData, ...briteliftRawData, ...blcRawData, ...ridexRawData, ...hamiltonRawData, ...skylineRawData, ...vipRawData, ...chRawData, ...rcRawData, ...tiRawData, ...dcRawData, ...acRawData, ...olRawData, ...ctRawData, ...ycovRawData, ...kelRawData, ...vtRawData, ...ycdcRawData, ...blRawData];
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
        else if(acct==='dmg') titleEl.textContent='DMG — Quality Assurance';
        else if(acct==='r4h') titleEl.textContent='R4H — Quality Assurance';
        else if(acct==='parentis') titleEl.textContent='Parentis Health — Quality Assurance';
        else if(acct==='britelift') titleEl.textContent='Britelift — Quality Assurance';
        else if(acct==='blc') titleEl.textContent='Britelift Chat — Quality Assurance';
        else if(acct==='ridex') titleEl.textContent='RideX — Quality Assurance';
        else if(acct==='hamilton') titleEl.textContent='Hamilton — Quality Assurance';
        else if(acct==='skyline') titleEl.textContent='Skyline — Quality Assurance';
        else if(acct==='vip') titleEl.textContent='VIP — Quality Assurance';
        else if(acct==='ch') titleEl.textContent='C&H — Quality Assurance';
        else if(acct==='rc') titleEl.textContent='Reno Cab — Quality Assurance';
        else if(acct==='ti') titleEl.textContent='Trans Iowa — Quality Assurance';
        else if(acct==='dc') titleEl.textContent='Data Carz — Quality Assurance';
        else if(acct==='ac') titleEl.textContent='Associated Cab — Quality Assurance';
        else if(acct==='ol') titleEl.textContent='Ollies — Quality Assurance';
        else if(acct==='ct') titleEl.textContent='Circle Taxi — Quality Assurance';
        else if(acct==='ycov') titleEl.textContent='YCOV — Quality Assurance';
        else if(acct==='kel') titleEl.textContent='Kelowna — Quality Assurance';
        else if(acct==='vt') titleEl.textContent='Vermont — Quality Assurance';
        else if(acct==='ycdc') titleEl.textContent='YCDC — Quality Assurance';
        else if(acct==='bl') titleEl.textContent='Blueline — Quality Assurance';
        else titleEl.textContent='All Accounts — Quality Assurance';
    }}
    if(badgeEl) {{
        if(acct==='m7') {{ badgeEl.textContent='M7 Account'; badgeEl.className='qa-badge qa-b-blue'; }}
        else if(acct==='dmg') {{ badgeEl.textContent='DMG'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='r4h') {{ badgeEl.textContent='R4H'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='parentis') {{ badgeEl.textContent='Parentis Health'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='britelift') {{ badgeEl.textContent='Britelift'; badgeEl.className='qa-badge qa-b-amber'; }}
        else if(acct==='blc') {{ badgeEl.textContent='Britelift Chat'; badgeEl.className='qa-badge qa-b-amber'; }}
        else if(acct==='ridex') {{ badgeEl.textContent='RideX'; badgeEl.className='qa-badge qa-b-purple'; }}
        else if(acct==='hamilton') {{ badgeEl.textContent='Hamilton'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='skyline') {{ badgeEl.textContent='Skyline'; badgeEl.className='qa-badge qa-b-skyline'; }}
        else if(acct==='vip') {{ badgeEl.textContent='VIP'; badgeEl.className='qa-badge qa-b-amber'; }}
        else if(acct==='ch') {{ badgeEl.textContent='C&H'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='rc') {{ badgeEl.textContent='Reno Cab'; badgeEl.className='qa-badge qa-b-green'; }}
        else if(acct==='ti') {{ badgeEl.textContent='Trans Iowa'; badgeEl.className='qa-badge qa-b-indigo'; }}
        else if(acct==='dc') {{ badgeEl.textContent='Data Carz'; badgeEl.className='qa-badge qa-b-orange'; }}
        else if(acct==='ac') {{ badgeEl.textContent='Associated Cab'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='ol') {{ badgeEl.textContent='Ollies'; badgeEl.className='qa-badge qa-b-rose'; }}
        else if(acct==='ct') {{ badgeEl.textContent='Circle Taxi'; badgeEl.className='qa-badge qa-b-cyan'; }}
        else if(acct==='ycov') {{ badgeEl.textContent='YCOV'; badgeEl.className='qa-badge qa-b-green'; }}
        else if(acct==='kel') {{ badgeEl.textContent='Kelowna'; badgeEl.className='qa-badge qa-b-green'; }}
        else if(acct==='vt') {{ badgeEl.textContent='Vermont'; badgeEl.className='qa-badge qa-b-teal'; }}
        else if(acct==='ycdc') {{ badgeEl.textContent='YCDC'; badgeEl.className='qa-badge qa-b-cyan'; }}
        else if(acct==='bl') {{ badgeEl.textContent='Blueline'; badgeEl.className='qa-badge qa-b-blue'; }}
        else {{ badgeEl.textContent='All Accounts'; badgeEl.className='qa-badge qa-b-amber'; }}
    }}
    const tblScroll=document.getElementById('qa-tbl-scroll-main');
    if(tblScroll) {{
        if(acct==='vip') tblScroll.classList.add('vip-mode');
        else tblScroll.classList.remove('vip-mode');
        if(acct==='ch') tblScroll.classList.add('ch-mode');
        else tblScroll.classList.remove('ch-mode');
        if(acct==='rc') tblScroll.classList.add('rc-mode');
        else tblScroll.classList.remove('rc-mode');
        if(acct==='ti') tblScroll.classList.add('ti-mode');
        else tblScroll.classList.remove('ti-mode');
        if(acct==='dc') tblScroll.classList.add('dc-mode');
        else tblScroll.classList.remove('dc-mode');
        if(acct==='ac') tblScroll.classList.add('ac-mode');
        else tblScroll.classList.remove('ac-mode');
        if(acct==='ol') tblScroll.classList.add('ol-mode');
        else tblScroll.classList.remove('ol-mode');
        if(acct==='ct') tblScroll.classList.add('ct-mode');
        else tblScroll.classList.remove('ct-mode');
        if(acct==='ycov') tblScroll.classList.add('ycov-mode');
        else tblScroll.classList.remove('ycov-mode');
        if(acct==='kel') tblScroll.classList.add('kel-mode');
        else tblScroll.classList.remove('kel-mode');
        if(acct==='vt') tblScroll.classList.add('vt-mode');
        else tblScroll.classList.remove('vt-mode');
        if(acct==='ycdc') tblScroll.classList.add('ycdc-mode');
        else tblScroll.classList.remove('ycdc-mode');
        if(acct==='bl') tblScroll.classList.add('bl-mode');
        else tblScroll.classList.remove('bl-mode');
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

function qaUpdateAiQeTrend(data) {{
    if(!qaAiQeTrendChart) return;
    const trend=qaBuildAiQeTrend(data);
    const labels=trend.map(t=>t.week);
    const aiAvgs=trend.map(t=>t.aiAvg);
    const qeAvgs=trend.map(t=>t.qeAvg);
    const aiCounts=trend.map(t=>t.aiCount);
    const qeCounts=trend.map(t=>t.qeCount);
    qaAiQeTrendChart.data.labels=labels;
    qaAiQeTrendChart.data.datasets[0].data=aiAvgs;
    qaAiQeTrendChart.data.datasets[1].data=qeAvgs;
    qaAiQeTrendChart.data.datasets[2].data=qeAvgs;
    qaAiQeTrendChart.data.datasets[3].data=aiCounts;
    qaAiQeTrendChart.data.datasets[4].data=qeCounts;
    qaAiQeTrendChart.options.scales.y1.max=Math.ceil(Math.max(...aiCounts,...qeCounts,1)/0.4);
    qaAiQeTrendChart.update();
    const totalAi=aiCounts.reduce((s,v)=>s+v,0);
    const totalQe=qeCounts.reduce((s,v)=>s+v,0);
    const sub=document.getElementById('qa-aiqe-trend-sub');
    if(sub)sub.textContent=`Date range & account only · AI ${{totalAi}} · QE ${{totalQe}}`;
    const badge=document.getElementById('qa-aiqe-trend-badge');
    if(badge)badge.innerHTML=totalAi||totalQe
        ?`<span style="color:#004C97">AI ${{totalAi}}</span> <span style="color:#94A3B8">/</span> <span style="color:#39B54A">QE ${{totalQe}}</span>`
        :'—';
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
        return`<div style="margin-bottom:7px"><div style="display:flex;justify-content:space-between;font-size:10px;color:#475569;margin-bottom:2px"><span>${{qaEscapeHtml(c.name)}}</span><span style="font-weight:700;color:${{color}}">${{c.pct}}%</span></div><div style="height:5px;background:#F1F5F9;border-radius:3px;overflow:hidden"><div style="height:100%;width:${{c.pct}}%;background:${{color}};border-radius:3px;transition:width .4s"></div></div></div>`;
    }}).join(''):`<div style="text-align:center;padding:16px;color:#0F9B58;font-size:12px">✓ All criteria above 95% — great job!</div>`;
    if(bdEl){{
        const byCoach={{}};
        data.forEach(r=>{{if(!r.coach)return;if(!byCoach[r.coach])byCoach[r.coach]={{n:0}};byCoach[r.coach].n++;}});
        bdEl.innerHTML=Object.entries(byCoach).sort((a,b)=>b[1].n-a[1].n).map(([name,d])=>
            `<div style="background:#F8FAFC;border-radius:8px;padding:5px 6px;text-align:center;min-height:58px;display:flex;flex-direction:column;align-items:center;justify-content:flex-start"><div title="${{qaEscapeHtml(name)}}" style="font-size:10px;line-height:1.08;color:#64748B;height:22px;width:100%;display:flex;align-items:flex-start;justify-content:center;overflow:hidden;text-wrap:balance">${{qaEscapeHtml(name)}}</div><div style="font-size:16px;line-height:1;font-weight:800;color:#0D3B6E;margin-top:2px">${{d.n}}</div><div style="font-size:9px;line-height:1.1;color:#94A3B8;margin-top:2px">${{d.n}} eval${{d.n===1?'':'s'}}</div></div>`
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
            :a.acct==='DMG'
            ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">DMG</span>`
            :a.acct==='R4H'
            ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">R4H</span>`
            :a.acct==='Britelift'
            ?`<span style="background:#FFF7ED;color:#C2410C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Britelift</span>`
            :a.acct==='Britelift Chat'
            ?`<span style="background:#FFF7ED;color:#9A3412;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Britelift Chat</span>`
            :a.acct==='RideX'
            ?`<span style="background:#F5F3FF;color:#6D28D9;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">RideX</span>`
            :a.acct==='Hamilton'
            ?`<span style="background:#ECFDF5;color:#065F46;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Hamilton</span>`
            :a.acct==='Skyline'
            ?`<span style="background:#F0F9FF;color:#0369A1;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Skyline</span>`
            :a.acct==='VIP'
            ?`<span style="background:#FFFBEB;color:#B45309;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">VIP</span>`
            :a.acct==='C&H'
            ?`<span style="background:#F0FDFA;color:#0F766E;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">C&amp;H</span>`
            :a.acct==='Reno Cab'
            ?`<span style="background:#F0FDF4;color:#15803D;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Reno Cab</span>`
            :a.acct==='Trans Iowa'
            ?`<span style="background:#EEF2FF;color:#4338CA;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Trans Iowa</span>`
            :a.acct==='Data Carz'
            ?`<span style="background:#FFF7ED;color:#C2410C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Data Carz</span>`
            :a.acct==='Associated Cab'
            ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Associated Cab</span>`
            :a.acct==='Ollies'
            ?`<span style="background:#FFF1F2;color:#BE123C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Ollies</span>`
            :a.acct==='Circle Taxi'
            ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Circle Taxi</span>`
            :a.acct==='YCOV'
            ?`<span style="background:#F0FDF4;color:#047857;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">YCOV</span>`
            :a.acct==='Kelowna'
            ?`<span style="background:#F0FDF4;color:#166534;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Kelowna</span>`
            :a.acct==='Vermont'
            ?`<span style="background:#F0FDFA;color:#0F766E;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Vermont</span>`
            :a.acct==='YCDC'
            ?`<span style="background:#ECFEFF;color:#155E75;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">YCDC</span>`
            :a.acct==='Blueline'
            ?`<span style="background:#EFF6FF;color:#1D4ED8;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Blueline</span>`
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
const QA_AI_ACCOUNTS=new Set(['Hamilton','Skyline','VIP','C&H','Reno Cab','Trans Iowa','Data Carz','Associated Cab','Ollies','Circle Taxi','YCOV','Kelowna','Vermont','YCDC','Blueline']);
const QA_COL_WIDTH_KEY='pacbiz.qa.detail.columnWidths.v1';
const QA_COL_MIN_WIDTH=80;
let qaDetailColumnWidths=null;

function qaDetailTable(){{
    return document.querySelector('#qa-tbl-scroll-main .qa-dtbl');
}}

function qaLoadColumnWidths(){{
    if(qaDetailColumnWidths) return qaDetailColumnWidths;
    try{{
        const saved=JSON.parse(localStorage.getItem(QA_COL_WIDTH_KEY)||'[]');
        qaDetailColumnWidths=Array.isArray(saved)?saved.map(v=>Math.max(QA_COL_MIN_WIDTH,Number(v)||0)):[];
    }}catch(e){{
        qaDetailColumnWidths=[];
    }}
    return qaDetailColumnWidths;
}}

function qaSaveColumnWidths(){{
    if(!qaDetailColumnWidths) return;
    try{{localStorage.setItem(QA_COL_WIDTH_KEY,JSON.stringify(qaDetailColumnWidths));}}catch(e){{}}
}}

function qaBaseColumnWidth(th){{
    const min=parseFloat((th.getAttribute('style')||'').match(/min-width:\\s*([0-9.]+)px/i)?.[1]||'0');
    return Math.max(QA_COL_MIN_WIDTH,min,Math.ceil(th.getBoundingClientRect().width)||0);
}}

function qaEnsureColGroup(table,ths){{
    let cg=table.querySelector('colgroup.qa-resize-cols');
    if(!cg){{
        cg=document.createElement('colgroup');
        cg.className='qa-resize-cols';
        table.insertBefore(cg,table.firstChild);
    }}
    while(cg.children.length<ths.length) cg.appendChild(document.createElement('col'));
    while(cg.children.length>ths.length) cg.removeChild(cg.lastElementChild);
    return cg;
}}

function qaApplyDetailColumnWidths(){{
    const table=qaDetailTable();
    if(!table) return;
    const ths=[...table.querySelectorAll('thead th')];
    if(!ths.length) return;
    const widths=qaLoadColumnWidths();
    const cg=qaEnsureColGroup(table,ths);
    let total=0;
    ths.forEach((th,i)=>{{
        if(!widths[i]) widths[i]=qaBaseColumnWidth(th);
        const w=Math.max(QA_COL_MIN_WIDTH,Math.round(widths[i]));
        widths[i]=w;
        const px=w+'px';
        cg.children[i].style.width=px;
        th.style.width=px;
        th.style.minWidth=px;
        th.style.maxWidth=px;
        total+=w;
    }});
    table.style.minWidth=Math.max(total,table.parentElement?.clientWidth||0)+'px';
    qaUpdateDetailStickyOffsets();
}}

function qaUpdateDetailStickyOffsets(){{
    const table=qaDetailTable();
    if(!table) return;
    const ths=[...table.querySelectorAll('thead th')];
    if(ths.length<5) return;
    const widths=qaLoadColumnWidths();
    let left=0;
    for(let i=0;i<5;i++){{
        const px=left+'px';
        table.querySelectorAll(`tr > *:nth-child(${{i+1}})`).forEach(cell=>{{cell.style.left=px;}});
        left+=Math.max(QA_COL_MIN_WIDTH,widths[i]||ths[i].offsetWidth||QA_COL_MIN_WIDTH);
    }}
}}

function qaInitResizableColumns(){{
    const table=qaDetailTable();
    if(!table) return;
    const ths=[...table.querySelectorAll('thead th')];
    if(!ths.length) return;
    qaApplyDetailColumnWidths();
    ths.forEach((th,i)=>{{
        if(th.querySelector('.qa-col-resizer')) return;
        th.classList.add('qa-resizable');
        const handle=document.createElement('span');
        handle.className='qa-col-resizer';
        handle.setAttribute('aria-hidden','true');
        handle.addEventListener('click',e=>{{e.preventDefault();e.stopPropagation();}});
        handle.addEventListener('pointerdown',e=>{{
            e.preventDefault();
            e.stopPropagation();
            const widths=qaLoadColumnWidths();
            widths[i]=widths[i]||qaBaseColumnWidth(th);
            const startX=e.clientX;
            const startW=widths[i];
            document.body.classList.add('qa-col-resizing');
            const move=ev=>{{
                ev.preventDefault();
                widths[i]=Math.max(QA_COL_MIN_WIDTH,Math.round(startW+ev.clientX-startX));
                qaApplyDetailColumnWidths();
                qaRenderVisibleRows();
                qaSaveColumnWidths();
            }};
            const up=()=>{{
                document.body.classList.remove('qa-col-resizing');
                document.removeEventListener('pointermove',move);
                document.removeEventListener('pointerup',up);
                document.removeEventListener('pointercancel',up);
                qaSaveColumnWidths();
            }};
            document.addEventListener('pointermove',move,{{passive:false}});
            document.addEventListener('pointerup',up);
            document.addEventListener('pointercancel',up);
        }});
        th.appendChild(handle);
    }});
}}

function qaRowHtml(r){{
    const score=Number(r.score),sc=!isNaN(score)&&score>0?score:null;
    const chipCls=sc!==null?qaChipCls(sc):'',disp=sc!==null?sc.toFixed(1)+'%':'—';
    const av=QA_AV[r.agent]||{{bg:'#F1F5F9',tc:'#475569',ini:(r.agent||'?').split(' ').map(w=>w[0]||'').join('').slice(0,2).toUpperCase()}};
    const acctPill=r._acct==='M7'
        ?`<span style="background:#EFF6FF;color:#1D4ED8;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">M7</span>`
    :r._acct==='DMG'
        ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">DMG</span>`
    :r._acct==='R4H'
        ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">R4H</span>`
    :r._acct==='Britelift'
        ?`<span style="background:#FFF7ED;color:#C2410C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Britelift</span>`
        :r._acct==='Britelift Chat'
        ?`<span style="background:#FFF7ED;color:#9A3412;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Britelift Chat</span>`
        :r._acct==='RideX'
        ?`<span style="background:#F5F3FF;color:#6D28D9;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">RideX</span>`
        :r._acct==='Hamilton'
        ?`<span style="background:#ECFDF5;color:#065F46;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Hamilton</span>`
        :r._acct==='Skyline'
        ?`<span style="background:#F0F9FF;color:#0369A1;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Skyline</span>`
        :r._acct==='VIP'
        ?`<span style="background:#FFFBEB;color:#B45309;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">VIP</span>`
        :r._acct==='C&H'
        ?`<span style="background:#F0FDFA;color:#0F766E;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">C&amp;H</span>`
        :r._acct==='Reno Cab'
        ?`<span style="background:#F0FDF4;color:#15803D;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Reno Cab</span>`
        :r._acct==='Trans Iowa'
        ?`<span style="background:#EEF2FF;color:#4338CA;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Trans Iowa</span>`
        :r._acct==='Data Carz'
        ?`<span style="background:#FFF7ED;color:#C2410C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Data Carz</span>`
        :r._acct==='Associated Cab'
        ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Associated Cab</span>`
        :r._acct==='Ollies'
        ?`<span style="background:#FFF1F2;color:#BE123C;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Ollies</span>`
        :r._acct==='Circle Taxi'
        ?`<span style="background:#ECFEFF;color:#0E7490;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Circle Taxi</span>`
        :r._acct==='YCOV'
        ?`<span style="background:#F0FDF4;color:#047857;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">YCOV</span>`
        :r._acct==='Kelowna'
        ?`<span style="background:#F0FDF4;color:#166534;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Kelowna</span>`
        :r._acct==='Vermont'
        ?`<span style="background:#F0FDFA;color:#0F766E;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Vermont</span>`
        :r._acct==='YCDC'
        ?`<span style="background:#ECFEFF;color:#155E75;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">YCDC</span>`
        :r._acct==='Blueline'
        ?`<span style="background:#EFF6FF;color:#1D4ED8;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Blueline</span>`
        :`<span style="background:#FFF0F3;color:#9F1239;border-radius:4px;padding:1px 6px;font-size:9px;font-weight:700">Parentis</span>`;
    const critKeys=['os_in','os_out','closing','approp','no_resp','fillers','ack','hold','ack_hold',
                    'resp_eff','empathy','adjust','mute','active','gen_q','answered','probing','verif',
                    'clarif','lost_sop','rude','trans','speech','info_prec'];
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
    const vipExtraKeys=['vip_os_out2','vip_profess','vip_verif_other','vip_comm_quality','vip_cust_verif_oth','vip_res_etiq','vip_sop'];
    const vipExtraCells=vipExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="vip-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="vip-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="vip-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const chExtraKeys=['ch_os_out2','ch_profess','ch_verif_other','ch_comm_quality','ch_cust_verif_meas','ch_res_etiq','ch_subjective'];
    const chExtraCells=chExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="ch-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="ch-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="ch-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const rcExtraKeys=['rc_profess','rc_verif_other','rc_res_etiq','rc_comm_quality','rc_cust_verif_meas'];
    const rcExtraCells=rcExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="rc-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="rc-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="rc-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const tiExtraKeys=['ti_os_out2','ti_profess','ti_verif_other','ti_res_etiq','ti_comm_quality'];
    const tiExtraCells=tiExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="ti-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="ti-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="ti-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const dcExtraKeys=['dc_os_out2','dc_profess','dc_res_etiq','dc_comm_quality'];
    const dcExtraCells=dcExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="dc-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="dc-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="dc-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const acExtraKeys=['ac_os_out2','ac_profess','ac_verif_other','ac_comm_quality','ac_cust_verif_meas','ac_res_etiq'];
    const acExtraCells=acExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="ac-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="ac-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="ac-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const olExtraKeys=['ol_prompt','ol_profess','ol_dead_air','ol_no_vern','ol_avoid_int','ol_cust_verif_meas','ol_ride_cancel','ol_timeliness','ol_res_etiq','ol_comm_quality'];
    const olExtraCells=olExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="ol-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="ol-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="ol-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const ctExtraKeys=['ct_os_out2','ct_profess','ct_verif_other','ct_comm_quality','ct_cust_verif_meas','ct_res_etiq','ct_subjective'];
    const ctExtraCells=ctExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="ct-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="ct-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="ct-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const ycovExtraKeys=['ycov_os_out2','ycov_profess','ycov_verif_other','ycov_comm_quality','ycov_cust_verif_meas','ycov_res_etiq','ycov_subjective'];
    const ycovExtraCells=ycovExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="ycov-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="ycov-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="ycov-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const kelExtraKeys=['kel_os_out2','kel_profess','kel_verif_other','kel_comm_quality','kel_cust_verif_meas','kel_res_etiq'];
    const kelExtraCells=kelExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="kel-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="kel-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="kel-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const vtExtraKeys=['vt_os_out2','vt_profess','vt_verif_other','vt_comm_quality','vt_res_etiq'];
    const vtExtraCells=vtExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="vt-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="vt-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="vt-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const ycdcExtraKeys=['ycdc_os_out2','ycdc_profess','ycdc_verif_other','ycdc_comm_quality','ycdc_res_etiq'];
    const ycdcExtraCells=ycdcExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="ycdc-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="ycdc-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="ycdc-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const blExtraKeys=['bl_os_out2','bl_profess','bl_verif_other','bl_comm_quality','bl_res_etiq'];
    const blExtraCells=blExtraKeys.map(k=>{{
        const v=r[k];
        if(!v||v===null||v==='')return'<td class="bl-extra-col" style="text-align:center"><span style="color:#94A3B8">&mdash;</span></td>';
        if(v==='Not Applicable')return'<td class="bl-extra-col" style="text-align:center;color:#94A3B8">N/A</td>';
        return`<td class="bl-extra-col" style="text-align:center">${{qaYN(v)}}</td>`;
    }});
    const fbTxt=qaEscapeHtml(r.feedback||'—');
    const rawEvalId=QA_AI_ACCOUNTS.has(r._acct)?(r.evaluation_id||r.qa_id):(r.qa_id||r.evaluation_id);
    const evalId=qaEscapeHtml(rawEvalId||'—');
    return`<tr><td style="white-space:nowrap;font-size:10px;color:#475569;max-width:140px;overflow:hidden;text-overflow:ellipsis" title="${{evalId}}">${{evalId}}</td><td style="white-space:nowrap;font-size:11px">${{qaEscapeHtml((r.ts||'—').slice(0,10))}}</td><td style="max-width:200px;overflow:hidden" title="${{qaEscapeHtml(r.agent||'')}}"><div style="display:flex;align-items:center;gap:5px;overflow:hidden"><span style="width:22px;height:22px;border-radius:50%;background:${{av.bg}};color:${{av.tc}};font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${{av.ini}}</span><span style="font-weight:600;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0">${{qaEscapeHtml(r.agent||'—')}}</span></div></td><td style="font-size:11px;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${{qaEscapeHtml(r.supervisor||'')}}">${{qaEscapeHtml(r.supervisor||'—')}}</td><td style="font-size:11px;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${{qaEscapeHtml(r.coach||'')}}">${{qaEscapeHtml(r.coach||'—')}}</td><td>${{acctPill}}</td><td><span class="qa-chip ${{chipCls}}">${{disp}}</span></td>${{critCells.join('')}}${{vipExtraCells.join('')}}${{chExtraCells.join('')}}${{rcExtraCells.join('')}}${{tiExtraCells.join('')}}${{dcExtraCells.join('')}}${{acExtraCells.join('')}}${{olExtraCells.join('')}}${{ctExtraCells.join('')}}${{ycovExtraCells.join('')}}${{kelExtraCells.join('')}}${{vtExtraCells.join('')}}${{ycdcExtraCells.join('')}}${{blExtraCells.join('')}}<td style="font-size:10px;color:#475569;white-space:nowrap">${{qaEscapeHtml(r.invest||'—')}}</td><td style="max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:10px;color:#475569" title="${{fbTxt}}">${{fbTxt}}</td></tr>`;
}}

function qaRenderVisibleRows(){{
    const el=document.getElementById('qa-detail-table');
    const scroller=document.getElementById('qa-tbl-scroll-main');
    if(!el||!scroller) return;
    if(!qaVsRows.length){{
        el.innerHTML=`<tr><td colspan="31" style="text-align:center;color:#94A3B8;padding:20px">No data</td></tr>`;
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
    if(topH>0) html+=`<tr style="height:${{topH}}px;pointer-events:none"><td colspan="31" style="padding:0;border:none"></td></tr>`;
    html+=qaVsRows.slice(startIdx,endIdx).map(qaRowHtml).join('');
    if(botH>0) html+=`<tr style="height:${{botH}}px;pointer-events:none"><td colspan="31" style="padding:0;border:none"></td></tr>`;
    el.innerHTML=html;
    qaUpdateDetailStickyOffsets();
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
    qaApplyDetailColumnWidths();
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

function qaRenderEvalDistLegend(accts, counts, total) {{
    const legEl=document.getElementById('qa-eval-dist-legend');
    if(!legEl)return;
    const isFocus=document.body.classList.contains('qa-dist-focus-mode');
    if(!isFocus){{
        legEl.className='qa-legend-compact';
        const rows=accts.map((a,i)=>({{...a,count:counts[i],pct:total?counts[i]/total*100:0}}))
            .sort((a,b)=>b.count-a.count||a.key.localeCompare(b.key));
        legEl.innerHTML=`<div class="qa-legend-head"><span>Account</span><span style="text-align:right">%</span></div>`+rows.map(r=>{{
            const pct=r.pct.toFixed(1)+'%';
            return`<div class="qa-legend-row"><span class="qa-legend-account"><span class="qa-legend-dot" style="background:${{r.color}}"></span><span class="qa-legend-name">${{qaEscapeHtml(r.key)}}</span></span><span class="qa-legend-pct" style="color:${{r.color}}">${{pct}}</span></div>`;
        }}).join('');
        return;
    }}
    legEl.className='';
    const rawRows=accts.map((a,i)=>({{...a,count:counts[i],pct:total?counts[i]/total*100:0}}))
        .filter(r=>r.count>0)
        .sort((a,b)=>b.count-a.count||a.key.localeCompare(b.key));
    const rows=rawRows.filter(r=>r.pct>=1);
    const otherCount=rawRows.filter(r=>r.pct<1).reduce((s,r)=>s+r.count,0);
    if(otherCount)rows.push({{key:'Other',color:'#94A3B8',count:otherCount,pct:total?otherCount/total*100:0}});
    legEl.innerHTML=`<table class="qa-dist-table">
        <thead><tr><th>Account</th><th>Evaluations</th><th>%</th></tr></thead>
        <tbody>
            ${{rows.map(r=>`<tr>
                <td><span class="qa-dist-account"><span class="qa-dist-dot" style="background:${{r.color}}"></span><span>${{qaEscapeHtml(r.key)}}</span></span></td>
                <td>${{r.count.toLocaleString()}}</td>
                <td>${{r.pct.toFixed(1)}}%</td>
            </tr>`).join('')}}
            <tr class="qa-dist-total"><td>Total</td><td>${{total.toLocaleString()}}</td><td>${{total?'100.0':'0.0'}}%</td></tr>
        </tbody>
    </table>`;
}}

function qaUpdateEvalDist(data) {{
    if(!qaEvalDistChart) return;
    const accts=[
        {{key:'M7',color:'#4F81BD'}},
        {{key:'DMG',color:'#0E7490'}},
        {{key:'R4H',color:'#0F766E'}},
        {{key:'Parentis',color:'#2C3E8C'}},
        {{key:'Britelift',color:'#C0392B'}},
        {{key:'Britelift Chat',color:'#9A3412'}},
        {{key:'RideX',color:'#8E44AD'}},
        {{key:'Hamilton',color:'#065F46'}},
        {{key:'Skyline',color:'#0EA5E9'}},
        {{key:'VIP',color:'#D97706'}},
        {{key:'C&H',color:'#0891B2'}},
        {{key:'Reno Cab',color:'#16A34A'}},
        {{key:'Trans Iowa',color:'#7C3AED'}},
        {{key:'Data Carz',color:'#EA580C'}},
        {{key:'Associated Cab',color:'#0E7490'}},
        {{key:'Ollies',color:'#BE123C'}},
        {{key:'Circle Taxi',color:'#0E7490'}},
        {{key:'YCOV',color:'#047857'}},
        {{key:'Kelowna',color:'#166534'}},
        {{key:'Vermont',color:'#0F766E'}},
        {{key:'YCDC',color:'#155E75'}},
        {{key:'Blueline',color:'#1D4ED8'}}
    ];
    const counts=accts.map(a=>data.filter(r=>r._acct===a.key).length);
    const total=counts.reduce((s,v)=>s+v,0);
    qaEvalDistChart.data.datasets[0].data=counts;
    qaEvalDistChart.update();
    const sub=document.getElementById('qa-donut-sub');
    if(sub) sub.textContent=total+' evaluation'+(total===1?'':'s')+' across all accounts';
    qaRenderEvalDistLegend(accts, counts, total);
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

function qaUpdateVipCrit(data) {{
    const crits=[
        {{key:'os_out',          valId:'qa-vcrit-greet-val', subId:'qa-vcrit-greet-sub'}},
        {{key:'vip_profess',     valId:'qa-vcrit-prof-val',  subId:'qa-vcrit-prof-sub'}},
        {{key:'gen_q',           valId:'qa-vcrit-genq-val',  subId:'qa-vcrit-genq-sub'}},
        {{key:'verif',           valId:'qa-vcrit-verif-val', subId:'qa-vcrit-verif-sub'}},
        {{key:'vip_res_etiq',    valId:'qa-vcrit-resol-val', subId:'qa-vcrit-resol-sub'}},
        {{key:'vip_comm_quality',valId:'qa-vcrit-comm-val',  subId:'qa-vcrit-comm-sub'}},
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

function qaUpdateCHCrit(data) {{
    const crits=[
        {{key:'os_out',        valId:'qa-chcrit-greet-val', subId:'qa-chcrit-greet-sub'}},
        {{key:'ch_profess',    valId:'qa-chcrit-prof-val',  subId:'qa-chcrit-prof-sub'}},
        {{key:'gen_q',         valId:'qa-chcrit-genq-val',  subId:'qa-chcrit-genq-sub'}},
        {{key:'verif',         valId:'qa-chcrit-verif-val', subId:'qa-chcrit-verif-sub'}},
        {{key:'ch_res_etiq',   valId:'qa-chcrit-resol-val', subId:'qa-chcrit-resol-sub'}},
        {{key:'ch_comm_quality',valId:'qa-chcrit-comm-val', subId:'qa-chcrit-comm-sub'}},
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

function qaUpdateRCCrit(data) {{
    const crits=[
        {{key:'os_out',       valId:'qa-rccrit-greet-val', subId:'qa-rccrit-greet-sub'}},
        {{key:'rc_profess',   valId:'qa-rccrit-prof-val',  subId:'qa-rccrit-prof-sub'}},
        {{key:'gen_q',        valId:'qa-rccrit-genq-val',  subId:'qa-rccrit-genq-sub'}},
        {{key:'verif',        valId:'qa-rccrit-verif-val', subId:'qa-rccrit-verif-sub'}},
        {{key:'rc_res_etiq',  valId:'qa-rccrit-resol-val', subId:'qa-rccrit-resol-sub'}},
        {{key:'rc_comm_quality',valId:'qa-rccrit-comm-val',subId:'qa-rccrit-comm-sub'}},
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

function qaUpdateDCCrit(data) {{
    const crits=[
        {{key:'os_out',         valId:'qa-dccrit-greet-val', subId:'qa-dccrit-greet-sub'}},
        {{key:'dc_profess',     valId:'qa-dccrit-prof-val',  subId:'qa-dccrit-prof-sub'}},
        {{key:'gen_q',          valId:'qa-dccrit-genq-val',  subId:'qa-dccrit-genq-sub'}},
        {{key:'verif',          valId:'qa-dccrit-verif-val', subId:'qa-dccrit-verif-sub'}},
        {{key:'dc_res_etiq',    valId:'qa-dccrit-resol-val', subId:'qa-dccrit-resol-sub'}},
        {{key:'dc_comm_quality',valId:'qa-dccrit-comm-val',  subId:'qa-dccrit-comm-sub'}},
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

function qaUpdateTICrit(data) {{
    const crits=[
        {{key:'os_out',         valId:'qa-ticrit-greet-val', subId:'qa-ticrit-greet-sub'}},
        {{key:'ti_profess',     valId:'qa-ticrit-prof-val',  subId:'qa-ticrit-prof-sub'}},
        {{key:'gen_q',          valId:'qa-ticrit-genq-val',  subId:'qa-ticrit-genq-sub'}},
        {{key:'verif',          valId:'qa-ticrit-verif-val', subId:'qa-ticrit-verif-sub'}},
        {{key:'ti_res_etiq',    valId:'qa-ticrit-resol-val', subId:'qa-ticrit-resol-sub'}},
        {{key:'ti_comm_quality',valId:'qa-ticrit-comm-val',  subId:'qa-ticrit-comm-sub'}},
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

function qaUpdateACCrit(data) {{
    const crits=[
        {{key:'os_out',             valId:'qa-accrit-greet-val', subId:'qa-accrit-greet-sub'}},
        {{key:'ac_profess',         valId:'qa-accrit-prof-val',  subId:'qa-accrit-prof-sub'}},
        {{key:'gen_q',              valId:'qa-accrit-genq-val',  subId:'qa-accrit-genq-sub'}},
        {{key:'verif',              valId:'qa-accrit-verif-val', subId:'qa-accrit-verif-sub'}},
        {{key:'ac_res_etiq',        valId:'qa-accrit-resol-val', subId:'qa-accrit-resol-sub'}},
        {{key:'ac_comm_quality',    valId:'qa-accrit-comm-val',  subId:'qa-accrit-comm-sub'}},
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

function qaUpdateOLCrit(data) {{
    const crits=[
        {{key:'os_out',          valId:'qa-olcrit-greet-val', subId:'qa-olcrit-greet-sub'}},
        {{key:'ol_profess',      valId:'qa-olcrit-prof-val',  subId:'qa-olcrit-prof-sub'}},
        {{key:'gen_q',           valId:'qa-olcrit-genq-val',  subId:'qa-olcrit-genq-sub'}},
        {{key:'verif',           valId:'qa-olcrit-verif-val', subId:'qa-olcrit-verif-sub'}},
        {{key:'ol_res_etiq',     valId:'qa-olcrit-resol-val', subId:'qa-olcrit-resol-sub'}},
        {{key:'ol_comm_quality', valId:'qa-olcrit-comm-val',  subId:'qa-olcrit-comm-sub'}},
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

function qaBuildAiQeTrend(data) {{
    const weeks={{}};
    data.forEach(r=>{{
        const k=r.week_start; if(!k) return;
        if(!weeks[k]) weeks[k]={{aiScores:[],qeScores:[]}};
        const ai=Number(r.score_ai);
        const human=Number(r.score_human);
        const score=Number(r.score);
        const hasAi=!isNaN(ai)&&ai>0;
        const hasHuman=!isNaN(human)&&human>0;
        if(hasAi) weeks[k].aiScores.push(ai);
        if(hasHuman) weeks[k].qeScores.push(human);
        else if(!hasAi&&!isNaN(score)&&score>0) weeks[k].qeScores.push(score);
    }});
    return Object.keys(weeks).sort().map(k=>{{
        const d=new Date(k+'T00:00:00');
        const lbl=QA_MONTHS[d.getMonth()]+' '+String(d.getDate()).padStart(2,'0')+', '+d.getFullYear();
        const aiScores=weeks[k].aiScores, qeScores=weeks[k].qeScores;
        return {{
            week:lbl,
            aiAvg:aiScores.length?parseFloat(qaAvg(aiScores).toFixed(1)):null,
            qeAvg:qeScores.length?parseFloat(qaAvg(qeScores).toFixed(1)):null,
            aiCount:aiScores.length,
            qeCount:qeScores.length,
        }};
    }});
}}

function qaGetDateAccountFilteredData() {{
    const startStr=qaFmtDate(qaDrpStart), endStr=qaFmtDate(qaDrpEnd);
    return qaGetActiveData().filter(r=>{{
        const rDate=(r.ts||'').slice(0,10);
        return rDate>=startStr&&rDate<=endStr;
    }});
}}

function qaUpdateCTCrit(data) {{
    const crits=[
        {{key:'os_out',          valId:'qa-ctcrit-greet-val', subId:'qa-ctcrit-greet-sub'}},
        {{key:'ct_profess',      valId:'qa-ctcrit-prof-val',  subId:'qa-ctcrit-prof-sub'}},
        {{key:'gen_q',           valId:'qa-ctcrit-genq-val',  subId:'qa-ctcrit-genq-sub'}},
        {{key:'verif',           valId:'qa-ctcrit-verif-val', subId:'qa-ctcrit-verif-sub'}},
        {{key:'ct_res_etiq',     valId:'qa-ctcrit-resol-val', subId:'qa-ctcrit-resol-sub'}},
        {{key:'ct_comm_quality', valId:'qa-ctcrit-comm-val',  subId:'qa-ctcrit-comm-sub'}},
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

function qaUpdateYCOVCrit(data) {{
    const crits=[
        {{key:'os_out',            valId:'qa-ycovcrit-greet-val', subId:'qa-ycovcrit-greet-sub'}},
        {{key:'ycov_profess',      valId:'qa-ycovcrit-prof-val',  subId:'qa-ycovcrit-prof-sub'}},
        {{key:'gen_q',             valId:'qa-ycovcrit-genq-val',  subId:'qa-ycovcrit-genq-sub'}},
        {{key:'verif',             valId:'qa-ycovcrit-verif-val', subId:'qa-ycovcrit-verif-sub'}},
        {{key:'ycov_res_etiq',     valId:'qa-ycovcrit-resol-val', subId:'qa-ycovcrit-resol-sub'}},
        {{key:'ycov_comm_quality', valId:'qa-ycovcrit-comm-val',  subId:'qa-ycovcrit-comm-sub'}},
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

function qaUpdateKELCrit(data) {{
    const crits=[
        {{key:'os_out',           valId:'qa-kelcrit-greet-val', subId:'qa-kelcrit-greet-sub'}},
        {{key:'kel_profess',      valId:'qa-kelcrit-prof-val',  subId:'qa-kelcrit-prof-sub'}},
        {{key:'gen_q',            valId:'qa-kelcrit-genq-val',  subId:'qa-kelcrit-genq-sub'}},
        {{key:'verif',            valId:'qa-kelcrit-verif-val', subId:'qa-kelcrit-verif-sub'}},
        {{key:'kel_res_etiq',     valId:'qa-kelcrit-resol-val', subId:'qa-kelcrit-resol-sub'}},
        {{key:'kel_comm_quality', valId:'qa-kelcrit-comm-val',  subId:'qa-kelcrit-comm-sub'}},
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

function qaUpdateVTCrit(data) {{
    const crits=[
        {{key:'os_out',          valId:'qa-vtcrit-greet-val', subId:'qa-vtcrit-greet-sub'}},
        {{key:'vt_profess',      valId:'qa-vtcrit-prof-val',  subId:'qa-vtcrit-prof-sub'}},
        {{key:'gen_q',           valId:'qa-vtcrit-genq-val',  subId:'qa-vtcrit-genq-sub'}},
        {{key:'verif',           valId:'qa-vtcrit-verif-val', subId:'qa-vtcrit-verif-sub'}},
        {{key:'vt_res_etiq',     valId:'qa-vtcrit-resol-val', subId:'qa-vtcrit-resol-sub'}},
        {{key:'vt_comm_quality', valId:'qa-vtcrit-comm-val',  subId:'qa-vtcrit-comm-sub'}},
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

function qaUpdateYCDCCrit(data) {{
    const crits=[
        {{key:'os_out',            valId:'qa-ycdccrit-greet-val', subId:'qa-ycdccrit-greet-sub'}},
        {{key:'ycdc_profess',      valId:'qa-ycdccrit-prof-val',  subId:'qa-ycdccrit-prof-sub'}},
        {{key:'gen_q',             valId:'qa-ycdccrit-genq-val',  subId:'qa-ycdccrit-genq-sub'}},
        {{key:'verif',             valId:'qa-ycdccrit-verif-val', subId:'qa-ycdccrit-verif-sub'}},
        {{key:'ycdc_res_etiq',     valId:'qa-ycdccrit-resol-val', subId:'qa-ycdccrit-resol-sub'}},
        {{key:'ycdc_comm_quality', valId:'qa-ycdccrit-comm-val',  subId:'qa-ycdccrit-comm-sub'}},
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

function qaUpdateBLCrit(data) {{
    const crits=[
        {{key:'os_out',          valId:'qa-blcrit-greet-val', subId:'qa-blcrit-greet-sub'}},
        {{key:'bl_profess',      valId:'qa-blcrit-prof-val',  subId:'qa-blcrit-prof-sub'}},
        {{key:'gen_q',           valId:'qa-blcrit-genq-val',  subId:'qa-blcrit-genq-sub'}},
        {{key:'verif',           valId:'qa-blcrit-verif-val', subId:'qa-blcrit-verif-sub'}},
        {{key:'bl_res_etiq',     valId:'qa-blcrit-resol-val', subId:'qa-blcrit-resol-sub'}},
        {{key:'bl_comm_quality', valId:'qa-blcrit-comm-val',  subId:'qa-blcrit-comm-sub'}},
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
    qaUpdateAiQeTrend(qaGetDateAccountFilteredData());
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
    const aivhAccts=['hamilton','skyline','vip','ch','rc','ti','dc','ac','ol','ct','ycov','kel','vt','ycdc','bl'];
    if(aivhCard)aivhCard.style.display=(aivhAccts.includes(acct))?'':'none';
    if(aivhAccts.includes(acct))qaUpdateAivh(filtered);
    const hCritEl=document.getElementById('qa-hamilton-crit');
    if(hCritEl)hCritEl.style.display=(acct==='hamilton')?'':'none';
    if(acct==='hamilton')qaUpdateHamiltonCrit(filtered);
    const sCritEl=document.getElementById('qa-skyline-crit');
    if(sCritEl)sCritEl.style.display=(acct==='skyline')?'':'none';
    if(acct==='skyline')qaUpdateSkylineCrit(filtered);
    const vCritEl=document.getElementById('qa-vip-crit');
    if(vCritEl)vCritEl.style.display=(acct==='vip')?'':'none';
    if(acct==='vip')qaUpdateVipCrit(filtered);
    const chCritEl=document.getElementById('qa-ch-crit');
    if(chCritEl)chCritEl.style.display=(acct==='ch')?'':'none';
    if(acct==='ch')qaUpdateCHCrit(filtered);
    const rcCritEl=document.getElementById('qa-rc-crit');
    if(rcCritEl)rcCritEl.style.display=(acct==='rc')?'':'none';
    if(acct==='rc')qaUpdateRCCrit(filtered);
    const tiCritEl=document.getElementById('qa-ti-crit');
    if(tiCritEl)tiCritEl.style.display=(acct==='ti')?'':'none';
    if(acct==='ti')qaUpdateTICrit(filtered);
    const dcCritEl=document.getElementById('qa-dc-crit');
    if(dcCritEl)dcCritEl.style.display=(acct==='dc')?'':'none';
    if(acct==='dc')qaUpdateDCCrit(filtered);
    const acCritEl=document.getElementById('qa-ac-crit');
    if(acCritEl)acCritEl.style.display=(acct==='ac')?'':'none';
    if(acct==='ac')qaUpdateACCrit(filtered);
    const olCritEl=document.getElementById('qa-ol-crit');
    if(olCritEl)olCritEl.style.display=(acct==='ol')?'':'none';
    if(acct==='ol')qaUpdateOLCrit(filtered);
    const ctCritEl=document.getElementById('qa-ct-crit');
    if(ctCritEl)ctCritEl.style.display=(acct==='ct')?'':'none';
    if(acct==='ct')qaUpdateCTCrit(filtered);
    const ycovCritEl=document.getElementById('qa-ycov-crit');
    if(ycovCritEl)ycovCritEl.style.display=(acct==='ycov')?'':'none';
    if(acct==='ycov')qaUpdateYCOVCrit(filtered);
    const kelCritEl=document.getElementById('qa-kel-crit');
    if(kelCritEl)kelCritEl.style.display=(acct==='kel')?'':'none';
    if(acct==='kel')qaUpdateKELCrit(filtered);
    const vtCritEl=document.getElementById('qa-vt-crit');
    if(vtCritEl)vtCritEl.style.display=(acct==='vt')?'':'none';
    if(acct==='vt')qaUpdateVTCrit(filtered);
    const ycdcCritEl=document.getElementById('qa-ycdc-crit');
    if(ycdcCritEl)ycdcCritEl.style.display=(acct==='ycdc')?'':'none';
    if(acct==='ycdc')qaUpdateYCDCCrit(filtered);
    const blCritEl=document.getElementById('qa-bl-crit');
    if(blCritEl)blCritEl.style.display=(acct==='bl')?'':'none';
    if(acct==='bl')qaUpdateBLCrit(filtered);
    const sumStripMain=document.getElementById('qa-sum-strip-main');
    if(sumStripMain)sumStripMain.style.display=(acct==='hamilton'||acct==='skyline'||acct==='vip'||acct==='ch'||acct==='rc'||acct==='ti'||acct==='dc'||acct==='ac'||acct==='ol'||acct==='ct'||acct==='ycov'||acct==='kel'||acct==='vt'||acct==='ycdc'||acct==='bl')?'none':'';
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
        qaInitResizableColumns();
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
    const qaTrendHoverPlugin={{
        id:'qaTrendHoverFocus',
        afterEvent(chart,args){{
            if(!['qa-trend-chart','qa-aiqe-trend-chart'].includes(chart.canvas?.id))return;
            const e=args.event;
            if(!e)return;
            const area=chart.chartArea;
            const clear=()=>{{
                if(chart.$qaHoverIndex!=null){{
                    chart.$qaHoverIndex=null;
                    args.changed=true;
                }}
            }};
            if(e.type==='mouseout'){{
                clear();
                return;
            }}
            if(!['mousemove','click','touchstart','touchmove'].includes(e.type))return;
            if(e.x<area.left||e.x>area.right||e.y<area.top||e.y>area.bottom){{
                if(e.type==='click')clear();
                return;
            }}
            const points=chart.getElementsAtEventForMode(e.native||e,'index',{{intersect:false}},false);
            let idx=points?.[0]?.index;
            if(idx==null){{
                const meta=chart.getDatasetMeta(0);
                let best=null,bestDx=Infinity;
                meta.data.forEach((pt,i)=>{{
                    const dx=Math.abs(pt.x-e.x);
                    if(dx<bestDx){{bestDx=dx;best=i;}}
                }});
                idx=best;
            }}
            if(idx==null)return;
            chart.$qaHoverIndex=idx;
            chart.$qaHoverDataset=points?.[0]?.datasetIndex??0;
            args.changed=true;
        }},
        afterDatasetsDraw(chart){{
            if(!['qa-trend-chart','qa-aiqe-trend-chart'].includes(chart.canvas?.id))return;
            const idx=chart.$qaHoverIndex;
            if(idx==null)return;
            const ctx2=chart.ctx,area=chart.chartArea;
            const firstMeta=chart.getDatasetMeta(0);
            const anchor=firstMeta?.data?.[idx];
            if(!anchor)return;
            const x=anchor.x;
            const clamp=(v,min,max)=>Math.max(min,Math.min(max,v));
            const roundedRect=(x,y,w,h,r)=>{{
                ctx2.beginPath();
                if(ctx2.roundRect)ctx2.roundRect(x,y,w,h,r);
                else ctx2.rect(x,y,w,h);
            }};
            const asPct=v=>v==null||Number.isNaN(Number(v))?'—':Number(v).toFixed(1)+'%';
            const asNum=v=>v==null||Number.isNaN(Number(v))?'—':Number(v).toLocaleString();
            const drawPoint=(pt,color,r)=>{{
                if(!pt)return;
                ctx2.save();
                ctx2.shadowColor=color;
                ctx2.shadowBlur=8;
                ctx2.fillStyle=color;
                ctx2.strokeStyle='#fff';
                ctx2.lineWidth=2;
                ctx2.beginPath();
                ctx2.arc(pt.x,pt.y,r,0,Math.PI*2);
                ctx2.fill();
                ctx2.shadowBlur=0;
                ctx2.stroke();
                ctx2.restore();
            }};
            const drawBar=(bar,ds)=>{{
                if(!bar||bar.width==null||bar.base==null)return;
                const left=bar.x-bar.width/2;
                const top=Math.min(bar.y,bar.base);
                const h=Math.abs(bar.base-bar.y);
                ctx2.save();
                ctx2.fillStyle=ds.backgroundColor||'rgba(15,23,42,.25)';
                ctx2.strokeStyle=ds.borderColor||'#0F172A';
                ctx2.lineWidth=1.5;
                ctx2.fillRect(left,top,bar.width,h);
                ctx2.strokeRect(left,top,bar.width,h);
                ctx2.restore();
            }};
            const drawLineSegment=(meta,ds,color)=>{{
                const pts=[idx-1,idx,idx+1].map(i=>meta.data[i]).filter(Boolean);
                if(pts.length<2)return;
                ctx2.save();
                ctx2.strokeStyle=color;
                ctx2.lineWidth=(ds.borderWidth||1.5)+0.6;
                ctx2.lineJoin='round';
                ctx2.lineCap='round';
                ctx2.beginPath();
                pts.forEach((pt,i)=>{{if(i===0)ctx2.moveTo(pt.x,pt.y);else ctx2.lineTo(pt.x,pt.y);}});
                ctx2.stroke();
                ctx2.restore();
            }};
            ctx2.save();
            ctx2.fillStyle='rgba(255,255,255,0.68)';
            ctx2.fillRect(area.left,area.top,area.right-area.left,area.bottom-area.top);
            ctx2.strokeStyle='rgba(13,59,110,0.38)';
            ctx2.lineWidth=1;
            ctx2.setLineDash([4,4]);
            ctx2.beginPath();
            ctx2.moveTo(x,area.top);
            ctx2.lineTo(x,area.bottom);
            ctx2.stroke();
            ctx2.setLineDash([]);

            if(chart.canvas.id==='qa-trend-chart'){{
                const avgDs=chart.data.datasets[0], barDs=chart.data.datasets[2];
                const avgMeta=chart.getDatasetMeta(0), barMeta=chart.getDatasetMeta(2);
                drawLineSegment(avgMeta,avgDs,'#0D3B6E');
                drawBar(barMeta.data[idx],barDs);
                drawPoint(avgMeta.data[idx],'#0D3B6E',6);
                const avg=avgDs.data[idx];
                const prev=idx>0?avgDs.data[idx-1]:null;
                const wow=prev==null||avg==null||Number.isNaN(Number(prev))||Number.isNaN(Number(avg))?'—':((avg-prev)>=0?'+':'')+(avg-prev).toFixed(1)+' pts';
                const lines=[
                    {{label:chart.data.labels[idx]||'—',value:'',strong:true}},
                    {{label:'Overall QA Score',value:asPct(avg)}},
                    {{label:'Entry Count',value:asNum(barDs.data[idx])}},
                    {{label:'Week-over-Week Change',value:wow}},
                ];
                drawTooltip(lines);
            }} else {{
                const aiDs=chart.data.datasets[0],qeDs=chart.data.datasets[1],aiBarDs=chart.data.datasets[3],qeBarDs=chart.data.datasets[4];
                const aiMeta=chart.getDatasetMeta(0),qeMeta=chart.getDatasetMeta(1),aiBarMeta=chart.getDatasetMeta(3),qeBarMeta=chart.getDatasetMeta(4);
                drawLineSegment(aiMeta,aiDs,'#004C97');
                drawLineSegment(qeMeta,qeDs,'#39B54A');
                drawBar(aiBarMeta.data[idx],aiBarDs);
                drawBar(qeBarMeta.data[idx],qeBarDs);
                drawPoint(aiMeta.data[idx],'#004C97',5.5);
                drawPoint(qeMeta.data[idx],'#39B54A',5.5);
                const ai=aiDs.data[idx],qe=qeDs.data[idx];
                const gap=ai==null||qe==null||Number.isNaN(Number(ai))||Number.isNaN(Number(qe))?'—':(qe-ai).toFixed(1)+' pts';
                const lines=[
                    {{label:chart.data.labels[idx]||'—',value:'',strong:true}},
                    {{label:'AI Avg (%)',value:asPct(ai)}},
                    {{label:'QE Avg (%)',value:asPct(qe)}},
                    {{label:'Gap (pts)',value:gap}},
                    {{label:'AI Entries',value:asNum(aiBarDs.data[idx])}},
                    {{label:'QE Entries',value:asNum(qeBarDs.data[idx])}},
                ];
                drawTooltip(lines);
            }}
            ctx2.restore();

            function drawTooltip(lines){{
                const pad=9,rowH=16,titleH=18;
                ctx2.save();
                ctx2.font='700 11px sans-serif';
                const labelW=Math.max(...lines.map(l=>ctx2.measureText(l.label).width));
                ctx2.font='800 11px sans-serif';
                const valueW=Math.max(...lines.map(l=>ctx2.measureText(l.value||'').width));
                const w=Math.max(190,labelW+valueW+pad*2+18);
                const h=pad*2+titleH+(lines.length-1)*rowH;
                let tx=clamp(x+14,area.left+4,area.right-w-4);
                if(x>area.right-w-24)tx=clamp(x-w-14,area.left+4,area.right-w-4);
                const ty=area.top+8;
                ctx2.shadowColor='rgba(15,23,42,0.18)';
                ctx2.shadowBlur=12;
                ctx2.fillStyle='rgba(15,23,42,0.94)';
                roundedRect(tx,ty,w,h,8);
                ctx2.fill();
                ctx2.shadowBlur=0;
                ctx2.strokeStyle='rgba(255,255,255,0.16)';
                ctx2.stroke();
                let cy=ty+pad+9;
                ctx2.fillStyle='#F8FAFC';
                ctx2.font='800 12px sans-serif';
                ctx2.textAlign='left';
                ctx2.textBaseline='middle';
                ctx2.fillText(lines[0].label,tx+pad,cy);
                cy+=titleH;
                ctx2.font='600 11px sans-serif';
                lines.slice(1).forEach(l=>{{
                    ctx2.fillStyle='rgba(226,232,240,0.86)';
                    ctx2.textAlign='left';
                    ctx2.fillText(l.label,tx+pad,cy);
                    ctx2.fillStyle='#fff';
                    ctx2.textAlign='right';
                    ctx2.fillText(l.value,tx+w-pad,cy);
                    cy+=rowH;
                }});
                ctx2.restore();
            }}
        }}
    }};
    const aiQeTrendLabelPlugin={{
        id:'qaAiQeTrendLabels',
        afterDatasetsDraw(chart){{
            if(chart.canvas?.id!=='qa-aiqe-trend-chart')return;
            const ctx2=chart.ctx;
            ctx2.save();
            const isFocus=document.body.classList.contains('qa-aiqe-focus-mode');
            const pointFont=isFocus?13:9;
            const trendFont=isFocus?14:10;
            const area=chart.chartArea;
            const clamp=(v,min,max)=>Math.max(min,Math.min(max,v));
            const drawSoftLabel=(text,x,y,color,align='center')=>{{
                ctx2.save();
                ctx2.font=`700 ${{pointFont}}px sans-serif`;
                ctx2.textAlign=align;
                ctx2.textBaseline='middle';
                ctx2.lineWidth=3;
                ctx2.strokeStyle='rgba(255,255,255,0.92)';
                ctx2.strokeText(text,x,y);
                ctx2.fillStyle=color;
                ctx2.fillText(text,x,y);
                ctx2.restore();
            }};
            [
                {{idx:3,color:'#004C97',inside:true}},
                {{idx:4,color:'#39B54A',inside:false}},
            ].forEach(cfg=>{{
                const ds=chart.data.datasets[cfg.idx],meta=chart.getDatasetMeta(cfg.idx);
                if(!ds||!meta)return;
                ctx2.font=`700 ${{pointFont}}px sans-serif`;
                ctx2.fillStyle=cfg.color;
                ctx2.textAlign='center';
                ctx2.textBaseline=cfg.inside?'top':'bottom';
                meta.data.forEach((bar,i)=>{{
                    const v=ds.data[i];if(!v)return;
                    const y=cfg.inside?bar.y+4:bar.y-4;
                    ctx2.fillText(v,bar.x,y);
                }});
            }});
            const aiDs=chart.data.datasets[0], qeDs=chart.data.datasets[1];
            const aiMeta=chart.getDatasetMeta(0), qeMeta=chart.getDatasetMeta(1);
            const gaps=[];
            aiMeta.data.forEach((aiPt,i)=>{{
                const qePt=qeMeta.data[i], ai=aiDs?.data[i], qe=qeDs?.data[i];
                if(!aiPt||!qePt||ai==null||qe==null||Number.isNaN(ai)||Number.isNaN(qe))return;
                const gap=qe-ai;
                gaps.push(gap);
                const close=Math.abs(aiPt.y-qePt.y)<24;
                let aiY=aiPt.y+12, qeY=qePt.y-9;
                if(close){{
                    if(aiPt.y<=qePt.y){{
                        aiY=aiPt.y-10;
                        qeY=qePt.y+12;
                    }} else {{
                        aiY=aiPt.y+12;
                        qeY=qePt.y-10;
                    }}
                }}
                aiY=clamp(aiY,area.top+8,area.bottom-8);
                qeY=clamp(qeY,area.top+8,area.bottom-8);
                if(Math.abs(aiY-qeY)<12){{
                    const mid=(aiY+qeY)/2;
                    aiY=clamp(mid+7,area.top+8,area.bottom-8);
                    qeY=clamp(mid-7,area.top+8,area.bottom-8);
                }}
                drawSoftLabel(ai.toFixed(1)+'%',aiPt.x,aiY,'#004C97');
                drawSoftLabel(qe.toFixed(1)+'%',qePt.x,qeY,'#39B54A');

                let gapX=aiPt.x, gapY=(aiPt.y+qePt.y)/2;
                if(close){{
                    gapX=clamp(aiPt.x+18,area.left+18,area.right-18);
                    gapY=clamp((aiY+qeY)/2,area.top+10,area.bottom-10);
                }}
                drawSoftLabel(gap.toFixed(1)+' pts',gapX,gapY,'#C2410C',close?'left':'center');
            }});
            const validGaps=gaps.filter(g=>g!=null&&!Number.isNaN(g));
            if(validGaps.length>=2){{
                const prev=Math.abs(validGaps[validGaps.length-2]), last=Math.abs(validGaps[validGaps.length-1]);
                const narrowing=last<=prev;
                const text='Gap Trend: '+(narrowing?'Narrowing ↓':'Widening ↑');
                ctx2.font=`700 ${{trendFont}}px sans-serif`;
                const pad=6, w=ctx2.measureText(text).width+pad*2, h=19;
                const x=area.right-w-4, y=Math.max(2,area.top-40);
                ctx2.fillStyle='rgba(255,247,237,0.94)';
                ctx2.strokeStyle='rgba(194,65,12,0.35)';
                ctx2.lineWidth=1;
                ctx2.beginPath();
                if(ctx2.roundRect){{
                    ctx2.roundRect(x,y,w,h,6);
                }} else {{
                    ctx2.rect(x,y,w,h);
                }}
                ctx2.fill();
                ctx2.stroke();
                ctx2.fillStyle='#C2410C';
                ctx2.textAlign='left';
                ctx2.textBaseline='middle';
                ctx2.fillText(text,x+pad,y+h/2);
            }}
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
    const qaAcctsLoaded=[qaRawData,dmgRawData,r4hRawData,parentisRawData,briteliftRawData,blcRawData,ridexRawData,hamiltonRawData,skylineRawData,vipRawData,chRawData,rcRawData,tiRawData,dcRawData,acRawData,olRawData,ctRawData,ycovRawData,kelRawData,vtRawData,ycdcRawData,blRawData].filter(d=>d.length>0).length;
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
            plugins:[trendLabelPlugin,barLabelPlugin,qaTrendHoverPlugin],
            data:{{labels:[],datasets:[
                {{label:'Avg QA Score',data:[],borderColor:'#0D3B6E',backgroundColor:'rgba(13,59,110,0.08)',tension:0.3,fill:true,pointRadius:4,pointHoverRadius:6,pointBackgroundColor:'#0D3B6E',yAxisID:'y',order:0}},
                {{label:'Target (85%)',data:[],borderColor:'#E85D3F',borderDash:[5,4],borderWidth:1.5,pointRadius:0,fill:false,yAxisID:'y',order:0}},
                {{type:'bar',label:'Total Evaluations',data:[],backgroundColor:'rgba(57,181,74,0.55)',borderColor:'#39B54A',borderWidth:1,yAxisID:'y1',barPercentage:0.6,categoryPercentage:0.7,order:1}}
            ]}},
            options:{{
                responsive:true,maintainAspectRatio:false,
                interaction:{{mode:'index',intersect:false}},
                animation:{{duration:200,easing:'easeOutQuart'}},
                layout:{{padding:{{top:24,bottom:0}}}},
                plugins:{{legend:{{display:true,position:'bottom',labels:{{font:{{size:10}},boxWidth:12}}}},tooltip:{{enabled:false,callbacks:{{label:ctx=>ctx.datasetIndex===2?ctx.parsed.y+' evals':ctx.parsed.y.toFixed(1)+'%'}}}}}},
                scales:{{
                    y:{{min:80,max:100,ticks:{{display:false}},border:{{display:false}},grid:{{color:'#F1F5F9'}}}},
                    y1:{{display:false,position:'right',beginAtZero:true,grid:{{display:false}}}},
                    x:{{ticks:{{font:{{size:9}},maxRotation:30}},grid:{{display:false}}}}
                }}
            }}
        }});
    }}
    const aiQeTrendCtx=document.getElementById('qa-aiqe-trend-chart');
    if(aiQeTrendCtx){{
        qaAiQeTrendChart=new Chart(aiQeTrendCtx,{{
            type:'line',
            plugins:[aiQeTrendLabelPlugin,qaTrendHoverPlugin],
            data:{{labels:[],datasets:[
                {{label:'AI avg',data:[],borderColor:'#004C97',borderWidth:1.5,backgroundColor:'rgba(0,76,151,0.08)',tension:0.3,fill:false,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:'#004C97',yAxisID:'y',order:0}},
                {{label:'QE avg',data:[],borderColor:'#39B54A',borderWidth:1.5,backgroundColor:'rgba(57,181,74,0.08)',tension:0.3,fill:false,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:'#39B54A',yAxisID:'y',order:0}},
                {{label:'Gap Area',data:[],borderColor:'rgba(249,115,22,0)',backgroundColor:'rgba(249,115,22,0.20)',tension:0.3,fill:{{target:0}},pointRadius:0,pointHoverRadius:0,yAxisID:'y',order:1}},
                {{type:'bar',label:'AI entries',data:[],backgroundColor:'rgba(0,76,151,0.32)',borderColor:'#004C97',borderWidth:1,yAxisID:'y1',barPercentage:1.0,categoryPercentage:0.72,order:2,minBarLength:10}},
                {{type:'bar',label:'QE entries',data:[],backgroundColor:'rgba(57,181,74,0.38)',borderColor:'#39B54A',borderWidth:1,yAxisID:'y1',barPercentage:1.0,categoryPercentage:0.72,order:2,minBarLength:12}},
            ]}},
            options:{{
                responsive:true,maintainAspectRatio:false,
                interaction:{{mode:'index',intersect:false}},
                animation:{{duration:200,easing:'easeOutQuart'}},
                layout:{{padding:{{top:44,bottom:0}}}},
                plugins:{{legend:{{display:true,position:'bottom',labels:{{font:{{size:9}},boxWidth:10}}}},tooltip:{{enabled:false,callbacks:{{label:ctx=>{{
                    if(ctx.dataset.label==='Gap Area'){{
                        const ai=ctx.chart.data.datasets[0].data[ctx.dataIndex], qe=ctx.chart.data.datasets[1].data[ctx.dataIndex];
                        return ai==null||qe==null?'Gap Area: No data':'Gap Area: '+(qe-ai).toFixed(1)+' pts';
                    }}
                    return ctx.dataset.yAxisID==='y1'?ctx.dataset.label+': '+ctx.parsed.y+' entries':ctx.dataset.label+': '+(ctx.parsed.y==null?'No data':ctx.parsed.y.toFixed(1)+'%');
                }}}}}}}},
                scales:{{
                    y:{{min:75,max:100,ticks:{{display:false}},border:{{display:false}},grid:{{color:'#F1F5F9'}}}},
                    y1:{{display:false,position:'right',beginAtZero:true,grid:{{display:false}}}},
                    x:{{ticks:{{font:{{size:9}},maxRotation:30}},grid:{{display:false}}}}
                }}
            }}
        }});
    }}
    if(!window.qaTrendHoverDismissBound){{
        window.qaTrendHoverDismissBound=true;
        document.addEventListener('click',ev=>{{
            if(ev.target?.closest?.('#qa-trend-chart,#qa-aiqe-trend-chart'))return;
            [qaTrendChart,qaAiQeTrendChart].forEach(ch=>{{
                if(ch&&ch.$qaHoverIndex!=null){{
                    ch.$qaHoverIndex=null;
                    ch.update('none');
                }}
            }});
        }},{{passive:true}});
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
            const isFocus=document.body.classList.contains('qa-dist-focus-mode');
            const donutLabelFont=isFocus?15:10;
            const donutValueFont=isFocus?16:11;
            const CX=validArcs[0].x, CY=validArcs[0].y, OR=validArcs[0].outerRadius;
            const area=chart.chartArea;
            ctx2.save();

            if(isFocus){{
                ctx2.textAlign='center';
                ctx2.textBaseline='middle';
                ctx2.fillStyle='#0F172A';
                ctx2.font='900 32px sans-serif';
                ctx2.fillText(total.toLocaleString(),CX,CY-8);
                ctx2.fillStyle='#334155';
                ctx2.font='600 17px sans-serif';
                ctx2.fillText('Evaluations',CX,CY+22);
            }}

            // Pass 1 — count inside arc (skip arcs too thin to read)
            meta.data.forEach((arc,i)=>{{
                const v=ds.data[i];if(!v)return;
                if(arc.endAngle-arc.startAngle<0.3)return;
                const mid=(arc.startAngle+arc.endAngle)/2;
                const r=(arc.innerRadius+arc.outerRadius)/2;
                ctx2.font=`bold ${{donutValueFont}}px sans-serif`;ctx2.fillStyle='#fff';
                ctx2.textAlign='center';ctx2.textBaseline='middle';
                ctx2.fillText(v,CX+Math.cos(mid)*r,CY+Math.sin(mid)*r);
            }});

            // Pass 2 — collect outside-label candidates (>= 2 % of total)
            const lbs=[];
            meta.data.forEach((arc,i)=>{{
                const v=ds.data[i];if(!v||v/total<0.02)return;
                const mid=(arc.startAngle+arc.endAngle)/2;
                const cm=Math.cos(mid),sm=Math.sin(mid);
                const pct=total?(v/total*100).toFixed(1)+'%':'0.0%';
                lbs.push({{i,v,mid,cm,sm,
                    label:chart.data.labels[i],pct,color:ds.backgroundColor[i],
                    right:cm>=0,y:CY+sm*(OR+(isFocus?28:18))}});
            }});

            // Fixed columns — always clear of the arc ring
            const COL_OFF=isFocus?40:24, TICK=isFocus?14:9;
            const rColX=CX+OR+COL_OFF, lColX=CX-OR-COL_OFF;

            // Pass 3 — deterministic spread: centre labels around their idealY centroid,
            // evenly spaced by MIN_GAP, then clamp the whole group inside the chart area.
            const MIN_GAP=isFocus?28:17;
            for(const isRight of [true,false]){{
                const grp=lbs.filter(l=>l.right===isRight).sort((a,b)=>a.y-b.y);
                if(!grp.length)continue;
                const n=grp.length;
                const span=(n-1)*MIN_GAP;
                const centY=grp.reduce((s,l)=>s+l.y,0)/n;
                let startY=centY-span/2;
                const pad=isFocus?18:6;
                startY=Math.max(area.top+pad,Math.min(area.bottom-pad-span,startY));
                grp.forEach((lb,i)=>{{lb.y=startY+i*MIN_GAP;}});
            }}

            // Pass 4 — draw leader line and label
            lbs.forEach(lb=>{{
                const dir=lb.right?1:-1;
                const colX=lb.right?rColX:lColX;
                const tickEndX=colX+dir*TICK;
                const ax=CX+lb.cm*OR, ay=CY+lb.sm*OR;   // arc surface
                const sx=CX+lb.cm*(OR+(isFocus?8:5)), sy=CY+lb.sm*(OR+(isFocus?8:5)); // radial stub

                ctx2.beginPath();
                ctx2.moveTo(ax,ay);
                ctx2.lineTo(sx,sy);         // radial stub
                ctx2.lineTo(colX,lb.y);     // diagonal to column at resolved y
                ctx2.lineTo(tickEndX,lb.y); // horizontal tick
                ctx2.strokeStyle=lb.color;ctx2.lineWidth=isFocus?1.8:1.5;ctx2.stroke();

                ctx2.font=`bold ${{donutLabelFont}}px sans-serif`;ctx2.fillStyle=lb.color;
                ctx2.textAlign=lb.right?'left':'right';ctx2.textBaseline='middle';
                ctx2.fillText(isFocus?`${{lb.label}}   ${{lb.pct}}`:lb.label,tickEndX+dir*(isFocus?8:3),lb.y);
            }});

            ctx2.restore();
        }}
    }};
    const evalDistCtx=document.getElementById('qa-eval-dist-chart');
    if(evalDistCtx){{
        qaEvalDistChart=new Chart(evalDistCtx,{{
            type:'doughnut',
            plugins:[evalDistLabelPlugin],
            data:{{labels:['M7','DMG','R4H','Parentis','Britelift','Britelift Chat','RideX','Hamilton','Skyline','VIP','C&H','Reno Cab','Trans Iowa','Data Carz','Associated Cab','Ollies','Circle Taxi','YCOV','Kelowna','Vermont','YCDC','Blueline'],datasets:[{{data:[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],backgroundColor:['#4F81BD','#0E7490','#0F766E','#2C3E8C','#C0392B','#9A3412','#8E44AD','#065F46','#0EA5E9','#D97706','#0891B2','#16A34A','#7C3AED','#EA580C','#0E7490','#BE123C','#0E7490','#047857','#166534','#0F766E','#155E75','#1D4ED8'],borderWidth:2,borderColor:'#fff'}}]}},
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
    if(active) setQADistFocusMode(false);
    if(active) setQAAiQeFocusMode(false);
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

function resizeQADistributionCharts() {{
    setTimeout(()=>{{
        if(qaDonutChart) qaDonutChart.resize();
        if(qaEvalDistChart) qaEvalDistChart.resize();
        if(qaAiQeTrendChart) qaAiQeTrendChart.resize();
    }}, 80);
}}

function setQADistFocusMode(active) {{
    if(active) document.body.classList.remove('qa-focus-mode');
    if(active) setQAAiQeFocusMode(false);
    if(active){{
        const tableBtn=document.getElementById('qa-focus-toggle');
        if(tableBtn){{
            tableBtn.setAttribute('aria-label','Expand table');
            tableBtn.setAttribute('title','Expand table');
        }}
    }}
    document.body.classList.toggle('qa-dist-focus-mode', active);
    const btn=document.getElementById('qa-dist-focus-toggle');
    if(btn){{
        btn.setAttribute('aria-label', active?'Collapse distribution':'Expand distribution');
        btn.setAttribute('title', active?'Collapse distribution':'Expand distribution');
    }}
    if(qaEvalDistChart){{
        qaEvalDistChart.options.cutout=active?'42%':'50%';
        qaEvalDistChart.options.layout.padding=active?{{top:34,bottom:34,left:118,right:118}}:{{top:28,bottom:28,left:28,right:28}};
        qaEvalDistChart.update();
    }}
    if((document.getElementById('qa-sel-account')?.value||'')==='') qaUpdateEvalDist(qaCurrentFiltered);
    resizeQADistributionCharts();
}}

function toggleQADistFocusMode() {{
    setQADistFocusMode(!document.body.classList.contains('qa-dist-focus-mode'));
}}

function qaResetAiQeFocusCards() {{
    const grid=document.querySelector('#qualityPanel .qa-g3');
    if(grid){{
        grid.style.display='';
        grid.style.justifyContent='';
        grid.style.alignItems='';
    }}
    document.querySelectorAll('#qualityPanel .qa-g3 > .qa-card').forEach(card=>{{
        card.style.display='';
        card.style.height='';
        card.style.maxHeight='';
        card.style.width='';
        card.style.maxWidth='';
        card.style.minHeight='';
    }});
    const host=document.getElementById('qa-aiqe-chart-host');
    if(host) host.style.height='';
    const canvas=document.getElementById('qa-aiqe-trend-chart');
    if(canvas){{
        canvas.style.width='';
        canvas.style.height='';
        canvas.removeAttribute('width');
        canvas.removeAttribute('height');
    }}
}}

function qaApplyAiQeFocusCards(active) {{
    qaResetAiQeFocusCards();
    if(!active) return;
    const grid=document.querySelector('#qualityPanel .qa-g3');
    const aiqeCard=document.getElementById('qa-aiqe-card');
    const host=document.getElementById('qa-aiqe-chart-host');
    if(!grid || !aiqeCard) return;
    grid.style.display='flex';
    grid.style.justifyContent='center';
    grid.style.alignItems='flex-start';
    document.querySelectorAll('#qualityPanel .qa-g3 > .qa-card').forEach(card=>{{
        card.style.display=(card===aiqeCard)?'flex':'none';
    }});
    aiqeCard.style.height='75vh';
    aiqeCard.style.maxHeight='900px';
    aiqeCard.style.width='95vw';
    aiqeCard.style.maxWidth='1900px';
    aiqeCard.style.minHeight='0';
    if(host) host.style.height='min(64vh, 780px)';
}}

function setQAAiQeFocusMode(active) {{
    if(active) {{
        document.body.classList.remove('qa-focus-mode');
        document.body.classList.remove('qa-dist-focus-mode');
        const tableBtn=document.getElementById('qa-focus-toggle');
        if(tableBtn){{
            tableBtn.setAttribute('aria-label','Expand table');
            tableBtn.setAttribute('title','Expand table');
        }}
        const distBtn=document.getElementById('qa-dist-focus-toggle');
        if(distBtn){{
            distBtn.setAttribute('aria-label','Expand distribution');
            distBtn.setAttribute('title','Expand distribution');
        }}
    }}
    document.body.classList.toggle('qa-aiqe-focus-mode',active);
    qaApplyAiQeFocusCards(active);
    const btn=document.getElementById('qa-aiqe-focus-toggle');
    if(btn){{
        btn.setAttribute('aria-label',active?'Collapse AI x QE trend':'Expand AI x QE trend');
        btn.setAttribute('title',active?'Collapse AI x QE trend':'Expand AI x QE trend');
    }}
    setTimeout(()=>{{ if(qaAiQeTrendChart) qaAiQeTrendChart.resize(); }},80);
}}

function toggleQAAiQeFocusMode() {{
    setQAAiQeFocusMode(!document.body.classList.contains('qa-aiqe-focus-mode'));
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
        // Close a stray Masterlist expand modal if the user navigates away
        // while it's open (setMasterlistFocusMode is no longer wired to any
        // button — the Master List "expand" button now opens mlOpenExpand()).
        mlCloseExpand();
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
// Master List "expand" button now opens the ml* expand modal (replaces the old
// masterlist-focus-mode toggle so the button isn't fought over by two mechanisms).
document.getElementById("masterlistFocusToggle")?.addEventListener("click", (e) => {{
    mlOpenExpand(document.getElementById("masterlistCard"), e.currentTarget);
}});
mlWireTable();
mlWireExpandModal();
document.getElementById("mlClearFiltersBtn")?.addEventListener("click", mlClearFilters);
{{
    const mainHdr = document.querySelector(".sticky-dashboard-header");
    if (mainHdr) {{
        document.documentElement.style.setProperty("--ml-stick-top", mainHdr.offsetHeight + "px");
    }}
}}
renderCoaching();
render();

// Item 5 — first-load render/overflow fix. The initial render() above runs
// synchronously during parse, before the browser has fully settled layout
// (sticky header height, logo image, scrollbar gutter), so the Masterlist
// canvases can measure the wrong width on that very first paint — producing
// a bad vertical-chart layout and, since a too-wide measurement can briefly
// exceed the viewport, a stray page-level horizontal scrollbar. A manual
// refresh "fixes" it only because layout is already settled by the time
// scripts re-run. Re-measure/redraw after window 'load', once more a
// rAF double-tick later (next two paints), and once more after a short
// settle delay — each gated on the Masterlist tab actually being active so
// this never fights the existing tab-switch/filter-change/focus-toggle path
// (reflowDashboard -> mlDrawAll). A real window resize previously had no
// listener at all for these canvases — add one (debounced) here too.
function mlRedrawIfActive() {{
    if (document.getElementById("masterlistPanel")?.classList.contains("active")) {{
        mlDrawAll();
    }}
}}
window.addEventListener("load", () => {{
    mlRedrawIfActive();
    requestAnimationFrame(() => requestAnimationFrame(mlRedrawIfActive));
    setTimeout(mlRedrawIfActive, 300);
}});
let mlResizeSettleTimer = null;
window.addEventListener("resize", () => {{
    if (mlResizeSettleTimer) clearTimeout(mlResizeSettleTimer);
    mlResizeSettleTimer = setTimeout(mlRedrawIfActive, 150);
}});
</script>

<script>
(function() {{
    // --- Data freshness indicator (updates every 60s without page refresh) ---
    (function() {{
        var schedTimes = ['03:00','06:00','11:00','15:00','19:00','22:00'];
        var el = document.getElementById('pb-data-freshness');
        if (!el) return;

        function updateFreshness() {{
            var now = new Date();

            // Find the most recent scheduled refresh time that has already passed
            var lastSched = null;
            schedTimes.forEach(function(t) {{
                var parts = t.split(':');
                var d = new Date(now);
                d.setHours(parseInt(parts[0]), parseInt(parts[1]), 0, 0);
                if (d <= now) {{
                    if (!lastSched || d > lastSched) lastSched = d;
                }}
            }});
            // If none today have passed yet, use last one from yesterday
            if (!lastSched) {{
                var parts = schedTimes[schedTimes.length - 1].split(':');
                lastSched = new Date(now);
                lastSched.setDate(lastSched.getDate() - 1);
                lastSched.setHours(parseInt(parts[0]), parseInt(parts[1]), 0, 0);
            }}

            var buildTs = PB_BUILD_TS;
            var isLive = buildTs >= lastSched;
            var minsAgo = Math.round((now - buildTs) / 60000);
            var hoursAgo = Math.floor(minsAgo / 60);
            var ageStr = hoursAgo > 0
                ? hoursAgo + 'h ' + (minsAgo % 60) + 'm'
                : minsAgo + 'm';

            if (isLive) {{
                el.innerHTML = '<span style="display:inline-flex;align-items:center;gap:5px">'
                    + '<span style="width:8px;height:8px;border-radius:50%;background:#16A34A;display:inline-block;animation:qa-pulse 2s ease-in-out infinite"></span>'
                    + '<span style="color:#15803D">Live Data</span>'
                    + '<span style="color:#94A3B8;font-weight:400">&nbsp;&mdash;&nbsp;' + ageStr + '</span>'
                    + '</span>';
            }} else {{
                el.innerHTML = '<span style="display:inline-flex;align-items:center;gap:5px">'
                    + '<span style="width:8px;height:8px;border-radius:50%;background:#D97706;display:inline-block;animation:qa-pulse 2s ease-in-out infinite"></span>'
                    + '<span style="color:#B45309">Stale Data</span>'
                    + '<span style="color:#94A3B8;font-weight:400">&nbsp;&mdash;&nbsp;' + ageStr + '</span>'
                    + '</span>';
            }}
        }}

        updateFreshness();
        setInterval(updateFreshness, 60000);
    }})();

    // Auto-reload at scheduled fresh-build times (30 min after Task Scheduler runs)
    // Task Scheduler: 3AM, 6AM, 11AM, 3PM, 7PM, 10PM → fresh builds ready at:
    var refreshTimes = ['03:30','06:30','11:30','15:30','19:30','22:30'];

    function getNextRefreshMs() {{
        var now = new Date();
        var candidates = refreshTimes.map(function(t) {{
            var parts = t.split(':');
            var d = new Date(now);
            d.setHours(parseInt(parts[0]), parseInt(parts[1]), 0, 0);
            if (d <= now) d.setDate(d.getDate() + 1);
            return d;
        }});
        return Math.min.apply(null, candidates.map(function(d) {{ return d - now; }}));
    }}

    function scheduleNextReload() {{
        var ms = getNextRefreshMs();
        setTimeout(function() {{ location.reload(); }}, ms);
    }}

    scheduleNextReload();
}})();
</script>

</body>
</html>
"""

    # Stamp the actual finish time just before writing so the "Refresh Time"
    # header and PB_BUILD_TS JS variable reflect when the file was completed,
    # not when dashboard.py started loading data (which can be minutes earlier).
    finish_time = datetime.now()
    html = html.replace(refresh_time, finish_time.strftime("%Y-%m-%d %I:%M %p"), 1)
    html = html.replace(refresh_iso,  finish_time.strftime("%Y-%m-%dT%H:%M:00"), 1)

    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")
    print(f"Dashboard generated: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

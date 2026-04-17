"""
Ashvale Civil Engineering - Daily Inspection Checker
Fetches LOLER/PUWER inspections from Safety Culture, analyses with Claude,
generates a colour-coded Excel report and emails it as an .xlsx attachment.
"""

import io
import json
import requests
import smtplib
import anthropic
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ─────────────────────────────────────────
# CONFIGURATION — fill these in
# ─────────────────────────────────────────

SAFETY_CULTURE_TOKEN = "YOUR_SAFETY_CULTURE_TOKEN"
ANTHROPIC_API_KEY     = "YOUR_ANTHROPIC_API_KEY"

TEMPLATE_IDS = [
    "template_7dbbb416041a44459216a2a0ba02bb10",  # LOLER Inspection
    "template_0a8a57e828e746f782b2659da47f398d",  # Equipment Inspection
]

SMTP_SERVER    = "smtp.gmail.com"
SMTP_PORT      = 587
EMAIL_FROM     = "YOUR_EMAIL@gmail.com"
EMAIL_PASSWORD = "YOUR_APP_PASSWORD"   # Gmail App Password, not your normal password
EMAIL_TO       = "YOUR_EMAIL@ashvale.co.uk"

MAX_DAYS_BACK = 30

# ─────────────────────────────────────────
# COLOURS
# ─────────────────────────────────────────

FILL_HEADER   = PatternFill("solid", fgColor="F4A460")   # salmon/orange header
FILL_RAMS_HDR = PatternFill("solid", fgColor="FFFF00")   # yellow RAMS header
FILL_YELLOW   = PatternFill("solid", fgColor="FFFF00")   # warning cell
FILL_RED      = PatternFill("solid", fgColor="FF0000")   # critical note
FILL_NONE     = PatternFill("none")

FONT_HEADER   = Font(bold=True, color="000000")
FONT_WHITE    = Font(bold=True, color="FFFFFF")
FONT_BLACK    = Font(color="000000")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

COLUMNS = [
    ("SITE",      22),
    ("JOB NO",    9),
    ("PUWER",     8),
    ("LOLER",     8),
    ("SITE SUP",  9),
    ("XCAVATOR",  10),
    ("DUMPER",    9),
    ("ROLLER",    9),
    ("TELEHAND",  11),
    ("HAVS",      8),
    ("TOOLBOX",   10),
    ("RAMS",      8),
    ("SUPERVISOR",16),
    ("NOTES",     45),
]

# Maps column header → Claude JSON key
COL_KEY = {
    "PUWER":    "puwer",
    "LOLER":    "loler",
    "SITE SUP": "site_sup",
    "XCAVATOR": "xcavator",
    "DUMPER":   "dumper",
    "ROLLER":   "roller",
    "TELEHAND": "telehand",
    "HAVS":     "havs",
    "TOOLBOX":  "toolbox",
    "RAMS":     "rams",
}

# ─────────────────────────────────────────
# SAFETY CULTURE API
# ─────────────────────────────────────────

def sc_headers():
    return {
        "Authorization": f"Bearer {SAFETY_CULTURE_TOKEN}",
        "Content-Type": "application/json",
    }


def get_inspection_ids(template_id, days_back=30):
    modified_after = (datetime.utcnow() - timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url = "https://api.safetyculture.io/audits/search"
    params = {
        "field": ["audit_id", "modified_at", "template_id"],
        "template": template_id,
        "modified_after": modified_after,
        "completed": "true",
        "limit": 100,
    }
    r = requests.get(url, headers=sc_headers(), params=params)
    r.raise_for_status()
    return r.json().get("audits", [])


def get_inspection_detail(audit_id):
    url = f"https://api.safetyculture.io/audits/{audit_id}"
    r = requests.get(url, headers=sc_headers())
    r.raise_for_status()
    return r.json()


def get_most_recent_ids(inspections_raw):
    sorted_insp = sorted(inspections_raw, key=lambda x: x.get("modified_at", ""), reverse=True)
    return [i["audit_id"] for i in sorted_insp[:20]]


# ─────────────────────────────────────────
# PARSE INSPECTION
# ─────────────────────────────────────────

def parse_inspection(inspection):
    audit_data    = inspection.get("audit_data", {})
    template_data = inspection.get("template_data", {})
    items         = inspection.get("items", [])

    site          = audit_data.get("site", {}).get("name", "Unknown Site")
    supervisor    = audit_data.get("authorship", {}).get("owner", "Unknown")
    date_str      = audit_data.get("date_started", "Unknown")
    completed     = audit_data.get("date_completed")
    score         = audit_data.get("score")
    total         = audit_data.get("total_score")
    pct           = audit_data.get("score_percentage")
    job_no        = audit_data.get("reference_id") or audit_data.get("audit_id", "")[:8]
    template_name = template_data.get("metadata", {}).get("name", "Unknown Template")

    try:
        dt = datetime.fromisoformat(date_str.replace("Z", ""))
        date_formatted = dt.strftime("%d %b %Y %H:%M")
    except Exception:
        date_formatted = date_str

    equipment_lines = []
    for item in items:
        label     = (item.get("label") or "").strip()
        if not label:
            continue
        responses = item.get("responses", {})
        selected  = responses.get("selected", [])
        text_resp = (responses.get("text") or "").strip()
        note      = (item.get("note") or "").strip()
        status    = selected[0].get("label", "").strip() if selected else ""

        line = label
        if status:   line += f": {status}"
        if text_resp: line += f" | text: {text_resp}"
        if note:     line += f" | note: {note}"
        equipment_lines.append(line)

    full_text = (
        f"Template: {template_name}\n"
        f"Site: {site}\n"
        f"Date: {date_formatted}\n"
        f"Supervisor: {supervisor}\n"
        f"Score: {score} / {total} ({pct}%)\n"
        f"Signed Off: {'Yes' if completed else 'No'}\n\n"
        f"EQUIPMENT / ITEMS:\n" + "\n".join(equipment_lines)
    )

    return {
        "audit_id":      inspection.get("audit_id", ""),
        "site":          site,
        "job_no":        job_no,
        "supervisor":    supervisor,
        "date":          date_formatted,
        "signed":        bool(completed),
        "template_name": template_name,
        "full_text":     full_text,
    }


# ─────────────────────────────────────────
# CLAUDE ANALYSIS — returns structured JSON
# ─────────────────────────────────────────

SYSTEM_PROMPT = """You are an H&S compliance checker for Ashvale Civil Engineering.
Analyse the Safety Culture inspection and return ONLY valid JSON — no markdown, no explanation.

JSON schema:
{
  "puwer":        "Y" | "N" | "N/A",
  "loler":        "Y" | "N" | "N/A",
  "site_sup":     "Y" | "N" | "N/A",
  "xcavator":     "Y" | "N" | "N/A",
  "dumper":       "Y" | "N" | "N/A",
  "roller":       "Y" | "N" | "N/A",
  "telehand":     "Y" | "N" | "N/A",
  "havs":         "Y" | "N" | "N/A",
  "toolbox":      "Y" | "N" | "N/A",
  "rams":         "Y" | "N" | "N/A",
  "note":         "string or empty string",
  "note_severity": "critical" | "warning" | "none",
  "warning_cells": ["list of keys from above that should be highlighted yellow"]
}

Rules:
- Y  = inspection completed and compliant
- N  = inspection done but failed / issue found
- N/A = equipment not present on site
- warning_cells = any key above where the value is N or there is a specific issue with that item
- note_severity = "critical" if any serious breach (missing serials, daily checks not done, At Risk equipment),
                  "warning" if minor / incomplete (e.g. not yet done, one item pending),
                  "none" if fully clear
- note = one concise sentence describing the issue, or empty string if none"""


def analyse_with_claude(parsed):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    message = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=512,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"INSPECTION DATA:\n{parsed['full_text']}"}],
    )

    raw = message.content[0].text.strip()
    # Strip any accidental markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ─────────────────────────────────────────
# BUILD EXCEL
# ─────────────────────────────────────────

def build_excel(results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = datetime.now().strftime("%d %b %Y")

    # ── Header row ──
    for col_idx, (header, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill      = FILL_RAMS_HDR if header == "RAMS" else FILL_HEADER
        cell.font      = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = THIN_BORDER
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 22

    # ── Data rows ──
    for row_idx, r in enumerate(results, start=2):
        analysis = r.get("analysis", {})
        warn_cells  = set(analysis.get("warning_cells", []))
        note        = analysis.get("note", "")
        severity    = analysis.get("note_severity", "none")

        row_data = [
            r["site"],
            r.get("job_no", ""),
            analysis.get("puwer",    ""),
            analysis.get("loler",    ""),
            analysis.get("site_sup", ""),
            analysis.get("xcavator", ""),
            analysis.get("dumper",   ""),
            analysis.get("roller",   ""),
            analysis.get("telehand", ""),
            analysis.get("havs",     ""),
            analysis.get("toolbox",  ""),
            analysis.get("rams",     ""),
            r["supervisor"],
            note,
        ]

        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            header = COLUMNS[col_idx - 1][0]
            key    = COL_KEY.get(header)

            # Colour logic
            if header == "NOTES":
                if severity == "critical":
                    cell.fill = FILL_RED
                    cell.font = FONT_WHITE
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif severity == "warning":
                    cell.fill = FILL_YELLOW
                    cell.font = FONT_BLACK
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif key and key in warn_cells:
                cell.fill = FILL_YELLOW
                cell.font = FONT_BLACK
            else:
                cell.font = FONT_BLACK

        ws.row_dimensions[row_idx].height = 18

    # Freeze header row
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────
# EMAIL WITH XLSX ATTACHMENT
# ─────────────────────────────────────────

def send_email(results, xlsx_buf):
    today          = datetime.now().strftime("%d %b %Y")
    has_critical   = any(r.get("analysis", {}).get("note_severity") == "critical" for r in results)
    filename       = f"Ashvale_Inspections_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

    subject = (
        f"⚠️ Ashvale Inspections — Action Required — {today}"
        if has_critical
        else f"✅ Ashvale Inspections — All Clear — {today}"
    )

    body_text = (
        f"Ashvale Civil Engineering — Daily Inspection Report\n"
        f"Date: {today}\n"
        f"Inspections processed: {len(results)}\n\n"
        f"Please find the formatted Excel report attached.\n"
        + ("⚠️  One or more sites require action — see red rows in the spreadsheet." if has_critical else "✅  All inspections clear.")
    )

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"]    = EMAIL_FROM
    msg["To"]      = EMAIL_TO
    msg.attach(MIMEText(body_text, "plain"))

    part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    part.set_payload(xlsx_buf.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

    print(f"  Email sent → {EMAIL_TO}  ({filename})")


# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────

def main():
    print(f"\n{'='*55}")
    print(f"Ashvale Inspection Checker — {datetime.now().strftime('%d %b %Y %H:%M')}")
    print(f"{'='*55}\n")

    all_results = []

    for template_id in TEMPLATE_IDS:
        print(f"Fetching inspections for: {template_id}")
        try:
            summaries = get_inspection_ids(template_id, days_back=MAX_DAYS_BACK)
            audit_ids = get_most_recent_ids(summaries)
            print(f"  Found {len(audit_ids)} inspection(s)")
        except Exception as e:
            print(f"  Error fetching list: {e}")
            continue

        for audit_id in audit_ids:
            print(f"  Processing {audit_id}...")
            try:
                detail   = get_inspection_detail(audit_id)
                parsed   = parse_inspection(detail)
                print(f"    Site: {parsed['site']} | {parsed['date']}")

                print(f"    Analysing with Claude...")
                analysis = analyse_with_claude(parsed)

                severity = analysis.get("note_severity", "none")
                icon     = "🔴" if severity == "critical" else ("🟡" if severity == "warning" else "✅")
                print(f"    → {icon} {severity.upper()}")

                all_results.append({**parsed, "analysis": analysis})

            except Exception as e:
                print(f"    Error processing {audit_id}: {e}")

    if not all_results:
        print("\nNo inspections found. No email sent.")
        return

    print(f"\nBuilding Excel for {len(all_results)} inspection(s)...")
    xlsx_buf = build_excel(all_results)

    print("Sending email with attachment...")
    send_email(all_results, xlsx_buf)
    print("\nDone ✓")


if __name__ == "__main__":
    main()

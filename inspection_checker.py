"""
Ashvale Civil Engineering - Daily Inspection Checker
Fetches LOLER/PUWER inspections from Safety Culture, analyses with Claude, emails report
"""

import requests
import smtplib
import anthropic
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ─────────────────────────────────────────
# CONFIGURATION — fill these in
# ─────────────────────────────────────────

SAFETY_CULTURE_TOKEN = "YOUR_SAFETY_CULTURE_TOKEN"
ANTHROPIC_API_KEY     = "YOUR_ANTHROPIC_API_KEY"

TEMPLATE_IDS = [
    "template_7dbbb416041a44459216a2a0ba02bb10",  # LOLER Inspection
    # "template_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",  # PUWER — add when ready
]

# Email settings
SMTP_SERVER    = "smtp.gmail.com"
SMTP_PORT      = 587
EMAIL_FROM     = "YOUR_EMAIL@gmail.com"
EMAIL_PASSWORD = "YOUR_APP_PASSWORD"   # Gmail App Password, not your normal password
EMAIL_TO       = "YOUR_EMAIL@ashvale.co.uk"

# How many days back to look
MAX_DAYS_BACK = 30

# ─────────────────────────────────────────
# SAFETY CULTURE API
# ─────────────────────────────────────────

def sc_headers():
    return {
        "Authorization": f"Bearer {SAFETY_CULTURE_TOKEN}",
        "Content-Type": "application/json"
    }

def get_inspection_ids(template_id, days_back=30):
    """Fetch list of inspection IDs for a template from the last N days"""
    modified_after = (datetime.utcnow() - timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url = "https://api.safetyculture.io/audits/search"
    params = {
        "field": ["audit_id", "modified_at", "template_id"],
        "template": template_id,
        "modified_after": modified_after,
        "completed": "true",
        "limit": 100
    }
    r = requests.get(url, headers=sc_headers(), params=params)
    r.raise_for_status()
    return r.json().get("audits", [])


def get_inspection_detail(audit_id):
    """Fetch full inspection detail"""
    url = f"https://api.safetyculture.io/audits/{audit_id}"
    r = requests.get(url, headers=sc_headers())
    r.raise_for_status()
    return r.json()


def get_most_recent_ids(inspections_raw):
    """Sort by most recent, return top 20 audit IDs"""
    sorted_inspections = sorted(
        inspections_raw,
        key=lambda x: x.get("modified_at", ""),
        reverse=True
    )
    return [i["audit_id"] for i in sorted_inspections[:20]]


# ─────────────────────────────────────────
# PARSE INSPECTION
# ─────────────────────────────────────────

def parse_inspection(inspection):
    """Extract key fields from inspection using correct data structure"""
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
        if status:
            line += f": {status}"
        if text_resp:
            line += f" | text: {text_resp}"
        if note:
            line += f" | note: {note}"
        equipment_lines.append(line)

    full_text = f"""Template: {template_name}
Site: {site}
Date: {date_formatted}
Supervisor: {supervisor}
Score: {score} / {total} ({pct}%)
Signed Off: {'Yes' if completed else 'No'}

EQUIPMENT / ITEMS:
{chr(10).join(equipment_lines)}
"""

    return {
        "audit_id":      inspection.get("audit_id", ""),
        "site":          site,
        "supervisor":    supervisor,
        "date":          date_formatted,
        "score":         f"{score} / {total} ({pct}%)" if score is not None else "Not recorded",
        "signed":        bool(completed),
        "template_name": template_name,
        "full_text":     full_text,
    }


# ─────────────────────────────────────────
# CLAUDE ANALYSIS
# ─────────────────────────────────────────

SYSTEM_PROMPT = """You are an H&S compliance checker for Ashvale Civil Engineering.
Analyse Safety Culture inspections and identify all issues.

Check for:
1. Missing serial numbers or plant IDs — equipment with a status (Safe/At Risk/Fail) but no serial number or ID recorded in the text or note field
2. Equipment with no status at all
3. Missing supervisor sign-off
4. Any equipment marked At Risk — must be flagged for plant team
5. Missing or incomplete dates
6. Any other LOLER/PUWER compliance gaps

Respond in this exact format:

CRITICAL:
- [issue] or None

WARNINGS:
- [issue] or None

PASSED:
- [item]

VERDICT: [one sentence]"""


def analyse_with_claude(parsed):
    """Send inspection to Claude for compliance analysis"""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    message = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=1000,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": f"INSPECTION DATA:\n{parsed['full_text']}"
            }
        ]
    )

    return message.content[0].text


# ─────────────────────────────────────────
# EMAIL REPORT
# ─────────────────────────────────────────

def build_email(results):
    """Build HTML email from list of results"""
    today = datetime.now().strftime("%d %b %Y")
    has_any_critical = any(r.get("has_critical") for r in results)

    subject = (
        f"⚠️ Ashvale Inspections — Action Required — {today}"
        if has_any_critical
        else f"✅ Ashvale Inspections — All Clear — {today}"
    )

    html = f"""<html><body style="font-family:Arial,sans-serif;background:#f4f4f4;padding:20px;">
<div style="max-width:700px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;">
  <div style="background:#1a1a2e;padding:24px;color:white;">
    <h1 style="margin:0;font-size:22px;">Ashvale Civil Engineering</h1>
    <p style="margin:4px 0 0;opacity:0.7;font-size:14px;">Daily Inspection Report — {today}</p>
  </div>
  <div style="padding:20px;">"""

    for r in results:
        border       = "#e05c3a" if r.get("has_critical") else "#28a745"
        status_label = "⚠️ ACTION REQUIRED" if r.get("has_critical") else "✅ CLEAR"
        status_color = "#e05c3a" if r.get("has_critical") else "#28a745"
        analysis_html = r.get("analysis", "No analysis").replace("\n", "<br>")

        html += f"""
    <div style="border:1px solid #e0e0e0;border-left:4px solid {border};border-radius:6px;margin-bottom:16px;">
      <div style="background:#f8f8f8;padding:12px 16px;border-bottom:1px solid #e0e0e0;">
        <strong style="font-size:15px;">{r['site']}</strong>
        <span style="float:right;color:{status_color};font-weight:bold;font-size:13px;">{status_label}</span>
      </div>
      <div style="padding:10px 16px;font-size:12px;color:#666;">
        📅 {r['date']} &nbsp;|&nbsp; 👤 {r['supervisor']} &nbsp;|&nbsp;
        📊 {r['score']} &nbsp;|&nbsp; ✍️ {'Signed' if r['signed'] else '❌ NOT SIGNED'}
      </div>
      <div style="padding:12px 16px;font-size:13px;color:#333;font-family:monospace;white-space:pre-wrap;line-height:1.6;">{analysis_html}</div>
    </div>"""

    html += """
  </div>
  <div style="background:#f8f8f8;padding:16px;text-align:center;font-size:12px;color:#888;border-top:1px solid #e0e0e0;">
    Ashvale Civil Engineering — Automated H&S Inspection Report
  </div>
</div>
</body></html>"""

    return subject, html


def send_email(subject, html_body):
    """Send the report email"""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = EMAIL_FROM
    msg["To"]      = EMAIL_TO
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())

    print(f"  Email sent to {EMAIL_TO}")


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

                has_critical = False
                if "CRITICAL:" in analysis:
                    critical_section = analysis.split("CRITICAL:")[1].split("WARNINGS:")[0]
                    has_critical = "none" not in critical_section.lower()

                all_results.append({
                    **parsed,
                    "analysis":     analysis,
                    "has_critical": has_critical,
                })

                print(f"    → {'⚠️  Issues found' if has_critical else '✅ Clear'}")

            except Exception as e:
                print(f"    Error processing {audit_id}: {e}")

    if not all_results:
        print("\nNo inspections found. No email sent.")
        return

    print(f"\nBuilding email for {len(all_results)} inspection(s)...")
    subject, html = build_email(all_results)
    print(f"Sending: {subject}")
    send_email(subject, html)
    print("\nDone ✓")


if __name__ == "__main__":
    main()

"""
Maps SafetyCulture template names → Excel column check types.

Each CHECK_TYPE entry lists keywords that, if found in the template name
(case-insensitive), classify the audit into that column.
"""

# Order matters for overlap: more specific patterns first.
CHECK_COLUMNS = [
    ("LOLER",    ["loler"]),
    ("PUWER",    ["puwer"]),
    ("SITE SUP", ["site supervisor", "supervisor weekly", "ss weekly", "site sup"]),
    ("EXCAVATOR",["excavator daily", "exc daily", "excavator check", "14t check",
                  "8t check", "20t check", "25t check", "3t exc", "6t exc"]),
    ("DUMPER",   ["dumper daily", "dumper check", "9t check", "6t dump", "3t dump"]),
    ("ROLLER",   ["roller daily", "roller check", "ror daily", "roller inspection"]),
    ("TELEHAND", ["telehandler daily", "telehandler check", "tele daily", "tl daily"]),
    ("HAVS",     ["hav operative", "hav/noise", "hav noise", "havs", "hav daily",
                  "hand arm", "noise dust"]),
    ("TOOLBOX",  ["toolbox", "tool box", "tbt"]),
    ("RAMS",     ["rams", "risk assessment", "method statement"]),
]

# Status symbols
STATUS_COMPLETE   = "✅"
STATUS_IN_PROGRESS = "🔄"
STATUS_MISSING    = "❌"
STATUS_NA         = "N/A"


def classify_template(template_name: str) -> str | None:
    """Return the Excel column name for a template, or None if unrecognised."""
    name_lower = template_name.lower()
    for col, keywords in CHECK_COLUMNS:
        if any(kw in name_lower for kw in keywords):
            return col
    return None


def audit_status(audit: dict) -> str:
    """Return ✅ / 🔄 based on audit completion status."""
    status = (audit.get("status") or "").lower()
    if status in ("completed", "complete"):
        return STATUS_COMPLETE
    return STATUS_IN_PROGRESS

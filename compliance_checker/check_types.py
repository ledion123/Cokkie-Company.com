"""
Maps SafetyCulture template names → Excel column check types.

Each CHECK_TYPE entry lists keywords that, if found in the template name
(case-insensitive), classify the audit into that column.
"""

# Order matters: plant-specific PUWER daily checks MUST come before the
# generic PUWER register, otherwise they'd all match "puwer" first.
# Confirmed SC template names are marked ✅.
CHECK_COLUMNS = [
    ("LOLER",    [
        "loler inspection",                                  # ✅ confirmed
        "loler",
    ]),
    # Plant daily checks (all start with "PUWER" — matched before generic PUWER)
    ("EXCAVATOR", [
        "puwer excavator daily check sheet",                 # ✅ confirmed
        "excavator daily", "exc daily", "excavator check",
    ]),
    ("DUMPER",    [
        "puwer dumper daily check sheet",                    # ✅ confirmed
        "dumper daily", "dumper check",
    ]),
    ("ROLLER",    [
        "puwer ride on roller daily check sheet",            # ✅ confirmed
        "roller daily", "roller check", "ror daily",
    ]),
    ("TELEHAND",  [
        "puwer telescopic- handler daily check sheet",       # ✅ confirmed
        "puwer telescopic handler daily check sheet",        # variant without hyphen
        "telehandler daily", "telehandler check", "tele daily",
    ]),
    # Generic PUWER register (weekly) — matched AFTER plant dailies
    ("PUWER",    [
        "puwer equipment register & inspection",             # ✅ confirmed
        "puwer equipment register",
    ]),
    ("SITE SUP", [
        "site supervisor weekly inspection",                 # ✅ confirmed
        "site supervisor", "supervisor weekly", "ss weekly",
    ]),
    ("HAVS",     [
        "hav / noise / dust daily monitoring",               # ✅ confirmed
        "hav / noise", "hav/noise", "hav noise", "havs",
        "hav operative", "hand arm", "noise dust",
    ]),
    ("TOOLBOX",  [
        "tool box talk",                                     # ✅ confirmed
        "toolbox talk", "toolbox", "tbt",
    ]),
    ("RAMS",     [
        "rams", "risk assessment", "method statement",       # unconfirmed – update when known
    ]),
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

"""
Parse the "WHAT'S OUT WHERE" PDF to build a per-site equipment register.
Returns: {canonical_site_name: {"excavators": [...], "dumpers": [...], ...}}

PDF format per line (no fixed column spacing):
  "SITE NAME [JOBCODE] EQUIPMENT DESCRIPTION"   (site has a job code)
  "SITE NAME EQUIPMENT DESCRIPTION"             (site has no job code)
"""
from __future__ import annotations

import re
import pdfplumber
from site_lookup import resolve_site_name

# Job code in brackets, e.g. [LP009], [CA0005], [DAN003]
JOB_CODE_RE = re.compile(r'\[[A-Z]{2,3}\d{3,5}\]')

# Keywords that mark the start of the equipment description.
# No trailing \b so "EXC" matches "EXCAVATOR" and "DUMP" matches "DUMPER".
EQUIP_KW_RE = re.compile(
    r'\b(\d+(?:\.\d+)?T\s+(?:EXC|DUMP)|TELEHANDLER|'
    r'(?:120|170)\s*ROLLER|RAMMAX\s*ROLLER|RAMMAX)',
    re.I,
)


def _classify_equipment(description: str) -> str | None:
    desc = description.upper()
    if "TELEHANDLER" in desc:
        return "telehandler"
    if "ROLLER" in desc or "RAMMAX" in desc:
        return "roller"
    if re.search(r'\d+T\s+EXC|EXCAVATOR', desc):
        return "excavator"
    if re.search(r'\d+T\s+DUMP|DUMPER', desc):
        return "dumper"
    return None


def _extract_id(text: str) -> str | None:
    """Return ASH NNN / ROR NN / TL NN / RAM NN if present, else 'HIRED'."""
    m = re.search(r'\b(ASH\s*\d+|ROR\s*\d+|TL\s*\d+|RAM\s*\d+)\b', text, re.I)
    if m:
        return re.sub(r'\s+', ' ', m.group(1).upper().strip())
    return "HIRED"


def _split_line(line: str) -> tuple[str, str] | tuple[None, None]:
    """
    Return (site_raw, description) from one PDF line, or (None, None) if
    the line does not contain a recognisable plant entry.
    """
    # Skip "ASHVALE MISC" lines — these are yard/spare entries, not site entries
    if re.match(r'ASHVALE\s+MISC', line, re.I):
        return None, None

    # Strategy 1: split after the closing ']' of a job code, e.g. "[LP009] "
    m = JOB_CODE_RE.search(line)
    if m:
        split_pos = m.end()
        site_raw = line[:split_pos].strip()
        desc = line[split_pos:].strip()
        if desc:
            return site_raw, desc

    # Strategy 2: split at the equipment keyword for sites without job codes
    km = EQUIP_KW_RE.search(line)
    if km:
        site_raw = line[:km.start()].strip()
        desc = line[km.start():].strip()
        if site_raw and len(site_raw) >= 4:
            return site_raw, desc

    return None, None


def parse_pdf(pdf_path: str) -> dict:
    """
    Returns:
        {
            "Denny End Rd, Waterbeach [LP0010]": {
                "excavators":   ["ASH 1424", "ASH 1425", ...],
                "dumpers":      ["ASH 916", ...],
                "rollers":      ["ROR 16", "RAM 03"],
                "telehandlers": ["TL 04"],
            },
            ...
        }
    """
    equipment: dict[str, dict] = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if not text:
                continue
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue

                site_raw, desc = _split_line(line)
                if site_raw is None:
                    continue

                plant_type = _classify_equipment(desc)
                if plant_type is None:
                    continue

                plant_id = _extract_id(desc)
                site_name = resolve_site_name(site_raw, "")

                if site_name not in equipment:
                    equipment[site_name] = {
                        "excavators": [],
                        "dumpers": [],
                        "rollers": [],
                        "telehandlers": [],
                    }

                bucket = plant_type + "s"
                if plant_id not in equipment[site_name][bucket]:
                    equipment[site_name][bucket].append(plant_id)

    return equipment


def summarise_equipment(site_equipment: dict) -> str:
    """Return a short human-readable summary for one site."""
    parts = []
    for label, key in [("Exc", "excavators"), ("Dump", "dumpers"),
                       ("Roll", "rollers"), ("TL", "telehandlers")]:
        items = site_equipment.get(key, [])
        if items:
            parts.append(f"{label}: {', '.join(items)}")
    return " | ".join(parts) if parts else "No plant recorded"

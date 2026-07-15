"""
Parse the "WHAT'S OUT WHERE" PDF to build a per-site equipment register.
Returns: {canonical_site_name: {"excavators": [...], "dumpers": [...], ...}}
"""
from __future__ import annotations

import re
import pdfplumber
from site_lookup import resolve_site_name


def _extract_id(text: str) -> str | None:
    """Pull ASH nnn / ROR nn / TL nn / RAM nn from a description string."""
    m = re.search(r"\b(ASH\s*\d+|ROR\s*\d+|TL\s*\d+|RAM\s*\d+)\b", text, re.I)
    if m:
        return re.sub(r"\s+", " ", m.group(1).upper().strip())
    return None


def _classify_equipment(description: str) -> str | None:
    desc = description.upper()
    if "TELEHANDLER" in desc:
        return "telehandler"
    if "ROLLER" in desc or "RAMMAX" in desc:
        return "roller"
    if re.search(r"\d+T\s+EXC|EXCAVATOR", desc):
        return "excavator"
    if re.search(r"\d+T\s+DUMP|DUMPER", desc):
        return "dumper"
    return None


def parse_pdf(pdf_path: str) -> dict:
    """
    Returns:
        {
            "Denny End Rd, Waterbeach [LP0010]": {
                "excavators": ["ASH 1424", "ASH 1425", ...],
                "dumpers":    ["ASH 916", ...],
                "rollers":    ["ROR 16", "RAM 03"],
                "telehandlers": [],
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

                # Try to split "SITE NAME   EQUIPMENT DESCRIPTION"
                # The PDF uses wide spacing between the two columns
                parts = re.split(r"\s{3,}", line, maxsplit=1)
                if len(parts) != 2:
                    continue
                site_raw, desc = parts[0].strip(), parts[1].strip()

                # Ignore header / page artefacts
                if not site_raw or len(site_raw) < 4:
                    continue

                plant_type = _classify_equipment(desc)
                if plant_type is None:
                    continue

                plant_id = _extract_id(desc)
                if plant_id is None:
                    continue

                # Resolve to canonical name using site_lookup
                site_name = resolve_site_name(site_raw, "")

                if site_name not in equipment:
                    equipment[site_name] = {
                        "excavators": [],
                        "dumpers": [],
                        "rollers": [],
                        "telehandlers": [],
                    }

                bucket = plant_type + "s"  # "excavator" → "excavators"
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

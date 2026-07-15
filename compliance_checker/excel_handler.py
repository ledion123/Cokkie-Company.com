"""
Read the weekly Excel template and write compliance results back.
"""

import re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from site_lookup import resolve_site_name, flag_linmere_ambiguity
from check_types import STATUS_COMPLETE, STATUS_IN_PROGRESS, STATUS_MISSING, STATUS_NA

# Expected column headers (case-insensitive)
COL_SITE      = "site"
COL_JOB       = "job no"
COL_PUWER     = "puwer"
COL_LOLER     = "loler"
COL_SITE_SUP  = "site sup"
COL_EXCAVATOR = "excavator"
COL_DUMPER    = "dumper"
COL_ROLLER    = "roller"
COL_TELEHAND  = "telehand"
COL_HAVS      = "havs"
COL_TOOLBOX   = "toolbox"
COL_RAMS      = "rams"
COL_SUPERVISOR = "supervisor"

CHECK_COLS = [COL_PUWER, COL_LOLER, COL_SITE_SUP, COL_EXCAVATOR,
              COL_DUMPER, COL_ROLLER, COL_TELEHAND, COL_HAVS,
              COL_TOOLBOX, COL_RAMS]

# Cell fill colours
GREEN  = PatternFill("solid", fgColor="C6EFCE")
ORANGE = PatternFill("solid", fgColor="FFEB9C")
RED    = PatternFill("solid", fgColor="FFC7CE")
GREY   = PatternFill("solid", fgColor="D9D9D9")


def _col_index(headers: list[str | None]) -> dict[str, int]:
    """Return {normalised_header: 0-based column index}."""
    idx = {}
    for i, h in enumerate(headers):
        if h is None:
            continue
        key = re.sub(r"[\s\r\n]+", " ", str(h)).strip().lower().rstrip()
        idx[key] = i
    return idx


def read_sites(xlsx_path: str) -> list[dict]:
    """
    Parse the Excel template and return a list of site dicts:
      [{"site": "WaterBeach", "job": "LP10", "canonical": "Denny End Rd...", ...}, ...]
    """
    wb = load_workbook(xlsx_path)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [str(h).strip().lower() if h else None for h in rows[0]]
    idx = _col_index(headers)

    sites = []
    for row in rows[1:]:
        site_val = row[idx.get(COL_SITE, 0)] if COL_SITE in idx else None
        job_val  = row[idx.get(COL_JOB, 1)]  if COL_JOB  in idx else None
        sup_val  = row[idx.get(COL_SUPERVISOR, len(row)-1)] if COL_SUPERVISOR in idx else None

        if not site_val and not job_val:
            continue

        canonical = resolve_site_name(str(site_val or ""), str(job_val or ""))
        ambiguity = flag_linmere_ambiguity(str(site_val or ""), str(job_val or ""))

        sites.append({
            "site":       str(site_val or "").strip(),
            "job":        str(job_val or "").strip(),
            "canonical":  canonical,
            "supervisor": str(sup_val or "").strip(),
            "ambiguity":  ambiguity,
        })

    return sites


def write_results(
    xlsx_path: str,
    output_path: str,
    results: list[dict],
    week_label: str = "",
):
    """
    Write ✅/🔄/❌/N/A into the template and save to output_path.

    results: list of dicts matching the sites from read_sites(), each with
      additional key "checks": {col_name: status_string}
      and optional "notes": str
    """
    wb = load_workbook(xlsx_path)
    ws = wb.active

    header_row = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    headers = [str(h).strip().lower() if h else None for h in header_row]
    idx = _col_index(headers)

    # Add week label in cell A1 if blank
    if week_label:
        ws["A1"] = f"Week: {week_label}  |  " + (str(ws["A1"].value) or "")

    for row_num, (row_data, result) in enumerate(
        zip(ws.iter_rows(min_row=2), results), start=2
    ):
        checks = result.get("checks", {})
        notes  = result.get("notes", "")

        for col_name in CHECK_COLS:
            if col_name.lower() not in idx:
                continue
            col_i = idx[col_name.lower()] + 1  # openpyxl is 1-based
            cell  = ws.cell(row=row_num, column=col_i)

            status = checks.get(col_name, STATUS_MISSING)

            cell.value     = status
            cell.alignment = Alignment(horizontal="center")

            if status == STATUS_COMPLETE:
                cell.fill = GREEN
            elif status == STATUS_IN_PROGRESS:
                cell.fill = ORANGE
            elif status == STATUS_MISSING:
                cell.fill = RED
            elif status == STATUS_NA:
                cell.fill = GREY

        # Write notes into the cell after RAMS if there's room
        if notes and "rams" in idx:
            note_col = idx["rams"] + 2  # one column after RAMS
            ws.cell(row=row_num, column=note_col).value = notes

        # Highlight Linmere ambiguity warning in site name cell
        ambiguity = result.get("ambiguity")
        if ambiguity and COL_SITE in idx:
            site_cell = ws.cell(row=row_num, column=idx[COL_SITE] + 1)
            site_cell.font = Font(color="FF0000", bold=True)

    # Auto-width site column
    if COL_SITE in idx:
        col_letter = get_column_letter(idx[COL_SITE] + 1)
        ws.column_dimensions[col_letter].width = 30

    wb.save(output_path)
    print(f"✅ Saved: {output_path}")

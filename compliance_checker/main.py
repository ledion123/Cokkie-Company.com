"""
Ashvale Compliance Checker
==========================
Usage:
  python main.py --excel weekly_template.xlsx --pdf WHATS_OUT_WHERE.pdf [--week 2026-07-13]

  --week  Any date in the target week (Mon–Sun). If a Sunday date is given,
          it is the week ENDING that Sunday. Defaults to current week.
          Format: YYYY-MM-DD

Outputs:
  <excel_basename>_FILLED_<week>.xlsx  in the same folder as the input Excel.
"""
from __future__ import annotations

import argparse
import sys
import os
from datetime import datetime
from pathlib import Path
from collections import defaultdict

from dotenv import load_dotenv

from api import SCClient
from check_types import (classify_template, audit_status,
                         STATUS_COMPLETE, STATUS_IN_PROGRESS,
                         STATUS_MISSING, STATUS_NA)
from excel_handler import read_sites, write_results
from pdf_parser import parse_pdf, summarise_equipment
from site_lookup import resolve_site_name
from week_utils import week_bounds, week_label, is_week_in_progress

load_dotenv(Path(__file__).parent / ".env")


# ── Columns that only apply when plant exists on site ────────────────────────
PLANT_COLS = {
    "EXCAVATOR": "excavators",
    "DUMPER":    "dumpers",
    "ROLLER":    "rollers",
    "TELEHAND":  "telehandlers",
}


def build_sc_site_map(client: SCClient, canonical_names: list[str]) -> dict[str, str]:
    """
    Fetch all SC sites and fuzzy-match to canonical site names.
    Returns {canonical_name: sc_site_id}.
    """
    from fuzzywuzzy import process

    print("Fetching SafetyCulture site list…")
    sc_sites = client.list_sites()
    sc_map = {s["name"]: s["id"] for s in sc_sites}
    sc_names = list(sc_map.keys())

    result = {}
    for canon in canonical_names:
        match, score = process.extractOne(canon, sc_names)
        if score >= 60:
            result[canon] = sc_map[match]
            print(f"  Matched '{canon}' → '{match}' (score={score})")
        else:
            print(f"  ⚠️  No SC site found for '{canon}' (best match: '{match}' score={score})")
    return result


def run(excel_path: str, pdf_path: str, week_ref: str | None = None):
    token = os.environ.get("SC_API_TOKEN")
    if not token:
        sys.exit("❌ SC_API_TOKEN not set. Copy .env.example to .env and fill in your token.")

    # ── Week range ─────────────────────────────────────────────────────────
    week_start, week_end = week_bounds(week_ref)
    label = week_label(week_start, week_end)
    in_progress = is_week_in_progress(week_end)
    print(f"\nWeek: {label}" + (" [IN PROGRESS – some checks not yet due]" if in_progress else ""))

    # ── Parse Excel template ───────────────────────────────────────────────
    print(f"\nReading Excel: {excel_path}")
    sites = read_sites(excel_path)
    print(f"  {len(sites)} sites found")

    for s in sites:
        if s["ambiguity"]:
            print(f"  {s['ambiguity']}")

    # ── Parse PDF equipment register ───────────────────────────────────────
    print(f"\nParsing PDF: {pdf_path}")
    pdf_equipment = parse_pdf(pdf_path)
    print(f"  {len(pdf_equipment)} sites with equipment data")

    # ── Connect to SafetyCulture ───────────────────────────────────────────
    client = SCClient(token)

    canonical_names = [s["canonical"] for s in sites]
    sc_site_map = build_sc_site_map(client, canonical_names)

    # ── Fetch all audits for the week ──────────────────────────────────────
    sc_site_ids = list(set(sc_site_map.values()))
    print(f"\nFetching audits {label}…")
    audits = client.search_audits(
        date_from=week_start,
        date_to=week_end,
        site_ids=sc_site_ids if sc_site_ids else None,
    )
    print(f"  {len(audits)} audit(s) retrieved")

    # Group audits by SC site ID
    audits_by_site: dict[str, list] = defaultdict(list)
    for audit in audits:
        sid = (audit.get("site") or {}).get("id", "")
        if sid:
            audits_by_site[sid].append(audit)

    # ── Build results row by row ───────────────────────────────────────────
    results = []
    for site_info in sites:
        canon = site_info["canonical"]
        sc_id = sc_site_map.get(canon)
        site_audits = audits_by_site.get(sc_id, []) if sc_id else []

        # Equipment for this site from PDF
        eq = pdf_equipment.get(canon, {})

        # Classify each audit into a check column
        checks_found: dict[str, list[str]] = defaultdict(list)
        for audit in site_audits:
            tpl_name = audit.get("template_name") or audit.get("name") or ""
            col = classify_template(tpl_name)
            if col:
                checks_found[col].append(audit_status(audit))

        # Collapse multiple audits → best status (✅ > 🔄)
        checks: dict[str, str] = {}
        for col in ["PUWER", "LOLER", "SITE SUP", "EXCAVATOR", "DUMPER",
                    "ROLLER", "TELEHAND", "HAVS", "TOOLBOX", "RAMS"]:
            statuses = checks_found.get(col, [])

            # Mark N/A for plant cols when no equipment on site
            if col in PLANT_COLS and not eq.get(PLANT_COLS[col]):
                checks[col] = STATUS_NA
                continue

            if not statuses:
                # Don't flag as missing if week is still in progress and it's
                # a daily check that could come in later
                checks[col] = STATUS_MISSING
            elif STATUS_COMPLETE in statuses:
                checks[col] = STATUS_COMPLETE
            else:
                checks[col] = STATUS_IN_PROGRESS

        # Build notes string
        notes_parts = []
        if not sc_id:
            notes_parts.append("⚠️ No SC site matched")
        if eq:
            notes_parts.append(summarise_equipment(eq))
        if site_info["ambiguity"]:
            notes_parts.append(site_info["ambiguity"])

        results.append({
            **site_info,
            "checks": checks,
            "notes":  " | ".join(notes_parts),
        })

    # ── Write output Excel ─────────────────────────────────────────────────
    in_path = Path(excel_path)
    safe_label = label.replace(" ", "_").replace("–", "-")
    out_path = in_path.parent / f"{in_path.stem}_FILLED_{safe_label}.xlsx"
    write_results(str(in_path), str(out_path), results, week_label=label)

    # ── Print summary to console ───────────────────────────────────────────
    print(f"\n{'Site':<35} {'PUWER':^7} {'LOLER':^7} {'SUP':^5} {'EXC':^5} "
          f"{'DUMP':^5} {'ROLL':^5} {'TL':^5} {'HAVS':^5} {'TBT':^5} {'RAMS':^5}")
    print("-" * 100)
    for r in results:
        c = r["checks"]
        print(
            f"{r['site']:<35} "
            f"{c.get('PUWER','?'):^7} {c.get('LOLER','?'):^7} {c.get('SITE SUP','?'):^5} "
            f"{c.get('EXCAVATOR','?'):^5} {c.get('DUMPER','?'):^5} {c.get('ROLLER','?'):^5} "
            f"{c.get('TELEHAND','?'):^5} {c.get('HAVS','?'):^5} {c.get('TOOLBOX','?'):^5} "
            f"{c.get('RAMS','?'):^5}"
        )

    missing_sc = [r["site"] for r in results if "⚠️ No SC site matched" in r.get("notes", "")]
    if missing_sc:
        print(f"\n⚠️  Sites not matched in SafetyCulture ({len(missing_sc)}): {', '.join(missing_sc)}")

    return results


def main():
    parser = argparse.ArgumentParser(description="Ashvale weekly compliance checker")
    parser.add_argument("--excel", required=True, help="Path to weekly Excel template (.xlsx)")
    parser.add_argument("--pdf",   required=True, help="Path to WHAT'S OUT WHERE PDF")
    parser.add_argument("--week",  default=None,
                        help="Any date in target week YYYY-MM-DD (default: current week)")
    args = parser.parse_args()

    run(args.excel, args.pdf, args.week)


if __name__ == "__main__":
    main()

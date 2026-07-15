"""
Maps Excel site names / job numbers to SafetyCulture site IDs.
SC site IDs are fetched live from the API on first run and cached.
"""
from __future__ import annotations

from fuzzywuzzy import process

# Excel short name → canonical site name used in the PDF / SC
# Keys are lowercased for matching.
KNOWN_SITES = {
    "vandome": "Austin Road, Twickenham [BG0217]",
    "austin road": "Austin Road, Twickenham [BG0217]",
    "bg217": "Austin Road, Twickenham [BG0217]",
    "pollards": "Pollards, Maple Cross [BG0219]",
    "maple cross": "Pollards, Maple Cross [BG0219]",
    "bg219": "Pollards, Maple Cross [BG0219]",
    "queens drive": "The Queens Drive, Mill End [BG0218]",
    "mill end": "The Queens Drive, Mill End [BG0218]",
    "bg218": "The Queens Drive, Mill End [BG0218]",
    "whittings": "Whitings Rd, Barnet [BG0221]",
    "whitings": "Whitings Rd, Barnet [BG0221]",
    "bg221": "Whitings Rd, Barnet [BG0221]",
    "moxon": "Moxon Street, Barnet [BG222]",
    "bg222": "Moxon Street, Barnet [BG222]",
    "jackson": "Jackson Rd, Aylesbury [BG0223]",
    "bg223": "Jackson Rd, Aylesbury [BG0223]",
    "rabans": "Rabans Lane, Aylesbury [AB0023]",
    "ab023": "Rabans Lane, Aylesbury [AB0023]",
    "ab0023": "Rabans Lane, Aylesbury [AB0023]",
    "berryfields": "Berryfields, Paradise Orchard",
    "cromwell": "Cromwell Place, Wixams",
    "wixams": "Cromwell Place, Wixams",
    "loverose": "Loverose Way, Wixams [BH006]",
    "bh006": "Loverose Way, Wixams [BH006]",
    "bh04": "Loverose Way, Wixams [BH006]",
    "waterbeach": "Denny End Rd, Waterbeach [LP0010]",
    "denny end": "Denny End Rd, Waterbeach [LP0010]",
    "lp10": "Denny End Rd, Waterbeach [LP0010]",
    "lp0010": "Denny End Rd, Waterbeach [LP0010]",
    "sawtry": "Bowland Close, Sawtry [LP009]",
    "lp09": "Bowland Close, Sawtry [LP009]",
    "lp009": "Bowland Close, Sawtry [LP009]",
    "wintringham": "Wintringham, St Neots",
    "st neots": "Wintringham, St Neots",
    "lp08": "Wintringham, St Neots",
    "goodwin": "Goodwin Drive, Arlesey",
    "wh05": "Goodwin Drive, Arlesey",
    "linmere 2": "Bedford Rd, Linmere [DAN005]",
    "bedford rd": "Bedford Rd, Linmere [DAN005]",
    "dan05": "Bedford Rd, Linmere [DAN005]",
    "dn 05": "Bedford Rd, Linmere [DAN005]",
    "dan005": "Bedford Rd, Linmere [DAN005]",
    "groveway": "Groveway, Walton, MK [DAN004]",
    "walton": "Groveway, Walton, MK [DAN004]",
    "dan04": "Groveway, Walton, MK [DAN004]",
    "dan004": "Groveway, Walton, MK [DAN004]",
    "linmere": "Waterslade Way, Linmere [DAN003]",
    "waterslade": "Waterslade Way, Linmere [DAN003]",
    "dan03": "Waterslade Way, Linmere [DAN003]",
    "dan003": "Waterslade Way, Linmere [DAN003]",
    "shenley": "Wildacre Road, Shenley Wood",
    "wildacre": "Wildacre Road, Shenley Wood",
    "dan02": "Wildacre Road, Shenley Wood",
    "pound farm": "London Rd, St Ippolyts",
    "eh002": "London Rd, St Ippolyts",
    "st ippolyts": "London Rd, St Ippolyts",
    "graves paddock": "Graves Paddock, Arlesey [CA006]",
    "ca006": "Graves Paddock, Arlesey [CA006]",
    "arlesey road": "Arlesey Road, Stotfold [CA0005]",
    "stotfold": "Arlesey Road, Stotfold [CA0005]",
    "ca005": "Arlesey Road, Stotfold [CA0005]",
    "ca0005": "Arlesey Road, Stotfold [CA0005]",
    "hawkshead": "Hawkeshead Road, Little Heath",
    "hawkeshead": "Hawkeshead Road, Little Heath",
    "little heath": "Hawkeshead Road, Little Heath",
    "ca004": "Hawkeshead Road, Little Heath",
    "shefford": "Shefford Road, Clifton",
    "clifton": "Shefford Road, Clifton",
    "wh17": "Shefford Road, Clifton",
    "silwood": "Silwood Park, Ascot",
    "ascot": "Silwood Park, Ascot",
    "bc01": "Silwood Park, Ascot",
    "wavell": "Wavell Rd, Beaconsfield",
    "beaconsfield": "Wavell Rd, Beaconsfield",
    "we004": "Wavell Rd, Beaconsfield",
    "meadows close": "Meadows Close",
    "mh014": "Meadows Close",
    "ardleigh": "Meadows Close",
}


def resolve_site_name(excel_site: str, excel_job: str) -> str:
    """Return canonical site name from Excel site + job number fields."""
    candidates = []
    if excel_site:
        candidates.append(excel_site.strip().lower())
    if excel_job:
        candidates.append(str(excel_job).strip().lower())

    for c in candidates:
        if c in KNOWN_SITES:
            return KNOWN_SITES[c]
        # partial key match
        for key, val in KNOWN_SITES.items():
            if key in c or c in key:
                return val

    # fuzzy fallback
    all_keys = list(KNOWN_SITES.keys())
    for c in candidates:
        match, score = process.extractOne(c, all_keys)
        if score >= 70:
            return KNOWN_SITES[match]

    return excel_site  # return raw if no match


# Two-Linmere disambiguation warning
LINMERE_SITES = {
    "DAN003": "Waterslade Way, Linmere [DAN003]",
    "DAN005": "Bedford Rd, Linmere [DAN005]",
}


def flag_linmere_ambiguity(excel_site: str, excel_job: str) -> str | None:
    s = (excel_site or "").lower()
    j = str(excel_job or "").lower()
    if "linmere" in s and not any(k in j for k in ("dan05", "dan005", "dan03", "dan003", "dn 05")):
        return "⚠️ Ambiguous Linmere — specify DAN003 (Waterslade Way) or DAN005 (Bedford Rd)"
    return None

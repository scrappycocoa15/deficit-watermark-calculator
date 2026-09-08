#!/usr/bin/env python3
"""
Deficit Watermark Calculator — Streamlit App  (v17)
Connects to Salesforce via Session ID, runs Sales + Terminations reports (+ optional
CSEs), computes FIFO watermarks, and displays results with Team / Leader / Rep filters.

Org-chart hierarchy (15 leaders, 100 reps) is hardcoded from
"2026 US SMB Sales Org Chart.xlsx" and used as the primary source for
Team / Leader assignments — no CSEs report is required for those filters.
CSEs is still optional and used only for the Excel team-sheet split.
"""

import os, sys, io, time, calendar, requests
from datetime import date as _date
import streamlit as st
import pandas as pd
from datetime import date

# ── Import core logic from deficit_calculator ─────────────────────────────────
APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from deficit_calculator import (
    process_accounts, write_excel,
    SALES_ID_COL, SALES_AMOUNT_COL, SALES_DATE_COL, SALES_OWNER_COL, SALES_NAME_COL,
    TERMS_ID_COL, TERMS_AMOUNT_COL, TERMS_DATE_COL, TERMS_OWNER_COL, TERMS_NAME_COL,
    CSES_OWNER_COL, CSES_MANAGER_COL, CSES_TEAM_COL,
)

SF_API_VERSION = "v59.0"

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Deficit Watermark Calculator",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# Configuration  — read from Streamlit secrets with hardcoded fallbacks.
# To update without touching code:
#   Streamlit Cloud → App settings → Secrets → add/edit the keys below.
# ─────────────────────────────────────────────────────────────────────────────
SALES_REPORT_ID  = st.secrets.get("SALES_REPORT_ID",  "00OPg00000BSGzZ")
TERMS_REPORT_ID  = st.secrets.get("TERMS_REPORT_ID",  "00OPg00000BSH7d")

# ── Org-chart hierarchy — hardcoded from "2026 US SMB Sales Org Chart.xlsx" ──
# LEADER_META  : leader name  →  {team, region}
# ORG_HIERARCHY: rep name     →  {team, leader}
# Used as the primary source for Team / Leader filter columns in master_df.
# CSEs Salesforce report is still used as a fallback for reps not listed here.
LEADER_META = {
    # VPs — shown as Leader when an RSD owns an account directly
    "Jason Rainey"    : {"team": "US SMB Client Sales Premier",  "region": "Premier"},
    "Andrew Hop"      : {"team": "US SMB Client Sales Strategic", "region": "Strategic"},
    "Peter Gadd"      : {"team": "US SMB Client Sales Key",       "region": "Key"},
    # RSDs — shown as Leader for their reps' accounts
    "Angie Koplan"    : {"team": "US SMB Client Sales Premier",  "region": "Premier West"},
    "Chris Smith"     : {"team": "US SMB Client Sales Premier",  "region": "Premier North"},
    "Kyle Loving"     : {"team": "US SMB Client Sales Premier",  "region": "Premier East"},
    "Brooke Nelson"   : {"team": "US SMB Client Sales Premier",  "region": "Premier South"},
    "Megan Frodge"    : {"team": "US SMB Client Sales Strategic", "region": "Strategic Northeast"},
    "Blake Karnes"    : {"team": "US SMB Client Sales Strategic", "region": "Strategic Northwest"},
    "Amanda Meek"     : {"team": "US SMB Client Sales Strategic", "region": "Strategic Mountain"},
    "Samantha Young"  : {"team": "US SMB Client Sales Strategic", "region": "Strategic Southeast"},
    "Dave Elinger"    : {"team": "US SMB Client Sales Strategic", "region": "Strategic Southwest"},
    "Christian Larson": {"team": "US SMB Client Sales Strategic", "region": "Strategic Midwest"},
    "Heather Lewis"   : {"team": "US SMB Client Sales Key",       "region": "Key West"},
    "Ronit Cohn"      : {"team": "US SMB Client Sales Key",       "region": "Key North"},
    "Jake Rutenbar"   : {"team": "US SMB Client Sales Key",       "region": "Key Mid Atlantic"},
    "Randi Kruger"    : {"team": "US SMB Client Sales Key",       "region": "Key South"},
    "Marissa Mock"    : {"team": "US SMB Client Sales Key",       "region": "Key Great Lakes"},
}

# Pre-build a lowercase-keyed lookup for fast, case-insensitive rep matching.
_ORG_LOOKUP = {}
_RAW_ORG = {
    "Danny Lewis":              {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},
    "Jon Salmon":               {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},
    "Katie Byrnes":             {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},
    "Julie Barter":             {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},
    "Sam Angelo":               {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},
    "Alex Capeloto":            {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},
    "Adam Sala":                {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},
    "Kim Koehn":                {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},
    "Deb Tegan":                {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},
    "Ashiey McCue":             {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},
    "Mandy Gallaner":           {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},
    "Katrina Brock":            {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},
    "Ben Goman":                {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},
    "Steve Snediker":           {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},
    "Deb Smith":                {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},
    "Ellen Bastian":            {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},
    "Natalie Jamieson":         {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},
    "Rob Meek":                 {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},
    "Breck Hansen":             {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},
    "Aghiles Benali":           {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},
    "Blair Sievert":            {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},
    "Alanna Parrott":           {"team": "US SMB Client Sales Premier",  "leader": "Brooke Nelson"},
    "Meaghan Rodgers":          {"team": "US SMB Client Sales Premier",  "leader": "Brooke Nelson"},
    "Robert Stoering":          {"team": "US SMB Client Sales Premier",  "leader": "Brooke Nelson"},
    "Michael Landes":           {"team": "US SMB Client Sales Premier",  "leader": "Brooke Nelson"},
    "Amy Lawrence":             {"team": "US SMB Client Sales Premier",  "leader": "Brooke Nelson"},
    "Lindsay Wilson":           {"team": "US SMB Client Sales Premier",  "leader": "Brooke Nelson"},
    "Jun Park":                 {"team": "US SMB Client Sales Strategic", "leader": "Megan Frodge"},
    "Emily Norris":             {"team": "US SMB Client Sales Strategic", "leader": "Megan Frodge"},
    "Lindsay Paxton":           {"team": "US SMB Client Sales Strategic", "leader": "Megan Frodge"},
    "Jessica Klein":            {"team": "US SMB Client Sales Strategic", "leader": "Megan Frodge"},
    "Matt Knight":              {"team": "US SMB Client Sales Strategic", "leader": "Megan Frodge"},
    "Jill Desjardine":          {"team": "US SMB Client Sales Strategic", "leader": "Megan Frodge"},
    "Evan Smith":               {"team": "US SMB Client Sales Strategic", "leader": "Megan Frodge"},
    "Mike Monello":             {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},
    "Mark Workman":             {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},
    "Debbie Saysanavanophet":   {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},
    "Evan Anderson":            {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},
    "Britt Whims":              {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},
    "Kaitlin Dailey":           {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},
    "Jamie Stewart":            {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},
    "Tom Wahl":                 {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},
    "Kevin Chheang":            {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},
    "Joe Dorey":                {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},
    "Andrea Flor":              {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},
    "Mara Obermeier":           {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},
    "Ty Sataaf":                {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},
    "Nate Heussner":            {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},
    "Joe Silva":                {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},
    "Tyler Nuquay":             {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},
    "Trevor Hecht":             {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},
    "Kylie Barrett":            {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},
    "Jordan Buri":              {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},
    "Tyler Hazen":              {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},
    "Lauren Pellowski":         {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},
    "Sam O'Connell":            {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},
    "Aaron Korus":              {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},
    "Colin Kraker":             {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},
    "Joseph Zangel":            {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},
    "Nicole Brightenstein":     {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},
    "Zack Scharf":              {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},
    "Tom Osterberg":            {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},
    "Anna Christofaro":         {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},
    "Joel Segall":              {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},
    "Dan Eagen":                {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},
    "Rachel Burns":             {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},
    "Laura Jungbauer":          {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},
    "Jeff Danner":              {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},
    "Tyler Klein":              {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},
    "Michaela Gormley":         {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},
    "Oliver Holdenson":         {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},
    "Joe Vigil":                {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},
    "Austin Aghamirzai":        {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},
    "Joe Bellefeuille":         {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},
    "Mark Hemmerle":            {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},
    "Natalie Rizk":             {"team": "US SMB Client Sales Key",       "leader": "Ronit Cohn"},
    "Ryan Doyle":               {"team": "US SMB Client Sales Key",       "leader": "Ronit Cohn"},
    "Alden Martinez":           {"team": "US SMB Client Sales Key",       "leader": "Ronit Cohn"},
    "Teylen Sheesley":          {"team": "US SMB Client Sales Key",       "leader": "Ronit Cohn"},
    "Jonathan Barth":           {"team": "US SMB Client Sales Key",       "leader": "Ronit Cohn"},
    "Brooke Mullis":            {"team": "US SMB Client Sales Key",       "leader": "Ronit Cohn"},
    "Mike Antkowiak":           {"team": "US SMB Client Sales Key",       "leader": "Ronit Cohn"},
    "Mckenzie Bowen":           {"team": "US SMB Client Sales Key",       "leader": "Jake Rutenbar"},
    "Karrah Manzanarez":        {"team": "US SMB Client Sales Key",       "leader": "Jake Rutenbar"},
    "Thang Nguyen":             {"team": "US SMB Client Sales Key",       "leader": "Jake Rutenbar"},
    "Danny Sinatro":            {"team": "US SMB Client Sales Key",       "leader": "Jake Rutenbar"},
    "Kelsey Fredrickson":       {"team": "US SMB Client Sales Key",       "leader": "Jake Rutenbar"},
    "Ben Angelo":               {"team": "US SMB Client Sales Key",       "leader": "Jake Rutenbar"},
    "Tyler Sanford":            {"team": "US SMB Client Sales Key",       "leader": "Randi Kruger"},
    "Rachel Meyer":             {"team": "US SMB Client Sales Key",       "leader": "Randi Kruger"},
    "Scott Bere":               {"team": "US SMB Client Sales Key",       "leader": "Randi Kruger"},
    "Tom Larson":               {"team": "US SMB Client Sales Key",       "leader": "Randi Kruger"},
    "David Jensen":             {"team": "US SMB Client Sales Key",       "leader": "Randi Kruger"},
    "Jack Zabel":               {"team": "US SMB Client Sales Key",       "leader": "Randi Kruger"},
    "Chris Spencer":            {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},
    "Bri Basolo":               {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},
    "Tyler Krob":               {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},
    "Libby Hartnagel":          {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},
    "Jake Nickoloff":           {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},
    "Anders Halvorson":         {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},

    # ── Salesforce full-name / spelling aliases ───────────────────────────────
    # Salesforce stores legal names; the org chart uses nicknames or has minor
    # spelling differences.  Both forms are kept so either will match.
    # Confirmed from live data (Sep 2026):
    "Ashley McCue":             {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},       # org: Ashiey (typo)
    "Anders Halvorsen":         {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},      # org: Halvorson
    "Brianna Basolo":           {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},      # org: Bri
    "Brittany Whims":           {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},      # org: Britt
    "Christopher Spencer":      {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},      # org: Chris
    "Debbie Saysanavongphet":   {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},      # org: Saysanavanophet
    "Jacob Nickoloff":          {"team": "US SMB Client Sales Key",       "leader": "Marissa Mock"},      # org: Jake
    "Jeffrey Danner":           {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},  # org: Jeff
    "Joseph Silva":             {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},    # org: Joe
    "Mackenzie Bowen":          {"team": "US SMB Client Sales Key",       "leader": "Jake Rutenbar"},     # org: Mckenzie
    "Mandy Gallanar":           {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},        # org: Gallaner
    "Michael Monello":          {"team": "US SMB Client Sales Strategic", "leader": "Blake Karnes"},      # org: Mike
    "Nathaniel Heussner":       {"team": "US SMB Client Sales Strategic", "leader": "Samantha Young"},    # org: Nate
    # Proactive aliases — same nickname patterns, not yet confirmed but low risk:
    "Jonathan Salmon":          {"team": "US SMB Client Sales Premier",  "leader": "Angie Koplan"},       # org: Jon
    "Robert Meek":              {"team": "US SMB Client Sales Premier",  "leader": "Kyle Loving"},        # org: Rob
    "Joseph Dorey":             {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},       # org: Joe
    "Joseph Vigil":             {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},     # org: Joe
    "Joseph Bellefeuille":      {"team": "US SMB Client Sales Key",       "leader": "Heather Lewis"},     # org: Joe
    "Daniel Eagen":             {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},  # org: Dan
    "Thomas Wahl":              {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},       # org: Tom
    "Thomas Osterberg":         {"team": "US SMB Client Sales Strategic", "leader": "Christian Larson"},  # org: Tom
    "Thomas Larson":            {"team": "US SMB Client Sales Key",       "leader": "Randi Kruger"},      # org: Tom

    # ── RSDs as account owners — their leader is their VP ────────────────────
    "Angie Koplan":             {"team": "US SMB Client Sales Premier",  "leader": "Jason Rainey"},
    "Chris Smith":              {"team": "US SMB Client Sales Premier",  "leader": "Jason Rainey"},
    "Kyle Loving":              {"team": "US SMB Client Sales Premier",  "leader": "Jason Rainey"},
    "Brooke Nelson":            {"team": "US SMB Client Sales Premier",  "leader": "Jason Rainey"},
    "Megan Frodge":             {"team": "US SMB Client Sales Strategic", "leader": "Andrew Hop"},
    "Blake Karnes":             {"team": "US SMB Client Sales Strategic", "leader": "Andrew Hop"},
    "Amanda Meek":              {"team": "US SMB Client Sales Strategic", "leader": "Andrew Hop"},
    "Samantha Young":           {"team": "US SMB Client Sales Strategic", "leader": "Andrew Hop"},
    "Dave Elinger":             {"team": "US SMB Client Sales Strategic", "leader": "Andrew Hop"},
    "Christian Larson":         {"team": "US SMB Client Sales Strategic", "leader": "Andrew Hop"},
    "Heather Lewis":            {"team": "US SMB Client Sales Key",       "leader": "Peter Gadd"},
    "Ronit Cohn":               {"team": "US SMB Client Sales Key",       "leader": "Peter Gadd"},
    "Jake Rutenbar":            {"team": "US SMB Client Sales Key",       "leader": "Peter Gadd"},
    "Randi Kruger":             {"team": "US SMB Client Sales Key",       "leader": "Peter Gadd"},
    "Marissa Mock":             {"team": "US SMB Client Sales Key",       "leader": "Peter Gadd"},

    # ── Additional Salesforce spelling aliases (confirmed Sep 2026) ───────────
    "Nicole Breitenstein":      {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},     # org: Brightenstein
    "Stephen Snediker":         {"team": "US SMB Client Sales Premier",  "leader": "Chris Smith"},       # org: Steve
    "Ty Saathoff":              {"team": "US SMB Client Sales Strategic", "leader": "Amanda Meek"},      # org: Sataaf
    "Samantha O'Connell":       {"team": "US SMB Client Sales Strategic", "leader": "Dave Elinger"},     # org: Sam
}
_ORG_LOOKUP = {k.lower(): v for k, v in _RAW_ORG.items()}

# ─────────────────────────────────────────────────────────────────────────────
# Styles
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: '72', Arial, sans-serif; }

.app-header {
    background: #0070F2;
    color: #fff;
    padding: 1rem 1.5rem;
    border-radius: 8px;
    margin-bottom: 1.25rem;
}
.app-header h1 { margin: 0; font-size: 1.35rem; font-weight: 700; color: #fff; }
.app-header p  { margin: 0.2rem 0 0; font-size: 0.82rem; opacity: 0.88; }

.kpi-wrap { display: flex; gap: 0.85rem; margin: 0.75rem 0 1.25rem; }
.kpi-card {
    flex: 1;
    background: #E1F4FF;
    border-left: 4px solid #0070F2;
    border-radius: 6px;
    padding: 0.75rem 1rem;
}
.kpi-val { font-size: 1.75rem; font-weight: 700; color: #00144A; line-height: 1.15; }
.kpi-lbl { font-size: 0.72rem; color: #555; margin-top: 4px; letter-spacing: 0.04em; text-transform: uppercase; }

.filter-wrap {
    background: #EAECEE;
    border-radius: 8px;
    padding: 0.75rem 1rem 0.25rem;
    margin-bottom: 0.75rem;
}

section[data-testid="stSidebar"] { background: #f5f6f7; }
section[data-testid="stSidebar"] .stMarkdown h3 { color: #00144A; }

.stButton > button {
    background: #0070F2 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
}
.stButton > button:hover { background: #0134BF !important; }
.stButton > button:disabled {
    background: #EAECEE !important;
    color: #aaa !important;
    cursor: not-allowed !important;
}
.stDownloadButton > button {
    background: #fff !important;
    color: #0070F2 !important;
    border: 1.5px solid #0070F2 !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
}
.stDownloadButton > button:hover {
    background: #E1F4FF !important;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Salesforce helpers
# ─────────────────────────────────────────────────────────────────────────────

def _hdr(session_id: str) -> dict:
    return {"Authorization": f"Bearer {session_id}", "Content-Type": "application/json"}


def detect_instance_url(session_id: str):
    """
    Auto-detect the Salesforce instance URL from a session ID alone.

    Calls the standard OAuth2 userinfo endpoint at login.salesforce.com (and
    test.salesforce.com as a fallback for sandboxes).  Salesforce returns the
    org's REST base URL in the response regardless of whether the org uses a
    custom domain, so the correct instance URL is always recovered.

    Returns (ok: bool, instance_url: str, info: dict).
    """
    from urllib.parse import urlparse
    for base in ("https://login.salesforce.com", "https://test.salesforce.com"):
        try:
            r = requests.get(
                f"{base}/services/oauth2/userinfo",
                headers=_hdr(session_id),
                timeout=10,
                allow_redirects=True,
            )
            if r.status_code == 200:
                info = r.json()
                # urls.rest is "https://<instance>/services/data/" — parse the origin
                rest = info.get("urls", {}).get("rest", "")
                if rest:
                    p = urlparse(rest)
                    return True, f"{p.scheme}://{p.netloc}", info
        except Exception:
            pass
    return False, "", {}


def search_reports(instance_url: str, session_id: str, term: str) -> list:
    term_safe = term.replace("'", "\\'")
    q = f"SELECT Id, Name FROM Report WHERE Name LIKE '%{term_safe}%' LIMIT 30"
    r = requests.get(
        f"{instance_url}/services/data/{SF_API_VERSION}/query",
        headers=_hdr(session_id), params={"q": q}, timeout=15,
    )
    r.raise_for_status()
    return r.json().get("records", [])


def _parse_sf_number(raw: str):
    """
    Parse a Salesforce-formatted number string to float.
    Handles:
      "$1,234.56"   →  1234.56
      "(1,234.56)"  →  -1234.56   (accounting notation for negatives)
      "($1,234.56)" →  -1234.56
      "-1234.56"    →  -1234.56
    Returns None if unparseable.
    """
    s = str(raw).replace(",", "").replace("$", "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _parse_report_response(data: dict):
    """
    Parse an Analytics API report instance response → (DataFrame, all_data: bool).

    factMap layout differs by report format:
      TABULAR  — all detail rows under "T!T"
      SUMMARY  — detail rows distributed across group keys ("0!T", "0_0!T", …);
                 "T!T" holds only grand-total *aggregates* (no rows).

    This function iterates every factMap node and collects all rows,
    so it works for both formats without needing to know which one is in use.
    """
    meta        = data.get("reportMetadata", {})
    ext         = data.get("reportExtendedMetadata", {})
    detail_cols = meta.get("detailColumns", [])
    col_info    = ext.get("detailColumnInfo", {})
    col_labels  = [col_info.get(c, {}).get("label", c) for c in detail_cols]
    numeric_t   = {"currency", "double", "int", "percent", "number"}
    col_types   = {c: col_info.get(c, {}).get("dataType", "string") for c in detail_cols}

    # Collect rows from every node in the factMap (handles tabular + summary).
    all_rows = []
    for node in data.get("factMap", {}).values():
        if isinstance(node, dict):
            all_rows.extend(node.get("rows", []))

    parsed = []
    for row in all_rows:
        pr = []
        for i, cell in enumerate(row.get("dataCells", [])):
            col_key = detail_cols[i] if i < len(detail_cols) else ""
            dtype   = col_types.get(col_key, "string")
            if dtype in numeric_t:
                val = cell.get("value")
                # Multi-currency fields return a dict: {"amount": 21168, "currency": "USD"}
                # Extract the numeric amount from it.
                if isinstance(val, dict):
                    val = val.get("amount")
                # value=null fallback: parse the formatted label string
                if val is None:
                    val = _parse_sf_number(cell.get("label", ""))
                pr.append(val)
            else:
                pr.append(cell.get("label"))
        parsed.append(pr)

    return pd.DataFrame(parsed, columns=col_labels), data.get("allData", True)


@st.cache_data(ttl=300, show_spinner=False)
def run_sf_report(instance_url: str, session_id: str, report_id: str,
                  date_col_label: str = ""):
    """
    Fetch ALL report rows — no 2,000-row ceiling.  Three-tier strategy:

    Tier 1 — SOQL + pagination
        GET the report metadata to discover the Salesforce object type and each
        column's SOQL field name (entityColumnName).  Issues a paginated /query
        and follows every nextRecordsUrl until done.  Unlimited row count.
        Works only when every column has a real SOQL equivalent (no formula /
        converted-currency columns).

    Tier 2 — Chunked Analytics API  ← the key fix for this report
        When converted-currency or formula columns have no entityColumnName,
        SOQL is unavailable.  Instead, split the date range into monthly windows
        and execute one async report instance per window.  Each month stays well
        under the 2,000-row cap; all chunks are concatenated.
        Requires date_col_label to be the exact column header of the date field
        in the report (e.g. "Order Effective Date").  The column's Analytics ID
        is resolved from the report metadata so no hardcoding is needed.
        The report's own existing filters are preserved; only a tight date
        bracket is added on top for each chunk.

    Tier 3 — plain async Analytics API  (last resort, 2,000-row cap)
        Used only when both SOQL and chunking fail.  Caller receives
        all_data=False as a signal.

    Returns (DataFrame, all_data: bool).
    """
    h   = _hdr(session_id)
    api = f"{instance_url}/services/data/{SF_API_VERSION}/analytics/reports/{report_id}"

    # ── Tier 0: fetch report metadata (always needed) ────────────────────────
    meta_r = requests.get(api, headers=h, timeout=15)
    meta_r.raise_for_status()
    m           = meta_r.json()
    rpt_meta    = m.get("reportMetadata", {})
    ext_meta    = m.get("reportExtendedMetadata", {})
    detail_cols = rpt_meta.get("detailColumns", [])
    col_info    = ext_meta.get("detailColumnInfo", {})
    col_labels  = [col_info.get(c, {}).get("label", c) for c in detail_cols]
    obj_type    = rpt_meta.get("reportType", {}).get("type", "")
    ent_cols    = [
        col_info.get(c, {}).get("entityColumnName", "") or ""
        for c in detail_cols
    ]

    # showDetails is only valid for Summary / Matrix reports.
    # Sending it for a Tabular report causes a 400 Bad Request.
    report_fmt   = rpt_meta.get("reportFormat", "TABULAR")
    needs_detail = report_fmt in ("SUMMARY", "MATRIX")

    # ── Tier 1: SOQL path ────────────────────────────────────────────────────
    if obj_type and all(ent_cols):
        try:
            soql     = f"SELECT {', '.join(ent_cols)} FROM {obj_type}"
            all_rows = []
            next_url = None
            while True:
                if next_url:
                    resp = requests.get(f"{instance_url}{next_url}",
                                        headers=h, timeout=30)
                else:
                    resp = requests.get(
                        f"{instance_url}/services/data/{SF_API_VERSION}/query",
                        headers=h, params={"q": soql}, timeout=30,
                    )
                resp.raise_for_status()
                data = resp.json()
                for rec in data.get("records", []):
                    row = []
                    for field in ent_cols:
                        val = rec
                        for part in field.split("."):
                            val = val.get(part) if isinstance(val, dict) else None
                        row.append(val)
                    all_rows.append(row)
                if data.get("done", True):
                    break
                next_url = data.get("nextRecordsUrl")
            return pd.DataFrame(all_rows, columns=col_labels), True, 1, []
        except Exception:
            pass  # fall through

    # ── Tier 2: chunked Analytics API ────────────────────────────────────────
    # Resolve the Analytics column ID for the date field by matching its label.
    date_col_id = None
    if date_col_label:
        for col_id in detail_cols:
            if col_info.get(col_id, {}).get("label", "") == date_col_label:
                date_col_id = col_id
                break

    if date_col_id:
        today   = _date.today()
        all_dfs = []
        t2_errs = []

        # ── Tier 2: synchronous chunked Analytics API ────────────────────────
        # Key insight: sending reportFilters in a POST body REPLACES the report's
        # saved filters (including the division filter), forcing us to re-send them
        # verbatim — a combination Salesforce has been rejecting.
        #
        # standardDateFilter is a SEPARATE metadata field: setting it only
        # overrides the date range and leaves reportFilters untouched, so the
        # division filter is preserved automatically.
        for i in range(9):
            y, mo = today.year, today.month - i
            while mo <= 0:
                mo += 12; y -= 1
            _, last = calendar.monthrange(y, mo)
            start = f"{y:04d}-{mo:02d}-01"
            end   = f"{y:04d}-{mo:02d}-{last:02d}"

            chunk_meta = {
                "standardDateFilter": {
                    "column":        date_col_id,
                    "durationValue": "CUSTOM",
                    "startDate":     start,
                    "endDate":       end,
                }
            }
            if needs_detail:
                chunk_meta["showDetails"] = True

            r = requests.post(api, headers=h,
                              json={"reportMetadata": chunk_meta}, timeout=60)

            if r.status_code >= 400:
                try:
                    err_body = r.json()
                except Exception:
                    err_body = r.text[:800]
                t2_errs.append({
                    "month":  f"{y}-{mo:02d}",
                    "status": r.status_code,
                    "body":   err_body,
                })
                continue

            chunk_df, chunk_all = _parse_report_response(r.json())

            if not chunk_all:
                t2_errs.append({
                    "month":  f"{y}-{mo:02d}",
                    "status": "cap_exceeded",
                    "body":   "Single month > 2,000 rows; weekly chunking would be needed.",
                })
                all_dfs = []
                break

            if not chunk_df.empty:
                all_dfs.append(chunk_df)

        if all_dfs:
            return pd.concat(all_dfs, ignore_index=True), True, 2, t2_errs

        # All chunks failed or empty — fall through to Tier 3.

    # ── Tier 3: plain async Analytics API (2,000-row cap) ────────────────────
    t2_errs = t2_errs if "t2_errs" in dir() else []
    t3_body = {"reportMetadata": {"showDetails": True}} if needs_detail else {}
    r = requests.post(f"{api}/instances", headers=h, json=t3_body, timeout=30)
    r.raise_for_status()
    inst_id = r.json()["id"]
    for _ in range(90):
        time.sleep(2)
        r = requests.get(f"{api}/instances/{inst_id}", headers=h, timeout=30)
        r.raise_for_status()
        payload = r.json()
        status  = payload.get("attributes", {}).get("status", "")
        if status == "Success":
            df, all_data = _parse_report_response(payload)
            return df, all_data, 3, t2_errs
        if status in ("Error", "Cancelled"):
            code = payload.get("attributes", {}).get("errorCode", "Unknown")
            raise RuntimeError(f"Report failed: {code}")
    raise TimeoutError("Report timed out after 3 minutes.")


# ─────────────────────────────────────────────────────────────────────────────
# Data prep (DataFrame-based, replacing the file-path loaders)
# ─────────────────────────────────────────────────────────────────────────────

def _coerce_numeric(s: pd.Series) -> pd.Series:
    """
    Convert a Series to numeric, handling all Salesforce Analytics API formats:
      21168                          →  21168.0   (int from multi-currency dict)
      {"amount": 21168, ...}         →  21168.0   (dict not yet unwrapped)
      "$1,234.56"                    →  1234.56
      "(1,234.56)" / "($1,234.56)"   →  -1234.56  (accounting notation)
    """
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    def _extract(x):
        if isinstance(x, dict):          # {"amount": 21168, "currency": "USD"}
            return x.get("amount")
        return _parse_sf_number(x)

    return pd.to_numeric(s.map(_extract), errors="coerce")


def prep_sales(df: pd.DataFrame) -> pd.DataFrame:
    need = [SALES_ID_COL, SALES_AMOUNT_COL, SALES_DATE_COL, SALES_OWNER_COL, SALES_NAME_COL]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"Sales report is missing columns: {miss}\nColumns found: {list(df.columns)}")
    df = df.rename(columns={
        SALES_ID_COL: "account_id", SALES_AMOUNT_COL: "amount",
        SALES_DATE_COL: "effective_date", SALES_OWNER_COL: "account_owner",
        SALES_NAME_COL: "account_name",
    })[["account_id", "amount", "effective_date", "account_owner", "account_name"]].copy()
    df["source"]         = "Sales"
    df["amount"]         = _coerce_numeric(df["amount"])
    df["effective_date"] = pd.to_datetime(df["effective_date"], errors="coerce")
    return df[
        df["amount"].notna() & (df["amount"] != 0)
        & df["effective_date"].notna()
        & df["account_id"].notna()
        & (df["account_id"].astype(str).str.strip() != "")
    ].reset_index(drop=True)


def prep_terms(df: pd.DataFrame) -> pd.DataFrame:
    need = [TERMS_ID_COL, TERMS_AMOUNT_COL, TERMS_DATE_COL, TERMS_OWNER_COL, TERMS_NAME_COL]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"Terminations report is missing columns: {miss}\nColumns found: {list(df.columns)}")
    df = df.rename(columns={
        TERMS_ID_COL: "account_id", TERMS_AMOUNT_COL: "amount",
        TERMS_DATE_COL: "effective_date", TERMS_OWNER_COL: "account_owner",
        TERMS_NAME_COL: "account_name",
    })[["account_id", "amount", "effective_date", "account_owner", "account_name"]].copy()
    df["source"]         = "Terminations"
    df["amount"]         = _coerce_numeric(df["amount"]).abs() * -1
    df["effective_date"] = pd.to_datetime(df["effective_date"], errors="coerce")
    return df[
        df["amount"].notna() & (df["amount"] != 0)
        & df["effective_date"].notna()
        & df["account_id"].notna()
        & (df["account_id"].astype(str).str.strip() != "")
    ].reset_index(drop=True)


def prep_cses(df: pd.DataFrame) -> pd.DataFrame:
    need = [CSES_OWNER_COL, CSES_MANAGER_COL, CSES_TEAM_COL]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"CSEs report is missing columns: {miss}\nColumns found: {list(df.columns)}")
    df = df.rename(columns={
        CSES_OWNER_COL: "owner_name", CSES_MANAGER_COL: "manager_name", CSES_TEAM_COL: "team",
    })[["owner_name", "manager_name", "team"]].copy()
    df["owner_name"]   = df["owner_name"].str.strip()
    df["manager_name"] = df["manager_name"].fillna("Unassigned").str.strip()
    df["team"]         = df["team"].fillna("").str.strip()
    return (df.dropna(subset=["owner_name"])
              .drop_duplicates(subset=["owner_name"])
              .reset_index(drop=True))


# ─────────────────────────────────────────────────────────────────────────────
# Master DataFrame — one row per account, all context + deficit columns
# ─────────────────────────────────────────────────────────────────────────────

def build_master_df(active: list, cses_df: pd.DataFrame = None) -> pd.DataFrame:
    # Build CSEs fallback lookup (rep name → team / manager from Salesforce CSEs report)
    cses_lookup = {}
    if cses_df is not None:
        for _, row in cses_df.iterrows():
            cses_lookup[row["owner_name"].strip().lower()] = {
                "Team":   row.get("team", ""),
                "Leader": row.get("manager_name", "Unassigned"),
                "Region": "",
            }

    rows = []
    for r in active:
        rep       = (r["account_owner"] or "").strip()
        rep_lower = rep.lower()

        # Primary: org-chart lookup (case-insensitive exact match on cleaned name)
        org = _ORG_LOOKUP.get(rep_lower)
        if org:
            team   = org["team"]
            leader = org["leader"]
            region = LEADER_META.get(leader, {}).get("region", "")
        else:
            # Fallback: CSEs Salesforce report
            fb     = cses_lookup.get(rep_lower, {})
            team   = fb.get("Team",   "")
            leader = fb.get("Leader", "Unassigned")
            region = fb.get("Region", "")

        base = {
            "Team":         team,
            "Leader":       leader,
            "Region":       region,
            "Rep":          rep,
            "Account Name": r["account_name"],
            "Account ID":   r["account_id"],
        }
        for n, d in enumerate(r["deficits"], 1):
            base[f"Deficit {n} Amount"]         = -round(d["amount"], 2)
            base[f"Deficit {n} Effective Date"] = (
                d["eff_date"].date()
                if hasattr(d["eff_date"], "date") else d["eff_date"]
            )
            base[f"Deficit {n} Clears/Expires"] = (
                d["clear_date"].date()
                if hasattr(d["clear_date"], "date") else d["clear_date"]
            )
            base[f"Deficit {n} Source"]         = d["source"]
        rows.append(base)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    return df.sort_values(["Team", "Leader", "Rep", "Account Name"]).reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# Excel download helper
# ─────────────────────────────────────────────────────────────────────────────

def excel_bytes(results: list, cses_df=None) -> bytes:
    active = [r for r in results if r["deficits"]]
    if not active:
        return b""
    buf = io.BytesIO()
    write_excel(results, buf, cses_df=cses_df)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar — Connection + Report Selection
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### Salesforce Connection")

    session_id = st.text_input(
        "Session ID",
        type="password",
        placeholder="00D…",
        key="sf_sid",
    )

    if st.button("Connect", use_container_width=True):
        if not session_id.strip():
            st.error("Session ID is required.")
        else:
            with st.spinner("Detecting instance and verifying session…"):
                ok, instance_url, info = detect_instance_url(session_id.strip())
            if ok:
                st.session_state.update({
                    "connected":   True,
                    "sf_instance": instance_url,
                    "sf_token":    session_id.strip(),
                })
                name = info.get("name") or info.get("preferred_username", "")
                st.success(f"Connected{' — ' + name if name else ''}")
            else:
                st.session_state["connected"] = False
                st.error("Connection failed. Check your Session ID.")

    # ── CSEs + Run ───────────────────────────────────────────────────────────
    if st.session_state.get("connected"):
        st.markdown("---")
        st.markdown("### Optional: CSEs Report")
        st.caption(
            "Enables team-sheet grouping in the Excel download. "
            "Team and Leader filters always work via the built-in org chart."
        )

        def report_picker(label: str, search_key: str, select_key: str):
            """Search-then-select widget for a Salesforce report. Returns report ID or None."""
            term = st.text_input(f"Search — {label}", key=search_key,
                                 placeholder="Type part of the report name…")
            if term and len(term) >= 2:
                try:
                    recs = search_reports(
                        st.session_state["sf_instance"],
                        st.session_state["sf_token"],
                        term,
                    )
                except Exception as e:
                    st.warning(f"Search error: {e}")
                    recs = []
                if recs:
                    names = [r["Name"] for r in recs]
                    ids   = [r["Id"]   for r in recs]
                    idx   = st.selectbox(
                        f"Select — {label}",
                        range(len(names)),
                        format_func=lambda i: names[i],
                        key=select_key,
                    )
                    return ids[idx]
                st.caption("No reports found.")
            return None

        cses_id = report_picker("CSEs Report", "c_srch", "c_sel")

        st.markdown("---")
        run_btn = st.button("Run Calculation", use_container_width=True)

        if st.session_state.get("results_ready"):
            if st.button("Refresh Data", use_container_width=True):
                run_sf_report.clear()
                for k in ["results_ready", "results", "active", "master_df",
                          "cses_df", "run_date", "warnings"]:
                    st.session_state.pop(k, None)
                st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# Page header
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>Deficit Watermark Calculator</h1>
  <p>Live FIFO netting of SOF losses and terminations &mdash; 6-month watermark windows per account.</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Pre-connection / pre-run states
# ─────────────────────────────────────────────────────────────────────────────
if not st.session_state.get("connected"):
    st.info("Enter your Salesforce Session ID in the sidebar to get started.")
    with st.expander("How to get your Session ID"):
        st.markdown("""
**Option A — Developer Console**
1. Log in to Salesforce and open the gear menu → **Developer Console**.
2. Go to **Debug → Open Execute Anonymous Window**.
3. Paste and run: `System.debug(UserInfo.getSessionId());`
4. Open the log and copy the value after `DEBUG|`.

**Option B — Browser URL** *(Classic only)*
The session ID may appear in the page URL after `sid=`.

The session ID expires when you log out or after your org's session timeout (typically 8 hours).
The instance URL is detected automatically — no need to enter it manually.
        """)
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Run reports & compute watermarks
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.get("connected") and "run_btn" in dir() and run_btn:
    bar   = st.progress(0, "Fetching Sales report…")
    warns = []
    try:
        with st.spinner("Running Sales report (monthly chunks)…"):
            s_df, s_all, s_tier, s_t2errs = run_sf_report(
                st.session_state["sf_instance"], st.session_state["sf_token"],
                SALES_REPORT_ID, date_col_label=SALES_DATE_COL,
            )
        if not s_all:
            warns.append(
                "Sales report used the plain Analytics API (2,000-row cap). "
                "Monthly chunking produced no rows — see Tier 2 errors in Diagnostics "
                "for the exact Salesforce rejection reason."
            )
        bar.progress(30, "Fetching Terminations report…")

        with st.spinner("Running Terminations report (monthly chunks)…"):
            t_df, t_all, t_tier, t_t2errs = run_sf_report(
                st.session_state["sf_instance"], st.session_state["sf_token"],
                TERMS_REPORT_ID, date_col_label=TERMS_DATE_COL,
            )
        if not t_all:
            warns.append(
                "Terminations report used the plain Analytics API (2,000-row cap). "
                "Monthly chunking produced no rows — see Tier 2 errors in Diagnostics "
                "for the exact Salesforce rejection reason."
            )
        bar.progress(60, "Running CSEs report…" if cses_id else "Processing data…")

        cses_raw = None
        if cses_id:
            with st.spinner("Running CSEs report…"):
                cses_raw, _, _, _ = run_sf_report(
                    st.session_state["sf_instance"], st.session_state["sf_token"], cses_id,
                )

        bar.progress(78, "Computing watermarks…")

        sales_df  = prep_sales(s_df)
        terms_df  = prep_terms(t_df)
        cses_df   = prep_cses(cses_raw) if cses_raw is not None else None
        combined  = pd.concat([sales_df, terms_df], ignore_index=True)
        results   = process_accounts(combined)
        active    = [r for r in results if r["deficits"]]
        master_df = build_master_df(active, cses_df)

        bar.progress(100, "Done.")
        bar.empty()

        st.session_state.update({
            "results_ready": True,
            "results":       results,
            "active":        active,
            "master_df":     master_df,
            "cses_df":       cses_df,
            "run_date":      date.today(),
            "warnings":      warns,
            "diag": {
                "sales_rows":    len(s_df),
                "sales_cols":    list(s_df.columns),
                "sales_sample":  s_df.iloc[0].to_dict() if not s_df.empty else None,
                "sales_all":     s_all,
                "sales_tier":    s_tier,
                "sales_t2errs":  s_t2errs,
                "terms_rows":    len(t_df),
                "terms_cols":    list(t_df.columns),
                "terms_sample":  t_df.iloc[0].to_dict() if not t_df.empty else None,
                "terms_all":     t_all,
                "terms_tier":    t_tier,
                "terms_t2errs":  t_t2errs,
                "sales_prepped": len(sales_df),
                "terms_prepped": len(terms_df),
                "combined_rows": len(combined),
                "results_total": len(results),
                "active_total":  len(active),
            },
        })

    except Exception as e:
        bar.empty()
        st.error(f"Error: {e}")
        st.stop()


if not st.session_state.get("results_ready"):
    st.info("Click **Run Calculation** in the sidebar to load data.")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Results
# ─────────────────────────────────────────────────────────────────────────────

for w in st.session_state.get("warnings", []):
    st.warning(w)

# ── Diagnostics panel ─────────────────────────────────────────────────────────
if st.session_state.get("diag"):
    d = st.session_state["diag"]
    with st.expander("Diagnostics — expand if results look wrong", expanded=(d["active_total"] == 0)):
        c1, c2, c3 = st.columns(3)
        c1.metric("Sales rows fetched",    f"{d['sales_rows']:,}")
        c1.metric("Sales rows after prep", f"{d['sales_prepped']:,}")
        c2.metric("Terms rows fetched",    f"{d['terms_rows']:,}")
        c2.metric("Terms rows after prep", f"{d['terms_prepped']:,}")
        c3.metric("Combined rows",         f"{d['combined_rows']:,}")
        c3.metric("Accounts with deficits",f"{d['active_total']:,}")

        tier_s = "Chunked (Tier 2)" if d.get("sales_tier") == 2 else "Plain async / Tier 3 (≤2,000 rows)"
        tier_t = "Chunked (Tier 2)" if d.get("terms_tier") == 2 else "Plain async / Tier 3 (≤2,000 rows)"
        st.caption(f"Fetch strategy — Sales: **{tier_s}** | Terms: **{tier_t}**")

        # If Tier 2 failed for any chunks, show the exact Salesforce errors so
        # the filter-rejection reason is visible without guessing.
        for label, err_key in [("Sales", "sales_t2errs"), ("Terms", "terms_t2errs")]:
            errs = d.get(err_key) or []
            if errs:
                with st.expander(f"Tier 2 chunk errors — {label} ({len(errs)} failed)"):
                    st.json(errs)

        st.markdown("**Expected column names (from deficit_calculator.py)**")
        exp_s = [SALES_ID_COL, SALES_AMOUNT_COL, SALES_DATE_COL, SALES_OWNER_COL, SALES_NAME_COL]
        exp_t = [TERMS_ID_COL, TERMS_AMOUNT_COL, TERMS_DATE_COL, TERMS_OWNER_COL, TERMS_NAME_COL]
        miss_s = [c for c in exp_s if c not in d["sales_cols"]]
        miss_t = [c for c in exp_t if c not in d["terms_cols"]]

        st.markdown("*Sales report columns returned from Salesforce:*")
        st.code(", ".join(d["sales_cols"]) or "(none)")
        if miss_s:
            st.error(f"Sales: missing expected columns — {miss_s}")

        st.markdown("*Terminations report columns returned from Salesforce:*")
        st.code(", ".join(d["terms_cols"]) or "(none)")
        if miss_t:
            st.error(f"Terminations: missing expected columns — {miss_t}")

        # Raw sample — shows actual cell values so numeric-format issues are visible
        if d.get("sales_sample"):
            st.markdown("*Sales report — first raw row (before prep):*")
            st.json(d["sales_sample"])
        if d.get("terms_sample"):
            st.markdown("*Terminations report — first raw row (before prep):*")
            st.json(d["terms_sample"])

master_df = st.session_state["master_df"]
results   = st.session_state["results"]
cses_df   = st.session_state["cses_df"]
run_date  = st.session_state["run_date"]
has_cses  = cses_df is not None   # used for Excel team-sheet split via write_excel
has_org   = not master_df.empty and "Leader" in master_df.columns


# ── Filters ──────────────────────────────────────────────────────────────────
st.markdown('<div class="filter-wrap">', unsafe_allow_html=True)

# Three-level cascade: Team → Leader → Account Owner
# Leader column is always populated from the hardcoded org chart (no CSEs needed).
fc = st.columns([1, 1, 1, 1.6, 0.5])

# Level 1 — Team
all_teams = sorted(master_df["Team"].dropna().unique()) if has_org else []
sel_teams = fc[0].multiselect(
    "Account Owner Team",
    all_teams,
    placeholder="All teams",
    key="f_team",
)

# Level 2 — Leader (cascades from team; shown as "Name (Region)")
if has_org:
    ldr_pool_df = master_df[master_df["Team"].isin(sel_teams)] if sel_teams else master_df
    # Build display label: "Angie Koplan (Premier West)"
    leader_display: dict[str, str] = {}
    for ldr in ldr_pool_df["Leader"].dropna().unique():
        region_vals = ldr_pool_df[ldr_pool_df["Leader"] == ldr]["Region"].dropna()
        region_str  = region_vals.iloc[0] if not region_vals.empty else ""
        leader_display[ldr] = f"{ldr} ({region_str})" if region_str else ldr
    all_leaders_display = sorted(leader_display.values())
    display_to_leader   = {v: k for k, v in leader_display.items()}
else:
    all_leaders_display = []
    display_to_leader   = {}

sel_leaders_display = fc[1].multiselect(
    "Leader",
    all_leaders_display,
    placeholder="All leaders",
    key="f_leader",
)
sel_leaders = [display_to_leader.get(d, d) for d in sel_leaders_display]

# Level 3 — Account Owner (cascades from leader, then team, then all)
if sel_leaders:
    rep_pool = master_df[master_df["Leader"].isin(sel_leaders)]["Rep"]
elif sel_teams:
    rep_pool = master_df[master_df["Team"].isin(sel_teams)]["Rep"]
else:
    rep_pool = master_df["Rep"] if not master_df.empty else pd.Series([], dtype=str)
all_reps = sorted(rep_pool.dropna().unique())
sel_reps = fc[2].multiselect(
    "Account Owner",
    all_reps,
    placeholder="All reps",
    key="f_rep",
)

search = fc[3].text_input(
    "Search account name", placeholder="Type to search…", key="f_search"
)
if fc[4].button("Clear", use_container_width=True, key="f_clear"):
    for k in ["f_team", "f_leader", "f_rep", "f_search"]:
        st.session_state.pop(k, None)
    st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Apply filters
filtered = master_df.copy()
if sel_teams:
    filtered = filtered[filtered["Team"].isin(sel_teams)]
if sel_leaders:
    filtered = filtered[filtered["Leader"].isin(sel_leaders)]
if sel_reps:
    filtered = filtered[filtered["Rep"].isin(sel_reps)]
if search:
    filtered = filtered[filtered["Account Name"].str.contains(search, case=False, na=False)]

filtered = filtered.reset_index(drop=True)


# ── KPI cards ────────────────────────────────────────────────────────────────
amt_cols         = [c for c in filtered.columns if "Amount" in c]
total_amt        = filtered[amt_cols].apply(pd.to_numeric, errors="coerce").sum().sum()
total_watermarks = int(filtered[amt_cols].apply(pd.to_numeric, errors="coerce").notna().sum().sum())

st.markdown(f"""
<div class="kpi-wrap">
  <div class="kpi-card">
    <div class="kpi-val">{len(filtered):,}</div>
    <div class="kpi-lbl">Accounts with Deficits</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-val">${abs(total_amt):,.0f}</div>
    <div class="kpi-lbl">Total Deficit Amount</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-val">{total_watermarks:,}</div>
    <div class="kpi-lbl">Active Watermarks</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-val">{run_date.strftime("%b %d, %Y")}</div>
    <div class="kpi-lbl">Data As Of</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Table ─────────────────────────────────────────────────────────────────────
if filtered.empty:
    st.info("No accounts match the current filters.")
else:
    # ── Dynamic column trimming ───────────────────────────────────────────────
    # Only show as many Deficit N groups as the filtered rows actually contain.
    # This prevents empty Amount/Clears columns from squeezing the visible data.
    populated_ns = set()
    for col in filtered.columns:
        if col.startswith("Deficit ") and " Amount" in col:
            try:
                n = int(col.split()[1])
                if pd.to_numeric(filtered[col], errors="coerce").notna().any():
                    populated_ns.add(n)
            except (ValueError, IndexError):
                pass
    max_deficit_n = max(populated_ns) if populated_ns else 0

    # Identity columns — Team kept in master_df for filtering but not displayed
    identity_cols = [
        c for c in ["Leader", "Rep", "Account Name", "Account ID"]
        if c in filtered.columns
    ]
    # Deficit group columns in order, capped at max_deficit_n
    deficit_group_order = ("Amount", "Effective Date", "Clears/Expires", "Source")
    deficit_cols = []
    for n in range(1, max_deficit_n + 1):
        for suffix in deficit_group_order:
            cname = f"Deficit {n} {suffix}"
            if cname in filtered.columns:
                deficit_cols.append(cname)

    display_cols = identity_cols + deficit_cols

    # ── Column config ─────────────────────────────────────────────────────────
    col_cfg = {}
    col_cfg["Leader"]       = st.column_config.TextColumn("Leader",       width="medium")
    col_cfg["Rep"]          = st.column_config.TextColumn("Rep",          width="medium")
    col_cfg["Account Name"] = st.column_config.TextColumn("Account Name", width="large")
    col_cfg["Account ID"]   = st.column_config.TextColumn("Account ID",   width="medium")

    for c in display_cols:
        if "Amount" in c:
            col_cfg[c] = st.column_config.NumberColumn(c, format="$%.2f",     width="medium")
        elif "Effective Date" in c or "Clears" in c:
            col_cfg[c] = st.column_config.DateColumn(c,   format="MM/DD/YYYY", width="medium")
        elif "Source" in c:
            col_cfg[c] = st.column_config.TextColumn(c, width="small")

    row_height = 35
    table_h    = min(620, 45 + len(filtered) * row_height)

    st.dataframe(
        filtered[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config=col_cfg,
        height=table_h,
    )


# ── Downloads ─────────────────────────────────────────────────────────────────
st.markdown("---")
dl1, dl2, _ = st.columns([1.2, 1.2, 4])

with dl1:
    xls = excel_bytes(results, cses_df=cses_df)
    if xls:
        st.download_button(
            label="Download Full Excel",
            data=xls,
            file_name=f"DeficitResults_{run_date.strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.button("Download Full Excel", disabled=True, use_container_width=True)

with dl2:
    if not filtered.empty:
        # display_cols was built by the table block above (same filtering logic)
        csv_buf = io.StringIO()
        filtered[display_cols].to_csv(csv_buf, index=False)
        st.download_button(
            label="Download Filtered CSV",
            data=csv_buf.getvalue(),
            file_name=f"Deficits_filtered_{run_date.strftime('%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

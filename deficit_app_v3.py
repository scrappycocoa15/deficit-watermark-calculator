#!/usr/bin/env python3
"""
Deficit Watermark Calculator — Streamlit App
Connects to Salesforce via Session ID, runs Sales + Terminations + CSEs reports,
computes FIFO watermarks, and displays results with Team / Manager / Rep filters.
"""

import os, sys, io, time, requests
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
_DEFAULT_SF_URL  = st.secrets.get("SF_INSTANCE_URL",  "")

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


def sf_verify(instance_url: str, session_id: str):
    """Ping the userinfo endpoint. Returns (ok: bool, info: dict)."""
    try:
        r = requests.get(
            f"{instance_url}/services/oauth2/userinfo",
            headers=_hdr(session_id), timeout=10,
        )
        return (True, r.json()) if r.status_code == 200 else (False, {})
    except Exception:
        return False, {}


def search_reports(instance_url: str, session_id: str, term: str) -> list:
    term_safe = term.replace("'", "\\'")
    q = f"SELECT Id, Name FROM Report WHERE Name LIKE '%{term_safe}%' LIMIT 30"
    r = requests.get(
        f"{instance_url}/services/data/{SF_API_VERSION}/query",
        headers=_hdr(session_id), params={"q": q}, timeout=15,
    )
    r.raise_for_status()
    return r.json().get("records", [])


def _parse_report_response(data: dict):
    """Parse Analytics API report response → (DataFrame, all_data: bool)."""
    meta        = data.get("reportMetadata", {})
    ext         = data.get("reportExtendedMetadata", {})
    detail_cols = meta.get("detailColumns", [])
    col_info    = ext.get("detailColumnInfo", {})
    col_labels  = [col_info.get(c, {}).get("label", c) for c in detail_cols]
    numeric_t   = {"currency", "double", "int", "percent", "number"}
    col_types   = {c: col_info.get(c, {}).get("dataType", "string") for c in detail_cols}

    rows = data.get("factMap", {}).get("T!T", {}).get("rows", [])
    parsed = []
    for row in rows:
        pr = []
        for i, cell in enumerate(row.get("dataCells", [])):
            key   = detail_cols[i] if i < len(detail_cols) else ""
            dtype = col_types.get(key, "string")
            pr.append(cell.get("value") if dtype in numeric_t else cell.get("label"))
        parsed.append(pr)

    return pd.DataFrame(parsed, columns=col_labels), data.get("allData", True)


@st.cache_data(ttl=300, show_spinner=False)
def run_sf_report(instance_url: str, session_id: str, report_id: str,
                  date_col_label: str = None):
    """
    Execute a Salesforce report asynchronously.
    If date_col_label is provided, a 12-month lookback filter is injected
    at runtime so the row count stays under the 2,000-row API limit.
    Results are cached for 5 minutes.
    """
    from datetime import timedelta

    h   = _hdr(session_id)
    api = f"{instance_url}/services/data/{SF_API_VERSION}/analytics/reports/{report_id}"

    # ── Discover date column API name if a label is given ─────────────────
    date_filter_col = None
    if date_col_label:
        meta_r = requests.get(api, headers=h, timeout=15)
        if meta_r.ok:
            col_info = (meta_r.json()
                        .get("reportExtendedMetadata", {})
                        .get("detailColumnInfo", {}))
            date_filter_col = next(
                (k for k, v in col_info.items()
                 if v.get("label", "").lower() == date_col_label.lower()),
                None,
            )

    # ── Build POST body — inject 12-month filter when possible ────────────
    cutoff = (date.today() - timedelta(days=365)).strftime("%Y-%m-%d")
    body   = {}
    if date_filter_col:
        body = {
            "reportMetadata": {
                "reportFilters": [
                    {
                        "column":        date_filter_col,
                        "filterType":    "fieldValue",
                        "isRunPageOnly": False,
                        "operator":      "greaterOrEqual",
                        "value":         cutoff,
                    }
                ]
            }
        }

    # ── Kick off async instance ────────────────────────────────────────────
    r = requests.post(f"{api}/instances", headers=h, json=body, timeout=30)
    r.raise_for_status()
    inst_id = r.json()["id"]

    # ── Poll up to 3 minutes ───────────────────────────────────────────────
    for _ in range(90):
        time.sleep(2)
        r = requests.get(f"{api}/instances/{inst_id}", headers=h, timeout=30)
        r.raise_for_status()
        payload = r.json()
        status  = payload.get("attributes", {}).get("status", "")
        if status == "Success":
            return _parse_report_response(payload)
        if status in ("Error", "Cancelled"):
            code = payload.get("attributes", {}).get("errorCode", "Unknown")
            raise RuntimeError(f"Report failed: {code}")
    raise TimeoutError("Report timed out after 3 minutes.")


# ─────────────────────────────────────────────────────────────────────────────
# Data prep (DataFrame-based, replacing the file-path loaders)
# ─────────────────────────────────────────────────────────────────────────────

def _coerce_numeric(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(
        s.astype(str).str.replace(",", "").str.replace("$", "").str.strip(),
        errors="coerce",
    )


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
    lookup = {}
    if cses_df is not None:
        for _, row in cses_df.iterrows():
            lookup[row["owner_name"]] = {
                "Team":    row.get("team", ""),
                "Manager": row.get("manager_name", "Unassigned"),
            }

    rows = []
    for r in active:
        info = lookup.get(r["account_owner"], {"Team": "", "Manager": "Unassigned"})
        base = {
            "Team":         info["Team"],
            "Manager":      info["Manager"],
            "Rep":          r["account_owner"],
            "Account Name": r["account_name"],
            "Account ID":   r["account_id"],
        }
        for n, d in enumerate(r["deficits"], 1):
            base[f"Deficit {n} Amount"]         = -round(d["amount"], 2)
            base[f"Deficit {n} Effective Date"] = d["eff_date"].date()
            base[f"Deficit {n} Clears/Expires"] = d["clear_date"].date()
            base[f"Deficit {n} Source"]         = d["source"]
        rows.append(base)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    sort_cols = (["Team", "Manager", "Rep", "Account Name"] if cses_df is not None
                 else ["Rep", "Account Name"])
    return df.sort_values(sort_cols).reset_index(drop=True)


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

    instance_url = st.text_input(
        "Instance URL",
        value=_DEFAULT_SF_URL,
        placeholder="https://myorg.my.salesforce.com",
        key="sf_url",
    )
    session_id = st.text_input(
        "Session ID",
        type="password",
        placeholder="00D…",
        key="sf_sid",
    )

    if st.button("Connect", use_container_width=True):
        if not instance_url or not session_id:
            st.error("Both fields are required.")
        else:
            with st.spinner("Verifying session…"):
                ok, info = sf_verify(instance_url.rstrip("/"), session_id)
            if ok:
                st.session_state.update({
                    "connected":   True,
                    "sf_instance": instance_url.rstrip("/"),
                    "sf_token":    session_id,
                })
                name = info.get("name") or info.get("preferred_username", "")
                st.success(f"Connected{' — ' + name if name else ''}")
            else:
                st.session_state["connected"] = False
                st.error("Connection failed. Check URL and Session ID.")

    # ── CSEs + Run ───────────────────────────────────────────────────────────
    if st.session_state.get("connected"):
        st.markdown("---")
        st.markdown("### Optional: CSEs Report")
        st.caption("Enables Team and Manager filters. Search by report name to select.")

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
    st.info("Enter your Salesforce Instance URL and Session ID in the sidebar to get started.")
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
        """)
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Run reports & compute watermarks
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.get("connected") and "run_btn" in dir() and run_btn:
    bar   = st.progress(0, "Fetching Sales report…")
    warns = []
    try:
        with st.spinner("Running Sales report…"):
            s_df, s_all = run_sf_report(
                st.session_state["sf_instance"], st.session_state["sf_token"],
                SALES_REPORT_ID, date_col_label="Order Effective Date",
            )
        if not s_all:
            warns.append("Sales report returned 2,000+ rows and may be truncated. "
                         "Consider adding a date filter to the report in Salesforce.")
        bar.progress(30, "Fetching Terminations report…")

        with st.spinner("Running Terminations report…"):
            t_df, t_all = run_sf_report(
                st.session_state["sf_instance"], st.session_state["sf_token"],
                TERMS_REPORT_ID, date_col_label="Billing End Date",
            )
        if not t_all:
            warns.append("Terminations report returned 2,000+ rows and may be truncated.")
        bar.progress(60, "Running CSEs report…" if cses_id else "Processing data…")

        cses_raw = None
        if cses_id:
            with st.spinner("Running CSEs report…"):
                cses_raw, _ = run_sf_report(
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

master_df = st.session_state["master_df"]
results   = st.session_state["results"]
cses_df   = st.session_state["cses_df"]
run_date  = st.session_state["run_date"]
has_cses  = cses_df is not None and not master_df.empty and "Team" in master_df.columns


# ── Filters ──────────────────────────────────────────────────────────────────
st.markdown('<div class="filter-wrap">', unsafe_allow_html=True)

if has_cses:
    fc = st.columns([1, 1, 1, 1.6, 0.5])

    # Team
    all_teams   = sorted(master_df["Team"].dropna().unique())
    sel_teams   = fc[0].multiselect("Team", all_teams, placeholder="All teams", key="f_team")

    # Manager — cascades from team
    mgr_pool    = (master_df[master_df["Team"].isin(sel_teams)]["Manager"]
                   if sel_teams else master_df["Manager"])
    all_managers = sorted(mgr_pool.dropna().unique())
    sel_managers = fc[1].multiselect("Manager", all_managers, placeholder="All managers", key="f_mgr")

    # Rep — cascades from manager (or team if no manager selected)
    if sel_managers:
        rep_pool = master_df[master_df["Manager"].isin(sel_managers)]["Rep"]
    elif sel_teams:
        rep_pool = master_df[master_df["Team"].isin(sel_teams)]["Rep"]
    else:
        rep_pool = master_df["Rep"]
    all_reps  = sorted(rep_pool.dropna().unique())
    sel_reps  = fc[2].multiselect("Rep", all_reps, placeholder="All reps", key="f_rep")

    search    = fc[3].text_input("Search account name", placeholder="Type to search…", key="f_search")
    if fc[4].button("Clear", use_container_width=True, key="f_clear"):
        for k in ["f_team", "f_mgr", "f_rep", "f_search"]:
            st.session_state.pop(k, None)
        st.rerun()

else:
    fc       = st.columns([1, 2, 0.5])
    all_reps = sorted(master_df["Rep"].dropna().unique()) if not master_df.empty else []
    sel_teams, sel_managers = [], []
    sel_reps = fc[0].multiselect("Rep", all_reps, placeholder="All reps", key="f_rep")
    search   = fc[1].text_input("Search account name", placeholder="Type to search…", key="f_search")
    if fc[2].button("Clear", use_container_width=True, key="f_clear"):
        for k in ["f_rep", "f_search"]:
            st.session_state.pop(k, None)
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Apply filters
filtered = master_df.copy()
if sel_teams:
    filtered = filtered[filtered["Team"].isin(sel_teams)]
if sel_managers:
    filtered = filtered[filtered["Manager"].isin(sel_managers)]
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
    # Decide which columns to show — hide Team/Manager if no CSEs data
    display_cols = [c for c in filtered.columns if has_cses or c not in ("Team", "Manager")]

    # Build column config
    col_cfg = {}
    if has_cses:
        col_cfg["Team"]    = st.column_config.TextColumn("Team",    width="small")
        col_cfg["Manager"] = st.column_config.TextColumn("Manager", width="medium")
    col_cfg["Rep"]          = st.column_config.TextColumn("Rep",          width="medium")
    col_cfg["Account Name"] = st.column_config.TextColumn("Account Name", width="large")
    col_cfg["Account ID"]   = st.column_config.TextColumn("Account ID",   width="medium")

    for c in filtered.columns:
        if "Amount" in c:
            col_cfg[c] = st.column_config.NumberColumn(c, format="$%.2f", width="small")
        elif "Effective Date" in c or "Clears" in c:
            col_cfg[c] = st.column_config.DateColumn(c, format="MM/DD/YYYY", width="small")
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
        display_cols_for_csv = [c for c in filtered.columns
                                if has_cses or c not in ("Team", "Manager")]
        csv_buf = io.StringIO()
        filtered[display_cols_for_csv].to_csv(csv_buf, index=False)
        st.download_button(
            label="Download Filtered CSV",
            data=csv_buf.getvalue(),
            file_name=f"Deficits_filtered_{run_date.strftime('%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

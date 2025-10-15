# bears_dashboard.py
import streamlit as st
import pandas as pd
import io
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(page_title="🐻 Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.caption("Uploads → Tables → Charts: Bears vs NFL Averages (Weekly & YTD)")

# -----------------------
# Helpers
# -----------------------
COMMON_WEEK_COLS = ["Week", "week", "WEEK"]

OFF_CANDIDATES = [
    "YPA", "YPC", "CMP%", "3D%", "RZ%", "EPA/Play", "SR%", "PTS/G", "Yds/G", "TO/G"
]
DEF_CANDIDATES = [
    "SACK", "INT", "FF", "FR", "QB Hits", "Pressures", "RZ% Allowed", "3D% Allowed",
    "EPA/Play Allowed", "SR% Allowed", "Pts Allowed/G", "Yds Allowed/G", "DVOA"
]

def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure DataFrame has unique column names to avoid Streamlit error."""
    if df is None or df.empty:
        return df
    cols = pd.Series(df.columns, dtype=str)
    for i in range(len(cols)):
        dup_count = (cols[:i] == cols[i]).sum()
        if dup_count:
            cols[i] = f"{cols[i]}_{dup_count+1}"
    df.columns = cols
    return df

def find_week_col(df: pd.DataFrame) -> str | None:
    for c in COMMON_WEEK_COLS:
        if c in df.columns:
            return c
    return None

def coerce_numeric(df: pd.DataFrame, skip_cols=("Team","team","Opponent","Opp","opponent")) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    for c in df.columns:
        if c in skip_cols:
            continue
        # Try numeric if it looks numeric-ish
        df[c] = pd.to_numeric(df[c], errors="ignore")
    return df

def ensure_week_numeric(df: pd.DataFrame) -> pd.DataFrame:
    wk = find_week_col(df)
    if wk:
        df[wk] = pd.to_numeric(df[wk], errors="coerce")
        df = df.dropna(subset=[wk])
        df[wk] = df[wk].astype(int)
    return df

def rename_week_to_standard(df: pd.DataFrame) -> pd.DataFrame:
    wk = find_week_col(df)
    if wk and wk != "Week":
        return df.rename(columns={wk: "Week"})
    return df

def ytd_pairwise_mean(bears_df: pd.DataFrame, nfl_df: pd.DataFrame) -> pd.DataFrame:
    """Compute YTD means of overlapping numeric columns."""
    if bears_df is None or bears_df.empty or nfl_df is None or nfl_df.empty:
        return pd.DataFrame()
    b = bears_df.copy()
    n = nfl_df.copy()
    # numeric only intersect
    b_num = b.select_dtypes(include=[np.number])
    n_num = n.select_dtypes(include=[np.number])
    common = [c for c in b_num.columns if c in n_num.columns]
    if not common:
        return pd.DataFrame()
    return pd.DataFrame({
        "Bears YTD": b_num[common].mean(),
        "NFL Avg YTD": n_num[common].mean()
    })

def auto_compute_nfl_averages(league_df: pd.DataFrame, metric_group: str) -> pd.DataFrame:
    """
    From a league-wide weekly dataframe (all teams, multiple weeks), compute per-week NFL averages.
    Expects a 'Week' column and team rows.
    metric_group: 'off' or 'def' → filters columns by candidates lists.
    """
    if league_df is None or league_df.empty:
        return pd.DataFrame()

    df = league_df.copy()
    df = rename_week_to_standard(df)
    df = ensure_week_numeric(df)
    df = coerce_numeric(df)

    cands = OFF_CANDIDATES if metric_group == "off" else DEF_CANDIDATES
    cols = ["Week"] + [c for c in cands if c in df.columns]
    if len(cols) == 1:  # only Week is present
        return pd.DataFrame()
    df = df[cols]

    # group by Week, mean across all teams for that week
    nfl_avg = df.groupby("Week", as_index=False).mean(numeric_only=True)
    return nfl_avg

# -----------------------
# Sidebar: Uploads
# -----------------------
st.sidebar.header("⬆️ Upload Weekly Data")

off_file = st.sidebar.file_uploader("Bears Offense (weekly CSV/XLSX)", type=["csv","xlsx"], key="off_upl")
def_file = st.sidebar.file_uploader("Bears Defense (weekly CSV/XLSX)", type=["csv","xlsx"], key="def_upl")

st.sidebar.divider()
st.sidebar.subheader("NFL Averages (pick one approach)")
nfl_off_file = st.sidebar.file_uploader("Upload NFL Offense Averages (per week)", type=["csv","xlsx"], key="nf_

# bears_dashboard.py
# ---------------------------------------------
# Chicago Bears 2025–26 Weekly Tracker Dashboard
# Full single-file Streamlit app with:
# - Weekly Controls, Opponent Preview, Injuries
# - Uploaders (Offense/Defense/Personnel/SnapCounts)
# - Predictions + PDF report
# - NFL per-week averages (upload or auto-compute)
# - Weekly & YTD charts (Bears vs NFL Averages) AT THE BOTTOM
# ---------------------------------------------

import os
import io
import json
from datetime import datetime
from typing import Optional, List

import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
import matplotlib.pyplot as plt

# ------------- Page -------------
st.set_page_config(page_title="🐻 Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.caption("Weekly controls → uploads → opponent preview → injuries → predictions → exports → **Charts at the very bottom**")

# ------------- Constants -------------
EXCEL_FILE = "bears_weekly_analytics.xlsx"

SHEETS_REQUIRED = [
    "Offense", "Defense", "Personnel", "SnapCounts",
    "Injuries", "OpponentPreview", "MediaSummaries", "Predictions",
    "YTD_Team_Offense", "YTD_Team_Defense",
    "NFL_Offense_Avg", "NFL_Defense_Avg"
]

# Candidate columns we try to chart if present
OFF_CANDIDATES = [
    "YPA", "YPC", "CMP%", "3D%", "RZ%", "EPA/Play", "SR%", "PTS/G", "Yds/G", "TO/G"
]
DEF_CANDIDATES = [
    "SACK", "INT", "FF", "FR", "QB Hits", "Pressures", "RZ% Allowed", "3D% Allowed",
    "EPA/Play Allowed", "SR% Allowed", "Pts Allowed/G", "Yds Allowed/G", "DVOA"
]

COMMON_WEEK_COLS = ["Week", "week", "WEEK"]
TEAM_COLS_SKIP_NUMERIC = ("Team", "team", "Opponent", "Opp", "opponent")

# ------------- Utilities -------------
def ensure_workbook(path: str):
    """Create workbook with required sheets if missing."""
    if not os.path.exists(path):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            for s in SHEETS_REQUIRED:
                pd.DataFrame().to_excel(w, sheet_name=s, index=False)
    else:
        # ensure all sheets exist
        try:
            with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
                existing = set(w.book.sheetnames)  # type: ignore[attr-defined]
                for s in SHEETS_REQUIRED:
                    if s not in existing:
                        pd.DataFrame().to_excel(w, sheet_name=s, index=False)
        except Exception:
            # Some versions of openpyxl/engine disallow direct book access in append;
            # fall back to read all -> rewrite.
            book = {}
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                for s in SHEETS_REQUIRED:
                    book[s] = pd.DataFrame()
                for k, v in book.items():
                    v.to_excel(w, sheet_name=k, index=False)

def load_sheet(sheet: str) -> pd.DataFrame:
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()

def save_append(df_new: pd.DataFrame, sheet: str, keys: Optional[List[str]] = None):
    """Append rows to sheet and deduplicate by keys if provided."""
    ensure_workbook(EXCEL_FILE)
    df_old = load_sheet(sheet)
    if df_old is None or df_old.empty:
        out = df_new.copy()
    else:
        out = pd.concat([df_old, df_new], ignore_index=True)
    if keys:
        out = out.drop_duplicates(subset=keys, keep="last")
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        out.to_excel(w, sheet_name=sheet, index=False)

def export_entire_excel_download():
    try:
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("📥 Download All Data (Excel)", data=f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except FileNotFoundError:
        st.caption("Create some data first to enable the full Excel download.")

def read_any(uploaded) -> pd.DataFrame:
    if uploaded is None:
        return pd.DataFrame()
    name = uploaded.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(uploaded)
        return pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Could not read {uploaded.name}: {e}")
        return pd.DataFrame()

def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = pd.Series(df.columns, dtype=str)
    for i in range(len(cols)):
        dup_count = (cols[:i] == cols[i]).sum()
        if dup_count:
            cols[i] = f"{cols[i]}_{dup_count+1}"
    df.columns = cols
    return df

def find_week_col(df: pd.DataFrame) -> Optional[str]:
    for c in COMMON_WEEK_COLS:
        if c in df.columns:
            return c
    return None

def ensure_week_numeric(df: pd.DataFrame) -> pd.DataFrame:
    wk = find_week_col(df)
    if wk:
        df[wk] = pd.to_numeric(df[wk], errors="coerce")
        df = df.dropna(subset=[wk])
        df[wk] = df[wk].astype(int)
    return df

def rename_week_to_standard(df: pd.DataFrame) -> pd.DataFrame:
    wk = find_week_col(df)
    if wk and wk != "Week":
        return df.rename(columns={wk: "Week"})
    return df

def coerce_numeric(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    for c in df.columns:
        if c in TEAM_COLS_SKIP_NUMERIC:
            continue
        df[c] = pd.to_numeric(df[c], errors="ignore")
    return df

def auto_compute_nfl_averages(league_df: pd.DataFrame, metric_group: str) -> pd.DataFrame:
    """From league-wide weekly dataframe (all teams), compute per-week NFL averages."""
    if league_df is None or league_df.empty:
        return pd.DataFrame()
    df = league_df.copy()
    df = rename_week_to_standard(df)
    df = ensure_week_numeric(df)
    df = coerce_numeric(df)
    cands = OFF_CANDIDATES if metric_group == "off" else DEF_CANDIDATES
    cols = ["Week"] + [c for c in cands if c in df.columns]
    if len(cols) == 1:
        return pd.DataFrame()
    nfl_avg = df[cols].groupby("Week", as_index=False).mean(numeric_only=True)
    return nfl_avg

def ytd_pairwise_mean(bears_df: pd.DataFrame, nfl_df: pd.DataFrame) -> pd.DataFrame:
    if bears_df is None or bears_df.empty or nfl_df is None or nfl_df.empty:
        return pd.DataFrame()
    b_num = bears_df.select_dtypes(include=[np.number])
    n_num = nfl_df.select_dtypes(include=[np.number])
    common = [c for c in b_num.columns if c in n_num.col]()


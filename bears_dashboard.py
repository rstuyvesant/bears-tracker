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

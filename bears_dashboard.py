# bears_dashboard.py
# bears_dashboard.py
# ---------------------------------------------
# Chicago Bears 2025–26 Weekly Tracker Dashboard
# - Weekly Controls, Opponent Preview, Injuries
# - Uploaders (Offense/Defense/Personnel/SnapCounts)
# - Predictions + PDF report
# - NFL per-week averages (upload or auto-compute)
# - Column alias normalization (for easier matching)
# - Data debug + maintenance tools
# - Weekly & YTD charts (Bears vs NFL Averages) AT THE BOTTOM
# ---------------------------------------------

import os
from datetime import datetime
from typing import Optional, List

import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF

# >>> Headless plotting fix (must be before pyplot import) <<<
import matplotlib
matplotlib.use("Agg")
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

# --- Column name aliasing to improve chart matches ---
OFF_ALIASES = {
    "ThirdDown%": "3D%",
    "Third Down%": "3D%",
    "3rdDown%": "3D%",
    "RedZone%": "RZ%",
    "Red Zone%": "RZ%",
    "Points/G": "PTS/G",
    "Pts/G": "PTS/G",
    "Yards/G": "Yds/G",
    "EPA per play": "EPA/Play",
    "EPA/play": "EPA/Play",
    "Success%": "SR%",
    "Success Rate": "SR%",
    "Cmp%": "CMP%",
    "CmpPct": "CMP%",
}
DEF_ALIASES = {
    "3rdDown% Allowed": "3D% Allowed",
    "ThirdDown% Allowed": "3D% Allowed",
    "RedZone% Allowed": "RZ% Allowed",
    "Red Zone% Allowed": "RZ% Allowed",
    "Points Allowed/G": "Pts Allowed/G",
    "Yards Allowed/G": "Yds Allowed/G",
    "EPA/play Allowed": "EPA/Play Allowed",
    "EPA per play Allowed": "EPA/Play Allowed",
    "QBHits": "QB Hits",
    "Pressures Allowed": "Pressures",
}

def rename_aliases(df: pd.DataFrame, is_offense: bool) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    mapping = OFF_ALIASES if is_offense else DEF_ALIASES
    cols = {c: mapping.get(c, c) for c in df.columns}
    return df.rename(columns=cols)

# ------------- Utilities -------------
def ensure_workbook(path: str):
    """Create workbook with required sheets if missing."""
    if not os.path.exists(path):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            for s in SHEETS_REQUIRED:
                pd.DataFrame().to_excel(w, sheet_name=s, index=False)
    else:
        # ensure all sheets exist; rewrite minimal workbook if needed
        try:
            with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
                existing = set(w.book.sheetnames)  # type: ignore[attr-defined]
                need = [s for s in SHEETS_REQUIRED if s not in existing]
            if need:
                with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
                    for s in need:
                        pd.DataFrame().to_excel(w, sheet_name=s, index=False)
        except Exception:
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                for s in SHEETS_REQUIRED:
                    pd.DataFrame().to_excel(w, sheet_name=s, index=False)

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
    df = rename_aliases(df, is_offense=(metric_group == "off"))
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
    common = [c for c in b_num.columns if c in n_num.columns]
    if not common:
        return pd.DataFrame()
    return pd.DataFrame({
        "Bears YTD": b_num[common].mean(),
        "NFL Avg YTD": n_num[common].mean()
    })

# ------------- Sidebar -------------
st.sidebar.header("Sections")
st.sidebar.markdown("""
1) Weekly Controls  
2) Upload Game Data  
3) Opponent Preview  
4) Injuries  
5) Predictions  
6) Exports  
7) **Charts (bottom)**
""")

st.sidebar.divider()
st.sidebar.subheader("Excel")
export_entire_excel_download()

st.sidebar.divider()
# ----- Sidebar: NFL averages inputs -----
st.sidebar.subheader("NFL Averages Inputs")
nfl_off_file = st.sidebar.file_uploader("Upload NFL Offense Averages (per week)", type=["csv", "xlsx"], key="nfl_off")
nfl_def_file = st.sidebar.file_uploader("Upload NFL Defense Averages (per week)", type=["csv", "xlsx"], key="nfl_def")
st.sidebar.caption("— OR — upload league-wide weekly files (all teams, all weeks)")
league_off_all_file = st.sidebar.file_uploader("Upload League-wide Offense Weekly (all teams)", type=["csv", "xlsx"], key="lg_off_all")
league_def_all_file = st.sidebar.file_uploader("Upload League-wide Defense Weekly (all teams)", type=["csv", "xlsx"], key="lg_def_all")

# ------------- Ensure workbook exists -------------
ensure_workbook(EXCEL_FILE)

# ------------- 1) Weekly Controls -------------
st.header("1) Weekly Controls")
with st.form("weekly_controls_form", clear_on_submit=False):
    colA, colB, colC, colD = st.columns(4)
    week = colA.number_input("Week", min_value=1, step=1, format="%d")
    opponent = colB.text_input("Opponent (3-letter code, e.g., MIN, GB)")
    venue = colC.selectbox("Home/Away", ["Home", "Away"])
    notes = colD.text_input("Key Notes (one-liner)")

    submitted_wc = st.form_submit_button("Save Weekly Control Note")
    if submitted_wc:
        df = pd.DataFrame([{
            "Week": int(week),
            "Opponent": opponent.strip(),
            "Venue": venue,
            "Key_Notes": notes.strip(),
            "Timestamp": datetime.now().isoformat(timespec="seconds")
        }])
        save_append(df, "MediaSummaries", keys=["Week", "Opponent", "Key_Notes"])
        st.success("Weekly control note saved.")

# ------------- 2) Upload Game Data -------------
st.header("2) Upload Game Data (CSV/XLSX)")
c1, c2, c3, c4 = st.columns(4)
with c1:
    off_up = st.file_uploader("Bears Offense (weekly)", type=["csv", "xlsx"], key="off_upl")
    if off_up:
        df = read_any(off_up)
        df = rename_week_to_standard(df)
        df = ensure_week_numeric(df)
        df = coerce_numeric(df)
        df = rename_aliases(df, is_offense=True)
        save_append(df, "Offense", keys=["Week", "Opponent"] if "Opponent" in df.columns else ["Week"])
        st.success("Offense data appended.")
with c2:
    def_up = st.file_uploader("Bears Defense (weekly)", type=["csv", "xlsx"], key="def_upl")
    if def_up:
        df = read_any(def_up)
        df = rename_week_to_standard(df)
        df = ensure_week_numeric(df)
        df = coerce_numeric(df)
        df = rename_aliases(df, is_offense=False)
        save_append(df, "Defense", keys=["Week", "Opponent"] if "Opponent" in df.columns else ["Week"])
        st.success("Defense data appended.")
with c3:
    per_up = st.file_uploader("Personnel (weekly)", type=["csv", "xlsx"], key="per_upl")
    if per_up:
        df = read_any(per_up)
        df = rename_week_to_standard(df)
        df = ensure_week_numeric(df)
        save_append(df, "Personnel", keys=["Week"])
        st.success("Personnel data appended.")
with c4:
    snaps_up = st.file_uploader("Snap Counts (weekly)", type=["csv", "xlsx"], key="snaps_upl")
    if snaps_up:
        df = read_any(snaps_up)
        df = rename_week_to_standard(df)
        df = ensure_week_numeric(df)
        save_append(df, "SnapCounts", keys=["Week"])
        st.success("Snap counts appended.")

# Show last 10 rows quick view
st.subheader("Latest Uploads (last 10 rows)")
off = load_sheet("Offense")
defn = load_sheet("Defense")
personnel = load_sheet("Personnel")
snaps = load_sheet("SnapCounts")

cA, cB = st.columns(2)
with cA:
    st.markdown("**Bears Offense (weekly)**")
    st.dataframe(make_unique_columns(off.tail(10)), width="stretch") if not off.empty else st.caption("—")
with cB:
    st.markdown("**Bears Defense (weekly)**")
    st.dataframe(make_unique_columns(defn.tail(10)), width="stretch") if not defn.empty else st.caption("—")

cC, cD = st.columns(2)
with cC:
    st.markdown("**Personnel (weekly)**")
    st.dataframe(make_unique_columns(personnel.tail(10)), width="stretch") if not personnel.empty else st.caption("—")
with cD:
    st.markdown("**Snap Counts (weekly)**")
    st.dataframe(make_unique_columns(snaps.tail(10)), width="stretch") if not snaps.empty else st.caption("—")

# ------------- NFL averages load/compute -------------
nfl_off_avg = read_any(nfl_off_file)
nfl_def_avg = read_any(nfl_def_file)
league_off_all = read_any(league_off_all_file)
league_def_all = read_any(league_def_all_file)

for df_name, is_off in [("nfl_off_avg", True), ("nfl_def_avg", False), ("league_off_all", True), ("league_def_all", False)]:
    df = locals()[df_name]
    if not df.empty:
        df = rename_week_to_standard(df)
        df = ensure_week_numeric(df)
        df = coerce_numeric(df)
        df = rename_aliases(df, is_offense=is_off)
        locals()[df_name] = df

if nfl_off_avg.empty and not league_off_all.empty:
    nfl_off_avg = auto_compute_nfl_averages(league_off_all, "off")
if nfl_def_avg.empty and not league_def_all.empty:
    nfl_def_avg = auto_compute_nfl_averages(league_def_all, "def")

# Save the latest NFL averages back into workbook for YTD use
if not nfl_off_avg.empty:
    save_append(nfl_off_avg, "NFL_Offense_Avg", keys=["Week"])
if not nfl_def_avg.empty:
    save_append(nfl_def_avg, "NFL_Defense_Avg", keys=["Week"])

# Quick peek last 10 NFL averages
st.subheader("NFL Averages (last 10 rows)")
cE, cF = st.columns(2)
with cE:
    st.markdown("**NFL Offense Averages (per week)**")
    st.dataframe(make_unique_columns(nfl_off_avg.tail(10)), width="stretch") if not nfl_off_avg.empty else st.caption("—")
with cF:
    st.markdown("**NFL Defense Averages (per week)**")
    st.dataframe(make_unique_columns(nfl_def_avg.tail(10)), width="stretch") if not nfl_def_avg.empty else st.caption("—")

# ----- Data Maintenance: delete rows by week (to remove old GB Week 1, etc.) -----
with st.expander("🧹 Data Maintenance (delete rows by week)"):
    del_week = st.number_input("Week to delete across all sheets", min_value=1, step=1, format="%d")
    if st.button("Delete rows for this week (ALL sheets)"):
        for sheet in ["Offense","Defense","Personnel","SnapCounts","Injuries","OpponentPreview","MediaSummaries","Predictions",
                      "NFL_Offense_Avg","NFL_Defense_Avg","YTD_Team_Offense","YTD_Team_Defense"]:
            df = load_sheet(sheet)
            if not df.empty and "Week" in df.columns:
                df = df[df["Week"] != int(del_week)]
                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                    df.to_excel(w, sheet_name=sheet, index=False)
        st.success(f"Deleted Week {int(del_week)} from all sheets that had it.")

st.divider()

# ------------- 3) Opponent Preview -------------
st.header("3) Opponent Preview")
with st.form("opponent_preview_form", clear_on_submit=False):
    c1, c2, c3 = st.columns(3)
    wk = c1.number_input("Week", min_value=1, step=1, format="%d")
    opp = c2.text_input("Opponent (3-letter code)")
    preview = c3.text_input("Preview Headline")
    details = st.text_area("Notes (matchups, tendencies, weather, etc.)", height=120)
    sub = st.form_submit_button("Save Opponent Preview")
    if sub:
        df = pd.DataFrame([{
            "Week": int(wk), "Opponent": opp.strip(),
            "Preview_Headline": preview.strip(),
            "Notes": details.strip(),
            "Timestamp": datetime.now().isoformat(timespec="seconds")
        }])
        save_append(df, "OpponentPreview", keys=["Week", "Opponent"])
        st.success("Opponent preview saved.")
opprev = load_sheet("OpponentPreview")
st.dataframe(make_unique_columns(opprev.tail(10)), width="stretch") if not opprev.empty else st.caption("—")

# ------------- 4) Injuries -------------
st.header("4) Injuries")
with st.form("injuries_form", clear_on_submit=True):
    c1, c2, c3, c4 = st.columns(4)
    wk_i = c1.number_input("Week", min_value=1, step=1, format="%d")
    opp_i = c2.text_input("Opponent (3-letter code)")
    player = c3.text_input("Player")
    status = c4.selectbox("Status", ["Questionable", "Doubtful", "Out", "IR", "Active"])
    c5, c6, c7 = st.columns(3)
    bodypart = c5.text_input("Body Part")
    practice = c6.selectbox("Practice", ["DNP", "Limited", "Full", "N/A"])
    game_status = c7.selectbox("Game Status", ["TBD", "Active", "Inactive", "N/A"])
    notes_i = st.text_area("Notes", height=100)
    save_i = st.form_submit_button("Save Injury")
    if save_i:
        df = pd.DataFrame([{
            "Week": int(wk_i), "Opponent": opp_i.strip(),
            "Player": player.strip(), "Status": status,
            "BodyPart": bodypart.strip(), "Practice": practice,
            "GameStatus": game_status, "Notes": notes_i.strip(),
            "Timestamp": datetime.now().isoformat(timespec="seconds")
        }])
        save_append(df, "Injuries", keys=["Week", "Opponent", "Player"])
        st.success("Injury saved.")
inj = load_sheet("Injuries")
st.dataframe(make_unique_columns(inj.tail(10)), width="stretch") if not inj.empty else st.caption("—")

# ------------- 5) Predictions -------------
st.header("5) Weekly Game Predictions")
with st.form("pred_form", clear_on_submit=True):
    c1, c2, c3, c4 = st.columns(4)
    wk_p = c1.number_input("Week", min_value=1, step=1, format="%d")
    opp_p = c2.text_input("Opponent (3-letter code)")
    winner = c3.selectbox("Predicted Winner", ["CHI", "Opponent"])
    conf = c4.slider("Confidence (0–100%)", 0, 100, 60)
    rationale = st.text_area("Rationale (use strategy, injuries, matchups, YTD)", height=100)
    save_p = st.form_submit_button("Save Prediction")
    if save_p:
        df = pd.DataFrame([{
            "Week": int(wk_p), "Opponent": opp_p.strip(),
            "Predicted_Winner": winner, "Confidence": conf,
            "Rationale": rationale.strip(),
            "Timestamp": datetime.now().isoformat(timespec="seconds")
        }])
        save_append(df, "Predictions", keys=["Week", "Opponent"])
        st.success("Prediction saved.")
preds = load_sheet("Predictions")
st.dataframe(make_unique_columns(preds.tail(10)), width="stretch") if not preds.empty else st.caption("—")

# Quick PDF generation for the current week’s report
def build_week_pdf(week_num: int, out_path: str):
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, f"Chicago Bears Week {week_num} Report", ln=1, align="C")
    pdf.set_font("Arial", "", 11)

    def add_table(title: str, df: pd.DataFrame, max_rows=12):
        pdf.ln(4)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, title, ln=1)
        pdf.set_font("Arial", "", 9)
        if df is None or df.empty:
            pdf.cell(0, 6, "—", ln=1)
            return
        cols = list(df.columns)
        head = " | ".join([str(c)[:15] for c in cols[:6]])
        pdf.cell(0, 5, head, ln=1)
        for _, row in df.head(max_rows).iterrows():
            line = " | ".join([str(row.get(c, ""))[:15] for c in cols[:6]])
            pdf.cell(0, 5, line, ln=1)

    off_w = off[off["Week"] == week_num] if "Week" in off.columns else pd.DataFrame()
    def_w = defn[defn["Week"] == week_num] if "Week" in defn.columns else pd.DataFrame()
    opp_w = opprev[opprev["Week"] == week_num] if "Week" in opprev.columns else pd.DataFrame()
    inj_w = inj[inj["Week"] == week_num] if "Week" in inj.columns else pd.DataFrame()
    pr_w = preds[preds["Week"] == week_num] if "Week" in preds.columns else pd.DataFrame()

    add_table("Opponent Preview", opp_w)
    add_table("Injuries", inj_w)
    add_table("Offense (week)", off_w)
    add_table("Defense (week)", def_w)
    add_table("Prediction", pr_w)

    pdf.output(out_path)

st.subheader("Exports")
colx, coly, colz = st.columns(3)
with colx:
    wk_exp_pre = st.number_input("Week to export (Pre)", min_value=1, step=1, format="%d", key="wkpre")
    if st.button("Export Pre PDF"):
        outfile = f"W{wk_exp_pre:02d}_Pre.pdf"
        build_week_pdf(int(wk_exp_pre), outfile)
        with open(outfile, "rb") as f:
            st.download_button("Download Pre PDF", data=f, file_name=outfile, mime="application/pdf")
        st.success("Pre PDF built.")
with coly:
    wk_exp_post = st.number_input("Week to export (Post)", min_value=1, step=1, format="%d", key="wkpost")
    if st.button("Export Post PDF"):
        outfile = f"W{wk_exp_post:02d}_Post.pdf"
        build_week_pdf(int(wk_exp_post), outfile)
        with open(outfile, "rb") as f:
            st.download_button("Download Post PDF", data=f, file_name=outfile, mime="application/pdf")
        st.success("Post PDF built.")
with colz:
    wk_exp_final = st.number_input("Week to export (Final)", min_value=1, step=1, format="%d", key="wkfinal")
    if st.button("Export Final PDF"):
        outfile = f"W{wk_exp_final:02d}_Final.pdf"
        build_week_pdf(int(wk_exp_final), outfile)
        with open(outfile, "rb") as f:
            st.download_button("Download Final PDF", data=f, file_name=outfile, mime="application/pdf")
        st.success("Final PDF built.")

st.divider()

# ============================
# 6) --- CHARTS AT THE BOTTOM
# ============================

# Quick debug panel so you can see why charts may not render
with st.expander("🩺 Data Sanity / Chart Debug"):
    def cols(df):
        return list(df.columns) if not df.empty else []
    # Use the freshest frames available for debug
    nfl_off_dbg = nfl_off_avg if not nfl_off_avg.empty else load_sheet("NFL_Offense_Avg")
    nfl_def_dbg = nfl_def_avg if not nfl_def_avg.empty else load_sheet("NFL_Defense_Avg")
    st.write("Off (rows, cols):", off.shape, "Columns:", cols(off))
    st.write("Def (rows, cols):", defn.shape, "Columns:", cols(defn))
    st.write("NFL Off Avg (rows, cols):", (nfl_off_dbg.shape if isinstance(nfl_off_dbg, pd.DataFrame) else (0,0)), "Columns:", cols(nfl_off_dbg))
    st.write("NFL Def Avg (rows, cols):", (nfl_def_dbg.shape if isinstance(nfl_def_dbg, pd.DataFrame) else (0,0)), "Columns:", cols(nfl_def_dbg))
    off_overlap = [c for c in OFF_CANDIDATES if c in off.columns and isinstance(nfl_off_dbg, pd.DataFrame) and c in nfl_off_dbg.columns]
    def_overlap = [c for c in DEF_CANDIDATES if c in defn.columns and isinstance(nfl_def_dbg, pd.DataFrame) and c in nfl_def_dbg.columns]
    st.write("Offense overlapping metrics:", off_overlap)
    st.write("Defense overlapping metrics:", def_overlap)
    st.write("Has Week in Off/NFL-Off:", ("Week" in off.columns), (isinstance(nfl_off_dbg, pd.DataFrame) and "Week" in nfl_off_dbg.columns))
    st.write("Has Week in Def/NFL-Def:", ("Week" in defn.columns), (isinstance(nfl_def_dbg, pd.DataFrame) and "Week" in nfl_def_dbg.columns))

st.header("6) Charts — Bears vs NFL Averages (Weekly & YTD)")
st.caption("These charts appear **after** Weekly Controls, Uploads, Opponent Preview, Injuries, Predictions, and Exports.")

# Re-load latest (in case just appended)
off = load_sheet("Offense")
defn = load_sheet("Defense")
nfl_off_avg_sheet = load_sheet("NFL_Offense_Avg")
nfl_def_avg_sheet = load_sheet("NFL_Defense_Avg")

# prefer current session uploads when present
nfl_off_avg_chart = nfl_off_avg.copy() if not nfl_off_avg.empty else nfl_off_avg_sheet.copy()
nfl_def_avg_chart = nfl_def_avg.copy() if not nfl_def_avg.empty else nfl_def_avg_sheet.copy()

# Normalize numeric + Week + aliases again for chart copies
for name, is_off in [("off", True), ("defn", False), ("nfl_off_avg_chart", True), ("nfl_def_avg_chart", False)]:
    df = locals()[name]
    if not df.empty:
        df = rename_week_to_standard(df)
        df = ensure_week_numeric(df)
        df = coerce_numeric(df)
        df = rename_aliases(df, is_offense=is_off)
        locals()[name] = df

def plot_weekly(bears_df: pd.DataFrame, nfl_df: pd.DataFrame, title: str, candidates: list[str]):
    if bears_df.empty or nfl_df.empty:
        st.info(f"{title}: upload Bears weekly and NFL averages (or league-wide) to see this chart.")
        return
    present = [c for c in candidates if c in bears_df.columns and c in nfl_df.columns]
    if not present:
        st.warning(f"{title}: no overlapping metric columns found.")
        return
    if "Week" not in bears_df.columns or "Week" not in nfl_df.columns:
        st.warning(f"{title}: missing 'Week' column.")
        return

    fig, ax = plt.subplots()
    for col in present:
        try:
            ax.plot(bears_df["Week"], bears_df[col], marker="o", label=f"Bears {col}")
        except Exception:
            pass
        try:
            ax.plot(nfl_df["Week"], nfl_df[col], linestyle="--", label=f"NFL Avg {col}")
        except Exception:
            pass
    ax.set_title(title)
    ax.set_xlabel("Week")
    ax.set_ylabel("Value")
    ax.legend()
    st.pyplot(fig)

def plot_ytd(bears_df: pd.DataFrame, nfl_df: pd.DataFrame, title: str):
    if bears_df.empty or nfl_df.empty:
        st.info(f"{title}: upload Bears weekly and NFL averages (or league-wide) to see this chart.")
        return
    ytd_df = ytd_pairwise_mean(bears_df, nfl_df)
    if ytd_df.empty:
        st.warning(f"{title}: no overlapping numeric columns to aggregate.")
        return
    fig, ax = plt.subplots()
    ytd_df.sort_index().plot(kind="bar", ax=ax)
    ax.set_title(title)
    ax.set_ylabel("Average Value")
    st.pyplot(fig)

# --- Weekly Offense vs NFL ---
plot_weekly(off, nfl_off_avg_chart, "Bears Offense vs NFL Average by Week", OFF_CANDIDATES)
# --- YTD Offense vs NFL ---
plot_ytd(off, nfl_off_avg_chart, "YTD Offense: Bears vs NFL Average")

st.divider()

# --- Weekly Defense vs NFL ---
plot_weekly(defn, nfl_def_avg_chart, "Bears Defense vs NFL Average by Week", DEF_CANDIDATES)
# --- YTD Defense vs NFL ---
plot_ytd(defn, nfl_def_avg_chart, "YTD Defense: Bears vs NFL Average")

# Tips
with st.expander("ℹ️ Charting Tips & Column Names"):
    st.markdown("""
- Be sure your weekly files include a **Week** column. We accept `week`/`WEEK` automatically.
- Offense metrics we look for:  
  `YPA, YPC, CMP%, 3D%, RZ%, EPA/Play, SR%, PTS/G, Yds/G, TO/G`.
- Defense metrics we look for:  
  `SACK, INT, FF, FR, QB Hits, Pressures, RZ% Allowed, 3D% Allowed, EPA/Play Allowed, SR% Allowed, Pts Allowed/G, Yds Allowed/G, DVOA`.
- If a metric doesn’t appear, make sure it exists **in both** Bears and NFL frames with the **same column name**.
""")

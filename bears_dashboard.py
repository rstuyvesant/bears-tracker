# Bears Dashboard Streamlit App — Full Revised Version
# Features:
# - Sidebar controls + color-coding toggle
# - Guided main page (1–6) with expanders
# - Quick Manual Entry panel (works even when fetch is down)
# - Upload Weekly Files (Offense/Defense/Personnel/Snap Counts)
# - Upload NFL Averages (2A) to enable color comparisons
# - Live Data Preview with optional green/red highlighting vs NFL averages
# - Export Week Excel and Final PDF
#
# Notes:
# - NFL fetch uses best-effort fallbacks. If upstream 404s, use Quick Manual Entry + Upload NFL Averages.
# - Exports saved to ./exports/

import os
from typing import List, Dict

import pandas as pd
import numpy as np
import streamlit as st
import openpyxl
from fpdf import FPDF

try:
    import nfl_data_py as nfl
except Exception:
    nfl = None

st.set_page_config(page_title="Bears Weekly Tracker", layout="wide")

CURRENT_SEASON = 2025
DATA_FILE = "bears_weekly_analytics.xlsx"
EXPORTS_DIR = "exports"
os.makedirs(EXPORTS_DIR, exist_ok=True)

OPPONENTS = [
    "MIN","GB","DET","KC","BUF","MIA","SF","DAL","PHI","NYG","NYJ","SEA","LAR","ARI","ATL",
    "NO","TB","CAR","HOU","TEN","JAX","IND","CLE","CIN","BAL","PIT","DEN","LV","LAC","NE","WAS"
]

OFF_EXPANDED_COLS = [
    "Week","Opponent","Team",
    "PassYds","RushYds","RecvYds","PassTD","RushTD","Points","Turnovers","SacksAgainst",
    "YPA","CMP%","YAC","Targets","Recs","1D","Total_YDS","QBR","SR%","EPA","DVOA",
    "Red_Zone%","3rd_Down_Conv%","Expl_Plays (20+)",
    "NFL_Avg._PassYds","NFL_Avg._RushYds","NFL_Avg._RecvYds","NFL_Avg._PassTD","NFL_Avg._RushTD","NFL_Avg._Points",
    "NFL_Avg._Turnovers","NFL_Avg._SacksAgainst","NFL_Avg._YPA","NFL_Avg._CMP%","NFL_Avg._YAC",
    "NFL_Avg._Targets","NFL_Avg._Recs","NFL_Avg._1D","NFL_Avg._Total_YDS","NFL_Avg._QBR","NFL_Avg._SR%",
    "NFL_Avg._EPA","NFL_Avg._DVOA","NFL_Avg._Red_Zone%","NFL_Avg._3rd_Down_Conv%","NFL_Avg._Expl_Plays (20+)"
]

DEF_EXPANDED_COLS = [
    "Week","Opponent","Team",
    "Sacks","QB_Hits","Pressures","INT","Turnovers","FF","FR",
    "3rd_Down%_Allowed","RZ%_Allowed","YdsAllowed_Pass","YdsAllowed_Rush","PointsAllowed",
    "NFL_Avg._Sacks","NFL_Avg._QB_Hits","NFL_Avg._Pressures","NFL_Avg._INT","NFL_Avg._Turnovers",
    "NFL_Avg._FF","NFL_Avg._FR","NFL_Avg._3rd_Down%_Allowed","NFL_Avg._RZ%_Allowed",
    "NFL_Avg._YdsAllowed_Pass","NFL_Avg._YdsAllowed_Rush","NFL_Avg._PointsAllowed"
]

PERSONNEL_COLS = [
    "Week","Opponent","Team",
    "11_Personnel","12_Personnel","13_Personnel","21_Personnel","Other_Personnel","Total_Snaps",
    "NFL_Avg._11_Personnel","NFL_Avg._12_Personnel","NFL_Avg._13_Personnel","NFL_Avg._21_Personnel",
    "NFL_Avg._Other_Personnel","NFL_Avg._Total_Snaps"
]

SNAPCOUNTS_COLS = [
    "Week","Opponent","Team",
    "Player1","Player2","Player3","Player4","Player5","Player6","Player7","Player8","Player9","Player10","Player11",
    "NFL_Avg._RBs","NFL_Avg._WRs","NFL_Avg._TEs","NFL_Avg._QB","NFL_Avg._OL","NFL_Avg._DL","NFL_Avg._LB","NFL_Avg._DB"
]

SHEET_ORDER = [
    ("Offense", OFF_EXPANDED_COLS),
    ("Defense", DEF_EXPANDED_COLS),
    ("Personnel", PERSONNEL_COLS),
    ("Snap_Counts", SNAPCOUNTS_COLS),
    ("Weekly_Notes", ["Week","Opponent","Notes"]),
    ("Media", ["Week","Opponent","Source","Summary"]),
    ("Injuries", ["Week","Opponent","Team","Player","Status","BodyPart","Practice","GameStatus","Notes"]),
    ("Opponent_Preview", ["Week","Opponent","Key_Threats","Tendencies","Injuries","Notes"]),
    ("Predictions", ["Week","Opponent","Predicted_Winner","Confidence","Rationale"]),
]

def normalize_week(val) -> str:
    s = str(val).strip().upper()
    if s.startswith("W"):
        return s
    try:
        i = int(float(s))
        return f"W{i:02d}"
    except Exception:
        return "W01"

def ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    # reorder
    front = [c for c in ["Week","Opponent","Team"] if c in cols]
    cols_order = front + [c for c in cols if c not in front]
    return df.reindex(columns=cols_order)

def clean_df(df: pd.DataFrame, opponent: str, team: str="CHI") -> pd.DataFrame:
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")].copy()
    if "Week" not in df.columns:
        df["Week"] = "W01"
    df["Week"] = df["Week"].apply(normalize_week)
    if "Opponent" not in df.columns:
        df["Opponent"] = opponent
    else:
        df["Opponent"] = df["Opponent"].fillna(opponent)
    if "Team" not in df.columns:
        df["Team"] = team
    else:
        df["Team"] = df["Team"].fillna(team)
    id_cols = {"Week","Opponent","Team"}
    metrics = [c for c in df.columns if c not in id_cols]
    if metrics:
        mask = df[metrics].apply(lambda r: r.isna().all() or (r.astype(str).str.strip()== "").all(), axis=1)
        df = df.loc[~mask].copy()
    return df

def init_workbook(path: str):
    if not os.path.exists(path):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sheet_name, cols in SHEET_ORDER:
            ws = wb.create_sheet(sheet_name)
            ws.append(cols)
        wb.save(path)

def upsert_sheet(path: str, sheet_name: str, df_new: pd.DataFrame, key_cols: List[str] = ["Week","Opponent","Team"]):
    init_workbook(path)
    wb = openpyxl.load_workbook(path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(df_new.columns.tolist())
        for _, r in df_new.iterrows():
            ws.append(r.tolist())
    else:
        ws = wb[sheet_name]
        data = list(ws.values)
        header = list(data[0]) if data else df_new.columns.tolist()
        rows = data[1:] if data else []
        existing = pd.DataFrame(rows, columns=header)
        all_cols = list(dict.fromkeys(list(existing.columns) + list(df_new.columns)))
        existing = existing.reindex(columns=all_cols)
        df_new = df_new.reindex(columns=all_cols)
        if all(k in existing.columns for k in key_cols):
            existing = existing.drop_duplicates(subset=key_cols, keep="last")
        if all(k in df_new.columns for k in key_cols):
            to_remove = set(tuple(x) for x in df_new[key_cols].to_numpy())
            existing = existing.loc[~existing[key_cols].apply(lambda r: tuple(r.values) in to_remove, axis=1)]
        combined = pd.concat([existing, df_new], ignore_index=True)
        ws.delete_rows(1, ws.max_row)
        ws.append(combined.columns.tolist())
        for _, r in combined.iterrows():
            ws.append(r.tolist())
    wb.save(path)

def compute_nfl_averages_for_week(week_int: int, season: int) -> Dict[str, float]:
    avg = {}
    if nfl is None:
        return avg
    try:
        weekly = nfl.import_weekly_data([season])
        if "week" in weekly and not weekly.empty:
            dfw = weekly[weekly["week"] == week_int]
            if dfw.empty:
                return avg
            if "passing_yards" in dfw:
                avg["NFL_Avg._PassYds"] = float(dfw["passing_yards"].mean())
            if "rushing_yards" in dfw:
                avg["NFL_Avg._RushYds"] = float(dfw["rushing_yards"].mean())
            if "passing_tds" in dfw:
                avg["NFL_Avg._PassTD"] = float(dfw["passing_tds"].mean())
            if "rushing_tds" in dfw:
                avg["NFL_Avg._RushTD"] = float(dfw["rushing_tds"].mean())
            # placeholders for other avg fields (left as NaN unless filled via 2A upload)
    except Exception:
        return {}
    return avg

def fetch_week_stats(team: str, opponent: str, week_int: int, season: int) -> Dict[str, float]:
    result: Dict[str, float] = {}
    if nfl is None:
        return result
    try:
        weekly = nfl.import_weekly_data([season])
        df = weekly.copy()
        team_col = "team" if "team" in df.columns else ("posteam" if "posteam" in df.columns else None)
        if team_col and "week" in df.columns:
            tdf = df[(df[team_col].astype(str).str.upper()==team.upper()) & (df["week"]==week_int)]
            if not tdf.empty:
                pass_yds = float(tdf.get("passing_yards", pd.Series(dtype=float)).sum() or 0)
                rush_yds = float(tdf.get("rushing_yards", pd.Series(dtype=float)).sum() or 0)
                pass_td = float(tdf.get("passing_tds", pd.Series(dtype=float)).sum() or 0)
                rush_td = float(tdf.get("rushing_tds", pd.Series(dtype=float)).sum() or 0)
                comp = tdf.get("completions", pd.Series(dtype=float)).sum() if "completions" in tdf else None
                att  = tdf.get("attempts", pd.Series(dtype=float)).sum() if "attempts" in tdf else None
                ypa  = float(pass_yds/att) if att and att!=0 else pd.NA
                cmp  = float((comp/att)*100) if comp is not None and att and att!=0 else pd.NA
                result.update({
                    "PassYds": pass_yds,
                    "RushYds": rush_yds,
                    "PassTD": pass_td,
                    "RushTD": rush_td,
                    "YPA": ypa,
                    "CMP%": cmp,
                    "Total_YDS": pass_yds + rush_yds
                })
    except Exception:
        return {}
    return result

def style_by_nfl_avg(df: pd.DataFrame, sheet_name: str):
    better_high = []
    better_low = []
    if sheet_name == "Offense":
        better_high = ["PassYds","RushYds","RecvYds","PassTD","RushTD","Points","YPA","CMP%","YAC","Targets","Recs","1D","Total_YDS","QBR","SR%","EPA","DVOA","Red_Zone%","3rd_Down_Conv%","Expl_Plays (20+)"]
        better_low = ["Turnovers","SacksAgainst"]
    elif sheet_name == "Defense":
        better_high = ["Sacks","QB_Hits","Pressures","INT","Turnovers","FF","FR"]
        better_low = ["3rd_Down%_Allowed","RZ%_Allowed","YdsAllowed_Pass","YdsAllowed_Rush","PointsAllowed"]
    avg_map = {c.replace("NFL_Avg._",""): c for c in df.columns if c.startswith("NFL_Avg._")}

    def highlight(row):
        styles = [""] * len(row)
        for i, col in enumerate(df.columns):
            if col in ("Week","Opponent","Team") or col.startswith("NFL_Avg._"):
                continue
            base = col
            avg_col = avg_map.get(base)
            if not avg_col or avg_col not in df.columns:
                continue
            try:
                val = float(row[col])
                avg = float(row[avg_col])
            except Exception:
                continue
            if base in better_high:
                styles[i] = "background-color: #e6ffe6;" if val > avg else "background-color: #ffe6e6;"
            elif base in better_low:
                styles[i] = "background-color: #e6ffe6;" if val < avg else "background-color: #ffe6e6;"
        return styles

    try:
        return df.style.apply(highlight, axis=1)
    except Exception:
        return df

# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.header("Controls")
    week_input = st.text_input("Week (e.g., W01)", value="W01")
    try:
        week_int = int(week_input.replace("W",""))
    except Exception:
        week_int = 1
    opponent_input = st.selectbox("Opponent (3-letter)", options=OPPONENTS, index=0)
    st.caption(f"Season assumed: {CURRENT_SEASON}")

    st.subheader("Data Operations")
    fetch_btn = st.button("Fetch NFL Data (Auto)")
    compute_avg_btn = st.button("Compute NFL Averages")

    st.subheader("Exports")
    export_xlsx_btn = st.button("Export Week Excel")
    export_pdf_btn = st.button("Export Final PDF")

    st.subheader("Display")
    color_toggle = st.toggle("Enable Color Coding", value=True, key="color_coding")

st.title("🐻 Chicago Bears Weekly Tracker — Revised")

st.markdown("""
**Sections (this page renders in this order):**
1) Weekly Controls 
2) Upload Weekly Files (Offense, Defense, Personnel, Snap Counts) 
2A) Upload NFL Averages (Manual) 
3) Key Notes & Media Summaries 
4) Injuries 
5) Opponent Preview & Weekly Game Predictions 
6) Exports & Downloads 
""")

# -----------------------------
# 1) Weekly Controls
# -----------------------------
with st.expander("1) Weekly Controls", expanded=True):
    c1, c2, c3 = st.columns([1,1,2])
    with c1: st.write(f"**Selected Week:** {week_input}")
    with c2: st.write(f"**Opponent:** {opponent_input}")
    with c3: st.caption("Use the sidebar to change Week and Opponent.")

    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        if st.button("Fetch NFL Data (Auto)") or fetch_btn:
            stats = fetch_week_stats("CHI", opponent_input, week_int, CURRENT_SEASON)
            if stats:
                off_row = {c: pd.NA for c in OFF_EXPANDED_COLS}
                off_row.update({"Week":week_input,"Opponent":opponent_input,"Team":"CHI"})
                for k,v in stats.items():
                    if k in off_row: off_row[k] = v
                df_off = pd.DataFrame([off_row])
                upsert_sheet(DATA_FILE, "Offense", df_off)
                st.success("Fetched NFL weekly data and merged into Offense.")
            else:
                st.warning("No NFL data merged (source unavailable). Use Quick Manual Entry.")

    with cc2:
        if st.button("Compute NFL Averages") or compute_avg_btn:
            avg_map = compute_nfl_averages_for_week(week_int, CURRENT_SEASON)
            if not avg_map:
                st.info("Could not compute NFL averages (source down). Use Section 2A upload instead.")
            else:
                wb = openpyxl.load_workbook(DATA_FILE)
                for sheet_name in ["Offense","Defense","Personnel"]:
                    if sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        data = list(ws.values)
                        if not data: continue
                        header = list(data[0])
                        rows = data[1:]
                        df = pd.DataFrame(rows, columns=header)
                        for k, v in avg_map.items():
                            if k in df.columns and df[k].isna().all():
                                df[k] = v
                        ws.delete_rows(1, ws.max_row)
                        ws.append(df.columns.tolist())
                        for _, r in df.iterrows():
                            ws.append(r.tolist())
                wb.save(DATA_FILE)
                st.success("NFL averages filled where available.")

    with cc3:
        st.markdown("**Quick Manual Entry (use if fetch is down)**")
        qm_cols1, qm_cols2 = st.columns(2)
        with qm_cols1:
            qm_points = st.number_input("Points", 0, 80, 0)
            qm_passyds = st.number_input("PassYds", 0, 800, 0)
            qm_rushyds = st.number_input("RushYds", 0, 500, 0)
            qm_passtd = st.number_input("PassTD", 0, 10, 0)
            qm_rushtd = st.number_input("RushTD", 0, 10, 0)
        with qm_cols2:
            qm_turn = st.number_input("Turnovers", 0, 10, 0)
            qm_sacksag = st.number_input("SacksAgainst", 0, 15, 0)
            qm_def_pa = st.number_input("PointsAllowed (DEF)", 0, 80, 0)
            qm_def_sacks = st.number_input("Sacks (DEF)", 0, 15, 0)
            qm_def_int = st.number_input("INT (DEF)", 0, 10, 0)
        if st.button("Save Quick Manual Stats"):
            off_row = {c: pd.NA for c in OFF_EXPANDED_COLS}
            off_row.update({"Week":week_input,"Opponent":opponent_input,"Team":"CHI",
                            "Points":qm_points,"PassYds":qm_passyds,"RushYds":qm_rushyds,
                            "PassTD":qm_passtd,"RushTD":qm_rushtd,"Turnovers":qm_turn,
                            "SacksAgainst":qm_sacksag,"Total_YDS":qm_passyds+qm_rushyds})
            df_off = pd.DataFrame([off_row])
            upsert_sheet(DATA_FILE,"Offense",df_off)

            def_row = {c: pd.NA for c in DEF_EXPANDED_COLS}
            def_row.update({"Week":week_input,"Opponent":opponent_input,"Team":"CHI",
                            "PointsAllowed":qm_def_pa,"Sacks":qm_def_sacks,"INT":qm_def_int})
            df_def = pd.DataFrame([def_row])
            upsert_sheet(DATA_FILE,"Defense",df_def)

            st.success("Quick manual stats saved for the selected week.")

# -----------------------------
# 2) Upload Weekly Files
# -----------------------------
with st.expander("2) Upload Weekly Files (Offense, Defense, Personnel, Snap Counts)", expanded=True):
    c = st.columns(4)
    with c[0]:
        off_file = st.file_uploader("Upload Offense CSV", type=["csv"], key="up_off")
    with c[1]:
        def_file = st.file_uploader("Upload Defense CSV", type=["csv"], key="up_def")
    with c[2]:
        per_file = st.file_uploader("Upload Personnel CSV", type=["csv"], key="up_per")
    with c[3]:
        sc_file = st.file_uploader("Upload Snap Counts CSV", type=["csv"], key="up_sc")

    if off_file:
        df = pd.read_csv(off_file)
        df = ensure_columns(df, OFF_EXPANDED_COLS)
        df = clean_df(df, opponent_input, "CHI")
        upsert_sheet(DATA_FILE,"Offense",df)
        st.success(f"Offense rows saved: {len(df)}")

    if def_file:
        df = pd.read_csv(def_file)
        df = ensure_columns(df, DEF_EXPANDED_COLS)
        df = clean_df(df, opponent_input, "CHI")
        upsert_sheet(DATA_FILE,"Defense",df)
        st.success(f"Defense rows saved: {len(df)}")

    if per_file:
        df = pd.read_csv(per_file)
        df = ensure_columns(df, PERSONNEL_COLS)
        df = clean_df(df, opponent_input, "CHI")
        upsert_sheet(DATA_FILE,"Personnel",df)
        st.success(f"Personnel rows saved: {len(df)}")

    if sc_file:
        df = pd.read_csv(sc_file)
        df = ensure_columns(df, SNAPCOUNTS_COLS)
        df = clean_df(df, opponent_input, "CHI")
        upsert_sheet(DATA_FILE,"Snap_Counts",df)
        st.success(f"Snap Counts rows saved: {len(df)}")

# -----------------------------
# 2A) Upload NFL Averages (Manual)
# -----------------------------
with st.expander("2A) Upload NFL Averages (Manual)"):
    nflavg_file = st.file_uploader("Upload NFL Averages CSV (columns should be NFL_Avg._*)", type=["csv"], key="nflavg_up")
    if nflavg_file:
        df_avg = pd.read_csv(nflavg_file)
        wb = openpyxl.load_workbook(DATA_FILE)
        changed = False
        for sheet_name in ["Offense","Defense","Personnel"]:
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                data = list(ws.values)
                if not data: continue
                header = list(data[0])
                rows = data[1:]
                df = pd.DataFrame(rows, columns=header)
                for c in df.columns:
                    if c.startswith("NFL_Avg._") and c in df_avg.columns:
                        df[c] = df[c].fillna(df_avg[c])
                        changed = True
                ws.delete_rows(1, ws.max_row)
                ws.append(df.columns.tolist())
                for _, r in df.iterrows():
                    ws.append(r.tolist())
        if changed:
            wb.save(DATA_FILE)
            st.success("NFL averages applied from uploaded CSV.")
        else:
            st.info("No matching NFL_Avg._* columns found to fill.")

# -----------------------------
# 3) Key Notes & Media Summaries
# -----------------------------
with st.expander("3) Key Notes & Media Summaries"):
    st.subheader("Key Notes")
    notes = st.text_area("Notes for this week", height=120, key="notes_text")
    if st.button("Save Notes"):
        df = pd.DataFrame([{ "Week": week_input, "Opponent": opponent_input, "Notes": notes }])
        upsert_sheet(DATA_FILE, "Weekly_Notes", df, key_cols=["Week","Opponent"])
        st.success("Notes saved.")

    st.subheader("Media Summaries")
    source = st.text_input("Source (e.g., ESPN, The Athletic)")
    summary = st.text_area("Summary", height=150)
    if st.button("Save Media Summary"):
        df = pd.DataFrame([{ "Week": week_input, "Opponent": opponent_input, "Source": source, "Summary": summary }])
        upsert_sheet(DATA_FILE, "Media", df, key_cols=["Week","Opponent","Source"])
        st.success("Media summary saved.")

# -----------------------------
# 4) Injuries
# -----------------------------
with st.expander("4) Injuries"):
    c1, c2, c3 = st.columns(3)
    with c1:
        inj_player = st.text_input("Player")
        inj_status = st.selectbox("Status", ["Questionable","Doubtful","Out","Active","IR"], index=0)
        inj_part = st.text_input("Body Part")
    with c2:
        inj_practice = st.selectbox("Practice", ["DNP","Limited","Full","NA"], index=3)
        inj_game = st.selectbox("Game Status", ["Active","Inactive","TBD"], index=2)
    with c3:
        inj_notes = st.text_area("Notes", height=100)
    if st.button("Save Injury"):
        df = pd.DataFrame([{
            "Week": week_input, "Opponent": opponent_input, "Team": "CHI",
            "Player": inj_player, "Status": inj_status, "BodyPart": inj_part,
            "Practice": inj_practice, "GameStatus": inj_game, "Notes": inj_notes
        }])
        upsert_sheet(DATA_FILE, "Injuries", df, key_cols=["Week","Opponent","Player"])
        st.success("Injury saved.")

# -----------------------------
# 5) Opponent Preview & Weekly Game Predictions
# -----------------------------
with st.expander("5) Opponent Preview & Weekly Game Predictions"):
    st.subheader("Opponent Preview")
    k_threats = st.text_area("Key Threats", height=80)
    tendencies = st.text_area("Tendencies", height=80)
    opp_inj = st.text_area("Opponent Injuries", height=80)
    opp_notes = st.text_area("Notes", height=80)
    if st.button("Save Opponent Preview"):
        df = pd.DataFrame([{
            "Week": week_input, "Opponent": opponent_input,
            "Key_Threats": k_threats, "Tendencies": tendencies, "Injuries": opp_inj, "Notes": opp_notes
        }])
        upsert_sheet(DATA_FILE, "Opponent_Preview", df, key_cols=["Week","Opponent"])
        st.success("Opponent preview saved.")

    st.subheader("Weekly Game Predictions")
    pred_winner = st.selectbox("Predicted Winner", options=["CHI", opponent_input])
    pred_conf = st.slider("Confidence (1–100)", 1, 100, 60)
    pred_rat = st.text_area("Rationale", height=100)
    if st.button("Save Prediction"):
        df = pd.DataFrame([{
            "Week": week_input, "Opponent": opponent_input,
            "Predicted_Winner": pred_winner, "Confidence": pred_conf, "Rationale": pred_rat
        }])
        upsert_sheet(DATA_FILE, "Predictions", df, key_cols=["Week","Opponent"])
        st.success("Prediction saved.")

# -----------------------------
# Live Data Preview (with optional color coding)
# -----------------------------
st.markdown("### Live Data Preview (current Excel)")
init_workbook(DATA_FILE)
wb = openpyxl.load_workbook(DATA_FILE)

def style_by_nfl_avg_wrapper(df, sheet_name):
    if st.session_state.get("color_coding", True) and sheet_name in ["Offense","Defense"]:
        try:
            return style_by_nfl_avg(df, sheet_name)
        except Exception:
            return df
    return df

def show_sheet(sheet_name: str, cols: List[str]):
    if sheet_name not in wb.sheetnames:
        st.write(f"({sheet_name} is empty)")
        return
    ws = wb[sheet_name]
    data = list(ws.values)
    if not data:
        st.write("(empty)")
        return
    header = list(data[0])
    rows = data[1:]
    df = pd.DataFrame(rows, columns=header).reindex(columns=cols, fill_value=pd.NA)
    styled = style_by_nfl_avg_wrapper(df, sheet_name)
    st.dataframe(styled, use_container_width=True)

tab_objs = st.tabs([s for s,_ in SHEET_ORDER])
for tab, (sname, scols) in zip(tab_objs, SHEET_ORDER):
    with tab:
        show_sheet(sname, scols)

# -----------------------------
# 6) Exports & Downloads
# -----------------------------
def export_week_excel(week_str: str, opponent: str) -> str:
    out_path = os.path.join(EXPORTS_DIR, f"{week_str}_{opponent}.xlsx")
    owb = openpyxl.Workbook()
    owb.remove(owb.active)
    for sheet_name, cols in SHEET_ORDER[:4]:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            data = list(ws.values)
            if not data:
                continue
            header = list(data[0])
            rows = data[1:]
            df = pd.DataFrame(rows, columns=header)
            if all(c in df.columns for c in ["Week","Opponent"]):
                df = df[(df["Week"]==week_str) & (df["Opponent"]==opponent)]
            ws_out = owb.create_sheet(sheet_name)
            ws_out.append(df.columns.tolist())
            for _, r in df.iterrows():
                ws_out.append(r.tolist())
    owb.save(out_path)
    return out_path

def export_final_pdf(week_str: str, opponent: str) -> str:
    out_path = os.path.join(EXPORTS_DIR, f"{week_str}_{opponent}_Final.pdf")
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, f"Bears — Week {week_str.replace('W','')} vs {opponent}", ln=True)
    pdf.set_font("Arial", size=10)

    def add_table(title: str, df: pd.DataFrame):
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, title, ln=True)
        pdf.set_font("Arial", size=9)
        if df.empty:
            pdf.cell(0, 6, "(no data)", ln=True)
            return
        col_w = max(25, pdf.w / max(6, len(df.columns)))
        for c in df.columns:
            pdf.cell(col_w, 6, str(c)[:18], border=1)
        pdf.ln(6)
        for _, r in df.iterrows():
            for c in df.columns:
                pdf.cell(col_w, 6, str(r[c])[:18], border=1)
            pdf.ln(6)
        pdf.ln(2)

    for sheet_name in ["Offense","Defense","Personnel","Snap_Counts"]:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            data = list(ws.values)
            if data:
                header = list(data[0]); rows = data[1:]
                df = pd.DataFrame(rows, columns=header)
                if all(c in df.columns for c in ["Week","Opponent"]):
                    df = df[(df["Week"]==week_str) & (df["Opponent"]==opponent)]
            else:
                df = pd.DataFrame()
        else:
            df = pd.DataFrame()
        add_table(sheet_name, df)

    pdf.output(out_path)
    return out_path

with st.expander("6) Exports & Downloads", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Export Week Excel") or export_xlsx_btn:
            path = export_week_excel(week_input, opponent_input)
            if path:
                st.success(f"Excel exported → {path}")
                with open(path, "rb") as f:
                    st.download_button("Download Week Excel", f, file_name=os.path.basename(path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        if st.button("Export Final PDF") or export_pdf_btn:
            path = export_final_pdf(week_input, opponent_input)
            if path:
                st.success(f"PDF exported → {path}")
                with open(path, "rb") as f:
                    st.download_button("Download Final PDF", f, file_name=os.path.basename(path), mime="application/pdf")
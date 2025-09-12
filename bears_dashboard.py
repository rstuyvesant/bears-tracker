import os
import io
from typing import Dict, List

import streamlit as st
import pandas as pd
from datetime import datetime

# Optional imports guarded for environments where they may not exist
try:
    import nfl_data_py as nfl
except Exception:
    nfl = None

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

try:
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows
except Exception:
    openpyxl = None

# -----------------------------
# App Config
# -----------------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker — Full Revision (2025‑09‑11)")

# Project paths
BASE_DIR = os.getcwd()
DATA_FILE = os.path.join(BASE_DIR, "bears_weekly_analytics.xlsx")
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
os.makedirs(EXPORTS_DIR, exist_ok=True)

# -----------------------------
# Utilities
# -----------------------------
NFL_TEAMS = [
    "ARI","ATL","BAL","BUF","CAR","CHI","CIN","CLE","DAL","DEN","DET",
    "GB","HOU","IND","JAX","KC","LV","LAC","LAR","MIA","MIN","NE","NO",
    "NYG","NYJ","PHI","PIT","SEA","SF","TB","TEN","WAS"
]

CURRENT_SEASON = datetime.now().year if datetime.now().month >= 9 else datetime.now().year - 1

OPPONENTS = [t for t in NFL_TEAMS if t != "CHI"]

DEF_EXPANDED_COLS = [
    "Week","Opponent","Team",
    "Sacks","QB_Hits","Pressures","INT","Turnovers","FF","FR",
    "3rd_Down%_Allowed","RZ%_Allowed","YdsAllowed_Pass","YdsAllowed_Rush","PointsAllowed",
    "NFL_Avg._Sacks","NFL_Avg._QB_Hits","NFL_Avg._Pressures","NFL_Avg._INT","NFL_Avg._Turnovers",
    "NFL_Avg._FF","NFL_Avg._FR","NFL_Avg._3rd_Down%_Allowed","NFL_Avg._RZ%_Allowed",
    "NFL_Avg._YdsAllowed_Pass","NFL_Avg._YdsAllowed_Rush","NFL_Avg._PointsAllowed"
]

OFF_EXPANDED_COLS = [
    "Week","Opponent","Team",
    "PassYds","RushYds","RecvYds",
    "PassTD","RushTD","Points",
    "Turnovers","SacksAgainst",
    "YPA","CMP%","YAC","Targets","Recs",
    "1D","Total_YDS",
    "QBR","SR%","EPA","DVOA","Red_Zone%","3rd_Down_Conv%","Expl_Plays (20+)",
    "NFL_Avg._PassYds","NFL_Avg._RushYds","NFL_Avg._PassTD","NFL_Avg._RushTD",
    "NFL_Avg._Points","NFL_Avg._Turnovers","NFL_Avg._SacksAgainst",
    "NFL_Avg._YPA","NFL_Avg._CMP%","NFL_Avg._YAC","NFL_Avg._Targets",
    "NFL_Avg._Recs","NFL_Avg._1D","NFL_Avg._Total_YDS",
    "NFL_Avg._QBR","NFL_Avg._SR%","NFL_Avg._EPA","NFL_Avg._DVOA",
    "NFL_Avg._Red_Zone%","NFL_Avg._3rd_Down_Conv%","NFL_Avg._Expl_Plays (20+)"
]

PERSONNEL_COLS = [
    "Week","Opponent","Team",
    "11_Personnel","12_Personnel","13_Personnel","21_Personnel","Other_Personnel","Total_Snaps",
    "NFL_Avg._11_Personnel","NFL_Avg._12_Personnel","NFL_Avg._13_Personnel",
    "NFL_Avg._21_Personnel","NFL_Avg._Other_Personnel","NFL_Avg._Total_Snaps"
]

SNAPCOUNTS_COLS = [
    "Week","Opponent","Team",
    # Flexible player columns; users can rename in CSVs without breaking the app
    "Player1","Player2","Player3","Player4","Player5","Player6","Player7","Player8","Player9","Player10","Player11",
    "NFL_Avg._RBs","NFL_Avg._WRs","NFL_Avg._TEs","NFL_Avg._QB","NFL_Avg._OL","NFL_Avg._DL","NFL_Avg._LB","NFL_Avg._DB"
]

SHEET_ORDER = [
    ("Offense", OFF_EXPANDED_COLS),
    ("Defense", DEF_EXPANDED_COLS),
    ("Personnel", PERSONNEL_COLS),
    ("Snap_Counts", SNAPCOUNTS_COLS),
    ("Injuries", ["Week","Opponent","Team","Player","Status","BodyPart","Practice","GameStatus","Notes"]),
    ("Media", ["Week","Opponent","Source","Summary"]),
    ("Opponent_Preview", ["Week","Opponent","Key_Threats","Tendencies","Injuries","Notes"]),
    ("Predictions", ["Week","Opponent","Predicted_Winner","Confidence","Rationale"]),
    ("Weekly_Notes", ["Week","Opponent","Notes"]),
    ("Weekly_Strategy", ["Week","Opponent","Off_Strategy","Def_Strategy","Key_Matchups","Keys_to_Win"]) 
]


def normalize_week(val) -> str:
    if pd.isna(val):
        return "W01"
    try:
        ival = int(float(val))
        return f"W{ival:02d}"
    except Exception:
        s = str(val).strip().upper()
        if s.startswith("W") and len(s) >= 3:
            return s
        return "W01"


def ensure_columns(df: pd.DataFrame, required: List[str]) -> pd.DataFrame:
    for col in required:
        if col not in df.columns:
            df[col] = pd.NA
    # Put Week/Opponent/Team first when present
    front = [c for c in ["Week","Opponent","Team"] if c in required]
    ordered = front + [c for c in required if c not in front]
    return df.reindex(columns=ordered)


def clean_df(df: pd.DataFrame, opponent: str, team: str = "CHI") -> pd.DataFrame:
    # Remove unnamed columns
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")].copy()

    # Fill Week/Opponent/Team
    if "Week" not in df.columns:
        df["Week"] = "W01"
    df["Week"] = df["Week"].apply(normalize_week)

    if "Opponent" not in df.columns:
        df["Opponent"] = opponent
    df["Opponent"] = df["Opponent"].fillna(opponent)

    if "Team" not in df.columns:
        df["Team"] = team
    df["Team"] = df["Team"].fillna(team)

    # Drop rows that are entirely blank across metrics (excluding identifiers)
    id_cols = {"Week","Opponent","Team"}
    metric_cols = [c for c in df.columns if c not in id_cols]
    mask = df[metric_cols].apply(lambda r: r.isna().all() or (r.astype(str).str.strip()=="").all(), axis=1)
    df = df.loc[~mask].copy()

    return df


# -----------------------------
# Excel helpers
# -----------------------------

def init_workbook(path: str):
    if openpyxl is None:
        st.warning("openpyxl not installed — Excel features limited.")
        return
    if not os.path.exists(path):
        wb = openpyxl.Workbook()
        # remove default sheet
        wb.remove(wb.active)
        for sheet_name, cols in SHEET_ORDER:
            ws = wb.create_sheet(title=sheet_name)
            ws.append(cols)
        wb.save(path)


def upsert_sheet(path: str, sheet_name: str, df_new: pd.DataFrame, key_cols: List[str] = ["Week","Opponent","Team"]):
    if openpyxl is None:
        st.error("openpyxl not available; cannot write Excel.")
        return
    if not os.path.exists(path):
        init_workbook(path)
    wb = openpyxl.load_workbook(path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(title=sheet_name)
        ws.append(df_new.columns.tolist())
    else:
        ws = wb[sheet_name]

    # Read existing sheet to DataFrame
    data = list(ws.values)
    if len(data) == 0:
        existing = pd.DataFrame(columns=df_new.columns)
    else:
        header = list(data[0])
        rows = data[1:]
        existing = pd.DataFrame(rows, columns=header)

    # Ensure consistent columns
    all_cols = list(dict.fromkeys(list(existing.columns) + list(df_new.columns)))
    existing = existing.reindex(columns=all_cols)
    df_new = df_new.reindex(columns=all_cols)

    # Normalize keys
    for k in key_cols:
        if k in existing.columns:
            existing[k] = existing[k].astype(str)
        if k in df_new.columns:
            df_new[k] = df_new[k].astype(str)

    # Drop duplicates in existing
    if all(k in existing.columns for k in key_cols):
        existing = existing.drop_duplicates(subset=key_cols, keep="last")

    # Upsert: remove overlaps, then concat
    if all(k in df_new.columns for k in key_cols):
        key_vals = set(tuple(x) for x in df_new[key_cols].to_numpy())
        if not existing.empty:
            mask = existing[key_cols].apply(lambda r: tuple(r.values) in key_vals, axis=1)
            existing = existing.loc[~mask]
    combined = pd.concat([existing, df_new], ignore_index=True)

    # Clear and write back
    ws.delete_rows(1, ws.max_row)
    ws.append(combined.columns.tolist())
    for _, row in combined.iterrows():
        ws.append(row.tolist())

    wb.save(path)


# -----------------------------
# NFL Fetch + Averages
# -----------------------------

def fetch_week_stats(team: str, opponent: str, week_int: int, season: int) -> Dict[str, float]:
    """Best‑effort multi‑source fetch using nfl_data_py.
    Tries several endpoints because upstream URLs sometimes change (404s).
    """
    result: Dict[str, float] = {}
    if nfl is None:
        st.warning("nfl_data_py not installed; cannot auto-fetch NFL data.")
        return result

    def safe_try(func, *args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            st.info(f"Fallback note: {getattr(func, '__name__', 'call')} failed → {e}")
            return None

    # 1) Primary: weekly team stats (fast)
    weekly = safe_try(nfl.import_weekly_data, [season])
    if weekly is not None and not weekly.empty:
        df = weekly.copy()
        # Try common team column names
        team_col = "team" if "team" in df.columns else ("posteam" if "posteam" in df.columns else None)
        if team_col is not None and "week" in df.columns:
            df_team = df[(df[team_col].astype(str).str.upper()==team.upper()) & (df["week"]==week_int)]
            if not df_team.empty:
                # Aggregate core offensive metrics
                pass_yds = float(df_team.get("passing_yards", pd.Series(dtype=float)).sum() or 0)
                if pass_yds == 0 and "pass_yards" in df_team: pass_yds = float(df_team["pass_yards"].sum())
                rush_yds = float(df_team.get("rushing_yards", pd.Series(dtype=float)).sum() or 0)
                if rush_yds == 0 and "rush_yards" in df_team: rush_yds = float(df_team["rush_yards"].sum())
                pass_td = float(df_team.get("passing_tds", pd.Series(dtype=float)).sum() or 0)
                rush_td = float(df_team.get("rushing_tds", pd.Series(dtype=float)).sum() or 0)
                comp = df_team.get("completions", pd.Series(dtype=float)).sum() if "completions" in df_team else None
                att  = df_team.get("attempts", pd.Series(dtype=float)).sum() if "attempts" in df_team else None
                ypa  = float(pass_yds/att) if att and att!=0 else pd.NA
                cmp  = float((comp/att)*100) if comp is not None and att and att!=0 else pd.NA
                result.update({
                    "PassYds": pass_yds,
                    "RushYds": rush_yds,
                    "PassTD": pass_td,
                    "RushTD": rush_td,
                    "YPA": ypa,
                    "CMP%": cmp,
                    "RecvYds": pass_yds if pd.isna(pass_yds)==False else pd.NA,
                    "Total_YDS": (pass_yds + rush_yds)
                })

    # 2) Fallback: team seasonal summary (per‑game averages) — then take week row if present
    if not result:
        team_summ = safe_try(nfl.import_seasonal_team_stats, [season])
        if team_summ is not None and not team_summ.empty:
            df = team_summ.copy()
            # If a weekly dimension exists, try filter; otherwise leave as placeholder
            if "team" in df.columns:
                df_team = df[df["team"].astype(str).str.upper()==team.upper()]
                if not df_team.empty:
                    # Use per‑game placeholders so exports have something (better than blank)
                    result.update({
                        "PassYds": float(df_team.get("pass_yds", pd.Series(dtype=float)).mean() or 0),
                        "RushYds": float(df_team.get("rush_yds", pd.Series(dtype=float)).mean() or 0),
                        "PassTD": float(df_team.get("pass_tds", pd.Series(dtype=float)).mean() or 0),
                        "RushTD": float(df_team.get("rush_tds", pd.Series(dtype=float)).mean() or 0),
                    })

    # 3) Fallback: schedules/boxscores (if exposed)
    if not result:
        sched = safe_try(nfl.import_schedules, [season])
        if sched is not None and not sched.empty:
            # Try to locate game id for week
            df = sched.copy()
            # Abbreviation columns vary: team, team_home, team_away
            if all(c in df.columns for c in ["team_home","team_away","week"]):
                game = df[((df["team_home"].astype(str).str.upper()==team.upper()) | (df["team_away"].astype(str).str.upper()==team.upper())) & (df["week"]==week_int)]
                if not game.empty:
                    # If points columns exist, take them
                    for side in ["home","away"]:
                        pts_col = f"result_{side}" if f"result_{side}" in game.columns else (f"points_{side}" if f"points_{side}" in game.columns else None)
                        if pts_col:
                            try:
                                result["Points"] = float(pd.to_numeric(game[pts_col], errors='coerce').iloc[0])
                                break
                            except Exception:
                                pass

    # Defensive placeholders (we still add them so columns exist downstream)
    if "SacksAgainst" not in result:
        result["SacksAgainst"] = pd.NA
    for k in [
        "Sacks","QB_Hits","Pressures","INT","Turnovers","FF","FR",
        "3rd_Down%_Allowed","RZ%_Allowed","YdsAllowed_Pass","YdsAllowed_Rush","PointsAllowed"
    ]:
        result.setdefault(f"DEF__{k}", pd.NA)

    return result
    try:
        # nflfastR weekly data (play-by-play derived). Using weekly is more stable than schedules.
        weekly = nfl.import_weekly_data([season])  # can be large; filtered next
        # Normalize team abbreviations
        df = weekly.copy()
        # Columns differ across versions; try common patterns
        # Filter for team & week
        if "team" in df.columns:
            df_team = df[(df["team"].str.upper()==team.upper()) & (df["week"]==week_int)]
        elif "posteam" in df.columns:
            df_team = df[(df["posteam"].str.upper()==team.upper()) & (df["week"]==week_int)]
        else:
            df_team = df[df.get("week", pd.Series(dtype=int))==week_int]

        if df_team.empty:
            st.info("No rows found for team/week in weekly dataset; data may not be published yet.")
            return result

        # Aggregate common metrics (best-effort)
        # Try both snake_case and alt names
        def first_available(row: pd.Series, names: List[str]):
            for n in names:
                if n in row and pd.notna(row[n]):
                    return row[n]
            return pd.NA

        # Use sums/means where appropriate
        # Passing yards
        result["PassYds"] = float(df_team.get("passing_yards", pd.Series(dtype=float)).sum() or 0)
        if result["PassYds"] == 0 and "pass_yards" in df_team:
            result["PassYds"] = float(df_team["pass_yards"].sum())

        result["RushYds"] = float(df_team.get("rushing_yards", pd.Series(dtype=float)).sum() or 0)
        if result["RushYds"] == 0 and "rush_yards" in df_team:
            result["RushYds"] = float(df_team["rush_yards"].sum())

        result["PassTD"] = float(df_team.get("passing_tds", pd.Series(dtype=float)).sum() or 0)
        result["RushTD"] = float(df_team.get("rushing_tds", pd.Series(dtype=float)).sum() or 0)

        # Sacks against (offense): sacks taken often stored as sacks, positive is defensive sacks.
        if "sacks" in df_team:
            # If positive means sacks by defense, sacks against offense is negative of own sacks? Use abs total for week.
            result["SacksAgainst"] = float(abs(df_team["sacks"].sum()))
        else:
            result["SacksAgainst"] = 0.0

        # Completions/Attempts for YPA & CMP%
        comp = df_team.get("completions", pd.Series(dtype=float)).sum() if "completions" in df_team else None
        att = df_team.get("attempts", pd.Series(dtype=float)).sum() if "attempts" in df_team else None
        if att and att != 0:
            result["YPA"] = float(result["PassYds"]/att)
            result["CMP%"] = float((comp/att)*100) if comp is not None else pd.NA
        else:
            result["YPA"] = pd.NA
            result["CMP%"] = pd.NA

        # Basic placeholders if data not present in that dataset
        result.setdefault("RecvYds", result.get("PassYds", 0))
        result.setdefault("Points", pd.NA)
        result.setdefault("Turnovers", pd.NA)
        result.setdefault("YAC", pd.NA)
        result.setdefault("Targets", pd.NA)
        result.setdefault("Recs", comp if comp is not None else pd.NA)
        result.setdefault("1D", pd.NA)
        result.setdefault("Total_YDS", (result.get("PassYds",0) + result.get("RushYds",0)))
        result.setdefault("QBR", pd.NA)
        result.setdefault("SR%", pd.NA)
        result.setdefault("EPA", pd.NA)
        result.setdefault("DVOA", pd.NA)
        result.setdefault("Red_Zone%", pd.NA)
        result.setdefault("3rd_Down_Conv%", pd.NA)
        result.setdefault("Expl_Plays (20+)", pd.NA)

        # Defense — derive limited items if possible (these are tricky from weekly team table)
        result_def = {
            "Sacks": pd.NA,
            "QB_Hits": pd.NA,
            "Pressures": pd.NA,
            "INT": pd.NA,
            "Turnovers": pd.NA,
            "FF": pd.NA,
            "FR": pd.NA,
            "3rd_Down%_Allowed": pd.NA,
            "RZ%_Allowed": pd.NA,
            "YdsAllowed_Pass": pd.NA,
            "YdsAllowed_Rush": pd.NA,
            "PointsAllowed": pd.NA,
        }
        result.update({f"DEF__{k}": v for k,v in result_def.items()})

    except Exception as e:
        st.warning(f"NFL fetch failed: {e}")
    return result


def compute_nfl_averages_for_week(week_int: int, season: int) -> Dict[str, float]:
    """Compute league averages for a subset of metrics. Returns mapping for NFL_Avg._* columns."""
    averages = {}
    if nfl is None:
        return averages
    try:
        weekly = nfl.import_weekly_data([season])
        dfw = weekly[weekly["week"] == week_int]
        if dfw.empty:
            return averages
        # Basic league averages (best-effort based on available columns)
        if "passing_yards" in dfw:
            averages["NFL_Avg._PassYds"] = float(dfw["passing_yards"].mean())
        if "rushing_yards" in dfw:
            averages["NFL_Avg._RushYds"] = float(dfw["rushing_yards"].mean())
        if "passing_tds" in dfw:
            averages["NFL_Avg._PassTD"] = float(dfw["passing_tds"].mean())
        if "rushing_tds" in dfw:
            averages["NFL_Avg._RushTD"] = float(dfw["rushing_tds"].mean())
        # Placeholders for others (left NaN if not computed here)
        for col in OFF_EXPANDED_COLS:
            if col.startswith("NFL_Avg._") and col not in averages:
                averages[col] = pd.NA
        for col in DEF_EXPANDED_COLS:
            if col.startswith("NFL_Avg._") and col not in averages:
                averages[col] = pd.NA
    except Exception as e:
        st.info(f"Could not compute NFL averages: {e}")
    return averages


def fill_nfl_avg_columns(df: pd.DataFrame, avg_map: Dict[str, float]) -> pd.DataFrame:
    for k, v in avg_map.items():
        if k in df.columns and df[k].isna().all():
            df[k] = v
    return df


# -----------------------------
# Sidebar Controls
# -----------------------------
with st.sidebar:
    st.header("Controls")
    week_input = st.text_input("Week (e.g., W01)", value="W01")
    try:
        week_int = int(week_input.replace("W",""))
    except Exception:
        week_int = 1
    opponent_input = st.selectbox("Opponent (3‑letter)", options=OPPONENTS, index=max(0, OPPONENTS.index("MIN") if "MIN" in OPPONENTS else 0))
    st.caption(f"Season assumed: {CURRENT_SEASON}")

    st.subheader("Data Operations")
    fetch_btn = st.button("Fetch NFL Data (Auto)")
    compute_avg_btn = st.button("Compute NFL Averages")

    st.subheader("Exports")
    pre_xls_btn = st.button("Export Week Excel (Pre/Post)")
    final_pdf_btn = st.button("Export Final PDF")

# Ensure workbook exists
init_workbook(DATA_FILE)

# -----------------------------
# Main Page — Uploaders & Live Views
# -----------------------------
st.markdown("### Weekly Uploads")
cols = st.columns(4)

with cols[0]:
    off_file = st.file_uploader("Upload Offense CSV", type=["csv"], key="off_up")
with cols[1]:
    def_file = st.file_uploader("Upload Defense CSV", type=["csv"], key="def_up")
with cols[2]:
    per_file = st.file_uploader("Upload Personnel CSV", type=["csv"], key="per_up")
with cols[3]:
    sc_file = st.file_uploader("Upload Snap Counts CSV", type=["csv"], key="sc_up")

# Process uploads
if off_file is not None:
    df_off = pd.read_csv(off_file)
    df_off = ensure_columns(df_off, OFF_EXPANDED_COLS)
    df_off = clean_df(df_off, opponent_input, team="CHI")
    upsert_sheet(DATA_FILE, "Offense", df_off)
    st.success(f"Offense rows saved: {len(df_off)}")

if def_file is not None:
    df_def = pd.read_csv(def_file)
    df_def = ensure_columns(df_def, DEF_EXPANDED_COLS)
    df_def = clean_df(df_def, opponent_input, team="CHI")
    upsert_sheet(DATA_FILE, "Defense", df_def)
    st.success(f"Defense rows saved: {len(df_def)}")

if per_file is not None:
    df_per = pd.read_csv(per_file)
    df_per = ensure_columns(df_per, PERSONNEL_COLS)
    df_per = clean_df(df_per, opponent_input, team="CHI")
    upsert_sheet(DATA_FILE, "Personnel", df_per)
    st.success(f"Personnel rows saved: {len(df_per)}")

if sc_file is not None:
    df_sc = pd.read_csv(sc_file)
    df_sc = ensure_columns(df_sc, SNAPCOUNTS_COLS)
    df_sc = clean_df(df_sc, opponent_input, team="CHI")
    upsert_sheet(DATA_FILE, "Snap_Counts", df_sc)
    st.success(f"Snap Counts rows saved: {len(df_sc)}")

# -----------------------------
# Fetch + Merge
# -----------------------------
if fetch_btn:
    stats = fetch_week_stats(team="CHI", opponent=opponent_input, week_int=week_int, season=CURRENT_SEASON)
    if stats:
        # Build DataFrame for Offense with current identifiers
        off_row = {c: pd.NA for c in OFF_EXPANDED_COLS}
        off_row.update({
            "Week": week_input,
            "Opponent": opponent_input,
            "Team": "CHI",
        })
        for k,v in stats.items():
            if k in OFF_EXPANDED_COLS:
                off_row[k] = v
        df_off = pd.DataFrame([off_row])

        # Minimal Defense row placeholder (fetch fills limited defensive metrics by default)
        def_row = {c: pd.NA for c in DEF_EXPANDED_COLS}
        def_row.update({"Week": week_input, "Opponent": opponent_input, "Team": "CHI"})
        for k,v in stats.items():
            if k.startswith("DEF__"):
                key = k.replace("DEF__", "")
                if key in DEF_EXPANDED_COLS:
                    def_row[key] = v
        df_def = pd.DataFrame([def_row])

        # Compute and apply NFL averages if requested afterwards
        upsert_sheet(DATA_FILE, "Offense", df_off)
        upsert_sheet(DATA_FILE, "Defense", df_def)
        st.success("Fetched NFL weekly data and merged into Excel (Offense/Defense).")
    else:
        st.warning("No NFL data merged (library missing or data not yet available).")

# -----------------------------
# Compute NFL Averages (fills NFL_Avg._* columns when empty)
# -----------------------------
if compute_avg_btn:
    avg_map = compute_nfl_averages_for_week(week_int=week_int, season=CURRENT_SEASON)
    if not avg_map:
        st.info("Could not compute NFL averages (library missing or data unavailable).")
    else:
        # Read, fill, and write back Offense/Defense
        if openpyxl is not None and os.path.exists(DATA_FILE):
            wb = openpyxl.load_workbook(DATA_FILE)
            for sheet_name in ["Offense","Defense","Personnel"]:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    data = list(ws.values)
                    if not data:
                        continue
                    header = list(data[0])
                    rows = data[1:]
                    df = pd.DataFrame(rows, columns=header)
                    df = fill_nfl_avg_columns(df, avg_map)
                    # write back
                    ws.delete_rows(1, ws.max_row)
                    ws.append(df.columns.tolist())
                    for _, r in df.iterrows():
                        ws.append(r.tolist())
            wb.save(DATA_FILE)
            st.success("NFL averages filled where empty.")
        else:
            st.info("Excel not available to update.")

# -----------------------------
# Live Previews (robust to missing sheets)
# -----------------------------

def ensure_sheet_exists(wb, sheet_name: str, cols: List[str]):
    if sheet_name not in wb.sheetnames:
        ws_new = wb.create_sheet(title=sheet_name)
        ws_new.append(cols)
        return ws_new
    return wb[sheet_name]

st.markdown("### Live Data Preview (current Excel)")
if openpyxl is not None and os.path.exists(DATA_FILE):
    wb = openpyxl.load_workbook(DATA_FILE)
    # Make sure the expected sheets exist (prevents KeyError)
    for sname, scols in SHEET_ORDER:
        ensure_sheet_exists(wb, sname, scols)
    wb.save(DATA_FILE)

    tabs = st.tabs([s for s,_ in SHEET_ORDER])
    for tab, (sheet_name, s_cols) in zip(tabs, SHEET_ORDER):
        with tab:
            if sheet_name not in wb.sheetnames:
                st.write("(sheet not found — created with headers; add data to view)")
                continue
            ws = wb[sheet_name]
            data = list(ws.values)
            if not data:
                st.write("(empty)")
            else:
                header = list(data[0])
                rows = data[1:]
                df = pd.DataFrame(rows, columns=header)
                st.dataframe(df, width='stretch')
else:
    st.info("Excel not found yet; upload a CSV or fetch to create it.")

# -----------------------------
# Exports
# -----------------------------

def export_week_excel(week_str: str, opponent: str) -> str:
    """Save a per-week Excel with key sheets filtered to the chosen week/opponent."""
    if openpyxl is None or not os.path.exists(DATA_FILE):
        st.error("Excel engine not available or data file missing.")
        return ""
    wb = openpyxl.load_workbook(DATA_FILE)
    out_path = os.path.join(EXPORTS_DIR, f"{week_str}_{opponent}.xlsx")
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    for sheet_name, _ in SHEET_ORDER[:4]:  # Offense/Defense/Personnel/Snap_Counts
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
            ws_out = out_wb.create_sheet(title=sheet_name)
            ws_out.append(df.columns.tolist())
            for _, r in df.iterrows():
                ws_out.append(r.tolist())

    out_wb.save(out_path)
    return out_path


def export_final_pdf(week_str: str, opponent: str) -> str:
    if FPDF is None:
        st.error("FPDF not installed; cannot create PDF.")
        return ""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Arial", style="B", size=16)
    pdf.cell(0, 10, f"Chicago Bears — Week {week_str.replace('W','')} vs {opponent}", ln=True)

    def add_section(title: str, df: pd.DataFrame):
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 8, title, ln=True)
        pdf.set_font("Arial", size=9)
        if df.empty:
            pdf.cell(0, 6, "(no data)", ln=True)
            return
        # Simple table
        col_w = max(25, pdf.w / max(6, len(df.columns)))
        # header
        for c in df.columns:
            pdf.cell(col_w, 6, str(c)[:18], border=1)
        pdf.ln(6)
        # rows
        for _, r in df.iterrows():
            for c in df.columns:
                pdf.cell(col_w, 6, str(r[c])[:18], border=1)
            pdf.ln(6)
        pdf.ln(2)

    # Pull filtered frames from Excel
    if openpyxl is not None and os.path.exists(DATA_FILE):
        wb = openpyxl.load_workbook(DATA_FILE)
        for sheet_name in ["Offense","Defense","Personnel","Snap_Counts"]:
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                data = list(ws.values)
                if not data:
                    df = pd.DataFrame()
                else:
                    header = list(data[0])
                    rows = data[1:]
                    df = pd.DataFrame(rows, columns=header)
                    if all(c in df.columns for c in ["Week","Opponent"]):
                        df = df[(df["Week"]==week_str) & (df["Opponent"]==opponent)]
                add_section(sheet_name, df)

    out_path = os.path.join(EXPORTS_DIR, f"{week_str}_{opponent}_Final.pdf")
    pdf.output(out_path)
    return out_path


# Buttons wiring
if pre_xls_btn:
    path = export_week_excel(week_input, opponent_input)
    if path:
        st.success(f"Excel exported → {path}")
        with open(path, "rb") as f:
            st.download_button("Download Week Excel", data=f, file_name=os.path.basename(path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if final_pdf_btn:
    path = export_final_pdf(week_input, opponent_input)
    if path:
        st.success(f"Final PDF exported → {path}")
        with open(path, "rb") as f:
            st.download_button("Download Final PDF", data=f, file_name=os.path.basename(path), mime="application/pdf")


st.markdown(
    """
---
**Notes**
- This revision expands the field mapping and ensures **Week / Opponent / Team** are always set on upload.
- `Fetch NFL Data` merges a best‑effort set of offensive stats and placeholders for defense; columns are created even if data is missing upstream.
- `Compute NFL Averages` fills any empty `NFL_Avg._*` columns for the current week where we could compute them.
- Exports are saved to the `exports/` folder under your repo folder. Use the download buttons to grab them directly.
- If you want deeper, play‑by‑play metrics (EPA, SR%, pressures), we can wire a second fetch path to merge from play‑by‑play and team‑defense tables.
"""
)

"""
)

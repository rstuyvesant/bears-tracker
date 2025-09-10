# bears_dashboard.py
# Bears Weekly Tracker with:
# - Auto fetch: NFL team data + Snap Counts (via nfl_data_py if available)
# - Flexible CSV/XLSX ingestion (accepts .csv and common .cvs typo)
# - Strategy tab, Global/Snap search
# - Current week Excel/PDF download (one click)
# - Dedupe by keys to avoid duplicate weeks
#
# NOTE: Auto-fetch uses the library 'nfl_data_py' (your requirements include it).
# If a function isn’t available in your version, the app won’t crash; it will show
# an instruction message instead.

from pathlib import Path
from datetime import datetime
import os
import re

import pandas as pd
import numpy as np
import streamlit as st
from fpdf import FPDF

# Try to import nfl_data_py for auto fetch
try:
    import nfl_data_py as nfl
    NFL_OK = True
except Exception:
    NFL_OK = False

# ================== App basics ==================
st.set_page_config(page_title="Bears Weekly Tracker", layout="wide")
st.title("Bears Weekly Tracker")

BASE_DIR = Path(__file__).parent.resolve()
DATA_DIR = BASE_DIR / "data"
EXPORTS_DIR = BASE_DIR / "exports"
CURR_DIR = EXPORTS_DIR / "Current_Week"
MASTER_XLSX = DATA_DIR / "bears_weekly_analytics.xlsx"

for p in (DATA_DIR, EXPORTS_DIR, CURR_DIR):
    p.mkdir(parents=True, exist_ok=True)

SHEETS = {
    "offense": "Offense",
    "defense": "Defense",
    "personnel": "Personnel",
    "snap": "Snap_Counts",
    "inj": "Injuries",
    "media": "Media",
    "opp": "Opponent_Preview",
    "pred": "Predictions",
    "notes": "Weekly_Notes",
    "strategy": "Weekly_Strategy",
}

NFL_CODES = [
    "ARI","ATL","BAL","BUF","CAR","CHI","CIN","CLE","DAL","DEN","DET","GB","HOU","IND",
    "JAX","KC","LAC","LAR","LV","MIA","MIN","NE","NO","NYG","NYJ","PHI","PIT","SEA","SF","TB","TEN","WAS"
]
FULLNAME_TO_CODE = {
    "arizona cardinals":"ARI","atlanta falcons":"ATL","baltimore ravens":"BAL","buffalo bills":"BUF",
    "carolina panthers":"CAR","chicago bears":"CHI","cincinnati bengals":"CIN","cleveland browns":"CLE",
    "dallas cowboys":"DAL","denver broncos":"DEN","detroit lions":"DET","green bay packers":"GB",
    "houston texans":"HOU","indianapolis colts":"IND","jacksonville jaguars":"JAX","kansas city chiefs":"KC",
    "los angeles chargers":"LAC","los angeles rams":"LAR","las vegas raiders":"LV","miami dolphins":"MIA",
    "minnesota vikings":"MIN","new england patriots":"NE","new orleans saints":"NO","new york giants":"NYG",
    "new york jets":"NYJ","philadelphia eagles":"PHI","pittsburgh steelers":"PIT","seattle seahawks":"SEA",
    "san francisco 49ers":"SF","tampa bay buccaneers":"TB","tennessee titans":"TEN","washington commanders":"WAS",
    # common nicknames
    "chargers":"LAC","rams":"LAR","raiders":"LV","49ers":"SF","niners":"SF","bucs":"TB","buccaneers":"TB",
    "commanders":"WAS","football team":"WAS","vikings":"MIN","packers":"GB","lions":"DET","bears":"CHI",
    "browns":"CLE","bengals":"CIN","steelers":"PIT","ravens":"BAL","bills":"BUF","patriots":"NE","jets":"NYJ",
    "giants":"NYG","eagles":"PHI","cowboys":"DAL","texans":"HOU","colts":"IND","jaguars":"JAX","chiefs":"KC",
    "dolphins":"MIA","falcons":"ATL","panthers":"CAR","broncos":"DEN","seahawks":"SEA","cardinals":"ARI","titans":"TEN"
}

def week_options():
    return [f"W{str(i).zfill(2)}" for i in range(1, 24)]

# ================== Excel helpers ==================
def ensure_master():
    """Create master workbook with all sheets on first run."""
    if MASTER_XLSX.exists():
        return
    with pd.ExcelWriter(MASTER_XLSX, engine="openpyxl", mode="w") as w:
        pd.DataFrame(columns=["Week","Opponent"]).to_excel(w, index=False, sheet_name=SHEETS["offense"])
        pd.DataFrame(columns=["Week","Opponent"]).to_excel(w, index=False, sheet_name=SHEETS["defense"])
        pd.DataFrame(columns=["Week","Opponent","11","12","13","21","Other"]).to_excel(w, index=False, sheet_name=SHEETS["personnel"])
        pd.DataFrame(columns=["Week","Opponent","Player","Snaps","Snap%","Side"]).to_excel(w, index=False, sheet_name=SHEETS["snap"])
        pd.DataFrame(columns=["Week","Opponent","Player","Status","BodyPart","Practice","GameStatus","Notes"]).to_excel(w, index=False, sheet_name=SHEETS["inj"])
        pd.DataFrame(columns=["Week","Opponent","Source","Summary"]).to_excel(w, index=False, sheet_name=SHEETS["media"])
        pd.DataFrame(columns=["Week","Opponent","Off_Summary","Def_Summary","Matchups"]).to_excel(w, index=False, sheet_name=SHEETS["opp"])
        pd.DataFrame(columns=["Week","Opponent","Predicted_Winner","Confidence","Rationale"]).to_excel(w, index=False, sheet_name=SHEETS["pred"])
        pd.DataFrame(columns=["Week","Opponent","Notes"]).to_excel(w, index=False, sheet_name=SHEETS["notes"])
        pd.DataFrame(columns=["Week","Opponent","Plan","Keys","Notes"]).to_excel(w, index=False, sheet_name=SHEETS["strategy"])

def read_sheet(name: str) -> pd.DataFrame:
    ensure_master()
    try:
        return pd.read_excel(MASTER_XLSX, sheet_name=name)
    except Exception:
        return pd.DataFrame()

def write_sheet(df: pd.DataFrame, name: str):
    ensure_master()
    with pd.ExcelWriter(MASTER_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        (df if not df.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name=name)

def append_to_sheet(new_rows: pd.DataFrame, name: str, dedup_cols=None):
    old = read_sheet(name)
    combined = pd.concat([old, new_rows], ignore_index=True)
    if dedup_cols and all(c in combined.columns for c in dedup_cols):
        combined = combined.drop_duplicates(subset=dedup_cols, keep="last").reset_index(drop=True)
    write_sheet(combined, name)

# ================== Flexible ingestion helpers ==================
def _normalize_columns(cols):
    return {c: re.sub(r"[_\s]+", "", str(c).lower()) for c in cols}

def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm = _normalize_columns(df.columns)
    cand_norm = [re.sub(r"[_\s]+","", c.lower()) for c in candidates]
    for original, simple in norm.items():
        if simple in cand_norm:
            return original
    return None

def map_team_to_code(val: str) -> str | None:
    if not isinstance(val, str):
        return None
    v = val.strip().upper()
    if v in NFL_CODES:
        return v
    return FULLNAME_TO_CODE.get(val.strip().lower())

def coerce_week_value(v) -> str | None:
    try:
        if isinstance(v, str) and v.strip().upper().startswith("W"):
            num = int(re.sub(r"[^\d]","", v))
        else:
            num = int(float(v))
        return f"W{num:02d}"
    except Exception:
        return None

def read_any_table(upload) -> pd.DataFrame:
    """Read CSV (.csv or .cvs) or Excel (.xlsx/.xls)."""
    name = (upload.name or "").lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(upload)
    return pd.read_csv(upload, engine="python")  # accepts .csv and .cvs

def ensure_week_opp(df: pd.DataFrame, fallback_week=None, fallback_opp=None, force_fill_if_missing=True) -> tuple[pd.DataFrame, list[str]]:
    """
    Ensure DF has Week and Opponent columns. If missing/not parseable and
    force_fill_if_missing=True, fill from fallback_week/opponent.
    Returns (df, warnings).
    """
    warnings = []
    df2 = df.copy()

    wk_col = _find_col(df2, ["Week","Wk","WK","GameWeek","Game_Week"])
    opp_col = _find_col(df2, ["Opponent","Opp","OppCode","Team","OpponentTeam","OpponentName","Opp Abbr","OppAbbr"])

    if wk_col is not None:
        df2["Week"] = df2[wk_col].map(coerce_week_value)
    if opp_col is not None:
        df2["Opponent"] = df2[opp_col].map(map_team_to_code).fillna(
            df2[opp_col].astype(str).str.upper()
        )

    # Fill from sidebar if missing and allowed
    if ("Week" not in df2.columns or df2["Week"].isna().all()) and force_fill_if_missing and fallback_week:
        df2["Week"] = str(fallback_week)
        warnings.append(f"Filled missing Week with sidebar value: {fallback_week}")
    if ("Opponent" not in df2.columns or df2["Opponent"].isna().all()) and force_fill_if_missing and fallback_opp:
        df2["Opponent"] = str(fallback_opp).upper()
        warnings.append(f"Filled missing Opponent with sidebar value: {fallback_opp}")

    # Final validation
    if "Week" not in df2.columns or df2["Week"].isna().all():
        raise ValueError("Missing or unreadable 'Week'. Include numeric week or values like W01, W1, or enable 'Auto-fill Week/Opponent' in sidebar.")
    if "Opponent" not in df2.columns or df2["Opponent"].isna().all():
        raise ValueError("Missing or unreadable 'Opponent'. Use team code (e.g., MIN) or full team name, or enable 'Auto-fill Week/Opponent' in sidebar.")

    df2["Week"] = df2["Week"].astype(str)
    df2["Opponent"] = df2["Opponent"].astype(str).str.upper()

    return df2, warnings

# ================== Sidebar: controls & auto fetch ==================
with st.sidebar:
    st.header("Weekly Controls")
    week = st.selectbox("Week", week_options(), index=0)
    opponent = st.selectbox("Opponent", NFL_CODES, index=NFL_CODES.index("MIN") if "MIN" in NFL_CODES else 0)

    st.markdown("---")
    st.subheader("Import Options")
    fill_missing = st.checkbox("Auto-fill missing Week/Opponent from sidebar", value=True)

    st.markdown("---")
    st.subheader("Auto-Fetch (via nfl_data_py)")
    colA, colB = st.columns(2)
    with colA:
        auto_nfl_btn = st.button("Fetch NFL Data (Auto)")
    with colB:
        auto_snaps_btn = st.button("Fetch Snap Counts (Auto)")
    if not NFL_OK:
        st.caption("Install/keep nfl_data_py to enable auto fetch.")

    st.markdown("---")
    st.subheader("Current Week Download")
    btn_xlsx = st.button("Download Current Week (Excel)")
    btn_pdf  = st.button("Download Current Week (PDF)")

    st.markdown("---")
    st.subheader("Master Workbook")
    ensure_master()
    with open(MASTER_XLSX, "rb") as f:
        st.download_button(
            "Download Master Excel",
            data=f.read(),
            file_name="bears_weekly_analytics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ================== Auto-Fetch implementations ==================
def _guess_season() -> int:
    # Simple guess: if month >= August, use current year; otherwise previous.
    today = datetime.now()
    return today.year if today.month >= 8 else (today.year - 1)

def fetch_nfl_data_auto(week_code: str, opp_code: str):
    """Best-effort team offense/defense snapshot from nfl_data_py weekly data."""
    if not NFL_OK:
        st.error("nfl_data_py not available. Please install/update it or upload CSVs manually.")
        return

    season = _guess_season()
    try:
        wnum = int(str(week_code).lstrip("Ww"))
    except Exception:
        st.error(f"Could not parse week: {week_code}")
        return

    weekly = None
    if hasattr(nfl, "import_weekly_data"):
        try:
            weekly = nfl.import_weekly_data([season])
        except Exception as e:
            st.warning(f"import_weekly_data failed ({e}).")
    if weekly is None or weekly.empty:
        st.error("NFL weekly dataset unavailable from nfl_data_py. Please update the library or upload offense/defense CSVs.")
        return

    # Normalize team column name
    tcol = None
    for c in weekly.columns:
        if str(c).lower() in ("team", "recent_team", "recent_team_abbr"):
            tcol = c; break
    if tcol is None:
        st.error("Could not find team column in weekly data.")
        return

    wcol = None
    for c in weekly.columns:
        if str(c).lower() == "week":
            wcol = c; break
    if wcol is None:
        st.error("Could not find week column in weekly data.")
        return

    wk_df = weekly[weekly[wcol] == wnum].copy()
    if wk_df.empty:
        st.warning(f"No weekly rows for week {wnum}.")
        return

    # Build quick offense aggregates for CHI and OPP
    def agg_off(team_code):
        df = wk_df[wk_df[tcol].astype(str).str.upper() == team_code]
        if df.empty:
            return None
        out = {}
        for field, alias in [
            ("passing_yards", "PassYds"),
            ("rushing_yards", "RushYds"),
            ("receiving_yards", "RecvYds"),
            ("passing_tds", "PassTD"),
            ("rushing_tds", "RushTD"),
            ("targets", "Targets"),
            ("receptions", "Recs"),
            ("sacks", "SacksAgainst"),    # if present at player level (QB)
        ]:
            if field in df.columns:
                out[alias] = pd.to_numeric(df[field], errors="coerce").fillna(0).sum()
        return pd.DataFrame([{
            "Week": week_code,
            "Opponent": opp_code if team_code == "CHI" else "CHI",
            "Team": team_code,
            **out
        }])

    off_rows = []
    for team in ("CHI", opp_code):
        r = agg_off(team)
        if r is not None:
            off_rows.append(r)
    if off_rows:
        off_df = pd.concat(off_rows, ignore_index=True)
        append_to_sheet(off_df, SHEETS["offense"], dedup_cols=["Week","Opponent","Team"])
        st.success(f"Fetched NFL team offense snapshot for {week_code} (rows: {len(off_df)}).")
    else:
        st.warning("No offense rows aggregated for this week/teams.")

    # Very light defense (if tackles/sacks columns exist)
    def agg_def(team_code):
        df = wk_df[wk_df[tcol].astype(str).str.upper() == team_code]
        if df.empty:
            return None
        out = {}
        for field, alias in [
            ("solo_tackles", "SoloTk"),
            ("assisted_tackles", "AstTk"),
            ("sacks", "Sacks"),
            ("interceptions", "INT"),
            ("forced_fumbles", "FF"),
            ("fumbles_recovered", "FR"),
            ("passes_defended", "PD"),
        ]:
            if field in df.columns:
                out[alias] = pd.to_numeric(df[field], errors="coerce").fillna(0).sum()
        return pd.DataFrame([{
            "Week": week_code,
            "Opponent": opp_code if team_code == "CHI" else "CHI",
            "Team": team_code,
            **out
        }])

    def_rows = []
    for team in ("CHI", opp_code):
        r = agg_def(team)
        if r is not None:
            def_rows.append(r)
    if def_rows:
        def_df = pd.concat(def_rows, ignore_index=True)
        append_to_sheet(def_df, SHEETS["defense"], dedup_cols=["Week","Opponent","Team"])
        st.success(f"Fetched NFL team defense snapshot for {week_code} (rows: {len(def_df)}).")
    else:
        st.info("Defense stats not available in your nfl_data_py weekly schema; upload defense CSV if needed.")

def fetch_snap_counts_auto(week_code: str, opp_code: str):
    """Best-effort snap count import; maps to (Week, Opponent, Player, Snaps, Snap%, Side)."""
    if not NFL_OK:
        st.error("nfl_data_py not available. Please install/update it or upload Snap Counts CSV.")
        return

    season = _guess_season()
    try:
        wnum = int(str(week_code).lstrip("Ww"))
    except Exception:
        st.error(f"Could not parse week: {week_code}")
        return

    snap_df = None
    for fname in ("import_snap_counts", "import_weekly_snap_counts", "load_snap_counts"):
        f = getattr(nfl, fname, None)
        if callable(f):
            try:
                snap_df = f([season])
                break
            except Exception as e:
                continue

    if snap_df is None or snap_df.empty:
        st.error("No snap counts function/data found in nfl_data_py. Update the lib or upload snap CSV.")
        return

    # Find columns (case-insensitive)
    def find_col(dframe, options):
        for c in dframe.columns:
            if str(c).lower() in [o.lower() for o in options]:
                return c
        return None

    tcol = find_col(snap_df, ["team","recent_team","recent_team_abbr"])
    ocol = find_col(snap_df, ["opponent","opp","opponent_team","opponent_abbr"])
    wcol = find_col(snap_df, ["week","wk"])
    pcol = find_col(snap_df, ["player","player_name","full_name"])
    snaps_cols = [c for c in snap_df.columns if str(c).lower() in
                  ("snaps","offense_snaps","defense_snaps","st_snaps","total_snaps")]
    pct_off = find_col(snap_df, ["offense_pct","offense_share","off_pct"])
    pct_def = find_col(snap_df, ["defense_pct","defense_share","def_pct"])
    pct_st  = find_col(snap_df, ["special_teams_pct","st_pct","special_pct"])

    if not (tcol and wcol and pcol):
        st.error("Unexpected snap counts schema; please upload CSV instead.")
        return

    filt = snap_df[snap_df[wcol] == wnum].copy()
    filt = filt[filt[tcol].astype(str).str.upper().isin(["CHI", opp_code])]

    # Opponent
    if ocol:
        filt["Opponent"] = filt[ocol].astype(str).str.upper()
    else:
        filt["Opponent"] = np.where(filt[tcol].astype(str).str.upper() == "CHI", opp_code, "CHI")

    # Side (choose the highest % among Off/Def/ST if present)
    def pick_side(row):
        triples = []
        if pct_off and pd.notna(row.get(pct_off)): triples.append(("Offense", row[pct_off]))
        if pct_def and pd.notna(row.get(pct_def)): triples.append(("Defense", row[pct_def]))
        if pct_st  and pd.notna(row.get(pct_st)):  triples.append(("ST", row[pct_st]))
        if not triples:
            return ""
        return max(triples, key=lambda x: (0 if pd.isna(x[1]) else float(x[1])))[0]

    side_series = filt.apply(pick_side, axis=1)

    # Snaps (choose best available)
    snaps_col = snaps_cols[0] if snaps_cols else None
    if not snaps_col:
        filt["snaps_proxy"] = 0
        snaps_col = "snaps_proxy"

    out = pd.DataFrame({
        "Week": [week_code] * len(filt),
        "Opponent": filt["Opponent"].astype(str).str.upper(),
        "Player": filt[pcol].astype(str),
        "Snaps": pd.to_numeric(filt[snaps_col], errors="coerce").fillna(0),
        "Snap%": pd.to_numeric(
            filt.get(pct_off, pd.Series([np.nan]*len(filt)))
        , errors="coerce").fillna(
            pd.to_numeric(filt.get(pct_def, pd.Series([np.nan]*len(filt))), errors="coerce")
        ).fillna(
            pd.to_numeric(filt.get(pct_st, pd.Series([np.nan]*len(filt))), errors="coerce")
        ),
        "Side": side_series
    })

    append_to_sheet(out, SHEETS["snap"], dedup_cols=["Week","Opponent","Player"])
    st.success(f"Fetched snap counts for {week_code} vs {opp_code}: {len(out)} rows.")

# Trigger auto-fetch if buttons pressed
if auto_nfl_btn:
    fetch_nfl_data_auto(week, opponent)
if auto_snaps_btn:
    fetch_snap_counts_auto(week, opponent)

# ================== 1) Weekly Notes ==================
st.markdown("### 1) Weekly Notes")
with st.expander("Notes"):
    note_text = st.text_area("Notes", height=110, placeholder="Key matchups, weather, personnel notes…")
    if st.button("Save Notes"):
        row = pd.DataFrame([{"Week": week, "Opponent": opponent, "Notes": note_text.strip()}])
        append_to_sheet(row, SHEETS["notes"], dedup_cols=["Week","Opponent"])
        st.success("Notes saved.")

# ================== 2) Upload Weekly Data ==================
st.markdown("### 2) Upload Weekly Data")
c1, c2 = st.columns(2)

with c1:
    st.subheader("Offense (.csv/.cvs/.xlsx)")
    st.caption("Headers can vary; app detects Week/Opponent. Example: Week,Opponent,Points,Yards,YPA,CMP% …")
    f = st.file_uploader("Upload Offense", type=["csv","cvs","xlsx","xls"], key="off_upl")
    if f:
        try:
            df = read_any_table(f)
            try:
                df, warns = ensure_week_opp(df, fallback_week=week, fallback_opp=opponent, force_fill_if_missing=fill_missing)
            except Exception as e:
                if fill_missing:
                    df["Week"] = week
                    df["Opponent"] = opponent
                    warns = [f"Filled Week/Opponent from sidebar ({week}, {opponent})."]
                else:
                    raise
            for w in warns: st.info(w)
            append_to_sheet(df, SHEETS["offense"], dedup_cols=["Week","Opponent"])
            st.success(f"Offense rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Offense upload failed: {e}")

    st.subheader("Defense (.csv/.cvs/.xlsx)")
    st.caption("Example: Week,Opponent,SACK,INT,3D%_Allowed,RZ%_Allowed,Pressures …")
    f = st.file_uploader("Upload Defense", type=["csv","cvs","xlsx","xls"], key="def_upl")
    if f:
        try:
            df = read_any_table(f)
            try:
                df, warns = ensure_week_opp(df, fallback_week=week, fallback_opp=opponent, force_fill_if_missing=fill_missing)
            except Exception as e:
                if fill_missing:
                    df["Week"] = week
                    df["Opponent"] = opponent
                    warns = [f"Filled Week/Opponent from sidebar ({week}, {opponent})."]
                else:
                    raise
            for w in warns: st.info(w)
            append_to_sheet(df, SHEETS["defense"], dedup_cols=["Week","Opponent"])
            st.success(f"Defense rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Defense upload failed: {e}")

with c2:
    st.subheader("Personnel (.csv/.cvs/.xlsx)")
    st.info(
        "Personnel: Week,Opponent,11,12,13,21,Other | "
        "Snap_Counts: Week,Opponent,Player,Snaps,Snap%,Side."
    )
    f = st.file_uploader("Upload Personnel", type=["csv","cvs","xlsx","xls"], key="per_upl")
    if f:
        try:
            df = read_any_table(f)
            try:
                df, warns = ensure_week_opp(df, fallback_week=week, fallback_opp=opponent, force_fill_if_missing=fill_missing)
            except Exception as e:
                if fill_missing:
                    df["Week"] = week
                    df["Opponent"] = opponent
                    warns = [f"Filled Week/Opponent from sidebar ({week}, {opponent})."]
                else:
                    raise
            for w in warns: st.info(w)
            append_to_sheet(df, SHEETS["personnel"], dedup_cols=["Week","Opponent"])
            st.success(f"Personnel rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Personnel upload failed: {e}")

    st.subheader("Snap Counts (.csv/.cvs/.xlsx)")
    st.caption("Columns include: Week,Opponent,Player,Snaps,Snap%,Side")
    f = st.file_uploader("Upload Snap Counts", type=["csv","cvs","xlsx","xls"], key="snap_upl")
    if f:
        try:
            df = read_any_table(f)
            try:
                df, warns = ensure_week_opp(df, fallback_week=week, fallback_opp=opponent, force_fill_if_missing=fill_missing)
            except Exception as e:
                if fill_missing:
                    df["Week"] = week
                    df["Opponent"] = opponent
                    warns = [f"Filled Week/Opponent from sidebar ({week}, {opponent})."]
                else:
                    raise
            for w in warns: st.info(w)
            if "Player" in df.columns:
                df["Player"] = df["Player"].astype(str).str.strip()
            append_to_sheet(df, SHEETS["snap"], dedup_cols=["Week","Opponent","Player"])
            st.success(f"Snap Count rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Snap upload failed: {e}")

# ================== 3) Opponent / Injuries / Media / Strategy ==================
st.markdown("### 3) Opponent Preview / Injuries / Media / Strategy")
a, b, c, d = st.columns(4)

with a:
    st.subheader("Opponent Preview")
    opp_off = st.text_area("Offense Summary", height=110)
    opp_def = st.text_area("Defense Summary", height=110)
    opp_match = st.text_area("Key Matchups", height=110)
    if st.button("Save Opponent Preview"):
        row = pd.DataFrame([{
            "Week": week, "Opponent": opponent,
            "Off_Summary": opp_off.strip(),
            "Def_Summary": opp_def.strip(),
            "Matchups": opp_match.strip()
        }])
        append_to_sheet(row, SHEETS["opp"], dedup_cols=["Week","Opponent"])
        st.success("Opponent preview saved.")

with b:
    st.subheader("Injuries")
    p = st.text_input("Player")
    s = st.selectbox("Status", ["Questionable","Doubtful","Out","IR","Healthy"])
    bp = st.text_input("Body Part / Injury")
    prac = st.text_input("Practice (DNP/Limited/Full)")
    gstat = st.text_input("Game Status")
    notes = st.text_area("Notes", height=90)
    if st.button("Save Injury"):
        if p.strip():
            row = pd.DataFrame([{
                "Week": week, "Opponent": opponent,
                "Player": p.strip(), "Status": s, "BodyPart": bp.strip(),
                "Practice": prac.strip(), "GameStatus": gstat.strip(), "Notes": notes.strip()
            }])
            append_to_sheet(row, SHEETS["inj"], dedup_cols=["Week","Opponent","Player"])
            st.success("Injury saved.")
        else:
            st.error("Enter a player name.")

with c:
    st.subheader("Media Summaries")
    src = st.text_input("Source")
    summ = st.text_area("Summary", height=140)
    if st.button("Save Media Summary"):
        if src.strip() and summ.strip():
            row = pd.DataFrame([{
                "Week": week, "Opponent": opponent,
                "Source": src.strip(), "Summary": summ.strip()
            }])
            append_to_sheet(row, SHEETS["media"])
            st.success("Media summary saved.")
        else:
            st.error("Enter both Source and Summary.")

with d:
    st.subheader("Weekly Strategy")
    plan = st.text_area("Plan (headline)", height=70)
    keys = st.text_area("Keys (bullets ok)", height=70)
    strat_notes = st.text_area("Notes", height=70)
    if st.button("Save Strategy"):
        row = pd.DataFrame([{
            "Week": week, "Opponent": opponent,
            "Plan": plan.strip(), "Keys": keys.strip(), "Notes": strat_notes.strip()
        }])
        append_to_sheet(row, SHEETS["strategy"], dedup_cols=["Week","Opponent"])
        st.success("Strategy saved.")

# ================== 4) Prediction ==================
st.markdown("### 4) Prediction")
x, y = st.columns(2)
with x:
    who = st.selectbox("Predicted Winner", ["CHI","OPP"], index=0)
    conf = st.slider("Confidence", 0, 100, 60)
with y:
    why = st.text_area("Rationale", height=120)
if st.button("Save Prediction"):
    row = pd.DataFrame([{
        "Week": week, "Opponent": opponent,
        "Predicted_Winner": "CHI" if who == "CHI" else opponent,
        "Confidence": conf,
        "Rationale": why.strip()
    }])
    append_to_sheet(row, SHEETS["pred"], dedup_cols=["Week","Opponent"])
    st.success("Prediction saved.")

# ================== Current-week package & downloads ==================
def build_week_package(week_code: str, opp_code: str) -> dict:
    pkg = {}
    for sheet_name in SHEETS.values():
        df = read_sheet(sheet_name)
        if df.empty:
            pkg[sheet_name] = df
            continue
        if {"Week","Opponent"}.issubset(df.columns):
            m = (df["Week"].astype(str) == week_code) & (df["Opponent"].astype(str).str.upper() == opp_code)
            pkg[sheet_name] = df[m].copy()
        else:
            pkg[sheet_name] = df.copy()
    return pkg

def export_week_excel(pkg: dict, path: Path):
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as w:
        for sheet_name, df in pkg.items():
            safe = sheet_name[:31]
            (df if not df.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name=safe)

def export_week_pdf(pkg: dict, path: Path, title: str):
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=15)
    def H(t): pdf.set_font("Arial","B",14); pdf.cell(0,10,t,ln=True)
    def P(t): pdf.set_font("Arial","",11); pdf.multi_cell(0,6,t if t else "-")

    pdf.add_page()
    pdf.set_font("Arial","B",16); pdf.cell(0,10,title,ln=True); pdf.ln(4)

    H("Weekly Notes")
    n = pkg.get(SHEETS["notes"], pd.DataFrame())
    P(str(n["Notes"].iloc[0]) if not n.empty and "Notes" in n.columns else "")

    pdf.ln(3); H("Opponent Preview")
    o = pkg.get(SHEETS["opp"], pd.DataFrame())
    if not o.empty:
        P("Offense: " + str(o.get("Off_Summary", pd.Series([""])).iloc[0]))
        P("Defense: " + str(o.get("Def_Summary", pd.Series([""])).iloc[0]))
        P("Matchups: " + str(o.get("Matchups", pd.Series([""])).iloc[0]))
    else:
        P("")

    pdf.ln(3); H("Weekly Strategy")
    s = pkg.get(SHEETS["strategy"], pd.DataFrame())
    if not s.empty:
        P("Plan: " + str(s.get("Plan", pd.Series([""])).iloc[0]))
        P("Keys: " + str(s.get("Keys", pd.Series([""])).iloc[0]))
        P("Notes: " + str(s.get("Notes", pd.Series([""])).iloc[0]))
    else:
        P("")

    pdf.ln(3); H("Injuries")
    inj = pkg.get(SHEETS["inj"], pd.DataFrame())
    if not inj.empty:
        for _, r in inj.iterrows():
            P(f"{r.get('Player','')} — {r.get('Status','')} — {r.get('BodyPart','')}; "
              f"Practice: {r.get('Practice','')}; Game: {r.get('GameStatus','')}; "
              f"Notes: {r.get('Notes','')}")
    else:
        P("")

    pdf.ln(3); H("Media Summaries")
    med = pkg.get(SHEETS["media"], pd.DataFrame())
    if not med.empty:
        for _, r in med.iterrows():
            P(f"{r.get('Source','')}: {r.get('Summary','')}")
    else:
        P("")

    pdf.ln(3); H("Prediction")
    pr = pkg.get(SHEETS["pred"], pd.DataFrame())
    if not pr.empty:
        r = pr.iloc[-1]
        P(f"Predicted Winner: {r.get('Predicted_Winner','')}")
        P(f"Confidence: {r.get('Confidence','')}")
        P("Rationale: " + str(r.get("Rationale","")))
    else:
        P("")

    pdf.ln(6); pdf.set_font("Arial","I",9); pdf.cell(0,6,"Generated by Bears Weekly Tracker",ln=True)
    pdf.output(str(path))

# Buttons for current-week downloads
if btn_xlsx or btn_pdf:
    pkg = build_week_package(week, opponent)
    if btn_xlsx:
        x = CURR_DIR / f"{week}_{opponent}.xlsx"
        export_week_excel(pkg, x)
        st.success(f"Saved: {x}")
        with open(x, "rb") as f:
            st.download_button("Download Current Week Excel (click to save)", f.read(), file_name=x.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if btn_pdf:
        p = CURR_DIR / f"{week}_{opponent}.pdf"
        export_week_pdf(pkg, p, title=f"Weekly Report — {week} vs {opponent}")
        st.success(f"Saved: {p}")
        with open(p, "rb") as f:
            st.download_button("Download Current Week PDF (click to save)", f.read(), file_name=p.name, mime="application/pdf")

# ================== 5) Search & Snap Finder ==================
st.markdown("### 5) Search & Snap Finder")
s1, s2 = st.columns([2,1])
with s1:
    st.subheader("Global Search (choose a table)")
    table_choice = st.selectbox("Table", list(SHEETS.values()), index=list(SHEETS.values()).index(SHEETS["snap"]))
    q = st.text_input("Search text (matches anywhere in the selected table)", "")
    df = read_sheet(table_choice)
    if not df.empty and q.strip():
        qlow = q.strip().lower()
        mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(qlow, na=False))
        df = df[mask.any(axis=1)]
    st.dataframe(df if not df.empty else pd.DataFrame({"Info":["No rows (or no matches)."]}),
                 use_container_width=True, hide_index=True)

with s2:
    st.subheader("Search Snap Counts")
    snap_df = read_sheet(SHEETS["snap"])
    player = st.text_input("Player contains…", "")
    side = st.selectbox("Side", ["(any)","Offense","Defense","ST"], index=0)
    snap_week = st.selectbox("Snap Week", ["(any)"] + week_options(), index=0)
    if not snap_df.empty:
        filt = snap_df.copy()
        if player.strip():
            filt = filt[filt["Player"].astype(str).str.contains(player.strip(), case=False, na=False)]
        if side != "(any)" and "Side" in filt.columns:
            filt = filt[filt["Side"].astype(str).str.contains(side, case=False, na=False)]
        if snap_week != "(any)" and "Week" in filt.columns:
            filt = filt[filt["Week"].astype(str) == snap_week]
        st.dataframe(filt if not filt.empty else pd.DataFrame({"Info":["No snap rows match."]}),
                     use_container_width=True, hide_index=True)
    else:
        st.info("No snap data yet. Use auto-fetch or upload.")

# ================== 6) Current Week Snapshots (with Search) ==================
st.markdown("### 6) Current Week Snapshots (with Search)")
tabs = st.tabs(["Offense","Defense","Personnel","Snap Counts","Injuries","Media","Opponent","Prediction","Notes","Strategy"])
search_terms = []

def show_df(tab, sheet):
    with tab:
        colX, colY = st.columns([2,1])
        with colX:
            q = st.text_input(f"Search {sheet}", key=f"q_{sheet}", placeholder="Search…")
        with colY:
            st.caption("Filters matching rows in this tab.")
        df = read_sheet(sheet)
        if not df.empty and {"Week","Opponent"}.issubset(df.columns):
            df = df[(df["Week"].astype(str) == week) & (df["Opponent"].astype(str).str.upper() == opponent)]
        if not df.empty and (q or "").strip():
            qlow = q.strip().lower()
            mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(qlow, na=False))
            df = df[mask.any(axis=1)]
        st.dataframe(df if not df.empty else pd.DataFrame({"Info": ["No rows for this week/opponent (or search filtered all)."]}),
                     use_container_width=True, hide_index=True)

show_df(tabs[0], SHEETS["offense"])
show_df(tabs[1], SHEETS["defense"])
show_df(tabs[2], SHEETS["personnel"])
show_df(tabs[3], SHEETS["snap"])
show_df(tabs[4], SHEETS["inj"])
show_df(tabs[5], SHEETS["media"])
show_df(tabs[6], SHEETS["opp"])
show_df(tabs[7], SHEETS["pred"])
show_df(tabs[8], SHEETS["notes"])
show_df(tabs[9], SHEETS["strategy"])

st.caption("Master workbook lives in ./data. Current-week downloads save to ./exports/Current_Week. Uploads accept CSV or Excel; .cvs is accepted. Week/Opponent are auto-filled from the sidebar when enabled.")

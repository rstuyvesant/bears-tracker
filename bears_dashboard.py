# 🐻 Chicago Bears 2025–26 Weekly Tracker
# Inline Weekly Controls + Auto Snap Counts + Extra Analytics

import os
import math
import pandas as pd
import streamlit as st
from datetime import datetime, date
try:
    from zoneinfo import ZoneInfo  # Py>=3.9
except Exception:
    ZoneInfo = None

# Optional deps
try:
    import openpyxl
except Exception:
    openpyxl = None
try:
    from fpdf import FPDF
except Exception:
    FPDF = None
try:
    import nfl_data_py as nfl
except Exception:
    nfl = None

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.caption("Inline Weekly Controls + Auto Snap Counts. Use 3-letter opponent codes (e.g., MIN). Filenames are case-sensitive on Linux/Streamlit Cloud.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"
SEASON_DEFAULT = 2025

# === Auto-calc current NFL week (approx) ===
SEASON_START_MONTH = 9   # adjust if needed
SEASON_START_DAY = 4     # approx first Thursday
def get_current_week(season: int = SEASON_DEFAULT) -> int:
    try:
        tz = ZoneInfo("America/Chicago") if ZoneInfo else None
    except Exception:
        tz = None
    today = (datetime.now(tz).date() if tz else datetime.now().date())
    start = date(season, SEASON_START_MONTH, SEASON_START_DAY)
    if today < start:
        return 1
    delta_weeks = (today - start).days // 7 + 1
    return max(1, min(30, delta_weeks))

CURRENT_WEEK = get_current_week(SEASON_DEFAULT)

# =========================
# Opponent normalization
# =========================
TEAM_MAP = {
    "chicago":"CHI","bears":"CHI","chi":"CHI",
    "detroit":"DET","lions":"DET","det":"DET",
    "green bay":"GB","packers":"GB","gb":"GB",
    "minnesota":"MIN","vikings":"MIN","minn":"MIN","min":"MIN",
    "dallas":"DAL","cowboys":"DAL","dal":"DAL",
    "new york giants":"NYG","giants":"NYG","nyg":"NYG",
    "philadelphia":"PHI","eagles":"PHI","phi":"PHI",
    "washington":"WAS","commanders":"WAS","was":"WAS","wsh":"WAS",
    "atlanta":"ATL","falcons":"ATL","atl":"ATL",
    "carolina":"CAR","panthers":"CAR","car":"CAR",
    "new orleans":"NO","saints":"NO","no":"NO",
    "tampa bay":"TB","buccaneers":"TB","bucs":"TB","tb":"TB",
    "arizona":"ARI","cardinals":"ARI","ari":"ARI",
    "los angeles rams":"LAR","rams":"LAR","lar":"LAR",
    "san francisco":"SF","49ers":"SF","niners":"SF","sf":"SF",
    "seattle":"SEA","seahawks":"SEA","sea":"SEA",
    "baltimore":"BAL","ravens":"BAL","bal":"BAL",
    "cincinnati":"CIN","bengals":"CIN","cin":"CIN",
    "cleveland":"CLE","browns":"CLE","cle":"CLE",
    "pittsburgh":"PIT","steelers":"PIT","pit":"PIT",
    "buffalo":"BUF","bills":"BUF","buf":"BUF",
    "miami":"MIA","dolphins":"MIA","mia":"MIA",
    "new england":"NE","patriots":"NE","ne":"NE",
    "new york jets":"NYJ","jets":"NYJ","nyj":"NYJ",
    "houston":"HOU","texans":"HOU","hou":"HOU",
    "indianapolis":"IND","colts":"IND","ind":"IND",
    "jacksonville":"JAX","jaguars":"JAX","jax":"JAX",
    "tennessee":"TEN","titans":"TEN","ten":"TEN",
    "denver":"DEN","broncos":"DEN","den":"DEN",
    "kansas city":"KC","chiefs":"KC","kc":"KC",
    "las vegas":"LV","raiders":"LV","lv":"LV",
    "los angeles chargers":"LAC","chargers":"LAC","lac":"LAC"
}
TEAM_CODES = ["ARI","ATL","BAL","BUF","CAR","CHI","CIN","CLE","DAL","DEN","DET","GB","HOU","IND","JAX",
              "KC","LAC","LAR","LV","MIA","MIN","NE","NO","NYG","NYJ","PHI","PIT","SEA","SF","TB","TEN","WAS"]

def canon_team(x: str) -> str:
    x = (x or "").strip()
    if not x:
        return x
    return TEAM_MAP.get(x.lower(), x.upper())

# =========================
# Excel helpers
# =========================
def _ensure_openpyxl():
    if openpyxl is None:
        st.error("openpyxl is required for Excel features. Add it to requirements.txt")
        st.stop()

def append_to_excel(new_df: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, dedup_keys=None):
    _ensure_openpyxl()
    new_df = new_df.copy()
    if "Opponent" in new_df.columns:
        new_df["Opponent"] = new_df["Opponent"].apply(canon_team)
    if "Week" in new_df.columns:
        new_df["Week"] = pd.to_numeric(new_df["Week"], errors="coerce").astype("Int64")

    if os.path.exists(file_name):
        book = openpyxl.load_workbook(file_name)
    else:
        book = openpyxl.Workbook()
        if "Sheet" in book.sheetnames:
            book.remove(book["Sheet"])

    if sheet_name in book.sheetnames:
        sheet = book[sheet_name]
        existing = pd.DataFrame(sheet.values)
        if not existing.empty:
            existing.columns = existing.iloc[0]
            existing = existing[1:]
        else:
            existing = pd.DataFrame(columns=new_df.columns)
        all_cols = list(dict.fromkeys(list(existing.columns) + list(new_df.columns)))
        existing = existing.reindex(columns=all_cols)
        new_df = new_df.reindex(columns=all_cols)
        combined = pd.concat([existing, new_df], ignore_index=True)
        if dedup_keys:
            combined = combined.drop_duplicates(subset=dedup_keys, keep="last")
        book.remove(sheet)
        sheet = book.create_sheet(sheet_name)
        sheet.append(list(combined.columns))
        for _, row in combined.iterrows():
            sheet.append(list(row.values))
    else:
        sheet = book.create_sheet(sheet_name)
        sheet.append(list(new_df.columns))
        for _, row in new_df.iterrows():
            sheet.append(list(row.values))

    book.save(file_name)

def read_sheet(sheet_name: str, file_name: str = EXCEL_FILE) -> pd.DataFrame:
    _ensure_openpyxl()
    if not os.path.exists(file_name):
        return pd.DataFrame()
    book = openpyxl.load_workbook(file_name)
    if sheet_name not in book.sheetnames:
        return pd.DataFrame()
    sheet = book[sheet_name]
    df = pd.DataFrame(sheet.values)
    if df.empty:
        return pd.DataFrame()
    df.columns = df.iloc[0]
    df = df[1:]
    return df

# =========================
# NFL averages fetch (extra stats aware)
# =========================
def fetch_nfl_averages_weekly():
    if nfl is None:
        st.warning("nfl_data_py not installed; skipping auto NFL averages fetch.")
        return
    try:
        season = SEASON_DEFAULT
        df = None
        for fn_name in ["import_weekly_team_stats", "import_seasonal_team_stats", "import_team_game_logs"]:
            fn = getattr(nfl, fn_name, None)
            if fn is None:
                continue
            try:
                tmp = fn([season])
                if isinstance(tmp, pd.DataFrame) and not tmp.empty:
                    df = tmp.copy(); break
            except Exception:
                continue
        if df is None or df.empty:
            st.warning("Could not fetch weekly team stats.")
            return

        # Map various library column names into ours
        colmap = {}
        for col in df.columns:
            lc = str(col).lower()
            if "penalt" in lc and "yards" in lc: colmap[col] = "Penalty_Yards"
            elif "penalt" in lc: colmap[col] = "Penalties"
            elif "yac" in lc or "yards_after_catch" in lc: colmap[col] = "YAC"
            elif "yac_allowed" in lc: colmap[col] = "YAC_Allowed"
            elif lc in ("team","team_abbr","team_code","club_code"): colmap[col] = "Team"
            elif lc in ("week","game_week"): colmap[col] = "Week"
            elif lc in ("season","year"): colmap[col] = "Season"
        if colmap:
            df = df.rename(columns=colmap)

        keep = [c for c in ["Season","Week","Team","Penalties","Penalty_Yards","YAC","YAC_Allowed"] if c in df.columns]
        if not keep:
            return
        small = df[keep].copy()

        if "Week" in small.columns:
            wk = small.groupby("Week").mean(numeric_only=True).reset_index()
            wk.insert(0, "Season", season)
            append_to_excel(wk, "NFL_Averages_Weekly", dedup_keys=["Season","Week"])

        ytd = small.mean(numeric_only=True).to_frame().T
        ytd.insert(0, "Season", season)
        append_to_excel(ytd, "NFL_Averages_YTD", dedup_keys=["Season"])

        st.success("NFL averages updated (weekly + YTD).")
    except Exception as e:
        st.warning(f"NFL averages fetch failed: {e}")

# =========================
# Auto Snap Counts (sidebar)
# =========================
def fetch_snap_counts_auto(team_code: str, season: int, week: int):
    """Fetch snap counts for a given team/week and append to Snap_Counts sheet.
       Saves rows: [Week, Opponent, Player, Snaps, Snap%, Side]. Opponent left blank if not derivable."""
    if nfl is None:
        st.warning("nfl_data_py not installed.")
        return 0

    df = None
    for fn_name in ["import_weekly_snap_counts", "import_snap_counts", "import_pbp"]:
        fn = getattr(nfl, fn_name, None)
        if fn is None:
            continue
        try:
            tmp = fn([season]) if fn_name != "import_pbp" else fn(seasons=[season])
            if isinstance(tmp, pd.DataFrame) and not tmp.empty:
                df = tmp.copy(); break
        except Exception:
            continue
    if df is None or df.empty:
        return 0

    # Normalize columns across library versions
    colmap = {}
    for col in df.columns:
        lc = str(col).lower()
        if lc in ("team","recent_team","club_code","team_abbr"): colmap[col] = "Team"
        elif lc in ("week","game_week"): colmap[col] = "Week"
        elif lc in ("player","player_name","name"): colmap[col] = "Player"
        elif lc.startswith("offense_snap"): colmap[col] = "OFF_Snaps"
        elif lc.startswith("offense_pct"): colmap[col] = "OFF_Pct"
        elif lc.startswith("defense_snap"): colmap[col] = "DEF_Snaps"
        elif lc.startswith("defense_pct"): colmap[col] = "DEF_Pct"
    if colmap:
        df = df.rename(columns=colmap)

    # Filter for team/week
    try:
        df = df[(df["Team"].astype(str).str.upper()==team_code) & (df["Week"].astype(str)==str(week))]
    except Exception:
        return 0
    if df.empty:
        return 0

    rows = []
    for _, r in df.iterrows():
        player = r.get("Player", "")
        if pd.notna(r.get("OFF_Snaps")):
            rows.append([week, "", player, r.get("OFF_Snaps"), r.get("OFF_Pct"), "OFF"])
        if pd.notna(r.get("DEF_Snaps")):
            rows.append([week, "", player, r.get("DEF_Snaps"), r.get("DEF_Pct"), "DEF"])

    if not rows:
        return 0
    out_df = pd.DataFrame(rows, columns=["Week","Opponent","Player","Snaps","Snap%","Side"])
    append_to_excel(out_df, "Snap_Counts", dedup_keys=["Week","Opponent","Player","Side"])
    return len(out_df)

# =========================
# DVOA-Proxy (with extra stats)
# =========================
def safe_num(x):
    try: return float(x)
    except Exception: return math.nan

def compute_dvoa_proxy(off_df: pd.DataFrame, def_df: pd.DataFrame) -> pd.DataFrame:
    if off_df is None: off_df = pd.DataFrame()
    if def_df is None: def_df = pd.DataFrame()

    # Best-effort merge on Week/Opponent
    if not off_df.empty and not def_df.empty and all(c in off_df.columns for c in ["Week","Opponent"]) and all(c in def_df.columns for c in ["Week","Opponent"]):
        base = pd.merge(off_df, def_df, on=["Week","Opponent"], how="outer", suffixes=("_OFF","_DEF"))
        # Map unified keys from OFF/DEF
        base["YPA"] = base.get("YPA_OFF")
        base["CMP%"] = base.get("CMP%_OFF")
        base["SR%"] = base.get("SR%_OFF")
        base["YAC"] = base.get("YAC_OFF")
        base["Penalties"] = base.get("Penalties_OFF")
        base["Penalty_Yards"] = base.get("Penalty_Yards_OFF")
        base["RZ%_Allowed"] = base.get("RZ%_Allowed_DEF")
        base["SACKs"] = base.get("SACKs_DEF")
        base["Pressures"] = base.get("Pressures_DEF")
        base["INTs"] = base.get("INTs_DEF")
        base["YAC_Allowed"] = base.get("YAC_Allowed_DEF")
    else:
        base = off_df.copy() if not off_df.empty else def_df.copy()

    if base.empty:
        return pd.DataFrame(columns=["Week","Opponent","DVOA_Proxy"])

    def score_row(row):
        ypa = safe_num(row.get("YPA", math.nan))
        cmpct = safe_num(row.get("CMP%", math.nan))
        sr = safe_num(row.get("SR%", math.nan))
        yac = safe_num(row.get("YAC", math.nan))
        pen = safe_num(row.get("Penalties", math.nan))
        pen_yards = safe_num(row.get("Penalty_Yards", math.nan))
        rz_allowed = safe_num(row.get("RZ%_Allowed", math.nan))
        sacks = safe_num(row.get("SACKs", math.nan))
        pressures = safe_num(row.get("Pressures", math.nan))
        ints = safe_num(row.get("INTs", math.nan))
        yac_allowed = safe_num(row.get("YAC_Allowed", math.nan))
        nz = lambda v: 0 if math.isnan(v) else v
        pos = 0.30*nz(ypa) + 0.20*nz(cmpct) + 0.20*nz(sr) + 0.10*nz(yac) + 0.10*nz(sacks) + 0.10*nz(pressures)
        neg = 0.25*nz(rz_allowed) + (-0.15)*nz(ints) + 0.15*nz(pen) + 0.10*nz(pen_yards) + 0.10*nz(yac_allowed)
        return pos - neg

    base["DVOA_Proxy"] = base.apply(score_row, axis=1)
    out = base.groupby([c for c in ["Week","Opponent"] if c in base.columns], dropna=False)["DVOA_Proxy"].mean().reset_index()
    return out

# =========================
# Sidebar controls
# =========================
with st.sidebar:
    st.header("Controls")
    if st.button("Fetch NFL Data (Auto)"):
        fetch_nfl_averages_weekly()

    st.divider()
    st.subheader("Auto Snap Counts")
    sc_team = st.selectbox("Team", TEAM_CODES, index=TEAM_CODES.index("CHI"))
    sc_week = st.number_input("Week", min_value=1, max_value=30, value=CURRENT_WEEK, step=1, key="snap_week")
    sc_season = st.number_input("Season", min_value=2000, max_value=2100, value=SEASON_DEFAULT, step=1, key="snap_season")
    if st.button("Fetch Snap Counts (Auto)"):
        added = fetch_snap_counts_auto(sc_team, int(sc_season), int(sc_week))
        if added:
            st.success(f"Added {added} Snap_Counts rows for {sc_team}, Week {int(sc_week)}.")

    st.divider()
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("📥 Download All Data (Excel)", f.read(), EXCEL_FILE,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================
# WEEKLY CONTROLS (restored)
# =========================
st.subheader("🗂️ Weekly Controls")
colA, colB = st.columns([1,2])
with colA:
    week = st.number_input("Week", min_value=1, max_value=30, value=CURRENT_WEEK, step=1, key="wk")
with colB:
    opp_input = st.text_input("Opponent (name or code)")
OPP = canon_team(opp_input)
st.caption("Tip: Use 3-letter codes (e.g., MIN). Names like 'Vikings' are auto-mapped.")

# =========================
# KEY NOTES → Strategy sheet
# =========================
st.subheader("📝 Key Notes")
with st.form("key_notes_form"):
    notes = st.text_area("Notes (quick bullets)")
    if st.form_submit_button("Save Key Notes"):
        row = pd.DataFrame([[week, OPP, "", "", notes, ""]],
                           columns=["Week","Opponent","Pre_Game_Strategy","Post_Game_Results","Key_Takeaways","Next_Week_Impact"])
        append_to_excel(row, "Strategy", dedup_keys=["Week","Opponent"])
        st.success("Key Notes saved to Strategy sheet.")

# =========================
# MEDIA SUMMARIES (inline)
# =========================
st.subheader("📚 Media Summaries")
with st.form("media_inline"):
    src = st.text_input("Source (e.g., ESPN, The Athletic)")
    summ = st.text_area("Summary")
    if st.form_submit_button("Save Media Summary"):
        row = pd.DataFrame([[week, OPP, src, summ]], columns=["Week","Opponent","Source","Summary"])
        append_to_excel(row, "Media_Summaries", dedup_keys=["Week","Opponent","Source","Summary"])
        st.success("Media summary saved.")

# =========================
# INJURIES (inline)
# =========================
st.subheader("🩹 Injuries")
with st.form("injuries_inline"):
    player = st.text_input("Player")
    status = st.selectbox("Status", ["","Questionable","Doubtful","Out","IR","Active"])
    body = st.text_input("Body Part / Injury")
    practice = st.selectbox("Practice", ["","DNP","Limited","Full"])
    game_stat = st.selectbox("Game Status", ["","Active","Inactive","TBD"])
    inj_notes = st.text_area("Notes")
    if st.form_submit_button("Save Injury"):
        row = pd.DataFrame([[week, OPP, player, status, body, practice, game_stat, inj_notes]],
                           columns=["Week","Opponent","Player","Status","BodyPart","Practice","GameStatus","Notes"])
        append_to_excel(row, "Injuries", dedup_keys=["Week","Opponent","Player"])
        st.success("Injury saved.")

# =========================
# OPPONENT PREVIEW (inline)
# =========================
st.subheader("🔎 Opponent Preview")
with st.form("opp_prev"):
    off_str = st.text_area("Offense – strengths/weaknesses")
    def_str = st.text_area("Defense – strengths/weaknesses")
    st_str = st.text_area("Special Teams – notes")
    xfac = st.text_area("X-Factors / Matchups")
    add_notes = st.text_area("Additional Notes")
    if st.form_submit_button("Save Opponent Preview"):
        row = pd.DataFrame([[week, OPP, off_str, def_str, st_str, xfac, add_notes]],
                           columns=["Week","Opponent","Off_Notes","Def_Notes","ST_Notes","X_Factors","Notes"])
        append_to_excel(row, "Opponent_Preview", dedup_keys=["Week","Opponent"])
        st.success("Opponent preview saved.")

# =========================
# PREDICTIONS (inline)
# =========================
st.subheader("🔮 Weekly Game Prediction")
with st.form("pred_inline"):
    pick = st.text_input("Predicted Winner (e.g., CHI)")
    conf = st.slider("Confidence (0–100)", 0, 100, 60)
    rationale = st.text_area("Rationale (tie to strategy/analytics)")
    if st.form_submit_button("Save Prediction"):
        row = pd.DataFrame([[week, OPP, (pick or "").upper(), conf, rationale]],
                           columns=["Week","Opponent","Predicted_Winner","Confidence","Rationale"])
        append_to_excel(row, "Predictions", dedup_keys=["Week","Opponent"])
        st.success("Prediction saved.")

# =========================
# (Optional) DATA UPLOAD – in an expander
# =========================
with st.expander("⬆️ Optional: Upload Offense/Defense/Personnel/Snap Counts CSVs"):
    st.markdown("Use these only if you prefer bulk uploads. Headers shown below.")
    def _up(lbl, sheet, req_cols, dedup):
        up = st.file_uploader(lbl, type=["csv","xlsx"], key=f"{sheet}_up")
        if up is not None:
            try:
                df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            except Exception as e:
                st.error(f"Failed to read: {e}")
                df = pd.DataFrame()
            if not df.empty:
                missing = [c for c in req_cols if c not in df.columns]
                if missing:
                    st.warning(f"Missing columns: {missing}. Saving what is present.")
                if "Opponent" in df.columns:
                    df["Opponent"] = df["Opponent"].apply(canon_team)
                if "Week" in df.columns:
                    df["Week"] = pd.to_numeric(df["Week"], errors="coerce").astype("Int64")
                append_to_excel(df, sheet, dedup_keys=dedup)
                st.success(f"Appended {len(df)} rows to {sheet}.")
    _up("Upload offense.csv", "Offense",
        ["Week","Opponent","YPA","CMP%","SR%","Points","Penalties","Penalty_Yards","YAC"],
        ["Week","Opponent"])
    _up("Upload defense.csv", "Defense",
        ["Week","Opponent","RZ%_Allowed","SACKs","Pressures","INTs","Penalties","Penalty_Yards","YAC_Allowed"],
        ["Week","Opponent"])
    _up("Upload personnel.csv", "Personnel",
        ["Week","Opponent","11","12","13","21","Other"],
        ["Week","Opponent"])
    _up("Upload snap_counts.csv", "Snap_Counts",
        ["Week","Opponent","Player","Snaps","Snap%","Side"],
        ["Week","Opponent","Player","Side"])

# =========================
# ANALYTICS & REPORTS
# =========================
st.divider()
st.subheader("📊 DVOA-Proxy (auto from Offense/Defense)")
try:
    off_df = read_sheet("Offense")
    def_df = read_sheet("Defense")
    if not off_df.empty or not def_df.empty:
        proxy = compute_dvoa_proxy(off_df, def_df)
        if not proxy.empty:
            append_to_excel(proxy, "DVOA_Proxy", dedup_keys=["Week","Opponent"])
            st.dataframe(proxy.sort_values(["Week","Opponent"]))
        else:
            st.info("DVOA-Proxy: no rows yet.")
    else:
        st.info("Upload or maintain Offense/Defense to compute DVOA-Proxy.")
except Exception as e:
    st.warning(f"DVOA-Proxy skipped: {e}")

st.subheader("🧾 Export Weekly Final PDF")
week_pdf = st.number_input("Week to export", min_value=1, max_value=30, value=CURRENT_WEEK, step=1)
if st.button("Export Final PDF"):
    if FPDF is None:
        st.error("Install fpdf to enable PDF exports (add 'fpdf' to requirements.txt)")
    else:
        def _wk(df, wk):
            if df is None or df.empty or "Week" not in df.columns: return pd.DataFrame()
            return df[df["Week"].astype(str) == str(wk)]

        pdf = FPDF(); pdf.add_page(); pdf.set_font("Arial", size=14)
        pdf.cell(0, 10, txt=f"Bears Weekly Report - Week {week_pdf}", ln=True)

        def add_table(title, df):
            pdf.set_font("Arial", size=12); pdf.ln(4); pdf.cell(0, 8, txt=title, ln=True)
            pdf.set_font("Arial", size=9)
            if df is None or df.empty:
                pdf.cell(0, 6, txt="(no data)", ln=True); return
            cols = list(df.columns)[:6]
            pdf.cell(0, 6, txt=", ".join(cols), ln=True)
            for _, r in df[cols].head(25).iterrows():
                row_txt = ", ".join([str(r[c]) for c in cols])
                pdf.cell(0, 5, txt=row_txt[:120], ln=True)

        add_table("Key Notes (Strategy)", _wk(read_sheet("Strategy"), week_pdf))
        add_table("Media Summaries", _wk(read_sheet("Media_Summaries"), week_pdf))
        add_table("Injuries", _wk(read_sheet("Injuries"), week_pdf))
        add_table("Opponent Preview", _wk(read_sheet("Opponent_Preview"), week_pdf))
        add_table("Predictions", _wk(read_sheet("Predictions"), week_pdf))
        add_table("Offense", _wk(read_sheet("Offense"), week_pdf))
        add_table("Defense", _wk(read_sheet("Defense"), week_pdf))
        add_table("DVOA Proxy", _wk(read_sheet("DVOA_Proxy"), week_pdf))

        out_name = f"W{int(week_pdf):02d}_Final.pdf"
        pdf.output(out_name)
        with open(out_name, "rb") as f:
            st.download_button("Download Weekly Final PDF", f.read(), file_name=out_name, mime="application/pdf")
        st.success(f"Exported {out_name}")

st.divider()
st.subheader("📋 Metric Layout (Team vs NFL Avg)")
try:
    off = read_sheet("Offense"); nfl_ytd = read_sheet("NFL_Averages_YTD")
    if not off.empty:
        show_cols = [c for c in ["Week","Opponent","YPA","CMP%","SR%","Penalties","Penalty_Yards","YAC"] if c in off.columns]
        st.dataframe(off[show_cols].sort_values(["Week","Opponent"]))
    if not nfl_ytd.empty:
        st.dataframe(nfl_ytd)
except Exception:
    pass

st.subheader("🗓️ Weekly Layout (NFL Avg by Week)")
try:
    nfl_w = read_sheet("NFL_Averages_Weekly")
    if not nfl_w.empty:
        st.dataframe(nfl_w.sort_values("Week"))
    else:
        st.info("Use 'Fetch NFL Data (Auto)' to populate weekly averages.")
except Exception:
    pass

st.divider()
st.info("Expected CSV headers (if you use uploads): "
        "Offense: Week,Opponent,YPA,CMP%,SR%,Points,Penalties,Penalty_Yards,YAC | "
        "Defense: Week,Opponent,RZ%_Allowed,SACKs,Pressures,INTs,Penalties,Penalty_Yards,YAC_Allowed. "
        "Inline forms save to Strategy, Media_Summaries, Injuries, Opponent_Preview, Predictions.")









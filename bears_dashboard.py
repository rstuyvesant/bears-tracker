# 🐻 Chicago Bears 2025–26 Weekly Tracker (with Extra Stats)
# - Adds support for extra metrics (Penalties, Penalty_Yards, YAC, YAC_Allowed, etc.)
# - Normalizes Opponent to 3-letter codes (e.g., MIN, CHI)
# - Optional CSV uploaders for Media Summaries, Injuries, Predictions
# - Auto fetches NFL weekly averages via nfl_data_py (if available)
# - Stores all data into bears_weekly_analytics.xlsx
# - Includes DVOA-Proxy computation and PDF/Excel exports

import os
import io
import sys
import math
import json
import pandas as pd
import streamlit as st

# Optional libraries used in features (guard-imported)
try:
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows
except Exception:
    openpyxl = None

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

# nfl_data_py is optional; app still runs without it
try:
    import nfl_data_py as nfl
except Exception:
    nfl = None

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.caption("Now with extra stats: Penalties, Penalty_Yards, YAC, and more. Use 3‑letter opponent codes (e.g., MIN, CHI). Filenames are case‑sensitive on Linux/Streamlit Cloud.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# =========================
# Opponent normalization
# =========================
TEAM_MAP = {
    # NFC North
    "chicago": "CHI", "bears": "CHI", "chi": "CHI",
    "detroit": "DET", "lions": "DET", "det": "DET",
    "green bay": "GB", "packers": "GB", "gb": "GB",
    "minnesota": "MIN", "vikings": "MIN", "minn": "MIN", "min": "MIN",
    # NFC East
    "dallas": "DAL", "cowboys": "DAL", "dal": "DAL",
    "new york giants": "NYG", "giants": "NYG", "nyg": "NYG",
    "philadelphia": "PHI", "eagles": "PHI", "phi": "PHI",
    "washington": "WAS", "commanders": "WAS", "was": "WAS", "wsh": "WAS",
    # NFC South
    "atlanta": "ATL", "falcons": "ATL", "atl": "ATL",
    "carolina": "CAR", "panthers": "CAR", "car": "CAR",
    "new orleans": "NO", "saints": "NO", "no": "NO",
    "tampa bay": "TB", "buccaneers": "TB", "bucs": "TB", "tb": "TB",
    # NFC West
    "arizona": "ARI", "cardinals": "ARI", "ari": "ARI",
    "los angeles rams": "LAR", "rams": "LAR", "lar": "LAR",
    "san francisco": "SF", "49ers": "SF", "niners": "SF", "sf": "SF",
    "seattle": "SEA", "seahawks": "SEA", "sea": "SEA",
    # AFC North
    "baltimore": "BAL", "ravens": "BAL", "bal": "BAL",
    "cincinnati": "CIN", "bengals": "CIN", "cin": "CIN",
    "cleveland": "CLE", "browns": "CLE", "cle": "CLE",
    "pittsburgh": "PIT", "steelers": "PIT", "pit": "PIT",
    # AFC East
    "buffalo": "BUF", "bills": "BUF", "buf": "BUF",
    "miami": "MIA", "dolphins": "MIA", "mia": "MIA",
    "new england": "NE", "patriots": "NE", "ne": "NE",
    "new york jets": "NYJ", "jets": "NYJ", "nyj": "NYJ",
    # AFC South
    "houston": "HOU", "texans": "HOU", "hou": "HOU",
    "indianapolis": "IND", "colts": "IND", "ind": "IND",
    "jacksonville": "JAX", "jaguars": "JAX", "jax": "JAX",
    "tennessee": "TEN", "titans": "TEN", "ten": "TEN",
    # AFC West
    "denver": "DEN", "broncos": "DEN", "den": "DEN",
    "kansas city": "KC", "chiefs": "KC", "kc": "KC",
    "las vegas": "LV", "raiders": "LV", "lv": "LV",
    "los angeles chargers": "LAC", "chargers": "LAC", "lac": "LAC",
}

def canon_team(x: str) -> str:
    x = (x or "").strip()
    if not x:
        return x
    k = x.lower()
    return TEAM_MAP.get(k, x.upper())

# =========================
# Utility: Excel IO
# =========================

def _ensure_openpyxl():
    if openpyxl is None:
        st.error("openpyxl is required for Excel features. Please add it to requirements.txt")
        st.stop()


def append_to_excel(new_df: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, dedup_keys=None):
    """Append DataFrame to Excel sheet; create file/sheet if missing; optional dedup by keys."""
    _ensure_openpyxl()
    new_df = new_df.copy()
    # Normalize common fields if present
    if "Opponent" in new_df.columns:
        new_df["Opponent"] = new_df["Opponent"].apply(canon_team)
    if "Week" in new_df.columns:
        new_df["Week"] = pd.to_numeric(new_df["Week"], errors="coerce").astype("Int64")

    if os.path.exists(file_name):
        book = openpyxl.load_workbook(file_name)
    else:
        book = openpyxl.Workbook()
        # remove default sheet
        if "Sheet" in book.sheetnames:
            std = book["Sheet"]
            book.remove(std)

    if sheet_name in book.sheetnames:
        sheet = book[sheet_name]
        existing = pd.DataFrame(sheet.values)
        if not existing.empty:
            existing.columns = existing.iloc[0]
            existing = existing[1:]
        else:
            existing = pd.DataFrame(columns=new_df.columns)
        # align columns
        all_cols = list(dict.fromkeys(list(existing.columns) + list(new_df.columns)))
        existing = existing.reindex(columns=all_cols)
        new_df = new_df.reindex(columns=all_cols)
        combined = pd.concat([existing, new_df], ignore_index=True)
        # optional de-dup
        if dedup_keys:
            combined = combined.drop_duplicates(subset=dedup_keys, keep="last")
        # rewrite sheet
        book.remove(sheet)
        sheet = book.create_sheet(sheet_name)
        sheet.append(list(combined.columns))
        for _, row in combined.iterrows():
            sheet.append(list(row.values))
    else:
        # create new sheet
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
# Extra stats + NFL averages
# =========================

EXTRA_COLUMNS = [
    # offense team perspective
    "Penalties", "Penalty_Yards", "YAC", "YAC_Allowed",  # new flexible fields
]


def fetch_nfl_averages_weekly():
    """Fetch league-wide weekly averages using nfl_data_py if available.
    Adds extra stats if columns exist in the pulled dataset. Stores into Excel sheets:
    - NFL_Averages_Weekly (per-week league averages)
    - NFL_Averages_YTD (to-date average across weeks)
    """
    if nfl is None:
        st.warning("nfl_data_py not installed; skipping auto NFL averages fetch.")
        return

    try:
        # Pull team weekly stats for current season; adjust season if provided
        # nfl.import_schedules/seasons might vary by library version; we try a robust call
        # We'll default to 2025 regular season weeks 1-18
        season = 2025
        team_week = nfl.import_team_desc()  # fallback small table; ensures package works
        # Prefer: nfl.import_seasonal_team_stats or nfl.import_weekly_team_stats if available
        df = None
        for fn_name in [
            "import_weekly_team_stats",
            "import_seasonal_team_stats",
            "import_team_game_logs",
        ]:
            fn = getattr(nfl, fn_name, None)
            if fn is None:
                continue
            try:
                tmp = fn([season])
                if isinstance(tmp, pd.DataFrame) and not tmp.empty:
                    df = tmp.copy()
                    break
            except Exception:
                continue

        if df is None or df.empty:
            st.warning("Could not fetch weekly team stats from nfl_data_py; skipping averages.")
            return

        # Normalize column naming to a friendly set
        # We search for common penalty/YAC-like fields and coerce to our target names
        colmap = {}
        for col in df.columns:
            lc = str(col).lower()
            if "penalt" in lc and "yards" in lc:
                colmap[col] = "Penalty_Yards"
            elif "penalt" in lc and ("count" in lc or lc.endswith("s")):
                colmap[col] = "Penalties"
            elif "yac" in lc or "yards_after_catch" in lc:
                colmap[col] = "YAC"
            elif "yac_allowed" in lc or ("yards_after_catch" in lc and "allowed" in lc):
                colmap[col] = "YAC_Allowed"
            elif lc in ("team", "team_abbr", "team_code", "club_code"):
                colmap[col] = "Team"
            elif lc in ("week", "game_week"):
                colmap[col] = "Week"
            elif lc in ("season", "year"):
                colmap[col] = "Season"
        if colmap:
            df = df.rename(columns=colmap)

        # Keep essential fields
        keep = [c for c in ["Season", "Week", "Team", "Penalties", "Penalty_Yards", "YAC", "YAC_Allowed"] if c in df.columns]
        if not keep:
            st.warning("Fetched data lacks recognizable extra-stat columns; saving nothing.")
            return
        df_small = df[keep].copy()

        # Compute league averages per week
        if "Week" in df_small.columns:
            weekly_avg = df_small.groupby("Week").mean(numeric_only=True).reset_index()
            weekly_avg.insert(0, "Season", season)
            append_to_excel(weekly_avg, "NFL_Averages_Weekly", dedup_keys=["Season", "Week"])

        # Compute YTD averages across available weeks
        ytd_avg = df_small.mean(numeric_only=True).to_frame().T
        ytd_avg.insert(0, "Season", season)
        append_to_excel(ytd_avg, "NFL_Averages_YTD", dedup_keys=["Season"])
        st.success("NFL averages updated (weekly + YTD), including extra stats when available.")
    except Exception as e:
        st.warning(f"NFL averages fetch failed gracefully: {e}")

# =========================
# DVOA-Proxy (simple, robust)
# =========================

def safe_num(x):
    try:
        return float(x)
    except Exception:
        return math.nan


def compute_dvoa_proxy(off_df: pd.DataFrame, def_df: pd.DataFrame) -> pd.DataFrame:
    """Compute a simple DVOA-like proxy combining a few offensive/defensive indicators.
    Robust to missing columns. Returns a frame with [Week, Opponent, DVOA_Proxy]."""
    if off_df is None:
        off_df = pd.DataFrame()
    if def_df is None:
        def_df = pd.DataFrame()

    # Expected columns (best effort):
    # Offense: YPA, CMP%, SR% (Success Rate), Points, Penalties, Penalty_Yards, YAC
    # Defense: RZ%_Allowed, SACKs, Pressures, INTs, Penalties, Penalty_Yards, YAC_Allowed
    # We'll compute per-row proxy and then average by [Week, Opponent]

    def score_row(row):
        # offense contributions
        ypa = safe_num(row.get("YPA", math.nan))
        cmpct = safe_num(row.get("CMP%", math.nan))
        sr = safe_num(row.get("SR%", math.nan))
        yac = safe_num(row.get("YAC", math.nan))
        pen = safe_num(row.get("Penalties", math.nan))
        pen_yards = safe_num(row.get("Penalty_Yards", math.nan))

        # defense contributions (negatives lower proxy)
        rz_allowed = safe_num(row.get("RZ%_Allowed", math.nan))
        sacks = safe_num(row.get("SACKs", math.nan))
        pressures = safe_num(row.get("Pressures", math.nan))
        ints = safe_num(row.get("INTs", math.nan))
        yac_allowed = safe_num(row.get("YAC_Allowed", math.nan))

        # Z-score-ish scaling via simple min-max clamps
        def nz(v):
            return 0 if math.isnan(v) else v
        # positive contributions
        pos = 0.30 * nz(ypa) + 0.20 * nz(cmpct) + 0.20 * nz(sr) + 0.10 * nz(yac) + 0.10 * nz(sacks) + 0.10 * nz(pressures)
        # negative contributions
        neg = 0.25 * nz(rz_allowed) + 0.15 * nz(ints) * -1 + 0.15 * nz(pen) + 0.10 * nz(pen_yards) + 0.10 * nz(yac_allowed)
        return pos - neg

    # Join offense/defense on Week/Opponent if both provided
    join_cols = [c for c in ["Week", "Opponent"] if c in off_df.columns and c in def_df.columns]
    if len(join_cols) == 2 and not off_df.empty and not def_df.empty:
        merged = pd.merge(off_df, def_df, on=join_cols, how="outer", suffixes=("_OFF", "_DEF"))
        # Map unified keys
        merged["YPA"] = merged.get("YPA_OFF")
        merged["CMP%"] = merged.get("CMP%_OFF")
        merged["SR%"] = merged.get("SR%_OFF")
        merged["YAC"] = merged.get("YAC_OFF")
        merged["Penalties"] = merged.get("Penalties_OFF")
        merged["Penalty_Yards"] = merged.get("Penalty_Yards_OFF")
        merged["RZ%_Allowed"] = merged.get("RZ%_Allowed_DEF")
        merged["SACKs"] = merged.get("SACKs_DEF")
        merged["Pressures"] = merged.get("Pressures_DEF")
        merged["INTs"] = merged.get("INTs_DEF")
        merged["YAC_Allowed"] = merged.get("YAC_Allowed_DEF")
        base = merged
    else:
        base = off_df.copy() if not off_df.empty else def_df.copy()

    if base.empty:
        return pd.DataFrame(columns=["Week", "Opponent", "DVOA_Proxy"])    

    base["DVOA_Proxy"] = base.apply(score_row, axis=1)
    out = base.groupby([c for c in ["Week", "Opponent"] if c in base.columns], dropna=False)["DVOA_Proxy"].mean().reset_index()
    return out

# =========================
# Upload helpers
# =========================

REQUIRED_SCHEMAS = {
    "Offense": ["Week","Opponent","YPA","CMP%","SR%","Points","Penalties","Penalty_Yards","YAC"],
    "Defense": ["Week","Opponent","RZ%_Allowed","SACKs","Pressures","INTs","Penalties","Penalty_Yards","YAC_Allowed"],
    "Personnel": ["Week","Opponent","11","12","13","21","Other"],
    "Snap_Counts": ["Week","Opponent","Player","Snaps","Snap%","Side"],
    "Strategy": ["Week","Opponent","Pre_Game_Strategy","Post_Game_Results","Key_Takeaways","Next_Week_Impact"],
    # Optional management sheets
    "Media_Summaries": ["Week","Opponent","Source","Summary"],
    "Injuries": ["Week","Opponent","Player","Status","BodyPart","Practice","GameStatus","Notes"],
    "Predictions": ["Week","Opponent","Predicted_Winner","Confidence","Rationale"],
}


def load_csv_or_excel(uploaded) -> pd.DataFrame:
    if uploaded is None:
        return pd.DataFrame()
    name = (uploaded.name or "").lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(uploaded)
        return pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        return pd.DataFrame()


def validate_and_normalize(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    req = REQUIRED_SCHEMAS.get(kind, [])
    missing = [c for c in req if c not in df.columns]
    if missing:
        st.warning(f"{kind}: Missing columns {missing}. I will still try to process existing columns.")
    df = df.copy()
    if "Opponent" in df.columns:
        df["Opponent"] = df["Opponent"].apply(canon_team)
    if "Week" in df.columns:
        df["Week"] = pd.to_numeric(df["Week"], errors="coerce").astype("Int64")
    return df

# =========================
# Sidebar controls
# =========================
with st.sidebar:
    st.header("Controls")
    if st.button("Fetch NFL Data (Auto)"):
        fetch_nfl_averages_weekly()

    st.divider()
    if st.button("Clear Manual NFL Averages"):
        # Overwrite manual averages sheet with empty
        append_to_excel(pd.DataFrame(columns=["Season","Week"]), "NFL_Averages_Manual", dedup_keys=["Season","Week"])
        st.success("Manual NFL averages cleared (sheet reset).")

    st.divider()
    # Download all data as a single Excel file
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 Download All Data (Excel)",
                data=f.read(),
                file_name=EXCEL_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# =========================
# Main sections: Uploaders
# =========================

st.subheader("📈 Offense (CSV)")
off_up = st.file_uploader("Upload offense.csv", type=["csv","xlsx"], key="off_up")
if off_up:
    df_off = load_csv_or_excel(off_up)
    df_off = validate_and_normalize(df_off, "Offense")
    if not df_off.empty:
        append_to_excel(df_off, "Offense", dedup_keys=["Week","Opponent"])
        st.success(f"Appended {len(df_off)} offense rows (with extra stats if present).")

st.subheader("🛡️ Defense (CSV)")
def_up = st.file_uploader("Upload defense.csv", type=["csv","xlsx"], key="def_up")
if def_up:
    df_def = load_csv_or_excel(def_up)
    df_def = validate_and_normalize(df_def, "Defense")
    if not df_def.empty:
        append_to_excel(df_def, "Defense", dedup_keys=["Week","Opponent"])
        st.success(f"Appended {len(df_def)} defense rows (with extra stats if present).")

st.subheader("👥 Personnel Usage (CSV)")
per_up = st.file_uploader("Upload personnel.csv", type=["csv","xlsx"], key="per_up")
if per_up:
    df_per = load_csv_or_excel(per_up)
    df_per = validate_and_normalize(df_per, "Personnel")
    if not df_per.empty:
        append_to_excel(df_per, "Personnel", dedup_keys=["Week","Opponent"])
        st.success(f"Appended {len(df_per)} personnel rows.")

st.subheader("⏱️ Snap Counts (CSV)")
snap_up = st.file_uploader("Upload snap_counts.csv", type=["csv","xlsx"], key="snap_up")
if snap_up:
    df_snap = load_csv_or_excel(snap_up)
    df_snap = validate_and_normalize(df_snap, "Snap_Counts")
    if not df_snap.empty:
        append_to_excel(df_snap, "Snap_Counts", dedup_keys=["Week","Opponent","Player","Side"])  # robust de-dup
        st.success(f"Appended {len(df_snap)} snap-count rows.")

st.subheader("🧠 Strategy Notes (CSV)")
strat_up = st.file_uploader("Upload strategy.csv", type=["csv","xlsx"], key="strat_up")
if strat_up:
    df_strat = load_csv_or_excel(strat_up)
    df_strat = validate_and_normalize(df_strat, "Strategy")
    if not df_strat.empty:
        append_to_excel(df_strat, "Strategy", dedup_keys=["Week","Opponent"]) 
        st.success(f"Appended {len(df_strat)} strategy rows.")

# Optional management data uploaders
st.subheader("📚 Import Media Summaries (optional)")
up_media = st.file_uploader("Upload media_summaries.csv", type=["csv","xlsx"], key="media_up")
if up_media:
    df_media = load_csv_or_excel(up_media)
    df_media = validate_and_normalize(df_media, "Media_Summaries")
    if not df_media.empty:
        append_to_excel(df_media, "Media_Summaries", dedup_keys=["Week","Opponent","Source","Summary"]) 
        st.success(f"Imported {len(df_media)} media summaries.")

st.subheader("🩹 Import Injuries (optional)")
up_inj = st.file_uploader("Upload injuries.csv", type=["csv","xlsx"], key="inj_up")
if up_inj:
    df_inj = load_csv_or_excel(up_inj)
    df_inj = validate_and_normalize(df_inj, "Injuries")
    if not df_inj.empty:
        append_to_excel(df_inj, "Injuries", dedup_keys=["Week","Opponent","Player"]) 
        st.success(f"Imported {len(df_inj)} injury rows.")

st.subheader("🔮 Import Weekly Game Predictions (optional)")
up_pred = st.file_uploader("Upload weekly_game_predictions.csv", type=["csv","xlsx"], key="pred_up")
if up_pred:
    df_pred = load_csv_or_excel(up_pred)
    df_pred = validate_and_normalize(df_pred, "Predictions")
    if not df_pred.empty:
        append_to_excel(df_pred, "Predictions", dedup_keys=["Week","Opponent"]) 
        st.success(f"Imported {len(df_pred)} predictions.")

# =========================
# Inline forms (for quick adds)
# =========================
st.divider()
st.subheader("Quick Add: Media Summary")
with st.form("media_form"):
    week_m = st.number_input("Week", min_value=1, max_value=30, step=1)
    opp_m = st.text_input("Opponent (3-letter code or name)")
    src_m = st.text_input("Source")
    sum_m = st.text_area("Summary")
    if st.form_submit_button("Save Summary"):
        row = pd.DataFrame([[week_m, canon_team(opp_m), src_m, sum_m]], columns=REQUIRED_SCHEMAS["Media_Summaries"]) 
        append_to_excel(row, "Media_Summaries", dedup_keys=["Week","Opponent","Source","Summary"]) 
        st.success("Media summary saved.")

st.subheader("Quick Add: Prediction")
with st.form("pred_form"):
    week_p = st.number_input("Week ", min_value=1, max_value=30, step=1, key="wk_pred")
    opp_p = st.text_input("Opponent (3-letter code or name)", key="opp_pred")
    win_p = st.text_input("Predicted Winner (e.g., CHI)")
    conf_p = st.slider("Confidence (0–100)", 0, 100, 60)
    rat_p = st.text_area("Rationale")
    if st.form_submit_button("Save Prediction"):
        row = pd.DataFrame([[week_p, canon_team(opp_p), win_p.upper(), conf_p, rat_p]], columns=REQUIRED_SCHEMAS["Predictions"]) 
        append_to_excel(row, "Predictions", dedup_keys=["Week","Opponent"]) 
        st.success("Prediction saved.")

# =========================
# DVOA Proxy + Summary Preview
# =========================
st.divider()
st.subheader("📊 DVOA-Proxy (Auto)")
try:
    off_df = read_sheet("Offense")
    def_df = read_sheet("Defense")
    if not off_df.empty or not def_df.empty:
        proxy = compute_dvoa_proxy(off_df, def_df)
        if not proxy.empty:
            append_to_excel(proxy, "DVOA_Proxy", dedup_keys=["Week","Opponent"]) 
            st.dataframe(proxy.sort_values(["Week","Opponent"]))
        else:
            st.info("DVOA-Proxy: no rows available yet.")
    else:
        st.info("Upload Offense/Defense to compute DVOA-Proxy.")
except Exception as e:
    st.warning(f"DVOA-Proxy calculation skipped: {e}")

# =========================
# Simple PDF Export (Final report)
# =========================
st.divider()
st.subheader("🧾 Export Weekly Final PDF")
week_pdf = st.number_input("Week to export", min_value=1, max_value=30, value=1, step=1)

if st.button("Export Final PDF"):
    if FPDF is None:
        st.error("Install fpdf to enable PDF exports (add 'fpdf' to requirements.txt)")
    else:
        # Gather quick summaries for the week
        def _wk(df, wk):
            if df is None or df.empty:
                return pd.DataFrame()
            if "Week" not in df.columns:
                return pd.DataFrame()
            return df[df["Week"].astype(str) == str(wk)]

        off_w = _wk(read_sheet("Offense"), week_pdf)
        def_w = _wk(read_sheet("Defense"), week_pdf)
        per_w = _wk(read_sheet("Personnel"), week_pdf)
        med_w = _wk(read_sheet("Media_Summaries"), week_pdf)
        pred_w = _wk(read_sheet("Predictions"), week_pdf)
        dvoa_w = _wk(read_sheet("DVOA_Proxy"), week_pdf)

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=14)
        pdf.cell(0, 10, txt=f"Bears Weekly Report - Week {week_pdf}", ln=True)

        def add_table(title, df):
            pdf.set_font("Arial", size=12)
            pdf.ln(4)
            pdf.cell(0, 8, txt=title, ln=True)
            pdf.set_font("Arial", size=9)
            if df is None or df.empty:
                pdf.cell(0, 6, txt="(no data)", ln=True)
                return
            cols = list(df.columns)[:6]
            pdf.cell(0, 6, txt=", ".join(cols), ln=True)
            for _, r in df[cols].head(25).iterrows():
                row_txt = ", ".join([str(r[c]) for c in cols])
                pdf.cell(0, 5, txt=row_txt[:120], ln=True)

        add_table("Offense", off_w)
        add_table("Defense", def_w)
        add_table("Personnel", per_w)
        add_table("Media Summaries", med_w)
        add_table("Predictions", pred_w)
        add_table("DVOA Proxy", dvoa_w)

        out_name = f"W{int(week_pdf):02d}_Final.pdf"
        pdf.output(out_name)
        with open(out_name, "rb") as f:
            st.download_button("Download Weekly Final PDF", f.read(), file_name=out_name, mime="application/pdf")
        st.success(f"Exported {out_name}")

# =========================
# Metric & Weekly layout previews (lightweight)
# =========================
st.divider()
st.subheader("📋 Metric Layout (Team vs NFL Avg)")
try:
    off = read_sheet("Offense")
    nfl_ytd = read_sheet("NFL_Averages_YTD")
    if not off.empty:
        # Show selected columns if present
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
        st.info("Run 'Fetch NFL Data (Auto)' to populate weekly averages.")
except Exception:
    pass

st.divider()
st.info("Expected CSV headers — Offense: Week,Opponent,YPA,CMP%,SR%,Points,Penalties,Penalty_Yards,YAC | Defense: Week,Opponent,RZ%_Allowed,SACKs,Pressures,INTs,Penalties,Penalty_Yards,YAC_Allowed. You can include additional columns; they will be preserved.")







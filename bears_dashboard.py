# bears_dashboard.py
# Streamlit Bears Tracker (clean, fixed, and consolidated)
# - Six main sections
# - Sidebar: NFL updates, Snap updates, Color Code legend, Downloads (Excel/PDF)
# - Safe Excel handling + nfl_data_py 0.3.2 compatibility
# - Uses st.dataframe(..., width="stretch")

import os
import io
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np

# --- IMPORTANT: fixes the Styler annotation crash on Streamlit Cloud ---
from pandas.io.formats.style import Styler

# Excel safety + load
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from zipfile import BadZipFile

# NFL data (0.3.2+ API)
import nfl_data_py as nfl


# ==============================
# App Constants & Config
# ==============================
st.set_page_config(page_title="Bears 2025–26 Tracker", layout="wide")

DATA_DIR = "./data"
EXPORTS_DIR = "./exports"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(DATA_DIR, "bears_weekly_analytics.xlsx")

# Canonical sheet names used throughout the app
SHEET_OFFENSE = "Offense"
SHEET_DEFENSE = "Defense"
SHEET_PERSONNEL = "Personnel"
SHEET_SNAP_COUNTS = "SnapCounts"
SHEET_INJURIES = "Injuries"
SHEET_MEDIA = "MediaSummaries"
SHEET_OPP_PREVIEW = "OpponentPreview"
SHEET_PREDICTIONS = "Predictions"
SHEET_NFL_AVG_MANUAL = "NFL_Averages_Manual"
SHEET_YTD_TEAM_OFF = "YTD_Team_Offense"
SHEET_YTD_TEAM_DEF = "YTD_Team_Defense"
SHEET_YTD_NFL_OFF = "YTD_NFL_Offense"
SHEET_YTD_NFL_DEF = "YTD_NFL_Defense"

ALL_SHEETS = [
    SHEET_OFFENSE, SHEET_DEFENSE, SHEET_PERSONNEL, SHEET_SNAP_COUNTS,
    SHEET_INJURIES, SHEET_MEDIA, SHEET_OPP_PREVIEW, SHEET_PREDICTIONS,
    SHEET_NFL_AVG_MANUAL, SHEET_YTD_TEAM_OFF, SHEET_YTD_TEAM_DEF,
    SHEET_YTD_NFL_OFF, SHEET_YTD_NFL_DEF
]

# For controls
NFL_TEAM = "CHI"   # standardized 3-letter team code
CURRENT_SEASON = 2025

# ==============================
# Utilities: Files / Excel
# ==============================

def ensure_xlsx(path: str) -> str:
    """
    Ensure an .xlsx exists at `path` and is a valid zip/xlsx.
    If the file is missing, has the wrong suffix, or is corrupt, recreate it.
    Returns the (possibly corrected) .xlsx path.
    """
    # Normalize to .xlsx if someone passed a CSV by mistake
    if not path.lower().endswith(".xlsx"):
        base, _ = os.path.splitext(path)
        path = base + ".xlsx"

    # Create if missing
    if not os.path.exists(path):
        wb = Workbook()
        # Create all sheets we care about up-front so app logic can always find them
        ws = wb.active
        ws.title = SHEET_OFFENSE
        for name in ALL_SHEETS:
            if name != SHEET_OFFENSE:
                wb.create_sheet(title=name)
        wb.save(path)
        return path

    # Validate that it opens as a real xlsx
    try:
        _ = load_workbook(path, read_only=True)
    except (BadZipFile, KeyError, OSError):
        # Recreate a clean workbook with expected sheets
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_OFFENSE
        for name in ALL_SHEETS:
            if name != SHEET_OFFENSE:
                wb.create_sheet(title=name)
        wb.save(path)

    return path


def safe_load_workbook(path: str):
    """
    Wrapper around openpyxl.load_workbook that guarantees `path` is a valid .xlsx.
    """
    valid_path = ensure_xlsx(path)
    return load_workbook(valid_path)


def read_sheet(path: str, sheet: str) -> pd.DataFrame:
    ensure_xlsx(path)
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        # Normalize columns
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


def write_sheet(path: str, sheet: str, df: pd.DataFrame):
    ensure_xlsx(path)
    # Keep simple: write the one sheet, preserve others by reading then writing back
    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xlw:
            df.to_excel(xlw, index=False, sheet_name=sheet)
    except FileNotFoundError:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as xlw:
            df.to_excel(xlw, index=False, sheet_name=sheet)


def append_to_sheet(path: str, sheet: str, df_new: pd.DataFrame, dedupe_on: list | None = None):
    """Append new rows, optionally deduplicate against selected columns."""
    df_old = read_sheet(path, sheet)
    if df_old.empty:
        df_out = df_new.copy()
    else:
        df_out = pd.concat([df_old, df_new], ignore_index=True)
    if dedupe_on:
        df_out = df_out.drop_duplicates(subset=dedupe_on, keep="last")
    write_sheet(path, sheet, df_out)
    return df_out


# ==============================
# NFL Data Helpers (0.3.2+)
# ==============================

def _normalize_team_col(df: pd.DataFrame) -> pd.DataFrame:
    """Unify team column naming across nfl_data_py frames."""
    if 'team' in df.columns:
        return df
    if 'recent_team' in df.columns:
        df = df.rename(columns={'recent_team': 'team'})
    if 'posteam' in df.columns:
        df = df.rename(columns={'posteam': 'team'})
    return df


def fetch_week_stats(season: int, week: int, team_abbr: str) -> pd.DataFrame:
    """
    Try weekly team stats first; fall back to seasonal aggregates if weekly missing.
    Works on nfl-data-py 0.3.2+.
    """
    # 1) Try weekly data
    try:
        weekly = nfl.import_weekly_data([season])  # large
        weekly = _normalize_team_col(weekly)
        if 'week' in weekly.columns:
            dfw = weekly[(weekly['team'] == team_abbr) & (weekly['week'] == week)].copy()
            if not dfw.empty:
                return dfw
    except Exception:
        pass

    # 2) Fallback: seasonal aggregates
    try:
        seasonal = nfl.import_seasonal_data([season])
        seasonal = _normalize_team_col(seasonal)
        dfs = seasonal[seasonal['team'] == team_abbr].copy()
        if not dfs.empty:
            dfs['week'] = week  # keep downstream compatibility
            return dfs
    except Exception:
        pass

    return pd.DataFrame()


def fetch_nfl_averages(season: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Compute simple per-team per-season aggregates, then NFL (league) averages for offense & defense views.
    Returns (nfl_off_avg, nfl_def_avg) as single-row DataFrames with columns prefixed 'NFL_Avg._'.
    """
    try:
        weekly = nfl.import_weekly_data([season])
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

    weekly = _normalize_team_col(weekly)

    # Minimal metric set — expand as needed
    # Example metrics (you can add CMP%, YPA, RZ%, SACKS if present in your weekly schema)
    metric_candidates = [c for c in weekly.columns if weekly[c].dtype.kind in "if"]
    if not metric_candidates:
        return pd.DataFrame(), pd.DataFrame()

    # Offensive perspective — group by team mean
    off_by_team = weekly.groupby("team")[metric_candidates].mean(numeric_only=True)
    nfl_off_avg = off_by_team.mean(numeric_only=True).to_frame().T
    nfl_off_avg.columns = [f"NFL_Avg._{c}" for c in nfl_off_avg.columns]

    # Defensive perspective: if you track allowed metrics, you can compute similarly
    # For now, just mirror the same (users often compare defense to NFL averages of those same columns)
    nfl_def_avg = nfl_off_avg.copy()

    return nfl_off_avg.reset_index(drop=True), nfl_def_avg.reset_index(drop=True)


def fetch_snap_counts(season: int, team: str) -> pd.DataFrame:
    try:
        snaps = nfl.import_snap_counts([season])
    except Exception:
        return pd.DataFrame()
    snaps = _normalize_team_col(snaps)
    return snaps[snaps["team"] == team].copy()


# ==============================
# Styling (Color Codes vs NFL Avg)
# ==============================

def _metric_pairs_for_sheet(sheet_name: str) -> tuple[list[str], list[str]]:
    """
    Define which metrics are 'better when higher' or 'better when lower'
    based on the sheet context (Offense vs Defense etc.).
    Extend as needed.
    """
    # Defaults
    better_high = ["YPA", "YPC", "CMP%", "QBR", "Points", "Yards", "Success_Rate", "YAC"]
    better_low = ["SACKs_Allowed", "TO", "INT_Thrown", "Fumbles", "Penalties"]

    if sheet_name == SHEET_DEFENSE:
        better_high = ["SACKs", "INTs", "FF", "FR", "Pressures"]
        better_low = ["3D%_Allowed", "RZ%_Allowed", "YPA_Allowed", "YPC_Allowed", "Points_Allowed", "Yards_Allowed"]

    return better_high, better_low


def _style_by_nfl_avg(df: pd.DataFrame, sheet_name: str) -> Styler:
    """
    Color-code dataframe cells by comparing to NFL_Avg._ columns if present.
    Green = better than NFL average; Red = worse than NFL average, based on metric orientation.
    """
    if df.empty:
        return df.style

    better_high, better_low = _metric_pairs_for_sheet(sheet_name)

    # Map raw metric names to NFL_Avg columns
    avg_map = {c.replace("NFL_Avg._", ""): c for c in df.columns if c.startswith("NFL_Avg._")}
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

    def _colorize(val, col_name):
        base_name = col_name
        avg_col = avg_map.get(base_name)
        if avg_col is None or avg_col not in df.columns:
            return ""

        try:
            nfl_avg = df[avg_col].iloc[0]  # NFL average row replicated? If not, adapt.
        except Exception:
            return ""

        if pd.isna(val) or pd.isna(nfl_avg):
            return ""

        # Decide direction
        if base_name in better_high:
            return "background-color: #e5ffe5" if val > nfl_avg else "background-color: #ffe5e5"
        if base_name in better_low:
            return "background-color: #e5ffe5" if val < nfl_avg else "background-color: #ffe5e5"

        # For metrics we don't know, do nothing
        return ""

    def _apply(row):
        styles = []
        for col in df.columns:
            if col in numeric_cols and col in avg_map or col in better_high or col in better_low:
                styles.append(_colorize(row[col], col))
            else:
                styles.append("")
        return styles

    return df.style.apply(_apply, axis=1)


# ==============================
# PDF Export (simple)
# ==============================
try:
    from fpdf import FPDF
except Exception:
    FPDF = None


def export_week_pdf(week: int, opponent: str, summary_lines: list[str]) -> bytes:
    if FPDF is None:
        return b""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    pdf.cell(0, 10, f"Bears Weekly Report — Week {week} vs {opponent}", ln=1)
    pdf.set_font("Arial", size=11)
    for line in summary_lines:
        pdf.multi_cell(0, 8, txt=line)
    # Write to bytes
    out = io.BytesIO()
    pdf.output(out, "F")
    return out.getvalue()


# ==============================
# Sidebar (Left) — Controls
# ==============================
st.sidebar.title("Controls")

# Week & Opponent selection (global)
week_input = st.sidebar.number_input("Week", min_value=1, max_value=22, value=1, step=1)
opponent_input = st.sidebar.text_input("Opponent (3-letter code)", value="MIN")

st.sidebar.markdown("---")
st.sidebar.subheader("NFL Updates")

if st.sidebar.button("Fetch NFL Data (Auto)"):
    off_avg, def_avg = fetch_nfl_averages(CURRENT_SEASON)
    if off_avg.empty:
        st.sidebar.error("Could not compute NFL averages (library missing or data unavailable).")
    else:
        # Save to corresponding sheets
        write_sheet(EXCEL_PATH, SHEET_YTD_NFL_OFF, off_avg)
        write_sheet(EXCEL_PATH, SHEET_YTD_NFL_DEF, def_avg)
        st.sidebar.success("NFL averages computed and saved to YTD_NFL_Offense / YTD_NFL_Defense.")

st.sidebar.subheader("Snap Updates")
if st.sidebar.button("Fetch Snap Counts"):
    snaps = fetch_snap_counts(CURRENT_SEASON, NFL_TEAM)
    if snaps.empty:
        st.sidebar.warning("No snap counts fetched.")
    else:
        # Keep useful columns if available
        keep = [c for c in snaps.columns if c in ("season", "week", "team", "player", "position", "offense_snaps", "defense_snaps", "special_teams_snaps", "offense_pct", "defense_pct", "special_teams_pct")]
        if keep:
            snaps = snaps[keep]
        append_to_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS, snaps, dedupe_on=["season", "week", "team", "player"])
        st.sidebar.success(f"Snap counts saved for {NFL_TEAM} {CURRENT_SEASON}.")

st.sidebar.subheader("Color Codes (Auto)")
st.sidebar.caption("Green = better than NFL average; Red = worse (by metric orientation).")

st.sidebar.markdown("---")
st.sidebar.subheader("Downloads")

# Download Excel
if os.path.exists(EXCEL_PATH):
    with open(EXCEL_PATH, "rb") as f:
        st.sidebar.download_button(
            label="Download Excel (All Data)",
            data=f.read(),
            file_name="bears_weekly_analytics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.sidebar.info("Excel not found yet; upload a CSV or fetch to create it.")

# Download current week's PDF (if generated below)
pdf_stub_path = os.path.join(EXPORTS_DIR, f"W{week_input:02d}_{opponent_input}_Final.pdf")
if os.path.exists(pdf_stub_path):
    with open(pdf_stub_path, "rb") as f:
        st.sidebar.download_button(
            label=f"Download Final PDF (W{week_input:02d})",
            data=f.read(),
            file_name=os.path.basename(pdf_stub_path),
            mime="application/pdf"
        )
else:
    st.sidebar.caption("Final PDF will appear here after you export it.")

# ==============================
# Main Header
# ==============================
st.title("Chicago Bears 2025–26 Weekly Tracker")

st.markdown(
    """
This page renders in this order:

**1) Weekly Controls**  
**2) Upload Weekly Data**  
**3) NFL Averages (Auto & Manual)**  
**4) DVOA Proxy & Color Codes**  
**5) Opponent Preview & Strategy Notes**  
**6) Exports & Downloads**
"""
)

# Show selected week/opponent at top
c1, c2, c3 = st.columns([1, 1, 2])
with c1:
    st.write(f"**Selected Week:** {week_input}")
with c2:
    st.write(f"**Opponent:** {opponent_input}")
with c3:
    st.caption("Use the sidebar to change Week and Opponent. These selections drive uploads and exports.")


# ==============================
# 1) Weekly Controls
# ==============================
with st.expander("1) Weekly Controls", expanded=True):
    cc1, cc2 = st.columns(2)

    with cc1:
        st.markdown("**Quick Fetch (This Week)**")
        if st.button("Fetch CHI Week Stats (Auto)"):
            dfw = fetch_week_stats(CURRENT_SEASON, int(week_input), NFL_TEAM)
            if dfw.empty:
                st.warning("No weekly data found; seasonal fallback used or empty result.")
            else:
                st.success(f"Fetched data for CHI week {int(week_input)}")
                st.dataframe(dfw.head(50), width="stretch")

    with cc2:
        st.markdown("**Notes / Key Items**")
        key_notes = st.text_area("Key Notes (autosaved into OpponentPreview sheet under 'Notes')", value="", height=120)
        if st.button("Save Notes to Opponent Preview"):
            if key_notes.strip():
                row = pd.DataFrame([{
                    "season": CURRENT_SEASON,
                    "week": int(week_input),
                    "team": NFL_TEAM,
                    "opponent": opponent_input.strip().upper(),
                    "Notes": key_notes.strip(),
                    "saved_at": datetime.now().isoformat(timespec="seconds")
                }])
                append_to_sheet(EXCEL_PATH, SHEET_OPP_PREVIEW, row, dedupe_on=["season", "week", "team", "opponent"])
                st.success("Notes saved.")
            else:
                st.info("Nothing to save.")


# ==============================
# 2) Upload Weekly Data
# ==============================
with st.expander("2) Upload Weekly Data", expanded=True):
    st.caption("Upload CSVs for Offense, Defense, Personnel, and Snap Counts (if you prefer manual uploads). Files are appended and deduplicated by common keys when possible.")

    upc1, upc2 = st.columns(2)

    with upc1:
        st.markdown("**Offense CSV**")
        f_off = st.file_uploader("Upload Offense CSV", type=["csv"], key="up_off")
        if f_off is not None:
            try:
                df = pd.read_csv(f_off)
                df["season"] = df.get("season", CURRENT_SEASON)
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.copy()
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_OFFENSE, df, dedupe_on=["season", "week", "team"])
                st.success(f"Offense rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")

        st.markdown("**Defense CSV**")
        f_def = st.file_uploader("Upload Defense CSV", type=["csv"], key="up_def")
        if f_def is not None:
            try:
                df = pd.read_csv(f_def)
                df["season"] = df.get("season", CURRENT_SEASON)
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.copy()
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_DEFENSE, df, dedupe_on=["season", "week", "team"])
                st.success(f"Defense rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")

    with upc2:
        st.markdown("**Personnel CSV**")
        f_per = st.file_uploader("Upload Personnel CSV", type=["csv"], key="up_per")
        if f_per is not None:
            try:
                df = pd.read_csv(f_per)
                df["season"] = df.get("season", CURRENT_SEASON)
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.copy()
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_PERSONNEL, df, dedupe_on=["season", "week", "team"])
                st.success(f"Personnel rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")

        st.markdown("**Snap Counts CSV (Manual)**")
        f_snap = st.file_uploader("Upload Snap Counts CSV", type=["csv"], key="up_snap")
        if f_snap is not None:
            try:
                df = pd.read_csv(f_snap)
                df["season"] = df.get("season", CURRENT_SEASON)
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.copy()
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS, df, dedupe_on=["season", "week", "team", "player"] if "player" in df.columns else ["season", "week", "team"])
                st.success(f"SnapCounts rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")


# ==============================
# 3) NFL Averages (Auto & Manual)
# ==============================
with st.expander("3) NFL Averages (Auto & Manual)", expanded=True):
    st.markdown("**Auto:** Use the sidebar ‘Fetch NFL Data (Auto)’ to compute/sync YTD NFL offense/defense averages.")
    off_view = read_sheet(EXCEL_PATH, SHEET_YTD_NFL_OFF)
    if not off_view.empty:
        st.write("**YTD NFL Offense Averages (Auto)**")
        st.dataframe(off_view, width="stretch")
    else:
        st.info("No auto NFL offense averages yet.")

    def_view = read_sheet(EXCEL_PATH, SHEET_YTD_NFL_DEF)
    if not def_view.empty:
        st.write("**YTD NFL Defense Averages (Auto)**")
        st.dataframe(def_view, width="stretch")
    else:
        st.info("No auto NFL defense averages yet.")

    st.markdown("---")
    st.markdown("**Manual NFL Averages CSV** (optional): Any columns here get stored to `NFL_Averages_Manual`.")
    f_nfl = st.file_uploader("Upload Manual NFL Averages CSV", type=["csv"], key="up_nflavg")
    if f_nfl is not None:
        try:
            df = pd.read_csv(f_nfl)
            # Normalize to NFL_Avg._ prefix to be used by the color styler
            df.columns = [c if c.startswith("NFL_Avg._") else f"NFL_Avg._{c}" for c in df.columns]
            df_out = append_to_sheet(EXCEL_PATH, SHEET_NFL_AVG_MANUAL, df, dedupe_on=None)
            st.success(f"Saved {len(df)} row(s) to NFL_Averages_Manual.")
            st.dataframe(df_out.tail(10), width="stretch")
        except Exception as e:
            st.error(f"Upload failed: {e}")


# ==============================
# 4) DVOA Proxy & Color Codes
# ==============================
with st.expander("4) DVOA Proxy & Color Codes", expanded=True):
    st.caption("This view merges your weekly data with NFL_Avg columns (from auto or manual) and shows green/red cells. You can extend the metric lists in the code.")

    # Compose a simple merged preview for Offense and Defense (current week)
    off_df = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
    def_df = read_sheet(EXCEL_PATH, SHEET_DEFENSE)

    # Build an NFL_Avg row to join (manual preferred, else use auto offense as a source of NFL_Avg._ columns)
    nfl_manual = read_sheet(EXCEL_PATH, SHEET_NFL_AVG_MANUAL)
    nfl_auto_off = read_sheet(EXCEL_PATH, SHEET_YTD_NFL_OFF)

    nfl_avg_row = pd.DataFrame()
    if not nfl_manual.empty:
        nfl_avg_row = nfl_manual.tail(1).copy()
    elif not nfl_auto_off.empty:
        nfl_avg_row = nfl_auto_off.tail(1).copy()

    # OFFENSE
    st.markdown("**Offense — Week Merge vs NFL Avg**")
    if not off_df.empty:
        cur_off = off_df[(off_df.get("season", CURRENT_SEASON) == CURRENT_SEASON) &
                         (off_df.get("week", week_input) == int(week_input))].copy()
        if not cur_off.empty and not nfl_avg_row.empty:
            # Only keep numeric & a few ID columns
            ids = [c for c in cur_off.columns if c in ("season", "week", "team", "opponent")]
            nums = [c for c in cur_off.columns if c not in ids and pd.api.types.is_numeric_dtype(cur_off[c])]
            preview = pd.concat([cur_off[ids + nums].reset_index(drop=True), nfl_avg_row.reset_index(drop=True)], axis=1)
            st.dataframe(_style_by_nfl_avg(preview, SHEET_OFFENSE), width="stretch")
        else:
            st.info("Need both weekly offense row(s) and an NFL_Avg row (auto or manual).")
    else:
        st.info("No Offense data yet.")

    # DEFENSE
    st.markdown("**Defense — Week Merge vs NFL Avg**")
    if not def_df.empty:
        cur_def = def_df[(def_df.get("season", CURRENT_SEASON) == CURRENT_SEASON) &
                         (def_df.get("week", week_input) == int(week_input))].copy()
        if not cur_def.empty and not nfl_avg_row.empty:
            ids = [c for c in cur_def.columns if c in ("season", "week", "team", "opponent")]
            nums = [c for c in cur_def.columns if c not in ids and pd.api.types.is_numeric_dtype(cur_def[c])]
            preview = pd.concat([cur_def[ids + nums].reset_index(drop=True), nfl_avg_row.reset_index(drop=True)], axis=1)
            st.dataframe(_style_by_nfl_avg(preview, SHEET_DEFENSE), width="stretch")
        else:
            st.info("Need both weekly defense row(s) and an NFL_Avg row (auto or manual).")
    else:
        st.info("No Defense data yet.")

    st.markdown("---")
    st.markdown("**Color Code Legend**")
    st.write("- Green = better than NFL average (direction-aware)")
    st.write("- Red = worse than NFL average (direction-aware)")


# ==============================
# 5) Opponent Preview & Strategy Notes
# ==============================
with st.expander("5) Opponent Preview & Strategy Notes", expanded=True):
    st.caption("Lightweight area for opponent notes, previews, and predictions you can extend later.")

    # Opponent Preview table
    opp = read_sheet(EXCEL_PATH, SHEET_OPP_PREVIEW)
    if not opp.empty:
        st.write("**Opponent Preview (Recent Entries)**")
        st.dataframe(opp.sort_values("saved_at" if "saved_at" in opp.columns else opp.columns[0]).tail(25), width="stretch")
    else:
        st.info("No opponent preview entries yet.")

    # Predictions (simple example)
    st.markdown("**Weekly Prediction (optional)**")
    pred_col1, pred_col2 = st.columns([2, 1])
    with pred_col1:
        rationale = st.text_area("Rationale", height=120, key="pred_rationale")
    with pred_col2:
        predicted_winner = st.selectbox("Predicted Winner", options=[NFL_TEAM, opponent_input], index=0)
        confidence = st.slider("Confidence", min_value=0, max_value=100, value=60, step=5)

    if st.button("Save Weekly Prediction"):
        row = pd.DataFrame([{
            "season": CURRENT_SEASON,
            "week": int(week_input),
            "team": NFL_TEAM,
            "opponent": opponent_input.strip().upper(),
            "Predicted_Winner": predicted_winner,
            "Confidence": confidence,
            "Rationale": rationale.strip(),
            "saved_at": datetime.now().isoformat(timespec="seconds")
        }])
        df_out = append_to_sheet(EXCEL_PATH, SHEET_PREDICTIONS, row, dedupe_on=["season", "week", "team", "opponent"])
        st.success("Prediction saved.")
        st.dataframe(df_out.tail(10), width="stretch")


# ==============================
# 6) Exports & Downloads
# ==============================
with st.expander("6) Exports & Downloads", expanded=True):
    st.caption("Create a quick PDF and download Excel/PDF here or from the sidebar.")

    if st.button("Export Final PDF (This Week)"):
        # Simple PDF: pull a couple summary lines from sheets to include
        off_this = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
        def_this = read_sheet(EXCEL_PATH, SHEET_DEFENSE)

        lines = [
            f"Week {int(week_input)} vs {opponent_input}",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            ""
        ]
        if not off_this.empty:
            offw = off_this[(off_this.get("season", CURRENT_SEASON) == CURRENT_SEASON) &
                            (off_this.get("week", week_input) == int(week_input))]
            lines.append(f"Offense rows: {len(offw)}")
        if not def_this.empty:
            defw = def_this[(def_this.get("season", CURRENT_SEASON) == CURRENT_SEASON) &
                            (def_this.get("week", week_input) == int(week_input))]
            lines.append(f"Defense rows: {len(defw)}")

        pdf_bytes = export_week_pdf(int(week_input), opponent_input, lines)
        if not pdf_bytes:
            st.error("FPDF not available — ensure fpdf is in requirements.")
        else:
            out_path = os.path.join(EXPORTS_DIR, f"W{int(week_input):02d}_{opponent_input}_Final.pdf")
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)
            st.success(f"Final PDF created: {os.path.basename(out_path)}")
            st.download_button(
                label="Download Final PDF (Just Created)",
                data=pdf_bytes,
                file_name=os.path.basename(out_path),
                mime="application/pdf",
                type="primary"
            )

    st.markdown("---")
    if os.path.exists(EXCEL_PATH):
        with open(EXCEL_PATH, "rb") as f:
            st.download_button(
                label="Download Excel (All Data)",
                data=f.read(),
                file_name="bears_weekly_analytics.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("Excel not found yet; upload a CSV or fetch to create it.")

    # Also show a quick peek of the workbook sheets
    st.markdown("**Workbook Peek**")
    peek_cols = st.columns(3)
    try:
        # --- Workbook Peek (safe rendering) ---
st.markdown("**Workbook Peek**")
peek_cols = st.columns(3)

with peek_cols[0]:
    st.write("Offense")
    off = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
    if not off.empty:
        st.dataframe(off.tail(10), width="stretch")
    else:
        st.caption("—")

with peek_cols[1]:
    st.write("Defense")
    deff = read_sheet(EXCEL_PATH, SHEET_DEFENSE)
    if not deff.empty:
        st.dataframe(deff.tail(10), width="stretch")
    else:
        st.caption("—")

with peek_cols[2]:
    st.write("SnapCounts")
    snaps = read_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS)
    if not snaps.empty:
        st.dataframe(snaps.tail(10), width="stretch")
    else:
        st.caption("—")

    except Exception as e:
        st.error(f"Peek failed: {e}")

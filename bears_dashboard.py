# bears_dashboard.py
# Chicago Bears Weekly Tracker (Week-vs-Opponent + YTD vs NFL build)
# - Adds: "Bears vs Opponent (This Week)" and "Bears YTD vs NFL (Weeks 1..W)" views
# - Keeps: six sections, sidebar controls, 404-safe imports, Season default=2024

import os
import io
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np

# Styler for color-coding
from pandas.io.formats.style import Styler

# Excel safety
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from zipfile import BadZipFile

# NFL data
import nfl_data_py as nfl

# Network error handling for remote CSVs
from urllib.error import HTTPError, URLError


# ==============================
# App Config
# ==============================
st.set_page_config(page_title="Bears Weekly Tracker", layout="wide")

DATA_DIR = "./data"
EXPORTS_DIR = "./exports"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(DATA_DIR, "bears_weekly_analytics.xlsx")

# Sheets
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
SHEET_NFL_WEEKLY_AVG = "NFL_Weekly_Averages"   # NEW (per-week league averages)

ALL_SHEETS = [
    SHEET_OFFENSE, SHEET_DEFENSE, SHEET_PERSONNEL, SHEET_SNAP_COUNTS,
    SHEET_INJURIES, SHEET_MEDIA, SHEET_OPP_PREVIEW, SHEET_PREDICTIONS,
    SHEET_NFL_AVG_MANUAL, SHEET_YTD_TEAM_OFF, SHEET_YTD_TEAM_DEF,
    SHEET_YTD_NFL_OFF, SHEET_YTD_NFL_DEF, SHEET_NFL_WEEKLY_AVG
]

# Controls
NFL_TEAM = "CHI"
DEFAULT_SEASON = 2024


# ==============================
# Safe importer helper (404/network resilient)
# ==============================
def import_df_safely(import_fn, seasons: list[int]) -> pd.DataFrame:
    try:
        return import_fn(seasons)
    except HTTPError as e:
        code = getattr(e, "code", "HTTPError")
        msg = f"{import_fn.__name__} for {seasons} returned HTTP {code}."
        if code == 404:
            msg += " The season may not be published yet."
        st.warning(msg)
    except URLError as e:
        st.warning(f"{import_fn.__name__} network error: {getattr(e, 'reason', e)}")
    except Exception as e:
        st.warning(f"{import_fn.__name__} failed: {e}")
    return pd.DataFrame()


# ==============================
# Excel helpers
# ==============================
def ensure_xlsx(path: str) -> str:
    if not path.lower().endswith(".xlsx"):
        base, _ = os.path.splitext(path)
        path = base + ".xlsx"
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_OFFENSE
        for name in ALL_SHEETS:
            if name != SHEET_OFFENSE:
                wb.create_sheet(title=name)
        wb.save(path)
        return path
    try:
        _ = load_workbook(path, read_only=True)
    except (BadZipFile, KeyError, OSError):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_OFFENSE
        for name in ALL_SHEETS:
            if name != SHEET_OFFENSE:
                wb.create_sheet(title=name)
        wb.save(path)
    return path


def read_sheet(path: str, sheet: str) -> pd.DataFrame:
    ensure_xlsx(path)
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


def write_sheet(path: str, sheet: str, df: pd.DataFrame):
    ensure_xlsx(path)
    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xlw:
            df.to_excel(xlw, index=False, sheet_name=sheet)
    except FileNotFoundError:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as xlw:
            df.to_excel(xlw, index=False, sheet_name=sheet)


def append_to_sheet(path: str, sheet: str, df_new: pd.DataFrame, dedupe_on: list | None = None):
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
# NFL helpers
# ==============================
def _normalize_team_col(df: pd.DataFrame) -> pd.DataFrame:
    if 'team' in df.columns:
        return df
    if 'recent_team' in df.columns:
        df = df.rename(columns={'recent_team': 'team'})
    if 'posteam' in df.columns:
        df = df.rename(columns={'posteam': 'team'})
    return df


def select_numeric(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]


def fetch_week_stats(season: int, week: int, team_abbr: str) -> pd.DataFrame:
    # Try weekly
    try:
        weekly = import_df_safely(nfl.import_weekly_data, [season])
        weekly = _normalize_team_col(weekly)
        if 'week' in weekly.columns and not weekly.empty:
            dfw = weekly[(weekly['team'] == team_abbr) & (weekly['week'] == week)].copy()
            if not dfw.empty:
                return dfw
    except Exception:
        pass
    # Fallback seasonal (tag week)
    try:
        seasonal = import_df_safely(nfl.import_seasonal_data, [season])
        seasonal = _normalize_team_col(seasonal)
        if not seasonal.empty:
            dfs = seasonal[seasonal['team'] == team_abbr].copy()
            if not dfs.empty:
                dfs['week'] = week
                return dfs
    except Exception:
        pass
    return pd.DataFrame()


def fetch_week_stats_both_teams(season: int, week: int, team_abbr: str, opp_abbr: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (bears_week, opponent_week) frames for the selected week."""
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    weekly = _normalize_team_col(weekly)
    if weekly.empty or 'week' not in weekly.columns:
        return pd.DataFrame(), pd.DataFrame()
    chi = weekly[(weekly['team'] == team_abbr) & (weekly['week'] == week)].copy()
    opp = weekly[(weekly['team'] == opp_abbr) & (weekly['week'] == week)].copy()
    return chi, opp


def fetch_nfl_averages_YTD(season: int, upto_week: int) -> pd.DataFrame:
    """
    League-wide averages across all team-weeks for weeks 1..upto_week.
    Returns a single-row DF with columns prefixed NFL_Avg._*
    """
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    if weekly.empty or 'week' not in weekly.columns:
        return pd.DataFrame()
    weekly = _normalize_team_col(weekly)
    df = weekly[(weekly['week'] >= 1) & (weekly['week'] <= upto_week)].copy()
    if df.empty:
        return pd.DataFrame()
    nums = select_numeric(df)
    if not nums:
        return pd.DataFrame()
    nfl_ytd = df[nums].mean(numeric_only=True).to_frame().T
    nfl_ytd.columns = [f"NFL_Avg._{c}" for c in nfl_ytd.columns]
    return nfl_ytd.reset_index(drop=True)


def fetch_bears_averages_YTD_from_source(season: int, upto_week: int, team: str) -> pd.DataFrame:
    """
    Bears averages across weeks 1..upto_week from remote weekly source.
    """
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    weekly = _normalize_team_col(weekly)
    if weekly.empty or 'week' not in weekly.columns:
        return pd.DataFrame()
    chi = weekly[(weekly['team'] == team) & (weekly['week'] >= 1) & (weekly['week'] <= upto_week)].copy()
    if chi.empty:
        return pd.DataFrame()
    nums = select_numeric(chi)
    if not nums:
        return pd.DataFrame()
    brs = chi[nums].mean(numeric_only=True).to_frame().T
    return brs.reset_index(drop=True)


def fetch_bears_averages_YTD_from_workbook(upto_week: int) -> pd.DataFrame:
    """
    Bears averages from your saved Offense + Defense sheets (weeks 1..upto_week).
    """
    off = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
    deff = read_sheet(EXCEL_PATH, SHEET_DEFENSE)
    df = pd.concat([off, deff], ignore_index=True) if (not off.empty or not deff.empty) else pd.DataFrame()
    if df.empty:
        return pd.DataFrame()
    if 'week' not in df.columns:
        return pd.DataFrame()
    df = df[(df['week'] >= 1) & (df['week'] <= upto_week)].copy()
    nums = select_numeric(df)
    if not nums:
        return pd.DataFrame()
    brs = df[nums].mean(numeric_only=True).to_frame().T
    return brs.reset_index(drop=True)


def fetch_snap_counts(season: int, team: str) -> pd.DataFrame:
    snaps = import_df_safely(nfl.import_snap_counts, [season])
    if snaps.empty:
        return pd.DataFrame()
    snaps = _normalize_team_col(snaps)
    return snaps[snaps["team"] == team].copy()


def fetch_snap_counts_week_both(season: int, week: int, team: str, opp: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    snaps = import_df_safely(nfl.import_snap_counts, [season])
    if snaps.empty:
        return pd.DataFrame(), pd.DataFrame()
    snaps = _normalize_team_col(snaps)
    has_week = 'week' in snaps.columns
    chi = snaps[(snaps["team"] == team) & ((snaps["week"] == week) if has_week else True)].copy()
    oppdf = snaps[(snaps["team"] == opp) & ((snaps["week"] == week) if has_week else True)].copy()
    return chi, oppdf


# ==============================
# Color styling vs reference row
# ==============================
def style_vs_reference(df: pd.DataFrame, ref_prefix: str, better_high: list[str], better_low: list[str]) -> Styler:
    """
    Compare numeric columns in df to a single reference row whose columns are prefixed ref_prefix (e.g., 'NFL_Avg._').
    """
    if df.empty:
        return df.style
    avg_map = {c.replace(ref_prefix, ""): c for c in df.columns if c.startswith(ref_prefix)}
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    def _cell(val, col):
        base = col
        avg_col = avg_map.get(base)
        if avg_col is None or avg_col not in df.columns:
            return ""
        try:
            ref_val = df[avg_col].iloc[0]
        except Exception:
            return ""
        if pd.isna(val) or pd.isna(ref_val):
            return ""
        if base in better_high:
            return "background-color: #e5ffe5" if val > ref_val else "background-color: #ffe5e5"
        if base in better_low:
            return "background-color: #e5ffe5" if val < ref_val else "background-color: #ffe5e5"
        return ""
    def _apply(row):
        return [
            _cell(row[c], c) if c in numeric_cols and (c in avg_map or c in better_high or c in better_low) else ""
            for c in df.columns
        ]
    return df.style.apply(_apply, axis=1)


def metric_orientation_for_context(context: str) -> tuple[list[str], list[str]]:
    # tunable lists per your earlier setup
    if context == "DEFENSE":
        better_high = ["SACKs", "INTs", "FF", "FR", "Pressures"]
        better_low = ["3D%_Allowed", "RZ%_Allowed", "YPA_Allowed", "YPC_Allowed", "Points_Allowed", "Yards_Allowed"]
    else:
        better_high = ["YPA", "YPC", "CMP%", "QBR", "Points", "Yards", "Success_Rate", "YAC"]
        better_low = ["SACKs_Allowed", "TO", "INT_Thrown", "Fumbles", "Penalties"]
    return better_high, better_low


# ==============================
# PDF Export (simple)
# ==============================
try:
    from fpdf import FPDF
except Exception:
    FPDF = None

def export_week_pdf(season: int, week: int, opponent: str, summary_lines: list[str]) -> bytes:
    if FPDF is None:
        return b""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    pdf.cell(0, 10, f"Bears Weekly Report — {season} Week {week} vs {opponent}", ln=1)
    pdf.set_font("Arial", size=11)
    for line in summary_lines:
        pdf.multi_cell(0, 8, txt=line)
    out = io.BytesIO()
    pdf.output(out, "F")
    return out.getvalue()


# ==============================
# Sidebar Controls
# ==============================
st.sidebar.title("Controls")
season_input = st.sidebar.number_input("Season", min_value=2012, max_value=2030, value=DEFAULT_SEASON, step=1)
week_input = st.sidebar.number_input("Week", min_value=1, max_value=22, value=1, step=1)
opponent_input = st.sidebar.text_input("Opponent (3-letter code)", value="MIN").strip().upper()

st.sidebar.markdown("---")
st.sidebar.subheader("NFL Updates")
if st.sidebar.button("Fetch NFL Data (Auto)"):
    # YTD league averages (weeks 1..W)
    nfl_ytd = fetch_nfl_averages_YTD(int(season_input), int(week_input))
    if nfl_ytd.empty:
        st.sidebar.error("Could not compute NFL YTD averages (source unavailable).")
    else:
        write_sheet(EXCEL_PATH, SHEET_YTD_NFL_OFF, nfl_ytd)  # store under offense slot for reuse
        write_sheet(EXCEL_PATH, SHEET_YTD_NFL_DEF, nfl_ytd)  # mirror for defense slot
        st.sidebar.success("NFL YTD averages saved.")

    # Optional: also store per-week league averages table
    weekly = import_df_safely(nfl.import_weekly_data, [int(season_input)])
    if not weekly.empty and 'week' in weekly.columns:
        week_means = weekly.groupby("week")[select_numeric(weekly)].mean(numeric_only=True).reset_index()
        write_sheet(EXCEL_PATH, SHEET_NFL_WEEKLY_AVG, week_means)
        st.sidebar.success("NFL Weekly Averages table saved.")

st.sidebar.subheader("Snap Updates")
if st.sidebar.button("Fetch Snap Counts"):
    snaps = fetch_snap_counts(int(season_input), NFL_TEAM)
    if snaps.empty:
        st.sidebar.warning("No snap counts fetched.")
    else:
        keep = [c for c in snaps.columns if c in (
            "season", "week", "team", "player", "position",
            "offense_snaps", "defense_snaps", "special_teams_snaps",
            "offense_pct", "defense_pct", "special_teams_pct"
        )]
        if keep:
            snaps = snaps[keep]
        append_to_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS, snaps,
                        dedupe_on=["season", "week", "team", "player"] if "player" in snaps.columns else ["season", "week", "team"])
        st.sidebar.success(f"Snap counts saved for {NFL_TEAM} {int(season_input)}.")

st.sidebar.subheader("Color Codes (Auto)")
st.sidebar.caption("Green = better than NFL average; Red = worse (direction-aware).")

st.sidebar.markdown("---")
st.sidebar.subheader("Downloads")
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

pdf_stub_path = os.path.join(EXPORTS_DIR, f"{int(season_input)}_W{int(week_input):02d}_{opponent_input}_Final.pdf")
if os.path.exists(pdf_stub_path):
    with open(pdf_stub_path, "rb") as f:
        st.sidebar.download_button(
            label=f"Download Final PDF ({int(season_input)} W{int(week_input):02d})",
            data=f.read(),
            file_name=os.path.basename(pdf_stub_path),
            mime="application/pdf"
        )
else:
    st.sidebar.caption("Final PDF will appear here after you export it.")


# ==============================
# Main
# ==============================
st.title("Chicago Bears Weekly Tracker")

st.markdown(
    """
**Order:**  
1) Weekly Controls → 2) Upload Weekly Data → 3) NFL Averages → 4) Color Codes & Comparisons → 5) Opponent Preview → 6) Exports
"""
)

c1, c2, c3 = st.columns([1, 1, 2])
with c1:
    st.write(f"**Season:** {int(season_input)} | **Week:** {int(week_input)}")
with c2:
    st.write(f"**Opponent:** {opponent_input}")
with c3:
    st.caption("Use the sidebar to change Season, Week, and Opponent. These drive uploads and exports.")


# ==============================
# 1) Weekly Controls
# ==============================
with st.expander("1) Weekly Controls", expanded=True):
    cc1, cc2 = st.columns(2)
    with cc1:
        st.markdown("**Quick Fetch (This Week)**")
        if st.button("Fetch CHI Week Stats (Auto)"):
            season_val = int(season_input)
            week_val = int(week_input)

            # Diagnostics
            raw_weekly = import_df_safely(nfl.import_weekly_data, [season_val])
            raw_season = import_df_safely(nfl.import_seasonal_data, [season_val])
            st.caption(f"Diagnostics — weekly rows: {len(raw_weekly)} | seasonal rows: {len(raw_season)}")

            dfw = fetch_week_stats(season_val, week_val, NFL_TEAM)
            if dfw.empty:
                st.warning(
                    "No rows for this Season/Week/Team, and seasonal fallback was empty. "
                    "If the season isn't published yet, upload CSVs in section 2."
                )
            else:
                dfw["season"] = dfw.get("season", season_val)
                dfw["week"]   = dfw.get("week", week_val)
                dfw["team"]   = dfw.get("team", NFL_TEAM)
                dfw = dfw.loc[:, ~dfw.columns.duplicated()]

                _off = append_to_sheet(EXCEL_PATH, SHEET_OFFENSE, dfw,
                                       dedupe_on=["season", "week", "team"])
                _def = append_to_sheet(EXCEL_PATH, SHEET_DEFENSE, dfw,
                                       dedupe_on=["season", "week", "team"])

                st.success(
                    f"Saved {season_val} Week {week_val}: Offense rows={len(_off)} • Defense rows={len(_def)}"
                )
                st.dataframe(dfw.head(50), width="stretch")

    with cc2:
        st.markdown("**Notes / Key Items**")
        key_notes = st.text_area("Quick Notes (saved under OpponentPreview)", height=120)
        if st.button("Save Notes to Opponent Preview"):
            if key_notes.strip():
                row = pd.DataFrame([{
                    "season": int(season_input),
                    "week": int(week_input),
                    "team": NFL_TEAM,
                    "opponent": opponent_input,
                    "Notes": key_notes.strip(),
                    "saved_at": datetime.now().isoformat(timespec="seconds")
                }])
                append_to_sheet(EXCEL_PATH, SHEET_OPP_PREVIEW, row,
                                dedupe_on=["season", "week", "team", "opponent"])
                st.success("Notes saved.")
            else:
                st.info("Nothing to save.")


# ==============================
# 2) Upload Weekly Data
# ==============================
with st.expander("2) Upload Weekly Data", expanded=True):
    st.caption("Upload Offense/Defense/Personnel/SnapCounts CSVs. Rows are appended and deduped.")
    upc1, upc2 = st.columns(2)
    with upc1:
        st.markdown("**Offense CSV**")
        f_off = st.file_uploader("Upload Offense CSV", type=["csv"], key="up_off")
        if f_off is not None:
            try:
                df = pd.read_csv(f_off)
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
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
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
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
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
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
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.loc[:, ~df.columns.duplicated()]
                dedupe_cols = ["season", "week", "team"]
                if "player" in df.columns:
                    dedupe_cols.append("player")
                df_out = append_to_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS, df, dedupe_on=dedupe_cols)
                st.success(f"SnapCounts rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")


# ==============================
# 3) NFL Averages (Auto & Manual)
# ==============================
with st.expander("3) NFL Averages (Auto & Manual)", expanded=True):
    st.markdown("Use the sidebar ‘Fetch NFL Data (Auto)’ to compute YTD league averages (weeks 1..W).")

    off_view = read_sheet(EXCEL_PATH, SHEET_YTD_NFL_OFF)
    if not off_view.empty:
        st.write("**NFL YTD Averages (Auto)** — used for color coding & YTD comparisons")
        st.dataframe(off_view, width="stretch")
    else:
        st.info("No auto NFL YTD averages yet.")

    st.markdown("---")
    st.write("**NFL Weekly Averages (per-week, Auto)**")
    wkavg = read_sheet(EXCEL_PATH, SHEET_NFL_WEEKLY_AVG)
    if not wkavg.empty:
        st.dataframe(wkavg.tail(5), width="stretch")
    else:
        st.caption("Will appear after Fetch NFL Data (Auto) succeeds.")

    st.markdown("---")
    st.markdown("**Manual NFL Averages CSV** (optional). Columns will be prefixed to `NFL_Avg._` for styling.")
    f_nfl = st.file_uploader("Upload Manual NFL Averages CSV", type=["csv"], key="up_nflavg")
    if f_nfl is not None:
        try:
            df = pd.read_csv(f_nfl)
            df.columns = [c if c.startswith("NFL_Avg._") else f"NFL_Avg._{c}" for c in df.columns]
            df_out = append_to_sheet(EXCEL_PATH, SHEET_NFL_AVG_MANUAL, df, dedupe_on=None)
            st.success(f"Saved {len(df)} row(s) to NFL_Averages_Manual.")
            st.dataframe(df_out.tail(10), width="stretch")
        except Exception as e:
            st.error(f"Upload failed: {e}")


# ==============================
# 4) Color Codes & Comparisons
# ==============================
with st.expander("4) DVOA Proxy, Color Codes & Comparisons", expanded=True):
    tabs = st.tabs(["Bears vs Opponent (This Week)", "Bears YTD vs NFL (Weeks 1..W)"])

    # --- Tab 1: Week comparison vs opponent ---
    with tabs[0]:
        st.caption("Side-by-side comparison for the selected week using remote weekly data when available.")

        chi_wk, opp_wk = fetch_week_stats_both_teams(int(season_input), int(week_input), NFL_TEAM, opponent_input)
        if chi_wk.empty and opp_wk.empty:
            st.info("No weekly rows found for Bears/opponent this week from the source. If you have CSVs, upload them in Section 2.")
        else:
            # Align numeric columns
            nums_chi = select_numeric(chi_wk) if not chi_wk.empty else []
            nums_opp = select_numeric(opp_wk) if not opp_wk.empty else []
            common = sorted(list(set(nums_chi).intersection(nums_opp)))
            show_cols = ["season", "week", "team"] + common

            chi_show = chi_wk[show_cols].copy() if not chi_wk.empty else pd.DataFrame(columns=show_cols)
            opp_show = opp_wk[show_cols].copy() if not opp_wk.empty else pd.DataFrame(columns=show_cols)
            # Add suffixes for clarity
            chi_show = chi_show.add_suffix("_CHI")
            opp_show = opp_show.add_suffix("_OPP")

            merged = pd.concat([chi_show.reset_index(drop=True), opp_show.reset_index(drop=True)], axis=1)
            st.dataframe(merged, width="stretch")

        # Snap counts week comparison (best effort)
        st.markdown("**Snap Counts (Week comparison, if available)**")
        chi_sn, opp_sn = fetch_snap_counts_week_both(int(season_input), int(week_input), NFL_TEAM, opponent_input)
        if chi_sn.empty and opp_sn.empty:
            st.caption("No snap counts for this week from the source.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                st.write("Bears Snap Counts")
                st.dataframe(chi_sn.head(50), width="stretch")
            with col2:
                st.write(f"{opponent_input} Snap Counts")
                st.dataframe(opp_sn.head(50), width="stretch")

    # --- Tab 2: YTD Bears vs NFL (Weeks 1..W) ---
    with tabs[1]:
        st.caption("Bears averages across weeks 1..W vs NFL averages across the same weeks.")

        # Bears YTD (prefer source; fallback to workbook)
        bears_ytd = fetch_bears_averages_YTD_from_source(int(season_input), int(week_input), NFL_TEAM)
        if bears_ytd.empty:
            bears_ytd = fetch_bears_averages_YTD_from_workbook(int(week_input))

        nfl_ytd = fetch_nfl_averages_YTD(int(season_input), int(week_input))

        if bears_ytd.empty or nfl_ytd.empty:
            st.info("Need both Bears YTD and NFL YTD rows. Try Fetch NFL Data (Auto) and/or add weekly rows (auto or uploads).")
        else:
            # Merge bears_ytd with nfl_ytd (NFL columns get NFL_Avg._ prefix)
            merged = pd.concat([bears_ytd.reset_index(drop=True), nfl_ytd.reset_index(drop=True)], axis=1)

            # Try two contexts: Offense and Defense
            st.markdown("**Offense Context (direction-aware coloring vs NFL YTD)**")
            bh, bl = metric_orientation_for_context("OFFENSE")
            st.dataframe(style_vs_reference(merged, "NFL_Avg._", bh, bl), width="stretch")

            st.markdown("**Defense Context (direction-aware coloring vs NFL YTD)**")
            bh, bl = metric_orientation_for_context("DEFENSE")
            st.dataframe(style_vs_reference(merged, "NFL_Avg._", bh, bl), width="stretch")


# ==============================
# 5) Opponent Preview & Strategy Notes
# ==============================
with st.expander("5) Opponent Preview & Strategy Notes", expanded=True):
    st.caption("Notes and predictions you can save and review.")

    opp = read_sheet(EXCEL_PATH, SHEET_OPP_PREVIEW)
    if not opp.empty:
        st.write("**Opponent Preview (Recent)**")
        st.dataframe(opp.sort_values(opp.columns[0]).tail(25), width="stretch")
    else:
        st.info("No opponent preview entries yet.")

    st.markdown("**Weekly Prediction (optional)**")
    colA, colB = st.columns([2, 1])
    with colA:
        rationale = st.text_area("Rationale", height=120, key="pred_rationale")
    with colB:
        predicted_winner = st.selectbox("Predicted Winner", options=[NFL_TEAM, opponent_input], index=0)
        confidence = st.slider("Confidence", min_value=0, max_value=100, value=60, step=5)
    if st.button("Save Weekly Prediction"):
        row = pd.DataFrame([{
            "season": int(season_input),
            "week": int(week_input),
            "team": NFL_TEAM,
            "opponent": opponent_input,
            "Predicted_Winner": predicted_winner,
            "Confidence": confidence,
            "Rationale": rationale.strip(),
            "saved_at": datetime.now().isoformat(timespec="seconds")
        }])
        df_out = append_to_sheet(EXCEL_PATH, SHEET_PREDICTIONS, row,
                                 dedupe_on=["season", "week", "team", "opponent"])
        st.success("Prediction saved.")
        st.dataframe(df_out.tail(10), width="stretch")


# ==============================
# 6) Exports & Downloads
# ==============================
with st.expander("6) Exports & Downloads", expanded=True):
    st.caption("Create a quick PDF and download Excel/PDF here or from the sidebar.")
    if st.button("Export Final PDF (This Week)"):
        off_this = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
        def_this = read_sheet(EXCEL_PATH, SHEET_DEFENSE)

        lines = [
            f"Season {int(season_input)} — Week {int(week_input)} vs {opponent_input}",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            ""
        ]
        if not off_this.empty:
            offw = off_this[(off_this.get("season", int(season_input)) == int(season_input)) &
                            (off_this.get("week", int(week_input)) == int(week_input))]
            lines.append(f"Offense rows: {len(offw)}")
        if not def_this.empty:
            defw = def_this[(def_this.get("season", int(season_input)) == int(season_input)) &
                            (def_this.get("week", int(week_input)) == int(week_input))]
            lines.append(f"Defense rows: {len(defw)}")

        pdf_bytes = export_week_pdf(int(season_input), int(week_input), opponent_input, lines)
        if not pdf_bytes:
            st.error("FPDF not available — ensure fpdf is in requirements.")
        else:
            out_path = os.path.join(EXPORTS_DIR, f"{int(season_input)}_W{int(week_input):02d}_{opponent_input}_Final.pdf")
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

    st.markdown("**Workbook Peek**")
    cols = st.columns(3)
    try:
        with cols[0]:
            st.write("Offense")
            off = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
            st.dataframe(off.tail(10), width="stretch") if not off.empty else st.caption("—")
        with cols[1]:
            st.write("Defense")
            deff = read_sheet(EXCEL_PATH, SHEET_DEFENSE)
            st.dataframe(deff.tail(10), width="stretch") if not deff.empty else st.caption("—")
        with cols[2]:
            st.write("SnapCounts")
            snaps = read_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS)
            st.dataframe(snaps.tail(10), width="stretch") if not snaps.empty else st.caption("—")
    except Exception as e:
        st.error(f"Peek failed: {e}")

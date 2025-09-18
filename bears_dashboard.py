import os
import io
from datetime import datetime, date, timedelta

import streamlit as st
import pandas as pd
import numpy as np

from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from zipfile import BadZipFile

import nfl_data_py as nfl
from urllib.error import HTTPError, URLError

# --- App setup ---
st.set_page_config(page_title="Bears Weekly Tracker", layout="wide")

DATA_DIR = "./data"
EXPORTS_DIR = "./exports"
CSV_DIR = "./csv"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)
os.makedirs(CSV_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(DATA_DIR, "bears_weekly_analytics.xlsx")

# Sheet names
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
SHEET_NFL_WEEKLY_AVG = "NFL_Weekly_Averages"

ALL_SHEETS = [
    SHEET_OFFENSE, SHEET_DEFENSE, SHEET_PERSONNEL, SHEET_SNAP_COUNTS,
    SHEET_INJURIES, SHEET_MEDIA, SHEET_OPP_PREVIEW, SHEET_PREDICTIONS,
    SHEET_NFL_AVG_MANUAL, SHEET_YTD_TEAM_OFF, SHEET_YTD_TEAM_DEF,
    SHEET_YTD_NFL_OFF, SHEET_YTD_NFL_DEF, SHEET_NFL_WEEKLY_AVG
]

NFL_TEAM = "CHI"
DEFAULT_SEASON = 2025   # default season

# ---------------------------
# Robust import wrappers
# ---------------------------
def import_df_safely(import_fn, seasons):
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

# ---------------------------
# Excel helpers
# ---------------------------
def ensure_xlsx(path):
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

def read_sheet(path, sheet):
    ensure_xlsx(path)
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

def write_sheet(path, sheet, df):
    ensure_xlsx(path)
    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xlw:
            df.to_excel(xlw, index=False, sheet_name=sheet)
    except FileNotFoundError:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as xlw:
            df.to_excel(xlw, index=False, sheet_name=sheet)

def append_to_sheet(path, sheet, df_new, dedupe_on=None):
    df_old = read_sheet(path, sheet)
    if df_old.empty:
        df_out = df_new.copy()
    else:
        df_out = pd.concat([df_old, df_new], ignore_index=True)
    if dedupe_on:
        df_out = df_out.drop_duplicates(subset=dedupe_on, keep="last")
    write_sheet(path, sheet, df_out)
    return df_out

# ---------------------------
# CSV helpers
# ---------------------------
def csv_name(base, season, week, team):
    return os.path.join(CSV_DIR, f"{base}_{int(season)}_W{int(week):02d}_{team}.csv")

def save_csv(df, base, season, week, team):
    path = csv_name(base, season, week, team)
    os.makedirs(CSV_DIR, exist_ok=True)
    df.to_csv(path, index=False)
    return path

# ---------------------------
# NFL helpers (normalize, etc.)
# ---------------------------
def _normalize_team_col(df):
    if df.empty:
        return df
    if 'team' in df.columns:
        return df
    if 'recent_team' in df.columns:
        df = df.rename(columns={'recent_team': 'team'})
    if 'posteam' in df.columns:
        df = df.rename(columns={'posteam': 'team'})
    return df

def select_numeric(df):
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

# ---------------------------
# PBP fallback aggregation
# ---------------------------
def _aggregate_team_week_from_pbp(pbp: pd.DataFrame, team: str, week: int) -> pd.DataFrame:
    """Build a reasonable weekly team line from pbp when weekly is unavailable."""
    if pbp.empty:
        return pd.DataFrame()
    # Filter to week
    if 'week' in pbp.columns:
        pbp = pbp[pbp['week'] == week]
    if pbp.empty:
        return pd.DataFrame()

    # Team offensive plays only (posteam == team)
    off = pbp[pbp.get('posteam') == team].copy()
    if off.empty:
        # Some rows (like penalties) may not have posteam; fallback to any team column
        off = pbp[(pbp.get('posteam') == team) | (pbp.get('defteam') != team)]
    off['is_pass'] = off.get('pass', 0).fillna(0).astype(int)
    off['is_rush'] = off.get('rush', 0).fillna(0).astype(int)

    # Basic aggregations
    plays = len(off)
    yards = float(off.get('yards_gained', 0).fillna(0).sum())
    pass_att = int(off.get('pass_attempt', 0).fillna(0).sum())
    completions = int(off.get('complete_pass', 0).fillna(0).sum())
    pass_yards = float(off.loc[off['is_pass'] == 1, 'yards_gained'].fillna(0).sum())
    rush_att = int(off.get('rush_attempt', 0).fillna(0).sum())
    rush_yards = float(off.loc[off['is_rush'] == 1, 'yards_gained'].fillna(0).sum())
    sacks = int(off.get('sack', 0).fillna(0).sum())
    interceptions = int(off.get('interception', 0).fillna(0).sum())
    fumbles_lost = int(off.get('fumble_lost', 0).fillna(0).sum())

    # 3rd down
    td_att = int(off.get('third_down_attempt', 0).fillna(0).sum())
    td_conv = int(off.get('third_down_converted', 0).fillna(0).sum())

    row = {
        "team": team,
        "week": week,
        "plays": plays,
        "yards_gained": yards,
        "pass_att": pass_att,
        "completions": completions,
        "pass_yards": pass_yards,
        "rush_att": rush_att,
        "rush_yards": rush_yards,
        "sacks": sacks,
        "interceptions": interceptions,
        "fumbles_lost": fumbles_lost,
        "third_down_att": td_att,
        "third_down_conv": td_conv,
    }
    return pd.DataFrame([row])

def _pbp_week_for_season(season: int) -> pd.DataFrame:
    pbp = import_df_safely(nfl.import_pbp_data, [season])
    # Some versions use 'season_type'; keep REG by default if present
    if not pbp.empty and 'season_type' in pbp.columns:
        pbp = pbp[pbp['season_type'].isin(['REG', 'regular', 'Regular'])]  # be defensive
    return pbp

def fetch_week_stats_via_pbp(season: int, week: int, team_abbr: str) -> pd.DataFrame:
    pbp = _pbp_week_for_season(season)
    if pbp.empty:
        return pd.DataFrame()
    return _aggregate_team_week_from_pbp(pbp, team_abbr, week)

def fetch_week_stats_both_via_pbp(season: int, week: int, team_abbr: str, opp_abbr: str):
    pbp = _pbp_week_for_season(season)
    if pbp.empty:
        return pd.DataFrame(), pd.DataFrame()
    return (
        _aggregate_team_week_from_pbp(pbp, team_abbr, week),
        _aggregate_team_week_from_pbp(pbp, opp_abbr, week)
    )

def nfl_week_average_via_pbp(season: int, week: int) -> pd.DataFrame:
    pbp = _pbp_week_for_season(season)
    if pbp.empty:
        return pd.DataFrame()
    teams = sorted(set(list(pbp.get('posteam').dropna().unique())))
    rows = []
    for t in teams:
        r = _aggregate_team_week_from_pbp(pbp, t, week)
        if not r.empty:
            rows.append(r)
    if not rows:
        return pd.DataFrame()
    wk = pd.concat(rows, ignore_index=True)
    nums = select_numeric(wk)
    out = wk[nums].mean(numeric_only=True).to_frame().T
    out.columns = [f"NFL_Avg._{c}" for c in out.columns]
    return out.reset_index(drop=True)

# ---------------------------
# Primary weekly fetch (tries weekly, then PBP)
# ---------------------------
def fetch_week_stats(season, week, team_abbr):
    # Try official weekly
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    weekly = _normalize_team_col(weekly)
    if not weekly.empty and 'week' in weekly.columns:
        dfw = weekly[(weekly['team'] == team_abbr) & (weekly['week'] == week)].copy()
        if not dfw.empty:
            return dfw
    # Fallback to PBP aggregation
    return fetch_week_stats_via_pbp(season, week, team_abbr)

def fetch_week_stats_both_teams(season, week, team_abbr, opp_abbr):
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    weekly = _normalize_team_col(weekly)
    if not weekly.empty and 'week' in weekly.columns:
        chi = weekly[(weekly['team'] == team_abbr) & (weekly['week'] == week)].copy()
        opp = weekly[(weekly['team'] == opp_abbr) & (weekly['week'] == week)].copy()
        if not chi.empty or not opp.empty:
            return chi, opp
    # Fallback to PBP aggregation for both
    return fetch_week_stats_both_via_pbp(season, week, team_abbr, opp_abbr)

def nfl_week_average(season, week):
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    if not weekly.empty and 'week' in weekly.columns:
        weekly = _normalize_team_col(weekly)
        wk = weekly[weekly['week'] == week].copy()
        if not wk.empty:
            nums = select_numeric(wk)
            out = wk[nums].mean(numeric_only=True).to_frame().T
            out.columns = [f"NFL_Avg._{c}" for c in out.columns]
            return out.reset_index(drop=True)
    # Fallback to PBP league mean
    return nfl_week_average_via_pbp(season, week)

# ---------------------------
# YTD helpers
# ---------------------------
def fetch_nfl_averages_YTD(season, upto_week):
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    weekly = _normalize_team_col(weekly)
    if not weekly.empty and 'week' in weekly.columns:
        df = weekly[(weekly['week'] >= 1) & (weekly['week'] <= upto_week)].copy()
        if not df.empty:
            nums = select_numeric(df)
            nfl_ytd = df[nums].mean(numeric_only=True).to_frame().T
            nfl_ytd.columns = [f"NFL_Avg._{c}" for c in nfl_ytd.columns]
            return nfl_ytd.reset_index(drop=True)
    # PBP fallback YTD: average of per-team, per-week aggregates 1..W
    pbp = _pbp_week_for_season(season)
    if pbp.empty:
        return pd.DataFrame()
    rows = []
    teams = sorted(set(list(pbp.get('posteam').dropna().unique())))
    for w in range(1, int(upto_week) + 1):
        for t in teams:
            r = _aggregate_team_week_from_pbp(pbp, t, w)
            if not r.empty:
                rows.append(r)
    if not rows:
        return pd.DataFrame()
    big = pd.concat(rows, ignore_index=True)
    nums = select_numeric(big)
    out = big[nums].mean(numeric_only=True).to_frame().T
    out.columns = [f"NFL_Avg._{c}" for c in out.columns]
    return out.reset_index(drop=True)

def fetch_bears_averages_YTD_from_source(season, upto_week, team):
    weekly = import_df_safely(nfl.import_weekly_data, [season])
    weekly = _normalize_team_col(weekly)
    if not weekly.empty and 'week' in weekly.columns:
        chi = weekly[(weekly['team'] == team) & (weekly['week'] >= 1) & (weekly['week'] <= upto_week)].copy()
        if not chi.empty:
            nums = select_numeric(chi)
            brs = chi[nums].mean(numeric_only=True).to_frame().T
            return brs.reset_index(drop=True)
    # PBP fallback YTD: aggregate each week for CHI
    pbp = _pbp_week_for_season(season)
    if pbp.empty:
        return pd.DataFrame()
    rows = []
    for w in range(1, int(upto_week) + 1):
        r = _aggregate_team_week_from_pbp(pbp, team, w)
        if not r.empty:
            rows.append(r)
    if not rows:
        return pd.DataFrame()
    big = pd.concat(rows, ignore_index=True)
    nums = select_numeric(big)
    brs = big[nums].mean(numeric_only=True).to_frame().T
    return brs.reset_index(drop=True)

def fetch_bears_averages_YTD_from_workbook(upto_week):
    off = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
    deff = read_sheet(EXCEL_PATH, SHEET_DEFENSE)
    df = pd.concat([off, deff], ignore_index=True) if (not off.empty or not deff.empty) else pd.DataFrame()
    if df.empty or 'week' not in df.columns:
        return pd.DataFrame()
    df = df[(df['week'] >= 1) & (df['week'] <= upto_week)].copy()
    nums = select_numeric(df)
    if not nums:
        return pd.DataFrame()
    brs = df[nums].mean(numeric_only=True).to_frame().T
    return brs.reset_index(drop=True)

# ---------------------------
# Snap counts & injuries
# ---------------------------
def fetch_snap_counts(season, team):
    snaps = import_df_safely(nfl.import_snap_counts, [season])
    if snaps.empty:
        return pd.DataFrame()
    snaps = _normalize_team_col(snaps)
    return snaps[snaps["team"] == team].copy()

def fetch_snap_counts_week_both(season, week, team, opp):
    snaps = import_df_safely(nfl.import_snap_counts, [season])
    if snaps.empty:
        return pd.DataFrame(), pd.DataFrame()
    snaps = _normalize_team_col(snaps)
    has_week = 'week' in snaps.columns
    chi = snaps[(snaps["team"] == team) & ((snaps["week"] == week) if has_week else True)].copy()
    oppdf = snaps[(snaps["team"] == opp) & ((snaps["week"] == week) if has_week else True)].copy()
    return chi, oppdf

def _week_date_bounds_from_schedule(season: int, week: int):
    sched = import_df_safely(nfl.import_schedules, [season])
    if sched.empty:
        return None, None
    dfw = sched[sched['week'] == week].copy() if 'week' in sched.columns else sched.copy()
    if dfw.empty:
        return None, None
    # Use min/max of game_date/timestamp columns available
    for date_col in ['gameday', 'game_date', 'start_time', 'game_time']:
        if date_col in dfw.columns:
            try:
                dts = pd.to_datetime(dfw[date_col], errors='coerce', utc=True)
                lo = dts.min()
                hi = dts.max()
                if pd.notna(lo) and pd.notna(hi):
                    # widen a little
                    return (lo - pd.Timedelta(days=1)), (hi + pd.Timedelta(days=2))
            except Exception:
                continue
    return None, None

def fetch_injuries_for_week(season: int, week: int, team: str) -> pd.DataFrame:
    inj = import_df_safely(nfl.import_injuries, [season])
    if inj.empty:
        return pd.DataFrame()
    inj = _normalize_team_col(inj)
    if 'week' in inj.columns:
        df = inj[(inj['team'] == team) & (inj['week'] == week)].copy()
        return df
    # fallback: filter by date window around the games in that week
    lo, hi = _week_date_bounds_from_schedule(season, week)
    if lo is not None and 'report_date' in inj.columns:
        dts = pd.to_datetime(inj['report_date'], errors='coerce', utc=True)
        df = inj[(inj['team'] == team) & (dts >= lo) & (dts <= hi)].copy()
        return df
    return inj[inj['team'] == team].copy()

# ---------------------------
# PDF export
# ---------------------------
try:
    from fpdf import FPDF
except Exception:
    FPDF = None

def export_week_pdf(season, week, opponent, summary_lines):
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

HEADLINE_KEYWORDS = [
    "point", "pts", "yard", "total_yards", "ypa", "ypg", "yac",
    "ypc", "rush_yards", "rushing_yards", "cmp", "completion",
    "passer_rating", "qbr", "sack", "pressure", "turnover", "to",
    "int", "fumble", "3d", "third_down", "rz", "red_zone",
]

def _first_n_common_numeric(chi, opp, n=6):
    if chi.empty or opp.empty:
        return []
    nums_chi = [c for c in chi.columns if pd.api.types.is_numeric_dtype(chi[c])]
    nums_opp = [c for c in opp.columns if pd.api.types.is_numeric_dtype(opp[c])]
    common = [c for c in nums_chi if c in nums_opp]
    def score(col):
        lc = col.lower()
        return max((1 if kw in lc else 0) for kw in HEADLINE_KEYWORDS)
    common.sort(key=lambda c: (-score(c), c.lower()))
    return common[:n]

def _safe_get_scalar(df, col):
    if df.empty or col not in df.columns:
        return None
    try:
        v = float(df[col].iloc[0])
        if np.isfinite(v):
            return v
    except Exception:
        pass
    return None

def _fmt(v):
    return "-" if v is None else (f"{v:.2f}" if abs(v) < 1000 else f"{v:,.0f}")

def build_pdf_summary_lines(season, week, opponent):
    lines = []
    chi_wk, opp_wk = fetch_week_stats_both_teams(season, week, NFL_TEAM, opponent)
    nfl_wk = nfl_week_average(season, week)
    bears_ytd = fetch_bears_averages_YTD_from_source(season, week, NFL_TEAM)
    if bears_ytd.empty:
        bears_ytd = fetch_bears_averages_YTD_from_workbook(week)
    nfl_ytd = fetch_nfl_averages_YTD(season, week)

    lines.append(f"Season {season} — Week {week} vs {opponent}")
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    lines.append("Bears vs Opponent — This Week")
    if chi_wk.empty and opp_wk.empty:
        lines.append("  • Weekly team rows not available from source/PBP.")
    else:
        cols = _first_n_common_numeric(chi_wk, opp_wk, n=6)
        if not cols:
            lines.append("  • No common numeric columns detected for this week.")
        else:
            for c in cols:
                chi_v = _safe_get_scalar(chi_wk, c)
                opp_v = _safe_get_scalar(opp_wk, c)
                diff = None if (chi_v is None or opp_v is None) else chi_v - opp_v
                lines.append(f"  • {c}: CHI {_fmt(chi_v)} | OPP {_fmt(opp_v)} | Δ {_fmt(diff)}")
    lines.append("")
    lines.append("Bears vs NFL Average — This Week")
    if chi_wk.empty or nfl_wk.empty:
        lines.append("  • Week average not available (missing CHI row or NFL week mean).")
    else:
        nfl_map = {k.replace("NFL_Avg._", ""): k for k in nfl_wk.columns if k.startswith("NFL_Avg._")}
        nums_chi = [c for c in chi_wk.columns if pd.api.types.is_numeric_dtype(chi_wk[c])]
        both = [c for c in nums_chi if c in nfl_map]
        def score(col):
            lc = col.lower()
            return max((1 if kw in lc else 0) for kw in HEADLINE_KEYWORDS)
        both.sort(key=lambda c: (-score(c), c.lower()))
        for c in both[:6]:
            chi_v = _safe_get_scalar(chi_wk, c)
            nfl_v = _safe_get_scalar(nfl_wk, nfl_map[c])
            diff = None if (chi_v is None or nfl_v is None) else chi_v - nfl_v
            lines.append(f"  • {c}: CHI {_fmt(chi_v)} | NFLwk {_fmt(nfl_v)} | Δ {_fmt(diff)}")
    lines.append("")
    lines.append("Bears YTD vs NFL YTD — Weeks 1..W")
    if bears_ytd.empty or nfl_ytd.empty:
        lines.append("  • YTD rows not available yet.")
    else:
        nfl_map_ytd = {k.replace("NFL_Avg._", ""): k for k in nfl_ytd.columns if k.startswith("NFL_Avg._")}
        nums_brs = [c for c in bears_ytd.columns if pd.api.types.is_numeric_dtype(bears_ytd[c])]
        both_ytd = [c for c in nums_brs if c in nfl_map_ytd]
        def score2(col):
            lc = col.lower()
            return max((1 if kw in lc else 0) for kw in HEADLINE_KEYWORDS)
        both_ytd.sort(key=lambda c: (-score2(c), c.lower()))
        for c in both_ytd[:6]:
            brs_v = _safe_get_scalar(bears_ytd, c)
            nfl_v = _safe_get_scalar(nfl_ytd, nfl_map_ytd[c])
            diff = None if (brs_v is None or nfl_v is None) else brs_v - nfl_v
            lines.append(f"  • {c}: CHI_YTD {_fmt(brs_v)} | NFL_YTD {_fmt(nfl_v)} | Δ {_fmt(diff)}")
    lines.append("")
    return lines

# ---------------------------
# Sidebar
# ---------------------------
st.sidebar.title("Controls")
season_input = st.sidebar.number_input("Season", min_value=2012, max_value=2030, value=DEFAULT_SEASON, step=1)
week_input = st.sidebar.number_input("Week", min_value=1, max_value=22, value=2, step=1)
opponent_input = st.sidebar.text_input("Opponent (3-letter code)", value="MIN").strip().upper()

st.sidebar.markdown("---")
st.sidebar.subheader("NFL Updates")
if st.sidebar.button("Fetch NFL Data (Auto)"):
    nfl_ytd = fetch_nfl_averages_YTD(int(season_input), int(week_input))
    if nfl_ytd.empty:
        st.sidebar.error("Could not compute NFL YTD averages (source/PBP unavailable).")
    else:
        write_sheet(EXCEL_PATH, SHEET_YTD_NFL_OFF, nfl_ytd)
        write_sheet(EXCEL_PATH, SHEET_YTD_NFL_DEF, nfl_ytd)
        st.sidebar.success("NFL YTD averages saved.")
    wkavg = nfl_week_average(int(season_input), int(week_input))
    if not wkavg.empty:
        write_sheet(EXCEL_PATH, SHEET_NFL_WEEKLY_AVG, wkavg.assign(_w=int(week_input)))
        st.sidebar.success("NFL weekly average (selected week) saved.")

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

st.sidebar.subheader("Injuries")
if st.sidebar.button("Fetch Injuries (Auto)"):
    inj = fetch_injuries_for_week(int(season_input), int(week_input), NFL_TEAM)
    if inj.empty:
        st.sidebar.warning("No injury rows found for this week/team.")
    else:
        append_to_sheet(EXCEL_PATH, SHEET_INJURIES, inj,
                        dedupe_on=[c for c in ["season", "week", "team", "player", "report_date"] if c in inj.columns])
        st.sidebar.success(f"Injuries saved for {NFL_TEAM} W{int(week_input)}.")

st.sidebar.subheader("CSV Outputs")
if st.sidebar.button("Write Weekly CSVs (Auto)"):
    season_val = int(season_input); week_val = int(week_input)
    chi_wk, opp_wk = fetch_week_stats_both_teams(season_val, week_val, NFL_TEAM, opponent_input)
    if chi_wk.empty and opp_wk.empty:
        st.sidebar.error("No weekly rows available (weekly + PBP both empty).")
    else:
        if not chi_wk.empty:
            path_off = save_csv(chi_wk.copy(), "Offense", season_val, week_val, NFL_TEAM)
            path_def = save_csv(chi_wk.copy(), "Defense", season_val, week_val, NFL_TEAM)
            st.sidebar.success(f"Wrote Bears weekly CSVs:\n• {os.path.basename(path_off)}\n• {os.path.basename(path_def)}")
        if not opp_wk.empty:
            path_off_o = save_csv(opp_wk.copy(), "Offense", season_val, week_val, opponent_input)
            path_def_o = save_csv(opp_wk.copy(), "Defense", season_val, week_val, opponent_input)
            st.sidebar.success(f"Wrote Opponent weekly CSVs:\n• {os.path.basename(path_off_o)}\n• {os.path.basename(path_def_o)}")

if st.sidebar.button("Write SnapCount CSVs (Auto)"):
    season_val = int(season_input); week_val = int(week_input)
    chi_sn, opp_sn = fetch_snap_counts_week_both(season_val, week_val, NFL_TEAM, opponent_input)
    wrote_any = False
    if not chi_sn.empty:
        p = save_csv(chi_sn.copy(), "SnapCounts", season_val, week_val, NFL_TEAM)
        st.sidebar.success(f"Wrote Bears SnapCounts CSV: {os.path.basename(p)}")
        wrote_any = True
    if not opp_sn.empty:
        p = save_csv(opp_sn.copy(), "SnapCounts", season_val, week_val, opponent_input)
        st.sidebar.success(f"Wrote Opponent SnapCounts CSV: {os.path.basename(p)}")
        wrote_any = True
    if not wrote_any:
        st.sidebar.error("No snap counts available from source/PBP to write.")

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
    st.sidebar.info("Excel not found yet; upload or fetch to create it.")

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

# ---------------------------
# Main
# ---------------------------
st.title("Chicago Bears Weekly Tracker")
st.markdown("**Order:** 1) Weekly Controls → 2) Upload Weekly Data → 3) NFL Averages → 4) Color Codes & Comparisons → 5) Opponent Preview → 6) Exports")

c1, c2, c3 = st.columns([1, 1, 2])
with c1:
    st.write(f"**Season:** {int(season_input)} | **Week:** {int(week_input)}")
with c2:
    st.write(f"**Opponent:** {opponent_input}")
with c3:
    st.caption("Use the sidebar to change Season, Week, and Opponent. These drive uploads and exports.")

# 1) Weekly Controls
with st.expander("1) Weekly Controls", expanded=True):
    st.info("Weekly fetch tries official weekly first; if unavailable, it **automatically builds from play-by-play**.")
    cc1, cc2 = st.columns(2)
    with cc1:
        st.markdown("**Quick Fetch (This Week)**")
        if st.button("Fetch & Append CHI Week to Workbook (Auto+PBP Fallback)"):
            season_val = int(season_input)
            week_val = int(week_input)
            dfw = fetch_week_stats(season_val, week_val, NFL_TEAM)
            if dfw.empty:
                st.warning("No CHI weekly/PBP rows for this week yet.")
            else:
                dfw["season"] = dfw.get("season", season_val)
                dfw["week"]   = dfw.get("week", week_val)
                dfw["team"]   = dfw.get("team", NFL_TEAM)
                dfw = dfw.loc[:, ~dfw.columns.duplicated()]
                _off = append_to_sheet(EXCEL_PATH, SHEET_OFFENSE, dfw, dedupe_on=["season", "week", "team"])
                _def = append_to_sheet(EXCEL_PATH, SHEET_DEFENSE, dfw, dedupe_on=["season", "week", "team"])
                path_off = save_csv(dfw, "Offense", season_val, week_val, NFL_TEAM)
                path_def = save_csv(dfw, "Defense", season_val, week_val, NFL_TEAM)
                st.success(f"Workbook updated. CSVs written:\n• {os.path.basename(path_off)}\n• {os.path.basename(path_def)}")
                st.dataframe(dfw.head(50), width="stretch")

        if st.button("Fetch & Write Opponent Week CSV (Auto+PBP Fallback)"):
            season_val = int(season_input)
            week_val = int(week_input)
            opp_df = fetch_week_stats(season_val, week_val, opponent_input)
            if opp_df.empty:
                st.warning("No opponent weekly/PBP rows yet.")
            else:
                opp_df["season"] = opp_df.get("season", season_val)
                opp_df["week"]   = opp_df.get("week", week_val)
                opp_df["team"]   = opp_df.get("team", opponent_input)
                opp_df = opp_df.loc[:, ~opp_df.columns.duplicated()]
                p1 = save_csv(opp_df, "Offense", season_val, week_val, opponent_input)
                p2 = save_csv(opp_df, "Defense", season_val, week_val, opponent_input)
                st.success(f"Opponent CSVs written:\n• {os.path.basename(p1)}\n• {os.path.basename(p2)}")

        if st.button("Fetch & Write SnapCount CSVs (Auto)"):
            season_val = int(season_input); week_val = int(week_input)
            chi_sn, opp_sn = fetch_snap_counts_week_both(season_val, week_val, NFL_TEAM, opponent_input)
            wrote = False
            if not chi_sn.empty:
                p = save_csv(chi_sn, "SnapCounts", season_val, week_val, NFL_TEAM)
                st.success(f"Bears SnapCounts CSV: {os.path.basename(p)}")
                wrote = True
            if not opp_sn.empty:
                p = save_csv(opp_sn, "SnapCounts", season_val, week_val, opponent_input)
                st.success(f"Opponent SnapCounts CSV: {os.path.basename(p)}")
                wrote = True
            if not wrote:
                st.warning("No snap counts available to write yet.")

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

# 2) Upload Weekly Data
with st.expander("2) Upload Weekly Data", expanded=True):
    st.caption("Upload Offense/Defense/Personnel/SnapCounts/Injuries CSVs. Rows are appended and deduped.")
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

        st.markdown("**Injuries CSV**")
        f_inj = st.file_uploader("Upload Injuries CSV", type=["csv"], key="up_inj")
        if f_inj is not None:
            try:
                df = pd.read_csv(f_inj)
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_INJURIES, df,
                                         dedupe_on=[c for c in ["season", "week", "team", "player"] if c in df.columns])
                st.success(f"Injuries rows now: {len(df_out)}")
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

# 3) NFL Averages (Auto & Manual)
with st.expander("3) NFL Averages (Auto & Manual)", expanded=True):
    st.markdown("Use the sidebar ‘Fetch NFL Data (Auto)’ to compute YTD league averages (weeks 1..W).")
    off_view = read_sheet(EXCEL_PATH, SHEET_YTD_NFL_OFF)
    if not off_view.empty:
        st.write("**NFL YTD Averages (Auto)**")
        st.dataframe(off_view, width="stretch")
    else:
        st.info("No auto NFL YTD averages yet (weekly & PBP may still be empty).")

    st.markdown("---")
    st.write("**NFL Weekly Averages (Selected Week, Auto or PBP)**")
    wkavg = read_sheet(EXCEL_PATH, SHEET_NFL_WEEKLY_AVG)
    if not wkavg.empty:
        st.dataframe(wkavg.tail(5), width="stretch")
    else:
        st.caption("Will appear after ‘Fetch NFL Data (Auto)’.")

    st.markdown("---")
    st.write("**Selected Week NFL Average (live)**")
    nfl_week_avg_row = nfl_week_average(int(season_input), int(week_input))
    if nfl_week_avg_row.empty:
        st.caption("No league average for this week yet.")
    else:
        st.dataframe(nfl_week_avg_row, width="stretch")

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

# 4) Color Codes & Comparisons
with st.expander("4) DVOA Proxy, Color Codes & Comparisons", expanded=True):
    tabs = st.tabs(["Bears vs Opponent (This Week)", "Bears YTD vs NFL (Weeks 1..W)"])

    with tabs[0]:
        st.caption("Uses weekly when available, otherwise PBP aggregates.")
        chi_wk, opp_wk = fetch_week_stats_both_teams(int(season_input), int(week_input), NFL_TEAM, opponent_input)
        if chi_wk.empty and opp_wk.empty:
            st.info("No weekly/PBP rows found for Bears/opponent this week yet.")
        else:
            nums_chi = select_numeric(chi_wk) if not chi_wk.empty else []
            nums_opp = select_numeric(opp_wk) if not opp_wk.empty else []
            common = sorted(list(set(nums_chi).intersection(nums_opp)))
            show_cols = ["season", "week", "team"] + common if common else list(chi_wk.columns)
            chi_show = chi_wk[show_cols].copy() if not chi_wk.empty else pd.DataFrame(columns=show_cols)
            opp_show = opp_wk[show_cols].copy() if not opp_wk.empty else pd.DataFrame(columns=show_cols)
            chi_show = chi_show.add_suffix("_CHI")
            opp_show = opp_show.add_suffix("_OPP")
            merged = pd.concat([chi_show.reset_index(drop=True), opp_show.reset_index(drop=True)], axis=1)
            st.dataframe(merged, width="stretch")

        st.markdown("**Bears vs NFL Average — This Week**")
        if not chi_wk.empty:
            if nfl_week_avg_row is None or nfl_week_avg_row.empty:
                st.caption("NFL week average unavailable yet.")
            else:
                ids = [c for c in chi_wk.columns if c in ("season", "week", "team")]
                nums = select_numeric(chi_wk)
                preview = pd.concat([chi_wk[ids + nums].reset_index(drop=True),
                                     nfl_week_avg_row.reset_index(drop=True)], axis=1)
                def _cell_color(val, base_col):
                    nfl_col = f"NFL_Avg._{base_col}"
                    if nfl_col not in preview.columns:
                        return ""
                    ref = preview[nfl_col].iloc[0]
                    if pd.isna(val) or pd.isna(ref):
                        return ""
                    return "background-color: #e5ffe5" if val > ref else "background-color: #ffe5e5"
                def _apply(row):
                    styles = []
                    for c in preview.columns:
                        if c in ids or c.startswith("NFL_Avg._"):
                            styles.append("")
                        else:
                            styles.append(_cell_color(row[c], c))
                    return styles
                st.dataframe(preview.style.apply(_apply, axis=1), width="stretch")
        else:
            st.caption("No CHI row to compare for this week.")

        # Bears-only Snap Counts view
        st.markdown("**Snap Counts (This Week — Bears only)**")
        chi_sn, _ = fetch_snap_counts_week_both(int(season_input), int(week_input), NFL_TEAM, opponent_input)
        if chi_sn.empty:
            st.caption("No CHI snap counts this week yet.")
        else:
            st.dataframe(chi_sn.head(50), width="stretch")

        st.markdown("**This Week Snap Counts (center display)**")
        chi_sn_all = read_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS)
        chi_sn_week = chi_sn_all.copy()
        if not chi_sn_week.empty:
            if 'season' in chi_sn_week.columns:
                chi_sn_week = chi_sn_week[chi_sn_week['season'] == int(season_input)]
            if 'week' in chi_sn_week.columns:
                chi_sn_week = chi_sn_week[chi_sn_week['week'] == int(week_input)]
        if chi_sn_week.empty:
            st.caption("No saved CHI snap counts for this week yet.")
        else:
            st.dataframe(chi_sn_week.head(50), width="stretch")

    with tabs[1]:
        st.caption("Bears averages across weeks 1..W vs NFL averages across the same weeks (weekly or PBP).")
        bears_ytd = fetch_bears_averages_YTD_from_source(int(season_input), int(week_input), NFL_TEAM)
        if bears_ytd.empty:
            bears_ytd = fetch_bears_averages_YTD_from_workbook(int(week_input))
        nfl_ytd = fetch_nfl_averages_YTD(int(season_input), int(week_input))
        if bears_ytd.empty or nfl_ytd.empty:
            st.info("Need both Bears YTD and NFL YTD rows (try ‘Fetch NFL Data (Auto)’ and/or Weekly Fetch).")
        else:
            merged = pd.concat([bears_ytd.reset_index(drop=True), nfl_ytd.reset_index(drop=True)], axis=1)
            st.dataframe(merged, width="stretch")

# 5) Opponent Preview & Strategy Notes
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

# 6) Exports & Downloads
with st.expander("6) Exports & Downloads", expanded=True):
    st.caption("Create week-only Excel, final PDF, and download the full workbook.")

    st.markdown("**Weekly Excel (This Week)**")
    if st.button("Create Weekly Excel (Off/Def/Personnel/Snaps/Injuries)"):
        season_val = int(season_input); week_val = int(week_input)
        off = read_sheet(EXCEL_PATH, SHEET_OFFENSE)
        deff = read_sheet(EXCEL_PATH, SHEET_DEFENSE)
        per = read_sheet(EXCEL_PATH, SHEET_PERSONNEL)
        snaps = read_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS)
        inj = read_sheet(EXCEL_PATH, SHEET_INJURIES)

        def fw(df):
            if df.empty: return df
            if 'season' in df.columns:
                df = df[df['season'] == season_val]
            if 'week' in df.columns:
                df = df[df['week'] == week_val]
            return df

        off = fw(off); deff = fw(deff); per = fw(per); snaps = fw(snaps); inj = fw(inj)
        chi_wk, opp_wk = fetch_week_stats_both_teams(season_val, week_val, NFL_TEAM, opponent_input)
        nfl_wk = nfl_week_average(season_val, week_val)

        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as xlw:
            if not off.empty:   off.to_excel(xlw, index=False, sheet_name="Offense")
            if not deff.empty:  deff.to_excel(xlw, index=False, sheet_name="Defense")
            if not per.empty:   per.to_excel(xlw, index=False, sheet_name="Personnel")
            if not snaps.empty: snaps.to_excel(xlw, index=False, sheet_name="SnapCounts")
            if not inj.empty:   inj.to_excel(xlw, index=False, sheet_name="Injuries")
            if not chi_wk.empty: chi_wk.to_excel(xlw, index=False, sheet_name="CHI_Week")
            if not opp_wk.empty: opp_wk.to_excel(xlw, index=False, sheet_name=f"{opponent_input}_Week")
            if nfl_wk is not None and not nfl_wk.empty: nfl_wk.to_excel(xlw, index=False, sheet_name="NFL_Week_Avg")
        bio.seek(0)
        st.download_button(
            label=f"Download Weekly Excel — {int(season_input)} W{int(week_input):02d}",
            data=bio.getvalue(),
            file_name=f"bears_week_{int(season_input)}_W{int(week_input):02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    st.markdown("---")
    if st.button("Export Final PDF (This Week)"):
        try:
            lines = build_pdf_summary_lines(int(season_input), int(week_input), opponent_input)
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
        except Exception as e:
            st.error(f"PDF export failed: {e}")

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
        st.info("Excel not found yet; upload or fetch to create it.")

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

        st.write("Injuries")
        inj = read_sheet(EXCEL_PATH, SHEET_INJURIES)
        st.dataframe(inj.tail(10), width="stretch") if not inj.empty else st.caption("—")
    except Exception as e:
        st.error(f"Peek failed: {e}")

# bears_dashboard.py
# Chicago Bears 2025–26 Weekly Tracker (Streamlit)
# Full app with NFL tools (manual + auto fetch), uploads, proxy, media/injuries/opponent preview,
# predictions, YTD Team vs NFL Avg (colorized), weekly previews, and Pre/Post/Final Excel/PDF exports.
# + Additions: Expanded NFL YTD Averages (Offense & Defense), Snap Counts Auto-Fetch button,
#   and PDF exports that include the YTD league-average sections.

import os
from datetime import datetime
import pandas as pd
import streamlit as st

# ---- Optional extras (present in requirements) ----
try:
    import openpyxl
    from openpyxl.formatting.rule import ColorScaleRule
except Exception:
    openpyxl = None

# ---- Optional PDF (reportlab) ----
HAS_REPORTLAB = True
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors as RL_COLORS
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
except Exception:
    HAS_REPORTLAB = False

# ========= Streamlit config =========
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# Display order + “which direction is good”
METRIC_SCHEMA = {
    "YPA": {"higher_is_better": True},
    "CMP%": {"higher_is_better": True},
    "RZ% Allowed": {"higher_is_better": False},
    "SACKs": {"higher_is_better": True},
    "INTs": {"higher_is_better": True},
    "QB Hits": {"higher_is_better": True},
    "Pressures": {"higher_is_better": True},
    "3D% Allowed": {"higher_is_better": False},
}

# ========= Excel helpers =========
def _ensure_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as w:
            pd.DataFrame(columns=["Week"]).to_excel(w, "Offense", index=False)
            pd.DataFrame(columns=["Week"]).to_excel(w, "Defense", index=False)
            pd.DataFrame(columns=["Week"]).to_excel(w, "Personnel", index=False)
            pd.DataFrame(columns=["Week"]).to_excel(w, "SnapCounts", index=False)
            pd.DataFrame(columns=["Week","Injury","Status","Notes"]).to_excel(w, "Injuries", index=False)
            pd.DataFrame(columns=["Week","Source","Summary"]).to_excel(w, "MediaSummaries", index=False)
            pd.DataFrame(columns=["Week","Notes"]).to_excel(w, "OpponentPreview", index=False)
            pd.DataFrame(columns=["Week","Rationale","Prediction"]).to_excel(w, "Predictions", index=False)
            # Manual league averages (optional)
            pd.DataFrame().to_excel(w, "NFL_Averages_Manual", index=False)
            # YTD sheets (auto)
            for s in ["YTD_Team_Offense","YTD_Team_Defense",
                      "YTD_NFL_Offense","YTD_NFL_Defense",
                      "NFL_Averages_Offense_YTD","NFL_Averages_Defense_YTD"]:
                pd.DataFrame().to_excel(w, s, index=False)

def _read_sheet(name: str) -> pd.DataFrame:
    _ensure_excel()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=name)
    except Exception:
        return pd.DataFrame()

def _write_sheet(name: str, df: pd.DataFrame):
    _ensure_excel()
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df.to_excel(w, sheet_name=name, index=False)

def _append_df(sheet_name: str, df_new: pd.DataFrame, key_cols=None):
    _ensure_excel()
    if df_new is None or df_new.empty:
        return
    df_old = _read_sheet(sheet_name)
    if df_old is None or df_old.empty:
        df_out = df_new.copy()
    else:
        df_out = pd.concat([df_old, df_new], ignore_index=True)
    if key_cols:
        df_out = df_out.drop_duplicates(subset=key_cols, keep="last")
    _write_sheet(sheet_name, df_out)

# --- Backward-compat shim (in case older code calls these) ---
SHEETS = {
    "Offense": "Offense",
    "Defense": "Defense",
    "Personnel": "Personnel",
    "SnapCounts": "SnapCounts",
    "Injuries": "Injuries",
    "MediaSummaries": "MediaSummaries",
    "OpponentPreview": "OpponentPreview",
    "Predictions": "Predictions",
    "NFL_Averages_Manual": "NFL_Averages_Manual",
    "YTD_Team_Offense": "YTD_Team_Offense",
    "YTD_Team_Defense": "YTD_Team_Defense",
    "YTD_NFL_Offense": "YTD_NFL_Offense",
    "YTD_NFL_Defense": "YTD_NFL_Defense",
    "NFL_Averages_Offense_YTD": "NFL_Averages_Offense_YTD",
    "NFL_Averages_Defense_YTD": "NFL_Averages_Defense_YTD",
}
def _load_sheet(name: str) -> pd.DataFrame:
    return _read_sheet(name)
# --- End shim ---

# ========= Utility / Computation =========
def _calc_proxy(off_row: pd.Series=None, def_row: pd.Series=None) -> float | None:
    if off_row is None and def_row is None:
        return None
    score = 0.0; count = 0
    def val(s, k):
        try:
            return float(s.get(k)) if s is not None and k in s else None
        except Exception:
            return None
    if off_row is not None:
        ypa = val(off_row, "YPA"); cmp_ = val(off_row, "CMP%")
        if ypa is not None: score += ypa; count += 1
        if cmp_ is not None: score += (cmp_/100); count += 1
    if def_row is not None:
        rz = val(def_row, "RZ% Allowed"); sacks = val(def_row, "SACKs")
        if rz is not None: score += (1 - (rz/100)); count += 1  # lower is better
        if sacks is not None: score += min(1.0, sacks/5.0); count += 1
    return round(score/count, 4) if count else None

def _colorize_df(df: pd.DataFrame):
    """Green if Team better than NFL for that metric, red if worse (per schema)."""
    if df is None or getattr(df, "empty", True):
        return df
    if not hasattr(df, "style"):
        return df
    def _row_style(row):
        bg = []
        for col in df.columns:
            if col == "Metric":
                bg.append("")
                continue
            metric = row.get("Metric")
            info = METRIC_SCHEMA.get(metric, {"higher_is_better": True})
            higher = info["higher_is_better"]
            t = row.get("Team"); n = row.get("NFL")
            if col == "Team" and pd.notnull(t) and pd.notnull(n):
                better = (t > n) if higher else (t < n)
                bg.append("background-color:#d4edda" if better else "background-color:#f8d7da")
            else:
                bg.append("")
        return bg
    return df.style.apply(_row_style, axis=1)

def _ytd(df: pd.DataFrame, up_to_week: int) -> pd.DataFrame:
    if df is None or df.empty or "Week" not in df.columns:
        return pd.DataFrame()
    return df[df["Week"].between(1, up_to_week)]

def _team_vs_nfl_table(team_df: pd.DataFrame, nfl_df: pd.DataFrame) -> pd.DataFrame:
    # Average numeric columns; then extract metrics in schema order
    t_row = team_df.select_dtypes(include="number").mean() if team_df is not None and not team_df.empty else pd.Series(dtype=float)
    n_row = nfl_df.select_dtypes(include="number").mean()  if nfl_df is not None and not nfl_df.empty else pd.Series(dtype=float)
    rows = []
    for m in METRIC_SCHEMA.keys():
        rows.append({"Metric": m, "Team": t_row.get(m), "NFL": n_row.get(m)})
    return pd.DataFrame(rows)

def _apply_excel_conditional(ws, first_data_row: int, team_col_letter: str, last_row: int):
    # Simple gradient on Team column as a visual cue
    try:
        cs = ColorScaleRule(start_type='min', start_color='F8D7DA', end_type='max', end_color='D4EDDA')
        ws.conditional_formatting.add(f"{team_col_letter}{first_data_row}:{team_col_letter}{last_row}", cs)
    except Exception:
        pass

# ---------- NEW: Helpers for Expanded NFL YTD Averages ----------
def _col_find(df: pd.DataFrame, aliases) -> str | None:
    if df is None or df.empty:
        return None
    for a in aliases:
        key = a.lower().strip()
        for c in df.columns:
            if c.lower().strip() == key:
                return c
    return None

def _safe_mean(series: pd.Series):
    try:
        s = pd.to_numeric(series, errors="coerce").dropna()
        if s.empty:
            return None
        return float(s.mean())
    except Exception:
        return None

# Aliases for Offense and Defense YTD (user-requested fields)
_OFFENSE_ALIAS_MAP = {
    "yards": ["yards","total yards","yds","tot yds"],
    "points": ["points","pts","score"],
    "turnovers": ["turnovers","tos"],
    "qbr": ["qbr"],
    "sr": ["sr%","success rate","success_rate"],
    "third_down": ["3d%","third down%","third_down%","third down rate","third_down_rate"],
    "dpa": ["dpa","drive points added","drive_points_added"],
    "yac": ["yac","yards after catch","yards_after_catch"],
    "dvoa": ["dvoa","dvoa_offense"],
    "dvoa_proxy": ["dvoa_proxy","dvoa (proxy)","dvoa proxy"],
    "rz": ["rz%","red zone%","red_zone%","red zone rate","red_zone_rate"],
    "penalties": ["penalties","pen"],
}

_DEFENSE_ALIAS_MAP = {
    "points_allowed": ["points allowed","pa","pts allowed","points, allowed","points - allowed"],
    "yards_allowed": ["yards allowed","ya","yds allowed","total yards allowed","tot yds allowed"],
    "pass_allowed": ["pass allowed","pass yards allowed","pass ya","pass yds allowed","pass"],
    "rush_allowed": ["rush allowed","rush yards allowed","rush ya","rush yds allowed","rush"],
    "turnovers": ["turnovers forced","takeaways","tos forced","turnovers"],
    "penalties": ["penalties","penalties allowed","def penalties"],
    "fr": ["fr","fumbles recovered","fumbles rec"],
    "ff": ["ff","forced fumbles","fumbles forced"],
}

_OFFENSE_YTD_ORDER = [
    "NFL Total Avg. Yards","NFL Total Avg. Points","NFL Total Avg. Turnovers","NFL Total Avg. QBR",
    "NFL Total Avg. SR%","NFL Total Avg. 3 Down%","NFL Total Avg. DPA","NFL Total Avg. YAC",
    "NFL Total Avg. DVOA","NFL Total Avg. RZ%","NFL Total Avg. Penalties"
]

_DEFENSE_YTD_ORDER = [
    "NFL Total Avg. Points Allowed","NFL Total Avg. Yards Allowed","NFL Total Avg. Pass Allowed",
    "NFL Total Avg. Rush Allowed","NFL Total Avg. Turnovers Allowed","NFL Total Avg. Penalties Allowed",
    "NFL Total Avg. FR Allowed","NFL Total Avg. FF Allowed"
]

def compute_nfl_ytd_offense_averages(offense_df: pd.DataFrame) -> pd.DataFrame:
    row = {}
    if offense_df is not None and not offense_df.empty:
        o = offense_df.copy()
        for key in ["yards","points","turnovers","qbr","sr","third_down","dpa","yac","dvoa","dvoa_proxy","rz","penalties"]:
            col = _col_find(o, _OFFENSE_ALIAS_MAP.get(key, []))
            if col:
                o[col] = pd.to_numeric(o[col], errors="coerce")
        row["NFL Total Avg. Yards"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["yards"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["yards"]) else None
        row["NFL Total Avg. Points"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["points"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["points"]) else None
        row["NFL Total Avg. Turnovers"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["turnovers"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["turnovers"]) else None
        qbr_col = _col_find(o, _OFFENSE_ALIAS_MAP["qbr"])
        row["NFL Total Avg. QBR"] = _safe_mean(o[qbr_col]) if qbr_col else None
        row["NFL Total Avg. SR%"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["sr"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["sr"]) else None
        row["NFL Total Avg. 3 Down%"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["third_down"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["third_down"]) else None
        row["NFL Total Avg. DPA"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["dpa"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["dpa"]) else None
        row["NFL Total Avg. YAC"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["yac"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["yac"]) else None
        dvoa_col = _col_find(o, _OFFENSE_ALIAS_MAP["dvoa"]) or _col_find(o, _OFFENSE_ALIAS_MAP["dvoa_proxy"])
        row["NFL Total Avg. DVOA"] = _safe_mean(o[dvoa_col]) if dvoa_col else None
        row["NFL Total Avg. RZ%"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["rz"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["rz"]) else None
        row["NFL Total Avg. Penalties"] = _safe_mean(o[_col_find(o, _OFFENSE_ALIAS_MAP["penalties"])]) if _col_find(o, _OFFENSE_ALIAS_MAP["penalties"]) else None
    return pd.DataFrame([row]).reindex(columns=_OFFENSE_YTD_ORDER)

def compute_nfl_ytd_defense_averages(defense_df: pd.DataFrame) -> pd.DataFrame:
    row = {}
    if defense_df is not None and not defense_df.empty:
        d = defense_df.copy()
        for key in ["points_allowed","yards_allowed","pass_allowed","rush_allowed","turnovers","penalties","fr","ff"]:
            col = _col_find(d, _DEFENSE_ALIAS_MAP.get(key, []))
            if col:
                d[col] = pd.to_numeric(d[col], errors="coerce")
        row["NFL Total Avg. Points Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["points_allowed"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["points_allowed"]) else None
        row["NFL Total Avg. Yards Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["yards_allowed"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["yards_allowed"]) else None
        row["NFL Total Avg. Pass Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["pass_allowed"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["pass_allowed"]) else None
        row["NFL Total Avg. Rush Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["rush_allowed"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["rush_allowed"]) else None
        row["NFL Total Avg. Turnovers Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["turnovers"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["turnovers"]) else None
        row["NFL Total Avg. Penalties Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["penalties"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["penalties"]) else None
        row["NFL Total Avg. FR Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["fr"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["fr"]) else None
        row["NFL Total Avg. FF Allowed"] = _safe_mean(d[_col_find(d, _DEFENSE_ALIAS_MAP["ff"])]) if _col_find(d, _DEFENSE_ALIAS_MAP["ff"]) else None
    return pd.DataFrame([row]).reindex(columns=_DEFENSE_YTD_ORDER)

# ========= NFL Auto-Fetch (nfl_data_py) =========
def _fetch_nfl_data_and_build_avgs(season: int, thru_week: int) -> tuple[bool, str]:
    """
    Tries to fetch weekly team stats with nfl_data_py and compute league averages
    for the METRIC_SCHEMA. Any metric with missing columns is left NaN.
    Saves result to NFL_Averages_Manual (preferred by the UI).
    """
    try:
        import nfl_data_py as nfl
    except Exception as e:
        return False, f"'nfl_data_py' not available: {e}"

    # Try a few likely functions; tolerate differences across versions
    weekly = None
    errors = []
    for fn_name in ["import_weekly_data", "import_weekly_team_stats", "load_weekly", "import_season_team_stats"]:
        try:
            fn = getattr(nfl, fn_name)
        except Exception as e:
            errors.append(f"{fn_name}: {e}")
            continue
        try:
            arg = [season] if fn_name != "load_weekly" else season
            weekly = fn(arg)
            break
        except Exception as e:
            errors.append(f"{fn_name} failed: {e}")
            weekly = None

    if weekly is None or weekly is False or (hasattr(weekly, "empty") and weekly.empty):
        return False, "Could not fetch weekly NFL data via nfl_data_py. " + ("; ".join(errors) if errors else "")

    # Filter to thru_week if 'week' column exists
    if "week" in weekly.columns:
        try:
            weekly = weekly[weekly["week"].astype(int).between(1, int(thru_week))]
        except Exception:
            pass

    # Helper to find a best-guess column
    def pick(df, *cands):
        for c in cands:
            if c in df.columns:
                return c
        return None

    out = {m: None for m in METRIC_SCHEMA.keys()}

    # YPA = pass_yards / pass_attempts
    yards_col = pick(weekly, "pass_yards", "passing_yards", "pass_yds", "yards_gained_pass")
    atts_col  = pick(weekly, "pass_attempts", "attempts", "att", "pass_att")
    if yards_col and atts_col:
        try:
            ypa = weekly[yards_col].sum() / max(1, weekly[atts_col].sum())
            out["YPA"] = float(round(ypa, 3))
        except Exception:
            pass

    # CMP% = completions / attempts
    comp_col = pick(weekly, "completions", "complete_pass", "cmp", "pass_completions")
    if comp_col and atts_col:
        try:
            cmp_pct = (weekly[comp_col].sum() / max(1, weekly[atts_col].sum())) * 100.0
            out["CMP%"] = float(round(cmp_pct, 2))
        except Exception:
            pass

    # SACKs
    sacks_col = pick(weekly, "sacks", "qb_sacks", "sack")
    if sacks_col:
        try:
            out["SACKs"] = float(round(weekly[sacks_col].mean(), 3))
        except Exception:
            pass

    # INTs (ambiguous; use available col)
    ints_col = pick(weekly, "interceptions", "int", "def_interceptions", "interceptions_thrown")
    if ints_col:
        try:
            out["INTs"] = float(round(weekly[ints_col].mean(), 3))
        except Exception:
            pass

    # QB Hits, Pressures
    qbh_col = pick(weekly, "qb_hits", "qb_hit", "qb_hits_defense")
    if qbh_col:
        try:
            out["QB Hits"] = float(round(weekly[qbh_col].mean(), 3))
        except Exception:
            pass

    prs_col = pick(weekly, "pressures", "pressure", "qb_pressures")
    if prs_col:
        try:
            out["Pressures"] = float(round(weekly[prs_col].mean(), 3))
        except Exception:
            pass

    manual_row = {k: out.get(k) for k in METRIC_SCHEMA.keys()}
    manual_df = pd.DataFrame([manual_row])
    _write_sheet("NFL_Averages_Manual", manual_df)

    try:
        _write_sheet("NFL_Raw_Fetch", weekly)
    except Exception:
        pass

    filled = sum(v is not None for v in manual_row.values())
    return True, f"Fetched season {season} up to week {thru_week}. Filled {filled}/{len(METRIC_SCHEMA)} metrics."

# ========= Sidebar (NFL tools) =========
with st.sidebar:
    st.header("NFL Tools")
    st.caption("Upload manual league averages, auto-fetch via nfl_data_py, or compute from uploads.")

    # Manual upload
    nfl_avgs_file = st.file_uploader("Upload Manual NFL Averages (CSV)", type=["csv"], key="nfl_avg_upload")

    col_sb1, col_sb2 = st.columns(2)
    col_sb3, col_sb4 = st.columns(2)

    if nfl_avgs_file:
        try:
            nfl_avgs_df = pd.read_csv(nfl_avgs_file)
            known = [c for c in nfl_avgs_df.columns if c in METRIC_SCHEMA.keys()]
            if known:
                nfl_avgs_df = nfl_avgs_df[known]
            _write_sheet("NFL_Averages_Manual", nfl_avgs_df)
            st.success("Manual NFL averages saved → sheet: NFL_Averages_Manual")
        except Exception as e:
            st.error(f"Failed to load NFL averages: {e}")

    # Auto fetch (weekly team data)
    with col_sb1:
        fetch_season = st.number_input("Fetch Season", min_value=2000, max_value=2100, value=datetime.now().year)
    with col_sb2:
        fetch_week = st.number_input("Through Week", min_value=1, max_value=25, value=1, step=1)
    with col_sb3:
        if st.button("Fetch NFL Data (Auto)"):
            ok, msg = _fetch_nfl_data_and_build_avgs(int(fetch_season), int(fetch_week))
            (st.success if ok else st.warning)(msg)
    with col_sb4:
        if st.button("Clear Manual NFL Averages"):
            _write_sheet("NFL_Averages_Manual", pd.DataFrame())
            st.success("Manual NFL averages cleared.")

    # NEW: Auto fetch Snap Counts
    st.divider()
    if st.button("Fetch Snap Counts (Auto)"):
        try:
            import nfl_data_py as nfl
            snaps = nfl.import_snap_counts(years=[int(fetch_season)])
            if "week" in snaps.columns:
                try:
                    snaps = snaps[snaps["week"].astype(int).between(1, int(fetch_week))]
                except Exception:
                    pass
            # Try to normalize column "Week" if present
            if "week" in snaps.columns and "Week" not in snaps.columns:
                snaps = snaps.rename(columns={"week": "Week"})
            _append_df("SnapCounts", snaps, key_cols=["Week","player_id","player","team","position"] if set(["player_id","player","team","position"]).issubset(snaps.columns) else None)
            st.success(f"Snap Counts fetched and saved ({len(snaps)} rows).")
        except Exception as e:
            st.error(f"Snap Counts fetch failed: {e}")

    st.divider()
    if st.button("Recompute NFL Averages (from uploads)"):
        off_all = _read_sheet("Offense"); def_all = _read_sheet("Defense")
        if off_all.empty and def_all.empty:
            st.warning("Upload Offense/Defense first.")
        else:
            _write_sheet("YTD_NFL_Offense", off_all.select_dtypes(include="number"))
            _write_sheet("YTD_NFL_Defense", def_all.select_dtypes(include="number"))
            st.success("Recomputed NFL averages from uploads (numeric columns).")

    st.divider()
    if st.button("Download All Data (Excel)"):
        _ensure_excel()
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("Click to Download", f, file_name=EXCEL_FILE)

# ========= Weekly Controls =========
st.subheader("Weekly Controls")
c1, c2, c3, c4 = st.columns([1,1,2,3])
with c1: selected_week = st.number_input("Week", min_value=1, step=1, value=1)
with c2: selected_team = st.text_input("Team", value="CHI")
with c3: opponent = st.text_input("Opponent", value="TBD")
with c4: key_notes = st.text_area("Key Notes (appear in PDF)")

# ========= Uploads =========
st.markdown("### 1) Upload Weekly Data")
u1, u2, u3, u4 = st.columns(4)
with u1: off_file  = st.file_uploader("Offense CSV", type=["csv"])
with u2: def_file  = st.file_uploader("Defense CSV", type=["csv"])
with u3: pers_file = st.file_uploader("Personnel CSV", type=["csv"])
with u4: snaps_file= st.file_uploader("Snap Counts CSV", type=["csv"])

def _load_csv_to_df(f, week: int):
    if not f: return None
    try:
        df = pd.read_csv(f)
        if "Week" not in df.columns:
            df.insert(0, "Week", week)
        else:
            df["Week"] = week
        return df
    except Exception as e:
        st.error(f"CSV load error: {e}")
        return None

if st.button("Save Weekly Uploads"):
    saved_any = False
    for nm, f in [("Offense",off_file), ("Defense",def_file), ("Personnel",pers_file), ("SnapCounts",snaps_file)]:
        df = _load_csv_to_df(f, selected_week)
        if df is not None and not df.empty:
            _append_df(nm, df, key_cols=["Week"])
            saved_any = True
    st.success("Weekly uploads saved." if saved_any else "No files to save.")

# ========= Proxy (auto) =========
off_df_all = _read_sheet("Offense")
def_df_all = _read_sheet("Defense")
off_row = off_df_all[off_df_all["Week"]==selected_week].tail(1).squeeze() if not off_df_all.empty else None
def_row = def_df_all[def_df_all["Week"]==selected_week].tail(1).squeeze() if not def_df_all.empty else None
def _calc_proxy(off_row: pd.Series=None, def_row: pd.Series=None) -> float | None:
    if off_row is None and def_row is None:
        return None
    score = 0.0; count = 0
    def val(s, k):
        try:
            return float(s.get(k)) if s is not None and k in s else None
        except Exception:
            return None
    if off_row is not None:
        ypa = val(off_row, "YPA"); cmp_ = val(off_row, "CMP%")
        if ypa is not None: score += ypa; count += 1
        if cmp_ is not None: score += (cmp_/100); count += 1
    if def_row is not None:
        rz = val(def_row, "RZ% Allowed"); sacks = val(def_row, "SACKs")
        if rz is not None: score += (1 - (rz/100)); count += 1  # lower is better
        if sacks is not None: score += min(1.0, sacks/5.0); count += 1
    return round(score/count, 4) if count else None
proxy_val = _calc_proxy(off_row, def_row)
st.markdown("#### DVOA-like Proxy (auto)")
st.write(f"Week {selected_week} proxy: **{proxy_val if proxy_val is not None else '—'}**")

# ========= Media Summaries =========
st.markdown("### Media Summaries (Store multiple per week)")
ms1, ms2, ms3 = st.columns([2,6,2])
with ms1: ms_source = st.text_input("Source (e.g., ESPN, The Athletic)")
with ms2: ms_text   = st.text_area("Summary", height=130)
with ms3:
    if st.button("Save Summary"):
        if ms_text.strip():
            _append_df("MediaSummaries",
                       pd.DataFrame([{"Week": selected_week, "Source": ms_source.strip(), "Summary": ms_text.strip()}]),
                       key_cols=["Week","Source","Summary"])
            st.success("Saved summary.")
        else:
            st.info("Nothing to save.")
_ms_all = _read_sheet("MediaSummaries")
if not _ms_all.empty and "Week" in _ms_all.columns:
    st.dataframe(_ms_all[_ms_all["Week"]==selected_week], use_container_width=True, hide_index=True)

# ========= Injuries =========
st.markdown("### Injuries – Weekly Rows")
inj1, inj2, inj3, inj4 = st.columns([2,2,6,2])
with inj1: inj_name = st.text_input("Injury")
with inj2: inj_status = st.selectbox("Status", ["Questionable","Doubtful","Out","IR","Active"])
with inj3: inj_notes = st.text_input("Notes")
with inj4:
    if st.button("Add Injury Row"):
        _append_df("Injuries",
                   pd.DataFrame([{"Week": selected_week, "Injury": inj_name, "Status": inj_status, "Notes": inj_notes}]))
        st.success("Injury row added.")
inj_all = _read_sheet("Injuries")
st.dataframe(inj_all[inj_all["Week"]==selected_week] if not inj_all.empty and "Week" in inj_all.columns else pd.DataFrame(),
             use_container_width=True, hide_index=True)

# ========= Opponent Preview =========
st.markdown("### Opponent Preview")
op1, op2 = st.columns([5,2])
with op1:
    opp_file = st.file_uploader("Upload opponent scouting/notes CSV (optional)", type=["csv"], key="opp_csv")
    opp_free = st.text_area("Or paste scouting notes here (optional)", height=130)
with op2:
    if st.button("Save Opponent Preview"):
        try:
            saved = False
            if opp_file is not None:
                df = pd.read_csv(opp_file)
                text_block = df.to_csv(index=False)
                _append_df("OpponentPreview", pd.DataFrame([{"Week": selected_week, "Notes": text_block}]), key_cols=["Week"])
                saved = True
            elif opp_free.strip():
                _append_df("OpponentPreview", pd.DataFrame([{"Week": selected_week, "Notes": opp_free.strip()}]), key_cols=["Week"])
                saved = True
            st.success("Opponent preview saved." if saved else "No CSV or text to save.")
        except Exception as e:
            st.error(f"Opponent preview save failed: {e}")
op_all = _read_sheet("OpponentPreview")
if not op_all.empty and "Week" in op_all.columns:
    st.caption("Current Week Opponent Preview")
    st.dataframe(op_all[op_all["Week"]==selected_week], use_container_width=True, hide_index=True)

# ========= Weekly Game Predictions =========
st.markdown("### Weekly Game Predictions")
p1, p2, p3 = st.columns([6,2,2])
with p1:
    pred_rationale = st.text_area("Prediction Rationale (uses strategy + proxy + injuries + opponent preview, etc.)", height=130)
with p2:
    pred_outcome = st.selectbox("Predicted Outcome", ["", "Win", "Loss"])
with p3:
    if st.button("Save Prediction"):
        if pred_outcome:
            _append_df("Predictions",
                       pd.DataFrame([{"Week": selected_week, "Rationale": pred_rationale.strip(), "Prediction": pred_outcome}]),
                       key_cols=["Week"])
            st.success("Prediction saved.")
        else:
            st.info("Choose an outcome.")
pred_all = _read_sheet("Predictions")
if not pred_all.empty:
    st.dataframe(pred_all.sort_values("Week"), use_container_width=True, hide_index=True)

# ========= 6) YTD Summary (Team vs NFL Avg) =========
st.markdown("### 6) YTD Summary (Auto-computed from uploads)")
off_ytd_team = _ytd(off_df_all, selected_week)
def_ytd_team = _ytd(def_df_all, selected_week)

# NFL average source: prefer Manual sheet (including auto-fetched) if present; else use uploads proxy; both filtered to YTD.
nfl_manual = _read_sheet("NFL_Averages_Manual")
if not nfl_manual.empty:
    nfl_off = nfl_manual  # already in metric-name columns
    nfl_def = nfl_manual
else:
    nfl_off_full = _read_sheet("YTD_NFL_Offense")
    nfl_def_full = _read_sheet("YTD_NFL_Defense")
    nfl_off = _ytd(nfl_off_full, selected_week) if not nfl_off_full.empty else off_ytd_team
    nfl_def = _ytd(nfl_def_full, selected_week) if not nfl_def_full.empty else def_ytd_team

if off_ytd_team.empty and def_ytd_team.empty:
    st.info("Upload Offense/Defense CSVs (with metrics like YPA, CMP%, RZ% Allowed, SACKs…) to see YTD.")
else:
    off_tbl = _team_vs_nfl_table(off_ytd_team, nfl_off)
    def_tbl = _team_vs_nfl_table(def_ytd_team, nfl_def)
    st.markdown(f"**{selected_team} Offense YTD vs NFL Avg (W1–W{selected_week})**")
    st.dataframe(_colorize_df(off_tbl), use_container_width=True)
    st.markdown(f"**{selected_team} Defense YTD vs NFL Avg (W1–W{selected_week})**")
    st.dataframe(_colorize_df(def_tbl), use_container_width=True)

# NEW: Compute expanded NFL YTD averages (offense & defense) from uploaded league-wide Offense/Defense sheets.
# These respect your additional columns (QBR, SR%, 3D%, DPA, YAC, DVOA, RZ%, Penalties / Allowed variants).
nfl_offense_ytd = compute_nfl_ytd_offense_averages(off_df_all)
nfl_defense_ytd = compute_nfl_ytd_defense_averages(def_df_all)

# Persist them so exports can pick up easily
_write_sheet("NFL_Averages_Offense_YTD", nfl_offense_ytd)
_write_sheet("NFL_Averages_Defense_YTD", nfl_defense_ytd)

# ========= This Week – Data Previews =========
st.markdown("### This Week – Data Previews")
def _week_preview(name):
    df = _read_sheet(name)
    if df is None or df.empty or "Week" not in df.columns:
        st.caption(f"{name}: no data yet")
    else:
        st.caption(f"{name}:")
        st.dataframe(df[df["Week"]==selected_week], use_container_width=True, hide_index=True)
pw1, pw2, pw3, pw4, pw5 = st.columns(5)
with pw1: _week_preview("Offense")
with pw2: _week_preview("Defense")
with pw3: _week_preview("Personnel")
with pw4: _week_preview("Injuries")
with pw5: _week_preview("SnapCounts")

# ========= Export helpers =========
def _save_ytd_sheets(off_team: pd.DataFrame, def_team: pd.DataFrame,
                     nfl_off_df: pd.DataFrame, nfl_def_df: pd.DataFrame):
    """Write current YTD snapshots to workbook for exporting."""
    _write_sheet("YTD_Team_Offense", off_team.select_dtypes(include="number") if not off_team.empty else pd.DataFrame())
    _write_sheet("YTD_Team_Defense", def_team.select_dtypes(include="number") if not def_team.empty else pd.DataFrame())
    _write_sheet("YTD_NFL_Offense", nfl_off_df.select_dtypes(include="number") if nfl_off_df is not None and not nfl_off_df.empty else pd.DataFrame())
    _write_sheet("YTD_NFL_Defense", nfl_def_df.select_dtypes(include="number") if nfl_def_df is not None and not nfl_def_df.empty else pd.DataFrame())

def _export_excel(tag: str):
    if openpyxl is None:
        st.error("openpyxl is required for Excel export.")
        return
    try:
        # Ensure YTD snapshot sheets updated
        _save_ytd_sheets(off_ytd_team, def_ytd_team, nfl_off, nfl_def)
        wb = openpyxl.load_workbook(EXCEL_FILE)
        # Add simple gradient on YTD team offense sheet
        if "YTD_Team_Offense" in wb.sheetnames:
            ws = wb["YTD_Team_Offense"]
            headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            if headers:
                first_data_row = 2
                last_row = ws.max_row
                if last_row >= first_data_row and "YPA" in headers:
                    from openpyxl.utils import get_column_letter
                    team_col_letter = get_column_letter(headers.index("YPA")+1)
                    _apply_excel_conditional(ws, first_data_row, team_col_letter, last_row)
        out_name = f"W{int(selected_week):02d}_{tag}.xlsx"
        wb.save(out_name)
        with open(out_name, "rb") as f:
            st.download_button(f"Download {out_name}", f, file_name=out_name)
        os.remove(out_name)
        st.success(f"{out_name} created.")
    except Exception as e:
        st.error(f"Excel export failed: {e}")

def _export_pdf_weekly(filename: str, header: str, notes: str,
                       off_tbl: pd.DataFrame|None, def_tbl: pd.DataFrame|None,
                       league_off_ytd: pd.DataFrame|None = None,
                       league_def_ytd: pd.DataFrame|None = None):
    if not HAS_REPORTLAB:
        return False, "reportlab not available"
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph(header, styles["Title"]))
    story.append(Spacer(1, 10))
    if notes:
        story.append(Paragraph("<b>Key Notes</b>", styles["Heading3"]))
        story.append(Paragraph(notes.replace("\n", "<br/>"), styles["Normal"]))
        story.append(Spacer(1, 8))

    def add_table_df(df: pd.DataFrame, title: str, colorize_team=False):
        if df is None or df.empty:
            return
        story.append(Paragraph(title, styles["Heading3"]))
        data = [list(df.columns)] + df.fillna("").values.tolist()
        tbl = Table(data, hAlign="LEFT")
        style = [
            ("BACKGROUND",(0,0),(-1,0), RL_COLORS.lightgrey),
            ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1), 0.25, RL_COLORS.grey),
            ("ALIGN",(1,1),(-1,-1),"RIGHT"),
        ]
        if colorize_team:
            try:
                team_idx = data[0].index("Team"); nfl_idx = data[0].index("NFL")
                for r in range(1, len(data)):
                    metric = data[r][0]
                    t = data[r][team_idx]; n = data[r][nfl_idx]
                    info = METRIC_SCHEMA.get(metric, {"higher_is_better": True})
                    higher = info["higher_is_better"]
                    try:
                        if t != "" and n != "":
                            tval = float(t); nval = float(n)
                            better = (tval > nval) if higher else (tval < nval)
                            style.append(("BACKGROUND",(team_idx,r),(team_idx,r),
                                          RL_COLORS.HexColor("#d4edda" if better else "#f8d7da")))
                    except Exception:
                        pass
            except Exception:
                pass
        tbl.setStyle(TableStyle(style))
        story.append(tbl)
        story.append(Spacer(1, 8))

    # Team vs NFL tables
    add_table_df(off_tbl, "Offense vs NFL Avg", colorize_team=True)
    add_table_df(def_tbl, "Defense vs NFL Avg", colorize_team=True)

    # NEW: League YTD averages (the expanded lists you requested)
    if league_off_ytd is not None and not league_off_ytd.empty:
        add_table_df(league_off_ytd, "NFL YTD Averages — Offense")
    if league_def_ytd is not None and not league_def_ytd.empty:
        add_table_df(league_def_ytd, "NFL YTD Averages — Defense")

    doc = SimpleDocTemplate(filename, pagesize=letter)
    doc.build(story)
    return True, "ok"

def _export_pdf(tag: str):
    if not HAS_REPORTLAB:
        st.warning("PDF export needs 'reportlab' in requirements.txt")
        return
    try:
        out_name = f"W{int(selected_week):02d}_{tag}.pdf"
        header = f"Week {selected_week}: {selected_team} vs {opponent}"
        # Build tables again off current sources
        off_tbl = _team_vs_nfl_table(off_ytd_team, nfl_off)
        def_tbl = _team_vs_nfl_table(def_ytd_team, nfl_def)
        ok, msg = _export_pdf_weekly(out_name, header, key_notes, off_tbl, def_tbl,
                                     league_off_ytd=nfl_offense_ytd,
                                     league_def_ytd=nfl_defense_ytd)
        if not ok:
            st.error(msg); return
        with open(out_name, "rb") as f:
            st.download_button(f"Download {out_name}", f, file_name=out_name, mime="application/pdf")
        os.remove(out_name)
        st.success(f"{out_name} created.")
    except Exception as e:
        st.error(f"PDF export failed: {e}")

# ========= Exports =========
st.markdown("### Exports")
e1, e2, e3, e4, e5, e6 = st.columns(6)
with e1:
    if st.button("Export Pre (Excel)"):
        _export_excel("Pre")
with e2:
    if st.button("Export Pre (PDF)"):
        _export_pdf("Pre")
with e3:
    if st.button("Export Post (Excel)"):
        _export_excel("Post")
with e4:
    if st.button("Export Post (PDF)"):
        _export_pdf("Post")
with e5:
    if st.button("Export Final (Excel)"):
        _export_excel("Final")
with e6:
    if st.button("Export Final (PDF)"):
        _export_pdf("Final")

st.caption("Tip: If PDF export says reportlab is missing, add `reportlab` to requirements.txt in the repo root, commit, push, then Manage app → Reboot/Rerun.")





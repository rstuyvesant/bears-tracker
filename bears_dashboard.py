import os
import streamlit as st
import pandas as pd
from fpdf import FPDF

# =========================
# App header
# =========================
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, snap counts, opponent previews, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

EXCEL_PATH = EXCEL_FILE if "EXCEL_FILE" in globals() else "bears_weekly_analytics.xlsx"

def _safe_read_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
    if not os.path.exists(file_path):
        return pd.DataFrame()
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def _find_candidate_sheets(file_path: str) -> dict:
    dfs = {}
    if not os.path.exists(file_path):
        return dfs
    try:
        xl = pd.ExcelFile(file_path)
        for s in xl.sheet_names:
            s_lower = s.lower()
            if any(k in s_lower for k in ["offense", "offensive"]):
                dfs.setdefault("offense", []).append(pd.read_excel(file_path, sheet_name=s))
            if any(k in s_lower for k in ["defense", "defensive"]):
                dfs.setdefault("defense", []).append(pd.read_excel(file_path, sheet_name=s))
            if any(k in s_lower for k in ["playbyplay", "play_by_play", "play-by-play", "pbp"]):
                dfs.setdefault("playbyplay", []).append(pd.read_excel(file_path, sheet_name=s))
    except Exception:
        pass
    for k, v in list(dfs.items()):
        dfs[k] = pd.concat(v, ignore_index=True) if v else pd.DataFrame()
    return dfs

def _select_metric_columns(df: pd.DataFrame, preferred_cols: list[str]) -> list[str]:
    return [c for c in preferred_cols if c in df.columns]

def _numeric_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def _clean_week(df: pd.DataFrame) -> pd.DataFrame:
    week_col = None
    for c in df.columns:
        if c.lower() in ["week", "wk", "game_week"]:
            week_col = c
            break
    if week_col:
        df = df.copy()
        df["Week"] = pd.to_numeric(df[week_col], errors="coerce").astype("Int64")
    return df

def _compute_avgs(df: pd.DataFrame, metrics: list[str]):
    if df.empty or not metrics:
        return pd.DataFrame(), pd.DataFrame()

    df = _clean_week(df)
    season_avgs = df[metrics].mean(numeric_only=True)
    season_table = season_avgs.reset_index().rename(columns={"index": "Metric", 0: "League_Average"})
    season_wide = pd.DataFrame({f"NFL Avg {m}": [season_avgs[m]] for m in season_avgs.index})

    if "Week" in df.columns:
        weekly = df.groupby("Week")[metrics].mean(numeric_only=True).reset_index()
    else:
        weekly = pd.DataFrame()

    return season_table, weekly, season_wide

def _write_nfl_averages_sheet(file_path: str,
                              season_off: pd.DataFrame, weekly_off: pd.DataFrame,
                              season_def: pd.DataFrame, weekly_def: pd.DataFrame,
                              season_pbp: pd.DataFrame, weekly_pbp: pd.DataFrame):
    mode = "a" if os.path.exists(file_path) else "w"
    with pd.ExcelWriter(file_path, engine="openpyxl", mode=mode, if_sheet_exists="replace") as writer:
        idx_df = pd.DataFrame({"Section": ["GeneratedAt"], "Value": [pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")]})
        idx_df.to_excel(writer, sheet_name="NFL_Averages", index=False, startrow=0)

        def _dump(df: pd.DataFrame, startrow: int, title: str) -> int:
            pd.DataFrame({"Section": [title]}).to_excel(writer, sheet_name="NFL_Averages", index=False, startrow=startrow)
            startrow += 1
            if not df.empty:
                df.to_excel(writer, sheet_name="NFL_Averages", index=False, startrow=startrow)
                startrow += len(df) + 2
            else:
                pd.DataFrame({"Info": ["(no data)"]}).to_excel(writer, sheet_name="NFL_Averages", index=False, startrow=startrow)
                startrow += 3
            return startrow

        row = 3
        row = _dump(season_off, row, "Season Averages - Offense")
        row = _dump(weekly_off, row, "Weekly Averages - Offense")
        row = _dump(season_def, row, "Season Averages - Defense")
        row = _dump(weekly_def, row, "Weekly Averages - Defense")
        row = _dump(season_pbp, row, "Season Averages - PlayByPlay")
        row = _dump(weekly_pbp, row, "Weekly Averages - PlayByPlay")

def _merge_nfl_avgs_into_preview(preview_df: pd.DataFrame, season_wide: pd.DataFrame) -> pd.DataFrame:
    if preview_df is None or preview_df.empty or season_wide.empty:
        return preview_df
    season_wide_broadcast = pd.concat([season_wide] * len(preview_df), ignore_index=True)
    season_wide_broadcast.index = preview_df.index
    return pd.concat([preview_df, season_wide_broadcast], axis=1)

PREFERRED_OFF_METRICS = ["YPA","YPC","CMP%","QBR","EPA/Play","Success Rate","Points/Game","Red Zone %","3rd Down %","Explosive Play %","SACKs Allowed","INTs Thrown","Fumbles Lost","YAC","DVOA_Proxy_Off"]
PREFERRED_DEF_METRICS = ["SACKs","INTs","FF","FR","QB Hits","Pressures","DVOA","DVOA_Proxy_Def","3D% Allowed","RZ% Allowed","EPA/Play Allowed","Success Rate Allowed","YPA Allowed","YPC Allowed","CMP% Allowed"]
PREFERRED_PBP_METRICS = ["EPA","Succ","AirYards","YAC","WPA"]

st.markdown("### 🧮 Compute NFL Averages")
with st.expander("Compute NFL Averages (write to Excel and optionally merge into previews)", expanded=False):
    do_merge = st.checkbox("Also add “NFL Avg …” columns to my Offense/Defense preview tables", value=True)
    if st.button("Compute & Save NFL Averages"):
        if not os.path.exists(EXCEL_PATH):
            st.error(f"Excel file not found: {EXCEL_PATH}")
        else:
            dfs = _find_candidate_sheets(EXCEL_PATH)

            # OFFENSE
            off_df = dfs.get("offense", pd.DataFrame())
            off_metrics = _select_metric_columns(off_df, PREFERRED_OFF_METRICS) if not off_df.empty else []
            if not off_metrics and not off_df.empty:
                off_metrics = _numeric_columns(off_df)
            season_off_tbl, weekly_off_tbl, season_off_wide = _compute_avgs(off_df, off_metrics)

            # DEFENSE
            def_df = dfs.get("defense", pd.DataFrame())
            def_metrics = _select_metric_columns(def_df, PREFERRED_DEF_METRICS) if not def_df.empty else []
            if not def_metrics and not def_df.empty:
                def_metrics = _numeric_columns(def_df)
            season_def_tbl, weekly_def_tbl, season_def_wide = _compute_avgs(def_df, def_metrics)

            # PLAY-BY-PLAY
            pbp_df = dfs.get("playbyplay", pd.DataFrame())
            pbp_metrics = _select_metric_columns(pbp_df, PREFERRED_PBP_METRICS) if not pbp_df.empty else []
            if not pbp_metrics and not pbp_df.empty:
                pbp_metrics = _numeric_columns(pbp_df)
            season_pbp_tbl, weekly_pbp_tbl, _ = _compute_avgs(pbp_df, pbp_metrics)

            try:
                _write_nfl_averages_sheet(EXCEL_PATH,
                                          season_off_tbl, weekly_off_tbl,
                                          season_def_tbl, weekly_def_tbl,
                                          season_pbp_tbl, weekly_pbp_tbl)
                st.success("✅ Wrote season & weekly NFL averages to 'NFL_Averages' sheet.")
            except Exception as e:
                st.error(f"Could not write NFL_Averages sheet: {e}")

            if do_merge:
                off_prev = st.session_state.get("offense_preview_df")
                def_prev = st.session_state.get("defense_preview_df")
                if off_prev is not None and not off_prev.empty and not season_off_wide.empty:
                    st.session_state["offense_preview_df"] = _merge_nfl_avgs_into_preview(off_prev, season_off_wide)
                    st.info("📊 Added NFL Avg columns to Offense preview.")
                if def_prev is not None and not def_prev.empty and not season_def_wide.empty:
                    st.session_state["defense_preview_df"] = _merge_nfl_avgs_into_preview(def_prev, season_def_wide)
                    st.info("📊 Added NFL Avg columns to Defense preview.")

            st.markdown("**Season NFL Averages (Quick Peek)**")
            if not season_off_tbl.empty:
                st.dataframe(season_off_tbl, use_container_width=True)
            if not season_def_tbl.empty:
                st.dataframe(season_def_tbl, use_container_width=True)
            if not season_pbp_tbl.empty:
                st.dataframe(season_pbp_tbl, use_container_width=True)

            st.markdown("**Weekly NFL Averages (Quick Peek)**")
            if not weekly_off_tbl.empty:
                st.dataframe(weekly_off_tbl, use_container_width=True)
            if not weekly_def_tbl.empty:
                st.dataframe(weekly_def_tbl, use_container_width=True)
            if not weekly_pbp_tbl.empty:
                st.dataframe(weekly_pbp_tbl, use_container_width=True)
# =========================
# Helpers
# =========================
def _to_week_key(x):
    try:
        # normalize to int-like string for dedup
        return str(int(float(x)))
    except Exception:
        return str(x)

def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True):
    """
    Append/replace a sheet inside the workbook with optional de-duplication by 'Week'.
    If 'Week' is present in both, we keep all existing rows NOT matching the new 'Week' values,
    then append the new rows, preserving columns.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    try:
        # Coerce 'Week' if present
        if "Week" in new_data.columns:
            new_data = new_data.copy()
            new_data["Week"] = new_data["Week"].map(_to_week_key)

        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            # load existing data for this sheet if it exists
            if sheet_name in book.sheetnames:
                sheet = book[sheet_name]
                existing = pd.DataFrame(sheet.values)
                if not existing.empty:
                    existing.columns = existing.iloc[0]
                    existing = existing[1:]
                else:
                    existing = pd.DataFrame()
            else:
                existing = pd.DataFrame()
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            existing = pd.DataFrame()

        if deduplicate and not existing.empty and "Week" in existing.columns and "Week" in new_data.columns:
            # Drop existing rows whose Week is in new_data
            keep = ~existing["Week"].astype(str).isin(new_data["Week"].astype(str))
            existing = existing[keep]

        # Combine columns safely
        if not existing.empty:
            all_cols = list(dict.fromkeys(list(existing.columns) + list(new_data.columns)))
            existing = existing.reindex(columns=all_cols)
            new_data = new_data.reindex(columns=all_cols)
            combined = pd.concat([existing, new_data], ignore_index=True)
        else:
            combined = new_data

        # Replace sheet
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        for r in dataframe_to_rows(combined, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def read_sheet(sheet_name: str) -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

# =========================
# Color styling helpers (simple conditional formatting)
# =========================
def style_offense(df: pd.DataFrame):
    def _colorize(val, good_thr=None, warn_thr=None, reverse=False):
        try:
            x = float(val)
        except Exception:
            return ""
        # default colors
        if reverse:
            # lower is better
            if x <= good_thr:
                return "background-color: #d4edda"  # green
            if warn_thr is not None and x <= warn_thr:
                return "background-color: #fff3cd"  # yellow
            return "background-color: #f8d7da"      # red
        else:
            if x >= good_thr:
                return "background-color: #d4edda"
            if warn_thr is not None and x >= warn_thr:
                return "background-color: #fff3cd"
            return "background-color: #f8d7da"

    if df.empty:
        return df
    sty = df.style
    # YPA: green >= 7.5, yellow 6.0–7.49, red < 6.0
    if "YPA" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=7.5, warn_thr=6.0), subset=["YPA"])
    # CMP%: green >= 66, yellow 60–65.9, red < 60
    if "CMP%" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=66.0, warn_thr=60.0), subset=["CMP%"])
    return sty

def style_defense(df: pd.DataFrame):
    def _colorize(val, good_thr=None, warn_thr=None, reverse=False):
        try:
            x = float(val)
        except Exception:
            return ""
        if reverse:
            if x <= good_thr:
                return "background-color: #d4edda"
            if warn_thr is not None and x <= warn_thr:
                return "background-color: #fff3cd"
            return "background-color: #f8d7da"
        else:
            if x >= good_thr:
                return "background-color: #d4edda"
            if warn_thr is not None and x >= warn_thr:
                return "background-color: #fff3cd"
            return "background-color: #f8d7da"

    if df.empty:
        return df
    sty = df.style
    # RZ% Allowed (lower better): green <= 50, yellow 50.1–65, red > 65
    if "RZ% Allowed" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=50.0, warn_thr=65.0, reverse=True), subset=["RZ% Allowed"])
    # SACK: green >= 3, yellow 2, red < 2
    if "SACK" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=3.0, warn_thr=2.0), subset=["SACK"])
    return sty

def style_dvoa_proxy(df: pd.DataFrame):
    def _colorize(val, good_thr=None, warn_thr=None, reverse=False):
        try:
            x = float(val)
        except Exception:
            return ""
        if reverse:
            if x <= good_thr:
                return "background-color: #d4edda"
            if warn_thr is not None and x <= warn_thr:
                return "background-color: #fff3cd"
            return "background-color: #f8d7da"
        else:
            if x >= good_thr:
                return "background-color: #d4edda"
            if warn_thr is not None and x >= warn_thr:
                return "background-color: #fff3cd"
            return "background-color: #f8d7da"

    if df.empty:
        return df
    sty = df.style
    # Off Adj EPA/play: green >= 0.15, yellow 0.00–0.149, red < 0
    if "Off Adj EPA/play" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=0.15, warn_thr=0.0), subset=["Off Adj EPA/play"])
    # Def Adj EPA/play (lower better): green <= -0.05, yellow -0.049–0.00, red > 0
    if "Def Adj EPA/play" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=-0.05, warn_thr=0.0, reverse=True), subset=["Def Adj EPA/play"])
    # Off Adj SR%: green >= 48, yellow 42–47.9, red < 42
    if "Off Adj SR%" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=48.0, warn_thr=42.0), subset=["Off Adj SR%"])
    # Def Adj SR% (lower better): green <= 42, yellow 42.1–48, red > 48
    if "Def Adj SR%" in df.columns:
        sty = sty.applymap(lambda v: _colorize(v, good_thr=42.0, warn_thr=48.0, reverse=True), subset=["Def Adj SR%"])
    return sty

# =========================
# Sidebar: Upload CSVs
# =========================
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv", key="up_off")
uploaded_defense   = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv", key="up_def")
uploaded_strategy  = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv", key="up_strat")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv", key="up_pers")
uploaded_injuries  = st.sidebar.file_uploader("Upload Injuries (.csv)", type="csv", key="up_inj")  # injuries upload
uploaded_snaps     = st.sidebar.file_uploader("Upload Snap Counts (.csv)", type="csv", key="up_snaps")  # snaps upload
uploaded_opp_prev  = st.sidebar.file_uploader("Upload Opponent Preview (.csv)", type="csv", key="up_opp_prev")  # opponent preview upload

if uploaded_offense is not None:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense")
    st.sidebar.success("✅ Offensive data uploaded.")

if uploaded_defense is not None:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense")
    st.sidebar.success("✅ Defensive data uploaded.")

if uploaded_strategy is not None:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy")
    st.sidebar.success("✅ Strategy data uploaded.")

if uploaded_personnel is not None:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel")
    st.sidebar.success("✅ Personnel data uploaded.")

if uploaded_injuries is not None:
    df_injuries = pd.read_csv(uploaded_injuries)
    append_to_excel(df_injuries, "Injuries")
    st.sidebar.success("✅ Injuries uploaded.")

if uploaded_snaps is not None:
    df_snaps = pd.read_csv(uploaded_snaps)
    append_to_excel(df_snaps, "SnapCounts")
    st.sidebar.success("✅ Snap counts uploaded.")

if uploaded_opp_prev is not None:
    df_opp_prev = pd.read_csv(uploaded_opp_prev)
    append_to_excel(df_opp_prev, "Opponent_Preview")
    st.sidebar.success("✅ Opponent Preview uploaded.")

# =========================
# Sidebar: Fetch (best-effort) via nfl_data_py
# =========================
with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pulls 2025 weekly team stats for CHI and saves to Excel.")
    fetch_week = st.number_input("Week to fetch (2025 season)", min_value=1, max_value=25, value=1, step=1, key="fetch_week_2025")

    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl

            # Optional cache update (best effort)
            try:
                nfl.update.schedule_data([2025])
            except Exception:
                pass
            try:
                nfl.update.weekly_data([2025])
            except Exception:
                pass

            weekly = nfl.import_weekly_data([2025])  # team-level weekly stats
            wk = int(fetch_week)
            team_week = weekly[(weekly["team"] == "CHI") & (weekly["week"] == wk)].copy()

            if team_week.empty:
                st.warning("No weekly team row found for CHI in that week yet.")
            else:
                team_week["Week"] = wk

                # Offense mapping (minimal)
                pass_yards = team_week["passing_yards"].iloc[0] if "passing_yards" in team_week.columns else None
                pass_att = None
                for cand in ["attempts", "passing_attempts", "pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].iloc[0]
                        break
                try:
                    ypa_val = float(pass_yards) / float(pass_att) if (pass_yards is not None and pass_att not in (None, 0)) else None
                except Exception:
                    ypa_val = None

                yards_total = None
                for cand in ["yards", "total_yards", "offense_yards"]:
                    if cand in team_week.columns:
                        yards_total = team_week[cand].iloc[0]
                        break

                completions = None
                for cand in ["completions", "passing_completions", "pass_completions"]:
                    if cand in team_week.columns:
                        completions = team_week[cand].iloc[0]
                        break

                cmp_pct = None
                if completions is not None and pass_att not in (None, 0):
                    try:
                        cmp_pct = round((float(completions) / float(pass_att)) * 100, 1)
                    except Exception:
                        cmp_pct = None

                off_row = pd.DataFrame([{
                    "Week": wk,
                    "YPA": round(ypa_val, 2) if ypa_val is not None else None,
                    "YDS": yards_total,
                    "CMP%": cmp_pct
                }])

                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].iloc[0]
                        break

                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val,
                    "RZ% Allowed": None  # requires PBP aggregation
                }])

                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)

                st.success(f"✅ Added CHI week {wk} to Offense/Defense (available fields).")
                st.caption("Note: RZ% Allowed and Pressures require play-by-play aggregation (panel below).")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

st.sidebar.markdown("### 📡 Fetch Defensive Metrics from Play-by-Play")
pbp_week = st.sidebar.number_input("Week to Fetch (2025 Season)", min_value=1, max_value=25, value=1, step=1, key="pbp_week_2025")
if st.sidebar.button("Fetch Play-by-Play Metrics"):
    try:
        import nfl_data_py as nfl
        pbp = nfl.import_pbp_data([2025], downcast=False)
        pbp_w = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == "CHI")].copy()

        if pbp_w.empty:
            st.warning("No PBP rows for CHI defense in that week yet.")
        else:
            # Red Zone Allowed: drives reaching <=20
            dmins = (pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                     .min().rename(columns={"yardline_100": "min_yardline_100"}))
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success Rate (offense success vs CHI defense)
            def play_success(row):
                if pd.isna(row.get("down")) or pd.isna(row.get("ydstogo")) or pd.isna(row.get("yards_gained")):
                    return False
                d = int(row["down"]); togo = float(row["ydstogo"]); gain = float(row["yards_gained"])
                if d == 1:
                    return gain >= 0.4 * togo
                elif d == 2:
                    return gain >= 0.6 * togo
                else:
                    return gain >= togo

            plays_mask = (~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))
            pbp_real = pbp_w[plays_mask].copy()
            success_rate = pbp_real.apply(play_success, axis=1).mean() * 100 if len(pbp_real) else 0.0

            # Pressures ≈ qb_hit + sacks
            qb_hits = pbp_w["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in pbp_w.columns else 0
            sacks = pbp_w["sack"].fillna(0).astype(int).sum() if "sack" in pbp_w.columns else 0
            pressures = int(qb_hits + sacks)

            metrics_df = pd.DataFrame([{
                "Week": int(pbp_week),
                "RZ% Allowed": round(rz_allowed, 1),
                "Success Rate% (Offense)": round(success_rate, 1),
                "Pressures": pressures
            }])
            append_to_excel(metrics_df, "Advanced_Defense", deduplicate=True)
            st.success(f"✅ Week {int(pbp_week)} PBP metrics saved — RZ% Allowed: {rz_allowed:.1f} | SR% (Off): {success_rate:.1f} | Pressures: {pressures}")
    except Exception as e:
        st.error(f"❌ Failed to fetch metrics: {e}")

# =========================
# Opponent Preview (NEW)
# =========================
st.markdown("### 🔍 Opponent Preview")
with st.form("opponent_preview_quick"):
    col1, col2, col3 = st.columns([1,1,2])
    with col1:
        opp_week = st.number_input("Week", min_value=1, max_value=25, step=1, value=1, key="opp_prev_week")
    with col2:
        opp_name = st.text_input("Opponent", value="", key="opp_prev_opp")
    with col3:
        opp_qb = st.text_input("Primary QB (optional)", value="", key="opp_prev_qb")

    col4, col5 = st.columns(2)
    with col4:
        opp_off_scheme = st.text_input("Offensive Scheme / Tendencies", value="", key="opp_prev_off")
        opp_strengths = st.text_area("Strengths (bullet/short lines)", height=100, key="opp_prev_str")
    with col5:
        opp_def_scheme = st.text_input("Defensive Scheme / Tendencies", value="", key="opp_prev_def")
        opp_weaknesses = st.text_area("Weaknesses (bullet/short lines)", height=100, key="opp_prev_weak")

    trends = st.text_area("Recent Trends / Notes", height=80, key="opp_prev_trends")
    key_matchups = st.text_area("Key Matchups", height=80, key="opp_prev_matchups")
    opp_submit = st.form_submit_button("Save Opponent Preview")

if opp_submit:
    opp_df = pd.DataFrame([{
        "Week": opp_week,
        "Opponent": opp_name,
        "QB": opp_qb,
        "Off_Scheme": opp_off_scheme,
        "Def_Scheme": opp_def_scheme,
        "Strengths": opp_strengths,
        "Weaknesses": opp_weaknesses,
        "Trends": trends,
        "KeyMatchups": key_matchups
    }])
    append_to_excel(opp_df, "Opponent_Preview", deduplicate=True)
    st.success(f"✅ Opponent Preview saved for Week {opp_week} vs {opp_name}")

# Show Opponent Preview table
df_opp_prev_show = read_sheet("Opponent_Preview")
if not df_opp_prev_show.empty:
    st.dataframe(df_opp_prev_show)
else:
    st.info("No Opponent Preview data yet. Upload CSV on the left or use the quick-entry form above.")

# =========================
# Injuries quick entry
# =========================
st.markdown("### 🏥 Injuries – Quick Entry")
with st.form("injury_quick"):
    iw1, iw2, iw3, iw4 = st.columns([1,2,1,2])
    with iw1:
        inj_week = st.number_input("Week", min_value=1, max_value=25, value=1, step=1)
    with iw2:
        inj_player = st.text_input("Player")
    with iw3:
        inj_status = st.selectbox("Status", ["Out", "Doubtful", "Questionable", "Probable", "IR", "Active"], index=5)
    with iw4:
        inj_body = st.text_input("Injury (e.g., hamstring)")

    inj_notes = st.text_area("Notes / Update (optional)", height=80)
    inj_submit = st.form_submit_button("Save Injury")

if inj_submit:
    inj_df = pd.DataFrame([{
        "Week": inj_week, "Player": inj_player, "Status": inj_status, "Injury": inj_body, "Notes": inj_notes
    }])
    append_to_excel(inj_df, "Injuries", deduplicate=False)
    st.success(f"✅ Injury entry saved for Week {inj_week}: {inj_player} — {inj_status} ({inj_body})")

df_inj = read_sheet("Injuries")
if not df_inj.empty:
    st.dataframe(df_inj)
else:
    st.caption("No injuries saved yet.")

# =========================
# Snap Counts – Quick Entry
# =========================
st.markdown("### ⏱️ Snap Counts – Quick Entry")
with st.form("snaps_quick"):
    sw1, sw2, sw3 = st.columns([1,2,2])
    with sw1:
        snap_week = st.number_input("Week", min_value=1, max_value=25, value=1, step=1, key="snap_week_input")
    with sw2:
        unit = st.selectbox("Unit", ["Offense", "Defense", "Special Teams"])
    with sw3:
        total_snaps = st.number_input("Total Snaps", min_value=0, step=1)

    snap_notes = st.text_area("Notes (optional)", height=70)
    snap_submit = st.form_submit_button("Save Snap Counts")

if snap_submit:
    sc_df = pd.DataFrame([{
        "Week": snap_week, "Unit": unit, "Total_Snaps": total_snaps, "Notes": snap_notes
    }])
    append_to_excel(sc_df, "SnapCounts", deduplicate=False)
    st.success(f"✅ Snap counts saved for Week {snap_week} ({unit})")

df_sc = read_sheet("SnapCounts")
if not df_sc.empty:
    st.dataframe(df_sc)

# =========================
# DVOA-like Proxy (opponent-adjusted EPA/SR)
# =========================
st.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.number_input("Week to Compute (2025 Season)", min_value=1, max_value=25, value=1, step=1, key="proxy_week_input")

def _success_flag(down, ydstogo, yards_gained):
    try:
        if pd.isna(down) or pd.isna(ydstogo) or pd.isna(yards_gained):
            return False
        d = int(down); togo = float(ydstogo); gain = float(yards_gained)
        if d == 1:
            return gain >= 0.4 * togo
        elif d == 2:
            return gain >= 0.6 * togo
        else:
            return gain >= togo
    except Exception:
        return False

if st.button("Compute DVOA-like Proxy"):
    try:
        import nfl_data_py as nfl
        wk = int(proxy_week)
        pbp = nfl.import_pbp_data([2025], downcast=False)

        plays = pbp[(~pbp["play_type"].isin(["no_play"])) & (~pbp["penalty"].fillna(False)) & (~pbp["epa"].isna())].copy()
        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == "CHI")].copy()
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == "CHI")].copy()

        if bears_off.empty and bears_def.empty:
            st.warning("No CHI plays found for that week yet.")
        else:
            opps = set()
            if not bears_off.empty:
                opps.update(bears_off["defteam"].unique().tolist())
            if not bears_def.empty:
                opps.update(bears_def["posteam"].unique().tolist())
            opponent = list(opps)[0] if opps else "UNK"

            prior = plays[plays["week"] < wk].copy()
            opp_def_plays = prior[prior["defteam"] == opponent].copy()
            opp_def_epa = opp_def_plays["epa"].mean() if len(opp_def_plays) else None
            opp_def_success = opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_def_plays) else None

            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            opp_off_success = opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_off_plays) else None

            if len(bears_off):
                chi_off_epa = bears_off["epa"].mean()
                chi_off_success = bears_off.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_off_epa = None; chi_off_success = None

            if len(bears_def):
                chi_def_epa_allowed = bears_def["epa"].mean()
                chi_def_success_allowed = bears_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_def_epa_allowed = None; chi_def_success_allowed = None

            def safe_diff(a, b):
                if a is None or pd.isna(a) or b is None or pd.isna(b):
                    return None
                return float(a) - float(b)

            off_adj_epa = safe_diff(chi_off_epa, opp_def_epa)
            off_adj_sr  = safe_diff(chi_off_success, opp_def_success)
            def_adj_epa = safe_diff(chi_def_epa_allowed, opp_off_epa)
            def_adj_sr  = safe_diff(chi_def_success_allowed, opp_off_success)

            out = pd.DataFrame([{
                "Week": wk,
                "Opponent": opponent,
                "Off Adj EPA/play": round(off_adj_epa, 3) if off_adj_epa is not None else None,
                "Off Adj SR%": round(off_adj_sr * 100, 1) if off_adj_sr is not None else None,
                "Def Adj EPA/play": round(def_adj_epa, 3) if def_adj_epa is not None else None,
                "Def Adj SR%": round(def_adj_sr * 100, 1) if def_adj_sr is not None else None,
                "Off EPA/play": round(chi_off_epa, 3) if chi_off_epa is not None else None,
                "Def EPA allowed/play": round(chi_def_epa_allowed, 3) if chi_def_epa_allowed is not None else None
            }])

            append_to_excel(out, "DVOA_Proxy", deduplicate=True)
            st.success(f"✅ DVOA-like proxy saved for Week {wk} vs {opponent}")

    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# =========================
# Download Excel
# =========================
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =========================
# Preview sections (styled where helpful)
# =========================
# Offense
df_off_show = read_sheet("Offense")
if not df_off_show.empty:
    st.subheader("📊 Offensive Analytics")
    try:
        st.dataframe(style_offense(df_off_show), use_container_width=True)
    except Exception:
        st.dataframe(df_off_show, use_container_width=True)

# Defense
df_def_show = read_sheet("Defense")
if not df_def_show.empty:
    st.subheader("🛡️ Defensive Analytics")
    try:
        st.dataframe(style_defense(df_def_show), use_container_width=True)
    except Exception:
        st.dataframe(df_def_show, use_container_width=True)

# Advanced Defense
df_advdef_show = read_sheet("Advanced_Defense")
if not df_advdef_show.empty:
    st.subheader("🧪 Advanced Defensive Metrics (PBP-derived)")
    st.dataframe(df_advdef_show, use_container_width=True)

# Personnel
df_pers_show = read_sheet("Personnel")
if not df_pers_show.empty:
    st.subheader("👥 Personnel Usage")
    st.dataframe(df_pers_show, use_container_width=True)

# DVOA Proxy Preview
df_proxy_show = read_sheet("DVOA_Proxy")
if not df_proxy_show.empty:
    st.subheader("📈 DVOA-like Proxy Metrics")
    try:
        st.dataframe(style_dvoa_proxy(df_proxy_show), use_container_width=True)
    except Exception:
        st.dataframe(df_proxy_show, use_container_width=True)

# Media Summaries
st.markdown("### 📰 Weekly Beat Writer / ESPN Summary")
with st.form("media_form"):
    media_week = st.number_input("Week", min_value=1, max_value=25, step=1, key="media_week_input")
    media_opponent = st.text_input("Opponent")
    media_summary = st.text_area("Beat Writer & ESPN Summary (Game Recap, Analysis, Strategy, etc.)")
    submit_media = st.form_submit_button("Save Summary")

if submit_media:
    media_df = pd.DataFrame([{
        "Week": media_week,
        "Opponent": media_opponent,
        "Summary": media_summary
    }])
    append_to_excel(media_df, "Media_Summaries", deduplicate=False)
    st.success(f"✅ Summary for Week {media_week} vs {media_opponent} saved.")

df_media_show = read_sheet("Media_Summaries")
if not df_media_show.empty:
    st.subheader("📰 Saved Media Summaries")
    st.dataframe(df_media_show, use_container_width=True)

# =========================
# Weekly Prediction
# =========================
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy   = read_sheet("Strategy")
        df_offense    = read_sheet("Offense")
        df_defense    = read_sheet("Defense")
        df_advdef     = read_sheet("Advanced_Defense")
        df_proxy      = read_sheet("DVOA_Proxy")

        row_s = df_strategy[df_strategy["Week"].map(_to_week_key) == _to_week_key(week_to_predict)] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"].map(_to_week_key) == _to_week_key(week_to_predict)] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"].map(_to_week_key) == _to_week_key(week_to_predict)] if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"].map(_to_week_key) == _to_week_key(week_to_predict)] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"].map(_to_week_key) == _to_week_key(week_to_predict)] if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            rz_allowed = None
            pressures  = None
            if not row_a.empty:
                rz_allowed = _safe_float(row_a.iloc[0].get("RZ% Allowed"), default=None)
                pressures  = _safe_float(row_a.iloc[0].get("Pressures"), default=None)
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), default=None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), default=None)

            reason_bits = []

            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win – efficiency edge on both sides"
                reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def {def_adj_epa:+.2f} EPA/play vs opp O")
            elif (pressures is not None and pressures >= 8) and any(tok in strategy_text for tok in ["blitz", "pressure"]):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win – red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient offense and poor red zone defense"
                reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient passing and weak red zone defense"
                reason_bits.append(f"YPA={ypa:.1f}")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            else:
                prediction = "Loss – no clear advantage in key strategy or stats"
                if off_adj_epa is not None:
                    reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play")
                if def_adj_epa is not None:
                    reason_bits.append(f"Def {def_adj_epa:+.2f} EPA/play")
                if pressures is not None:
                    reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            reason_text = " | ".join(reason_bits)
            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")
            if reason_text:
                st.caption(reason_text)

            prediction_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(prediction_entry, "Predictions", deduplicate=True)
        else:
            st.info("Please upload or fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# Show saved predictions
df_preds_show = read_sheet("Predictions")
if not df_preds_show.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(df_preds_show, use_container_width=True)

# =========================
# Weekly Report PDF
# =========================
st.markdown("### 🧾 Download Weekly Game Report (PDF)")
report_week = st.number_input("Select Week for Report", min_value=1, max_value=25, step=1, key="report_week_input")

if st.button("Generate Weekly Report (PDF)"):
    try:
        df_strategy = read_sheet("Strategy")
        df_offense  = read_sheet("Offense")
        df_defense  = read_sheet("Defense")
        df_media    = read_sheet("Media_Summaries")
        df_preds    = read_sheet("Predictions")
        df_opp      = read_sheet("Opponent_Preview")
        df_inj      = read_sheet("Injuries")
        df_snaps    = read_sheet("SnapCounts")
        df_advdef   = read_sheet("Advanced_Defense")
        df_proxy    = read_sheet("DVOA_Proxy")

        r = str(int(report_week))

        strat_row = df_strategy[df_strategy["Week"].map(_to_week_key) == r]
        off_row   = df_offense[df_offense["Week"].map(_to_week_key) == r]
        def_row   = df_defense[df_defense["Week"].map(_to_week_key) == r]
        media_rows= df_media[df_media["Week"].map(_to_week_key) == r]
        pred_row  = df_preds[df_preds["Week"].map(_to_week_key) == r]
        opp_row   = df_opp[df_opp["Week"].map(_to_week_key) == r]
        inj_rows  = df_inj[df_inj["Week"].map(_to_week_key) == r]
        sc_rows   = df_snaps[df_snaps["Week"].map(_to_week_key) == r]
        adv_row   = df_advdef[df_advdef["Week"].map(_to_week_key) == r]
        proxy_row = df_proxy[df_proxy["Week"].map(_to_week_key) == r]

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, f"Chicago Bears Weekly Report – Week {report_week}", ln=True)

        pdf.set_font("Arial", "", 12)

        # Opponent Preview
        if not opp_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "🔍 Opponent Preview:", ln=True)
            pdf.set_font("Arial", "", 12)
            opp = opp_row.iloc[0].to_dict()
            for k in ["Opponent", "QB", "Off_Scheme", "Def_Scheme", "Strengths", "Weaknesses", "Trends", "KeyMatchups"]:
                if k in opp and pd.notna(opp[k]) and str(opp[k]).strip():
                    pdf.multi_cell(0, 8, f"{k}: {opp[k]}")
            pdf.ln(2)

        # Prediction
        if not pred_row.empty:
            outcome = pred_row.iloc[0].get("Prediction", "")
            reason  = pred_row.iloc[0].get("Reason", "")
            notes   = pred_row.iloc[0].get("Notes", "")
            pdf.multi_cell(0, 8, f"🔮 Prediction: {outcome}\n📝 Reason: {reason}\n{notes}\n")

        # Strategy
        if not strat_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "📘 Strategy Notes:", ln=True)
            pdf.set_font("Arial", "", 12)
            strategy_text = strat_row.iloc[0].astype(str).str.cat(sep=" | ")
            pdf.multi_cell(0, 8, strategy_text)

        # Offense
        if not off_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "📊 Offensive Analytics:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in off_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        # Defense
        if not def_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "🛡️ Defensive Analytics:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in def_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        # Advanced Defense
        if not adv_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "🧪 Advanced Defensive Metrics:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in adv_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        # DVOA-like proxy
        if not proxy_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "📈 DVOA-like Proxy:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in proxy_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        # Injuries
        if not inj_rows.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "🏥 Injuries:", ln=True)
            pdf.set_font("Arial", "", 12)
            for _, rrow in inj_rows.iterrows():
                pdf.multi_cell(0, 8, f"{rrow.get('Player','')} — {rrow.get('Status','')} ({rrow.get('Injury','')}) {rrow.get('Notes','')}")

        # Snap counts
        if not sc_rows.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "⏱️ Snap Counts:", ln=True)
            pdf.set_font("Arial", "", 12)
            for _, rrow in sc_rows.iterrows():
                pdf.cell(0, 8, f"{rrow.get('Unit','')}: {rrow.get('Total_Snaps','')} — {rrow.get('Notes','')}", ln=True)

        # Media summaries
        if not media_rows.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "📰 Media Summaries:", ln=True)
            pdf.set_font("Arial", "", 12)
            for _, mrow in media_rows.iterrows():
                source = mrow.get("Opponent", "Source")
                summary = mrow.get("Summary", "")
                pdf.multi_cell(0, 8, f"{source}:\n{summary}\n")

        pdf_output = f"week_{report_week}_report.pdf"
        pdf.output(pdf_output)
        with open(pdf_output, "rb") as f:
            st.download_button(
                label=f"📥 Download Week {report_week} Report (PDF)",
                data=f,
                file_name=pdf_output,
                mime="application/pdf"
            )
    except Exception as e:
        st.error(f"❌ Failed to generate PDF. Error: {e}")
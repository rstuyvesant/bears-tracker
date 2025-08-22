import os
import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime

# ---------- Page Setup ----------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, previews, league averages, and exports.")

# Main Excel workbook
EXCEL_FILE = "bears_weekly_analytics.xlsx"

# ---------- Helpers ----------
def make_export_filename(week: int, phase: str, ext: str = "xlsx") -> str:
    """
    Standardized export filenames: W01_post_2025-08-21.xlsx / .pdf
    """
    today = datetime.today().strftime("%Y-%m-%d")
    return f"W{week:02d}_{phase}_{today}.{ext}"

def load_excel_sheet(file_name, sheet_name):
    if not os.path.exists(file_name):
        return pd.DataFrame()
    try:
        return pd.read_excel(file_name, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def append_to_excel(new_df, sheet_name, file_name=EXCEL_FILE, dedupe_cols=None):
    """
    Append new_df to sheet `sheet_name` inside Excel `file_name`.
    - Creates file/sheet if missing.
    - De-duplicates by dedupe_cols (if given), else drops exact duplicate rows.
    """
    if new_df is None or new_df.empty:
        return False, "Empty DataFrame — nothing to append."

    # Normalize common columns
    if "Week" in new_df.columns:
        new_df["Week"] = pd.to_numeric(new_df["Week"], errors="coerce").astype("Int64")

    # Ensure file or create with this sheet
    if not os.path.exists(file_name):
        with pd.ExcelWriter(file_name, engine="openpyxl", mode="w") as writer:
            new_df.to_excel(writer, sheet_name=sheet_name, index=False)
        return True, "Created workbook and wrote first sheet."

    # Read existing if present
    try:
        xl = pd.ExcelFile(file_name)
        old_df = pd.read_excel(file_name, sheet_name=sheet_name) if sheet_name in xl.sheet_names else pd.DataFrame()
    except Exception as e:
        return False, f"Could not read existing workbook: {e}"

    # Align columns
    all_cols = list(dict.fromkeys(list(old_df.columns if not old_df.empty else []) + list(new_df.columns)))
    old_df = old_df.reindex(columns=all_cols, fill_value=pd.NA)
    new_df = new_df.reindex(columns=all_cols, fill_value=pd.NA)

    combined = pd.concat([old_df, new_df], ignore_index=True)

    # Deduplicate
    if dedupe_cols:
        keys = [c for c in dedupe_cols if c in combined.columns]
        combined = combined.drop_duplicates(subset=keys or None, keep="last").reset_index(drop=True)
    else:
        combined = combined.drop_duplicates(keep="last").reset_index(drop=True)

    # Write back
    try:
        with pd.ExcelWriter(file_name, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, sheet_name=sheet_name, index=False)
        return True, f"Appended {len(new_df)} rows → {sheet_name} (now {len(combined)} total)."
    except Exception as e:
        return False, f"Write failed: {e}"

# ---------- Sidebar (Quick Actions + Export Naming Controls) ----------
with st.sidebar:
    st.header("⚙️ Quick Actions")
    st.write(f"Workbook: **{EXCEL_FILE}**")

    # Week/Phase controls shared by sidebar quick download + PDF section + bottom export
    week_num_global = st.number_input("Week #", min_value=1, max_value=18, value=1)
    phase_global = st.selectbox("Phase", ["pre", "post", "final"])

    # Quick compute
    run_compute_now = st.button("🧮 Compute NFL Averages (sidebar)")
    st.session_state["trigger_compute_from_sidebar"] = bool(run_compute_now or st.session_state.get("trigger_compute_from_sidebar", False))

    # Quick Excel download with auto filename
    if os.path_exists := os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                "💾 Download All Data (Excel)",
                data=f.read(),
                file_name=make_export_filename(int(week_num_global), phase_global, "xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.caption("No Excel yet — upload or compute to create it.")

    st.markdown("---")
    st.header("📂 Sections")
    st.caption("This page renders in this order:")
    st.write("• 📥 Uploads")
    st.write("• 🧾 Previews")
    st.write("• 📊 DVOA Proxy")
    st.write("• 🧮 NFL Averages")
    st.write("• 🎨 Color Settings")
    st.write("• 🧾 PDF Exports")
    st.write("• 💾 Export (bottom)")

# =========================================================
# 1) 📥 UPLOADS
# =========================================================
st.markdown("## 📥 Upload Weekly Data")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Offense")
    up_off = st.file_uploader("Upload Offense CSV", type=["csv"], key="upload_off")
    if up_off is not None:
        try:
            off_df = pd.read_csv(up_off)
            st.dataframe(off_df.head(20), use_container_width=True)
            ok, msg = append_to_excel(off_df, "Offense", EXCEL_FILE, dedupe_cols=["Week", "Team", "Opponent"])
            (st.success if ok else st.error)(msg)
            st.session_state["offense_preview_df"] = off_df.copy()
        except Exception as e:
            st.error(f"Offense upload failed: {e}")

with col2:
    st.subheader("Defense")
    up_def = st.file_uploader("Upload Defense CSV", type=["csv"], key="upload_def")
    if up_def is not None:
        try:
            def_df = pd.read_csv(up_def)
            st.dataframe(def_df.head(20), use_container_width=True)
            ok, msg = append_to_excel(def_df, "Defense", EXCEL_FILE, dedupe_cols=["Week", "Team", "Opponent"])
            (st.success if ok else st.error)(msg)
            st.session_state["defense_preview_df"] = def_df.copy()
        except Exception as e:
            st.error(f"Defense upload failed: {e}")

col3, col4 = st.columns(2)

with col3:
    st.subheader("Personnel")
    up_pers = st.file_uploader("Upload Personnel Usage CSV", type=["csv"], key="upload_pers")
    if up_pers is not None:
        try:
            pers_df = pd.read_csv(up_pers)
            st.dataframe(pers_df.head(20), use_container_width=True)
            ok, msg = append_to_excel(pers_df, "Personnel", EXCEL_FILE, dedupe_cols=["Week", "Team"])
            (st.success if ok else st.error)(msg)
        except Exception as e:
            st.error(f"Personnel upload failed: {e}")

with col4:
    st.subheader("Strategy")
    up_strat = st.file_uploader("Upload Strategy Notes CSV", type=["csv"], key="upload_strat")
    if up_strat is not None:
        try:
            strat_df = pd.read_csv(up_strat)
            st.dataframe(strat_df.head(20), use_container_width=True)
            ok, msg = append_to_excel(strat_df, "Strategy", EXCEL_FILE, dedupe_cols=["Week", "Team", "Opponent"])
            (st.success if ok else st.error)(msg)
        except Exception as e:
            st.error(f"Strategy upload failed: {e}")

st.subheader("Opponent Preview")
up_opp = st.file_uploader("Upload Opponent Preview CSV", type=["csv"], key="upload_opp")
if up_opp is not None:
    try:
        opp_df = pd.read_csv(up_opp)
        st.dataframe(opp_df.head(20), use_container_width=True)
        ok, msg = append_to_excel(opp_df, "Opponent_Preview", EXCEL_FILE, dedupe_cols=["Week", "Opponent"])
        (st.success if ok else st.error)(msg)
    except Exception as e:
        st.error(f"Opponent Preview upload failed: {e}")

# --- Injuries & Snap Counts ---
col_ic1, col_ic2 = st.columns(2)

with col_ic1:
    st.subheader("Injuries")
    up_inj = st.file_uploader("Upload Injuries CSV", type=["csv"], key="upload_inj")
    if up_inj is not None:
        try:
            inj_df = pd.read_csv(up_inj)
            st.dataframe(inj_df.head(20), use_container_width=True)
            ok, msg = append_to_excel(inj_df, "Injuries", EXCEL_FILE, dedupe_cols=["Week", "Team", "Player"])
            (st.success if ok else st.error)(msg)
        except Exception as e:
            st.error(f"Injuries upload failed: {e}")

with col_ic2:
    st.subheader("Snap Counts")
    up_snap = st.file_uploader("Upload Snap Counts CSV", type=["csv"], key="upload_snap")
    if up_snap is not None:
        try:
            snap_df = pd.read_csv(up_snap)
            st.dataframe(snap_df.head(20), use_container_width=True)
            ok, msg = append_to_excel(snap_df, "Snap_Counts", EXCEL_FILE, dedupe_cols=["Week", "Team", "Player"])
            (st.success if ok else st.error)(msg)
        except Exception as e:
            st.error(f"Snap Counts upload failed: {e}")

# =========================================================
# 2) 🧾 PREVIEWS
# =========================================================
st.markdown("## 🧾 Previews")

prev_col1, prev_col2 = st.columns(2)

with prev_col1:
    offense_all = load_excel_sheet(EXCEL_FILE, "Offense")
    if not offense_all.empty:
        st.session_state["offense_preview_df"] = offense_all.tail(30).reset_index(drop=True)
        st.markdown("**Offense Preview (latest rows)**")
        st.dataframe(st.session_state["offense_preview_df"], use_container_width=True)
    else:
        st.info("No Offense data yet.")

with prev_col2:
    defense_all = load_excel_sheet(EXCEL_FILE, "Defense")
    if not defense_all.empty:
        st.session_state["defense_preview_df"] = defense_all.tail(30).reset_index(drop=True)
        st.markdown("**Defense Preview (latest rows)**")
        st.dataframe(st.session_state["defense_preview_df"], use_container_width=True)
    else:
        st.info("No Defense data yet.")

ic_col1, ic_col2 = st.columns(2)

with ic_col1:
    injuries_all = load_excel_sheet(EXCEL_FILE, "Injuries")
    st.markdown("**Injuries (latest rows)**")
    if not injuries_all.empty:
        st.dataframe(injuries_all.tail(50).reset_index(drop=True), use_container_width=True)
    else:
        st.info("No Injuries data yet.")

with ic_col2:
    snaps_all = load_excel_sheet(EXCEL_FILE, "Snap_Counts")
    st.markdown("**Snap Counts (latest rows)**")
    if not snaps_all.empty:
        st.dataframe(snaps_all.tail(50).reset_index(drop=True), use_container_width=True)
    else:
        st.info("No Snap Counts data yet.")

# =========================================================
# 3) 📊 DVOA PROXY (writes columns to Offense & Defense)
# =========================================================
st.markdown("## 📊 Compute DVOA Proxy")
with st.expander("Add/refresh DVOA_Proxy_Off and DVOA_Proxy_Def in the workbook", expanded=False):

    # Tunable weights (transparent, tweak as desired)
    w_off = {
        "YPA": 0.30, "CMP%": 0.20, "EPA/Play": 0.30, "3rd Down %": 0.10, "Red Zone %": 0.10,
        "SACKs Allowed": -0.10, "INTs Thrown": -0.10, "Fumbles Lost": -0.05,
    }
    w_def = {
        "SACKs": 0.25, "INTs": 0.20, "FF": 0.10, "FR": 0.05, "QB Hits": 0.05, "Pressures": 0.05,
        "EPA/Play Allowed": -0.25, "Success Rate Allowed": -0.15,
        "3D% Allowed": -0.10, "RZ% Allowed": -0.10, "YPA Allowed": -0.10, "YPC Allowed": -0.10, "CMP% Allowed": -0.10,
    }

    def _normalize(series):
        s = pd.to_numeric(series, errors="coerce")
        if s.dropna().empty:
            return s
        z = (s - s.mean()) / (s.std(ddof=0) if s.std(ddof=0) else 1.0)
        return z.clip(-3, 3)

    def _compute_proxy(df, weights):
        if df is None or df.empty:
            return pd.Series(dtype="float64")
        score = pd.Series(0.0, index=df.index)
        for col, w in weights.items():
            if col in df.columns:
                score = score + w * _normalize(df[col])
        score = (score - score.mean()) / (score.std(ddof=0) if score.std(ddof=0) else 1.0)
        return score * 8.0  # ~ -24..+24 range

    if st.button("Compute & Save DVOA Proxies"):
        off_all = load_excel_sheet(EXCEL_FILE, "Offense")
        def_all = load_excel_sheet(EXCEL_FILE, "Defense")
        if off_all.empty and def_all.empty:
            st.warning("No Offense/Defense data yet.")
        else:
            if not off_all.empty:
                off_all = off_all.copy()
                off_all["DVOA_Proxy_Off"] = _compute_proxy(off_all, w_off)
                ok, msg = append_to_excel(off_all, "Offense", EXCEL_FILE, dedupe_cols=["Week","Team","Opponent"])
                (st.success if ok else st.error)(f"Offense: {msg}")
                st.session_state["offense_preview_df"] = off_all.tail(30).reset_index(drop=True)

            if not def_all.empty:
                def_all = def_all.copy()
                def_all["DVOA_Proxy_Def"] = _compute_proxy(def_all, w_def)
                ok, msg = append_to_excel(def_all, "Defense", EXCEL_FILE, dedupe_cols=["Week","Team","Opponent"])
                (st.success if ok else st.error)(f"Defense: {msg}")
                st.session_state["defense_preview_df"] = def_all.tail(30).reset_index(drop=True)

            st.info("DVOA proxies updated. Re-run **🧮 Compute NFL Averages** to refresh league means.")

# =========================================================
# 4) 🧮 COMPUTE NFL AVERAGES (Season + Weekly) + Optional Merge
# =========================================================
EXCEL_PATH = EXCEL_FILE  # alias

def _find_candidate_sheets(file_path):
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

def _select_metric_columns(df, preferred_cols):
    return [c for c in preferred_cols if isinstance(df, pd.DataFrame) and c in df.columns]

def _numeric_columns(df):
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def _clean_week(df):
    wk = None
    for c in df.columns:
        if c.lower() in ["week", "wk", "game_week"]:
            wk = c; break
    if wk:
        df = df.copy()
        df["Week"] = pd.to_numeric(df[wk], errors="coerce").astype("Int64")
    return df

def _compute_avgs(df, metrics):
    if df is None or df.empty or not metrics:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df = _clean_week(df)
    season_avgs = df[metrics].mean(numeric_only=True)
    season_table = season_avgs.reset_index().rename(columns={"index": "Metric", 0: "League_Average"})
    season_table["SourceRows"] = len(df)
    season_wide = pd.DataFrame({f"NFL Avg {m}": [season_avgs[m]] for m in season_avgs.index})
    weekly = df.groupby("Week")[metrics].mean(numeric_only=True).reset_index() if "Week" in df.columns else pd.DataFrame()
    return season_table, weekly, season_wide

def _write_nfl_averages_sheet(file_path,
                              season_off, weekly_off,
                              season_def, weekly_def,
                              season_pbp, weekly_pbp):
    mode = "a" if os.path.exists(file_path) else "w"
    with pd.ExcelWriter(file_path, engine="openpyxl", mode=mode, if_sheet_exists="replace") as writer:
        idx_df = pd.DataFrame({"Section": ["GeneratedAt"], "Value": [pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")]})
        idx_df.to_excel(writer, sheet_name="NFL_Averages", index=False, startrow=0)

        def _dump(df, startrow, title):
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

def _merge_nfl_avgs_into_preview(preview_df, season_wide):
    if preview_df is None or preview_df.empty or season_wide.empty:
        return preview_df
    season_wide_broadcast = pd.concat([season_wide] * len(preview_df), ignore_index=True)
    season_wide_broadcast.index = preview_df.index
    return pd.concat([preview_df, season_wide_broadcast], axis=1)

PREFERRED_OFF_METRICS = ["YPA","YPC","CMP%","QBR","EPA/Play","Success Rate","Points/Game","Red Zone %","3rd Down %","Explosive Play %","SACKs Allowed","INTs Thrown","Fumbles Lost","YAC","DVOA_Proxy_Off"]
PREFERRED_DEF_METRICS = ["SACKs","INTs","FF","FR","QB Hits","Pressures","DVOA","DVOA_Proxy_Def","3D% Allowed","RZ% Allowed","EPA/Play Allowed","Success Rate Allowed","YPA Allowed","YPC Allowed","CMP% Allowed"]
PREFERRED_PBP_METRICS = ["EPA","Succ","AirYards","YAC","WPA"]

st.markdown("## 🧮 Compute NFL Averages")
with st.expander("Compute NFL Averages (write to Excel and optionally merge into previews)", expanded=False):
    do_merge = st.checkbox("Also add “NFL Avg …” columns to my Offense/Defense preview tables", value=True)
    clicked_main_compute = st.button("Compute & Save NFL Averages")

    if clicked_main_compute or st.session_state.get("trigger_compute_from_sidebar"):
        # reset sidebar trigger so it doesn't auto-run next time
        st.session_state["trigger_compute_from_sidebar"] = False

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

            # PLAY-BY-PLAY (optional)
            pbp_df = dfs.get("playbyplay", pd.DataFrame())
            pbp_metrics = _select_metric_columns(pbp_df, PREFERRED_PBP_METRICS) if not pbp_df.empty else []
            if not pbp_metrics and not pbp_df.empty:
                pbp_metrics = _numeric_columns(pbp_df)
            season_pbp_tbl, weekly_pbp_tbl, _ = _compute_avgs(pbp_df, pbp_metrics)

            # Write the sheet
            try:
                _write_nfl_averages_sheet(EXCEL_PATH,
                                          season_off_tbl, weekly_off_tbl,
                                          season_def_tbl, weekly_def_tbl,
                                          season_pbp_tbl, weekly_pbp_tbl)
                st.success("✅ Wrote season & weekly NFL averages to 'NFL_Averages' sheet.")
            except Exception as e:
                st.error(f"Could not write NFL_Averages sheet: {e}")

            # Save for other widgets
            st.session_state["season_off_tbl"] = season_off_tbl
            st.session_state["season_def_tbl"] = season_def_tbl
            st.session_state["weekly_off_tbl"] = weekly_off_tbl
            st.session_state["weekly_def_tbl"] = weekly_def_tbl
            st.session_state["season_pbp_tbl"] = season_pbp_tbl
            st.session_state["weekly_pbp_tbl"] = weekly_pbp_tbl
            st.session_state["season_off_wide"] = season_off_wide
            st.session_state["season_def_wide"] = season_def_wide

            # Optional merge into previews
            if do_merge:
                off_prev = st.session_state.get("offense_preview_df")
                def_prev = st.session_state.get("defense_preview_df")
                if isinstance(off_prev, pd.DataFrame) and not off_prev.empty and not season_off_wide.empty:
                    st.session_state["offense_preview_df"] = _merge_nfl_avgs_into_preview(off_prev, season_off_wide)
                    st.info("📊 Added NFL Avg columns to Offense preview.")
                if isinstance(def_prev, pd.DataFrame) and not def_prev.empty and not season_def_wide.empty:
                    st.session_state["defense_preview_df"] = _merge_nfl_avgs_into_preview(def_prev, season_def_wide)
                    st.info("📊 Added NFL Avg columns to Defense preview.")

            # Quick peeks
            st.markdown("**Season NFL Averages (Quick Peek)**")
            if not season_off_tbl.empty: st.dataframe(season_off_tbl, use_container_width=True)
            if not season_def_tbl.empty: st.dataframe(season_def_tbl, use_container_width=True)
            if not season_pbp_tbl.empty: st.dataframe(season_pbp_tbl, use_container_width=True)

            st.markdown("**Weekly NFL Averages (Quick Peek)**")
            if not weekly_off_tbl.empty: st.dataframe(weekly_off_tbl, use_container_width=True)
            if not weekly_def_tbl.empty: st.dataframe(weekly_def_tbl, use_container_width=True)
            if not weekly_pbp_tbl.empty: st.dataframe(weekly_pbp_tbl, use_container_width=True)

# =========================================================
# 5) 🎨 COLOR THRESHOLDS
# =========================================================
THRESHOLDS_OFF = {
    "YPA": (6.8, 7.8), "YPC": (4.0, 5.0), "CMP%": (62.0, 68.0), "QBR": (50.0, 65.0),
    "EPA/Play": (0.00, 0.08), "Success Rate": (42.0, 50.0), "Points/Game": (20.0, 26.0),
    "Red Zone %": (50.0, 62.0), "3rd Down %": (36.0, 44.0), "Explosive Play %": (9.0, 13.0),
    "YAC": (4.2, 5.2), "SACKs Allowed": (2.8, 1.8), "INTs Thrown": (1.2, 0.6),
    "Fumbles Lost": (0.8, 0.4), "DVOA_Proxy_Off": (0.0, 10.0),
}
THRESHOLDS_DEF = {
    "SACKs": (2.0, 3.0), "INTs": (0.6, 1.2), "FF": (0.4, 0.8), "FR": (0.4, 0.8),
    "QB Hits": (5.0, 8.0), "Pressures": (14.0, 20.0),
    "3D% Allowed": (42.0, 34.0), "RZ% Allowed": (60.0, 50.0), "EPA/Play Allowed": (0.00, -0.05),
    "Success Rate Allowed": (47.0, 41.0), "YPA Allowed": (7.5, 6.7), "YPC Allowed": (4.7, 4.0), "CMP% Allowed": (66.0, 62.0),
    "DVOA": (5.0, -5.0), "DVOA_Proxy_Def": (5.0, -5.0),
}

def _style_thresholds(df, thresholds, invert_for_cols=None):
    invert_for_cols = invert_for_cols or set()
    styles = pd.DataFrame("", index=df.index, columns=df.columns)

    def color_cell(val, low, high, higher_is_better=True):
        try:
            v = float(val)
        except Exception:
            return ""
        if higher_is_better:
            if v >= high: return "background-color: #d9f2d9"
            if v <= low:  return "background-color: #f8d7da"
        else:
            if v <= high: return "background-color: #d9f2d9"
            if v >= low:  return "background-color: #f8d7da"
        return ""

    for col, (low, high) in thresholds.items():
        if col not in df.columns:
            continue
        higher_is_better = not (("Allowed" in col) or (col in invert_for_cols))
        styles[col] = df[col].apply(lambda v: color_cell(v, low, high, higher_is_better))
    return styles

st.markdown("## 🎨 Color Settings")
with st.expander("Color thresholds for previews & weekly NFL tables", expanded=False):
    enable_colors = st.checkbox("Enable color thresholds (green = better, red = worse)", value=True)

    if enable_colors:
        off_prev = st.session_state.get("offense_preview_df")
        def_prev = st.session_state.get("defense_preview_df")

        if isinstance(off_prev, pd.DataFrame) and not off_prev.empty:
            st.subheader("Offense Preview (with thresholds)")
            st.dataframe(off_prev.style.apply(lambda _: _style_thresholds(off_prev, THRESHOLDS_OFF, {"SACKs Allowed","INTs Thrown","Fumbles Lost"}), axis=None),
                         use_container_width=True)

        if isinstance(def_prev, pd.DataFrame) and not def_prev.empty:
            st.subheader("Defense Preview (with thresholds)")
            st.dataframe(def_prev.style.apply(lambda _: _style_thresholds(def_prev, THRESHOLDS_DEF), axis=None),
                         use_container_width=True)

        wk_off = st.session_state.get("weekly_off_tbl")
        wk_def = st.session_state.get("weekly_def_tbl")

        if isinstance(wk_off, pd.DataFrame) and not wk_off.empty:
            st.subheader("Weekly NFL Averages – Offense (with thresholds)")
            st.dataframe(wk_off.style.apply(lambda _: _style_thresholds(wk_off, THRESHOLDS_OFF), axis=None), use_container_width=True)

        if isinstance(wk_def, pd.DataFrame) and not wk_def.empty:
            st.subheader("Weekly NFL Averages – Defense (with thresholds)")
            st.dataframe(wk_def.style.apply(lambda _: _style_thresholds(wk_def, THRESHOLDS_DEF), axis=None), use_container_width=True)

# =========================================================
# 6) 🧾 PDF EXPORTS (colored, auto-named)
# =========================================================
def _rgb_for_cell(value, low, high, higher_is_better=True):
    try:
        v = float(value)
    except Exception:
        return None
    if higher_is_better:
        if v >= high: return (217, 242, 217)  # green
        if v <= low:  return (248, 215, 218)  # red
    else:
        if v <= high: return (217, 242, 217)
        if v >= low:  return (248, 215, 218)
    return None

def _higher_is_better_for(col):
    return not (("Allowed" in col) or (col in {"SACKs Allowed","INTs Thrown","Fumbles Lost"}))

def _df_subset_for_pdf(df, thresholds_dict, include_week=True, max_cols=10):
    if df is None or df.empty:
        return pd.DataFrame()
    cols = []
    if include_week and "Week" in df.columns:
        cols.append("Week")
    metric_cols = [c for c in thresholds_dict.keys() if c in df.columns]
    cols.extend(metric_cols)
    if len(cols) > max_cols:
        cols = cols[:max_cols]
    return df[cols].copy()

def _add_table(pdf, title, df, thresholds_dict):
    if df is None or df.empty:
        return
    pdf.add_page(orientation="L")
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, txt=title, ln=1)

    pdf.set_font("Arial", "B", 10)
    col_count = len(df.columns)
    table_width = 270
    col_w = table_width / max(col_count, 1)
    row_h = 8

    # header
    for col in df.columns:
        pdf.set_fill_color(230, 230, 230)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(col_w, row_h, txt=str(col), border=1, ln=0, align="C", fill=True)
    pdf.ln(row_h)

    # rows
    pdf.set_font("Arial", "", 9)
    for _, row in df.iterrows():
        for col in df.columns:
            text = "" if pd.isna(row[col]) else str(row[col])
            fill_rgb = None
            if col != "Week" and col in thresholds_dict:
                low, high = thresholds_dict[col]
                fill_rgb = _rgb_for_cell(row[col], low, high, _higher_is_better_for(col))
            if fill_rgb:
                pdf.set_fill_color(*fill_rgb)
                fill_flag = True
            else:
                fill_flag = False
            pdf.cell(col_w, row_h, txt=text, border=1, ln=0, align="C", fill=fill_flag)
        pdf.ln(row_h)

def _export_pdf(filename, tables):
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    for t in tables:
        df = _df_subset_for_pdf(t["df"], t["thresholds"], include_week=t.get("include_week", True))
        _add_table(pdf, t["title"], df, t["thresholds"])
    pdf.output(filename)

st.markdown("## 🧾 PDF Exports")
with st.expander("Download colored PDFs (league tables & previews)", expanded=False):
    # Use the same week/phase picked in the sidebar for consistency
    wk = int(st.session_state.get("week_num_global", 0) or week_num_global)
    ph = st.session_state.get("phase_global", "") or phase_global

    st.caption(f"Using Week {wk} — Phase '{ph}' for file names (change in sidebar).")

    wk_off = st.session_state.get("weekly_off_tbl")
    wk_def = st.session_state.get("weekly_def_tbl")

    if isinstance(wk_off, pd.DataFrame) and not wk_off.empty:
        if st.button("📥 Download NFL Averages (PDF)"):
            try:
                fname = make_export_filename(wk, ph, "pdf")
                _export_pdf(
                    fname,
                    [
                        {"title": f"Weekly NFL Averages – Offense (W{wk:02d}, {ph})", "df": wk_off, "thresholds": THRESHOLDS_OFF, "include_week": True},
                        {"title": f"Weekly NFL Averages – Defense (W{wk:02d}, {ph})", "df": wk_def if isinstance(wk_def, pd.DataFrame) else pd.DataFrame(), "thresholds": THRESHOLDS_DEF, "include_week": True},
                    ],
                )
                with open(fname, "rb") as f:
                    st.download_button(f"Download {fname}", f.read(), file_name=fname, mime="application/pdf")
            except Exception as e:
                st.error(f"PDF export failed: {e}")
    else:
        st.info("Run **Compute NFL Averages** first to enable the league PDF download.")

    off_prev = st.session_state.get("offense_preview_df")
    def_prev = st.session_state.get("defense_preview_df")

    if isinstance(off_prev, pd.DataFrame) and not off_prev.empty:
        if st.button("📥 Download Bears vs NFL Previews (PDF)"):
            try:
                fname_prev = make_export_filename(wk, ph, "pdf").replace(".pdf", "_previews.pdf")
                _export_pdf(
                    fname_prev,
                    [
                        {"title": f"Bears Offense Preview (W{wk:02d}, {ph})", "df": off_prev, "thresholds": THRESHOLDS_OFF, "include_week": False},
                        {"title": f"Bears Defense Preview (W{wk:02d}, {ph})", "df": def_prev if isinstance(def_prev, pd.DataFrame) else pd.DataFrame(), "thresholds": THRESHOLDS_DEF, "include_week": False},
                    ],
                )
                with open(fname_prev, "rb") as f:
                    st.download_button(f"Download {fname_prev}", f.read(), file_name=fname_prev, mime="application/pdf")
            except Exception as e:
                st.error(f"PDF export failed: {e}")
    else:
        st.info("Previews not found in session. Visit the preview sections (and enable NFL Avg merge) first.")

# =========================================================
# 7) 💾 EXPORT (bottom of page, auto-named)
# =========================================================
st.markdown("## 💾 Export")
wk = int(st.session_state.get("week_num_global", 0) or week_num_global)
ph = st.session_state.get("phase_global", "") or phase_global

if os.path.exists(EXCEL_FILE):
    try:
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                "⬇️ Download All Data (Excel)",
                data=f.read(),
                file_name=make_export_filename(wk, ph, "xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Export failed: {e}")
else:
    st.info("No Excel file yet. Upload something or compute averages to create it.")
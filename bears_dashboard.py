import os
import streamlit as st
import pandas as pd
from fpdf import FPDF

# ----------------------------
# App header
# ----------------------------
st.set_page_config(page_title="Chicago Bears 2025-26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, league comparisons, and advanced metrics.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"
TEAM = "CHI"
SEASON = 2025

# ----------------------------
# Helpers
# ----------------------------
def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True, keys=("Week",)):
    """Append or replace a sheet with optional deduping by key columns."""
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

                existing.columns = existing.iloc[0]
    try:
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                # Load existing sheet into DataFrame
                ws = book[sheet_name]
                existing = pd.DataFrame(ws.values)
                if existing.shape[0] > 0:
                    existing.columns = existing.iloc[0]
                    existing = existing[1:]
                else:
                    existing = pd.DataFrame()

                # Combine with new_data
                combined = pd.concat([existing, new_data], ignore_index=True)

                # Deduplicate if requested and keys present
                if deduplicate and all(k in combined.columns for k in keys):
                    combined = combined.drop_duplicates(subset=list(keys), keep="last")
            else:
                combined = new_data.copy()
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            combined = new_data.copy()

        # Replace sheet
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        ws = book.create_sheet(sheet_name)
        for r in dataframe_to_rows(combined, index=False, header=True):
            ws.append(r)
        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def load_sheet(name: str) -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=name)
    except Exception:
        return pd.DataFrame()

def colorize_df(df: pd.DataFrame, context: str):
    """Apply red/green highlighting for readability.
    context: 'offense' or 'defense'
    """
    if df.empty:
        return df

    # Default benchmarks (tune later in app if you want)
    # Explosive rate ~ 12% league avg; DSR ~ 72% league avg
    exp_avg = 12.0
    dsr_avg = 72.0

    def style_func(val, positive_good=True, avg=0):
        try:
            v = float(val)
        except Exception:
            return ""
        if positive_good:
            # green if >= avg, red if lower
            return "background-color: #e6ffe6" if v >= avg else "background-color: #ffe6e6"
        else:
            # green if <= avg, red if higher
            return "background-color: #e6ffe6" if v <= avg else "background-color: #ffe6e6"

    styled = df.style

    # Offense columns (higher is better)
    off_cols_pos_good = [c for c in df.columns if c.lower() in
                         ["explosive_play_rate", "drive_success_rate"]]
    for c in off_cols_pos_good:
        if c in df.columns:
            styled = styled.map(lambda v, avg=exp_avg if "Explosive" in c else dsr_avg:
                                style_func(v, positive_good=True, avg=avg), subset=[c])

    # Defense allowed columns (lower is better)
    def_cols_neg_good = [c for c in df.columns if c.lower() in
                         ["explosive_play_rate_allowed", "drive_success_rate_allowed"]]
    for c in def_cols_neg_good:
        if c in df.columns:
            styled = styled.map(lambda v, avg=exp_avg if "Explosive" in c else dsr_avg:
                                style_func(v, positive_good=False, avg=avg), subset=[c])

    return styled

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

# ----------------------------
# Sidebar: Uploads
# ----------------------------
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense   = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy  = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")
uploaded_injuries  = st.sidebar.file_uploader("Upload Injuries (.csv) [optional]", type="csv")

if uploaded_offense:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense", deduplicate=True)
    st.sidebar.success("✅ Offensive data uploaded.")
if uploaded_defense:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense", deduplicate=True)
    st.sidebar.success("✅ Defensive data uploaded.")
if uploaded_strategy:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy", deduplicate=True)
    st.sidebar.success("✅ Strategy data uploaded.")
if uploaded_personnel:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel", deduplicate=True)
    st.sidebar.success("✅ Personnel data uploaded.")
if uploaded_injuries:
    df_inj = pd.read_csv(uploaded_injuries)
    append_to_excel(df_inj, "Injuries", deduplicate=True, keys=("Week","Player"))
    st.sidebar.success("✅ Injuries uploaded.")

# ----------------------------
# Sidebar: Fetch weekly team summary (basic)
# ----------------------------
with st.sidebar.expander("⚡ Fetch Weekly Team Data (nfl_data_py)"):
    st.caption(f"Pulls {SEASON} weekly team stats for CHI and saves to Excel (basic fields).")
    fetch_week = st.number_input("Week to fetch", min_value=1, max_value=25, value=1, step=1, key="fetch_week_basic")
    if st.button("Fetch CHI Week (basic)"):
        try:
            import nfl_data_py as nfl

            try:
                nfl.update.weekly_data([SEASON])
            except Exception:
                pass

            weekly = nfl.import_weekly_data([SEASON])
            wk = int(fetch_week)
            team_week = weekly[(weekly["team"] == TEAM) & (weekly["week"] == wk)].copy()

            if team_week.empty:
                st.warning("No weekly team row found for CHI in that week. Try again later.")
            else:
                team_week["Week"] = wk

                # Map a few offense/defense basics if present
                # Offense
                pass_yards = team_week["passing_yards"].iloc[0] if "passing_yards" in team_week.columns else None
                pass_att = None
                for cand in ["attempts", "passing_attempts", "pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].iloc[0]
                        break
                try:
                    ypa_val = float(pass_yards) / float(pass_att) if pass_yards is not None and pass_att not in (None, 0) else None
                except Exception:
                    ypa_val = None

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
                    "CMP%": cmp_pct
                }])

                # Defense
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].iloc[0]
                        break

                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val
                }])

                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)

                st.success(f"✅ Added CHI week {wk} basic fields to Offense/Defense.")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

# ----------------------------
# Sidebar: Fetch play-by-play metrics (Advanced)
# ----------------------------
st.sidebar.markdown("### 📡 Fetch Advanced Metrics from Play-by-Play")
pbp_week = st.sidebar.number_input("Week to fetch", min_value=1, max_value=25, value=1, step=1, key="pbp_week")

if st.sidebar.button("Fetch PBP: RZ%, Success, Pressures, Explosive, DSR"):
    try:
        import nfl_data_py as nfl
        wk = int(pbp_week)

        pbp = nfl.import_pbp_data([SEASON], downcast=False)
        # real plays only
        plays = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False))
        ].copy()

        # ---------- DEFENSE (CHI on defense) ----------
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == TEAM)].copy()
        if len(bears_def) == 0:
            st.warning("No CHI defensive PBP rows found for that week.")
        else:
            # Red Zone % Allowed: drives with min_yardline_100 <= 20
            dmins = (
                bears_def.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success rate allowed (opponent success vs CHI)
            def_success = bears_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            def_success_pct = round(def_success * 100, 1) if pd.notna(def_success) else 0.0

            # Pressures = sacks + qb_hit flags
            qb_hits = bears_def["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in bears_def.columns else 0
            sacks = bears_def["sack"].fillna(0).astype(int).sum() if "sack" in bears_def.columns else 0
            pressures = int(qb_hits + sacks)

            # Explosive plays allowed: rush >=10 or pass >=20
            def_exp = bears_def[
                ((bears_def["rush_attempt"].fillna(0) == 1) & (bears_def["yards_gained"].fillna(0) >= 10)) |
                ((bears_def["pass_attempt"].fillna(0) == 1) & (bears_def["yards_gained"].fillna(0) >= 20))
            ]
            total_def_plays = len(bears_def)
            def_explosive_rate = round(len(def_exp) / total_def_plays * 100, 1) if total_def_plays else 0.0

            # Drive Success Rate (allowed): pct of opponent drives with any first down or points
            def_drives = bears_def.groupby(["game_id", "drive"], as_index=False).agg(
                any_fd=("first_down", lambda s: (s.fillna(0) == 1).any()),
                any_points=("touchdown", lambda s: (s.fillna(0) == 1).any())
            )
            # If FG or safety not flagged, "any_points" above is minimal; we treat first downs OR touchdowns as success proxy.
            def_dsr = round((def_drives["any_fd"] | def_drives["any_points"]).mean() * 100, 1) if len(def_drives) else 0.0

            adv_def = pd.DataFrame([{
                "Week": wk,
                "RZ% Allowed": round(rz_allowed, 1),
                "Success Rate% (Offense)": def_success_pct,  # offense success against CHI
                "Pressures": pressures,
                "Explosive_Play_Rate_Allowed": def_explosive_rate,
                "Drive_Success_Rate_Allowed": def_dsr
            }])
            append_to_excel(adv_def, "Advanced_Defense", deduplicate=True, keys=("Week",))

        # ---------- OFFENSE (CHI on offense) ----------
        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == TEAM)].copy()
        if len(bears_off) == 0:
            st.info("No CHI offensive PBP rows found for that week.")
        else:
            # Explosive plays: rush >=10 or pass >=20
            off_exp = bears_off[
                ((bears_off["rush_attempt"].fillna(0) == 1) & (bears_off["yards_gained"].fillna(0) >= 10)) |
                ((bears_off["pass_attempt"].fillna(0) == 1) & (bears_off["yards_gained"].fillna(0) >= 20))
            ]
            total_off_plays = len(bears_off)
            off_explosive_rate = round(len(off_exp) / total_off_plays * 100, 1) if total_off_plays else 0.0

            # Drive Success Rate (offense): pct of CHI drives with any first down or points
            off_drives = bears_off.groupby(["game_id", "drive"], as_index=False).agg(
                any_fd=("first_down", lambda s: (s.fillna(0) == 1).any()),
                any_points=("touchdown", lambda s: (s.fillna(0) == 1).any())
            )
            off_dsr = round((off_drives["any_fd"] | off_drives["any_points"]).mean() * 100, 1) if len(off_drives) else 0.0

            adv_off = pd.DataFrame([{
                "Week": wk,
                "Explosive_Play_Rate": off_explosive_rate,
                "Drive_Success_Rate": off_dsr
            }])
            append_to_excel(adv_off, "Advanced_Offense", deduplicate=True, keys=("Week",))

        st.success(f"✅ Advanced metrics saved for Week {wk}.")
    except Exception as e:
        st.error(f"❌ Failed to fetch PBP metrics: {e}")

# ----------------------------
# Download full Excel
# ----------------------------
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ----------------------------
# Visible data sections (merge advanced where available)
# ----------------------------
st.subheader("📊 Offensive Analytics")
df_off = load_sheet("Offense")
df_adv_off = load_sheet("Advanced_Offense")
if not df_off.empty and not df_adv_off.empty and "Week" in df_off.columns and "Week" in df_adv_off.columns:
    off_display = df_off.merge(df_adv_off, on="Week", how="left")
else:
    off_display = df_off if not df_off.empty else df_adv_off

if not off_display.empty:
    try:
        st.dataframe(colorize_df(off_display, "offense"), use_container_width=True)
    except Exception:
        st.dataframe(off_display, use_container_width=True)
else:
    st.info("No offense data yet.")

st.subheader("🛡️ Defensive Analytics")
df_def = load_sheet("Defense")
df_adv_def = load_sheet("Advanced_Defense")
if not df_def.empty and not df_adv_def.empty and "Week" in df_def.columns and "Week" in df_adv_def.columns:
    def_display = df_def.merge(df_adv_def, on="Week", how="left")
else:
    def_display = df_def if not df_def.empty else df_adv_def

if not def_display.empty:
    try:
        st.dataframe(colorize_df(def_display, "defense"), use_container_width=True)
    except Exception:
        st.dataframe(def_display, use_container_width=True)
else:
    st.info("No defense data yet.")

# ----------------------------
# Media summaries
# ----------------------------
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

# Show media
df_media = load_sheet("Media_Summaries")
if not df_media.empty:
    st.subheader("📰 Saved Media Summaries")
    st.dataframe(df_media, use_container_width=True)

# ----------------------------
# DVOA-like Proxy (opponent-adjusted EPA/SR)
# ----------------------------
st.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.number_input("Week to Compute (proxy)", min_value=1, max_value=25, value=1, step=1, key="proxy_week")

if st.button("Compute DVOA-like Proxy"):
    try:
        import nfl_data_py as nfl

        wk = int(proxy_week)
        pbp = nfl.import_pbp_data([SEASON], downcast=False)

        plays = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False)) &
            (~pbp["epa"].isna())
        ].copy()

        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == TEAM)].copy()
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == TEAM)].copy()

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

            # Opponent defensive benchmarks (allowed vs them)
            opp_def_plays = prior[prior["defteam"] == opponent].copy()
            opp_def_epa = opp_def_plays["epa"].mean() if len(opp_def_plays) else None
            if len(opp_def_plays):
                opp_def_success = opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                opp_def_success = None

            # Opponent offensive benchmarks (their offense)
            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            if len(opp_off_plays):
                opp_off_success = opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                opp_off_success = None

            # Bears week EPA/SR on offense & defense
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

            append_to_excel(out, "DVOA_Proxy", deduplicate=True, keys=("Week",))
            st.success(
                f"✅ DVOA-like proxy saved for Week {wk} vs {opponent} "
                f"(Off Adj EPA/play={out.iloc[0]['Off Adj EPA/play']}, Off Adj SR%={out.iloc[0]['Off Adj SR%']}, "
                f"Def Adj EPA/play={out.iloc[0]['Def Adj EPA/play']}, Def Adj SR%={out.iloc[0]['Def Adj SR%']})"
            )
    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# Preview DVOA Proxy
df_dvoa = load_sheet("DVOA_Proxy")
if not df_dvoa.empty:
    st.subheader("📊 DVOA-like Proxy Metrics")
    st.dataframe(df_dvoa.tail(10), use_container_width=True)

# ----------------------------
# Weekly Prediction (simple rule set)
# ----------------------------
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
        df_strategy = load_sheet("Strategy")
        df_offense  = load_sheet("Offense")
        df_defense  = load_sheet("Defense")
        df_advdef   = load_sheet("Advanced_Defense")
        df_proxy    = load_sheet("DVOA_Proxy")

        row_s = df_strategy[df_strategy["Week"] == week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"] == week_to_predict]  if not df_offense.empty  else pd.DataFrame()
        row_d = df_defense[df_defense["Week"] == week_to_predict]  if not df_defense.empty  else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"] == week_to_predict]    if not df_advdef.empty   else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"] == week_to_predict]      if not df_proxy.empty    else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            # Prefer Advanced_Defense values when available
            rz_allowed = _safe_float(row_a.iloc[0].get("RZ% Allowed"), default=None) if not row_a.empty else None
            pressures  = _safe_float(row_a.iloc[0].get("Pressures"), default=None)   if not row_a.empty else None
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), default=None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), default=None)

            # Rule set
            reason_bits = []
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win - efficiency edge on both sides"
                reason_bits.append(f"Off+{off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play vs opp O")
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win - pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win - red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient offense and poor red zone defense"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient passing and weak red zone defense"
                reason_bits.append(f"YPA={ypa:.1f}")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            else:
                prediction = "Loss - no clear advantage in key strategy or stats"
                if off_adj_epa is not None:
                    reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                if def_adj_epa is not None:
                    reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play")
                if pressures is not None:
                    reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            reason_text = " | ".join(reason_bits)
            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")
            if reason_text:
                st.caption(reason_text)

            # Save prediction
            pred_save = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("-")[0].strip(),
                "Reason": prediction[prediction.find("-")+1:].strip() if "-" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(pred_save, "Predictions", deduplicate=True, keys=("Week",))
        else:
            st.info("Please upload/fetch Strategy, Offense, and Defense for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check data. Error: {e}")

# Show saved predictions
df_preds = load_sheet("Predictions")
if not df_preds.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(df_preds, use_container_width=True)
else:
    st.info("No predictions saved yet.")
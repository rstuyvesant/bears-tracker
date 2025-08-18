import os
import pandas as pd
import streamlit as st
from fpdf import FPDF

# =========================
# App setup
# =========================
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, league comparisons, and auto-fetched data.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# =========================
# Utilities
# =========================
def _coerce_week(val):
    """Return an int week if possible; else None."""
    try:
        if pd.isna(val):
            return None
        return int(val)
    except Exception:
        try:
            return int(float(val))
        except Exception:
            return None

def append_to_excel(new_df: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True):
    """
    Append/replace a sheet in the Excel workbook.
    If deduplicate=True and 'Week' column exists, drop existing rows for those weeks before writing.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    # Ensure Week is normalized (if present)
    if "Week" in new_df.columns:
        new_df = new_df.copy()
        new_df["Week"] = new_df["Week"].apply(_coerce_week)

    try:
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
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

        # Deduplicate by Week if both have Week
        if deduplicate and not existing.empty and "Week" in existing.columns and "Week" in new_df.columns:
            existing["Week"] = existing["Week"].apply(_coerce_week)
            weeks_in_new = set(new_df["Week"].dropna().tolist())
            if weeks_in_new:
                existing = existing[~existing["Week"].isin(weeks_in_new)]

        combined = pd.concat([existing, new_df], ignore_index=True)

        # Replace the sheet
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        for r in dataframe_to_rows(combined, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def read_sheet(name: str) -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=name)
    except Exception:
        return pd.DataFrame()

# =========================
# Color styling helpers (NO type annotations here)
# =========================
def style_offense(df: pd.DataFrame):
    """Highlights Offense sheet (YPA, CMP%)."""
    def ypa_color(val):
        try:
            v = float(val)
        except Exception:
            return ""
        if v >= 7.5: return "background-color: #e6f4ea"  # green
        if v >= 6.0: return "background-color: #fff8e1"  # yellow
        return "background-color: #fdecea"               # red

    def cmp_color(val):
        try:
            v = float(val)
        except Exception:
            return ""
        if v >= 67.0: return "background-color: #e6f4ea"
        if v >= 60.0: return "background-color: #fff8e1"
        return "background-color: #fdecea"

    styler = df.style
    if "YPA" in df.columns:
        styler = styler.applymap(ypa_color, subset=["YPA"])
    if "CMP%" in df.columns:
        styler = styler.applymap(cmp_color, subset=["CMP%"])
    return styler

def style_defense(df: pd.DataFrame):
    """Highlights Defense sheet (RZ% Allowed, SACK)."""
    def rz_color(val):
        try:
            v = float(val)
        except Exception:
            return ""
        if v < 50:  return "background-color: #e6f4ea"
        if v <= 65: return "background-color: #fff8e1"
        return "background-color: #fdecea"

    def sack_color(val):
        try:
            v = float(val)
        except Exception:
            return ""
        if v >= 4:  return "background-color: #e6f4ea"
        if v >= 2:  return "background-color: #fff8e1"
        return "background-color: #fdecea"

    styler = df.style
    if "RZ% Allowed" in df.columns:
        styler = styler.applymap(rz_color, subset=["RZ% Allowed"])
    if "SACK" in df.columns:
        styler = styler.applymap(sack_color, subset=["SACK"])
    return styler

def style_personnel(df: pd.DataFrame):
    """Highlights Personnel usage percentages."""
    def usage_color(val):
        try:
            v = float(val)
        except Exception:
            return ""
        if v >= 60: return "background-color: #e6f4ea"
        if v >= 30: return "background-color: #fff8e1"
        return "background-color: #f3f4f6"

    styler = df.style
    for col in df.columns:
        if any(tag in str(col) for tag in ["11", "12", "13", "21"]):
            styler = styler.applymap(usage_color, subset=[col])
    return styler

def style_dvoa(df: pd.DataFrame):
    """Highlights proxy outputs (Adj EPA/play)."""
    def off_adj_epa_color(val):
        try:
            v = float(val)
        except Exception:
            return ""
        if v >= 0.15: return "background-color: #e6f4ea"
        if v >= 0.05: return "background-color: #fff8e1"
        return "background-color: #fdecea"

    def def_adj_epa_color(val):
        try:
            v = float(val)
        except Exception:
            return ""
        if v <= -0.05: return "background-color: #e6f4ea"
        if v <= 0.00:  return "background-color: #fff8e1"
        return "background-color: #fdecea"

    styler = df.style
    if "Off Adj EPA/play" in df.columns:
        styler = styler.applymap(off_adj_epa_color, subset=["Off Adj EPA/play"])
    if "Def Adj EPA/play" in df.columns:
        styler = styler.applymap(def_adj_epa_color, subset=["Def Adj EPA/play"])
    return styler

# =========================
# Sidebar: Uploads
# =========================
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv", key="u_off")
uploaded_defense   = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv", key="u_def")
uploaded_strategy  = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv", key="u_strat")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv", key="u_pers")

if uploaded_offense:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense")
    st.sidebar.success("✅ Offensive data uploaded.")

if uploaded_defense:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense")
    st.sidebar.success("✅ Defensive data uploaded.")

if uploaded_strategy:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy")
    st.sidebar.success("✅ Strategy data uploaded.")

if uploaded_personnel:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel")
    st.sidebar.success("✅ Personnel data uploaded.")

# =========================
# Sidebar: Auto-fetchers
# =========================
with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pulls 2025 weekly team stats for CHI and saves to Excel (best effort).")
    fetch_week = st.number_input("Week to fetch (2025 season)", min_value=1, max_value=25, value=1, step=1)
    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl

            try:
                nfl.update.schedule_data([2025])
            except Exception:
                pass
            try:
                nfl.update.weekly_data([2025])
            except Exception:
                pass

            weekly = nfl.import_weekly_data([2025])
            wk = int(fetch_week)
            team_week = weekly[(weekly["team"] == "CHI") & (weekly["week"] == wk)].copy()

            if team_week.empty:
                st.warning("No weekly team row found for CHI in that week yet.")
            else:
                team_week["Week"] = wk

                # Simple offense derivations
                pass_yards = team_week["passing_yards"].values[0] if "passing_yards" in team_week.columns else None
                pass_att   = None
                for cand in ["attempts", "passing_attempts", "pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].values[0]
                        break
                try:
                    ypa_val = float(pass_yards) / float(pass_att) if pass_yards is not None and pass_att not in (None, 0) else None
                except Exception:
                    ypa_val = None

                yards_total = None
                for cand in ["yards", "total_yards", "offense_yards"]:
                    if cand in team_week.columns:
                        yards_total = team_week[cand].values[0]
                        break

                completions = None
                for cand in ["completions", "passing_completions", "pass_completions"]:
                    if cand in team_week.columns:
                        completions = team_week[cand].values[0]
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

                # Defense
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].values[0]
                        break

                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val,
                    "RZ% Allowed": None  # needs PBP
                }])

                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)

                st.success(f"✅ Added CHI week {wk} to Offense/Defense (available fields).")
                st.caption("Note: Red Zone % Allowed and pressures require play-by-play aggregation (see next panel).")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

st.sidebar.markdown("---")

st.sidebar.markdown("### 📡 Fetch Defensive Metrics from Play-by-Play")
pbp_week = st.sidebar.number_input("Week to Fetch (2025 Season)", min_value=1, max_value=25, value=1, step=1)
if st.sidebar.button("Fetch Play-by-Play Metrics"):
    try:
        import nfl_data_py as nfl
        pbp = nfl.import_pbp_data([2025], downcast=False)
        pbp_w = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == "CHI")].copy()

        if pbp_w.empty:
            st.warning("No PBP rows for CHI defense in that week yet.")
        else:
            # Red Zone % Allowed (drive reaches <= 20)
            dmins = (
                pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success Rate (opposing offense success vs CHI)
            def play_success(row):
                if pd.isna(row.get("down")) or pd.isna(row.get("ydstogo")) or pd.isna(row.get("yards_gained")):
                    return False
                d = int(row["down"]); togo = float(row["ydstogo"]); gain = float(row["yards_gained"])
                if d == 1: return gain >= 0.4 * togo
                if d == 2: return gain >= 0.6 * togo
                return gain >= togo

            plays_mask = (~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))
            pbp_real = pbp_w[plays_mask].copy()
            success_rate = pbp_real.apply(play_success, axis=1).mean() * 100 if len(pbp_real) else 0.0

            # Pressures approx = QB hits + sacks
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

            st.success(
                f"✅ Week {int(pbp_week)} PBP metrics — RZ% Allowed: {rz_allowed:.1f} | "
                f"Success Rate% (Off): {success_rate:.1f} | Pressures: {pressures}"
            )
    except Exception as e:
        st.error(f"❌ Failed to fetch PBP metrics: {e}")

st.sidebar.markdown("---")

# =========================
# DVOA-like proxy (opponent-adjusted EPA/SR)
# =========================
st.sidebar.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.sidebar.number_input("Week to Compute", min_value=1, max_value=25, value=1, step=1, key="proxy_week")

def _success_flag(down, ydstogo, yards_gained):
    try:
        if pd.isna(down) or pd.isna(ydstogo) or pd.isna(yards_gained):
            return False
        d = int(down); togo = float(ydstogo); gain = float(yards_gained)
        if d == 1: return gain >= 0.4 * togo
        if d == 2: return gain >= 0.6 * togo
        return gain >= togo
    except Exception:
        return False

if st.sidebar.button("Compute DVOA-like Proxy"):
    try:
        import nfl_data_py as nfl
        wk = int(proxy_week)
        pbp = nfl.import_pbp_data([2025], downcast=False)

        plays = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False)) &
            (~pbp["epa"].isna())
        ].copy()

        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == "CHI")].copy()
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == "CHI")].copy()

        if bears_off.empty and bears_def.empty:
            st.warning("No CHI plays found for that week yet.")
        else:
            opps = set()
            if not bears_off.empty: opps.update(bears_off["defteam"].dropna().unique().tolist())
            if not bears_def.empty: opps.update(bears_def["posteam"].dropna().unique().tolist())
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
            st.success(
                f"✅ DVOA-like proxy saved for Week {wk} vs {opponent} "
                f"(Off Adj EPA/play={out.iloc[0]['Off Adj EPA/play']}, Off Adj SR%={out.iloc[0]['Off Adj SR%']}, "
                f"Def Adj EPA/play={out.iloc[0]['Def Adj EPA/play']}, Def Adj SR%={out.iloc[0]['Def Adj SR%']})"
            )
    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# =========================
# Download workbook (sidebar)
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
# Main panels: show current data (styled)
# =========================
df_off = read_sheet("Offense")
df_def = read_sheet("Defense")
df_str = read_sheet("Strategy")
df_per = read_sheet("Personnel")
df_adv = read_sheet("Advanced_Defense")
df_dvo = read_sheet("DVOA_Proxy")

if not df_off.empty:
    st.subheader("📊 Offensive Analytics")
    try:
        st.dataframe(style_offense(df_off), use_container_width=True)
    except Exception:
        st.dataframe(df_off, use_container_width=True)

if not df_def.empty:
    st.subheader("🛡️ Defensive Analytics")
    try:
        st.dataframe(style_defense(df_def), use_container_width=True)
    except Exception:
        st.dataframe(df_def, use_container_width=True)

if not df_per.empty:
    st.subheader("👥 Personnel Usage")
    try:
        st.dataframe(style_personnel(df_per), use_container_width=True)
    except Exception:
        st.dataframe(df_per, use_container_width=True)

if not df_adv.empty:
    st.subheader("📡 Advanced Defensive Metrics (from PBP)")
    st.dataframe(df_adv, use_container_width=True)

if not df_dvo.empty:
    st.subheader("📈 DVOA-like Proxy (Opponent-Adjusted)")
    try:
        st.dataframe(style_dvoa(df_dvo), use_container_width=True)
    except Exception:
        st.dataframe(df_dvo, use_container_width=True)

# =========================
# Media summaries
# =========================
st.markdown("### 📰 Weekly Beat Writer / ESPN Summary")
with st.form("media_form"):
    media_week = st.number_input("Week", min_value=1, max_value=25, step=1, key="media_week_input")
    media_opponent = st.text_input("Opponent")
    media_summary = st.text_area("Beat Writer & ESPN Summary (Game Recap, Analysis, Strategy, etc.)")
    submit_media = st.form_submit_button("Save Summary")

if submit_media:
    media_df = pd.DataFrame([{"Week": media_week, "Opponent": media_opponent, "Summary": media_summary}])
    append_to_excel(media_df, "Media_Summaries", deduplicate=False)
    st.success(f"✅ Summary for Week {media_week} vs {media_opponent} saved.")

df_media = read_sheet("Media_Summaries")
if not df_media.empty:
    st.subheader("📰 Saved Media Summaries")
    st.dataframe(df_media, use_container_width=True)

# =========================
# Prediction
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
        # base
        df_strategy = read_sheet("Strategy")
        df_offense  = read_sheet("Offense")
        df_defense  = read_sheet("Defense")
        # optional
        df_advdef = read_sheet("Advanced_Defense")
        df_proxy  = read_sheet("DVOA_Proxy")

        row_s = df_strategy[df_strategy["Week"].apply(_coerce_week) == int(week_to_predict)] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"].apply(_coerce_week) == int(week_to_predict)] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"].apply(_coerce_week) == int(week_to_predict)] if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"].apply(_coerce_week) == int(week_to_predict)] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"].apply(_coerce_week) == int(week_to_predict)] if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            rz_allowed = None
            pressures = None
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

            # Strong two-way efficiency edge
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win – efficiency edge on both sides"
                reason_bits.append(f"Off+{off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play vs opp O")

            # Pass-rush advantage
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            # Coverage + red zone discipline
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win – red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            # Clear drag
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient offense and poor red zone defense"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            # Legacy fallback
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient passing and weak red zone defense"
                reason_bits.append(f"YPA={ypa:.1f}")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            else:
                prediction = "Loss – no clear advantage in key strategy or stats"
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
            pred_outcome = prediction.split("–")[0].strip() if "–" in prediction else prediction.split("-")[0].strip()
            pred_reason  = prediction.split("–")[1].strip() if "–" in prediction else (prediction.split("-", 1)[1].strip() if "-" in prediction else "")

            prediction_entry = pd.DataFrame([{
                "Week": int(week_to_predict),
                "Prediction": pred_outcome,
                "Reason": pred_reason,
                "Notes": reason_text
            }])
            append_to_excel(prediction_entry, "Predictions", deduplicate=True)
        else:
            st.info("Please upload or fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# =========================
# Saved predictions
# =========================
df_preds = read_sheet("Predictions")
if not df_preds.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(df_preds, use_container_width=True)

# =========================
# Simple PDF export (current week)
# =========================
st.markdown("### 🧾 Download Weekly Game Report (PDF)")
report_week = st.number_input("Select Week for Report", min_value=1, max_value=25, step=1, key="report_week")

if st.button("Generate Weekly Report"):
    try:
        df_strategy = read_sheet("Strategy")
        df_offense  = read_sheet("Offense")
        df_defense  = read_sheet("Defense")
        df_media    = read_sheet("Media_Summaries")
        df_preds    = read_sheet("Predictions")

        strat_row = df_strategy[df_strategy["Week"].apply(_coerce_week) == int(report_week)]
        off_row   = df_offense[df_offense["Week"].apply(_coerce_week) == int(report_week)]
        def_row   = df_defense[df_defense["Week"].apply(_coerce_week) == int(report_week)]
        media_rows= df_media[df_media["Week"].apply(_coerce_week) == int(report_week)] if not df_media.empty else pd.DataFrame()
        pred_row  = df_preds[df_preds["Week"].apply(_coerce_week) == int(report_week)] if not df_preds.empty else pd.DataFrame()

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, f"Chicago Bears Weekly Report – Week {int(report_week)}", ln=True)
        pdf.set_font("Arial", "", 12)

        if not pred_row.empty:
            outcome = str(pred_row.iloc[0].get("Prediction", ""))
            reason  = str(pred_row.iloc[0].get("Reason", ""))
            pdf.multi_cell(0, 10, f"Prediction: {outcome}\nReason: {reason}\n")

        if not strat_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Strategy Notes:", ln=True)
            pdf.set_font("Arial", "", 12)
            strategy_text = strat_row.iloc[0].astype(str).str.cat(sep=" | ")
            pdf.multi_cell(0, 10, strategy_text)

        if not off_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Offensive Analytics:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in off_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        if not def_row.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Defensive Analytics:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in def_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        if not media_rows.empty:
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Media Summaries:", ln=True)
            pdf.set_font("Arial", "", 12)
            for _, row in media_rows.iterrows():
                source = row.get("Opponent", "Source")
                summary = row.get("Summary", "")
                pdf.multi_cell(0, 10, f"{source}:\n{summary}\n")

        pdf_output = f"week_{int(report_week)}_report.pdf"
        pdf.output(pdf_output)
        with open(pdf_output, "rb") as f:
            st.download_button(
                label=f"📥 Download Week {int(report_week)} Report (PDF)",
                data=f,
                file_name=pdf_output,
                mime="application/pdf"
            )
    except Exception as e:
        st.error(f"❌ Failed to generate PDF. Error: {e}")
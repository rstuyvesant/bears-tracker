# bears_dashboard.pyimport streamlit as st
import pandas as pd
import os
from fpdf import FPDF

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# =========================
# Excel helper
# =========================
def append_to_excel(new_data, sheet_name, file_name=EXCEL_FILE, deduplicate=True):
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    try:
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                sheet = book[sheet_name]
                existing_data = pd.DataFrame(sheet.values)
                if not existing_data.empty:
                    existing_data.columns = existing_data.iloc[0]
                    existing_data = existing_data[1:]
                else:
                    existing_data = pd.DataFrame(columns=new_data.columns)
                if deduplicate and "Week" in existing_data.columns and "Week" in new_data.columns:
                    new_week = str(new_data.iloc[0]["Week"])
                    existing_data = existing_data[existing_data["Week"].astype(str) != new_week]
                combined_data = pd.concat([existing_data, new_data], ignore_index=True)
            else:
                combined_data = new_data
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            combined_data = new_data

        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        for r in dataframe_to_rows(combined_data, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

# =========================
# Styling helpers (green/red)
# =========================
def _cell_color(val, *, good_when="high", low=None, high=None):
    try:
        x = float(val)
    except Exception:
        return ""
    green = "background-color: #c7f5c4"  # light green
    red   = "background-color: #f7c4c4"  # light red

    if good_when == "high":
        if high is not None and x >= high:
            return green
        if low is not None and x <= low:
            return red
    else:  # "low"
        if low is not None and x <= low:
            return green
        if high is not None and x >= high:
            return red
    return ""

def style_df_by_rules(df, rules):
    if df is None or len(df) == 0:
        return df
    styler = df.style
    for col, spec in rules.items():
        if col in df.columns:
            styler = styler.applymap(lambda v: _cell_color(v, **spec), subset=[col])
    return styler

# =========================
# Sidebar: Upload CSVs
# =========================
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense   = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy  = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")

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

# =========================
# Optional: auto-fetch blocks (keep if you already use nfl_data_py/pbp)
# =========================
with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pulls 2025 weekly team stats for CHI and saves to Excel.")
    fetch_week = st.number_input("Week to fetch (2025 season)", min_value=1, max_value=25, value=1, step=1, key="fetch_week_2025")
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
                st.warning("No weekly team row found for CHI in that week. Try later or verify the week.")
            else:
                team_week["Week"] = wk

                # Build Offense row
                pass_yards = team_week.get("passing_yards")
                pass_yards = pass_yards.values[0] if pass_yards is not None else None

                pass_att = None
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

                # Defense row (limited without PBP)
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].values[0]
                        break

                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val,
                    "RZ% Allowed": None
                }])

                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)

                st.success(f"✅ Added CHI week {wk} to Offense/Defense (available fields).")
                st.caption("Note: RZ% Allowed & Pressures need play-by-play; use the PBP panel below.")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

# ---- PBP-derived defensive metrics ----
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
            dmins = (
                pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            def play_success(row):
                if pd.isna(row.get("down")) or pd.isna(row.get("ydstogo")) or pd.isna(row.get("yards_gained")):
                    return False
                d = int(row["down"])
                togo = float(row["ydstogo"])
                gain = float(row["yards_gained"])
                if d == 1:
                    return gain >= 0.4 * togo
                elif d == 2:
                    return gain >= 0.6 * togo
                else:
                    return gain >= togo

            plays_mask = (~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))
            pbp_real = pbp_w[plays_mask].copy()
            success_rate = pbp_real.apply(play_success, axis=1).mean() * 100 if len(pbp_real) else 0.0

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
                f"✅ Week {int(pbp_week)} PBP metrics saved — RZ% Allowed: {rz_allowed:.1f} | "
                f"Success Rate% (Off): {success_rate:.1f} | Pressures: {pressures}"
            )
            st.caption("Note: Hurries aren’t separately flagged; pressures = sacks + QB hits.")
    except Exception as e:
        st.error(f"❌ Failed to fetch metrics: {e}")

# ---- DVOA-like Proxy (opponent-adjusted) ----
st.sidebar.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.sidebar.number_input("Week to Compute (2025 Season)", min_value=1, max_value=25, value=1, step=1, key="proxy_week_2025")

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
            if not bears_off.empty:
                opps.update(bears_off["defteam"].unique().tolist())
            if not bears_def.empty:
                opps.update(bears_def["posteam"].unique().tolist())
            opponent = list(opps)[0] if opps else "UNK"

            prior = plays[plays["week"] < wk].copy()

            opp_def_plays = prior[prior["defteam"] == opponent].copy()
            opp_def_epa = opp_def_plays["epa"].mean() if len(opp_def_plays) else None
            opp_def_success = (opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
                               if len(opp_def_plays) else None)

            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            opp_off_success = (opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
                               if len(opp_off_plays) else None)

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
# Download all Excel data
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
# Styled previews of uploaded/loaded data
# =========================
# Try loading from workbook if not just uploaded (so page refresh still shows tables)
def _maybe_load(sheet):
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE, sheet_name=sheet)
        except Exception:
            return None
    return None

if "df_offense" not in locals():
    df_offense = _maybe_load("Offense")
if "df_defense" not in locals():
    df_defense = _maybe_load("Defense")
if "df_strategy" not in locals():
    df_strategy = _maybe_load("Strategy")
if "df_personnel" not in locals():
    df_personnel = _maybe_load("Personnel")

if df_offense is not None and not df_offense.empty:
    st.subheader("Offensive Analytics")
    offense_rules = {
        "YPA":  {"good_when": "high", "low": 6.0,  "high": 7.5},
        "CMP%": {"good_when": "high", "low": 58.0, "high": 66.0},
        "YPC":  {"good_when": "high", "low": 3.8,  "high": 4.5},
        "QBR":  {"good_when": "high", "low": 40.0, "high": 60.0},
        "SR%":  {"good_when": "high", "low": 40.0, "high": 47.0},
    }
    st.dataframe(style_df_by_rules(df_offense, offense_rules), use_container_width=True)

if df_defense is not None and not df_defense.empty:
    st.subheader("Defensive Analytics")
    defense_rules = {
        "SACK":        {"good_when": "high", "low": None, "high": 3.0},
        "Pressures":   {"good_when": "high", "low": None, "high": 8.0},
        "3D% Allowed": {"good_when": "low",  "low": 35.0, "high": 45.0},
        "RZ% Allowed": {"good_when": "low",  "low": 50.0, "high": 65.0},
    }
    st.dataframe(style_df_by_rules(df_defense, defense_rules), use_container_width=True)

if df_strategy is not None and not df_strategy.empty:
    st.subheader("Weekly Strategy")
    st.dataframe(df_strategy, use_container_width=True)

if df_personnel is not None and not df_personnel.empty:
    st.subheader("Personnel Usage")
    personnel_rules = {
        "11 Personnel": {"good_when": "high", "low": None, "high": 60.0},
        "12 Personnel": {"good_when": "high", "low": None, "high": 30.0},
    }
    st.dataframe(style_df_by_rules(df_personnel, personnel_rules), use_container_width=True)

# DVOA-like proxy preview (styled)
if os.path.exists(EXCEL_FILE):
    try:
        df_dvoa = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        if not df_dvoa.empty:
            st.subheader("📊 DVOA-like Proxy Metrics")
            dvoa_rules = {
                "Off Adj EPA/play": {"good_when": "high", "low": 0.00, "high": 0.15},
                "Off Adj SR%":      {"good_when": "high", "low": 45.0, "high": 50.0},
                "Def Adj EPA/play": {"good_when": "low",  "low": -0.05, "high": 0.00},
                "Def Adj SR%":      {"good_when": "low",  "low": 45.0,  "high": 55.0},
            }
            st.dataframe(style_df_by_rules(df_dvoa, dvoa_rules), use_container_width=True)
    except Exception:
        pass

# =========================
# Media Summaries
# =========================
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

if os.path.exists(EXCEL_FILE):
    try:
        df_media = pd.read_excel(EXCEL_FILE, sheet_name="Media_Summaries")
        if not df_media.empty:
            st.subheader("📰 Saved Media Summaries")
            st.dataframe(df_media, use_container_width=True)
    except Exception:
        st.info("No media summaries found.")

# =========================
# Weekly Game Prediction
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
        df_strategy = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        df_offense  = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        df_defense  = pd.read_excel(EXCEL_FILE, sheet_name="Defense")
        try:
            df_advdef = pd.read_excel(EXCEL_FILE, sheet_name="Advanced_Defense")
        except Exception:
            df_advdef = pd.DataFrame()
        try:
            df_proxy = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        except Exception:
            df_proxy = pd.DataFrame()

        row_s = df_strategy[df_strategy["Week"] == week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"] == week_to_predict] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"] == week_to_predict] if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"] == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"] == week_to_predict] if not df_proxy.empty else pd.DataFrame()

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
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win – efficiency edge on both sides"
                reason_bits.append(f"Off+{off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play vs opp O")
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win – red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient offense and poor red zone defense"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
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

# Saved predictions (plain table)
if os.path.exists(EXCEL_FILE):
    try:
        df_preds = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
        if not df_preds.empty:
            st.subheader("📈 Saved Game Predictions")
            st.dataframe(df_preds, use_container_width=True)
        else:
            st.info("No predictions saved yet.")
    except Exception:
        st.info("No predictions saved yet.")
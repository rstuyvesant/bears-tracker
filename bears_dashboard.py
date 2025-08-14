import streamlit as st
import pandas as pd
import os
from fpdf import FPDF

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# --------------------------
# Helpers
# --------------------------
def _coerce_week(val):
    try:
        return int(val)
    except Exception:
        try:
            return int(float(val))
        except Exception:
            return None

def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

def _safe_int(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return int(x)
    except Exception:
        return default

# --------------------------
# Robust Excel append with dedup by 'Week'
# --------------------------
def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True):
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    try:
        # Normalize Week column to int if present
        if "Week" in new_data.columns:
            new_data["Week"] = new_data["Week"].apply(_coerce_week)

        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                sheet = book[sheet_name]
                existing_data = pd.DataFrame(sheet.values)
                if len(existing_data) > 0:
                    existing_data.columns = existing_data.iloc[0]
                    existing_data = existing_data[1:]
                else:
                    existing_data = pd.DataFrame(columns=new_data.columns)
                # align columns
                for col in new_data.columns:
                    if col not in existing_data.columns:
                        existing_data[col] = pd.NA
                for col in existing_data.columns:
                    if col not in new_data.columns:
                        new_data[col] = pd.NA
                existing_data = existing_data[new_data.columns]

                # Dedup by Week if requested and Week exists
                if deduplicate and "Week" in existing_data.columns and "Week" in new_data.columns:
                    keep_week = _coerce_week(new_data.iloc[0]["Week"])
                    if keep_week is not None:
                        existing_data = existing_data[existing_data["Week"].apply(_coerce_week) != keep_week]

                combined_data = pd.concat([existing_data, new_data], ignore_index=True)
            else:
                combined_data = new_data
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            combined_data = new_data

        # Replace sheet and write combined
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        # If Week exists, sort by Week
        if "Week" in combined_data.columns:
            combined_data["Week"] = combined_data["Week"].apply(_coerce_week)
            combined_data = combined_data.sort_values(by=["Week"], kind="stable")

        for r in dataframe_to_rows(combined_data, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

# --------------------------
# Sidebar: Upload CSVs
# --------------------------
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Offensive Analytics (.csv)", type="csv")
uploaded_defense   = st.sidebar.file_uploader("Defensive Analytics (.csv)", type="csv")
uploaded_strategy  = st.sidebar.file_uploader("Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Personnel Usage (.csv)", type="csv")
uploaded_injuries  = st.sidebar.file_uploader("Injuries (.csv)", type="csv")

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

if uploaded_injuries:
    df_inj = pd.read_csv(uploaded_injuries)
    append_to_excel(df_inj, "Injuries", deduplicate=False)  # allow multiple rows per week/player
    st.sidebar.success("✅ Injuries uploaded.")

# --------------------------
# Sidebar: Fetch blocks (restored)
# --------------------------
with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pulls 2025 weekly team stats for CHI and saves to Excel (Offense/Defense basics).")
    fetch_week = st.number_input("Week to fetch (2025)", min_value=1, max_value=25, value=1, step=1, key="fetch_week_2025")

    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl

            # (Optional) refresh local cache; ignore errors
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
                st.warning("No weekly team row found for CHI in that week. Try again later or verify the week.")
            else:
                team_week["Week"] = wk

                # ---- Offense (best-effort mapping) ----
                pass_yards = team_week["passing_yards"].values[0] if "passing_yards" in team_week.columns else None

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

                # ---- Defense (basic) ----
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].values[0]
                        break

                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val,
                    "RZ% Allowed": None  # PBP fetch will fill this
                }])

                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)

                st.success(f"✅ Added CHI week {wk} to Offense/Defense (available fields).")
                st.caption("Tip: Next, click 'Fetch Defensive Metrics (Play-by-Play)' for RZ% Allowed, Success Rate, and Pressures.")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📡 Fetch Defensive Metrics (Play-by-Play)")
pbp_week = st.sidebar.number_input("Week to fetch (2025)", min_value=1, max_value=25, value=1, step=1, key="pbp_week_2025")

if st.sidebar.button("Fetch Play-by-Play Metrics"):
    try:
        import nfl_data_py as nfl

        pbp = nfl.import_pbp_data([2025], downcast=False)
        pbp_w = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == "CHI")].copy()

        if pbp_w.empty:
            st.warning("No PBP rows for CHI defense in that week yet. Try again later.")
        else:
            # Red Zone % Allowed (drives reaching <=20)
            dmins = (
                pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success Rate against CHI defense (offense success)
            def play_success(row):
                try:
                    d = int(row["down"])
                    togo = float(row["ydstogo"])
                    gain = float(row["yards_gained"])
                except Exception:
                    return False
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
                f"✅ Week {int(pbp_week)}: RZ% Allowed {rz_allowed:.1f} | "
                f"Success Rate% (Off) {success_rate:.1f} | Pressures {pressures}"
            )
    except Exception as e:
        st.error(f"❌ Failed to fetch metrics: {e}")

# --------------------------
# DVOA-like Proxy (Opponent-adjusted)
# --------------------------
st.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.number_input("Week to compute", min_value=1, max_value=25, value=1, step=1, key="proxy_week_2025")

def _success_flag(down, ydstogo, yards_gained):
    try:
        if down is None or pd.isna(down) or ydstogo is None or pd.isna(ydstogo) or yards_gained is None or pd.isna(yards_gained):
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

        plays = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False)) &
            (~pbp["epa"].isna())
        ].copy()

        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == "CHI")].copy()
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == "CHI")].copy()

        if bears_off.empty and bears_def.empty:
            st.warning("No CHI plays found for that week yet. Try again later.")
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
            st.success(
                f"✅ DVOA-like proxy saved for Week {wk} vs {opponent} "
                f"(Off Adj EPA/play={out.iloc[0]['Off Adj EPA/play']}, Off Adj SR%={out.iloc[0]['Off Adj SR%']}, "
                f"Def Adj EPA/play={out.iloc[0]['Def Adj EPA/play']}, Def Adj SR%={out.iloc[0]['Def Adj SR%']})"
            )
    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# --------------------------
# Download Excel (sidebar)
# --------------------------
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --------------------------
# On-page previews of uploaded/fetched data
# --------------------------
def _df_preview(title, df):
    st.subheader(title)
    if df is not None and not df.empty:
        st.dataframe(df)
    else:
        st.info("No data yet.")

# Try to load common sheets for preview
_df_off, _df_def, _df_strat, _df_pers, _df_inj, _df_proxy = None, None, None, None, None, None
if os.path.exists(EXCEL_FILE):
    try:
        _df_off = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
    except Exception:
        _df_off = pd.DataFrame()
    try:
        _df_def = pd.read_excel(EXCEL_FILE, sheet_name="Defense")
    except Exception:
        _df_def = pd.DataFrame()
    try:
        _df_strat = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
    except Exception:
        _df_strat = pd.DataFrame()
    try:
        _df_pers = pd.read_excel(EXCEL_FILE, sheet_name="Personnel")
    except Exception:
        _df_pers = pd.DataFrame()
    try:
        _df_inj = pd.read_excel(EXCEL_FILE, sheet_name="Injuries")
    except Exception:
        _df_inj = pd.DataFrame()
    try:
        _df_proxy = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
    except Exception:
        _df_proxy = pd.DataFrame()

if _df_off is not None:
    _df_preview("📊 Offensive Analytics", _df_off)
if _df_def is not None:
    _df_preview("🛡️ Defensive Analytics", _df_def)
if _df_strat is not None:
    _df_preview("📘 Weekly Strategy", _df_strat)
if _df_pers is not None:
    _df_preview("👥 Personnel Usage", _df_pers)
if _df_proxy is not None:
    _df_preview("📈 DVOA-Like Proxy (latest rows)", _df_proxy.tail(5) if _df_proxy is not None else pd.DataFrame())
if _df_inj is not None:
    _df_preview("🩹 Injuries", _df_inj)

# --------------------------
# Injuries – add via form (optional)
# --------------------------
st.markdown("### 🩹 Add Injury Note (optional)")
with st.form("injury_form"):
    inj_week = st.number_input("Week", min_value=1, max_value=25, step=1)
    inj_player = st.text_input("Player")
    inj_pos = st.text_input("Position")
    inj_type = st.text_input("Injury Type (e.g., Hamstring)")
    inj_status = st.selectbox("Game Status", ["Questionable", "Doubtful", "Out", "IR", "Active"])
    inj_prac = st.text_input("Practice Participation (e.g., DNP/Limited/Full)")
    inj_notes = st.text_area("Notes")
    inj_submit = st.form_submit_button("Save Injury Note")

if inj_submit:
    inj_df = pd.DataFrame([{
        "Week": inj_week,
        "Player": inj_player,
        "Position": inj_pos,
        "Injury": inj_type,
        "Status": inj_status,
        "Practice": inj_prac,
        "Notes": inj_notes
    }])
    append_to_excel(inj_df, "Injuries", deduplicate=False)
    st.success(f"✅ Injury note saved for Week {inj_week} – {inj_player}")

# --------------------------
# Media summaries
# --------------------------
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
        st.subheader("📰 Saved Media Summaries")
        st.dataframe(df_media)
    except Exception:
        st.info("No media summaries stored yet.")

# --------------------------
# Weekly Prediction
# --------------------------
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

if os.path.exists(EXCEL_FILE):
    try:
        # Base sheets
        df_strategy = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        df_offense  = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        df_defense  = pd.read_excel(EXCEL_FILE, sheet_name="Defense")

        # Optional sheets
        try:
            df_advdef = pd.read_excel(EXCEL_FILE, sheet_name="Advanced_Defense")
        except Exception:
            df_advdef = pd.DataFrame()

        try:
            df_proxy = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        except Exception:
            df_proxy = pd.DataFrame()

        # Filter selected week
        row_s = df_strategy[df_strategy["Week"].apply(_coerce_week) == week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"].apply(_coerce_week) == week_to_predict] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"].apply(_coerce_week) == week_to_predict] if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"].apply(_coerce_week) == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"].apply(_coerce_week) == week_to_predict] if not df_proxy.empty else pd.DataFrame()

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
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play vs opp D")
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

# --------------------------
# Show saved predictions
# --------------------------
if os.path.exists(EXCEL_FILE):
    try:
        df_preds = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
        st.subheader("📈 Saved Game Predictions")
        st.dataframe(df_preds)
    except Exception:
        st.info("No predictions saved yet.")
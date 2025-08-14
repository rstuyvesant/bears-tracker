# bears_dashboard.py

import streamlit as st
import pandas as pd
import os
from fpdf import FPDF

# ------------------ App Setup ------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"
SEASON = 2025  # change when needed

# ------------------ Helpers ------------------
def _read_sheet(sheet_name: str) -> pd.DataFrame:
    """Read a sheet; return empty DataFrame if missing or file absent."""
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        return df
    except Exception:
        return pd.DataFrame()

def _ensure_workbook_sheets():
    """Create workbook and base sheets if missing."""
    import openpyxl
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        wb.create_sheet("Offense")
        wb.create_sheet("Defense")
        wb.create_sheet("Strategy")
        wb.create_sheet("Personnel")
        wb.create_sheet("Advanced_Defense")
        wb.create_sheet("DVOA_Proxy")
        wb.create_sheet("Predictions")
        wb.create_sheet("Media_Summaries")
        wb.save(EXCEL_FILE)

def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True):
    """Append or replace rows into sheet; if 'Week' present, replace existing week rows."""
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    # Make sure workbook exists
    _ensure_workbook_sheets()

    try:
        book = openpyxl.load_workbook(file_name)
        # Read existing
        if sheet_name in book.sheetnames:
            sheet = book[sheet_name]
            existing_raw = list(sheet.values)
            if existing_raw:
                header = existing_raw[0]
                rows = existing_raw[1:]
                existing = pd.DataFrame(rows, columns=header)
            else:
                existing = pd.DataFrame(columns=new_data.columns)
        else:
            existing = pd.DataFrame(columns=new_data.columns)

        # Align columns
        all_cols = list(dict.fromkeys(list(existing.columns) + list(new_data.columns)))
        existing = existing.reindex(columns=all_cols)
        new_data = new_data.reindex(columns=all_cols)

        # De-dup by Week if requested and column exists
        if deduplicate and "Week" in existing.columns and "Week" in new_data.columns:
            incoming_weeks = set(pd.to_numeric(new_data["Week"], errors="coerce").dropna().astype(int).tolist())
            if not existing.empty:
                existing["Week_num"] = pd.to_numeric(existing["Week"], errors="coerce")
                existing = existing[~existing["Week_num"].isin(incoming_weeks)].drop(columns=["Week_num"], errors="ignore")

        combined = pd.concat([existing, new_data], ignore_index=True)

        # Rewrite sheet
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)
        for r in dataframe_to_rows(combined, index=False, header=True):
            sheet.append(r)
        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

# ------------------ Sidebar Uploaders ------------------
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")

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

# ------------------ Auto-Fetch Missing Weeks (nfl_data_py) ------------------
with st.sidebar.expander("⚡ Auto-Fetch Missing Weeks (nfl_data_py)"):
    st.caption("Pulls team-level weekly stats for CHI and fills any missing Offense/Defense rows.")
    max_week = st.number_input("Fetch up to week", min_value=1, max_value=25, value=1, step=1, key="fetch_to_week")

    if st.button("Fetch Missing Weeks"):
        try:
            import nfl_data_py as nfl

            # Try to refresh local cache (safe if it fails)
            try:
                nfl.update.schedule_data([SEASON])
            except Exception:
                pass
            try:
                nfl.update.weekly_data([SEASON])
            except Exception:
                pass

            weekly = nfl.import_weekly_data([SEASON])  # team-level weekly stats
            weekly = weekly[(weekly["team"] == "CHI") & (weekly["week"] <= int(max_week))].copy()

            # Current weeks present
            off_curr = _read_sheet("Offense")
            def_curr = _read_sheet("Defense")
            have_off = set(pd.to_numeric(off_curr.get("Week", pd.Series([])), errors="coerce").dropna().astype(int).tolist())
            have_def = set(pd.to_numeric(def_curr.get("Week", pd.Series([])), errors="coerce").dropna().astype(int).tolist())

            added = 0
            for wk in sorted(weekly["week"].unique().tolist()):
                tw = weekly[weekly["week"] == wk].copy()
                if tw.empty:
                    continue

                # ----- OFFENSE -----
                off_row = None
                if wk not in have_off:
                    # attempt to compute YPA, CMP%, YDS
                    pass_yards = tw.get("passing_yards")
                    pass_yards = pass_yards.values[0] if pass_yards is not None and len(pass_yards) else None

                    pass_att = None
                    for cand in ["attempts", "passing_attempts", "pass_attempts"]:
                        if cand in tw.columns:
                            pass_att = tw[cand].values[0]
                            break

                    try:
                        ypa_val = float(pass_yards) / float(pass_att) if pass_yards not in (None, "") and pass_att not in (None, 0, "") else None
                    except Exception:
                        ypa_val = None

                    yards_total = None
                    for cand in ["yards", "total_yards", "offense_yards"]:
                        if cand in tw.columns:
                            yards_total = tw[cand].values[0]
                            break

                    completions = None
                    for cand in ["completions", "passing_completions", "pass_completions"]:
                        if cand in tw.columns:
                            completions = tw[cand].values[0]
                            break

                    cmp_pct = None
                    if completions not in (None, "") and pass_att not in (None, 0, ""):
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

                # ----- DEFENSE -----
                def_row = None
                if wk not in have_def:
                    sacks_val = None
                    for cand in ["sacks", "defense_sacks"]:
                        if cand in tw.columns:
                            sacks_val = tw[cand].values[0]
                            break

                    def_row = pd.DataFrame([{
                        "Week": wk,
                        "SACK": sacks_val,
                        "RZ% Allowed": None  # PBP-based later
                    }])

                # Save if new
                if off_row is not None:
                    append_to_excel(off_row, "Offense", deduplicate=True)
                    added += 1
                if def_row is not None:
                    append_to_excel(def_row, "Defense", deduplicate=True)
                    added += 1

            st.success(f"✅ Auto-fetch complete. New rows added: {added}")
            st.caption("Note: RZ% Allowed, Success Rate, and Pressures come from play-by-play (panel below).")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

# ------------------ PBP Metrics (RZ% Allowed, Success Rate, Pressures) ------------------
st.sidebar.markdown("### 📡 Fetch Defensive Metrics from PBP")
pbp_week = st.sidebar.number_input("Week to fetch (PBP)", min_value=1, max_value=25, value=1, step=1, key="pbp_week")

if st.sidebar.button("Fetch PBP Metrics"):
    try:
        import nfl_data_py as nfl

        pbp = nfl.import_pbp_data([SEASON], downcast=False)
        pbp_w = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == "CHI")].copy()

        if pbp_w.empty:
            st.warning("No PBP rows for CHI defense in that week yet. Try again later.")
        else:
            # Red Zone % Allowed via min yardline per drive
            dmins = (
                pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Offense success rate against CHI (CHI on defense)
            def _success(row):
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
            success_rate = pbp_real.apply(_success, axis=1).mean() * 100 if len(pbp_real) else 0.0

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
            st.caption("Note: Hurries aren’t separately flagged in standard PBP; pressures = sacks + QB hits.")
    except Exception as e:
        st.error(f"❌ Failed to fetch PBP metrics: {e}")

# ------------------ DVOA-like Proxy (Opponent-Adjusted) ------------------
st.sidebar.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.sidebar.number_input("Week to compute (Proxy)", min_value=1, max_value=25, value=1, step=1, key="proxy_week")

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
        pbp = nfl.import_pbp_data([SEASON], downcast=False)

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
            if len(opp_def_plays):
                opp_def_success = opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                opp_def_success = None

            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            if len(opp_off_plays):
                opp_off_success = opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                opp_off_success = None

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

# ------------------ Download Excel ------------------
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------ Quick Previews ------------------
off_prev = _read_sheet("Offense")
def_prev = _read_sheet("Defense")
str_prev = _read_sheet("Strategy")
per_prev = _read_sheet("Personnel")
adv_prev = _read_sheet("Advanced_Defense")
proxy_prev = _read_sheet("DVOA_Proxy")

if not off_prev.empty:
    st.subheader("Offensive Analytics")
    st.dataframe(off_prev)
if not def_prev.empty:
    st.subheader("Defensive Analytics")
    st.dataframe(def_prev)
if not str_prev.empty:
    st.subheader("Weekly Strategy")
    st.dataframe(str_prev)
if not per_prev.empty:
    st.subheader("Personnel Usage")
    st.dataframe(per_prev)
if not proxy_prev.empty:
    st.subheader("📊 DVOA-like Proxy (last 5)")
    st.dataframe(proxy_prev.tail(5))
if not adv_prev.empty:
    st.subheader("Advanced Defensive Metrics (PBP)")
    st.dataframe(adv_prev)

# ------------------ Media Summaries ------------------
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

# ------------------ Weekly Prediction ------------------
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
        df_strategy = _read_sheet("Strategy")
        df_offense  = _read_sheet("Offense")
        df_defense  = _read_sheet("Defense")
        df_advdef   = _read_sheet("Advanced_Defense")
        df_proxy    = _read_sheet("DVOA_Proxy")

        row_s = df_strategy[df_strategy.get("Week", pd.Series([])) == week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense.get("Week", pd.Series([])) == week_to_predict] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense.get("Week", pd.Series([])) == week_to_predict] if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef.get("Week", pd.Series([])) == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy.get("Week", pd.Series([])) == week_to_predict] if not df_proxy.empty else pd.DataFrame()

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
            # Efficiency edge
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win - efficiency edge on both sides"
                reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def {def_adj_epa:+.2f} EPA/play vs opp O")
            # Pass-rush advantage
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win - pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            # Coverage + red zone discipline
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win - red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            # Clear drag
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient offense and poor red zone defense"
                reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            # Legacy fallback
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient passing and weak red zone defense"
                reason_bits.append(f"YPA={ypa:.1f}")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            else:
                prediction = "Loss - no clear advantage in key strategy or stats"
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

            # Save
            prediction_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("-")[0].strip(),
                "Reason": prediction.split("-")[1].strip() if "-" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(prediction_entry, "Predictions", deduplicate=True)
        else:
            st.info("Please upload or fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# ------------------ Saved Predictions Preview ------------------
pred_prev = _read_sheet("Predictions")
if not pred_prev.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(pred_prev)

# ------------------ Optional Workbook Sanity Checker ------------------
with st.expander("🧪 Excel Sanity Checker (optional)"):
    cwd = os.getcwd()
    st.write("Current Working Directory:", cwd)
    st.write("Excel File Being Used:", EXCEL_FILE)
    exists = os.path.exists(EXCEL_FILE)
    st.write("Exists:", exists)
    if exists:
        st.write("Size (bytes):", os.path.getsize(EXCEL_FILE))
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_FILE)
            st.write("Sheets:", wb.sheetnames)
            previews = ["Offense", "Defense", "Strategy", "Personnel", "Advanced_Defense", "DVOA_Proxy", "Predictions"]
            for sn in previews:
                dfp = _read_sheet(sn)
                st.write(f"Preview - {sn}", dfp.head(5) if not dfp.empty else "(empty)")
        except Exception as e:
            st.error(f"Workbook read error: {e}")
import streamlit as st
import pandas as pd
import os
from fpdf import FPDF

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# =========================
# Helpers
# =========================
def append_to_excel(new_data, sheet_name, file_name=EXCEL_FILE, deduplicate=True, dedup_cols=("Week",)):
    """
    Append/replace a sheet in the Excel workbook.
    If deduplicate=True and dedup_cols exist, removes existing rows that match new_data's key(s).
    """
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
                # Normalize columns
                if set(new_data.columns) != set(existing_data.columns):
                    existing_data = existing_data.reindex(columns=new_data.columns)
                # Dedup logic
                if deduplicate:
                    if all(c in existing_data.columns and c in new_data.columns for c in dedup_cols):
                        keys_to_drop = set(
                            tuple(map(str, r)) for r in new_data[dedup_cols].astype(str).itertuples(index=False, name=None)
                        )
                        mask = existing_data[dedup_cols].astype(str).apply(tuple, axis=1).isin(keys_to_drop)
                        existing_data = existing_data[~mask]
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


def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default


# =========================
# Sidebar: Uploads
# =========================
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


# =========================
# Download Excel (sidebar)
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
# Previews (main)
# =========================
if uploaded_offense:
    st.subheader("Offensive Analytics")
    st.dataframe(df_offense, use_container_width=True)

if uploaded_defense:
    st.subheader("Defensive Analytics")
    st.dataframe(df_defense, use_container_width=True)

if uploaded_strategy:
    st.subheader("Weekly Strategy")
    st.dataframe(df_strategy, use_container_width=True)

if uploaded_personnel:
    st.subheader("Personnel Usage")
    st.dataframe(df_personnel, use_container_width=True)


# =========================
# 🩹 Injuries (NEW)
# =========================
st.markdown("### 🩹 Injuries")

# Upload injuries CSV (must include Week)
uploaded_injuries = st.file_uploader("Upload Injuries (.csv)", type="csv", key="inj_csv")
if uploaded_injuries is not None:
    try:
        df_inj = pd.read_csv(uploaded_injuries)
        if "Week" not in df_inj.columns:
            st.error("Your injuries CSV must include a 'Week' column.")
        else:
            append_to_excel(df_inj, "Injuries", deduplicate=False)
            st.success("✅ Injuries uploaded.")
    except Exception as e:
        st.error(f"Couldn't read injuries CSV: {e}")

# Quick add one injury line
with st.form("inj_quick_add"):
    st.caption("Quick add a single injury")
    inj_week = st.number_input("Week", 1, 25, 1, step=1, key="inj_week_add")
    inj_oppo = st.text_input("Opponent (optional)", key="inj_oppo_add")
    inj_player = st.text_input("Player", key="inj_player_add")
    inj_pos = st.text_input("Position (e.g., QB, LT, WR)", key="inj_pos_add")
    inj_status = st.selectbox("Game Status", ["Questionable", "Doubtful", "Out", "IR", "Active"], key="inj_status_add")
    inj_prac = st.selectbox("Practice (Fri)", ["Full", "Limited", "DNP", "N/A"], key="inj_prac_add")
    inj_notes = st.text_area("Notes", key="inj_notes_add")
    inj_submit = st.form_submit_button("Save Injury")

if inj_submit:
    if inj_player.strip() == "":
        st.warning("Please enter a player name.")
    else:
        inj_row = pd.DataFrame([{
            "Week": inj_week,
            "Opponent": inj_oppo,
            "Player": inj_player,
            "Position": inj_pos,
            "Status": inj_status,
            "Practice_Fri": inj_prac,
            "Notes": inj_notes
        }])
        append_to_excel(inj_row, "Injuries", deduplicate=False)
        st.success(f"✅ Saved injury for Week {inj_week}: {inj_player} ({inj_status})")

# Preview injuries by week
inj_filter_week = st.number_input("Preview injuries for week", 1, 25, 1, step=1, key="inj_prev_wk")
if os.path.exists(EXCEL_FILE):
    try:
        df_injuries_all = pd.read_excel(EXCEL_FILE, sheet_name="Injuries")
        if not df_injuries_all.empty:
            view = df_injuries_all[df_injuries_all["Week"] == inj_filter_week]
            if view.empty:
                st.info("No injuries recorded for that week yet.")
            else:
                st.dataframe(view, use_container_width=True)
        else:
            st.info("No injuries saved yet.")
    except Exception:
        st.info("No injuries saved yet.")


# =========================
# 📰 Weekly Beat Writer / ESPN Summary
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
        st.subheader("📰 Saved Media Summaries")
        st.dataframe(df_media, use_container_width=True)
    except Exception:
        st.info("No media summaries found.")


# =========================
# 📊 Compute DVOA-like Proxy (simple)
# =========================
st.markdown("### 📊 Compute DVOA-like Proxy")
week_to_compute = st.number_input("Week to compute", min_value=1, max_value=25, step=1, key="dvoa_week_input")

if st.button("🔴 Compute DVOA-like Proxy"):
    if os.path.exists(EXCEL_FILE):
        try:
            df_off = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
            df_def = pd.read_excel(EXCEL_FILE, sheet_name="Defense")

            row_o = df_off[df_off["Week"] == week_to_compute]
            row_d = df_def[df_def["Week"] == week_to_compute]

            st.write("Offense Row:", row_o)
            st.write("Defense Row:", row_d)

            if not row_o.empty and not row_d.empty:
                try:
                    idx_o = row_o.index[0]
                    idx_d = row_d.index[0]
                    ypa = float(row_o.at[idx_o, "YPA"]) if "YPA" in row_o.columns else 0
                    cmp_pct = float(row_o.at[idx_o, "CMP%"]) if "CMP%" in row_o.columns else 0
                    rz_def = float(row_d.at[idx_d, "RZ% Allowed"]) if "RZ% Allowed" in row_d.columns else 0
                    sacks = int(row_d.at[idx_d, "SACK"]) if "SACK" in row_d.columns else 0
                except Exception as e:
                    st.error(f"❌ Data extraction error: {e}")
                    ypa = cmp_pct = rz_def = sacks = 0

                proxy = round((ypa * 0.4) + (cmp_pct * 0.3) - (rz_def * 0.2) + (sacks * 0.5), 2)

                dvoa_row = pd.DataFrame([{
                    "Week": week_to_compute,
                    "DVOA_Proxy": proxy,
                    "YPA": ypa,
                    "CMP%": cmp_pct,
                    "RZ% Allowed": rz_def,
                    "SACK": sacks
                }])

                append_to_excel(dvoa_row, "DVOA_Proxy", deduplicate=True)
                st.success(f"✅ DVOA Proxy for Week {week_to_compute}: {proxy}")
            else:
                st.warning("Missing offense or defense data for that week.")
        except Exception as e:
            st.error(f"Error computing DVOA Proxy: {e}")
    else:
        st.warning("Excel file not found.")

# Preview saved DVOA Proxy data
if os.path.exists(EXCEL_FILE):
    try:
        df_dvoa = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        if not df_dvoa.empty:
            st.subheader("📊 DVOA Proxy Metrics")
            st.dataframe(df_dvoa, use_container_width=True)
        else:
            st.info("No DVOA Proxy data available yet.")
    except Exception:
        st.info("No DVOA Proxy data available yet.")


# =========================
# 🔮 Weekly Game Prediction (with Injuries impact)
# =========================
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

if os.path.exists(EXCEL_FILE):
    try:
        # Required base sheets
        df_strategy = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        df_offense  = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        df_defense  = pd.read_excel(EXCEL_FILE, sheet_name="Defense")

        # Optional advanced sheets
        try:
            df_advdef = pd.read_excel(EXCEL_FILE, sheet_name="Advanced_Defense")
        except Exception:
            df_advdef = pd.DataFrame()
        try:
            df_proxy = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        except Exception:
            df_proxy = pd.DataFrame()
        try:
            df_injuries = pd.read_excel(EXCEL_FILE, sheet_name="Injuries")
        except Exception:
            df_injuries = pd.DataFrame()

        # Filter to selected week
        row_s = df_strategy[df_strategy["Week"] == week_to_predict]
        row_o = df_offense[df_offense["Week"] == week_to_predict]
        row_d = df_defense[df_defense["Week"] == week_to_predict]
        row_a = df_advdef[df_advdef["Week"] == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"] == week_to_predict] if not df_proxy.empty else pd.DataFrame()
        row_i = df_injuries[df_injuries["Week"] == week_to_predict] if not df_injuries.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            # Offense basics
            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            # Prefer Advanced_Defense for RZ% and Pressures
            rz_allowed = None
            pressures = None
            if not row_a.empty:
                rz_allowed = _safe_float(row_a.iloc[0].get("RZ% Allowed"), default=None)
                pressures  = _safe_float(row_a.iloc[0].get("Pressures"), default=None)
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), default=None)

            # DVOA-like proxy
            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), default=None)

            # Injury impact
            injury_penalty = 0.0
            injury_bits = []
            if not row_i.empty:
                key_positions = {"QB": 2.0, "LT": 1.5, "WR": 1.0, "CB": 1.0, "EDGE": 1.0, "RB": 0.5, "TE": 0.5}
                for _, r in row_i.iterrows():
                    status = str(r.get("Status", "")).strip().lower()
                    prac = str(r.get("Practice_Fri", "")).strip().upper()
                    pos = str(r.get("Position", "")).strip().upper()
                    w = key_positions.get(pos, 0.5)
                    if status in ("out", "ir") or prac == "DNP":
                        injury_penalty += w
                        injury_bits.append(f"{r.get('Player','?')}({pos},{status or prac})")

            # Rule set
            reason_bits = []

            # Major injuries tilt to Loss unless strong efficiency edge
            if injury_penalty >= 2.5 and not ((off_adj_epa is not None and off_adj_epa >= 0.25) and (def_adj_epa is not None and def_adj_epa <= -0.10)):
                prediction = "Loss – key injuries"
                if injury_bits:
                    reason_bits.append("Injuries: " + ", ".join(injury_bits))
                if off_adj_epa is not None:
                    reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                if def_adj_epa is not None:
                    reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play")

            elif (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
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

            # Always include injury notes if any (even when small)
            if injury_bits and all("Injuries:" not in rb for rb in reason_bits):
                reason_bits.append("Injuries: " + ", ".join(injury_bits))

            reason_text = " | ".join(reason_bits)
            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")
            if reason_text:
                st.caption(reason_text)

            # Save prediction
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
if os.path.exists(EXCEL_FILE):
    try:
        df_preds = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
        st.subheader("📈 Saved Game Predictions")
        st.dataframe(df_preds, use_container_width=True)
    except Exception:
        st.info("No predictions saved yet.")
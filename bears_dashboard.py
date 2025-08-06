import streamlit as st
import pandas as pd
import os
from fpdf import FPDF
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# Append new data to Excel workbook
def append_to_excel(new_data, sheet_name, file_name=EXCEL_FILE, deduplicate=True):
    try:
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                sheet = book[sheet_name]
                existing_data = pd.DataFrame(sheet.values)
                existing_data.columns = existing_data.iloc[0]
                existing_data = existing_data[1:]

                if deduplicate and "Week" in existing_data.columns and "Week" in new_data.columns:
                    existing_data = existing_data[existing_data["Week"] != str(new_data.iloc[0]["Week"])]
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

# Sidebar upload
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

# ----- DVOA Proxy Computation -----
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
                    cmp = float(row_o.at[idx_o, "CMP%"]) if "CMP%" in row_o.columns else 0
                    rz_def = float(row_d.at[idx_d, "RZ% Allowed"]) if "RZ% Allowed" in row_d.columns else 0
                    sacks = int(row_d.at[idx_d, "SACK"]) if "SACK" in row_d.columns else 0
                except Exception as e:
                    st.error(f"❌ Data extraction error: {e}")
                    ypa = cmp = rz_def = sacks = 0

                proxy = round((ypa * 0.4) + (cmp * 0.3) - (rz_def * 0.2) + (sacks * 0.5), 2)

                dvoa_row = pd.DataFrame([{
                    "Week": week_to_compute,
                    "DVOA_Proxy": proxy,
                    "YPA": ypa,
                    "CMP%": cmp,
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
        st.subheader("📊 DVOA Proxy Metrics")
        st.dataframe(df_dvoa)
    except:
        st.info("No DVOA Proxy data available yet.")

# Download Excel
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Show uploaded data
if uploaded_offense:
    st.subheader("Offensive Analytics")
    st.dataframe(df_offense)

if uploaded_defense:
    st.subheader("Defensive Analytics")
    st.dataframe(df_defense)

if uploaded_strategy:
    st.subheader("Weekly Strategy")
    st.dataframe(df_strategy)

if uploaded_personnel:
    st.subheader("Personnel Usage")
    st.dataframe(df_personnel)

# Media Summary Section
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

# Show media summaries
if os.path.exists(EXCEL_FILE):
    try:
        df_media = pd.read_excel(EXCEL_FILE, sheet_name="Media_Summaries")
        st.subheader("📰 Saved Media Summaries")
        st.dataframe(df_media)
    except:
        st.info("No media summaries found.")

# Prediction Section
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        df_offense = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        df_defense = pd.read_excel(EXCEL_FILE, sheet_name="Defense")

        row_s = df_strategy[df_strategy["Week"] == week_to_predict]
        row_o = df_offense[df_offense["Week"] == week_to_predict]
        row_d = df_defense[df_defense["Week"] == week_to_predict]

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()
            try:
                ypa = float(row_o.iloc[0].get("YPA", 0))
                red_zone_allowed = float(row_d.iloc[0].get("RZ% Allowed", 0))
                sacks = int(row_d.iloc[0].get("SACK", 0))
            except:
                ypa = red_zone_allowed = sacks = 0

            if "blitz" in strategy_text and sacks >= 3:
                prediction = "Win – pressure defense likely disrupts opponent"
            elif ypa < 6 and red_zone_allowed > 65:
                prediction = "Loss – inefficient passing and weak red zone defense"
            elif "zone" in strategy_text and red_zone_allowed < 50:
                prediction = "Win – disciplined zone and red zone efficiency"
            elif any(word in strategy_text for word in ["struggled", "injuries", "turnovers"]):
                prediction = "Loss – opponent issues likely to affect performance"
            else:
                prediction = "Loss – no clear advantage in key strategy or stats"

            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")

            prediction_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else ""
            }])
            append_to_excel(prediction_entry, "Predictions", deduplicate=True)
        else:
            st.info("Missing data for that week.")
    except Exception as e:
        st.warning("Prediction failed. Check uploaded data.")

# Show saved predictions
if os.path.exists(EXCEL_FILE):
    try:
        df_preds = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
        st.subheader("📈 Saved Game Predictions")
        st.dataframe(df_preds)
    except:
        st.info("No predictions saved yet.")
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

            # Clear offensive/defensive drag
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient offense and poor red zone defense"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            # Legacy fallback using YPA + RZ
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

            # Save prediction to Excel
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
        st.dataframe(df_preds)
    except Exception:
        st.info("No predictions saved yet.")
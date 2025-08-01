import streamlit as st
import pandas as pd
import os
from fpdf import FPDF

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# Append new data to Excel workbook
def append_to_excel(new_data, sheet_name, file_name=EXCEL_FILE, deduplicate=True):
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

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

# ----- Fetch weekly data via nfl_data_py -----
with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pulls 2025 weekly team stats for CHI and saves to Excel.")
    fetch_week = st.number_input(
        "Week to fetch (2025 season)", min_value=1, max_value=25, value=1, step=1, key="fetch_week_2025"
    )

    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl

            # Try to refresh local cache (safe if it fails)
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

                # Offense fields (best-effort mapping)
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

                # Defense (team-level; advanced items need PBP)
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].values[0]
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
                st.caption("Note: Red Zone % Allowed and pressures require play-by-play aggregation; we can add that next.")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

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
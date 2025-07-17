
import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Upload weekly CSV data and track performance across the season.")

excel_file = "bears_historical_data.xlsx"

# Load existing or create new Excel workbook
def load_excel():
    if os.path.exists(excel_file):
        return pd.ExcelFile(excel_file)
    return None

# Upload section
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")

if uploaded_personnel:

    df_personnel = pd.read_csv(uploaded_personnel)

    append_to_excel(df_personnel, "Personnel")

    st.sidebar.success("✅ Personnel data uploaded and added.")



# Process uploaded files
def append_to_excel(new_df, sheet_name):
    if os.path.exists(excel_file):
        with pd.ExcelWriter(excel_file, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            existing = pd.read_excel(writer, sheet_name=sheet_name) if sheet_name in writer.book.sheetnames else pd.DataFrame()
            combined = pd.concat([existing, new_df], ignore_index=True)
            writer.book.remove(writer.book[sheet_name])
            combined.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(excel_file, mode="w", engine="openpyxl") as writer:
            new_df.to_excel(writer, sheet_name=sheet_name, index=False)

# Handle uploads
if uploaded_offense:
    df_off = pd.read_csv(uploaded_offense)
    append_to_excel(df_off, "Offense")
    st.success("✅ Offensive data uploaded and added.")

if uploaded_defense:
    df_def = pd.read_csv(uploaded_defense)
    append_to_excel(df_def, "Defense")
    st.success("✅ Defensive data uploaded and added.")

if uploaded_strategy:
    df_strat = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strat, "Strategy")
    st.success("✅ Strategy data uploaded and added.")

# Load Excel file and display
excel_data = load_excel()
if excel_data:
    for sheet in excel_data.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        st.subheader(f"📂 {sheet} Data")
        st.dataframe(df)

    with open(excel_file, "rb") as f:
        st.download_button(
            label="📥 Download Full Excel File",
            data=f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload weekly data files to begin tracking.")
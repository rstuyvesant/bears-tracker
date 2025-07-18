import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# Append new data to Excel workbook
def append_to_excel(new_data, sheet_name, file_name="bears_weekly_analytics.xlsx"):
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    # Load existing Excel file or create new one
    if os.path.exists(file_name):
        book = openpyxl.load_workbook(file_name)
        if sheet_name in book.sheetnames:
            sheet = book[sheet_name]
            # Read existing data into DataFrame
            existing_data = pd.DataFrame(sheet.values)
            existing_data.columns = existing_data.iloc[0]
            existing_data = existing_data[1:]

            # Drop duplicate week if it exists
            week_col = "Week"
            existing_data = existing_data[existing_data[week_col] != str(new_data.iloc[0][week_col])]

            # Combine and overwrite
            combined_data = pd.concat([existing_data, new_data], ignore_index=True)
        else:
            combined_data = new_data
    else:
        book = openpyxl.Workbook()
        book.remove(book.active)
        combined_data = new_data

    # Overwrite or add the updated sheet
    if sheet_name in book.sheetnames:
        del book[sheet_name]
    sheet = book.create_sheet(sheet_name)

    for r in dataframe_to_rows(combined_data, index=False, header=True):
        sheet.append(r)

    book.save(file_name)

# Upload section
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")

if uploaded_offense:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense")
    st.sidebar.success("✅ Offensive data uploaded and added.")

if uploaded_defense:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense")
    st.sidebar.success("✅ Defensive data uploaded and added.")

if uploaded_strategy:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy")
    st.sidebar.success("✅ Strategy data uploaded and added.")

if uploaded_personnel:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel")
    st.sidebar.success("✅ Personnel data uploaded and added.")

# Download section
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Optional preview
st.markdown("### 📊 Data Preview (latest upload)")
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
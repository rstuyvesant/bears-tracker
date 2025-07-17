import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Chicago Bears 2025–26 Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Upload new weekly data, view current stats, and export to Excel.")

# Set file paths
data_files = {
    "Personnel Usage": "bears_personnel_usage.csv",
    "Offensive Analytics": "bears_offensive_analytics.csv",
    "Defensive Analytics": "bears_defensive_analytics.csv",
    "Weekly Strategy": "bears_weekly_strategy.csv",
}

# Sidebar for uploads
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_files = {}

for label, filename in data_files.items():
    uploaded = st.sidebar.file_uploader(f"Upload {label}", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded)
        df.to_csv(filename, index=False)
        st.sidebar.success(f"{label} uploaded!")
        uploaded_files[label] = df

# Load and display data
st.header("📊 Current Weekly Data")

for label, filename in data_files.items():
    st.subheader(f"📁 {label}")
    if os.path.exists(filename):
        df = pd.read_csv(filename)
        st.dataframe(df)
    else:
        st.warning(f"{label} file not found.")

# 📥 Export Button
st.markdown("---")
if st.button("📤 Export All Data to Excel"):
    with pd.ExcelWriter("bears_combined_export.xlsx") as writer:
        for label, filename in data_files.items():
            if os.path.exists(filename):
                df = pd.read_csv(filename)
                df.to_excel(writer, sheet_name=label[:31], index=False)
    st.success("✅ Exported all data to 'bears_combined_export.xlsx'")
    with open("bears_combined_export.xlsx", "rb") as f:
        st.download_button(
            label="⬇️ Download Excel File",
            data=f,
            file_name="bears_combined_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
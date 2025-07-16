import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

# Use relative file paths for Streamlit deployment
personnel_path = "bears_personnel_usage.csv"
offense_path = "bears_offensive_analytics.csv"
defense_path = "bears_defensive_analytics.csv"
strategy_path = "bears_weekly_strategy.csv"

if os.path.exists(personnel_path):
    st.subheader("👥 Weekly Personnel Usage")
    st.dataframe(pd.read_csv(personnel_path))
else:
    st.warning("Personnel usage file not found.")

if os.path.exists(offense_path):
    st.subheader("📊 Offensive Analytics")
    st.dataframe(pd.read_csv(offense_path))
else:
    st.warning("Offensive analytics file not found.")

if os.path.exists(defense_path):
    st.subheader("🛡️ Defensive Analytics")
    st.dataframe(pd.read_csv(defense_path))
else:
    st.warning("Defensive analytics file not found.")

if os.path.exists(strategy_path):
    st.subheader("📋 Weekly Strategy Summary")
    st.dataframe(pd.read_csv(strategy_path))
else:
    st.info("Strategy summary file not found.")
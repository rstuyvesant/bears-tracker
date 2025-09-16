# bears_dashboard.py
# Chicago Bears Weekly Tracker (Week-vs-Opponent + YTD vs NFL build)
# - # ==============================
# 2) Upload Weekly Data
# ==============================
with st.expander("2) Upload Weekly Data", expanded=True):
    st.caption("Upload Offense/Defense/Personnel/SnapCounts CSVs. Rows are appended and deduped.")
    upc1, upc2 = st.columns(2)

    # ---------- Left column ----------
    with upc1:
        st.markdown("**Offense CSV**")
        f_off = st.file_uploader("Upload Offense CSV", type=["csv"], key="up_off")
        if f_off is not None:
            try:
                df = pd.read_csv(f_off)
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_OFFENSE, df, dedupe_on=["season", "week", "team"])
                st.success(f"Offense rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")

        st.markdown("**Defense CSV**")
        f_def = st.file_uploader("Upload Defense CSV", type=["csv"], key="up_def")
        if f_def is not None:
            try:
                df = pd.read_csv(f_def)
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_DEFENSE, df, dedupe_on=["season", "week", "team"])
                st.success(f"Defense rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")

    # ---------- Right column ----------
    with upc2:
        st.markdown("**Personnel CSV**")
        f_per = st.file_uploader("Upload Personnel CSV", type=["csv"], key="up_per")
        if f_per is not None:
            try:
                df = pd.read_csv(f_per)
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.loc[:, ~df.columns.duplicated()]
                df_out = append_to_sheet(EXCEL_PATH, SHEET_PERSONNEL, df, dedupe_on=["season", "week", "team"])
                st.success(f"Personnel rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")

        st.markdown("**Snap Counts CSV (Manual)**")
        f_snap = st.file_uploader("Upload Snap Counts CSV", type=["csv"], key="up_snap")
        if f_snap is not None:
            try:
                df = pd.read_csv(f_snap)
                df["season"] = df.get("season", int(season_input))
                df["week"] = df.get("week", int(week_input))
                df["team"] = df.get("team", NFL_TEAM)
                df = df.loc[:, ~df.columns.duplicated()]
                dedupe_cols = ["season", "week", "team"]
                if "player" in df.columns:
                    dedupe_cols.append("player")
                df_out = append_to_sheet(EXCEL_PATH, SHEET_SNAP_COUNTS, df, dedupe_on=dedupe_cols)
                st.success(f"SnapCounts rows now: {len(df_out)}")
            except Exception as e:
                st.error(f"Upload failed: {e}")

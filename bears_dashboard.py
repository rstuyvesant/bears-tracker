import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown(
    "Track weekly stats, strategy, personnel usage, injuries, snap counts, and league comparisons."
)

EXCEL_FILE = "bears_weekly_analytics.xlsx"


# =========================
# Utilities
# =========================
def safe_read_excel(path: str, sheet: str) -> pd.DataFrame:
    """Return a DataFrame if the sheet exists, else empty DataFrame."""
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()


def append_to_excel(new_df: pd.DataFrame, sheet_name: str, key_cols=None):
    """
    Append/merge rows into a sheet.
    - If the workbook/sheet exists, read it first, concat, drop duplicates on key_cols (if provided), write back.
    - If it doesn't exist, create it.
    """
    if new_df is None or len(new_df) == 0:
        return

    # Normalize columns to strings to avoid mismatches
    new_df = new_df.copy()
    new_df.columns = [str(c) for c in new_df.columns]

    if os.path.exists(EXCEL_FILE):
        try:
            existing = safe_read_excel(EXCEL_FILE, sheet_name)
            if not existing.empty:
                existing.columns = [str(c) for c in existing.columns]
                combined = pd.concat([existing, new_df], ignore_index=True)
            else:
                combined = new_df
        except Exception:
            combined = new_df

        # De-duplicate if key columns are provided
        if key_cols:
            has_all_keys = all(k in combined.columns for k in key_cols)
            if has_all_keys:
                combined = combined.drop_duplicates(subset=key_cols, keep="last")
            else:
                combined = combined.drop_duplicates(keep="last")
        else:
            combined = combined.drop_duplicates(keep="last")

        # Write back: replace the target sheet, keep others
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            new_df.to_excel(writer, sheet_name=sheet_name, index=False)


# =========================
# Color styling helpers
# =========================
def style_offense(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """
    Highlights Offense sheet:
      - YPA:  >= 7.5 green, 6.0–7.49 yellow, < 6.0 red
      - CMP%: >= 68  green, 60–67.9 yellow, < 60  red
      - Explosive_Play%: >= 12 green, 8–11.9 yellow, < 8 red (if present)
      - Drive_Success%:  >= 75 green, 65–74.9 yellow, < 65 red (if present)
    """
    df = df.copy()

    def color_ypa(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 7.5:
            return "background-color:#d1f5e0"  # green-tint
        if x >= 6.0:
            return "background-color:#fff9d6"  # yellow
        return "background-color:#fde2e2"      # red-tint

    def color_cmp(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 68:
            return "background-color:#d1f5e0"
        if x >= 60:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def color_explosive(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 12:
            return "background-color:#d1f5e0"
        if x >= 8:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def color_dsr(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 75:
            return "background-color:#d1f5e0"
        if x >= 65:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    styler = df.style

    if "YPA" in df.columns:
        styler = styler.applymap(color_ypa, subset=["YPA"])
    if "CMP%" in df.columns:
        styler = styler.applymap(color_cmp, subset=["CMP%"])
    # Optional columns if you added them
    if "Explosive_Play%" in df.columns:
        styler = styler.applymap(color_explosive, subset=["Explosive_Play%"])
    if "Drive_Success%" in df.columns:
        styler = styler.applymap(color_dsr, subset=["Drive_Success%"])

    return styler


def style_defense(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """
    Highlights Defense sheet:
      - RZ% Allowed: < 50 green, 50–65 yellow, > 65 red
      - SACK:        >= 3 green, 1–2 yellow, 0 red
      - Pressures:   >= 10 green, 6–9 yellow, <= 5 red (if present)
      - Explosive_Allowed%: < 8 green, 8–11.9 yellow, >= 12 red (if present)
    """
    df = df.copy()

    def color_rz(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x < 50:
            return "background-color:#d1f5e0"
        if x <= 65:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def color_sacks(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 3:
            return "background-color:#d1f5e0"
        if x >= 1:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def color_pressures(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 10:
            return "background-color:#d1f5e0"
        if x >= 6:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def color_expl_allowed(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x < 8:
            return "background-color:#d1f5e0"
        if x < 12:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    styler = df.style
    if "RZ% Allowed" in df.columns:
        styler = styler.applymap(color_rz, subset=["RZ% Allowed"])
    if "SACK" in df.columns:
        styler = styler.applymap(color_sacks, subset=["SACK"])
    if "Pressures" in df.columns:
        styler = styler.applymap(color_pressures, subset=["Pressures"])
    if "Explosive_Allowed%" in df.columns:
        styler = styler.applymap(color_expl_allowed, subset=["Explosive_Allowed%"])
    return styler


def style_dvoa_proxy(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """
    Highlights DVOA-like proxy sheet:
      - Off Adj EPA/play:  >= +0.15 green, +0.05–+0.149 yellow, < +0.05 red
      - Def Adj EPA/play:  <= -0.05 green, -0.049–+0.05 yellow, > +0.05 red
      - Off Adj SR%:       >= +5 green, +1–+4.9 yellow, < +1 red
      - Def Adj SR%:       <= -5 green, -4.9–+1 yellow, > +1 red
    """
    df = df.copy()

    def col_off_epa(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 0.15:
            return "background-color:#d1f5e0"
        if x >= 0.05:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def col_def_epa(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x <= -0.05:
            return "background-color:#d1f5e0"
        if x <= 0.05:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def col_off_sr(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x >= 5:
            return "background-color:#d1f5e0"
        if x >= 1:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    def col_def_sr(v):
        try:
            x = float(v)
        except Exception:
            return ""
        if x <= -5:
            return "background-color:#d1f5e0"
        if x <= 1:
            return "background-color:#fff9d6"
        return "background-color:#fde2e2"

    styler = df.style
    if "Off Adj EPA/play" in df.columns:
        styler = styler.applymap(col_off_epa, subset=["Off Adj EPA/play"])
    if "Def Adj EPA/play" in df.columns:
        styler = styler.applymap(col_def_epa, subset=["Def Adj EPA/play"])
    if "Off Adj SR%" in df.columns:
        styler = styler.applymap(col_off_sr, subset=["Off Adj SR%"])
    if "Def Adj SR%" in df.columns:
        styler = styler.applymap(col_def_sr, subset=["Def Adj SR%"])
    return styler


def style_plain(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Neutral zebra-striping for general tables."""
    return df.style.set_properties(**{"background-color": "#ffffff"}).apply(
        lambda _: ["background-color: #f9f9fb" if i % 2 else "" for i in range(len(df))],
        axis=0,
    )


# =========================
# Uploads (center panel)
# =========================
st.header("📤 Upload Weekly Data")

up_off = st.file_uploader("Upload Offense CSV", type=["csv"], key="up_off")
if up_off:
    df_off = pd.read_csv(up_off)
    append_to_excel(df_off, "Offense", key_cols=[c for c in ["Week", "Opponent"] if c in df_off.columns])
    st.success("✅ Offense data uploaded.")

up_def = st.file_uploader("Upload Defense CSV", type=["csv"], key="up_def")
if up_def:
    df_def = pd.read_csv(up_def)
    append_to_excel(df_def, "Defense", key_cols=[c for c in ["Week", "Opponent"] if c in df_def.columns])
    st.success("✅ Defense data uploaded.")

up_str = st.file_uploader("Upload Strategy CSV", type=["csv"], key="up_str")
if up_str:
    df_str = pd.read_csv(up_str)
    append_to_excel(df_str, "Strategy", key_cols=[c for c in ["Week"] if c in df_str.columns])
    st.success("✅ Strategy data uploaded.")

up_per = st.file_uploader("Upload Personnel CSV", type=["csv"], key="up_per")
if up_per:
    df_per = pd.read_csv(up_per)
    append_to_excel(df_per, "Personnel", key_cols=[c for c in ["Week"] if c in df_per.columns])
    st.success("✅ Personnel data uploaded.")

up_inj = st.file_uploader("Upload Injuries CSV", type=["csv"], key="up_inj")
if up_inj:
    df_inj = pd.read_csv(up_inj)
    append_to_excel(df_inj, "Injuries", key_cols=[c for c in ["Week", "Player"] if c in df_inj.columns])
    st.success("✅ Injuries uploaded.")

up_snaps = st.file_uploader("Upload Snap Counts CSV", type=["csv"], key="up_snaps")
if up_snaps:
    df_snaps = pd.read_csv(up_snaps)
    append_to_excel(df_snaps, "SnapCounts", key_cols=[c for c in ["Week", "Player"] if c in df_snaps.columns])
    st.success("✅ Snap counts uploaded.")


# =========================
# Quick Injury Entry
# =========================
st.header("➕ Quick Injury Entry")
c1, c2, c3 = st.columns([2, 1.2, 1.2])
with c1:
    q_player = st.text_input("Player")
with c2:
    q_status = st.selectbox("Status", ["Questionable", "Doubtful", "Out", "IR", "PUP", "Healthy"])
with c3:
    q_week = st.number_input("Week", min_value=1, max_value=25, step=1, value=1)

if st.button("Add Injury"):
    if q_player:
        row = pd.DataFrame([{"Week": q_week, "Player": q_player, "Status": q_status}])
        append_to_excel(row, "Injuries", key_cols=["Week", "Player"])
        st.success(f"✅ Injury added for {q_player} (Week {q_week}).")
    else:
        st.info("Enter a player name.")


# =========================
# Fetch placeholders (sidebar)
# =========================
with st.sidebar:
    st.header("🔄 Fetch & Compute (best-effort)")
    if st.button("Fetch Weekly Data (nfl_data_py)"):
        st.info("Best-effort fetch would run here (requires nfl_data_py).")
    if st.button("Fetch Defensive Metrics from Play-by-Play"):
        st.info("Best-effort PBP defensive metrics would run here.")
    if st.button("Fetch Snap Counts"):
        st.info("Best-effort snap counts would run here.")
    if st.button("Compute DVOA-like Proxy (Opponent-Adjusted)"):
        st.info("Best-effort DVOA-like proxy would run here.")


# =========================
# Data Review (with color)
# =========================
st.header("📊 Data Review")

if os.path.exists(EXCEL_FILE):
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        for sheet in xls.sheet_names:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
            st.subheader(f"Sheet: {sheet}")

            if df.empty:
                st.info("No rows yet.")
                continue

            # Apply sheet-specific styling
            if sheet.lower() == "offense":
                st.dataframe(style_offense(df), use_container_width=True)
            elif sheet.lower() == "defense":
                st.dataframe(style_defense(df), use_container_width=True)
            elif sheet.lower() == "dvoa_proxy":
                st.dataframe(style_dvoa_proxy(df), use_container_width=True)
            elif sheet.lower() in {"injuries", "snapcounts", "personnel", "strategy", "predictions"}:
                st.dataframe(style_plain(df), use_container_width=True)
            else:
                st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"Failed to read workbook: {e}")
else:
    st.info("No Excel workbook yet. Upload CSVs or use Quick Injury Entry to create it.")


# =========================
# Download workbook
# =========================
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            "⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
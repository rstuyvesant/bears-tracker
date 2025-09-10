## bears_dashboard.py
# Simple Bears Weekly Tracker (NO Pre/Final/PDF exports)

from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st

# ----------------------------
# App basics
# ----------------------------
st.set_page_config(page_title="Bears Weekly Tracker", layout="wide")
st.title("Bears Weekly Tracker (Simple)")

BASE_DIR = Path(__file__).parent.resolve()
DATA_DIR = BASE_DIR / "data"
MASTER_XLSX = DATA_DIR / "bears_weekly_analytics.xlsx"

DATA_DIR.mkdir(parents=True, exist_ok=True)

SHEETS = {
    "offense": "Offense",
    "defense": "Defense",
    "personnel": "Personnel",
    "snap": "Snap_Counts",
    "inj": "Injuries",
    "media": "Media",
    "opp": "Opponent_Preview",
    "pred": "Predictions",
    "notes": "Weekly_Notes",
}

NFL_TEAMS = [
    "ARI","ATL","BAL","BUF","CAR","CHI","CIN","CLE","DAL","DEN","DET","GB","HOU","IND",
    "JAX","KC","LAC","LAR","LV","MIA","MIN","NE","NO","NYG","NYJ","PHI","PIT","SEA","SF","TB","TEN","WAS"
]
def week_options():
    return [f"W{str(i).zfill(2)}" for i in range(1, 24)]

# ----------------------------
# Excel helpers
# ----------------------------
def ensure_master():
    if MASTER_XLSX.exists():
        return
    with pd.ExcelWriter(MASTER_XLSX, engine="openpyxl", mode="w") as w:
        pd.DataFrame(columns=["Week","Opponent"]).to_excel(w, index=False, sheet_name=SHEETS["offense"])
        pd.DataFrame(columns=["Week","Opponent"]).to_excel(w, index=False, sheet_name=SHEETS["defense"])
        pd.DataFrame(columns=["Week","Opponent","11","12","13","21","Other"]).to_excel(w, index=False, sheet_name=SHEETS["personnel"])
        pd.DataFrame(columns=["Week","Opponent","Player","Snaps","Snap%","Side"]).to_excel(w, index=False, sheet_name=SHEETS["snap"])
        pd.DataFrame(columns=["Week","Opponent","Player","Status","BodyPart","Practice","GameStatus","Notes"]).to_excel(w, index=False, sheet_name=SHEETS["inj"])
        pd.DataFrame(columns=["Week","Opponent","Source","Summary"]).to_excel(w, index=False, sheet_name=SHEETS["media"])
        pd.DataFrame(columns=["Week","Opponent","Off_Summary","Def_Summary","Matchups"]).to_excel(w, index=False, sheet_name=SHEETS["opp"])
        pd.DataFrame(columns=["Week","Opponent","Predicted_Winner","Confidence","Rationale"]).to_excel(w, index=False, sheet_name=SHEETS["pred"])
        pd.DataFrame(columns=["Week","Opponent","Notes"]).to_excel(w, index=False, sheet_name=SHEETS["notes"])

def read_sheet(name):
    ensure_master()
    try:
        return pd.read_excel(MASTER_XLSX, sheet_name=name)
    except Exception:
        return pd.DataFrame()

def write_sheet(df, name):
    ensure_master()
    with pd.ExcelWriter(MASTER_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        (df if not df.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name=name)

def append_to_sheet(new_rows, name, dedup_cols=None):
    old = read_sheet(name)
    combined = pd.concat([old, new_rows], ignore_index=True)
    if dedup_cols and all(c in combined.columns for c in dedup_cols):
        combined = combined.drop_duplicates(subset=dedup_cols, keep="last").reset_index(drop=True)
    write_sheet(combined, name)

# ----------------------------
# Sidebar controls
# ----------------------------
with st.sidebar:
    st.header("Weekly Controls")
    week = st.selectbox("Week", week_options(), index=0)
    opponent = st.selectbox("Opponent", NFL_TEAMS, index=NFL_TEAMS.index("MIN") if "MIN" in NFL_TEAMS else 0)

    st.markdown("---")
    st.subheader("Download Master (optional)")
    if MASTER_XLSX.exists():
        with open(MASTER_XLSX, "rb") as f:
            st.download_button(
                "Download Master Excel",
                data=f.read(),
                file_name="bears_weekly_analytics.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.caption("Master workbook will be created the first time you save something.")

# ----------------------------
# 1) Weekly Notes
# ----------------------------
st.markdown("### 1) Weekly Notes")
with st.expander("Notes"):
    st.info("Add reminders/highlights. Saved under the selected Week & Opponent.")
    note_text = st.text_area("Notes", height=110, placeholder="Key matchups, weather, personnel notes…")
    if st.button("Save Notes"):
        row = pd.DataFrame([{"Week": week, "Opponent": opponent, "Notes": note_text.strip()}])
        append_to_sheet(row, SHEETS["notes"], dedup_cols=["Week","Opponent"])
        st.success("Notes saved.")

# ----------------------------
# 2) Upload weekly data
# ----------------------------
st.markdown("### 2) Upload Weekly Data")
c1, c2 = st.columns(2)

with c1:
    st.subheader("Offense CSV")
    st.caption("Example: Week,Opponent,Points,Yards,YPA,CMP%,SR%,…")
    f = st.file_uploader("Upload Offense (.csv)", type=["csv"], key="off_upl")
    if f:
        try:
            df = pd.read_csv(f)
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["offense"], dedup_cols=["Week","Opponent"])
            st.success(f"Offense rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Offense upload failed: {e}")

    st.subheader("Defense CSV")
    st.caption("Example: Week,Opponent,SACK,INT,3D%_Allowed,RZ%_Allowed,Pressures,…")
    f = st.file_uploader("Upload Defense (.csv)", type=["csv"], key="def_upl")
    if f:
        try:
            df = pd.read_csv(f)
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["defense"], dedup_cols=["Week","Opponent"])
            st.success(f"Defense rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Defense upload failed: {e}")

with c2:
    st.subheader("Personnel CSV")
    st.info(
        "Personnel: Week,Opponent,11,12,13,21,Other | "
        "Snap_Counts: Week,Opponent,Player,Snaps,Snap%,Side."
    )
    f = st.file_uploader("Upload Personnel (.csv)", type=["csv"], key="per_upl")
    if f:
        try:
            df = pd.read_csv(f)
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["personnel"], dedup_cols=["Week","Opponent"])
            st.success(f"Personnel rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Personnel upload failed: {e}")

    st.subheader("Snap Counts CSV")
    st.caption("Columns: Week,Opponent,Player,Snaps,Snap%,Side")
    f = st.file_uploader("Upload Snap Counts (.csv)", type=["csv"], key="snap_upl")
    if f:
        try:
            df = pd.read_csv(f)
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["snap"], dedup_cols=["Week","Opponent","Player"])
            st.success(f"Snap Count rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Snap upload failed: {e}")

# ----------------------------
# 3) Opponent / Injuries / Media
# ----------------------------
st.markdown("### 3) Opponent Preview / Injuries / Media")
a, b, c = st.columns(3)

with a:
    st.subheader("Opponent Preview")
    opp_off = st.text_area("Offense Summary", height=110)
    opp_def = st.text_area("Defense Summary", height=110)
    opp_match = st.text_area("Key Matchups", height=110)
    if st.button("Save Opponent Preview"):
        row = pd.DataFrame([{
            "Week": week, "Opponent": opponent,
            "Off_Summary": opp_off.strip(),
            "Def_Summary": opp_def.strip(),
            "Matchups": opp_match.strip()
        }])
        append_to_sheet(row, SHEETS["opp"], dedup_cols=["Week","Opponent"])
        st.success("Opponent preview saved.")

with b:
    st.subheader("Injuries")
    p = st.text_input("Player")
    s = st.selectbox("Status", ["Questionable","Doubtful","Out","IR","Healthy"])
    bp = st.text_input("Body Part / Injury")
    prac = st.text_input("Practice (DNP/Limited/Full)")
    gstat = st.text_input("Game Status")
    notes = st.text_area("Notes", height=90)
    if st.button("Save Injury"):
        if p.strip():
            row = pd.DataFrame([{
                "Week": week, "Opponent": opponent,
                "Player": p.strip(), "Status": s, "BodyPart": bp.strip(),
                "Practice": prac.strip(), "GameStatus": gstat.strip(), "Notes": notes.strip()
            }])
            append_to_sheet(row, SHEETS["inj"], dedup_cols=["Week","Opponent","Player"])
            st.success("Injury saved.")
        else:
            st.error("Enter a player name.")

with c:
    st.subheader("Media Summaries")
    src = st.text_input("Source")
    summ = st.text_area("Summary", height=140)
    if st.button("Save Media Summary"):
        if src.strip() and summ.strip():
            row = pd.DataFrame([{
                "Week": week, "Opponent": opponent,
                "Source": src.strip(), "Summary": summ.strip()
            }])
            append_to_sheet(row, SHEETS["media"])
            st.success("Media summary saved.")
        else:
            st.error("Enter both Source and Summary.")

# ----------------------------
# 4) Prediction
# ----------------------------
st.markdown("### 4) Prediction")
x, y = st.columns(2)
with x:
    who = st.selectbox("Predicted Winner", ["CHI","OPP"], index=0)
    conf = st.slider("Confidence", 0, 100, 60)
with y:
    why = st.text_area("Rationale", height=120)
if st.button("Save Prediction"):
    row = pd.DataFrame([{
        "Week": week, "Opponent": opponent,
        "Predicted_Winner": "CHI" if who == "CHI" else opponent,
        "Confidence": conf,
        "Rationale": why.strip()
    }])
    append_to_sheet(row, SHEETS["pred"], dedup_cols=["Week","Opponent"])
    st.success("Prediction saved.")

# ----------------------------
# 5) Current Week Snapshots
# ----------------------------
st.markdown("### 5) Current Week Snapshots")
tabs = st.tabs(["Offense","Defense","Personnel","Snap Counts","Injuries","Media","Opponent","Prediction","Notes"])

def show_df(tab, sheet):
    with tab:
        df = read_sheet(sheet)
        if not df.empty and {"Week","Opponent"}.issubset(df.columns):
            df = df[(df["Week"].astype(str) == week) & (df["Opponent"].astype(str).str.upper() == opponent)]
        st.dataframe(df if not df.empty else pd.DataFrame({"Info": ["No rows for this week/opponent yet."]}))

show_df(tabs[0], SHEETS["offense"])
show_df(tabs[1], SHEETS["defense"])
show_df(tabs[2], SHEETS["personnel"])
show_df(tabs[3], SHEETS["snap"])
show_df(tabs[4], SHEETS["inj"])
show_df(tabs[5], SHEETS["media"])
show_df(tabs[6], SHEETS["opp"])
show_df(tabs[7], SHEETS["pred"])
show_df(tabs[8], SHEETS["notes"])

st.caption("Files are stored in ./data (master workbook). No PDF/Excel weekly exports in this simple version.")

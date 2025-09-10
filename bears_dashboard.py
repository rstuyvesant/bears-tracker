# bears_dashboard.py
# Bears Weekly Tracker — uploads, search, master workbook, and per-week downloads (Excel/PDF)

from pathlib import Path
import io
import pandas as pd
import numpy as np
import streamlit as st
from fpdf import FPDF

# ----------------------------
# App basics
# ----------------------------
st.set_page_config(page_title="Bears Weekly Tracker", layout="wide")
st.title("Bears Weekly Tracker")

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
    "strategy": "Weekly_Strategy",
}

NFL_TEAMS = [
    "ARI","ATL","BAL","BUF","CAR","CHI","CIN","CLE","DAL","DEN","DET","GB","HOU","IND",
    "JAX","KC","LAC","LAR","LV","MIA","MIN","NE","NO","NYG","NYJ","PHI","PIT","SEA","SF","TB","TEN","WAS"
]
def week_options():
    return [f"W{str(i).zfill(2)}" for i in range(1, 24)]

# ----------------------------
# Workbook helpers
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
        pd.DataFrame(columns=["Week","Opponent","Category","Detail"]).to_excel(w, index=False, sheet_name=SHEETS["strategy"])

def read_sheet(name: str) -> pd.DataFrame:
    ensure_master()
    try:
        return pd.read_excel(MASTER_XLSX, sheet_name=name)
    except Exception:
        return pd.DataFrame()

def write_sheet(df: pd.DataFrame, name: str):
    ensure_master()
    with pd.ExcelWriter(MASTER_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        (df if not df.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name=name)

def append_to_sheet(new_rows: pd.DataFrame, name: str, dedup_cols=None):
    old = read_sheet(name)
    combined = pd.concat([old, new_rows], ignore_index=True)
    # Dedup (avoid duplicate weeks / rows)
    if dedup_cols and all(c in combined.columns for c in dedup_cols):
        combined = combined.drop_duplicates(subset=dedup_cols, keep="last").reset_index(drop=True)
    write_sheet(combined, name)

# ----------------------------
# CSV/XLSX ingestion helpers
# ----------------------------
def _read_any_table(file) -> pd.DataFrame:
    """Read CSV or XLSX into DataFrame."""
    fname = getattr(file, "name", "").lower()
    if fname.endswith(".csv") or ".csv" in fname:
        return pd.read_csv(file)
    if fname.endswith(".xlsx") or ".xls" in fname:
        return pd.read_excel(file)
    # fallback: try CSV
    try:
        file.seek(0)
        return pd.read_csv(file)
    except Exception:
        file.seek(0)
        return pd.read_excel(file)

def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Make headers case-insensitive and map common synonyms."""
    def norm(s): return str(s).strip()
    df = df.copy()
    df.columns = [norm(c) for c in df.columns]

    # Build casefold mapping
    lower_map = {c.casefold(): c for c in df.columns}

    def has(*alts):
        return next((lower_map[a.casefold()] for a in alts if a.casefold() in lower_map), None)

    # Map Week
    wk = has("Week","week","wk","w")
    if wk and wk != "Week":
        df.rename(columns={wk:"Week"}, inplace=True)

    # Map Opponent
    opp = has("Opponent","opponent","opp","OPP","Opp")
    if opp and opp != "Opponent":
        df.rename(columns={opp:"Opponent"}, inplace=True)

    # Map Player
    pl = has("Player","player","Name","name")
    if pl and pl != "Player":
        df.rename(columns={pl:"Player"}, inplace=True)

    # Map Snaps
    sn = has("Snaps","snaps","Plays","plays")
    if sn and sn != "Snaps":
        df.rename(columns={sn:"Snaps"}, inplace=True)

    # Map Snap%
    sp = has("Snap%","snap%","snap_pct","SnapPct","snap pct","Snap Percent","snap_percent")
    if sp and sp != "Snap%":
        df.rename(columns={sp:"Snap%"}, inplace=True)

    # Side
    sd = has("Side","side","Unit","unit","Pos_Side","pos_side")
    if sd and sd != "Side":
        df.rename(columns={sd:"Side"}, inplace=True)

    return df

def _fill_missing_week_opp(df: pd.DataFrame, week: str, opp: str, allow_fill: bool) -> pd.DataFrame:
    df = df.copy()
    if allow_fill:
        if "Week" not in df.columns:
            df["Week"] = week
        df["Week"] = df["Week"].astype(str)
        if "Opponent" not in df.columns:
            df["Opponent"] = opp
        df["Opponent"] = df["Opponent"].astype(str).str.upper()
    return df

# ----------------------------
# Sidebar controls + downloads
# ----------------------------
with st.sidebar:
    st.header("Weekly Controls")
    week = st.selectbox("Week", week_options(), index=0)
    opponent = st.selectbox("Opponent", NFL_TEAMS, index=NFL_TEAMS.index("MIN") if "MIN" in NFL_TEAMS else 0)
    autofill = st.checkbox("Auto-fill missing Week/Opponent on upload", value=True,
                           help="If your CSV lacks these columns, the selected values will be used.")

    st.markdown("---")
    st.subheader("Downloads")
    dl_excel = st.button("Download Current Week (Excel)")
    dl_pdf = st.button("Download Current Week (PDF)")

    st.markdown("---")
    st.subheader("Master Workbook")
    if MASTER_XLSX.exists():
        with open(MASTER_XLSX, "rb") as f:
            st.download_button(
                "Download Master Excel",
                data=f.read(),
                file_name="bears_weekly_analytics.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.caption("Master workbook de-duplicates by key columns (e.g., Week/Opponent).")
    else:
        st.caption("The master workbook will be created the first time you save or upload.")

# ----------------------------
# 1) Weekly Notes
# ----------------------------
st.markdown("### 1) Weekly Notes")
with st.expander("Notes"):
    note_text = st.text_area("Notes", height=110, placeholder="Key matchups, reminders, personnel notes…")
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
    st.subheader("Offense")
    st.caption("Typical columns include: Week, Opponent, Points, Yards, YPA, CMP%, SR%, …")
    f = st.file_uploader("Upload Offense (.csv/.xlsx)", type=["csv","xlsx"], key="off_upl")
    if f:
        try:
            df = _read_any_table(f)
            df = _normalize_headers(df)
            df = _fill_missing_week_opp(df, week, opponent, autofill)
            # Require Week & Opponent after normalization/fill
            if not {"Week","Opponent"}.issubset(df.columns):
                raise ValueError("Your file must include columns Week and Opponent (or enable Auto-fill).")
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["offense"], dedup_cols=["Week","Opponent"])
            st.success(f"Offense rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Offense upload failed: {e}")

    st.subheader("Defense")
    st.caption("Typical columns: Week, Opponent, SACK, INT, 3D%_Allowed, RZ%_Allowed, Pressures, …")
    f = st.file_uploader("Upload Defense (.csv/.xlsx)", type=["csv","xlsx"], key="def_upl")
    if f:
        try:
            df = _read_any_table(f)
            df = _normalize_headers(df)
            df = _fill_missing_week_opp(df, week, opponent, autofill)
            if not {"Week","Opponent"}.issubset(df.columns):
                raise ValueError("Your file must include columns Week and Opponent (or enable Auto-fill).")
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["defense"], dedup_cols=["Week","Opponent"])
            st.success(f"Defense rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Defense upload failed: {e}")

with c2:
    st.subheader("Personnel")
    st.info(
        "Personnel: Week,Opponent,11,12,13,21,Other  |  "
        "Snap_Counts: Week,Opponent,Player,Snaps,Snap%,Side  |  "
        "Strategy: Week,Opponent,Category,Detail"
    )
    f = st.file_uploader("Upload Personnel (.csv/.xlsx)", type=["csv","xlsx"], key="per_upl")
    if f:
        try:
            df = _read_any_table(f)
            df = _normalize_headers(df)
            df = _fill_missing_week_opp(df, week, opponent, autofill)
            if not {"Week","Opponent"}.issubset(df.columns):
                raise ValueError("Your file must include Week and Opponent (or enable Auto-fill).")
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["personnel"], dedup_cols=["Week","Opponent"])
            st.success(f"Personnel rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Personnel upload failed: {e}")

    st.subheader("Snap Counts")
    f = st.file_uploader("Upload Snap Counts (.csv/.xlsx)", type=["csv","xlsx"], key="snap_upl")
    if f:
        try:
            df = _read_any_table(f)
            df = _normalize_headers(df)
            df = _fill_missing_week_opp(df, week, opponent, autofill)
            if not {"Week","Opponent"}.issubset(df.columns):
                raise ValueError("Your file must include Week and Opponent (or enable Auto-fill).")
            df["Week"] = df["Week"].astype(str)
            df["Opponent"] = df["Opponent"].astype(str).str.upper()
            append_to_sheet(df, SHEETS["snap"], dedup_cols=["Week","Opponent","Player"])
            st.success(f"Snap Count rows saved: {len(df)}")
        except Exception as e:
            st.error(f"Snap Counts upload failed: {e}")

st.subheader("Strategy")
st.caption("Optional weekly strategy notes. Columns: Week,Opponent,Category,Detail")
f = st.file_uploader("Upload Strategy (.csv/.xlsx)", type=["csv","xlsx"], key="strat_upl")
if f:
    try:
        df = _read_any_table(f)
        df = _normalize_headers(df)
        df = _fill_missing_week_opp(df, week, opponent, autofill)
        if not {"Week","Opponent"}.issubset(df.columns):
            raise ValueError("Your file must include Week and Opponent (or enable Auto-fill).")
        df["Week"] = df["Week"].astype(str)
        df["Opponent"] = df["Opponent"].astype(str).str.upper()
        append_to_sheet(df, SHEETS["strategy"])
        st.success(f"Strategy rows saved: {len(df)}")
    except Exception as e:
        st.error(f"Strategy upload failed: {e}")

# ----------------------------
# 3) Opponent / Injuries / Media / Prediction
# ----------------------------
st.markdown("### 3) Opponent / Injuries / Media / Prediction")
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
    st.subheader("Media")
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

st.subheader("Prediction")
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
# Build per-week package + Downloads
# ----------------------------
def _week_slice(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    if {"Week","Opponent"}.issubset(df.columns):
        return df[(df["Week"].astype(str)==week) & (df["Opponent"].astype(str).str.upper()==opponent)]
    return df

def build_week_package() -> dict:
    pkg = {}
    for k, sheet in SHEETS.items():
        pkg[sheet] = _week_slice(read_sheet(sheet))
    return pkg

def week_excel_bytes(pkg: dict) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl", mode="w") as w:
        for sheet_name, df in pkg.items():
            safe = sheet_name[:31]
            (df if not df.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name=safe)
    bio.seek(0)
    return bio.getvalue()

def week_pdf_bytes(pkg: dict, title: str) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=15)

    def H(t): pdf.set_font("Arial","B",14); pdf.cell(0,10,t,ln=True)
    def P(t): pdf.set_font("Arial","",11); pdf.multi_cell(0,6,t if t else "-")

    pdf.add_page()
    pdf.set_font("Arial","B",16)
    pdf.cell(0,10,title,ln=True)
    pdf.ln(4)

    # Notes
    H("Weekly Notes")
    notes = pkg.get(SHEETS["notes"], pd.DataFrame())
    P(str(notes["Notes"].iloc[0]) if not notes.empty and "Notes" in notes.columns else "")

    # Opponent
    pdf.ln(4); H("Opponent Preview")
    opp = pkg.get(SHEETS["opp"], pd.DataFrame())
    if not opp.empty:
        P("Offense: " + str(opp.get("Off_Summary", pd.Series([""])).iloc[0]))
        P("Defense: " + str(opp.get("Def_Summary", pd.Series([""])).iloc[0]))
        P("Matchups: " + str(opp.get("Matchups", pd.Series([""])).iloc[0]))
    else:
        P("")

    # Injuries
    pdf.ln(4); H("Injuries")
    inj = pkg.get(SHEETS["inj"], pd.DataFrame())
    if not inj.empty:
        for _, r in inj.iterrows():
            P(f"{r.get('Player','')} — {r.get('Status','')} — {r.get('BodyPart','')}; "
              f"Practice: {r.get('Practice','')}; Game: {r.get('GameStatus','')}; Notes: {r.get('Notes','')}")
    else:
        P("")

    # Media
    pdf.ln(4); H("Media Summaries")
    med = pkg.get(SHEETS["media"], pd.DataFrame())
    if not med.empty:
        for _, r in med.iterrows():
            P(f"{r.get('Source','')}: {r.get('Summary','')}")
    else:
        P("")

    # Prediction
    pdf.ln(4); H("Prediction")
    pr = pkg.get(SHEETS["pred"], pd.DataFrame())
    if not pr.empty:
        r = pr.iloc[-1]
        P(f"Predicted Winner: {r.get('Predicted_Winner','')}")
        P(f"Confidence: {r.get('Confidence','')}")
        P("Rationale: " + str(r.get("Rationale","")))
    else:
        P("")

    pdf.ln(6); pdf.set_font("Arial","I",9); pdf.cell(0,6,"Generated by Bears Weekly Tracker",ln=True)
    raw = pdf.output(dest="S").encode("latin1", "ignore")
    return raw

# Sidebar download actions
if dl_excel:
    pkg = build_week_package()
    data = week_excel_bytes(pkg)
    st.sidebar.download_button(
        "Save Current Week (Excel)",
        data=data,
        file_name=f"{week}_{opponent}_package.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if dl_pdf:
    pkg = build_week_package()
    data = week_pdf_bytes(pkg, title=f"Weekly Report — {week} vs {opponent}")
    st.sidebar.download_button(
        "Save Current Week (PDF)",
        data=data,
        file_name=f"{week}_{opponent}_report.pdf",
        mime="application/pdf",
    )

# ----------------------------
# 4) Searchable Snapshots
# ----------------------------
st.markdown("### 4) Current Week Snapshots (with Search)")
tabs = st.tabs(["Offense","Defense","Personnel","Snap Counts","Injuries","Media","Opponent","Prediction","Notes","Strategy"])

def _search_df(df: pd.DataFrame, q: str) -> pd.DataFrame:
    if not q or df.empty: 
        return df
    q = str(q).strip().lower()
    return df[[True]*len(df)] if q == "" else df[df.astype(str).apply(lambda col: col.str.lower().str.contains(q, na=False)).any(axis=1)]

def show_tab(tab, sheet, placeholder="Search…"):
    with tab:
        df = _week_slice(read_sheet(sheet))
        q = st.text_input(f"Search {sheet}", value="", placeholder=placeholder, key=f"q_{sheet}")
        sdf = _search_df(df, q)
        st.dataframe(sdf if not sdf.empty else pd.DataFrame({"Info":["No rows for this week/opponent (or search filtered all)."]}))

show_tab(tabs[0], SHEETS["offense"])
show_tab(tabs[1], SHEETS["defense"])
show_tab(tabs[2], SHEETS["personnel"])
show_tab(tabs[3], SHEETS["snap"], placeholder="Filter by player, side, etc.")
show_tab(tabs[4], SHEETS["inj"], placeholder="Filter by player/status…")
show_tab(tabs[5], SHEETS["media"], placeholder="Filter by source/summary…")
show_tab(tabs[6], SHEETS["opp"], placeholder="Filter opponent notes…")
show_tab(tabs[7], SHEETS["pred"], placeholder="Filter predictions…")
show_tab(tabs[8], SHEETS["notes"], placeholder="Filter notes…")
show_tab(tabs[9], SHEETS["strategy"], placeholder="Filter strategy…")

st.caption("Master workbook lives in ./data. Uploaders accept CSV or Excel. Week/Opponent are auto-filled if missing (toggle in sidebar).")

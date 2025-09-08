# bears_dashboard.py
# Restores Pre-Export and Final buttons, keeps "Export Weekly Final PDF"
# Compact, drop-in, and compatible with your existing weekly workflow.

import os
import io
import datetime as dt
import pandas as pd
import streamlit as st

# Optional PDF deps (graceful fallback to TXT if missing)
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# --- PAGE CONFIG ---
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.caption("Pre + Final exports restored • Works with your current weekly flow")

# --- CONSTANTS ---
EXCEL_FILE = "bears_weekly_analytics.xlsx"  # your main workbook
SHEETS = {
    "offense": "Offense",
    "defense": "Defense",
    "personnel": "Personnel",
    "snap_counts": "Snap_Counts",
    "injuries": "Injuries",
    "media": "Media_Summaries",
    "opponent": "Opponent_Preview",
    "pred": "Predictions",
    "nfl_manual": "NFL_Averages_Manual",
    "nfl_off": "YTD_NFL_Offense",
    "nfl_def": "YTD_NFL_Defense",
}

# ---------- UTILITIES ----------
def _ensure_excel(file_name: str):
    if not os.path.exists(file_name):
        with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
            # initialize all expected sheets empty (header-only)
            for s in SHEETS.values():
                pd.DataFrame().to_excel(writer, sheet_name=s, index=False)

def _load_sheet(sheet_name: str) -> pd.DataFrame:
    _ensure_excel(EXCEL_FILE)
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        # Normalize empty -> 0 rows DataFrame with no NaN column names
        if df.empty:
            return pd.DataFrame()
        df.columns = [str(c) for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

def _append_rows(df_new: pd.DataFrame, sheet_name: str, dedupe_on=None):
    _ensure_excel(EXCEL_FILE)
    df_old = _load_sheet(sheet_name)
    if df_old.empty:
        df_out = df_new.copy()
    else:
        df_out = pd.concat([df_old, df_new], ignore_index=True)

    # optional dedupe
    if dedupe_on and set(dedupe_on).issubset(df_out.columns):
        df_out = df_out.drop_duplicates(subset=dedupe_on, keep="last")

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_out.to_excel(writer, sheet_name=sheet_name, index=False)

def _week_code(week: int) -> str:
    # W01, W02, ...
    return f"W{int(week):02d}"

def _now_str() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M")

def _pdf_from_sections(title: str, week: int, opponent: str, blocks: list) -> bytes:
    """
    Generate a simple, reliable PDF using reportlab if available, else return txt bytes.
    blocks: list of tuples (section_title, df_or_text)
    """
    if not REPORTLAB_AVAILABLE:
        # Fallback TXT so you can still download something if reportlab isn't present.
        text = [f"{title}",
                f"Week: {week} ({_week_code(week)})",
                f"Opponent: {opponent}",
                f"Generated: {_now_str()}",
                "-"*60]
        for head, data in blocks:
            text.append(f"\n[{head}]")
            if isinstance(data, pd.DataFrame) and not data.empty:
                text.append(data.to_string(index=False))
            elif isinstance(data, pd.DataFrame) and data.empty:
                text.append("(no data)")
            else:
                text.append(str(data) if data else "(no data)")
        txt = "\n".join(text)
        return txt.encode("utf-8")

    # PDF path
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=LETTER)
    width, height = LETTER
    left = 0.75 * inch
    top = height - 0.75 * inch
    y = top

    def draw_line(s, font="Helvetica", size=10, leading=12):
        nonlocal y
        c.setFont(font, size)
        for line in s.splitlines() if isinstance(s, str) else [str(s)]:
            if y < 1 * inch:
                c.showPage()
                y = top
                c.setFont(font, size)
            c.drawString(left, y, line)
            y -= leading

    # Header
    draw_line(title, "Helvetica-Bold", 14, 18)
    draw_line(f"Week: {week} ({_week_code(week)})", "Helvetica", 10, 14)
    draw_line(f"Opponent: {opponent}", "Helvetica", 10, 14)
    draw_line(f"Generated: {_now_str()}", "Helvetica-Oblique", 9, 12)
    draw_line("-" * 90, "Helvetica", 10, 12)

    # Body
    for head, data in blocks:
        draw_line(f"[{head}]", "Helvetica-Bold", 11, 14)
        if isinstance(data, pd.DataFrame):
            if data.empty:
                draw_line("(no data)")
            else:
                # Render DataFrame as text table
                draw_line(" | ".join(data.columns))
                for _, row in data.iterrows():
                    draw_line(" | ".join([str(x) for x in row.values]))
        else:
            draw_line(str(data) if data else "(no data)")
        draw_line("")  # spacer

    c.save()
    buf.seek(0)
    return buf.read()

def _gather_week_blocks(week: int) -> list:
    """Collects key sections for the exports."""
    wk = int(week)
    wk_col = "Week"  # assumed column present across weekly tables

    # pull data
    offense = _load_sheet(SHEETS["offense"])
    defense = _load_sheet(SHEETS["defense"])
    personnel = _load_sheet(SHEETS["personnel"])
    snaps = _load_sheet(SHEETS["snap_counts"])
    injuries = _load_sheet(SHEETS["injuries"])
    media = _load_sheet(SHEETS["media"])
    opp = _load_sheet(SHEETS["opponent"])
    pred = _load_sheet(SHEETS["pred"])

    # filter by week if column present
    def f(df):
        return df[df[wk_col] == wk] if (not df.empty and wk_col in df.columns) else df

    offense_w = f(offense)
    defense_w = f(defense)
    personnel_w = f(personnel)
    snaps_w = f(snaps)
    injuries_w = f(injuries)
    media_w = f(media)
    opp_w = f(opp)
    pred_w = f(pred)

    return [
        ("Opponent Preview", opp_w),
        ("Key Injuries", injuries_w),
        ("Offense (weekly rows)", offense_w),
        ("Defense (weekly rows)", defense_w),
        ("Personnel Usage", personnel_w),
        ("Snap Counts", snaps_w),
        ("Media Summaries", media_w),
        ("Predictions", pred_w),
    ]

def _opponent_for_week(week: int) -> str:
    df = _load_sheet(SHEETS["opponent"])
    if not df.empty:
        wk_col = "Week"
        opp_col = "Opponent"
        if wk_col in df.columns and opp_col in df.columns:
            row = df[df[wk_col] == int(week)]
            if not row.empty:
                return str(row.iloc[0][opp_col])
    return ""

# ---------- SIDEBAR ----------
with st.sidebar:
    st.header("League Data & Utilities")
    st.write("These are independent of the weekly center sections.")

    # (Placeholder) Fetch NFL Data (Auto)
    # Keep as a stub so it doesn't break your current workflow
    if st.button("Fetch NFL Data (Auto)"):
        st.success("Fetched/updated league data (placeholder).")

    # Download all data (Excel)
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                "Download All Data (Excel)",
                data=f.read(),
                file_name=EXCEL_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("No Excel yet — it will be created as you save data.")

# ---------- WEEKLY CONTROLS ----------
st.subheader("Weekly Controls")
colA, colB, colC = st.columns([1,1,2])
with colA:
    week = st.number_input("Week", min_value=1, max_value=25, value=1, step=1)
with colB:
    opp_default = _opponent_for_week(week)
    opponent = st.text_input("Opponent (3-letter code preferred)", value=opp_default)

st.divider()

# ---------- DATA ENTRY / UPLOAD SECTIONS (compact versions) ----------

st.subheader("Upload Weekly Data")
c1, c2, c3, c4 = st.columns(4)

with c1:
    up_off = st.file_uploader("Upload Offense CSV", type=["csv"])
    if up_off is not None:
        try:
            df = pd.read_csv(up_off)
            _append_rows(df, SHEETS["offense"], dedupe_on=["Week","Opponent"] if {"Week","Opponent"}.issubset(df.columns) else None)
            st.success("Offense saved.")
        except Exception as e:
            st.error(f"Offense upload failed: {e}")

with c2:
    up_def = st.file_uploader("Upload Defense CSV", type=["csv"])
    if up_def is not None:
        try:
            df = pd.read_csv(up_def)
            _append_rows(df, SHEETS["defense"], dedupe_on=["Week","Opponent"] if {"Week","Opponent"}.issubset(df.columns) else None)
            st.success("Defense saved.")
        except Exception as e:
            st.error(f"Defense upload failed: {e}")

with c3:
    up_per = st.file_uploader("Upload Personnel CSV", type=["csv"])
    if up_per is not None:
        try:
            df = pd.read_csv(up_per)
            _append_rows(df, SHEETS["personnel"], dedupe_on=["Week","Opponent"] if {"Week","Opponent"}.issubset(df.columns) else None)
            st.success("Personnel saved.")
        except Exception as e:
            st.error(f"Personnel upload failed: {e}")

with c4:
    up_snap = st.file_uploader("Upload Snap Counts CSV", type=["csv"])
    if up_snap is not None:
        try:
            df = pd.read_csv(up_snap)
            _append_rows(df, SHEETS["snap_counts"], dedupe_on=["Week","Opponent","Player"] if {"Week","Opponent","Player"}.issubset(df.columns) else None)
            st.success("Snap counts saved.")
        except Exception as e:
            st.error(f"Snap counts upload failed: {e}")

st.divider()

st.subheader("Injuries (quick add)")
inj_cols = st.columns(5)
with inj_cols[0]:
    inj_player = st.text_input("Player")
with inj_cols[1]:
    inj_status = st.selectbox("Status", ["", "Out", "Doubtful", "Questionable", "Probable"])
with inj_cols[2]:
    inj_body = st.text_input("BodyPart / Notes")
with inj_cols[3]:
    inj_week = st.number_input("Week (inj)", min_value=1, max_value=25, value=int(week))
with inj_cols[4]:
    if st.button("Save Injury"):
        row = pd.DataFrame([{
            "Week": int(inj_week),
            "Opponent": opponent,
            "Player": inj_player,
            "Status": inj_status,
            "Notes": inj_body,
            "SavedAt": _now_str()
        }])
        _append_rows(row, SHEETS["injuries"], dedupe_on=["Week","Opponent","Player"])
        st.success("Injury saved.")

st.divider()

st.subheader("Opponent Preview & Media Summaries")
pcol1, pcol2 = st.columns(2)

with pcol1:
    st.markdown("**Opponent Preview**")
    opp_notes = st.text_area("Notes", height=140, key="opp_notes")
    if st.button("Save Opponent Preview"):
        row = pd.DataFrame([{
            "Week": int(week),
            "Opponent": opponent,
            "Notes": opp_notes,
            "SavedAt": _now_str()
        }])
        _append_rows(row, SHEETS["opponent"], dedupe_on=["Week","Opponent"])

with pcol2:
    st.markdown("**Media Summary**")
    media_source = st.text_input("Source (e.g., ESPN, The Athletic)")
    media_summary = st.text_area("Summary", height=140)
    if st.button("Save Media Summary"):
        row = pd.DataFrame([{
            "Week": int(week),
            "Opponent": opponent,
            "Source": media_source,
            "Summary": media_summary,
            "SavedAt": _now_str()
        }])
        _append_rows(row, SHEETS["media"])
        st.success("Media summary saved.")

st.divider()

st.subheader("Weekly Game Predictions")
pred_cols = st.columns([1,1,1,2])
with pred_cols[0]:
    predicted_winner = st.selectbox("Predicted Winner", ["", "CHI", opponent])
with pred_cols[1]:
    confidence = st.slider("Confidence (%)", 0, 100, 60)
with pred_cols[2]:
    rationale_short = st.text_input("Rationale (short)")
with pred_cols[3]:
    rationale = st.text_area("Rationale (details)", height=100)

if st.button("Save Prediction"):
    row = pd.DataFrame([{
        "Week": int(week),
        "Opponent": opponent,
        "Predicted_Winner": predicted_winner,
        "Confidence": confidence,
        "Rationale_Short": rationale_short,
        "Rationale": rationale,
        "SavedAt": _now_str()
    }])
    _append_rows(row, SHEETS["pred"], dedupe_on=["Week","Opponent"])
    st.success("Prediction saved.")

st.divider()

# ---------- EXPORTS (Pre + Final restored) ----------
st.subheader("Exports")

def _build_pre_report(week: int, opponent: str) -> bytes:
    blocks = _gather_week_blocks(week)
    title = "Chicago Bears — PRE-GAME Weekly Report"
    return _pdf_from_sections(title, week, opponent, blocks)

def _build_final_report(week: int, opponent: str) -> bytes:
    blocks = _gather_week_blocks(week)
    title = "Chicago Bears — FINAL Weekly Report"
    return _pdf_from_sections(title, week, opponent, blocks)

ec1, ec2, ec3 = st.columns(3)

with ec1:
    if st.button("Export Pre-Game PDF"):  # <-- RESTORED Pre-Export button
        pdf = _build_pre_report(week, opponent)
        fname = f"{_week_code(week)}_Pre.pdf" if REPORTLAB_AVAILABLE else f"{_week_code(week)}_Pre.txt"
        mime = "application/pdf" if REPORTLAB_AVAILABLE else "text/plain"
        st.download_button("Download Pre-Game Report", data=pdf, file_name=fname, mime=mime)

with ec2:
    if st.button("Export Final PDF"):     # <-- RESTORED Final Export button
        pdf = _build_final_report(week, opponent)
        fname = f"{_week_code(week)}_Final.pdf" if REPORTLAB_AVAILABLE else f"{_week_code(week)}_Final.txt"
        mime = "application/pdf" if REPORTLAB_AVAILABLE else "text/plain"
        st.download_button("Download Final Report", data=pdf, file_name=fname, mime=mime)

with ec3:
    # Your existing button kept for compatibility (calls the same final export)
    st.markdown("**Export Weekly Final PDF** *(kept for compatibility)*")
    st.caption("This produces the same PDF as “Export Final PDF”.")
    wk_to_export = st.number_input("Week to export", min_value=1, max_value=25, value=int(week), key="wk_export_final")
    if st.button("Export final pdf"):
        pdf = _build_final_report(wk_to_export, _opponent_for_week(wk_to_export) or opponent)
        fname = f"{_week_code(wk_to_export)}_Final.pdf" if REPORTLAB_AVAILABLE else f"{_week_code(wk_to_export)}_Final.txt"
        mime = "application/pdf" if REPORTLAB_AVAILABLE else "text/plain"
        st.download_button("Download Final Report (compat)", data=pdf, file_name=fname, mime=mime, key="dl_final_compat")

st.info("Tip: You can export as many times as you want. Use Pre-Game during the week; use Final after you’ve finished entering data.") Personnel: Week,Opponent,11,12,13,21,Other | Snap_Counts: Week,Opponent,Player,Snaps,Snap%,Side.")










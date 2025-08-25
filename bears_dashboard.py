# bears_dashboard.py
# Chicago Bears 2025–26 Weekly Tracker (All-in-One Streamlit App)
# -------------------------------------------------------------------
# Features
# - Upload weekly CSVs: Offense, Defense, Personnel, Strategy, Injuries, Snap Counts, Opponent Preview
# - Store everything in a single Excel file with de-duplication per sheet
# - Media Summaries (ESPN/Beat writers/etc) + storage
# - Predictions (basic example) + storage
# - DVOA-like Proxy (example using uploaded metrics)
# - Auto-computed YTD (team & NFL average) and included in PDF export
# - Sidebar: Compute NFL Averages, Download All Data (Excel)
# - Main sections: clear, ordered layout
#
# Notes
# - Adjust column names in METRIC_SCHEMA below to match your actual CSV headers.
# - This app does not rely on internet access; it works with your local CSV uploads.
# - If a sheet/file doesn't exist yet, it is created automatically.
# -------------------------------------------------------------------

import os
import io
import math
from typing import Dict, List, Tuple

import streamlit as st
import pandas as pd

# Excel/Report deps
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet


# ---------------------------- App Config ----------------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel, injuries, snap counts, media, predictions, and league comparisons.")

# Single master workbook for everything
EXCEL_FILE = "bears_weekly_analytics.xlsx"

# Default sheets used by the app
SHEETS = {
    "Offense": "Offense",
    "Defense": "Defense",
    "Personnel": "Personnel",
    "Strategy": "Strategy",
    "Injuries": "Injuries",
    "SnapCounts": "SnapCounts",
    "MediaSummaries": "MediaSummaries",
    "Predictions": "Predictions",
    "OpponentPreview": "OpponentPreview",
    "NFL_Averages": "NFL_Averages",   # optional league-wide manual reference
}

# Some helpful defaults
DEFAULT_TEAM = "CHI"  # Use abbreviation present in your CSVs
DEFAULT_WEEK = 1


# ---------------------------- Excel Helpers ----------------------------
def _safe_load_book(file_name: str):
    if os.path.exists(file_name):
        try:
            return openpyxl.load_workbook(file_name)
        except Exception:
            # Corrupt or open elsewhere; create new
            return openpyxl.Workbook()
    else:
        return openpyxl.Workbook()

def _sheet_to_df(book, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in book.sheetnames:
        return pd.DataFrame()
    ws = book[sheet_name]
    data = list(ws.values)
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data)
    # first row as header
    df.columns = df.iloc[0]
    df = df[1:]
    # reset index
    df = df.reset_index(drop=True)
    return df

def _ensure_sheet(book, sheet_name: str):
    if sheet_name not in book.sheetnames:
        book.create_sheet(sheet_name)
    # remove default 'Sheet' if empty/unwanted
    if "Sheet" in book.sheetnames and len(book.sheetnames) > 1:
        ws = book["Sheet"]
        if ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None:
            book.remove(ws)

def _write_df_to_sheet(book, sheet_name: str, df: pd.DataFrame):
    _ensure_sheet(book, sheet_name)
    ws = book[sheet_name]
    ws.delete_rows(1, ws.max_row or 1)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

def append_to_excel(new_df: pd.DataFrame, sheet_name: str, dedup_cols: List[str]):
    """
    Append new_df rows into sheet_name in EXCEL_FILE, then deduplicate by dedup_cols.
    Creates the workbook/sheet if needed.
    """
    book = _safe_load_book(EXCEL_FILE)
    _ensure_sheet(book, sheet_name)

    existing_df = _sheet_to_df(book, sheet_name)
    if existing_df.empty:
        combined = new_df.copy()
    else:
        combined = pd.concat([existing_df, new_df], ignore_index=True)

    # Deduplicate if possible
    if dedup_cols:
        # only keep columns that actually exist
        dedup_cols_present = [c for c in dedup_cols if c in combined.columns]
        if dedup_cols_present:
            combined = combined.drop_duplicates(subset=dedup_cols_present, keep="last")

    # Normalize column order: keep existing headers first, then new ones
    if not existing_df.empty:
        ordered_cols = list(existing_df.columns) + [c for c in combined.columns if c not in existing_df.columns]
        combined = combined.reindex(columns=ordered_cols)

    _write_df_to_sheet(book, sheet_name, combined)
    book.save(EXCEL_FILE)
    return combined


# ---------------------------- YTD / NFL Avg Schema ----------------------------
# Define how metrics should aggregate in YTD
# "sum" -> add across weeks; "mean" -> average across weeks
METRIC_SCHEMA: Dict[str, str] = {
    # Common Offense metrics (adjust to match your files)
    "Points": "sum",
    "Yards": "sum",
    "YPA": "mean",
    "YPC": "mean",
    "CMP%": "mean",
    "QBR": "mean",
    "SR%": "mean",
    "3D%": "mean",
    "RZ%": "mean",

    # Common Defense metrics
    "SACK": "sum",
    "INT": "sum",
    "FF": "sum",
    "FR": "sum",
    "QB Hits": "sum",
    "Pressures": "sum",
    "DVOA": "mean",
    "3D% Allowed": "mean",
    "RZ% Allowed": "mean",
}

def _select_metric_columns(df, metric_schema):
    return [c for c in metric_schema.keys() if c in df.columns]

def _agg_dict_for(df, metric_schema):
    cols = _select_metric_columns(df, metric_schema)
    return {c: metric_schema[c] for c in cols}

def compute_ytd_team(df, team_name: str, week: int) -> Tuple[int, dict]:
    """Compute team YTD up to and including `week`."""
    if df is None or df.empty:
        return 0, {}
    try:
        sub = df[(df["Team"] == team_name) & (df["Week"].astype(int) <= int(week))]
    except Exception:
        # If Week isn't int, try coercion
        temp = df.copy()
        temp["Week"] = pd.to_numeric(temp["Week"], errors="coerce")
        sub = temp[(temp["Team"] == team_name) & (temp["Week"] <= int(week))]

    if sub.empty:
        return 0, {}

    agg_map = _agg_dict_for(sub, METRIC_SCHEMA)
    if not agg_map:
        return 0, {}

    ytd = sub.agg(agg_map)
    games = sub["Week"].nunique()
    return games, ytd.to_dict()

def compute_ytd_nfl_avg(df, week: int) -> Tuple[int, dict]:
    """
    Compute NFL average YTD (per-team) up to and including `week`.
    Steps:
      - Filter to weeks <= selected week
      - Group by Team; aggregate by METRIC_SCHEMA
      - Average those per-team values across teams
    """
    if df is None or df.empty:
        return 0, {}
    temp = df.copy()
    temp["Week"] = pd.to_numeric(temp["Week"], errors="coerce")
    sub = temp[temp["Week"] <= int(week)]
    if sub.empty:
        return 0, {}

    agg_map = _agg_dict_for(sub, METRIC_SCHEMA)
    if not agg_map:
        return 0, {}

    per_team = sub.groupby("Team").agg(agg_map)
    if per_team.empty:
        return 0, {}

    nfl_avg = per_team.mean(numeric_only=True)
    return per_team.shape[0], nfl_avg.to_dict()

def _format_display_table(team_dict, nfl_dict, label_cols: List[str]) -> pd.DataFrame:
    keys = sorted(set(team_dict.keys()) | set(nfl_dict.keys()))
    rows = []
    for k in keys:
        tv = team_dict.get(k, None)
        nv = nfl_dict.get(k, None)

        def _fmt(x):
            if x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))):
                return None
            if isinstance(x, (int, float)):
                if "%" in k:
                    return f"{x:.1f}%"
                return f"{x:.2f}"
            return x

        rows.append({"Metric": k, label_cols[0]: _fmt(tv), label_cols[1]: _fmt(nv)})

    out = pd.DataFrame(rows)
    # Keep schema order where possible
    schema_order = [k for k in METRIC_SCHEMA.keys() if k in out["Metric"].values]
    out["order"] = out["Metric"].apply(lambda m: schema_order.index(m) if m in schema_order else 999)
    out = out.sort_values("order").drop(columns=["order"]).reset_index(drop=True)
    return out


# ---------------------------- ReportLab Helpers ----------------------------
def _df_to_rl_table(df, col_widths=None):
    data = [list(df.columns)] + df.astype(str).values.tolist()
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EEEEEE")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 9),
        ("FONTSIZE", (0,1), (-1,-1), 8),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#FAFAFA")]),
    ]))
    return t

def generate_weekly_pdf(
    output_path: str,
    selected_week: int,
    selected_team: str,
    opponent_name: str,
    key_notes: str,
    offense_ytd_display_df: pd.DataFrame,
    defense_ytd_display_df: pd.DataFrame,
    other_sections: dict = None
):
    doc = SimpleDocTemplate(output_path, pagesize=letter, topMargin=36, bottomMargin=36, leftMargin=36, rightMargin=36)
    styles = getSampleStyleSheet()
    story = []

    title = f"Chicago Bears — Weekly Game Report (Week {selected_week})"
    story.append(Paragraph(title, styles["Title"]))
    story.append(Paragraph(f"Team: {selected_team}", styles["Normal"]))
    story.append(Paragraph(f"Opponent: {opponent_name}", styles["Normal"]))
    story.append(Spacer(1, 6))

    if key_notes:
        story.append(Paragraph("<b>Key Notes:</b> " + key_notes, styles["BodyText"]))
        story.append(Spacer(1, 10))

    # YTD Offense/Defense
    story.append(Paragraph("Offense — YTD vs NFL Avg", styles["Heading3"]))
    story.append(_df_to_rl_table(offense_ytd_display_df, col_widths=[150, 170, 170]))
    story.append(Spacer(1, 12))

    story.append(Paragraph("Defense — YTD vs NFL Avg", styles["Heading3"]))
    story.append(_df_to_rl_table(defense_ytd_display_df, col_widths=[150, 170, 170]))
    story.append(Spacer(1, 14))

    if other_sections:
        for heading, df in other_sections.items():
            story.append(Paragraph(heading, styles["Heading3"]))
            story.append(_df_to_rl_table(df))
            story.append(Spacer(1, 10))

    doc.build(story)


# ---------------------------- Sidebar ----------------------------
st.sidebar.header("⚙️ Actions")
with st.sidebar.expander("Compute NFL Averages (Manual/Optional)"):
    st.write("If you keep a league-wide reference CSV, you can upload it and store it to the **NFL_Averages** sheet.")
    nfl_csv = st.file_uploader("Upload NFL Averages CSV", type=["csv"], key="nfl_avg_upload")
    if nfl_csv is not None:
        df_nfl_avg = pd.read_csv(nfl_csv)
        combined = append_to_excel(df_nfl_avg, SHEETS["NFL_Averages"], dedup_cols=[])
        st.success(f"Stored {len(df_nfl_avg)} rows to sheet '{SHEETS['NFL_Averages']}'.")
        st.dataframe(combined.tail(10), use_container_width=True)

st.sidebar.markdown("---")
# Download the master Excel file
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button("⬇️ Download All Data (Excel)", data=f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.sidebar.info("Excel file will appear here after your first upload/save.")


# ---------------------------- State (Week/Team) ----------------------------
if "selected_week" not in st.session_state:
    st.session_state.selected_week = DEFAULT_WEEK
if "selected_team" not in st.session_state:
    st.session_state.selected_team = DEFAULT_TEAM
if "opponent_name" not in st.session_state:
    st.session_state.opponent_name = "TBD"
if "key_notes" not in st.session_state:
    st.session_state.key_notes = ""


# ---------------------------- Page Controls ----------------------------
st.markdown("### Weekly Controls")
colA, colB, colC = st.columns(3)
with colA:
    st.session_state.selected_week = st.number_input("Week", min_value=1, max_value=22, value=st.session_state.selected_week, step=1)
with colB:
    st.session_state.selected_team = st.text_input("Team (abbrev as used in your CSVs)", value=st.session_state.selected_team)
with colC:
    st.session_state.opponent_name = st.text_input("Opponent", value=st.session_state.opponent_name)

st.session_state.key_notes = st.text_area("Key Notes (appears in PDF)", value=st.session_state.key_notes, height=80)


# ---------------------------- Upload Sections ----------------------------
st.markdown("---")
st.header("1) Upload Weekly Data")

def _upload_section(title: str, sheet_key: str, dedup_cols: List[str], help_text: str = ""):
    st.subheader(title)
    if help_text:
        st.caption(help_text)
    file = st.file_uploader(f"Upload {title} CSV", type=["csv"], key=f"up_{sheet_key}")
    if file is not None:
        try:
            df = pd.read_csv(file)
        except Exception:
            file.seek(0)
            df = pd.read_csv(file, encoding_errors="ignore")
        # Add Week if missing; many files already include Week
        if "Week" not in df.columns:
            df["Week"] = st.session_state.selected_week
        combined = append_to_excel(df, SHEETS[sheet_key], dedup_cols=dedup_cols)
        st.success(f"Saved to sheet '{SHEETS[sheet_key]}' ({len(df)} new rows).")
        st.dataframe(df.head(20), use_container_width=True)

_upload_section("Offensive Analytics", "Offense",
                dedup_cols=["Week", "Team"],
                help_text="Include columns like Team, Week, Points, Yards, YPA, YPC, CMP%, QBR, SR%, 3D%, RZ%, etc.")

_upload_section("Defensive Analytics", "Defense",
                dedup_cols=["Week", "Team"],
                help_text="Include columns like Team, Week, SACK, INT, FF, FR, QB Hits, Pressures, DVOA, 3D% Allowed, RZ% Allowed, etc.")

_upload_section("Personnel Usage", "Personnel",
                dedup_cols=["Week", "Team"],
                help_text="e.g., counts by personnel groupings (11, 12, 13, 21).")

_upload_section("Strategy Notes", "Strategy",
                dedup_cols=["Week", "Team"],
                help_text="Pre/Post-game strategy notes used by your prediction block.")

_upload_section("Injuries", "Injuries",
                dedup_cols=["Week", "Team", "Player"],
                help_text="Injury reports (e.g., Player, Status, Practice, GameStatus).")

_upload_section("Snap Counts", "SnapCounts",
                dedup_cols=["Week", "Team", "Player"],
                help_text="Player snap counts and % (e.g., Player, Snaps, Snap%).")

_upload_section("Opponent Preview", "OpponentPreview",
                dedup_cols=["Week", "Team", "Opponent"],
                help_text="Any opponent scouting/notes for the selected week.")

st.markdown("---")


# ---------------------------- Load Current Week Data ----------------------------
def _load_sheet(sheet_name: str) -> pd.DataFrame:
    book = _safe_load_book(EXCEL_FILE)
    return _sheet_to_df(book, sheet_name)

offense_df_full = _load_sheet(SHEETS["Offense"])
defense_df_full = _load_sheet(SHEETS["Defense"])
personnel_df_full = _load_sheet(SHEETS["Personnel"])
strategy_df_full = _load_sheet(SHEETS["Strategy"])
injuries_df_full = _load_sheet(SHEETS["Injuries"])
snap_df_full = _load_sheet(SHEETS["SnapCounts"])
opp_prev_df_full = _load_sheet(SHEETS["OpponentPreview"])
media_df_full = _load_sheet(SHEETS["MediaSummaries"])
pred_df_full = _load_sheet(SHEETS["Predictions"])

# Best-effort coercions
def _to_int(x):
    try:
        return int(x)
    except Exception:
        return pd.NA

for frame in [offense_df_full, defense_df_full, personnel_df_full, strategy_df_full,
              injuries_df_full, snap_df_full, opp_prev_df_full]:
    if not frame.empty and "Week" in frame.columns:
        frame["Week"] = pd.to_numeric(frame["Week"], errors="coerce").astype("Int64")


selected_week = int(st.session_state.selected_week)
selected_team = st.session_state.selected_team


# ---------------------------- Current-Week Previews ----------------------------
st.header("2) This Week — Data Previews")

def _preview_by_week(df: pd.DataFrame, label: str):
    st.subheader(label)
    if df.empty:
        st.info("No data yet.")
        return pd.DataFrame()
    if "Week" not in df.columns:
        st.dataframe(df.tail(20), use_container_width=True)
        return df
    sub = df[df["Week"] == selected_week]
    if sub.empty:
        st.info(f"No rows for Week {selected_week}.")
    else:
        st.dataframe(sub, use_container_width=True)
    return sub

offense_cur = _preview_by_week(offense_df_full, "Offense — Weekly Rows")
defense_cur = _preview_by_week(defense_df_full, "Defense — Weekly Rows")
personnel_cur = _preview_by_week(personnel_df_full, "Personnel — Weekly Rows")
strategy_cur = _preview_by_week(strategy_df_full, "Strategy — Weekly Rows")
injuries_cur = _preview_by_week(injuries_df_full, "Injuries — Weekly Rows")
snap_cur = _preview_by_week(snap_df_full, "Snap Counts — Weekly Rows")
opp_prev_cur = _preview_by_week(opp_prev_df_full, "Opponent Preview — Weekly Rows")

st.markdown("---")


# ---------------------------- Media Summaries ----------------------------
st.header("3) Media Summaries (Store multiple per week)")
with st.expander("Add a Media Summary"):
    source = st.text_input("Source (ESPN, The Athletic, BearsWire, etc.)", value="")
    url = st.text_input("URL (optional)", value="")
    summary = st.text_area("Summary (short paragraph)", height=120)
    if st.button("Save Summary"):
        row = pd.DataFrame([{
            "Week": selected_week,
            "Team": selected_team,
            "Source": source,
            "URL": url,
            "Summary": summary
        }])
        combined = append_to_excel(row, SHEETS["MediaSummaries"], dedup_cols=["Week", "Team", "Source", "URL", "Summary"])
        st.success("Saved summary.")
        st.dataframe(combined.tail(10), use_container_width=True)

st.subheader("All Summaries (Week)")
if not media_df_full.empty:
    ms = media_df_full.copy()
    if "Week" in ms.columns:
        ms["Week"] = pd.to_numeric(ms["Week"], errors="coerce").astype("Int64")
        ms = ms[ms["Week"] == selected_week]
    st.dataframe(ms, use_container_width=True)
else:
    st.info("No media summaries yet.")

st.markdown("---")


# ---------------------------- DVOA-like Proxy (Example) ----------------------------
st.header("4) DVOA-like Proxy (Demo)")
st.caption("Example proxy using a few metrics from Offense/Defense if present. Adjust as you see fit.")

def _safe_float(x):
    try:
        return float(x)
    except Exception:
        return None

def compute_dvoa_proxy(off_sub: pd.DataFrame, def_sub: pd.DataFrame) -> pd.DataFrame:
    """
    Example proxy: combine a few rate/impact metrics into a single score.
    Modify weights/metrics to your preference.
    """
    # pick one row for selected_team if multiple; or average them
    def pick_team_mean(df):
        if df.empty:
            return {}
        if "Team" in df.columns:
            df = df[df["Team"] == selected_team] if selected_team in df["Team"].unique() else df
        numeric = {}
        for col in ["YPA", "CMP%", "RZ%", "SACK", "INT", "QB Hits", "Pressures", "DVOA"]:
            if col in df.columns:
                vals = pd.to_numeric(df[col], errors="coerce")
                numeric[col] = vals.mean(skipna=True)
        return numeric

    off_vals = pick_team_mean(off_sub)
    def_vals = pick_team_mean(def_sub)

    # very simple illustrative formula
    # offense positive: YPA, CMP%, RZ%
    # defense positive: SACK, INT, QB Hits, Pressures; (lower DVOA is better, invert if present)
    score = 0.0
    if "YPA" in off_vals and off_vals["YPA"] is not None:
        score += 1.5 * off_vals["YPA"]
    if "CMP%" in off_vals and off_vals["CMP%"] is not None:
        score += 0.2 * off_vals["CMP%"]
    if "RZ%" in off_vals and off_vals["RZ%"] is not None:
        score += 0.2 * off_vals["RZ%"]

    if "SACK" in def_vals and def_vals["SACK"] is not None:
        score += 0.6 * def_vals["SACK"]
    if "INT" in def_vals and def_vals["INT"] is not None:
        score += 1.0 * def_vals["INT"]
    if "QB Hits" in def_vals and def_vals["QB Hits"] is not None:
        score += 0.2 * def_vals["QB Hits"]
    if "Pressures" in def_vals and def_vals["Pressures"] is not None:
        score += 0.1 * def_vals["Pressures"]

    if "DVOA" in def_vals and def_vals["DVOA"] is not None:
        score += (-0.5) * def_vals["DVOA"]  # lower DVOA better

    return pd.DataFrame([{"Week": selected_week, "Team": selected_team, "DVOA_Proxy": round(score, 3)}])

dvoa_proxy_df = compute_dvoa_proxy(offense_cur, defense_cur)
st.dataframe(dvoa_proxy_df, use_container_width=True)

st.markdown("---")


# ---------------------------- Predictions (Demo) ----------------------------
st.header("5) Weekly Game Predictions (Demo)")
st.caption("Simple example that you can expand. Stores one row per week.")

pred_reason = st.text_area("Prediction Rationale (uses strategy + proxy, etc.)", height=100)
pred_outcome = st.selectbox("Predicted Outcome", ["Win", "Loss", "Toss-up"])
if st.button("Save Prediction"):
    row = pd.DataFrame([{
        "Week": selected_week,
        "Team": selected_team,
        "Opponent": st.session_state.opponent_name,
        "Prediction": pred_outcome,
        "Reason": pred_reason,
        "DVOA_Proxy": dvoa_proxy_df["DVOA_Proxy"].iloc[0] if not dvoa_proxy_df.empty else None,
    }])
    combined = append_to_excel(row, SHEETS["Predictions"], dedup_cols=["Week", "Team"])
    st.success("Prediction saved.")
    st.dataframe(combined.tail(10), use_container_width=True)

st.markdown("---")


# ---------------------------- Auto YTD + NFL Avg (and preview) ----------------------------
st.header("6) YTD Summary (Auto‑computed from uploads)")
st.caption("Derived from your weekly Offense/Defense sheets—no YTD values need to be stored in CSVs.")

# Build combined OFFENSE and DEFENSE frames limited to Week <= selected
offense_df = offense_df_full.copy()
defense_df = defense_df_full.copy()
for df_ in [offense_df, defense_df]:
    if not df_.empty and "Week" in df_.columns:
        df_["Week"] = pd.to_numeric(df_["Week"], errors="coerce")

off_gms, off_ytd_team = compute_ytd_team(offense_df, selected_team, selected_week)
off_tms, off_ytd_nfl = compute_ytd_nfl_avg(offense_df, selected_week)

def_gms, def_ytd_team = compute_ytd_team(defense_df, selected_team, selected_week)
def_tms, def_ytd_nfl = compute_ytd_nfl_avg(defense_df, selected_week)

offense_ytd_display = _format_display_table(
    off_ytd_team, off_ytd_nfl,
    [f"{selected_team} YTD (W1–W{selected_week}, {off_gms} gms)",
     f"NFL Avg YTD (W1–W{selected_week}, {off_tms} teams)"]
)
defense_ytd_display = _format_display_table(
    def_ytd_team, def_ytd_nfl,
    [f"{selected_team} YTD (W1–W{selected_week}, {def_gms} gms)",
     f"NFL Avg YTD (W1–W{selected_week}, {def_tms} teams)"]
)

with st.expander("📈 Show YTD Tables"):
    st.subheader("Offense — YTD vs NFL Avg")
    st.dataframe(offense_ytd_display, use_container_width=True)
    st.subheader("Defense — YTD vs NFL Avg")
    st.dataframe(defense_ytd_display, use_container_width=True)

st.markdown("---")


# ---------------------------- Weekly PDF Export ----------------------------
st.header("7) Weekly Game Report PDF")

# You can include extra tables (raw weekly views) in PDF if desired:
other_sections = {}

if not offense_cur.empty:
    other_sections["Weekly Offense (This Week)"] = offense_cur
if not defense_cur.empty:
    other_sections["Weekly Defense (This Week)"] = defense_cur
if not personnel_cur.empty:
    other_sections["Personnel (This Week)"] = personnel_cur
if not strategy_cur.empty:
    other_sections["Strategy (This Week)"] = strategy_cur
if not injuries_cur.empty:
    other_sections["Injuries (This Week)"] = injuries_cur
if not snap_cur.empty:
    other_sections["Snap Counts (This Week)"] = snap_cur
if not opp_prev_cur.empty:
    other_sections["Opponent Preview (This Week)"] = opp_prev_cur
if not dvoa_proxy_df.empty:
    other_sections["DVOA Proxy (This Week)"] = dvoa_proxy_df

pdf_name = f"W{int(selected_week):02d}_Final.pdf"
if st.button("Generate Weekly Game Report PDF"):
    generate_weekly_pdf(
        output_path=pdf_name,
        selected_week=selected_week,
        selected_team=selected_team,
        opponent_name=st.session_state.opponent_name,
        key_notes=st.session_state.key_notes,
        offense_ytd_display_df=offense_ytd_display,
        defense_ytd_display_df=defense_ytd_display,
        other_sections=other_sections
    )
    st.success(f"Created {pdf_name}. Use the button below to save it.")

if os.path.exists(pdf_name):
    with open(pdf_name, "rb") as f:
        st.download_button("⬇️ Save PDF", data=f, file_name=pdf_name, mime="application/pdf")


# ---------------------------- Page Footer ----------------------------
st.markdown("---")
st.caption("Tip: Keep your CSV headers consistent week to week. Adjust METRIC_SCHEMA keys to match your files exactly.")
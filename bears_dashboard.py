# bears_dashboard.py
# Chicago Bears 2025–26 Weekly Tracker (Streamlit)
# Includes: guarded reportlab import, colored tables, YTD + NFL Avg in Excel & PDF,
# Pre/Post/Final exports, uploads (off/def/personnel/snap counts/injuries), media summaries,
# opponent preview, simple prediction capture, and safe guards around YTD.

import os
import io
import json
from datetime import datetime
import pandas as pd
import streamlit as st

# ========== Safe imports & environment guards ==========
try:
    import openpyxl
    from openpyxl.styles import PatternFill
    from openpyxl.formatting.rule import ColorScaleRule
except Exception:
    pass

HAS_REPORTLAB = True
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors as RL_COLORS
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
except Exception:
    HAS_REPORTLAB = False

# ========== Constants & config ==========
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# Basic metric schema used for ordering and coloring
# (Add/remove metrics as your CSVs evolve; keys are the "Metric" labels we’ll display)
METRIC_SCHEMA = {
    "YPA": {"higher_is_better": True},
    "CMP%": {"higher_is_better": True},
    "RZ% Allowed": {"higher_is_better": False},
    "SACKs": {"higher_is_better": True},
    "INTs": {"higher_is_better": True},
    "QB Hits": {"higher_is_better": True},
    "Pressures": {"higher_is_better": True},
    "3D% Allowed": {"higher_is_better": False},
}

# ========== Utilities ==========
def _ensure_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            pd.DataFrame(columns=["Week"]).to_excel(writer, sheet_name="Offense", index=False)
            pd.DataFrame(columns=["Week"]).to_excel(writer, sheet_name="Defense", index=False)
            pd.DataFrame(columns=["Week"]).to_excel(writer, sheet_name="Personnel", index=False)
            pd.DataFrame(columns=["Week"]).to_excel(writer, sheet_name="SnapCounts", index=False)
            pd.DataFrame(columns=["Week","Injury","Status","Notes"]).to_excel(writer, sheet_name="Injuries", index=False)
            pd.DataFrame(columns=["Week","Source","Summary"]).to_excel(writer, sheet_name="MediaSummaries", index=False)
            pd.DataFrame(columns=["Week","Notes"]).to_excel(writer, sheet_name="OpponentPreview", index=False)
            pd.DataFrame(columns=["Week","Rationale","Prediction"]).to_excel(writer, sheet_name="Predictions", index=False)
            # YTD sheets will be rewritten each time
            pd.DataFrame().to_excel(writer, sheet_name="YTD_Team_Offense", index=False)
            pd.DataFrame().to_excel(writer, sheet_name="YTD_NFL_Offense", index=False)
            pd.DataFrame().to_excel(writer, sheet_name="YTD_Team_Defense", index=False)
            pd.DataFrame().to_excel(writer, sheet_name="YTD_NFL_Defense", index=False)

def _read_sheet(name: str) -> pd.DataFrame:
    _ensure_excel()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=name)
    except Exception:
        return pd.DataFrame()

def _append_df(sheet_name: str, df_new: pd.DataFrame, key_cols=None):
    _ensure_excel()
    if df_new is None or df_new.empty:
        return
    df_new = df_new.copy()
    df_old = _read_sheet(sheet_name)
    if df_old is None or df_old.empty:
        df_out = df_new
    else:
        df_out = pd.concat([df_old, df_new], ignore_index=True)
    # Deduplicate if key_cols provided
    if key_cols:
        df_out = df_out.drop_duplicates(subset=key_cols, keep="last")
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_out.to_excel(writer, sheet_name=sheet_name, index=False)

def _coerce_week_col(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    if "Week" not in df.columns:
        # try to find a week-like col
        for c in df.columns:
            if str(c).strip().lower() == "week":
                df = df.rename(columns={c:"Week"})
                break
    return df

def _calc_proxy(off_row: pd.Series=None, def_row: pd.Series=None) -> float | None:
    """
    Simple DVOA-like proxy using a few columns if present.
    Positive is "good". For defense, we flip some signs.
    """
    if off_row is None and def_row is None:
        return None
    score = 0.0
    count = 0
    def val(s, k): 
        try:
            return float(s.get(k)) if s is not None and k in s else None
        except Exception:
            return None

    # Offense boosts
    if off_row is not None:
        ypa = val(off_row, "YPA")
        cmp_ = val(off_row, "CMP%")
        if ypa is not None:
            score += ypa * 1.0; count += 1
        if cmp_ is not None:
            score += (cmp_ / 100.0) * 1.0; count += 1

    # Defense: lower RZ% Allowed is good, more SACKs good
    if def_row is not None:
        rz = val(def_row, "RZ% Allowed")
        sacks = val(def_row, "SACKs")
        if rz is not None:
            score += (1.0 - (rz / 100.0))  # lower better
            count += 1
        if sacks is not None:
            score += min(1.0, sacks / 5.0)  # normalize
            count += 1

    return round(score / count, 4) if count else None

def _merge_team_vs_nfl(team_df: pd.DataFrame, nfl_df: pd.DataFrame, as_metrics=True) -> pd.DataFrame:
    """
    Expect both to have columns named like metrics. If as_metrics=True, we reshape into
    rows of ["Metric", "Team", "NFL"] for display and coloring.
    """
    if team_df is None: team_df = pd.DataFrame()
    if nfl_df is None: nfl_df = pd.DataFrame()
    if team_df.empty and nfl_df.empty:
        return pd.DataFrame(columns=["Metric","Team","NFL"])
    # reduce to one row each (e.g., mean across YTD)
    if not team_df.empty:
        team_row = team_df.mean(numeric_only=True)
    else:
        team_row = pd.Series(dtype=float)
    if not nfl_df.empty:
        nfl_row = nfl_df.mean(numeric_only=True)
    else:
        nfl_row = pd.Series(dtype=float)
    # Build table
    metrics = list(METRIC_SCHEMA.keys())
    out = []
    for m in metrics:
        out.append({
            "Metric": m,
            "Team": team_row.get(m),
            "NFL": nfl_row.get(m)
        })
    return pd.DataFrame(out)

def _colorize_df(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Color Team vs NFL columns based on better/worse per METRIC_SCHEMA."""
    if df is None or df.empty:
        return df
    def _color_row(row):
        bg = []
        for col in df.columns:
            if col == "Metric":
                bg.append("")
                continue
            metric = row["Metric"]
            info = METRIC_SCHEMA.get(metric, {"higher_is_better": True})
            higher = info["higher_is_better"]
            team = row.get("Team")
            nfl = row.get("NFL")
            if col == "Team" and pd.notnull(team) and pd.notnull(nfl):
                better = (team > nfl) if higher else (team < nfl)
                color = "background-color: #d4edda" if better else "background-color: #f8d7da"
                bg.append(color)
            else:
                bg.append("")
        return bg
    return df.style.apply(_color_row, axis=1)

def _apply_excel_conditional_formatting(ws, first_data_row: int, team_col_letter: str, nfl_col_letter: str, last_row: int):
    """
    Green if team better than NFL (or lower for bad-is-better metrics), red otherwise.
    We use a two-color scale as a light approximation per row.
    """
    # This is a simple demo; real per-row comparisons would need per-row rules.
    # We at least show a gradient for the Team column.
    try:
        cs = ColorScaleRule(start_type='min', start_color='F8D7DA', end_type='max', end_color='D4EDDA')
        ws.conditional_formatting.add(f"{team_col_letter}{first_data_row}:{team_col_letter}{last_row}", cs)
    except Exception:
        pass

def _export_pdf_weekly(filename: str, team_title: str, notes: str,
                       off_table: pd.DataFrame|None, def_table: pd.DataFrame|None):
    if not HAS_REPORTLAB:
        return False, "reportlab not available"
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph(team_title, styles["Title"]))
    story.append(Spacer(1, 10))
    if notes:
        story.append(Paragraph("<b>Key Notes</b>", styles["Heading3"]))
        story.append(Paragraph(notes.replace("\n","<br/>"), styles["Normal"]))
        story.append(Spacer(1, 8))
    def _tbl(df, heading):
        if df is None or df.empty:
            return
        story.append(Paragraph(heading, styles["Heading3"]))
        data = [list(df.columns)] + df.fillna("").values.tolist()
        tbl = Table(data, hAlign="LEFT")
        style = [
            ("BACKGROUND",(0,0),(-1,0), RL_COLORS.lightgrey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1),0.25, RL_COLORS.grey),
            ("ALIGN",(1,1),(-1,-1),"RIGHT"),
        ]
        # Light green/red banding for Team column values (very simple heuristic)
        # Find Team column index
        try:
            team_idx = data[0].index("Team")
            nfl_idx = data[0].index("NFL")
            # Apply background per row comparing Team vs NFL using schema
            for r in range(1, len(data)):
                metric = data[r][0]
                team = data[r][team_idx]
                nfl = data[r][nfl_idx]
                info = METRIC_SCHEMA.get(metric, {"higher_is_better": True})
                higher = info["higher_is_better"]
                try:
                    if team != "" and nfl != "":
                        better = (float(team) > float(nfl)) if higher else (float(team) < float(nfl))
                        style.append(
                            ("BACKGROUND",(team_idx,r),(team_idx,r),
                             RL_COLORS.HexColor("#d4edda" if better else "#f8d7da"))
                        )
                except Exception:
                    pass
        except Exception:
            pass
        tbl.setStyle(TableStyle(style))
        story.append(tbl)
        story.append(Spacer(1, 10))
    _tbl(off_table, "Offense vs NFL Avg")
    _tbl(def_table, "Defense vs NFL Avg")
    doc = SimpleDocTemplate(filename, pagesize=letter)
    doc.build(story)
    return True, "ok"

# ========== Sidebar ==========
with st.sidebar:
    st.header("Tools")
    st.caption("Use league-wide data if you have it; otherwise YTD is computed from uploads.")
    nfl_avgs_file = st.file_uploader("Compute NFL Averages (Manual/Optional): upload NFL averages CSV", type=["csv"], key="nfl_avg_upload")
    if nfl_avgs_file:
        try:
            nfl_avgs_df = pd.read_csv(nfl_avgs_file)
            _append_df("NFL_Averages_Manual", nfl_avgs_df, None)
            st.success("Manual NFL averages uploaded.")
        except Exception as e:
            st.error(f"Failed to load NFL averages: {e}")

    if st.button("Download All Data (Excel)"):
        _ensure_excel()
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("Click to Download", f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ========== Weekly Controls ==========
st.subheader("Weekly Controls")
cols_wc = st.columns([1,1,2,3])
with cols_wc[0]:
    selected_week = st.number_input("Week", min_value=1, step=1, value=1)
with cols_wc[1]:
    selected_team = st.text_input("Team", value="CHI")
with cols_wc[2]:
    opponent = st.text_input("Opponent", value="W01 TBD")
with cols_wc[3]:
    key_notes = st.text_area("Key Notes (appear in PDF)", placeholder="Short bullets for the weekly PDF…")

# ========== 1) Upload Weekly Data ==========
st.markdown("### 1) Upload Weekly Data")
cols_up = st.columns(4)
with cols_up[0]:
    off_file = st.file_uploader("Offense CSV", type=["csv"], key="off_csv")
with cols_up[1]:
    def_file = st.file_uploader("Defense CSV", type=["csv"], key="def_csv")
with cols_up[2]:
    pers_file = st.file_uploader("Personnel CSV", type=["csv"], key="per_csv")
with cols_up[3]:
    snaps_file = st.file_uploader("Snap Counts CSV", type=["csv"], key="snap_csv")

def _load_csv_to_df(f, week: int):
    if not f: return None
    try:
        df = pd.read_csv(f)
        df = _coerce_week_col(df)
        if "Week" not in df.columns:
            df.insert(0, "Week", week)
        else:
            df["Week"] = week
        return df
    except Exception as e:
        st.error(f"CSV load error: {e}")
        return None

if st.button("Save Weekly Uploads"):
    any_ok = False
    for name, f in [("Offense", off_file), ("Defense", def_file), ("Personnel", pers_file), ("SnapCounts", snaps_file)]:
        df = _load_csv_to_df(f, selected_week)
        if df is not None and not df.empty:
            _append_df(name, df, key_cols=["Week"])
            any_ok = True
    if any_ok:
        st.success("Weekly uploads saved.")
    else:
        st.info("No files uploaded.")

# Compute proxy (demo) for selected week
off_df_all = _read_sheet("Offense")
def_df_all = _read_sheet("Defense")
off_row = off_df_all[off_df_all["Week"]==selected_week].tail(1).squeeze() if not off_df_all.empty and "Week" in off_df_all else None
def_row = def_df_all[def_df_all["Week"]==selected_week].tail(1).squeeze() if not def_df_all.empty and "Week" in def_df_all else None
proxy_val = _calc_proxy(off_row, def_row)

st.markdown("#### DVOA-like Proxy (auto)")
st.write(f"**Week {selected_week} proxy:** {proxy_val if proxy_val is not None else '—'}")

# ========== Media Summaries ==========
st.markdown("### Media Summaries (Store multiple per week)")
ms_cols = st.columns([2,6,2])
with ms_cols[0]:
    ms_source = st.text_input("Source (e.g., ESPN, The Athletic)")
with ms_cols[1]:
    ms_text = st.text_area("Summary")
with ms_cols[2]:
    if st.button("Save Summary"):
        if ms_text.strip():
            _append_df("MediaSummaries", pd.DataFrame([{"Week": selected_week, "Source": ms_source, "Summary": ms_text.strip()}]), key_cols=["Week","Source","Summary"])
            st.success("Saved summary.")
        else:
            st.info("Nothing to save.")
ms_list = _read_sheet("MediaSummaries")
if not ms_list.empty:
    st.dataframe(ms_list[ms_list["Week"]==selected_week], use_container_width=True, hide_index=True)

# ========== Injuries ==========
st.markdown("### Injuries – Weekly Rows")
inj_df = _read_sheet("Injuries")
inj_cols = st.columns([2,2,6,2])
with inj_cols[0]:
    inj_name = st.text_input("Injury")
with inj_cols[1]:
    inj_status = st.selectbox("Status", ["Questionable","Doubtful","Out","IR","Active"])
with inj_cols[2]:
    inj_notes = st.text_input("Notes")
with inj_cols[3]:
    if st.button("Add Injury Row"):
        _append_df("Injuries", pd.DataFrame([{"Week": selected_week, "Injury": inj_name, "Status": inj_status, "Notes": inj_notes}]), None)
        st.success("Injury row added.")
inj_preview = inj_df[inj_df["Week"]==selected_week] if not inj_df.empty and "Week" in inj_df else pd.DataFrame()
st.dataframe(inj_preview, use_container_width=True, hide_index=True)

# ========== Opponent Preview ==========
st.markdown("### Opponent Preview")
opp_file = st.file_uploader("Upload opponent scouting/notes CSV (optional)", type=["csv"], key="opp_csv")
if st.button("Save Opponent Preview"):
    df = _load_csv_to_df(opp_file, selected_week)
    if df is not None and not df.empty:
        # If CSV has many cols, squash to a Notes column for simplicity
        try:
            text_block = df.to_csv(index=False)
            _append_df("OpponentPreview", pd.DataFrame([{"Week": selected_week, "Notes": text_block}]), key_cols=["Week"])
            st.success("Opponent preview saved.")
        except Exception as e:
            st.error(f"Opponent preview save failed: {e}")
    else:
        st.info("No opponent file provided.")

# ========== This Week – Data Previews ==========
st.markdown("### This Week – Data Previews")
def _week_preview(name):
    df = _read_sheet(name)
    if df is None or df.empty or "Week" not in df.columns:
        st.caption(f"{name}: no data yet")
    else:
        st.caption(f"{name}:")
        st.dataframe(df[df["Week"]==selected_week], use_container_width=True, hide_index=True)

cols_prev = st.columns(5)
with cols_prev[0]: _week_preview("Offense")
with cols_prev[1]: _week_preview("Defense")
with cols_prev[2]: _week_preview("Personnel")
with cols_prev[3]: _week_preview("Injuries")
with cols_prev[4]: _week_preview("SnapCounts")

# ========== Weekly Game Predictions ==========
st.markdown("### Weekly Game Predictions (Demo)")
pred_cols = st.columns([6,2,2])
with pred_cols[0]:
    pred_rationale = st.text_area("Prediction Rationale (uses strategy + proxy + injuries + opponent preview, etc.)")
with pred_cols[1]:
    pred_outcome = st.selectbox("Predicted Outcome", ["", "Win", "Loss"])
with pred_cols[2]:
    if st.button("Save Prediction"):
        if pred_outcome:
            _append_df("Predictions", pd.DataFrame([{"Week": selected_week, "Rationale": pred_rationale, "Prediction": pred_outcome}]), key_cols=["Week"])
            st.success("Prediction saved.")
        else:
            st.info("Choose an outcome.")

pred_df = _read_sheet("Predictions")
if not pred_df.empty:
    st.dataframe(pred_df.sort_values("Week"), use_container_width=True, hide_index=True)

# ========== 6) YTD Summary (Auto) ==========
st.markdown("### 6) YTD Summary (Auto-computed from uploads)")

def _ytd_team(df: pd.DataFrame, up_to_week: int) -> pd.DataFrame:
    if df is None or df.empty or "Week" not in df.columns:
        return pd.DataFrame()
    return df[df["Week"].between(1, up_to_week, inclusive="both")]

def _ytd_nfl(df: pd.DataFrame, up_to_week: int) -> pd.DataFrame:
    # In a full version, we'd average by team; here we average the uploaded rows as league approximation.
    return _ytd_team(df, up_to_week)

off_ytd_team = _ytd_team(off_df_all, selected_week)
off_ytd_nfl  = _ytd_nfl(off_df_all, selected_week)
def_ytd_team = _ytd_team(def_df_all, selected_week)
def_ytd_nfl  = _ytd_nfl(def_df_all, selected_week)

def _team_vs_nfl_display(team_df, nfl_df, heading):
    if (team_df is None or team_df.empty) and (nfl_df is None or nfl_df.empty):
        st.info(f"{heading}: will appear after you upload offense/defense CSVs with expected headers (e.g., {', '.join(METRIC_SCHEMA.keys())}).")
        return None
    merged = _merge_team_vs_nfl(team_df, nfl_df)
    st.markdown(f"**{heading}**")
    st.dataframe(_colorize_df(merged), use_container_width=True)
    return merged

off_disp = _team_vs_nfl_display(off_ytd_team, off_ytd_nfl,
                                f"{selected_team} Offense YTD vs NFL Avg (W1–W{selected_week})")
def_disp = _team_vs_nfl_display(def_ytd_team, def_ytd_nfl,
                                f"{selected_team} Defense YTD vs NFL Avg (W1–W{selected_week})")

# Write YTD sheets (for Excel downloads)
def _write_ytd_sheets():
    _ensure_excel()
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        # team-only YTD numeric
        off_team_num = off_ytd_team.select_dtypes(include="number")
        def_team_num = def_ytd_team.select_dtypes(include="number")
        off_nfl_num  = off_ytd_nfl.select_dtypes(include="number")
        def_nfl_num  = def_ytd_nfl.select_dtypes(include="number")

        off_team_num.to_excel(writer, sheet_name="YTD_Team_Offense", index=False)
        off_nfl_num.to_excel(writer,  sheet_name="YTD_NFL_Offense",  index=False)
        def_team_num.to_excel(writer, sheet_name="YTD_Team_Defense", index=False)
        def_nfl_num.to_excel(writer,  sheet_name="YTD_NFL_Defense",  index=False)
_write_ytd_sheets()

# ========== Exports: Pre / Post / Final ==========
st.markdown("### Exports")
exp_cols = st.columns(6)
def _export_excel(tag: str):
    _ensure_excel()
    # apply simple conditional formatting to YTD Team Offense sheet
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        for sheet_name in ["YTD_Team_Offense","YTD_Team_Defense"]:
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                # Find columns
                headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
                if not headers: continue
                first_data_row = 2
                last_row = ws.max_row
                if "YPA" in headers:
                    team_col_letter = openpyxl.utils.get_column_letter(headers.index("YPA")+1)
                    _apply_excel_conditional_formatting(ws, first_data_row, team_col_letter, team_col_letter, last_row)
        out_name = f"W{int(selected_week):02d}_{tag}.xlsx"
        wb.save(out_name)
        with open(out_name,"rb") as f:
            st.download_button(f"Download {out_name}", f, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        os.remove(out_name)
        st.success(f"{out_name} created.")
    except Exception as e:
        st.error(f"Excel export failed: {e}")

def _export_pdf(tag: str):
    if not HAS_REPORTLAB:
        st.warning("PDF export needs 'reportlab' installed and detected. Add to requirements.txt and redeploy.")
        return
    try:
        out_name = f"W{int(selected_week):02d}_{tag}.pdf"
        team_title = f"Week {selected_week}: {selected_team} vs {opponent}"
        _ = _export_pdf_weekly(
            out_name,
            team_title,
            key_notes,
            off_table=off_disp,
            def_table=def_disp
        )
        with open(out_name,"rb") as f:
            st.download_button(f"Download {out_name}", f, file_name=out_name, mime="application/pdf")
        os.remove(out_name)
        st.success(f"{out_name} created.")
    except Exception as e:
        st.error(f"PDF export failed: {e}")

with exp_cols[0]:
    if st.button("Export Pre (Excel)"):
        _export_excel("Pre")
with exp_cols[1]:
    if st.button("Export Pre (PDF)"):
        _export_pdf("Pre")
with exp_cols[2]:
    if st.button("Export Post (Excel)"):
        _export_excel("Post")
with exp_cols[3]:
    if st.button("Export Post (PDF)"):
        _export_pdf("Post")
with exp_cols[4]:
    if st.button("Export Final (Excel)"):
        _export_excel("Final")
with exp_cols[5]:
    if st.button("Export Final (PDF)"):
        _export_pdf("Final")

st.caption("Tip: If PDF export says reportlab is missing, add `reportlab` to requirements.txt in the repo root, commit, push, then reboot/rerun on Streamlit Cloud.")

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
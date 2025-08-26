# bears_dashboard.py
# Chicago Bears 2025–26 Weekly Tracker (Streamlit)

import os
import pandas as pd
import streamlit as st

# Optional extras
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

# ========= Config =========
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

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

# ========= Excel helpers =========
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

def _read_sheet(name: str) -> pd.DataFrame:
    _ensure_excel()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=name)
    except Exception:
        return pd.DataFrame()

# --- Backward-compat shim ---
SHEETS = {
    "Offense": "Offense",
    "Defense": "Defense",
    "Personnel": "Personnel",
    "SnapCounts": "SnapCounts",
    "Injuries": "Injuries",
    "MediaSummaries": "MediaSummaries",
    "OpponentPreview": "OpponentPreview",
    "Predictions": "Predictions",
}
def _load_sheet(name: str) -> pd.DataFrame:
    return _read_sheet(name)
# --- End shim ---

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
    if key_cols:
        df_out = df_out.drop_duplicates(subset=key_cols, keep="last")
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_out.to_excel(writer, sheet_name=sheet_name, index=False)

# ========= Proxy calc =========
def _calc_proxy(off_row: pd.Series=None, def_row: pd.Series=None) -> float | None:
    if off_row is None and def_row is None:
        return None
    score = 0.0
    count = 0
    def val(s, k):
        try:
            return float(s.get(k)) if s is not None and k in s else None
        except Exception:
            return None
    if off_row is not None:
        ypa = val(off_row, "YPA")
        cmp_ = val(off_row, "CMP%")
        if ypa is not None: score += ypa; count += 1
        if cmp_ is not None: score += (cmp_/100); count += 1
    if def_row is not None:
        rz = val(def_row, "RZ% Allowed")
        sacks = val(def_row, "SACKs")
        if rz is not None: score += (1-(rz/100)); count += 1
        if sacks is not None: score += min(1.0, sacks/5.0); count += 1
    return round(score/count, 4) if count else None

# ========= Colorize =========
def _colorize_df(df: pd.DataFrame):
    """Color Team vs NFL columns based on better/worse per METRIC_SCHEMA."""
    if df is None or getattr(df, "empty", True):
        return df
    if not hasattr(df, "style"):
        return df
    def _color_row(row):
        bg = []
        for col in df.columns:
            if col == "Metric":
                bg.append("")
                continue
            metric = row.get("Metric")
            info = METRIC_SCHEMA.get(metric, {"higher_is_better": True})
            higher = info["higher_is_better"]
            team = row.get("Team")
            nfl = row.get("NFL")
            if col == "Team" and pd.notnull(team) and pd.notnull(nfl):
                better = (team > nfl) if higher else (team < nfl)
                color = "background-color:#d4edda" if better else "background-color:#f8d7da"
                bg.append(color)
            else:
                bg.append("")
        return bg
    return df.style.apply(_color_row, axis=1)

# ========= Sidebar =========
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
            st.error(f"Failed: {e}")
    if st.button("Download All Data (Excel)"):
        _ensure_excel()
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("Click to Download", f, file_name=EXCEL_FILE)

# ========= Weekly Controls =========
st.subheader("Weekly Controls")
cols_wc = st.columns([1,1,2,3])
with cols_wc[0]:
    selected_week = st.number_input("Week", min_value=1, step=1, value=1)
with cols_wc[1]:
    selected_team = st.text_input("Team", value="CHI")
with cols_wc[2]:
    opponent = st.text_input("Opponent", value="TBD")
with cols_wc[3]:
    key_notes = st.text_area("Key Notes (appear in PDF)")

# ========= Uploads (Off/Def/etc.) =========
st.markdown("### Upload Weekly Data")
cols_up = st.columns(4)
with cols_up[0]:
    off_file = st.file_uploader("Offense CSV", type=["csv"])
with cols_up[1]:
    def_file = st.file_uploader("Defense CSV", type=["csv"])
with cols_up[2]:
    pers_file = st.file_uploader("Personnel CSV", type=["csv"])
with cols_up[3]:
    snaps_file = st.file_uploader("Snap Counts CSV", type=["csv"])

def _load_csv_to_df(f, week: int):
    if not f:
        return None
    try:
        df = pd.read_csv(f)
        if "Week" not in df.columns:
            df.insert(0, "Week", week)
        else:
            df["Week"] = week
        return df
    except Exception as e:
        st.error(f"CSV load error: {e}")
        return None

if st.button("Save Weekly Uploads"):
    for name, f in [("Offense", off_file), ("Defense", def_file), ("Personnel", pers_file), ("SnapCounts", snaps_file)]:
        df = _load_csv_to_df(f, selected_week)
        if df is not None:
            _append_df(name, df, ["Week"])
    st.success("Weekly uploads saved.")

# ========= Proxy =========
off_df_all = _read_sheet("Offense")
def_df_all = _read_sheet("Defense")
off_row = off_df_all[off_df_all["Week"] == selected_week].tail(1).squeeze() if not off_df_all.empty else None
def_row = def_df_all[def_df_all["Week"] == selected_week].tail(1).squeeze() if not def_df_all.empty else None
proxy_val = _calc_proxy(off_row, def_row)
st.markdown("#### DVOA-like Proxy")
st.write(f"Week {selected_week} proxy: {proxy_val if proxy_val is not None else '—'}")

# ========= Media Summaries =========
st.markdown("### Media Summaries (Store multiple per week)")
ms_cols = st.columns([2, 6, 2])
with ms_cols[0]:
    ms_source = st.text_input("Source (e.g., ESPN, The Athletic)")
with ms_cols[1]:
    ms_text = st.text_area("Summary", height=130)
with ms_cols[2]:
    if st.button("Save Summary"):
        if ms_text.strip():
            _append_df(
                "MediaSummaries",
                pd.DataFrame([{
                    "Week": selected_week,
                    "Source": ms_source.strip(),
                    "Summary": ms_text.strip()
                }]),
                key_cols=["Week", "Source", "Summary"]
            )
            st.success("Saved media summary.")
        else:
            st.info("Nothing to save.")
_ms_all = _read_sheet("MediaSummaries")
if not _ms_all.empty and "Week" in _ms_all.columns:
    st.dataframe(
        _ms_all[_ms_all["Week"] == selected_week].sort_index(),
        use_container_width=True,
        hide_index=True
    )

# ========= Opponent Preview =========
st.markdown("### Opponent Preview")
opp_cols = st.columns([5, 2])
with opp_cols[0]:
    opp_file = st.file_uploader("Upload opponent scouting/notes CSV (optional)", type=["csv"], key="opp_csv")
    opp_free_text = st.text_area("Or paste scouting notes here (optional)", height=130)
with opp_cols[1]:
    if st.button("Save Opponent Preview"):
        try:
            saved = False
            if opp_file is not None:
                df = pd.read_csv(opp_file)
                text_block = df.to_csv(index=False)
                _append_df(
                    "OpponentPreview",
                    pd.DataFrame([{"Week": selected_week, "Notes": text_block}]),
                    key_cols=["Week"]
                )
                saved = True
            elif opp_free_text.strip():
                _append_df(
                    "OpponentPreview",
                    pd.DataFrame([{"Week": selected_week, "Notes": opp_free_text.strip()}]),
                    key_cols=["Week"]
                )
                saved = True
            if saved:
                st.success("Opponent preview saved.")
            else:
                st.info("No CSV or text to save.")
        except Exception as e:
            st.error(f"Opponent preview save failed: {e}")
_op_all = _read_sheet("OpponentPreview")
if not _op_all.empty and "Week" in _op_all.columns:
    st.caption("Current Week Opponent Preview")
    st.dataframe(
        _op_all[_op_all["Week"] == selected_week],
        use_container_width=True,
        hide_index=True
    )

# ========= Weekly Game Predictions =========
st.markdown("### Weekly Game Predictions")
pred_cols = st.columns([6, 2, 2])
with pred_cols[0]:
    pred_rationale = st.text_area(
        "Prediction Rationale (uses strategy + proxy + injuries + opponent preview, etc.)",
        height=130
    )
with pred_cols[1]:
    pred_outcome = st.selectbox("Predicted Outcome", ["", "Win", "Loss"])
with pred_cols[2]:
    if st.button("Save Prediction"):
        if pred_outcome:
            _append_df(
                "Predictions",
                pd.DataFrame([{
                    "Week": selected_week,
                    "Rationale": pred_rationale.strip(),
                    "Prediction": pred_outcome
                }]),
                key_cols=["Week"]  # keep the latest prediction per week
            )
            st.success("Prediction saved.")
        else:
            st.info("Choose an outcome before saving.")
_pred_all = _read_sheet("Predictions")
if not _pred_all.empty and "Week" in _pred_all.columns:
    st.dataframe(_pred_all.sort_values("Week"), use_container_width=True, hide_index=True)

# ========= YTD Summary =========
st.markdown("### YTD Summary (auto)")
def _ytd(df, week: int):
    return df[df["Week"].between(1, week)] if not df.empty and "Week" in df.columns else pd.DataFrame()
off_ytd = _ytd(off_df_all, selected_week)
def_ytd = _ytd(def_df_all, selected_week)

if off_ytd.empty and def_ytd.empty:
    st.info("Upload Offense/Defense to see YTD.")
else:
    merged = pd.DataFrame({"Metric": list(METRIC_SCHEMA.keys())})
    # For demo: treat Offense as Team and Defense as NFL avg proxy if present.
    # You can change this to proper per-team vs. league if your data has team columns.
    team_means = off_ytd.mean(numeric_only=True) if not off_ytd.empty else pd.Series(dtype=float)
    nfl_means  = def_ytd.mean(numeric_only=True) if not def_ytd.empty else pd.Series(dtype=float)
    merged["Team"] = [team_means.get(m) for m in METRIC_SCHEMA.keys()]
    merged["NFL"]  = [nfl_means.get(m)  for m in METRIC_SCHEMA.keys()]
    st.dataframe(_colorize_df(merged), use_container_width=True)

# ========= Exports (placeholders) =========
st.markdown("### Exports")
if st.button("Export Pre (Excel)"):
    st.info("Would create Pre Excel here.")
if st.button("Export Pre (PDF)"):
    if not HAS_REPORTLAB:
        st.warning("Install reportlab for PDF export")
    else:
        st.info("Would create Pre PDF here.")



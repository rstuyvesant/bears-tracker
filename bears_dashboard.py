import os
import streamlit as st
import pandas as pd
from fpdf import FPDF

# -------------------- App Header --------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, and opponent-adjusted efficiency.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# -------------------- Styling Helpers --------------------
def _fmt_pct(x):
    try:
        return f"{float(x):.1f}%"
    except Exception:
        return x

def _fmt_3(x):
    try:
        return f"{float(x):.3f}"
    except Exception:
        return x

def _color_pos_neg(val):
    """Green for positive, red for negative (offense better when > 0)."""
    try:
        v = float(val)
    except Exception:
        return ""
    if v > 0:
        return "background-color: lightgreen; color: black"
    if v < 0:
        return "background-color: salmon; color: black"
    return ""

def _color_neg_pos(val):
    """Green for negative, red for positive (defense better when < 0)."""
    try:
        v = float(val)
    except Exception:
        return ""
    if v < 0:
        return "background-color: lightgreen; color: black"
    if v > 0:
        return "background-color: salmon; color: black"
    return ""

def _color_sr_higher_better(val):
    """Success Rate where higher is good (e.g., offense SR%)."""
    try:
        v = float(val)
    except Exception:
        return ""
    if v >= 55:
        return "background-color: lightgreen; color: black"
    if v <= 45:
        return "background-color: salmon; color: black"
    return ""

def _color_sr_lower_better(val):
    """Success Rate where lower is good (e.g., defense allowed SR%)."""
    try:
        v = float(val)
    except Exception:
        return ""
    if v <= 45:
        return "background-color: lightgreen; color: black"
    if v >= 55:
        return "background-color: salmon; color: black"
    return ""

# -------------------- Excel Helpers --------------------
def append_to_excel(new_data, sheet_name, file_name=EXCEL_FILE, deduplicate=True, key_cols=("Week",)):
    """Append/replace a sheet with optional dedup by key columns."""
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    try:
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                # Read existing sheet
                ws = book[sheet_name]
                existing = pd.DataFrame(ws.values)
                if len(existing) > 0 and existing.iloc[0].notna().any():
                    existing.columns = existing.iloc[0]
                    existing = existing[1:]
                else:
                    existing = pd.DataFrame(columns=new_data.columns)

                # Deduplicate by key
                if deduplicate and all(k in existing.columns for k in key_cols) and all(k in new_data.columns for k in key_cols):
                    # Drop rows whose keys are present in new_data
                    merge_keys = list(key_cols)
                    key_tuples = set(tuple(map(str, r)) for r in new_data[merge_keys].astype(str).values)
                    mask = existing[merge_keys].astype(str).apply(tuple, axis=1).apply(lambda t: t not in key_tuples)
                    existing = existing[mask]

                combined = pd.concat([existing, new_data], ignore_index=True)
            else:
                combined = new_data
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            combined = new_data

        if sheet_name in book.sheetnames:
            del book[sheet_name]
        ws_new = book.create_sheet(sheet_name)
        for r in dataframe_to_rows(combined, index=False, header=True):
            ws_new.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def ensure_sheets():
    """Make sure workbook exists with expected empty sheets."""
    base = {
        "Offense": ["Week", "YDS", "YPA", "YPC", "CMP%"],
        "Defense": ["Week", "SACK", "INT", "FF", "FR", "RZ% Allowed"],
        "Strategy": ["Week", "Opponent", "Off_Strategy", "Off_Result", "Def_Strategy", "Def_Result", "Key_Notes", "Next_Week_Impact"],
        "Personnel": ["Week", "11", "12", "13", "21", "Division 11", "Division 12", "Division 13", "Division 21",
                      "Conf 11", "Conf 12", "Conf 13", "Conf 21", "NFL 11", "NFL 12", "NFL 13", "NFL 21"],
        "Media_Summaries": ["Week", "Opponent", "Summary"],
        "Advanced_Defense": ["Week", "RZ% Allowed", "Success Rate% (Offense)", "Pressures"],
        "DVOA_Proxy": ["Week", "Opponent", "Off Adj EPA/play", "Off Adj SR%", "Def Adj EPA/play", "Def Adj SR%",
                       "Off EPA/play", "Def EPA allowed/play"],
        "Predictions": ["Week", "Prediction", "Reason", "Notes"],
        "Injuries": ["Week", "Player", "Pos", "Status", "BodyPart", "Practice", "GameStatus", "Notes"],
        "SnapCounts": ["Week", "Unit", "Player", "Pos", "Snaps", "Snap%"]
    }
    for sheet, cols in base.items():
        try:
            if os.path.exists(EXCEL_FILE):
                x = pd.ExcelFile(EXCEL_FILE)
                if sheet in x.sheet_names:
                    continue
            df = pd.DataFrame(columns=cols)
            append_to_excel(df, sheet, deduplicate=False)
        except Exception as e:
            st.error(f"Ensure sheet failed for {sheet}: {e}")

ensure_sheets()

# -------------------- Sidebar: Uploaders --------------------
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")
uploaded_injuries = st.sidebar.file_uploader("Upload Injuries (.csv)", type="csv")
uploaded_snaps = st.sidebar.file_uploader("Upload Snap Counts (.csv)", type="csv")

if uploaded_offense is not None:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense")
    st.sidebar.success("✅ Offensive data uploaded.")

if uploaded_defense is not None:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense")
    st.sidebar.success("✅ Defensive data uploaded.")

if uploaded_strategy is not None:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy")
    st.sidebar.success("✅ Strategy data uploaded.")

if uploaded_personnel is not None:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel")
    st.sidebar.success("✅ Personnel data uploaded.")

if uploaded_injuries is not None:
    df_inj = pd.read_csv(uploaded_injuries)
    append_to_excel(df_inj, "Injuries", key_cols=("Week", "Player"))
    st.sidebar.success("✅ Injuries uploaded (dedup by Week+Player).")

if uploaded_snaps is not None:
    df_snaps = pd.read_csv(uploaded_snaps)
    append_to_excel(df_snaps, "SnapCounts", key_cols=("Week", "Unit", "Player"))
    st.sidebar.success("✅ Snap counts uploaded (dedup by Week+Unit+Player).")

# -------------------- Sidebar: Fetch (nfl_data_py) --------------------
with st.sidebar.expander("⚡ Fetch Weekly Team Data (nfl_data_py)"):
    st.caption("Pulls 2025 weekly team stats for CHI and saves to Excel (basic fields).")
    fetch_week = st.number_input("Week to fetch (2025)", min_value=1, max_value=25, value=1, step=1, key="fetch_week_2025")

    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl
            try:
                nfl.update.weekly_data([2025])
            except Exception:
                pass
            weekly = nfl.import_weekly_data([2025])
            wk = int(fetch_week)
            team_week = weekly[(weekly["team"] == "CHI") & (weekly["week"] == wk)].copy()

            if team_week.empty:
                st.warning("No weekly row found for CHI in that week yet.")
            else:
                # Offense (best-effort)
                pass_yards = team_week["passing_yards"].iloc[0] if "passing_yards" in team_week.columns else None
                pass_att = None
                for cand in ["attempts", "passing_attempts", "pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].iloc[0]
                        break
                ypa_val = None
                if pass_yards is not None and pass_att not in (None, 0):
                    try:
                        ypa_val = float(pass_yards) / float(pass_att)
                    except Exception:
                        ypa_val = None

                yards_total = None
                for cand in ["yards", "total_yards", "offense_yards"]:
                    if cand in team_week.columns:
                        yards_total = team_week[cand].iloc[0]
                        break

                completions = None
                for cand in ["completions", "passing_completions", "pass_completions"]:
                    if cand in team_week.columns:
                        completions = team_week[cand].iloc[0]
                        break
                cmp_pct = None
                if completions is not None and pass_att not in (None, 0):
                    try:
                        cmp_pct = round((float(completions) / float(pass_att)) * 100, 1)
                    except Exception:
                        cmp_pct = None

                off_row = pd.DataFrame([{
                    "Week": wk,
                    "YDS": yards_total,
                    "YPA": round(ypa_val, 2) if ypa_val is not None else None,
                    "CMP%": cmp_pct
                }])

                # Defense
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].iloc[0]
                        break
                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val,
                    "RZ% Allowed": None  # filled by PBP fetch
                }])

                append_to_excel(off_row, "Offense", key_cols=("Week",))
                append_to_excel(def_row, "Defense", key_cols=("Week",))
                st.success(f"✅ Added CHI week {wk} to Offense/Defense (available fields).")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📡 Fetch Defensive Metrics (Play-by-Play)")
pbp_week = st.sidebar.number_input("Week to fetch (2025)", min_value=1, max_value=25, value=1, step=1, key="pbp_week_2025")

if st.sidebar.button("Fetch PBP Metrics"):
    try:
        import nfl_data_py as nfl
        pbp = nfl.import_pbp_data([2025], downcast=False)
        df = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == "CHI")].copy()

        if df.empty:
            st.warning("No PBP rows for CHI defense in that week yet.")
        else:
            # Red zone drives allowed (% of opponent drives penetrating the 20)
            dmins = (
                df.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success Rate (offense success vs CHI defense) — exclude no_play & penalties
            def play_success(r):
                try:
                    d = int(r["down"])
                    togo = float(r["ydstogo"])
                    gain = float(r["yards_gained"])
                except Exception:
                    return False
                if d == 1:
                    return gain >= 0.4 * togo
                elif d == 2:
                    return gain >= 0.6 * togo
                else:
                    return gain >= togo

            mask = (~df["play_type"].isin(["no_play"])) & (~df["penalty"].fillna(False))
            real = df[mask].copy()
            success_rate = real.apply(play_success, axis=1).mean() * 100 if len(real) else 0.0

            qb_hits = real["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in real.columns else 0
            sacks = real["sack"].fillna(0).astype(int).sum() if "sack" in real.columns else 0
            pressures = int(qb_hits + sacks)

            metrics = pd.DataFrame([{
                "Week": int(pbp_week),
                "RZ% Allowed": round(rz_allowed, 1),
                "Success Rate% (Offense)": round(success_rate, 1),
                "Pressures": pressures
            }])
            append_to_excel(metrics, "Advanced_Defense", key_cols=("Week",))
            st.success(f"✅ Week {int(pbp_week)} PBP metrics saved.")
    except Exception as e:
        st.error(f"❌ Failed to fetch PBP metrics: {e}")

# -------------------- Opponent-adjusted DVOA-like Proxy --------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.sidebar.number_input("Week to compute (2025)", min_value=1, max_value=25, value=1, step=1, key="proxy_week_2025")

def _success_flag(down, ydstogo, yards_gained):
    try:
        if pd.isna(down) or pd.isna(ydstogo) or pd.isna(yards_gained):
            return False
        d = int(down); togo = float(ydstogo); gain = float(yards_gained)
        if d == 1:
            return gain >= 0.4 * togo
        elif d == 2:
            return gain >= 0.6 * togo
        else:
            return gain >= togo
    except Exception:
        return False

if st.sidebar.button("Compute DVOA-like Proxy"):
    try:
        import nfl_data_py as nfl
        wk = int(proxy_week)
        pbp = nfl.import_pbp_data([2025], downcast=False)
        plays = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False)) &
            (~pbp["epa"].isna())
        ].copy()

        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == "CHI")].copy()
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == "CHI")].copy()

        if bears_off.empty and bears_def.empty:
            st.warning("No CHI plays found for that week yet.")
        else:
            opps = set()
            if not bears_off.empty:
                opps.update(bears_off["defteam"].unique().tolist())
            if not bears_def.empty:
                opps.update(bears_def["posteam"].unique().tolist())
            opponent = list(opps)[0] if opps else "UNK"

            prior = plays[plays["week"] < wk].copy()
            opp_def_plays = prior[prior["defteam"] == opponent].copy()
            opp_def_epa = opp_def_plays["epa"].mean() if len(opp_def_plays) else None
            opp_def_success = opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_def_plays) else None

            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            opp_off_success = opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_off_plays) else None

            if len(bears_off):
                chi_off_epa = bears_off["epa"].mean()
                chi_off_success = bears_off.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_off_epa = None; chi_off_success = None

            if len(bears_def):
                chi_def_epa_allowed = bears_def["epa"].mean()
                chi_def_success_allowed = bears_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_def_epa_allowed = None; chi_def_success_allowed = None

            def safe_diff(a, b):
                if a is None or pd.isna(a) or b is None or pd.isna(b):
                    return None
                return float(a) - float(b)

            off_adj_epa = safe_diff(chi_off_epa, opp_def_epa)
            off_adj_sr  = safe_diff(chi_off_success, opp_def_success)
            def_adj_epa = safe_diff(chi_def_epa_allowed, opp_off_epa)
            def_adj_sr  = safe_diff(chi_def_success_allowed, opp_off_success)

            out = pd.DataFrame([{
                "Week": wk,
                "Opponent": opponent,
                "Off Adj EPA/play": round(off_adj_epa, 3) if off_adj_epa is not None else None,
                "Off Adj SR%": round(off_adj_sr * 100, 1) if off_adj_sr is not None else None,
                "Def Adj EPA/play": round(def_adj_epa, 3) if def_adj_epa is not None else None,
                "Def Adj SR%": round(def_adj_sr * 100, 1) if def_adj_sr is not None else None,
                "Off EPA/play": round(chi_off_epa, 3) if chi_off_epa is not None else None,
                "Def EPA allowed/play": round(chi_def_epa_allowed, 3) if chi_def_epa_allowed is not None else None
            }])

            append_to_excel(out, "DVOA_Proxy", key_cols=("Week",))
            st.success(f"✅ DVOA-like proxy saved for Week {wk} vs {opponent}.")
    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# -------------------- Download Excel --------------------
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------- Quick Injury Entry --------------------
st.markdown("### 🏥 Injuries (Quick Entry)")
with st.form("injury_quick"):
    iq_week = st.number_input("Week", min_value=1, max_value=25, step=1, key="injury_week")
    iq_player = st.text_input("Player")
    iq_pos = st.text_input("Position")
    iq_status = st.selectbox("Status", ["Questionable", "Doubtful", "Out", "IR", "PUP", "Probable", "Healthy"], index=0)
    iq_body = st.text_input("Body Part")
    iq_practice = st.selectbox("Practice", ["DNP", "Limited", "Full", "N/A"], index=3)
    iq_game = st.selectbox("Game Status", ["TBD", "Active", "Inactive"], index=0)
    iq_notes = st.text_area("Notes")
    iq_submit = st.form_submit_button("Save Injury")
if iq_submit:
    inj_row = pd.DataFrame([{
        "Week": iq_week, "Player": iq_player, "Pos": iq_pos, "Status": iq_status,
        "BodyPart": iq_body, "Practice": iq_practice, "GameStatus": iq_game, "Notes": iq_notes
    }])
    append_to_excel(inj_row, "Injuries", key_cols=("Week", "Player"))
    st.success(f"✅ Saved injury for Week {iq_week}: {iq_player}")

# -------------------- Create Next Week Templates --------------------
st.markdown("### 🧩 Templates")
tw = st.number_input("Template Week", min_value=1, max_value=25, value=1, step=1, key="tmpl_week")

cols_map = {
    "Offense": ["Week", "YDS", "YPA", "YPC", "CMP%"],
    "Defense": ["Week", "SACK", "INT", "FF", "FR", "RZ% Allowed"],
    "Strategy": ["Week", "Opponent", "Off_Strategy", "Off_Result", "Def_Strategy", "Def_Result", "Key_Notes", "Next_Week_Impact"],
    "Personnel": ["Week", "11", "12", "13", "21", "Division 11", "Division 12", "Division 13", "Division 21",
                  "Conf 11", "Conf 12", "Conf 13", "Conf 21", "NFL 11", "NFL 12", "NFL 13", "NFL 21"]
}

c1, c2, c3 = st.columns(3)
with c1:
    if st.button("➕ Create Next Week Template Row (blank)"):
        for sh, cols in cols_map.items():
            empty = pd.DataFrame([{c: (tw if c == "Week" else None) for c in cols}])
            append_to_excel(empty, sh, key_cols=("Week",))
        st.success(f"Created blank templates for Week {tw}.")

with c2:
    if st.button("➕ Add Next Week Template Rows (append only)"):
        for sh, cols in cols_map.items():
            empty = pd.DataFrame([{c: (tw if c == "Week" else None) for c in cols}])
            append_to_excel(empty, sh, deduplicate=False)
        st.success(f"Appended template rows for Week {tw} (no dedup).")

with c3:
    if st.button("🧹 Remove Empty Template Rows for the Week"):
        try:
            x = pd.ExcelFile(EXCEL_FILE)
            for sh, cols in cols_map.items():
                if sh in x.sheet_names:
                    df = pd.read_excel(EXCEL_FILE, sheet_name=sh)
                    # Keep rows where either Week != tw OR some non-Week cell is not all-null
                    if "Week" in df.columns:
                        mask_keep = ~(
                            (df["Week"] == tw) &
                            (df.drop(columns=["Week"], errors="ignore").isna().all(axis=1))
                        )
                        df2 = df[mask_keep].copy()
                        append_to_excel(df2, sh, deduplicate=False)
            st.success(f"Removed empty templates for Week {tw}.")
        except Exception as e:
            st.error(f"Cleanup failed: {e}")

# -------------------- On-screen Previews --------------------
def _show_sheet(title, sheet):
    if not os.path.exists(EXCEL_FILE):
        return
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
        st.subheader(title)
        st.dataframe(df, use_container_width=True)
    except Exception:
        pass

_show_sheet("📊 Offensive Analytics", "Offense")
_show_sheet("🛡️ Defensive Analytics", "Defense")
_show_sheet("🧭 Weekly Strategy", "Strategy")
_show_sheet("👥 Personnel Usage", "Personnel")
_show_sheet("🏥 Injuries", "Injuries")
_show_sheet("🧮 Advanced Defense (PBP-derived)", "Advanced_Defense")

# -------------------- Styled DVOA-like Proxy Preview --------------------
if os.path.exists(EXCEL_FILE):
    try:
        df_dvoa = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        if not df_dvoa.empty:
            view = df_dvoa.copy()

            # Format columns if present
            if "Off Adj SR%" in view.columns:
                view["Off Adj SR%"] = view["Off Adj SR%"].apply(_fmt_pct)
            if "Def Adj SR%" in view.columns:
                view["Def Adj SR%"] = view["Def Adj SR%"].apply(_fmt_pct)
            for col in ["Off Adj EPA/play", "Def Adj EPA/play", "Off EPA/play", "Def EPA allowed/play"]:
                if col in view.columns:
                    view[col] = view[col].apply(_fmt_3)

            st.subheader("📈 DVOA-like Proxy (Opponent-Adjusted) — colored for quick scan")

            # Build styler with column-specific rules
            styler = view.style

            if "Off Adj EPA/play" in view.columns:
                styler = styler.applymap(_color_pos_neg, subset=["Off Adj EPA/play"])
            if "Def Adj EPA/play" in view.columns:
                styler = styler.applymap(_color_neg_pos, subset=["Def Adj EPA/play"])
            if "Off Adj SR%" in view.columns:
                # Strip % to numeric for coloring
                def _strip_pct(x):
                    try:
                        return float(str(x).replace("%", ""))
                    except Exception:
                        return x
                sr_series = view["Off Adj SR%"].map(_strip_pct)
                sr_df = pd.DataFrame({"Off Adj SR%": sr_series})
                styler = styler.applymap(_color_pos_neg, subset=["Off Adj SR%"])  # Off adj SR% positive is good (green)

            if "Def Adj SR%" in view.columns:
                # For defense, negative (better than opp offense) is green
                styler = styler.applymap(_color_neg_pos, subset=["Def Adj SR%"])

            st.dataframe(styler, use_container_width=True, height=340)
        else:
            st.info("No DVOA Proxy data available yet.")
    except Exception as e:
        st.info(f"No DVOA Proxy data available yet. ({e})")

# -------------------- Weekly Prediction --------------------
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        df_offense = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        df_defense = pd.read_excel(EXCEL_FILE, sheet_name="Defense")
        try:
            df_advdef = pd.read_excel(EXCEL_FILE, sheet_name="Advanced_Defense")
        except Exception:
            df_advdef = pd.DataFrame()
        try:
            df_proxy = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        except Exception:
            df_proxy = pd.DataFrame()

        row_s = df_strategy[df_strategy["Week"] == week_to_predict]
        row_o = df_offense[df_offense["Week"] == week_to_predict]
        row_d = df_defense[df_defense["Week"] == week_to_predict]
        row_a = df_advdef[df_advdef["Week"] == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"] == week_to_predict] if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            rz_allowed = None
            pressures = None
            if not row_a.empty:
                rz_allowed = _safe_float(row_a.iloc[0].get("RZ% Allowed"), default=None)
                pressures = _safe_float(row_a.iloc[0].get("Pressures"), default=None)
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), default=None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), default=None)

            reason_bits = []

            # Strong two-way efficiency edge
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win - efficiency edge on both sides"
                reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def {def_adj_epa:+.2f} EPA/play vs opp O")

            # Pass-rush advantage
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win - pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            # Coverage + red zone discipline
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win - red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            # Clear offensive/defensive drag
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient offense and poor red zone defense"
                reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            # Legacy fallback using YPA + RZ
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient passing and weak red zone defense"
                reason_bits.append(f"YPA={ypa:.1f}")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            else:
                prediction = "Loss - no clear advantage in key strategy or stats"
                if off_adj_epa is not None:
                    reason_bits.append(f"Off {off_adj_epa:+.2f} EPA/play")
                if def_adj_epa is not None:
                    reason_bits.append(f"Def {def_adj_epa:+.2f} EPA/play")
                if pressures is not None:
                    reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            reason_text = " | ".join(reason_bits)
            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")
            if reason_text:
                st.caption(reason_text)

            prediction_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("-")[0].strip(),
                "Reason": prediction.split("-")[1].strip() if "-" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(prediction_entry, "Predictions", key_cols=("Week",))
        else:
            st.info("Please upload or fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check data. Error: {e}")

# -------------------- Saved Predictions Preview --------------------
_show_sheet("📈 Saved Game Predictions", "Predictions")
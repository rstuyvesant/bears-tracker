import os
import streamlit as st
import pandas as pd
from fpdf import FPDF

# ------------------- App Header -------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# ------------------- Helpers -------------------
def ensure_workbook_structure():
    """
    Create workbook and expected sheets if missing.
    """
    import openpyxl
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        # Drop default "Sheet"
        wb.remove(wb.active)
        # Create empty dataframes for known sheets
        empty_sheets = {
            "Offense": ["Week", "Opponent", "YDS", "YPA", "YPC", "CMP%", "QBR",
                        "SR%", "DIV Avg YDS", "DIV Avg QBR", "DIV Avg SR%",
                        "CONF Avg YDS", "CONF Avg QBR", "CONF Avg SR%",
                        "NFL Avg YDS", "NFL Avg QBR", "NFL Avg SR%"],
            "Defense": ["Week", "SACK", "INT", "FF", "FR", "DVOA",
                        "3D% Allowed", "RZ% Allowed", "QB Hits", "Pressures"],
            "Strategy": ["Week", "Opponent", "Off_Strategy", "Off_Results",
                         "Def_Strategy", "Def_Results", "Key_Notes",
                         "Next_Week_Impact"],
            "Personnel": ["Week", "11 Personnel", "12 Personnel", "13 Personnel", "21 Personnel",
                          "Division 11", "Division 12", "Division 13", "Division 21",
                          "Conf 11", "Conf 12", "Conf 13", "Conf 21",
                          "NFL 11", "NFL 12", "NFL 13", "NFL 21"],
            "Injuries": ["Week", "Player", "Position", "Status", "Body_Part", "Practice", "Game_Status", "Notes"],
            "Advanced_Defense": ["Week", "RZ% Allowed", "Success Rate% (Offense)", "Pressures"],
            "DVOA_Proxy": ["Week", "Opponent", "Off Adj EPA/play", "Off Adj SR%",
                           "Def Adj EPA/play", "Def Adj SR%",
                           "Off EPA/play", "Def EPA allowed/play"],
            "Predictions": ["Week", "Prediction", "Reason", "Notes"],
            "Media_Summaries": ["Week", "Opponent", "Summary"],
            "Raw_Weekly": [],  # dump from nfl_data_py import_weekly_data
        }
        for name, cols in empty_sheets.items():
            ws = wb.create_sheet(name)
            if cols:
                ws.append(cols)
        wb.save(EXCEL_FILE)
    else:
        # Make sure all expected sheets exist
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_FILE)
        needed = ["Offense","Defense","Strategy","Personnel","Injuries",
                  "Advanced_Defense","DVOA_Proxy","Predictions","Media_Summaries","Raw_Weekly"]
        for name in needed:
            if name not in wb.sheetnames:
                ws = wb.create_sheet(name)
                # Add headers if known
                if name == "Offense":
                    ws.append(["Week","Opponent","YDS","YPA","YPC","CMP%","QBR",
                               "SR%","DIV Avg YDS","DIV Avg QBR","DIV Avg SR%",
                               "CONF Avg YDS","CONF Avg QBR","CONF Avg SR%",
                               "NFL Avg YDS","NFL Avg QBR","NFL Avg SR%"])
                elif name == "Defense":
                    ws.append(["Week","SACK","INT","FF","FR","DVOA",
                               "3D% Allowed","RZ% Allowed","QB Hits","Pressures"])
                elif name == "Strategy":
                    ws.append(["Week","Opponent","Off_Strategy","Off_Results",
                               "Def_Strategy","Def_Results","Key_Notes","Next_Week_Impact"])
                elif name == "Personnel":
                    ws.append(["Week","11 Personnel","12 Personnel","13 Personnel","21 Personnel",
                               "Division 11","Division 12","Division 13","Division 21",
                               "Conf 11","Conf 12","Conf 13","Conf 21",
                               "NFL 11","NFL 12","NFL 13","NFL 21"])
                elif name == "Injuries":
                    ws.append(["Week","Player","Position","Status","Body_Part","Practice","Game_Status","Notes"])
                elif name == "Advanced_Defense":
                    ws.append(["Week","RZ% Allowed","Success Rate% (Offense)","Pressures"])
                elif name == "DVOA_Proxy":
                    ws.append(["Week","Opponent","Off Adj EPA/play","Off Adj SR%",
                               "Def Adj EPA/play","Def Adj SR%","Off EPA/play","Def EPA allowed/play"])
                elif name == "Predictions":
                    ws.append(["Week","Prediction","Reason","Notes"])
                elif name == "Media_Summaries":
                    ws.append(["Week","Opponent","Summary"])
                else:
                    pass
        wb.save(EXCEL_FILE)

def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True):
    """
    Replace records for any Week values in new_data when deduplicate=True.
    Keeps headers and preserves column order.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    ensure_workbook_structure()

    try:
        book = openpyxl.load_workbook(file_name)
        # Read existing
        if sheet_name in book.sheetnames:
            sheet = book[sheet_name]
            existing = pd.DataFrame(sheet.values)
            if not existing.empty:
                existing.columns = existing.iloc[0]
                existing = existing[1:]
            else:
                existing = pd.DataFrame()
        else:
            # Create sheet if missing
            sheet = book.create_sheet(sheet_name)
            existing = pd.DataFrame()

        # Normalize types
        if not existing.empty and "Week" in existing.columns:
            existing["Week"] = pd.to_numeric(existing["Week"], errors="coerce").astype("Int64")
        if "Week" in new_data.columns:
            new_data["Week"] = pd.to_numeric(new_data["Week"], errors="coerce").astype("Int64")

        # Merge/replace same weeks
        if deduplicate and not existing.empty and "Week" in existing.columns and "Week" in new_data.columns:
            weeks_to_replace = set(new_data["Week"].dropna().unique().tolist())
            existing = existing[~existing["Week"].isin(weeks_to_replace)]

        # Union columns (to keep headers stable)
        combined = pd.concat([existing, new_data], ignore_index=True)
        # Sort by Week if present
        if "Week" in combined.columns:
            combined = combined.sort_values(by="Week", kind="mergesort", na_position="last").reset_index(drop=True)

        # Rewrite sheet
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        for r in dataframe_to_rows(combined, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def safe_float(val, default=None):
    try:
        if val is None:
            return default
        if isinstance(val, float) and pd.isna(val):
            return default
        return float(val)
    except Exception:
        return default

# ------------------- Sidebar: Upload CSVs -------------------
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv", key="up_off")
uploaded_defense   = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv", key="up_def")
uploaded_strategy  = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv", key="up_strat")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv", key="up_pers")
uploaded_injuries  = st.sidebar.file_uploader("Upload Injuries (.csv)", type="csv", key="up_inj")

if uploaded_offense:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense")
    st.sidebar.success("✅ Offensive data uploaded.")
if uploaded_defense:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense")
    st.sidebar.success("✅ Defensive data uploaded.")
if uploaded_strategy:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy")
    st.sidebar.success("✅ Strategy data uploaded.")
if uploaded_personnel:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel")
    st.sidebar.success("✅ Personnel data uploaded.")
if uploaded_injuries:
    df_inj = pd.read_csv(uploaded_injuries)
    append_to_excel(df_inj, "Injuries")
    st.sidebar.success("✅ Injuries data uploaded.")

# ------------------- Sidebar: Fetch (nfl_data_py) -------------------
with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pull 2025 weekly team stats for CHI and save to Excel.")
    fetch_week = st.number_input("Week to fetch (2025)", min_value=1, max_value=25, value=1, step=1)

    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl
            # Try local refresh (safe if fails)
            try:
                nfl.update.schedule_data([2025])
            except Exception:
                pass
            try:
                nfl.update.weekly_data([2025])
            except Exception:
                pass

            weekly = nfl.import_weekly_data([2025])  # team-level weekly stats
            wk = int(fetch_week)
            team_week = weekly[(weekly["team"] == "CHI") & (weekly["week"] == wk)].copy()

            if team_week.empty:
                st.warning("No weekly team row found for CHI in that week yet.")
            else:
                # Prepare Offense
                opp = team_week.get("opponent")
                opponent = opp.values[0] if opp is not None else None

                pass_yards = team_week.get("passing_yards")
                pass_yards = pass_yards.values[0] if pass_yards is not None else None

                pass_att = None
                for cand in ["attempts", "passing_attempts", "pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].values[0]
                        break

                try:
                    ypa_val = float(pass_yards) / float(pass_att) if pass_yards is not None and pass_att not in (None, 0) else None
                except Exception:
                    ypa_val = None

                yards_total = None
                for cand in ["yards", "total_yards", "offense_yards"]:
                    if cand in team_week.columns:
                        yards_total = team_week[cand].values[0]
                        break

                completions = None
                for cand in ["completions", "passing_completions", "pass_completions"]:
                    if cand in team_week.columns:
                        completions = team_week[cand].values[0]
                        break

                cmp_pct = None
                if completions is not None and pass_att not in (None, 0):
                    try:
                        cmp_pct = round((float(completions) / float(pass_att)) * 100, 1)
                    except Exception:
                        cmp_pct = None

                off_row = pd.DataFrame([{
                    "Week": wk,
                    "Opponent": opponent,
                    "YDS": yards_total,
                    "YPA": round(ypa_val, 2) if ypa_val is not None else None,
                    "CMP%": cmp_pct
                }])

                # Defense basics
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].values[0]
                        break
                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val,
                    "RZ% Allowed": None
                }])

                # Save
                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)
                st.success(f"✅ Added CHI Week {wk} to Offense/Defense (available fields).")
                st.caption("Note: Red Zone % Allowed and Pressures require Play-by-Play (see panel below).")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

st.sidebar.markdown("### 📡 Fetch Defensive Metrics from Play-by-Play")
pbp_week = st.sidebar.number_input("Week to Fetch (2025 Season)", min_value=1, max_value=25, value=1, step=1)
if st.sidebar.button("Fetch Play-by-Play Metrics"):
    try:
        import nfl_data_py as nfl
        pbp = nfl.import_pbp_data([2025], downcast=False)
        pbp_w = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == "CHI")].copy()

        if pbp_w.empty:
            st.warning("No CHI defensive PBP for that week yet.")
        else:
            # Red Zone % Allowed
            dmins = (
                pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success Rate (offense success vs CHI defense)
            def play_success(row):
                try:
                    if pd.isna(row.get("down")) or pd.isna(row.get("ydstogo")) or pd.isna(row.get("yards_gained")):
                        return False
                    d = int(row["down"]); togo = float(row["ydstogo"]); gain = float(row["yards_gained"])
                    if d == 1:
                        return gain >= 0.4 * togo
                    elif d == 2:
                        return gain >= 0.6 * togo
                    else:
                        return gain >= togo
                except Exception:
                    return False

            plays_mask = (~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))
            pbp_real = pbp_w[plays_mask].copy()
            success_rate = pbp_real.apply(play_success, axis=1).mean() * 100 if len(pbp_real) else 0.0

            # Pressures approx = sacks + qb hits
            qb_hits = pbp_w["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in pbp_w.columns else 0
            sacks = pbp_w["sack"].fillna(0).astype(int).sum() if "sack" in pbp_w.columns else 0
            pressures = int(qb_hits + sacks)

            metrics_df = pd.DataFrame([{
                "Week": int(pbp_week),
                "RZ% Allowed": round(rz_allowed, 1),
                "Success Rate% (Offense)": round(success_rate, 1),
                "Pressures": pressures
            }])
            append_to_excel(metrics_df, "Advanced_Defense", deduplicate=True)
            st.success(
                f"✅ Week {int(pbp_week)} PBP metrics saved — RZ% Allowed: {rz_allowed:.1f} | "
                f"Success Rate% (Off): {success_rate:.1f} | Pressures: {pressures}"
            )
    except Exception as e:
        st.error(f"❌ Failed to fetch metrics: {e}")

# ------------------- Compute DVOA-like Proxy (Opponent Adjusted) -------------------
st.sidebar.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.sidebar.number_input("Week to Compute (2025 Season)", min_value=1, max_value=25, value=1, step=1)

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
            if not bears_off.empty: opps.update(bears_off["defteam"].unique().tolist())
            if not bears_def.empty: opps.update(bears_def["posteam"].unique().tolist())
            opponent = list(opps)[0] if opps else "UNK"

            prior = plays[plays["week"] < wk].copy()

            opp_def_plays = prior[prior["defteam"] == opponent].copy()
            opp_def_epa = opp_def_plays["epa"].mean() if len(opp_def_plays) else None
            opp_def_success = (opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
                               if len(opp_def_plays) else None)

            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            opp_off_success = (opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
                               if len(opp_off_plays) else None)

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

            append_to_excel(out, "DVOA_Proxy", deduplicate=True)
            st.success(
                f"✅ DVOA-like proxy saved for Week {wk} vs {opponent} "
                f"(Off Adj EPA/play={out.iloc[0]['Off Adj EPA/play']}, Off Adj SR%={out.iloc[0]['Off Adj SR%']}, "
                f"Def Adj EPA/play={out.iloc[0]['Def Adj EPA/play']}, Def Adj SR%={out.iloc[0]['Def Adj SR%']})"
            )
    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# ------------------- Excel Sanity Checker (Optional) -------------------
with st.expander("🛠️ Excel Sanity Checker (optional)"):
    st.write(f"**Current Working Directory:** `{os.getcwd()}`")
    st.write(f"**Excel File Being Used:** `{EXCEL_FILE}`")
    path = os.path.abspath(EXCEL_FILE)
    st.write(f"**Excel file path:** `{path}`")
    st.write(f"**Exists:** {os.path.exists(path)}")
    if os.path.exists(path):
        st.write(f"**Size (bytes):** {os.path.getsize(path)}")
        try:
            xl = pd.ExcelFile(path)
            st.write("**Sheets:**")
            for i, s in enumerate(xl.sheet_names):
                st.write(f"{i}: {s}")
            # small previews
            for s in ["Offense","Defense","Strategy","Personnel","DVOA_Proxy","Predictions"]:
                try:
                    dfp = xl.parse(s)
                    st.write(f"**Preview – {s}**")
                    st.dataframe(dfp.head(10))
                except Exception:
                    pass
            st.success("Workbook structure looks good.")
        except Exception as e:
            st.error(f"Sanity read failed: {e}")

    if st.button("Create/Repair Workbook"):
        ensure_workbook_structure()
        st.success("✅ Workbook ensured (created missing sheets/headers).")

# ------------------- Download: Raw + Formatted (green/red) -------------------
if os.path.exists(EXCEL_FILE):
    # Raw
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel - raw)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_raw"
        )

    # Formatted copy
    try:
        from openpyxl import load_workbook
        from openpyxl.formatting.rule import ColorScaleRule
        from openpyxl.utils import get_column_letter

        def add_color_scale(ws, col_idx: int, reverse: bool = False):
            if col_idx < 1 or ws.max_row <= 1:
                return
            green = "63BE7B"; yellow = "FFEB84"; red = "F8696B"
            start, mid, end = (red, yellow, green) if not reverse else (green, yellow, red)
            rng = f"{get_column_letter(col_idx)}2:{get_column_letter(col_idx)}{ws.max_row}"
            rule = ColorScaleRule(start_type="min", start_color=start,
                                  mid_type="percentile", mid_value=50, mid_color=mid,
                                  end_type="max", end_color=end)
            ws.conditional_formatting.add(rng, rule)

        def apply_defaults(wb):
            targets = {
                "Offense": {
                    "YDS": False, "YPA": False, "YPC": False, "CMP%": False, "QBR": False,
                    "SR%": False, "DIV Avg YDS": False, "DIV Avg QBR": False, "DIV Avg SR%": False,
                    "CONF Avg YDS": False, "CONF Avg QBR": False, "CONF Avg SR%": False,
                    "NFL Avg YDS": False, "NFL Avg QBR": False, "NFL Avg SR%": False,
                },
                "Defense": {
                    "SACK": False, "INT": False, "FF": False, "FR": False,
                    "QB Hits": False, "Pressures": False,
                    "3D% Allowed": True, "RZ% Allowed": True,
                },
                "Advanced_Defense": {
                    "RZ% Allowed": True, "Success Rate% (Offense)": True, "Pressures": False
                },
                "DVOA_Proxy": {
                    "Off Adj EPA/play": False, "Off Adj SR%": False,
                    "Def Adj EPA/play": True, "Def Adj SR%": True
                },
                "Personnel": {
                    "11 Personnel": False, "12 Personnel": False, "13 Personnel": False, "21 Personnel": False,
                    "Division 11": False, "Division 12": False, "Division 13": False, "Division 21": False,
                    "Conf 11": False, "Conf 12": False, "Conf 13": False, "Conf 21": False,
                    "NFL 11": False, "NFL 12": False, "NFL 13": False, "NFL 21": False,
                }
            }
            for sheet_name, cols in targets.items():
                if sheet_name not in wb.sheetnames:
                    continue
                ws = wb[sheet_name]
                if ws.max_row < 2:
                    continue
                # headers map
                headers = {}
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(row=1, column=c).value
                    if isinstance(v, str) and v.strip():
                        headers[v.strip()] = c
                for hdr, rev in cols.items():
                    col_idx = headers.get(hdr)
                    if col_idx:
                        add_color_scale(ws, col_idx, reverse=rev)

        formatted_path = EXCEL_FILE.replace(".xlsx", "_formatted.xlsx")
        wb = load_workbook(EXCEL_FILE)
        apply_defaults(wb)
        wb.save(formatted_path)

        with open(formatted_path, "rb") as f2:
            st.sidebar.download_button(
                label="⬇️ Download All Data (Excel - formatted)",
                data=f2,
                file_name=os.path.basename(formatted_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_formatted"
            )
    except Exception as e:
        st.sidebar.error(f"Could not create formatted Excel: {e}")

# ------------------- Main: Show Uploaded/Stored Data -------------------
if os.path.exists(EXCEL_FILE):
    try:
        df_off_show = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        st.subheader("📊 Offensive Analytics")
        st.dataframe(df_off_show)
    except Exception:
        st.info("No Offense sheet yet.")
    try:
        df_def_show = pd.read_excel(EXCEL_FILE, sheet_name="Defense")
        st.subheader("🛡️ Defensive Analytics")
        st.dataframe(df_def_show)
    except Exception:
        st.info("No Defense sheet yet.")
    try:
        df_str_show = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        st.subheader("📘 Weekly Strategy")
        st.dataframe(df_str_show)
    except Exception:
        st.info("No Strategy sheet yet.")
    try:
        df_per_show = pd.read_excel(EXCEL_FILE, sheet_name="Personnel")
        st.subheader("👥 Personnel Usage")
        st.dataframe(df_per_show)
    except Exception:
        st.info("No Personnel sheet yet.")
    try:
        df_inj_show = pd.read_excel(EXCEL_FILE, sheet_name="Injuries")
        st.subheader("🩺 Injuries")
        st.dataframe(df_inj_show)
    except Exception:
        st.info("No Injuries sheet yet.")
    try:
        df_dvoa_show = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
        if not df_dvoa_show.empty:
            st.subheader("📈 DVOA-like Proxy Metrics")
            st.dataframe(df_dvoa_show.tail(5))
    except Exception:
        pass

# ------------------- Media Summaries -------------------
st.markdown("### 📰 Weekly Beat Writer / ESPN Summary")
with st.form("media_form"):
    media_week = st.number_input("Week", min_value=1, max_value=25, step=1, key="media_week_input")
    media_opponent = st.text_input("Opponent")
    media_summary = st.text_area("Beat Writer & ESPN Summary (Game Recap, Analysis, Strategy, etc.)")
    submit_media = st.form_submit_button("Save Summary")

if submit_media:
    media_df = pd.DataFrame([{
        "Week": media_week,
        "Opponent": media_opponent,
        "Summary": media_summary
    }])
    append_to_excel(media_df, "Media_Summaries", deduplicate=False)
    st.success(f"✅ Summary for Week {media_week} vs {media_opponent} saved.")

# ------------------- Weekly Prediction -------------------
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        df_offense  = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        df_defense  = pd.read_excel(EXCEL_FILE, sheet_name="Defense")

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

            ypa = safe_float(row_o.iloc[0].get("YPA"), default=None)

            rz_allowed = None
            pressures = None
            if not row_a.empty:
                rz_allowed = safe_float(row_a.iloc[0].get("RZ% Allowed"), default=None)
                pressures = safe_float(row_a.iloc[0].get("Pressures"), default=None)
            if rz_allowed is None:
                rz_allowed = safe_float(row_d.iloc[0].get("RZ% Allowed"), default=None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = safe_float(row_p.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = safe_float(row_p.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = safe_float(row_p.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = safe_float(row_p.iloc[0].get("Def Adj SR%"), default=None)

            # Rule set
            reason_bits = []
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win - efficiency edge on both sides"
                reason_bits.append(f"Off+{off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play vs opp O")
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win - pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win - red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient offense and poor red zone defense"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss - inefficient passing and weak red zone defense"
                reason_bits.append(f"YPA={ypa:.1f}")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            else:
                prediction = "Loss - no clear advantage in key strategy or stats"
                if off_adj_epa is not None:
                    reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                if def_adj_epa is not None:
                    reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play")
                if pressures is not None:
                    reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            reason_text = " | ".join(reason_bits)
            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")
            if reason_text:
                st.caption(reason_text)

            # Save prediction
            pred_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("-")[0].strip(),
                "Reason": prediction.split("-")[1].strip() if "-" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(pred_entry, "Predictions", deduplicate=True)
        else:
            st.info("Please upload or fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# Show saved predictions table
if os.path.exists(EXCEL_FILE):
    try:
        df_preds = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
        st.subheader("📈 Saved Game Predictions")
        st.dataframe(df_preds)
    except Exception:
        st.info("No predictions saved yet.")
import os
import pandas as pd
import streamlit as st
from fpdf import FPDF

# ------------------------------
# App setup
# ------------------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, snap counts, and opponent-adjusted metrics.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# ------------------------------
# Helpers
# ------------------------------
def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True, dedup_keys=("Week",)):
    """
    Append/replace data into an Excel sheet.
    If deduplicate=True and dedup_keys exist, it removes rows in the existing sheet that match the new rows on those keys.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    try:
        # ensure columns are strings
        new_data = new_data.copy()
        new_data.columns = [str(c) for c in new_data.columns]

        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                sh = book[sheet_name]
                existing = pd.DataFrame(sh.values)
                if not existing.empty:
                    existing.columns = existing.iloc[0]
                    existing = existing[1:]
                else:
                    existing = pd.DataFrame(columns=new_data.columns)

                # try to align columns
                for c in new_data.columns:
                    if c not in existing.columns:
                        existing[c] = pd.NA
                for c in existing.columns:
                    if c not in new_data.columns:
                        new_data[c] = pd.NA
                existing = existing[new_data.columns]

                if deduplicate and all(k in existing.columns for k in dedup_keys) and all(k in new_data.columns for k in dedup_keys):
                    # remove any existing rows that match any (Week) of the incoming rows
                    key_values = new_data[list(dedup_keys)].drop_duplicates()
                    # build boolean mask to keep rows NOT matching any key row
                    mask = pd.Series([True] * len(existing))
                    for _, kv in key_values.iterrows():
                        cond = pd.Series([True] * len(existing))
                        for k in dedup_keys:
                            cond &= (existing[k].astype(str) == str(kv[k]))
                        mask &= ~cond
                    existing = existing[mask]

                combined = pd.concat([existing, new_data], ignore_index=True)
            else:
                combined = new_data
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            combined = new_data

        # rewrite sheet
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)
        for r in dataframe_to_rows(combined, index=False, header=True):
            sheet.append(r)
        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def ensure_workbook_basics():
    """
    Make sure workbook exists and key sheets are present (empty if missing).
    """
    base_sheets = {
        "Offense": ["Week", "YDS", "YPA", "YPC", "CMP%"],
        "Defense": ["Week", "SACK", "INT", "FF", "FR", "3D% Allowed", "RZ% Allowed", "QB Hits", "Pressures"],
        "Strategy": ["Week", "Opponent", "Off_Strategy", "Def_Strategy", "Notes"],
        "Personnel": ["Week", "11 Personnel", "12 Personnel", "13 Personnel", "21 Personnel"],
        "DVOA_Proxy": ["Week", "Opponent", "Off Adj EPA/play", "Off Adj SR%", "Def Adj EPA/play", "Def Adj SR%"],
        "Predictions": ["Week", "Prediction", "Reason", "Notes"],
        "Injuries": ["Week", "Player", "Position", "Status", "Body_Part", "Practice", "Game_Status", "Updated"],
        "SnapCounts": ["Week", "Side", "Player", "Position", "Snaps", "Pct"]
    }
    for sheet, cols in base_sheets.items():
        if not os.path.exists(EXCEL_FILE):
            append_to_excel(pd.DataFrame(columns=cols), sheet, deduplicate=False)
        else:
            try:
                _ = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
            except Exception:
                append_to_excel(pd.DataFrame(columns=cols), sheet, deduplicate=False)

# simple style helpers (green/red)
def style_pos_neg(val):
    try:
        v = float(val)
        if v > 0:
            return "background-color: #b6e3b6; color: #000000"
        if v < 0:
            return "background-color: #f7b0a9; color: #000000"
    except Exception:
        return ""
    return ""

# ------------------------------
# Ensure workbook/sheets
# ------------------------------
ensure_workbook_basics()

# ------------------------------
# Sidebar: Upload CSVs
# ------------------------------
st.sidebar.header("📤 Upload Weekly CSVs")
uploaded_offense = st.sidebar.file_uploader("Offense (.csv)", type="csv")
uploaded_defense = st.sidebar.file_uploader("Defense (.csv)", type="csv")
uploaded_strategy = st.sidebar.file_uploader("Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Personnel (.csv)", type="csv")
uploaded_injuries = st.sidebar.file_uploader("Injuries (.csv)", type="csv")

if uploaded_offense:
    try:
        df = pd.read_csv(uploaded_offense)
        append_to_excel(df, "Offense", deduplicate=True, dedup_keys=("Week",))
        st.sidebar.success("✅ Offense uploaded")
    except Exception as e:
        st.sidebar.error(f"Offense upload failed: {e}")

if uploaded_defense:
    try:
        df = pd.read_csv(uploaded_defense)
        append_to_excel(df, "Defense", deduplicate=True, dedup_keys=("Week",))
        st.sidebar.success("✅ Defense uploaded")
    except Exception as e:
        st.sidebar.error(f"Defense upload failed: {e}")

if uploaded_strategy:
    try:
        df = pd.read_csv(uploaded_strategy)
        append_to_excel(df, "Strategy", deduplicate=True, dedup_keys=("Week",))
        st.sidebar.success("✅ Strategy uploaded")
    except Exception as e:
        st.sidebar.error(f"Strategy upload failed: {e}")

if uploaded_personnel:
    try:
        df = pd.read_csv(uploaded_personnel)
        append_to_excel(df, "Personnel", deduplicate=True, dedup_keys=("Week",))
        st.sidebar.success("✅ Personnel uploaded")
    except Exception as e:
        st.sidebar.error(f"Personnel upload failed: {e}")

if uploaded_injuries:
    try:
        df = pd.read_csv(uploaded_injuries)
        append_to_excel(df, "Injuries", deduplicate=True, dedup_keys=("Week", "Player"))
        st.sidebar.success("✅ Injuries uploaded")
    except Exception as e:
        st.sidebar.error(f"Injuries upload failed: {e}")

# ------------------------------
# Sidebar: Fetch blocks
# ------------------------------
with st.sidebar.expander("⚡ Fetch Weekly Team Data (nfl_data_py)"):
    st.caption("Pull basic weekly CHI team stats (Off/Def) for 2025.")
    fw = st.number_input("Week to fetch", min_value=1, max_value=25, value=1, step=1, key="fetch_wk")
    if st.button("Fetch CHI Week (Team)"):
        try:
            import nfl_data_py as nfl
            weekly = nfl.import_weekly_data([2025])
            wk = int(fw)
            team_week = weekly[(weekly["team"] == "CHI") & (weekly["week"] == wk)].copy()
            if team_week.empty:
                st.warning("No weekly row found yet for CHI.")
            else:
                # Offense quick picks
                yds = None
                for c in ["yards", "total_yards", "offense_yards"]:
                    if c in team_week.columns:
                        yds = team_week[c].iloc[0]
                        break
                # YPA best effort
                pass_yards = team_week["passing_yards"].iloc[0] if "passing_yards" in team_week.columns else None
                pass_att = None
                for cand in ["attempts", "passing_attempts", "pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].iloc[0]
                        break
                ypa = None
                if pass_yards is not None and pass_att not in (None, 0):
                    try:
                        ypa = round(float(pass_yards) / float(pass_att), 2)
                    except Exception:
                        ypa = None

                cmp_pct = None
                if "completions" in team_week.columns and pass_att not in (None, 0):
                    try:
                        cmp_pct = round(float(team_week["completions"].iloc[0]) / float(pass_att) * 100, 1)
                    except Exception:
                        cmp_pct = None

                off_row = pd.DataFrame([{"Week": wk, "YDS": yds, "YPA": ypa, "CMP%": cmp_pct}])
                append_to_excel(off_row, "Offense", deduplicate=True)

                sacks_val = None
                for cc in ["sacks", "defense_sacks"]:
                    if cc in team_week.columns:
                        sacks_val = team_week[cc].iloc[0]
                        break
                def_row = pd.DataFrame([{"Week": wk, "SACK": sacks_val}])
                append_to_excel(def_row, "Defense", deduplicate=True)

                st.success(f"✅ Added team Off/Def for week {wk}")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

with st.sidebar.expander("📡 Fetch Defensive PBP Metrics"):
    st.caption("Play-by-play: RZ% Allowed, Success Rate, Pressures (2025).")
    pbp_wk = st.number_input("Week", min_value=1, max_value=25, value=1, step=1, key="pbp_wk")
    if st.button("Fetch PBP for CHI Defense"):
        try:
            import nfl_data_py as nfl
            pbp = nfl.import_pbp_data([2025], downcast=False)
            pbp_w = pbp[(pbp["week"] == int(pbp_wk)) & (pbp["defteam"] == "CHI")].copy()
            if pbp_w.empty:
                st.warning("No PBP rows for that week.")
            else:
                # RZ% Allowed via min yardline in drive
                dmins = (
                    pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                    .min()
                    .rename(columns={"yardline_100": "min_yardline_100"})
                )
                total_drives = len(dmins)
                rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
                rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

                # Success Rate (offense success vs CHI)
                def success(down, togo, gain):
                    try:
                        d = int(down)
                        t = float(togo)
                        g = float(gain)
                        if d == 1:
                            return g >= 0.4 * t
                        elif d == 2:
                            return g >= 0.6 * t
                        return g >= t
                    except Exception:
                        return False

                plays = pbp_w[(~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))].copy()
                sr = plays.apply(lambda r: success(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() * 100 if len(plays) else 0.0

                qb_hits = pbp_w["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in pbp_w.columns else 0
                sacks = pbp_w["sack"].fillna(0).astype(int).sum() if "sack" in pbp_w.columns else 0
                pressures = int(qb_hits + sacks)

                adv = pd.DataFrame([{
                    "Week": int(pbp_wk),
                    "RZ% Allowed": round(rz_allowed, 1),
                    "Success Rate% (Offense)": round(sr, 1),
                    "Pressures": pressures
                }])
                append_to_excel(adv, "Advanced_Defense", deduplicate=True)
                st.success(f"✅ Saved PBP metrics for week {int(pbp_wk)}")
        except Exception as e:
            st.error(f"PBP fetch failed: {e}")

# ------------------------------
# DVOA-like proxy (opponent-adjusted EPA/SR)
# ------------------------------
st.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.number_input("Week to compute (2025)", min_value=1, max_value=25, value=1, step=1, key="proxy_week")

def _success_flag(down, ydstogo, yards_gained):
    try:
        if pd.isna(down) or pd.isna(ydstogo) or pd.isna(yards_gained):
            return False
        d = int(down); t = float(ydstogo); g = float(yards_gained)
        if d == 1:
            return g >= 0.4 * t
        elif d == 2:
            return g >= 0.6 * t
        return g >= t
    except Exception:
        return False

if st.button("Compute DVOA-like Proxy"):
    try:
        import nfl_data_py as nfl
        wk = int(proxy_week)
        pbp = nfl.import_pbp_data([2025], downcast=False)

        plays = pbp[(~pbp["play_type"].isin(["no_play"])) & (~pbp["penalty"].fillna(False)) & (~pbp["epa"].isna())].copy()
        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == "CHI")].copy()
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == "CHI")].copy()

        if bears_off.empty and bears_def.empty:
            st.warning("No CHI plays for that week yet.")
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
            opp_def_sr = opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_def_plays) else None

            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            opp_off_sr = opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_off_plays) else None

            if len(bears_off):
                chi_off_epa = bears_off["epa"].mean()
                chi_off_sr = bears_off.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_off_epa = None; chi_off_sr = None

            if len(bears_def):
                chi_def_epa_allowed = bears_def["epa"].mean()
                chi_def_sr_allowed = bears_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_def_epa_allowed = None; chi_def_sr_allowed = None

            def diff(a, b):
                if a is None or pd.isna(a) or b is None or pd.isna(b): return None
                return float(a) - float(b)

            off_adj_epa = diff(chi_off_epa, opp_def_epa)
            off_adj_sr  = diff(chi_off_sr,  opp_def_sr)
            def_adj_epa = diff(chi_def_epa_allowed, opp_off_epa)
            def_adj_sr  = diff(chi_def_sr_allowed,  opp_off_sr)

            out = pd.DataFrame([{
                "Week": wk,
                "Opponent": opponent,
                "Off Adj EPA/play": round(off_adj_epa, 3) if off_adj_epa is not None else None,
                "Off Adj SR%": round(off_adj_sr * 100, 1) if off_adj_sr is not None else None,
                "Def Adj EPA/play": round(def_adj_epa, 3) if def_adj_epa is not None else None,
                "Def Adj SR%": round(def_adj_sr * 100, 1) if def_adj_sr is not None else None
            }])
            append_to_excel(out, "DVOA_Proxy", deduplicate=True)
            st.success(f"✅ DVOA-like proxy saved for Week {wk} vs {opponent}")
    except Exception as e:
        st.error(f"Failed to compute proxy: {e}")

# ------------------------------
# DVOA Proxy Preview (with colors)
# ------------------------------
st.markdown("### 📊 DVOA Proxy Results")
try:
    dvoa_df = pd.read_excel(EXCEL_FILE, sheet_name="DVOA_Proxy")
    if not dvoa_df.empty:
        # Color the two EPA columns green/red; SR% is already a percent (higher offense SR is good, lower def SR is good).
        style_cols = []
        if "Off Adj EPA/play" in dvoa_df.columns: style_cols.append("Off Adj EPA/play")
        if "Def Adj EPA/play" in dvoa_df.columns: style_cols.append("Def Adj EPA/play")

        if style_cols:
            st.dataframe(
                dvoa_df.style.applymap(style_pos_neg, subset=style_cols)
            )
        else:
            st.dataframe(dvoa_df)
    else:
        st.info("No DVOA Proxy data yet.")
except Exception:
    st.info("No DVOA Proxy data yet.")

# ------------------------------
# Main previews for uploaded data
# ------------------------------
colA, colB = st.columns(2)
with colA:
    st.subheader("Offensive Analytics")
    try:
        df_off = pd.read_excel(EXCEL_FILE, sheet_name="Offense")
        st.dataframe(df_off)
    except Exception:
        st.info("No Offense yet.")

    st.subheader("Strategy")
    try:
        df_str = pd.read_excel(EXCEL_FILE, sheet_name="Strategy")
        st.dataframe(df_str)
    except Exception:
        st.info("No Strategy yet.")

with colB:
    st.subheader("Defensive Analytics")
    try:
        df_def = pd.read_excel(EXCEL_FILE, sheet_name="Defense")
        st.dataframe(df_def)
    except Exception:
        st.info("No Defense yet.")

    st.subheader("Personnel")
    try:
        df_per = pd.read_excel(EXCEL_FILE, sheet_name="Personnel")
        st.dataframe(df_per)
    except Exception:
        st.info("No Personnel yet.")

# ------------------------------
# Injuries (quick entry)
# ------------------------------
st.markdown("### 🏥 Injuries (Quick Entry)")
with st.form("inj_form"):
    inj_week = st.number_input("Week", min_value=1, max_value=25, value=1, step=1)
    inj_player = st.text_input("Player")
    inj_pos = st.text_input("Position")
    inj_status = st.selectbox("Status", ["Questionable", "Doubtful", "Out", "IR", "Active"])
    inj_body = st.text_input("Body Part")
    inj_prac = st.selectbox("Practice", ["DNP", "Limited", "Full", "N/A"])
    inj_game = st.selectbox("Game Status", ["TBD", "Active", "Inactive"])
    submitted = st.form_submit_button("Save Injury")
if submitted and inj_player.strip():
    row = pd.DataFrame([{
        "Week": inj_week, "Player": inj_player.strip(), "Position": inj_pos.strip(),
        "Status": inj_status, "Body_Part": inj_body.strip(), "Practice": inj_prac,
        "Game_Status": inj_game, "Updated": pd.Timestamp.now(tz="US/Central").strftime("%Y-%m-%d %H:%M")
    }])
    append_to_excel(row, "Injuries", deduplicate=True, dedup_keys=("Week", "Player"))
    st.success("✅ Injury saved")

st.subheader("Injuries Table")
try:
    df_inj = pd.read_excel(EXCEL_FILE, sheet_name="Injuries")
    st.dataframe(df_inj)
except Exception:
    st.info("No injuries yet.")

# ------------------------------
# Snap Counts (best effort)
# ------------------------------
st.markdown("### ⏱️ Snap Counts (Best Effort)")
with st.form("snap_quick_add"):
    sc_week = st.number_input("Week (Snaps)", min_value=1, max_value=25, value=1, step=1)
    sc_side = st.selectbox("Side", ["Offense", "Defense"])
    sc_player = st.text_input("Player (optional)")
    sc_pos = st.text_input("Position (optional)")
    sc_snaps = st.number_input("Snaps", min_value=0, max_value=200, value=0, step=1)
    sc_pct = st.number_input("Pct", min_value=0.0, max_value=100.0, value=0.0, step=0.1)
    sc_submit = st.form_submit_button("Save Snap Row")
if sc_submit:
    row = pd.DataFrame([{
        "Week": sc_week, "Side": sc_side, "Player": sc_player, "Position": sc_pos,
        "Snaps": sc_snaps, "Pct": sc_pct
    }])
    append_to_excel(row, "SnapCounts", deduplicate=True, dedup_keys=("Week", "Side", "Player"))
    st.success("✅ Snap row saved")

try:
    df_snaps = pd.read_excel(EXCEL_FILE, sheet_name="SnapCounts")
    st.dataframe(df_snaps)
except Exception:
    st.info("No snap counts yet.")

# ------------------------------
# Weekly Prediction
# ------------------------------
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, value=1, step=1, key="predict_week_input")

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

        rs = df_strategy[df_strategy["Week"] == week_to_predict]
        ro = df_offense[df_offense["Week"] == week_to_predict]
        rd = df_defense[df_defense["Week"] == week_to_predict]
        ra = df_advdef[df_advdef["Week"] == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        rp = df_proxy[df_proxy["Week"] == week_to_predict] if not df_proxy.empty else pd.DataFrame()

        if not rs.empty and not ro.empty and not rd.empty:
            strat_text = rs.iloc[0].astype(str).str.cat(sep=" ").lower()
            ypa = _safe_float(ro.iloc[0].get("YPA"), default=None)

            rz_allowed = None
            pressures = None
            if not ra.empty:
                rz_allowed = _safe_float(ra.iloc[0].get("RZ% Allowed"), default=None)
                pressures  = _safe_float(ra.iloc[0].get("Pressures"), default=None)
            if rz_allowed is None:
                rz_allowed = _safe_float(rd.iloc[0].get("RZ% Allowed"), default=None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not rp.empty:
                off_adj_epa = _safe_float(rp.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = _safe_float(rp.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = _safe_float(rp.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = _safe_float(rp.iloc[0].get("Def Adj SR%"), default=None)

            reason_bits = []
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win – efficiency edge on both sides"
                reason_bits += [f"Off+{off_adj_epa:+.2f} EPA/play vs opp D", f"Def{def_adj_epa:+.2f} EPA/play vs opp O"]
            elif (pressures is not None and pressures >= 8) and ("blitz" in strat_text or "pressure" in strat_text):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strat_text for tok in ["zone", "two-high", "split-safety"]):
                prediction = "Win – red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient offense and poor red zone defense"
                reason_bits += [f"Off{off_adj_epa:+.2f} EPA/play", f"RZ% Allowed={rz_allowed:.0f}"]
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient passing and weak red zone defense"
                reason_bits += [f"YPA={ypa:.1f}", f"RZ% Allowed={rz_allowed:.0f}"]
            else:
                prediction = "Loss – no clear advantage in key strategy or stats"
                if off_adj_epa is not None: reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                if def_adj_epa is not None: reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play")
                if pressures is not None:   reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:  reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")

            reason_text = " | ".join(reason_bits)
            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")
            if reason_text:
                st.caption(reason_text)

            save_row = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(save_row, "Predictions", deduplicate=True)
        else:
            st.info("Please upload/fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check data. Error: {e}")

st.subheader("📈 Saved Game Predictions")
try:
    df_preds = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
    st.dataframe(df_preds)
except Exception:
    st.info("No predictions saved yet.")

# ------------------------------
# Download workbook
# ------------------------------
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
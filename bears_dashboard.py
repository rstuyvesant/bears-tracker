import os
import streamlit as st
import pandas as pd
from fpdf import FPDF

# ==============================================
# Basic App Setup
# ==============================================
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, snap counts, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# ==============================================
# Helpers: Excel append with de-duplication
# ==============================================
def _coerce_week_col(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure a 'Week' column exists and is numeric where possible."""
    if "Week" not in df.columns:
        return df
    out = df.copy()
    try:
        out["Week"] = pd.to_numeric(out["Week"], errors="coerce").astype("Int64")
    except Exception:
        pass
    return out

def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True):
    """Append or replace rows in an Excel sheet with optional deduplication.
       - If sheet has 'Week' and new_data has 'Week' -> de-dupe by Week
       - If sheet is 'Injuries' and has ['Week','Player'] -> de-dupe by (Week, Player)
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    if new_data is None or new_data.empty:
        return

    new_df = new_data.copy()
    new_df = _coerce_week_col(new_df)

    try:
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                sheet = book[sheet_name]
                existing_data = pd.DataFrame(sheet.values)
                if not existing_data.empty:
                    existing_data.columns = existing_data.iloc[0]
                    existing_data = existing_data[1:]
                else:
                    existing_data = pd.DataFrame(columns=new_df.columns)

                # align columns
                existing_data = existing_data.reindex(columns=new_df.columns, fill_value=pd.NA)
                existing_data = _coerce_week_col(existing_data)

                if deduplicate:
                    if sheet_name == "Injuries" and {"Week", "Player"}.issubset(new_df.columns) and \
                       {"Week", "Player"}.issubset(existing_data.columns):
                        # Remove existing rows for same (Week, Player)
                        key = ["Week", "Player"]
                        left = existing_data
                        to_remove = new_df[key].drop_duplicates()
                        left = left.merge(to_remove.assign(_drop=1), on=key, how="left")
                        existing_data = left[left["_drop"].isna()].drop(columns=["_drop"])
                    elif "Week" in existing_data.columns and "Week" in new_df.columns:
                        # Remove existing rows that match incoming Week(s)
                        wks = new_df["Week"].dropna().unique().tolist()
                        existing_data = existing_data[~existing_data["Week"].isin(wks)]

                combined_data = pd.concat([existing_data, new_df], ignore_index=True)
            else:
                combined_data = new_df
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            combined_data = new_df

        # recreate the sheet
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        # write rows
        for r in dataframe_to_rows(combined_data, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

# ==============================================
# Styling helpers (no pandas.io.formats.type hints)
# ==============================================
def _color_scale(val, good_high=True, green_thresh=None, yellow_thresh=None):
    """Return CSS background for a numeric cell based on thresholds.
       - good_high: True means higher is better; False means lower is better
       - green_thresh, yellow_thresh: numbers or None
    """
    try:
        x = float(val)
    except Exception:
        return ""
    if green_thresh is None or yellow_thresh is None:
        return ""
    if good_high:
        if x >= green_thresh:
            return "background-color: #d9f7be"  # green-ish
        elif x >= yellow_thresh:
            return "background-color: #fff7ae"  # yellow-ish
        else:
            return "background-color: #ffd6d6"  # red-ish
    else:
        # lower is better
        if x <= green_thresh:
            return "background-color: #d9f7be"
        elif x <= yellow_thresh:
            return "background-color: #fff7ae"
        else:
            return "background-color: #ffd6d6"

def style_offense(df: pd.DataFrame):
    # Targets: YPA (high good), CMP% (high good), Explosive% (high good), Drive SR% (high good), Off Adj EPA/play (high good), Off Adj SR% (high good)
    cols = [c for c in df.columns if c in ["YPA","CMP%","Explosive%","Drive SR%","Off Adj EPA/play","Off Adj SR%"]]
    def _apply(s):
        if s.name == "YPA":
            return [_color_scale(v, True, 7.5, 6.0) for v in s]
        if s.name == "CMP%":
            return [_color_scale(v, True, 68, 60) for v in s]
        if s.name == "Explosive%":
            return [_color_scale(v, True, 12, 8) for v in s]
        if s.name == "Drive SR%":
            return [_color_scale(v, True, 42, 35) for v in s]
        if s.name == "Off Adj EPA/play":
            return [_color_scale(v, True, 0.15, 0.05) for v in s]
        if s.name == "Off Adj SR%":
            return [_color_scale(v, True, 3.0, 0.0) for v in s]  # in % points vs opp
        return ["" for _ in s]
    return df.style.apply(_apply, subset=cols)

def style_defense(df: pd.DataFrame):
    # Targets: RZ% Allowed (low good), Success Rate% (Offense) allowed (low good), Pressures (high good), Def Adj EPA/play (low good), Def Adj SR% (low good)
    cols = [c for c in df.columns if c in ["RZ% Allowed","Success Rate% (Offense)","Pressures","Def Adj EPA/play","Def Adj SR%"]]
    def _apply(s):
        if s.name == "RZ% Allowed":
            return [_color_scale(v, False, 50, 60) for v in s]
        if s.name == "Success Rate% (Offense)":
            return [_color_scale(v, False, 42, 48) for v in s]
        if s.name == "Pressures":
            return [_color_scale(v, True, 10, 7) for v in s]
        if s.name == "Def Adj EPA/play":
            return [_color_scale(v, False, -0.05, 0.0) for v in s]
        if s.name == "Def Adj SR%":
            return [_color_scale(v, False, -2.0, 0.0) for v in s]
        return ["" for _ in s]
    return df.style.apply(_apply, subset=cols)

def style_injuries(df: pd.DataFrame):
    # Color by GameStatus: Out/IR/PUP/Knee questionable
    def _row_styles(r):
        status = str(r.get("GameStatus","")).strip().lower()
        if any(k in status for k in ["out","injured reserve","pup","dnp"]):
            return ["background-color: #ffd6d6"] * len(r)  # red
        if any(k in status for k in ["questionable","limited"]):
            return ["background-color: #fff7ae"] * len(r)  # yellow
        if any(k in status for k in ["full","probable","active"]):
            return ["background-color: #d9f7be"] * len(r)  # green
        return [""] * len(r)
    return df.style.apply(lambda _df: [_row_styles(_df.iloc[i]) for i in range(len(_df))], axis=None)

# ==============================================
# Sidebar Uploads
# ==============================================
st.sidebar.header("📤 Upload New Weekly Data (.csv)")
uploaded_offense   = st.sidebar.file_uploader("Offense", type="csv")
uploaded_defense   = st.sidebar.file_uploader("Defense", type="csv")
uploaded_strategy  = st.sidebar.file_uploader("Strategy", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Personnel", type="csv")
uploaded_injuries  = st.sidebar.file_uploader("Injuries", type="csv")
uploaded_snaps     = st.sidebar.file_uploader("Snap Counts", type="csv")

if uploaded_offense:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense")
    st.sidebar.success("✅ Offense uploaded")

if uploaded_defense:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense")
    st.sidebar.success("✅ Defense uploaded")

if uploaded_strategy:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy")
    st.sidebar.success("✅ Strategy uploaded")

if uploaded_personnel:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel")
    st.sidebar.success("✅ Personnel uploaded")

if uploaded_injuries:
    df_inj = pd.read_csv(uploaded_injuries)
    append_to_excel(df_inj, "Injuries")
    st.sidebar.success("✅ Injuries uploaded")

if uploaded_snaps:
    df_snaps = pd.read_csv(uploaded_snaps)
    append_to_excel(df_snaps, "SnapCounts")
    st.sidebar.success("✅ Snap Counts uploaded")

# ==============================================
# Sidebar: Fetch blocks
# ==============================================
with st.sidebar.expander("⚡ Fetch Weekly Team Data (nfl_data_py)"):
    st.caption("Pulls 2025 team-level weekly stats (CHI) and saves to Excel (basic Offense/Defense).")
    fetch_week = st.number_input("Week (2025)", 1, 25, 1, 1)
    if st.button("Fetch CHI week (basic)"):
        try:
            import nfl_data_py as nfl
            # try to refresh caches (ok if it fails)
            try:
                nfl.update.weekly_data([2025])
            except Exception:
                pass
            weekly = nfl.import_weekly_data([2025])
            wk = int(fetch_week)
            team_week = weekly[(weekly["team"]=="CHI") & (weekly["week"]==wk)].copy()
            if team_week.empty:
                st.warning("No weekly row for CHI yet. Try again later.")
            else:
                # Offense approximation
                pass_yards = team_week["passing_yards"].iloc[0] if "passing_yards" in team_week.columns else None
                pass_att = None
                for cand in ["attempts","passing_attempts","pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].iloc[0]
                        break
                try:
                    ypa_val = float(pass_yards)/float(pass_att) if pass_yards is not None and pass_att not in (None,0) else None
                except Exception:
                    ypa_val = None
                yards_total = None
                for cand in ["yards","total_yards","offense_yards"]:
                    if cand in team_week.columns:
                        yards_total = team_week[cand].iloc[0]
                        break
                completions = None
                for cand in ["completions","passing_completions","pass_completions"]:
                    if cand in team_week.columns:
                        completions = team_week[cand].iloc[0]
                        break
                cmp_pct = None
                if completions is not None and pass_att not in (None,0):
                    try:
                        cmp_pct = round((float(completions)/float(pass_att))*100,1)
                    except Exception:
                        cmp_pct = None

                off_row = pd.DataFrame([{
                    "Week": wk, "YPA": round(ypa_val,2) if ypa_val is not None else None,
                    "YDS": yards_total, "CMP%": cmp_pct
                }])

                # Defense basics
                sacks_val = None
                for cand in ["sacks","defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].iloc[0]
                        break
                def_row = pd.DataFrame([{"Week": wk, "SACK": sacks_val, "RZ% Allowed": None}])

                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)
                st.success(f"✅ Saved week {wk} (basic Offense/Defense).")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

with st.sidebar.expander("📡 Fetch Defensive Metrics from PBP"):
    pbp_week = st.number_input("Week (2025 season)", 1, 25, 1, 1)
    if st.button("Fetch PBP Defensive Metrics"):
        try:
            import nfl_data_py as nfl
            pbp = nfl.import_pbp_data([2025], downcast=False)
            dfw = pbp[(pbp["week"]==int(pbp_week)) & (pbp["defteam"]=="CHI")].copy()
            if dfw.empty:
                st.warning("No CHI defensive PBP found for that week yet.")
            else:
                # Red Zone % Allowed (drives reaching <=20)
                dmins = (
                    dfw.groupby(["game_id","drive"], as_index=False)["yardline_100"]
                    .min().rename(columns={"yardline_100":"min_yardline_100"})
                )
                total_drives = len(dmins)
                rz = (len(dmins[dmins["min_yardline_100"] <= 20])/total_drives*100) if total_drives>0 else 0.0
                # Success Rate allowed to offense
                def success_flag(r):
                    d = r.get("down"); togo=r.get("ydstogo"); gain=r.get("yards_gained")
                    try:
                        if pd.isna(d) or pd.isna(togo) or pd.isna(gain):
                            return False
                        d = int(d); togo=float(togo); gain=float(gain)
                        if d==1: return gain>=0.4*togo
                        if d==2: return gain>=0.6*togo
                        return gain>=togo
                    except Exception:
                        return False
                plays_mask = (~dfw["play_type"].isin(["no_play"])) & (~dfw["penalty"].fillna(False))
                real_plays = dfw[plays_mask].copy()
                sr = real_plays.apply(success_flag, axis=1).mean()*100 if len(real_plays) else 0.0
                qb_hits = dfw["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in dfw.columns else 0
                sacks = dfw["sack"].fillna(0).astype(int).sum() if "sack" in dfw.columns else 0
                pressures = int(qb_hits + sacks)

                adv = pd.DataFrame([{
                    "Week": int(pbp_week),
                    "RZ% Allowed": round(rz,1),
                    "Success Rate% (Offense)": round(sr,1),
                    "Pressures": pressures
                }])
                append_to_excel(adv, "Advanced_Defense", deduplicate=True)
                st.success(f"✅ Saved Advanced Defense for week {int(pbp_week)}.")
        except Exception as e:
            st.error(f"❌ Failed: {e}")

with st.sidebar.expander("📈 Compute DVOA-like Proxy (Opponent-Adjusted)"):
    proxy_week = st.number_input("Week to compute", 1, 25, 1, 1, key="proxy_week")
    if st.button("Compute DVOA-like Proxy"):
        try:
            import nfl_data_py as nfl
            wk = int(proxy_week)
            pbp = nfl.import_pbp_data([2025], downcast=False)
            plays = pbp[
                (~pbp["play_type"].isin(["no_play"])) &
                (~pbp["penalty"].fillna(False)) &
                (~pbp["epa"].isna())
            ].copy()
            bears_off = plays[(plays["week"]==wk) & (plays["posteam"]=="CHI")].copy()
            bears_def = plays[(plays["week"]==wk) & (plays["defteam"]=="CHI")].copy()
            if bears_off.empty and bears_def.empty:
                st.warning("No CHI plays found for that week.")
            else:
                opps=set()
                if not bears_off.empty: opps.update(bears_off["defteam"].unique().tolist())
                if not bears_def.empty: opps.update(bears_def["posteam"].unique().tolist())
                opponent = list(opps)[0] if opps else "UNK"

                prior = plays[plays["week"]<wk].copy()
                # opponent defense (allowed)
                opp_def = prior[prior["defteam"]==opponent].copy()
                opp_def_epa = opp_def["epa"].mean() if len(opp_def) else None
                opp_def_sr  = opp_def.apply(lambda r: _success(r), axis=1).mean() if len(opp_def) else None
                # opponent offense
                opp_off = prior[prior["posteam"]==opponent].copy()
                opp_off_epa = opp_off["epa"].mean() if len(opp_off) else None
                opp_off_sr  = opp_off.apply(lambda r: _success(r), axis=1).mean() if len(opp_off) else None

                if len(bears_off):
                    chi_off_epa = bears_off["epa"].mean()
                    chi_off_sr  = bears_off.apply(lambda r: _success(r), axis=1).mean()
                else:
                    chi_off_epa=None; chi_off_sr=None
                if len(bears_def):
                    chi_def_epa = bears_def["epa"].mean()
                    chi_def_sr  = bears_def.apply(lambda r: _success(r), axis=1).mean()
                else:
                    chi_def_epa=None; chi_def_sr=None

                def safe_diff(a,b):
                    if a is None or pd.isna(a) or b is None or pd.isna(b): return None
                    return float(a)-float(b)

                off_adj_epa = safe_diff(chi_off_epa, opp_def_epa)
                off_adj_sr  = safe_diff(chi_off_sr,  opp_def_sr)
                def_adj_epa = safe_diff(chi_def_epa, opp_off_epa)
                def_adj_sr  = safe_diff(chi_def_sr,  opp_off_sr)

                out = pd.DataFrame([{
                    "Week": wk, "Opponent": opponent,
                    "Off Adj EPA/play": round(off_adj_epa,3) if off_adj_epa is not None else None,
                    "Off Adj SR%": round(off_adj_sr*100,1) if off_adj_sr is not None else None,
                    "Def Adj EPA/play": round(def_adj_epa,3) if def_adj_epa is not None else None,
                    "Def Adj SR%": round(def_adj_sr*100,1) if def_adj_sr is not None else None,
                    "Off EPA/play": round(chi_off_epa,3) if chi_off_epa is not None else None,
                    "Def EPA allowed/play": round(chi_def_epa,3) if chi_def_epa is not None else None
                }])
                append_to_excel(out, "DVOA_Proxy", deduplicate=True)
                st.success(f"✅ Saved DVOA-like proxy for Week {wk} vs {opponent}.")
        except Exception as e:
            st.error(f"❌ Failed to compute proxy: {e}")

def _success(row):
    try:
        d=row.get("down"); togo=row.get("ydstogo"); gain=row.get("yards_gained")
        if pd.isna(d) or pd.isna(togo) or pd.isna(gain): return False
        d=int(d); togo=float(togo); gain=float(gain)
        if d==1: return gain>=0.4*togo
        if d==2: return gain>=0.6*togo
        return gain>=togo
    except Exception:
        return False

# ==============================================
# Injuries quick entry (on main panel)
# ==============================================
st.markdown("### 🏥 Add an Injury (Quick Entry)")
with st.form("injury_quick_entry"):
    iq_week   = st.number_input("Week", 1, 25, 1)
    iq_player = st.text_input("Player")
    iq_pos    = st.text_input("Position")
    iq_injury = st.text_input("InjuryType")
    iq_status = st.text_input("GameStatus (Out/Questionable/Probable/IR/etc.)")
    iq_prac   = st.text_input("PracticeStatus (DNP/Limited/Full/etc.)")
    iq_notes  = st.text_area("Notes")
    iq_submit = st.form_submit_button("Save Injury")
if iq_submit:
    inj_row = pd.DataFrame([{
        "Week": iq_week, "Player": iq_player, "Position": iq_pos,
        "InjuryType": iq_injury, "GameStatus": iq_status,
        "PracticeStatus": iq_prac, "Notes": iq_notes
    }])
    append_to_excel(inj_row, "Injuries", deduplicate=True)
    st.success(f"✅ Saved injury for Week {iq_week}: {iq_player}")

# ==============================================
# Download All Data
# ==============================================
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button("⬇️ Download All Data (Excel)", data=f, file_name=EXCEL_FILE, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==============================================
# PREVIEWS (styled where applicable)
# ==============================================
def _read_sheet(name):
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=name)
    except Exception:
        return pd.DataFrame()

st.markdown("### 📊 Offensive Analytics")
df_off_prev = _read_sheet("Offense")
if not df_off_prev.empty:
    st.dataframe(style_offense(df_off_prev))
else:
    st.info("No offense data yet.")

st.markdown("### 🛡️ Defensive Analytics")
df_def_prev = _read_sheet("Defense")
if not df_def_prev.empty:
    # Merge in Advanced_Defense if available to show color columns together
    df_adv_prev = _read_sheet("Advanced_Defense")
    if not df_adv_prev.empty:
        try:
            merged = pd.merge(df_def_prev, df_adv_prev, on="Week", how="left")
            st.dataframe(style_defense(merged))
        except Exception:
            st.dataframe(df_def_prev)
    else:
        st.dataframe(df_def_prev)
else:
    st.info("No defense data yet.")

st.markdown("### 🧮 DVOA-like Proxy (Opponent-Adjusted)")
df_proxy_prev = _read_sheet("DVOA_Proxy")
if not df_proxy_prev.empty:
    st.dataframe(df_proxy_prev)
else:
    st.info("No DVOA-like proxy data yet.")

st.markdown("### 🧑‍🤝‍🧑 Personnel Usage")
df_pers_prev = _read_sheet("Personnel")
if not df_pers_prev.empty:
    st.dataframe(df_pers_prev)
else:
    st.info("No personnel data yet.")

st.markdown("### 📝 Weekly Strategy")
df_strat_prev = _read_sheet("Strategy")
if not df_strat_prev.empty:
    st.dataframe(df_strat_prev)
else:
    st.info("No strategy data yet.")

st.markdown("### 🏥 Injuries")
df_inj_prev = _read_sheet("Injuries")
if not df_inj_prev.empty:
    st.dataframe(style_injuries(df_inj_prev))
else:
    st.info("No injuries logged yet.")

st.markdown("### ⏱️ Snap Counts")
df_snap_prev = _read_sheet("SnapCounts")
if not df_snap_prev.empty:
    st.dataframe(df_snap_prev)
else:
    st.info("No snap counts yet.")

# ==============================================
# Weekly Game Prediction
# ==============================================
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", 1, 25, 1, 1, key="predict_week_input")

def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy = _read_sheet("Strategy")
        df_offense  = _read_sheet("Offense")
        df_defense  = _read_sheet("Defense")
        df_advdef   = _read_sheet("Advanced_Defense")
        df_proxy    = _read_sheet("DVOA_Proxy")

        row_s = df_strategy[df_strategy["Week"]==week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"]==week_to_predict]   if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"]==week_to_predict]   if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"]==week_to_predict]     if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"]==week_to_predict]       if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()
            ypa = _safe_float(row_o.iloc[0].get("YPA"), None)
            rz_allowed = None
            pressures = None
            if not row_a.empty:
                rz_allowed = _safe_float(row_a.iloc[0].get("RZ% Allowed"), None)
                pressures  = _safe_float(row_a.iloc[0].get("Pressures"), None)
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), None)

            reason_bits = []
            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win – efficiency edge on both sides"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play vs opp D")
                reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play vs opp O")
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone","two-high","split-safety"]):
                prediction = "Win – red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (off_adj_epa is not None and off_adj_epa <= -0.10) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient offense and poor red zone defense"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (ypa is not None and ypa < 6) and (rz_allowed is not None and rz_allowed > 65):
                prediction = "Loss – inefficient passing and weak red zone defense"
                reason_bits.append(f"YPA={ypa:.1f}")
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
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

            # save prediction
            prediction_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(prediction_entry, "Predictions", deduplicate=True)
        else:
            st.info("Please ensure Strategy, Offense, and Defense for that week are present.")
    except Exception as e:
        st.warning(f"Prediction failed. Error: {e}")

# Show saved predictions
df_preds = _read_sheet("Predictions")
if not df_preds.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(df_preds)
else:
    st.info("No predictions saved yet.")

# ==============================================
# Optional: Excel Sanity Checker
# ==============================================
with st.expander("🧪 Excel Sanity Checker (optional)"):
    cwd = os.getcwd()
    st.write(f"**Current Working Directory:** `{cwd}`")
    st.write(f"**Excel File Being Used:** `{EXCEL_FILE}`")
    if os.path.exists(EXCEL_FILE):
        st.success("Workbook exists.")
        try:
            xl = pd.ExcelFile(EXCEL_FILE)
            st.write("**Sheets:**", xl.sheet_names)
            for s in xl.sheet_names[:6]:
                try:
                    st.caption(f"Preview — {s}")
                    st.dataframe(pd.read_excel(xl, sheet_name=s).head(10))
                except Exception:
                    st.caption(f"Preview — {s} (unavailable)")
        except Exception as e:
            st.error(f"Failed to open workbook: {e}")
    else:
        st.warning("Workbook not found yet (it will be created after your first upload/fetch).")
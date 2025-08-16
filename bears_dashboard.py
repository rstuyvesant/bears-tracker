import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"
TEAM = "CHI"
SEASON = 2025
MAX_WEEKS = 25  # pre/post included safety

# =============================
# Helpers
# =============================

def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

def _dedup_by_week(df: pd.DataFrame) -> pd.DataFrame:
    """Keep the last row per Week if 'Week' exists; otherwise drop exact duplicates."""
    if "Week" in df.columns:
        # normalize week to int where possible
        def _to_int(v):
            try:
                return int(v)
            except Exception:
                return v
        df = df.copy()
        df["Week"] = df["Week"].apply(_to_int)
        return df.drop_duplicates(subset=["Week"], keep="last")
    return df.drop_duplicates(keep="last")

def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE):
    """
    Append/merge a DataFrame into an Excel sheet.
    - If sheet exists and has 'Week', we keep only the latest row per Week (dedupe).
    - Otherwise we drop exact duplicates.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    if new_data is None or new_data.empty:
        return

    # Ensure string columns for safety with openpyxl when mixed types
    new_data = new_data.copy()
    if "Week" in new_data.columns:
        # force week to int when possible
        try:
            new_data["Week"] = new_data["Week"].astype(int)
        except Exception:
            pass

    if os.path.exists(file_name):
        book = openpyxl.load_workbook(file_name)
        if sheet_name in book.sheetnames:
            sheet = book[sheet_name]
            existing = pd.DataFrame(sheet.values)
            if len(existing) > 0:
                existing.columns = existing.iloc[0]
                existing = existing[1:]
            else:
                existing = pd.DataFrame(columns=new_data.columns)

            # align cols
            combined = pd.concat([existing, new_data], ignore_index=True)
            combined = _dedup_by_week(combined)

            # wipe sheet and re-write
            book.remove(sheet)
            sheet = book.create_sheet(sheet_name)
            for r in dataframe_to_rows(combined, index=False, header=True):
                sheet.append(r)
        else:
            sheet = book.create_sheet(sheet_name)
            for r in dataframe_to_rows(new_data, index=False, header=True):
                sheet.append(r)
        book.save(file_name)
    else:
        with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
            new_data.to_excel(writer, sheet_name=sheet_name, index=False)

def read_sheet(sheet_name: str) -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

# =============================
# File Uploaders
# =============================

st.sidebar.header("📤 Upload Weekly CSVs")
uploaded_off = st.sidebar.file_uploader("Upload Offense CSV", type=["csv"])
uploaded_def = st.sidebar.file_uploader("Upload Defense CSV", type=["csv"])
uploaded_strat = st.sidebar.file_uploader("Upload Strategy CSV", type=["csv"])
uploaded_pers = st.sidebar.file_uploader("Upload Personnel CSV", type=["csv"])
uploaded_inj = st.sidebar.file_uploader("Upload Injuries CSV", type=["csv"])
uploaded_snap = st.sidebar.file_uploader("Upload Snap Counts CSV", type=["csv"])

if uploaded_off:
    df_off = pd.read_csv(uploaded_off)
    append_to_excel(df_off, "Offense")
    st.success("✅ Offense data saved.")

if uploaded_def:
    df_def = pd.read_csv(uploaded_def)
    append_to_excel(df_def, "Defense")
    st.success("✅ Defense data saved.")

if uploaded_strat:
    df_strat = pd.read_csv(uploaded_strat)
    append_to_excel(df_strat, "Strategy")
    st.success("✅ Strategy data saved.")

if uploaded_pers:
    df_pers = pd.read_csv(uploaded_pers)
    append_to_excel(df_pers, "Personnel")
    st.success("✅ Personnel data saved.")

if uploaded_inj:
    df_inj = pd.read_csv(uploaded_inj)
    append_to_excel(df_inj, "Injuries")
    st.success("✅ Injuries data saved.")

if uploaded_snap:
    df_snap = pd.read_csv(uploaded_snap)
    append_to_excel(df_snap, "SnapCounts")
    st.success("✅ Snap counts data saved.")

# =============================
# Quick Injury Entry
# =============================

st.header("🩺 Injuries — quick entry")
with st.form("injury_quick"):
    iw = st.number_input("Week", min_value=1, max_value=MAX_WEEKS, value=1, step=1)
    player = st.text_input("Player")
    status = st.selectbox("Status", ["Questionable", "Doubtful", "Out", "IR", "PUP", "Active"])
    body = st.text_input("Body Part / Note")
    source = st.text_input("Source (optional)")
    submitted = st.form_submit_button("Add Injury")
if submitted and player:
    row = pd.DataFrame([{
        "Week": iw, "Player": player, "Status": status, "BodyPart": body, "Source": source
    }])
    append_to_excel(row, "Injuries")
    st.success(f"✅ Added injury for Week {iw}: {player} — {status}")

# =============================
# Data Fetching (nfl_data_py)
# =============================

st.header("📡 Data Fetching")

week_to_fetch = st.number_input("Enter Week # to fetch", min_value=1, max_value=MAX_WEEKS, step=1)

# --- Weekly team basics (Offense/Defense) ---
if st.button("Fetch Weekly Data (nfl_data_py)"):
    try:
        import nfl_data_py as nfl

        # refresh cache (best effort)
        try: nfl.update.weekly_data([SEASON])
        except Exception: pass

        weekly = nfl.import_weekly_data([SEASON])
        wk = int(week_to_fetch)
        team_week = weekly[(weekly["team"] == TEAM) & (weekly["week"] == wk)].copy()

        if team_week.empty:
            st.warning("No weekly team row found yet for that week.")
        else:
            team_week = team_week.copy()
            team_week["Week"] = wk

            # --- Offense mapping (best effort) ---
            # Passing yards/attempts/completions
            pass_yards = None
            for c in ["passing_yards", "pass_yards", "yards"]:
                if c in team_week.columns:
                    pass_yards = team_week[c].values[0]; break
            pass_att = None
            for c in ["attempts", "passing_attempts", "pass_attempts", "pass_att"]:
                if c in team_week.columns:
                    pass_att = team_week[c].values[0]; break
            completions = None
            for c in ["completions", "passing_completions", "pass_completions", "pass_cmp"]:
                if c in team_week.columns:
                    completions = team_week[c].values[0]; break

            # YPA / CMP%
            try:
                ypa_val = (float(pass_yards) / float(pass_att)) if pass_yards is not None and pass_att not in (None, 0) else None
            except Exception:
                ypa_val = None
            cmp_pct = None
            if completions is not None and pass_att not in (None, 0):
                try:
                    cmp_pct = round((float(completions) / float(pass_att)) * 100, 1)
                except Exception:
                    cmp_pct = None

            yards_total = None
            for c in ["yards", "total_yards", "offense_yards"]:
                if c in team_week.columns:
                    yards_total = team_week[c].values[0]; break

            off_row = pd.DataFrame([{
                "Week": wk,
                "YPA": round(ypa_val, 2) if ypa_val is not None else None,
                "YDS": yards_total,
                "CMP%": cmp_pct
            }])

            # --- Defense basics (sacks) ---
            sacks_val = None
            for c in ["sacks", "defense_sacks", "def_sacks"]:
                if c in team_week.columns:
                    sacks_val = team_week[c].values[0]; break

            def_row = pd.DataFrame([{
                "Week": wk,
                "SACK": sacks_val,
                "RZ% Allowed": None  # will be filled by PBP fetch
            }])

            append_to_excel(off_row, "Offense")
            append_to_excel(def_row, "Defense")
            st.success(f"✅ Added Week {wk} offense/defense basics.")
    except Exception as e:
        st.error(f"❌ Fetch failed: {e}")

# --- Play-by-Play derived defense & offense advanced ---
if st.button("Fetch Defensive Metrics from Play-by-Play"):
    try:
        import nfl_data_py as nfl

        wk = int(week_to_fetch)
        pbp = nfl.import_pbp_data([SEASON], downcast=False)

        # Keep only real plays
        plays = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False))
        ].copy()

        # ----- DEFENSE (CHI on defense) -----
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == TEAM)].copy()
        if bears_def.empty:
            st.warning("No CHI defensive plays yet for that week.")
        else:
            # Red Zone % Allowed (drive-based: any play inside 20 = red zone drive)
            dmins = (
                bears_def.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min().rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success Rate allowed
            def _success_flag(down, togo, gain):
                try:
                    d = int(down); t = float(togo); g = float(gain)
                except Exception:
                    return False
                if pd.isna(down) or pd.isna(togo) or pd.isna(gain):
                    return False
                if d == 1:
                    return g >= 0.4 * t
                elif d == 2:
                    return g >= 0.6 * t
                return g >= t

            def_sr = bears_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            def_sr_pct = round(def_sr * 100, 1) if pd.notna(def_sr) else None

            # Pressures approx = sacks + QB hits
            qb_hits = bears_def["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in bears_def.columns else 0
            sacks = bears_def["sack"].fillna(0).astype(int).sum() if "sack" in bears_def.columns else 0
            pressures = int(qb_hits + sacks)

            # Explosive plays allowed: rush >=10 OR pass >=20
            def is_explosive(row):
                yds = _safe_float(row.get("yards_gained"), 0)
                is_pass = row.get("pass") == 1 or row.get("pass_attempt") == 1
                is_rush = row.get("rush") == 1 or row.get("rush_attempt") == 1
                if is_pass and yds is not None:
                    return yds >= 20
                if is_rush and yds is not None:
                    return yds >= 10
                return False
            def_expl_pct = round(bears_def.apply(is_explosive, axis=1).mean() * 100, 1) if len(bears_def) else 0.0

            # Red Zone TD% allowed: plays inside 20 that resulted in touchdown
            rz_def = bears_def[bears_def["yardline_100"] <= 20]
            if len(rz_def):
                rz_td_allowed_pct = round((rz_def["touchdown"].fillna(0).astype(int).sum() / len(rz_def)) * 100, 1)
            else:
                rz_td_allowed_pct = 0.0

            adv_def = pd.DataFrame([{
                "Week": wk,
                "RZ% Allowed (Drives)": round(rz_allowed, 1),
                "Success Rate% Allowed": def_sr_pct,
                "Pressures": pressures,
                "Explosive% Allowed": def_expl_pct,
                "RZ TD% Allowed (Plays)": rz_td_allowed_pct
            }])
            append_to_excel(adv_def, "Advanced_Defense")

        # ----- OFFENSE (CHI on offense) -----
        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == TEAM)].copy()
        if bears_off.empty:
            st.info("No CHI offensive plays yet for that week.")
        else:
            # Success Rate
            def _success_flag_off(down, togo, gain):
                try:
                    d = int(down); t = float(togo); g = float(gain)
                except Exception:
                    return False
                if pd.isna(down) or pd.isna(togo) or pd.isna(gain):
                    return False
                if d == 1:
                    return g >= 0.4 * t
                elif d == 2:
                    return g >= 0.6 * t
                return g >= t
            off_sr = bears_off.apply(lambda r: _success_flag_off(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            off_sr_pct = round(off_sr * 100, 1) if pd.notna(off_sr) else None

            # Explosive made
            def is_explosive_off(row):
                yds = _safe_float(row.get("yards_gained"), 0)
                is_pass = row.get("pass") == 1 or row.get("pass_attempt") == 1
                is_rush = row.get("rush") == 1 or row.get("rush_attempt") == 1
                if is_pass and yds is not None:
                    return yds >= 20
                if is_rush and yds is not None:
                    return yds >= 10
                return False
            off_expl_pct = round(bears_off.apply(is_explosive_off, axis=1).mean() * 100, 1) if len(bears_off) else 0.0

            # Red Zone TD% scored (play-based)
            rz_off = bears_off[bears_off["yardline_100"] <= 20]
            if len(rz_off):
                rz_td_scored_pct = round((rz_off["touchdown"].fillna(0).astype(int).sum() / len(rz_off)) * 100, 1)
            else:
                rz_td_scored_pct = 0.0

            adv_off = pd.DataFrame([{
                "Week": wk,
                "Success Rate%": off_sr_pct,
                "Explosive%": off_expl_pct,
                "RZ TD%": rz_td_scored_pct
            }])
            append_to_excel(adv_off, "Advanced_Offense")

        st.success(f"✅ Week {wk} PBP advanced metrics saved (offense & defense).")

    except Exception as e:
        st.error(f"❌ Failed to fetch PBP metrics: {e}")

# --- Snap Counts (placeholder stub) ---
if st.button("Fetch Snap Counts (best effort)"):
    st.info(f"Fetching snap counts for week {week_to_fetch}… (placeholder)")
    # You can later map an API here and then append_to_excel(df, "SnapCounts")

# --- Opponent Preview (placeholder fetch) ---
if st.button("Fetch Opponent Preview (best effort)"):
    st.info(f"Fetching opponent preview for week {week_to_fetch}… (placeholder)")
    # Later: pull basic opp splits, last-3 trend, injuries, etc., and save to "Opponent_Preview"

# =============================
# DVOA-like Proxy (opponent-adjusted)
# =============================

st.header("📊 DVOA-like Proxy")
week_for_proxy = st.number_input("Week to compute Proxy", min_value=1, max_value=MAX_WEEKS, step=1)

if st.button("Compute DVOA-like Proxy (Opponent Adjusted)"):
    try:
        import nfl_data_py as nfl

        wk = int(week_for_proxy)
        pbp = nfl.import_pbp_data([SEASON], downcast=False)

        base = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False)) &
            (~pbp["epa"].isna())
        ].copy()

        bears_off = base[(base["week"] == wk) & (base["posteam"] == TEAM)].copy()
        bears_def = base[(base["week"] == wk) & (base["defteam"] == TEAM)].copy()

        if bears_off.empty and bears_def.empty:
            st.warning("No CHI plays found for that week yet.")
        else:
            # figure opponent
            opps = set()
            if not bears_off.empty: opps.update(bears_off["defteam"].unique().tolist())
            if not bears_def.empty: opps.update(bears_def["posteam"].unique().tolist())
            opponent = list(opps)[0] if opps else "UNK"

            prior = base[base["week"] < wk].copy()

            # opponent defensive benchmarks (allowed vs them)
            opp_def = prior[prior["defteam"] == opponent]
            opp_def_epa = opp_def["epa"].mean() if len(opp_def) else None

            def _success_flag(down, ydstogo, yards_gained):
                try:
                    d = int(down); t = float(ydstogo); g = float(yards_gained)
                except Exception:
                    return False
                if pd.isna(down) or pd.isna(ydstogo) or pd.isna(yards_gained):
                    return False
                if d == 1: return g >= 0.4 * t
                if d == 2: return g >= 0.6 * t
                return g >= t

            opp_def_sr = opp_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_def) else None

            # opponent offensive benchmarks (their offense)
            opp_off = prior[prior["posteam"] == opponent]
            opp_off_epa = opp_off["epa"].mean() if len(opp_off) else None
            opp_off_sr = opp_off.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_off) else None

            # CHI week epa/sr
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
            off_adj_sr  = diff(chi_off_sr, opp_def_sr)
            def_adj_epa = diff(chi_def_epa_allowed, opp_off_epa)
            def_adj_sr  = diff(chi_def_sr_allowed, opp_off_sr)

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
            append_to_excel(out, "DVOA_Proxy")
            st.success(f"✅ DVOA-like proxy saved for Week {wk} vs {opponent}")
            try:
                st.dataframe(out.style.background_gradient(cmap="RdYlGn"))
            except Exception:
                st.dataframe(out)
    except Exception as e:
        st.error(f"Error computing DVOA Proxy: {e}")

# =============================
# Opponent Preview
# =============================

st.header("🔎 Opponent Preview")

try:
    strat = read_sheet("Strategy")
    if not strat.empty and "Week" in strat and "Opponent" in strat:
        wk_prev = st.number_input("Select week to preview opponent", min_value=1, max_value=MAX_WEEKS, step=1)
        row = strat[strat["Week"] == wk_prev]
        if not row.empty:
            opp = str(row.iloc[0]["Opponent"])
            st.subheader(f"Opponent: {opp}")
            # Injuries
            inj = read_sheet("Injuries")
            inj_wk = inj[inj["Week"] == wk_prev] if not inj.empty else pd.DataFrame()
            st.write(f"**Injury count (our team): {len(inj_wk)}**")
            # Trend placeholders
            st.write("📈 Efficiency trend (TBD)")
        else:
            st.info("No opponent info for that week yet.")
    else:
        st.info("No Strategy sheet with Week/Opponent yet.")
except Exception as e:
    st.error(f"Error loading opponent preview: {e}")

# =============================
# Weekly Prediction
# =============================

st.header("🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=MAX_WEEKS, step=1)

if os.path.exists(EXCEL_FILE):
    try:
        # base sheets
        df_strategy = read_sheet("Strategy")
        df_offense = read_sheet("Offense")
        df_defense = read_sheet("Defense")
        # advanced
        df_aoff = read_sheet("Advanced_Offense")
        df_adef = read_sheet("Advanced_Defense")
        df_proxy = read_sheet("DVOA_Proxy")

        row_s = df_strategy[df_strategy["Week"] == week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"] == week_to_predict] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"] == week_to_predict] if not df_defense.empty else pd.DataFrame()
        row_aoff = df_aoff[df_aoff["Week"] == week_to_predict] if not df_aoff.empty else pd.DataFrame()
        row_adef = df_adef[df_adef["Week"] == week_to_predict] if not df_adef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"] == week_to_predict] if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            # pulls
            ypa = _safe_float(row_o.iloc[0].get("YPA"), None)

            rz_allowed = None
            pressures = None
            def_sr_allowed = None
            expl_allowed = None
            rz_td_allowed = None

            if not row_adef.empty:
                rz_allowed = _safe_float(row_adef.iloc[0].get("RZ% Allowed (Drives)"), None)
                pressures = _safe_float(row_adef.iloc[0].get("Pressures"), None)
                def_sr_allowed = _safe_float(row_adef.iloc[0].get("Success Rate% Allowed"), None)
                expl_allowed = _safe_float(row_adef.iloc[0].get("Explosive% Allowed"), None)
                rz_td_allowed = _safe_float(row_adef.iloc[0].get("RZ TD% Allowed (Plays)"), None)
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), None)

            # Rule set
            reason_bits = []

            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win – efficiency edge on both sides"
                reason_bits.append(f"Off+{off_adj_epa:+.2f} EPA/play")
                reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play")
            elif (pressures is not None and pressures >= 8) and ("blitz" in strategy_text or "pressure" in strategy_text):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None: reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone", "two-high", "split-safety"]):
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
                prediction = "Loss – no clear advantage"
                if off_adj_epa is not None: reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play")
                if def_adj_epa is not None: reason_bits.append(f"Def{def_adj_epa:+.2f} EPA/play")
                if pressures is not None: reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None: reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
                if def_sr_allowed is not None: reason_bits.append(f"Def SR%={def_sr_allowed:.1f}")
                if expl_allowed is not None: reason_bits.append(f"Explosive% Allowed={expl_allowed:.1f}")
                if rz_td_allowed is not None: reason_bits.append(f"RZ TD% Allowed={rz_td_allowed:.1f}")

            reason_text = " | ".join(reason_bits)
            st.success(f"**Predicted Outcome for Week {week_to_predict}: {prediction}**")
            if reason_text:
                st.caption(reason_text)

            # Save prediction
            pred_row = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(pred_row, "Predictions")
        else:
            st.info("Please upload/fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check data. Error: {e}")

# =============================
# Preview Panels
# =============================

st.header("👀 Quick Previews")
off = read_sheet("Offense")
defn = read_sheet("Defense")
aoff = read_sheet("Advanced_Offense")
adef = read_sheet("Advanced_Defense")
dvoa = read_sheet("DVOA_Proxy")
preds = read_sheet("Predictions")
inj = read_sheet("Injuries")
pers = read_sheet("Personnel")

if not off.empty:
    st.subheader("Offense")
    try:
        st.dataframe(off.style.background_gradient(subset=["YPA","CMP%"], cmap="RdYlGn"))
    except Exception:
        st.dataframe(off)

if not defn.empty:
    st.subheader("Defense")
    try:
        st.dataframe(defn.style.background_gradient(subset=["SACK","RZ% Allowed"], cmap="RdYlGn_r"))
    except Exception:
        st.dataframe(defn)

if not aoff.empty:
    st.subheader("Advanced Offense (from PBP)")
    try:
        st.dataframe(aoff.style.background_gradient(subset=["Success Rate%","Explosive%","RZ TD%"], cmap="RdYlGn"))
    except Exception:
        st.dataframe(aoff)

if not adef.empty:
    st.subheader("Advanced Defense (from PBP)")
    try:
        st.dataframe(adef.style.background_gradient(subset=["Success Rate% Allowed","Explosive% Allowed","RZ TD% Allowed (Plays)"], cmap="RdYlGn_r"))
    except Exception:
        st.dataframe(adef)

if not dvoa.empty:
    st.subheader("DVOA-like Proxy")
    try:
        st.dataframe(dvoa.style.background_gradient(cmap="RdYlGn"))
    except Exception:
        st.dataframe(dvoa)

if not preds.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(preds)

if not inj.empty:
    st.subheader("Injuries")
    st.dataframe(inj)

if not pers.empty:
    st.subheader("Personnel")
    st.dataframe(pers)

# =============================
# Download
# =============================

st.sidebar.header("📥 Download Data")
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

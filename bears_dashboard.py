import os
import streamlit as st
import pandas as pd
from fpdf import FPDF

# -------------------------
# Basic page setup
# -------------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, snap counts, opponent previews, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"
TEAM = "CHI"
SEASON = 2025

# -------------------------
# Helpers
# -------------------------
def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = True):
    """Append (or replace by week) rows to a sheet. Dedupes on Week if present."""
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    try:
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
            if sheet_name in book.sheetnames:
                sheet = book[sheet_name]
                existing = pd.DataFrame(sheet.values)
                if not existing.empty:
                    existing.columns = existing.iloc[0]
                    existing = existing[1:]
                else:
                    existing = pd.DataFrame()

                if deduplicate and not existing.empty and "Week" in existing.columns and "Week" in new_data.columns:
                    wk = str(new_data.iloc[0]["Week"])
                    existing = existing[existing["Week"].astype(str) != wk]

                combined = pd.concat([existing, new_data], ignore_index=True)
            else:
                combined = new_data
        else:
            book = openpyxl.Workbook()
            book.remove(book.active)
            combined = new_data

        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        for r in dataframe_to_rows(combined, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def safe_read_excel(sheet_name: str) -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def _success_flag(down, ydstogo, yards_gained):
    try:
        if pd.isna(down) or pd.isna(ydstogo) or pd.isna(yards_gained):
            return False
        d = int(down); togo = float(ydstogo); gain = float(yards_gained)
        if d == 1:   return gain >= 0.4 * togo
        if d == 2:   return gain >= 0.6 * togo
        return gain >= togo
    except Exception:
        return False

def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

# -------------------------
# Sidebar: CSV uploads
# -------------------------
st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense   = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy  = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")
uploaded_injuries  = st.sidebar.file_uploader("Upload Injuries (.csv)", type="csv")
uploaded_snaps     = st.sidebar.file_uploader("Upload Snap Counts (.csv)", type="csv")

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
    append_to_excel(df_inj, "Injuries", deduplicate=False)  # multiple rows per week ok
    st.sidebar.success("✅ Injuries data uploaded.")

if uploaded_snaps:
    df_snap = pd.read_csv(uploaded_snaps)
    append_to_excel(df_snap, "Snap_Counts", deduplicate=False)  # many rows
    st.sidebar.success("✅ Snap counts data uploaded.")

# -------------------------
# Sidebar: Fetch weekly team data (nfl_data_py)
# -------------------------
with st.sidebar.expander("⚡ Fetch Weekly Team Data (nfl_data_py)"):
    st.caption("Pulls weekly team stats for CHI and saves to Excel.")
    fetch_week = st.number_input("Week to fetch (season 2025)", min_value=1, max_value=25, value=1, step=1, key="fetch_week_2025")
    auto_fill_missing = st.checkbox("Auto-fetch missing weeks up to selected week", value=False)

    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl
            try:
                nfl.update.weekly_data([SEASON])
            except Exception:
                pass

            weekly = nfl.import_weekly_data([SEASON])
            weeks_to_do = [int(fetch_week)]
            if auto_fill_missing:
                have = safe_read_excel("Offense")
                have_weeks = set()
                if not have.empty and "Week" in have.columns:
                    have_weeks = set(pd.to_numeric(have["Week"], errors="coerce").dropna().astype(int).tolist())
                weeks_to_do = [w for w in range(1, int(fetch_week) + 1) if w not in have_weeks] or [int(fetch_week)]

            added = 0
            for wk in weeks_to_do:
                team_week = weekly[(weekly["team"] == TEAM) & (weekly["week"] == wk)].copy()
                if team_week.empty:
                    continue
                team_week["Week"] = wk

                # Best-effort offense fields
                pass_yards = team_week["passing_yards"].values[0] if "passing_yards" in team_week.columns else None
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
                    "YPA": round(ypa_val, 2) if ypa_val is not None else None,
                    "YDS": yards_total,
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
                    "RZ% Allowed": None  # filled later from PBP
                }])

                append_to_excel(off_row, "Offense", deduplicate=True)
                append_to_excel(def_row, "Defense", deduplicate=True)
                append_to_excel(team_week.rename(columns=str), "Raw_Weekly", deduplicate=False)
                added += 1

            if added:
                st.success(f"✅ Added/updated {added} week(s) for CHI.")
            else:
                st.info("No rows available to add for the selected range.")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

# -------------------------
# Sidebar: Fetch PBP-derived defensive metrics
# -------------------------
st.sidebar.markdown("### 📡 Fetch Defensive Metrics (Play-by-Play)")
pbp_week = st.sidebar.number_input("Week to Fetch (2025)", min_value=1, max_value=25, value=1, step=1, key="pbp_week_2025")
if st.sidebar.button("Fetch Play-by-Play Metrics"):
    try:
        import nfl_data_py as nfl
        pbp = nfl.import_pbp_data([SEASON], downcast=False)
        pbp_w = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == TEAM)].copy()

        if pbp_w.empty:
            st.warning("No PBP rows for CHI defense in that week yet.")
        else:
            # Red zone drives allowed (% of drives that reached ≤20 yd line)
            dmins = (
                pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success rate: offense success vs CHI defense
            plays_mask = (~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))
            pbp_real = pbp_w[plays_mask].copy()
            success_rate = pbp_real.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() * 100 if len(pbp_real) else 0.0

            # Pressures approximation: hits + sacks
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
                f"✅ Week {int(pbp_week)} PBP metrics — RZ% Allowed: {rz_allowed:.1f} | "
                f"Success Rate% (Off): {success_rate:.1f} | Pressures: {pressures}"
            )
    except Exception as e:
        st.error(f"❌ Failed to fetch metrics: {e}")

# -------------------------
# Injuries + Snap Counts (upload or fetch)
# -------------------------
with st.sidebar.expander("🩹 Injuries & ⏱️ Snap Counts"):
    inj_week = st.number_input("Week (injuries/snap counts)", min_value=1, max_value=25, value=1, step=1, key="inj_snap_week")
    colA, colB = st.columns(2)
    with colA:
        fetch_inj = st.button("Fetch Injuries (best-effort)")
    with colB:
        fetch_snaps = st.button("Fetch Snap Counts (best-effort)")

    if fetch_inj:
        try:
            # Not all seasons have injuries in nfl_data_py. Attempt, then fallback.
            import nfl_data_py as nfl
            try:
                inj = nfl.import_injury_reports([SEASON])  # if available in your version
                inj_w = inj[(inj["team"] == TEAM) & (inj["week"] == int(inj_week))].copy()
            except Exception:
                inj_w = pd.DataFrame()

            if inj_w.empty:
                st.info("No injuries available via nfl_data_py for this week/version. Use CSV upload or the on-page form.")
            else:
                inj_w = inj_w.rename(columns=str)
                if "Week" not in inj_w.columns:
                    inj_w["Week"] = int(inj_week)
                append_to_excel(inj_w, "Injuries", deduplicate=False)
                st.success(f"✅ Injuries saved for Week {inj_week}.")
        except Exception as e:
            st.error(f"❌ Injury fetch failed: {e}")

    if fetch_snaps:
        try:
            import nfl_data_py as nfl
            try:
                snaps = nfl.import_snap_counts([SEASON])  # if available in your version
                snaps_w = snaps[(snaps["team"] == TEAM) & (snaps["week"] == int(inj_week))].copy()
            except Exception:
                snaps_w = pd.DataFrame()

            if snaps_w.empty:
                st.info("No snap counts available via nfl_data_py (this season/version). Use CSV upload instead.")
            else:
                snaps_w = snaps_w.rename(columns=str)
                if "Week" not in snaps_w.columns:
                    snaps_w["Week"] = int(inj_week)
                append_to_excel(snaps_w, "Snap_Counts", deduplicate=False)
                st.success(f"✅ Snap counts saved for Week {inj_week}.")
        except Exception as e:
            st.error(f"❌ Snap counts fetch failed: {e}")

# On-page quick entry for injuries (optional)
st.markdown("### 🩹 Add an Injury (Quick Entry)")
with st.form("injury_form"):
    iw = st.number_input("Week", min_value=1, max_value=25, value=1)
    iplayer = st.text_input("Player")
    ipos = st.text_input("Position")
    istatus = st.selectbox("Status", ["Questionable", "Doubtful", "Out", "IR", "Active"], index=0)
    inotes = st.text_area("Notes")
    if st.form_submit_button("Save Injury"):
        if iplayer.strip():
            inj_row = pd.DataFrame([{
                "Week": int(iw),
                "Player": iplayer.strip(),
                "Position": ipos.strip(),
                "Status": istatus,
                "Notes": inotes.strip()
            }])
            append_to_excel(inj_row, "Injuries", deduplicate=False)
            st.success(f"✅ Injury saved for Week {iw}.")
        else:
            st.warning("Player name is required.")

# -------------------------
# Opponent Preview (best-effort)
# -------------------------
st.markdown("### 🧭 Opponent Preview")
opp_week = st.number_input("Week to preview", min_value=1, max_value=25, value=1, step=1, key="opp_prev_week")
if st.button("Build Opponent Preview"):
    try:
        import nfl_data_py as nfl
        sched = nfl.import_schedules([SEASON])
        g = sched[(sched["week"] == int(opp_week)) & ((sched["home_team"] == TEAM) | (sched["away_team"] == TEAM))].copy()
        if g.empty:
            st.info("Schedule not found for that week (yet).")
        else:
            row = g.iloc[0]
            opp = row["away_team"] if row["home_team"] == TEAM else row["home_team"]
            st.write(f"**Week {int(opp_week)} Opponent:** {opp}")

            # Season-to-date opponent context up to prior week
            try:
                weekly = nfl.import_weekly_data([SEASON])
                prior = weekly[weekly["week"] < int(opp_week)].copy()

                # Opponent offense (their team rows)
                opp_off = prior[prior["team"] == opp].copy()
                # Opponent defense (we approximate by looking at rows where opponent was opponent, but weekly is team-centric.
                # For quick context, just show their offense means.)
                summary = {}
                for col in ["passing_yards", "rushing_yards", "sacks", "points"]:
                    if col in opp_off.columns and len(opp_off):
                        summary[col] = round(opp_off[col].mean(), 1)
                if summary:
                    st.write("**Opponent Offense (season-to-date averages before this week):**")
                    st.dataframe(pd.DataFrame([summary]))
                else:
                    st.info("No simple opponent stats available in this environment.")
            except Exception:
                st.info("Weekly dataset unavailable for preview.")
    except Exception as e:
        st.error(f"❌ Opponent preview failed: {e}")

# -------------------------
# Download full Excel
# -------------------------
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------------
# Main grid: show current uploaded/fetched tables
# -------------------------
df_offense   = safe_read_excel("Offense")
df_defense   = safe_read_excel("Defense")
df_strategy  = safe_read_excel("Strategy")
df_personnel = safe_read_excel("Personnel")
df_advdef    = safe_read_excel("Advanced_Defense")
df_inj       = safe_read_excel("Injuries")
df_snaps     = safe_read_excel("Snap_Counts")
df_proxy     = safe_read_excel("DVOA_Proxy")
df_preds     = safe_read_excel("Predictions")

if not df_offense.empty:
    st.subheader("📊 Offensive Analytics")
    st.dataframe(df_offense)
if not df_defense.empty:
    st.subheader("🛡️ Defensive Analytics")
    st.dataframe(df_defense)
if not df_strategy.empty:
    st.subheader("📘 Weekly Strategy")
    st.dataframe(df_strategy)
if not df_personnel.empty:
    st.subheader("👥 Personnel Usage")
    st.dataframe(df_personnel)
if not df_inj.empty:
    st.subheader("🩹 Injuries")
    st.dataframe(df_inj)
if not df_snaps.empty:
    st.subheader("⏱️ Snap Counts")
    st.dataframe(df_snaps)
if not df_advdef.empty:
    st.subheader("📡 Advanced Defense (PBP)")
    st.dataframe(df_advdef)
if not df_proxy.empty:
    st.subheader("📈 DVOA-Like Proxy")
    st.dataframe(df_proxy)
if not df_preds.empty:
    st.subheader("🧮 Saved Predictions")
    st.dataframe(df_preds)

# -------------------------
# DVOA-like proxy (opponent-adjusted EPA/SR)
# -------------------------
st.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.number_input("Week to Compute (2025 Season)", min_value=1, max_value=25, value=1, step=1, key="proxy_week_2025")

if st.button("Compute DVOA-like Proxy"):
    try:
        import nfl_data_py as nfl
        wk = int(proxy_week)
        pbp = nfl.import_pbp_data([SEASON], downcast=False)

        plays = pbp[
            (~pbp["play_type"].isin(["no_play"])) &
            (~pbp["penalty"].fillna(False)) &
            (~pbp["epa"].isna())
        ].copy()

        bears_off = plays[(plays["week"] == wk) & (plays["posteam"] == TEAM)].copy()
        bears_def = plays[(plays["week"] == wk) & (plays["defteam"] == TEAM)].copy()

        if bears_off.empty and bears_def.empty:
            st.warning("No CHI plays found for that week yet. Try again later.")
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

            append_to_excel(out, "DVOA_Proxy", deduplicate=True)
            st.success(
                f"✅ DVOA-like proxy saved for Week {wk} vs {opponent} "
                f"(Off Adj EPA/play={out.iloc[0]['Off Adj EPA/play']}, Off Adj SR%={out.iloc[0]['Off Adj SR%']}, "
                f"Def Adj EPA/play={out.iloc[0]['Def Adj EPA/play']}, Def Adj SR%={out.iloc[0]['Def Adj SR%']})"
            )
    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# -------------------------
# Weekly Prediction
# -------------------------
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy = safe_read_excel("Strategy")
        df_offense  = safe_read_excel("Offense")
        df_defense  = safe_read_excel("Defense")
        df_advdef   = safe_read_excel("Advanced_Defense")
        df_proxy    = safe_read_excel("DVOA_Proxy")

        row_s = df_strategy[df_strategy["Week"] == week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"] == week_to_predict] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"] == week_to_predict] if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"] == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"] == week_to_predict] if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            rz_allowed = None
            pressures = None
            if not row_a.empty:
                rz_allowed = _safe_float(row_a.iloc[0].get("RZ% Allowed"), default=None)
                pressures  = _safe_float(row_a.iloc[0].get("Pressures"), default=None)
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), default=None)

            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), default=None)

            reason_bits = []

            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win – efficiency edge on both sides"
                reason_bits += [f"Off{off_adj_epa:+.2f} EPA/play vs opp D", f"Def{def_adj_epa:+.2f} EPA/play vs opp O"]
            elif (pressures is not None and pressures >= 8) and any(tok in strategy_text for tok in ["blitz","pressure"]):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif (rz_allowed is not None and rz_allowed < 50) and any(tok in strategy_text for tok in ["zone","two-high","split-safety"]):
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

            prediction_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(prediction_entry, "Predictions", deduplicate=True)
        else:
            st.info("Please upload or fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# -------------------------
# PDF Report (unchanged core, optional to keep using)
# -------------------------
st.markdown("### 🧾 Download Weekly Game Report (PDF)")
report_week = st.number_input("Select Week for Report", min_value=1, max_value=25, step=1, key="report_week")

if st.button("Generate Weekly Report"):
    try:
        df_strategy = safe_read_excel("Strategy")
        df_offense  = safe_read_excel("Offense")
        df_defense  = safe_read_excel("Defense")
        df_media    = safe_read_excel("Media_Summaries")
        df_preds    = safe_read_excel("Predictions")

        strat_row = df_strategy[df_strategy["Week"] == report_week]
        off_row   = df_offense[df_offense["Week"] == report_week]
        def_row   = df_defense[df_defense["Week"] == report_week]
        media_rows= df_media[df_media["Week"] == report_week] if not df_media.empty else pd.DataFrame()
        pred_row  = df_preds[df_preds["Week"] == report_week] if not df_preds.empty else pd.DataFrame()

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, f"Chicago Bears Weekly Report – Week {report_week}", ln=True)
        pdf.set_font("Arial", "", 12)

        if not pred_row.empty:
            outcome = pred_row.iloc[0]["Prediction"]
            reason  = pred_row.iloc[0].get("Reason","")
            pdf.multi_cell(0, 10, f"🔮 Prediction: {outcome}\n📝 Reason: {reason}\n")

        if not strat_row.empty:
            pdf.set_font("Arial", "B", 12); pdf.cell(0, 10, "📘 Strategy Notes:", ln=True)
            pdf.set_font("Arial", "", 12)
            strategy_text = strat_row.iloc[0].astype(str).str.cat(sep=" | ")
            pdf.multi_cell(0, 10, strategy_text)

        if not off_row.empty:
            pdf.set_font("Arial", "B", 12); pdf.cell(0, 10, "📊 Offensive Analytics:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in off_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        if not def_row.empty:
            pdf.set_font("Arial", "B", 12); pdf.cell(0, 10, "🛡️ Defensive Analytics:", ln=True)
            pdf.set_font("Arial", "", 12)
            for col, val in def_row.iloc[0].items():
                pdf.cell(0, 8, f"{col}: {val}", ln=True)

        if not media_rows.empty:
            pdf.set_font("Arial", "B", 12); pdf.cell(0, 10, "📰 Media Summaries:", ln=True)
            pdf.set_font("Arial", "", 12)
            for _, row in media_rows.iterrows():
                source = row.get("Opponent", "Source")
                summary = row.get("Summary", "")
                pdf.multi_cell(0, 10, f"{source}:\n{summary}\n")

        pdf_output = f"week_{report_week}_report.pdf"
        pdf.output(pdf_output)
        with open(pdf_output, "rb") as f:
            st.download_button(
                label=f"📥 Download Week {report_week} Report (PDF)",
                data=f,
                file_name=pdf_output,
                mime="application/pdf"
            )
    except Exception as e:
        st.error(f"❌ Failed to generate PDF. {e}")
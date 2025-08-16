import os
import streamlit as st
import pandas as pd
from fpdf import FPDF

# -------------------- App Setup --------------------
st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, injuries, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# -------------------- Helpers --------------------
def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

def _read_sheet(sheet_name: str) -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def _append_to_excel(new_df: pd.DataFrame, sheet_name: str, key_cols=None, deduplicate=True):
    """
    Append (or replace duplicates) into `sheet_name`.
    If key_cols is provided, rows with identical key(s) are replaced by latest.
    Fallback key is ["Week"] if present.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    if new_df is None or len(new_df) == 0:
        return

    # Normalize columns
    new_df = new_df.copy()
    new_df.columns = [str(c).strip() for c in new_df.columns]

    # Load or create workbook
    if os.path.exists(EXCEL_FILE):
        book = openpyxl.load_workbook(EXCEL_FILE)
    else:
        book = openpyxl.Workbook()
        if "Sheet" in book.sheetnames:
            std = book["Sheet"]
            book.remove(std)

    # Read existing from the target sheet
    if sheet_name in book.sheetnames:
        ws = book[sheet_name]
        existing = pd.DataFrame(ws.values)
        if not existing.empty:
            existing.columns = existing.iloc[0]
            existing = existing[1:]
        else:
            existing = pd.DataFrame(columns=new_df.columns)
    else:
        existing = pd.DataFrame(columns=new_df.columns)

    # Harmonize columns
    for c in new_df.columns:
        if c not in existing.columns:
            existing[c] = pd.NA
    for c in existing.columns:
        if c not in new_df.columns:
            new_df[c] = pd.NA
    existing = existing[new_df.columns]

    # Dedup by key(s)
    if deduplicate:
        if key_cols is None:
            key_cols = ["Week"] if "Week" in new_df.columns else None
        if key_cols:
            # Drop rows from existing where keys match incoming keys
            # Build a mask for all rows in existing that are NOT present in new_df keys
            merge_keys = key_cols if isinstance(key_cols, list) else [key_cols]
            new_keys = new_df[merge_keys].astype(str).apply("|".join, axis=1).unique().tolist()
            if len(existing):
                ex_keys = existing[merge_keys].astype(str).apply("|".join, axis=1)
                existing = existing[~ex_keys.isin(new_keys)]

    combined = pd.concat([existing, new_df], ignore_index=True)

    # Rewrite sheet
    if sheet_name in book.sheetnames:
        del book[sheet_name]
    ws = book.create_sheet(sheet_name)

    for r in dataframe_to_rows(combined, index=False, header=True):
        ws.append(r)

    book.save(EXCEL_FILE)

# -------------------- Sidebar: Uploads --------------------
st.sidebar.header("📤 Upload New Weekly Data (.csv)")
uploaded_offense   = st.sidebar.file_uploader("Offense CSV",   type="csv", key="up_off")
uploaded_defense   = st.sidebar.file_uploader("Defense CSV",   type="csv", key="up_def")
uploaded_strategy  = st.sidebar.file_uploader("Strategy CSV",  type="csv", key="up_str")
uploaded_personnel = st.sidebar.file_uploader("Personnel CSV", type="csv", key="up_per")
uploaded_injuries  = st.sidebar.file_uploader("Injuries CSV (optional)", type="csv", key="up_inj")

if uploaded_offense:
    df_off = pd.read_csv(uploaded_offense)
    _append_to_excel(df_off, "Offense", key_cols=["Week"], deduplicate=True)
    st.sidebar.success("✅ Offense uploaded")

if uploaded_defense:
    df_def = pd.read_csv(uploaded_defense)
    _append_to_excel(df_def, "Defense", key_cols=["Week"], deduplicate=True)
    st.sidebar.success("✅ Defense uploaded")

if uploaded_strategy:
    df_str = pd.read_csv(uploaded_strategy)
    # Strategy often has opponent per week; dedup by Week if present
    keys = ["Week"] if "Week" in df_str.columns else None
    _append_to_excel(df_str, "Strategy", key_cols=keys, deduplicate=True)
    st.sidebar.success("✅ Strategy uploaded")

if uploaded_personnel:
    df_per = pd.read_csv(uploaded_personnel)
    _append_to_excel(df_per, "Personnel", key_cols=["Week"], deduplicate=True)
    st.sidebar.success("✅ Personnel uploaded")

if uploaded_injuries:
    df_inj = pd.read_csv(uploaded_injuries)
    # Injuries: dedup by Week + Player
    keys = [k for k in ["Week","Player"] if k in df_inj.columns]
    _append_to_excel(df_inj, "Injuries", key_cols=keys or None, deduplicate=True)
    st.sidebar.success("✅ Injuries uploaded")

# Download workbook
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------- Sidebar: Fetchers --------------------
with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pulls **2025** team-level stats for CHI (best effort).")
    fetch_week = st.number_input("Week to fetch (2025)", 1, 25, 1, 1, key="fetch_wk")
    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl

            # Try to update local cache—ignore if not available
            for fn in (getattr(nfl, "update", None) or {}).__dict__.values() if hasattr(nfl, "update") else []:
                try:
                    fn([2025])
                except Exception:
                    pass

            weekly = nfl.import_weekly_data([2025])
            team_week = weekly[(weekly["team"]=="CHI") & (weekly["week"]==int(fetch_week))].copy()

            if team_week.empty:
                st.warning("No team row for CHI in that week yet.")
            else:
                team_week["Week"] = int(fetch_week)

                # --- OFFENSE (simple best-effort mappings) ---
                pass_yards = team_week["passing_yards"].iloc[0] if "passing_yards" in team_week.columns else None
                pass_att = None
                for cand in ["attempts","passing_attempts","pass_attempts"]:
                    if cand in team_week.columns:
                        pass_att = team_week[cand].iloc[0]
                        break
                ypa = None
                try:
                    ypa = round(float(pass_yards)/float(pass_att),2) if pass_yards not in (None,pd.NA) and pass_att not in (None,0,pd.NA) else None
                except Exception:
                    ypa = None

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
                try:
                    if completions is not None and pass_att not in (None,0):
                        cmp_pct = round(float(completions)/float(pass_att)*100,1)
                except Exception:
                    cmp_pct = None

                off_row = pd.DataFrame([{
                    "Week": int(fetch_week),
                    "YPA": ypa,
                    "YDS": yards_total,
                    "CMP%": cmp_pct
                }])

                # --- DEFENSE (basic) ---
                sacks_val = None
                for cand in ["sacks","defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].iloc[0]
                        break

                def_row = pd.DataFrame([{
                    "Week": int(fetch_week),
                    "SACK": sacks_val,
                    "RZ% Allowed": None  # filled by PBP fetch below
                }])

                _append_to_excel(off_row, "Offense", key_cols=["Week"], deduplicate=True)
                _append_to_excel(def_row, "Defense", key_cols=["Week"], deduplicate=True)
                _append_to_excel(team_week.rename(columns=str), "Raw_Weekly", key_cols=None, deduplicate=False)

                st.success(f"✅ Added CHI week {int(fetch_week)} to Offense/Defense (available fields).")
                st.caption("Tip: use 'Fetch PBP Metrics' to fill RZ% Allowed & Pressures.")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📡 Fetch Defensive Metrics (Play-by-Play)")
pbp_week = st.sidebar.number_input("Week (2025)", 1, 25, 1, 1, key="pbp_wk")
if st.sidebar.button("Fetch PBP Metrics"):
    try:
        import nfl_data_py as nfl
        pbp = nfl.import_pbp_data([2025], downcast=False)
        pbp_w = pbp[(pbp["week"]==int(pbp_week)) & (pbp["defteam"]=="CHI")].copy()
        if pbp_w.empty:
            st.warning("No PBP rows for CHI defense that week yet.")
        else:
            # Red Zone Allowed (% of drives that reach <=20)
            dmins = (pbp_w.groupby(["game_id","drive"], as_index=False)["yardline_100"]
                         .min().rename(columns={"yardline_100":"min_yardline_100"}))
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives>0 else 0
            rz_allowed = (rz_drives/total_drives*100) if total_drives>0 else 0.0

            # Success Rate (offense success vs CHI)
            def _success(row):
                if pd.isna(row.get("down")) or pd.isna(row.get("ydstogo")) or pd.isna(row.get("yards_gained")):
                    return False
                d = int(row["down"]); togo = float(row["ydstogo"]); gain = float(row["yards_gained"])
                if d==1: return gain >= 0.4*togo
                if d==2: return gain >= 0.6*togo
                return gain >= togo

            plays_mask = (~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))
            pbp_real = pbp_w[plays_mask].copy()
            success_rate = pbp_real.apply(_success, axis=1).mean()*100 if len(pbp_real) else 0.0

            # Pressures ≈ sacks + qb hits (best effort)
            qb_hits = pbp_w["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in pbp_w.columns else 0
            sacks = pbp_w["sack"].fillna(0).astype(int).sum() if "sack" in pbp_w.columns else 0
            pressures = int(qb_hits + sacks)

            adv = pd.DataFrame([{
                "Week": int(pbp_week),
                "RZ% Allowed": round(rz_allowed,1),
                "Success Rate% (Offense)": round(success_rate,1),
                "Pressures": pressures
            }])
            _append_to_excel(adv, "Advanced_Defense", key_cols=["Week"], deduplicate=True)
            st.success(f"✅ PBP saved — RZ% Allowed {rz_allowed:.1f} | SR {success_rate:.1f} | Pressures {pressures}")
    except Exception as e:
        st.error(f"❌ PBP fetch failed: {e}")

# -------------------- Center: Quick Previews (what you just uploaded/fetched) --------------------
df_off_prev = _read_sheet("Offense")
df_def_prev = _read_sheet("Defense")
df_str_prev = _read_sheet("Strategy")
df_per_prev = _read_sheet("Personnel")

if not df_off_prev.empty:
    st.subheader("📊 Offensive Analytics (preview)")
    st.dataframe(df_off_prev.sort_values("Week"), use_container_width=True)
if not df_def_prev.empty:
    st.subheader("🛡️ Defensive Analytics (preview)")
    st.dataframe(df_def_prev.sort_values("Week"), use_container_width=True)
if not df_per_prev.empty:
    st.subheader("👥 Personnel Usage (preview)")
    st.dataframe(df_per_prev.sort_values("Week"), use_container_width=True)
if not df_str_prev.empty:
    st.subheader("📘 Weekly Strategy (preview)")
    st.dataframe(df_str_prev.sort_values("Week"), use_container_width=True)

# -------------------- Media Summary --------------------
st.markdown("### 📰 Weekly Beat Writer / ESPN Summary")
with st.form("media_form"):
    media_week = st.number_input("Week", 1, 25, 1, 1, key="media_week_input")
    media_opponent = st.text_input("Opponent")
    media_summary = st.text_area("Beat Writer & ESPN Summary (Game Recap, Analysis, Strategy, etc.)")
    submit_media = st.form_submit_button("Save Summary")

if submit_media:
    media_df = pd.DataFrame([{
        "Week": int(media_week),
        "Opponent": media_opponent,
        "Summary": media_summary
    }])
    _append_to_excel(media_df, "Media_Summaries", key_cols=["Week","Opponent"], deduplicate=True)
    st.success(f"✅ Summary for Week {media_week} saved.")

if os.path.exists(EXCEL_FILE):
    try:
        df_media = _read_sheet("Media_Summaries")
        if not df_media.empty:
            st.subheader("📰 Saved Media Summaries")
            st.dataframe(df_media.sort_values("Week"), use_container_width=True)
    except Exception:
        st.info("No media summaries stored yet.")

# -------------------- DVOA-like Proxy (Opponent-Adjusted EPA/SR) --------------------
st.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.number_input("Week to Compute (2025)", 1, 25, 1, 1, key="proxy_week")

def _success_flag(down, ydstogo, yards_gained):
    try:
        if pd.isna(down) or pd.isna(ydstogo) or pd.isna(yards_gained):
            return False
        d = int(down); togo = float(ydstogo); gain = float(yards_gained)
        if d == 1: return gain >= 0.4 * togo
        if d == 2: return gain >= 0.6 * togo
        return gain >= togo
    except Exception:
        return False

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
            st.warning("No CHI plays found for that week yet.")
        else:
            opps = set()
            if not bears_off.empty: opps.update(bears_off["defteam"].unique().tolist())
            if not bears_def.empty: opps.update(bears_def["posteam"].unique().tolist())
            opponent = list(opps)[0] if opps else "UNK"

            prior = plays[plays["week"] < wk].copy()

            opp_def = prior[prior["defteam"]==opponent].copy()
            opp_def_epa = opp_def["epa"].mean() if len(opp_def) else None
            opp_def_sr  = opp_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_def) else None

            opp_off = prior[prior["posteam"]==opponent].copy()
            opp_off_epa = opp_off["epa"].mean() if len(opp_off) else None
            opp_off_sr  = opp_off.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_off) else None

            if len(bears_off):
                chi_off_epa = bears_off["epa"].mean()
                chi_off_sr  = bears_off.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_off_epa = None; chi_off_sr = None

            if len(bears_def):
                chi_def_epa_allowed = bears_def["epa"].mean()
                chi_def_sr_allowed  = bears_def.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                chi_def_epa_allowed = None; chi_def_sr_allowed = None

            def diff(a,b):
                if a is None or pd.isna(a) or b is None or pd.isna(b): return None
                return float(a) - float(b)

            out = pd.DataFrame([{
                "Week": wk,
                "Opponent": opponent,
                "Off Adj EPA/play": round(diff(chi_off_epa, opp_def_epa), 3) if diff(chi_off_epa, opp_def_epa) is not None else None,
                "Off Adj SR%": round(diff(chi_off_sr,  opp_def_sr)*100, 1) if diff(chi_off_sr, opp_def_sr) is not None else None,
                "Def Adj EPA/play": round(diff(chi_def_epa_allowed, opp_off_epa), 3) if diff(chi_def_epa_allowed, opp_off_epa) is not None else None,
                "Def Adj SR%": round(diff(chi_def_sr_allowed,  opp_off_sr)*100, 1) if diff(chi_def_sr_allowed, opp_off_sr) is not None else None,
                "Off EPA/play": round(chi_off_epa, 3) if chi_off_epa is not None else None,
                "Def EPA allowed/play": round(chi_def_epa_allowed, 3) if chi_def_epa_allowed is not None else None
            }])

            _append_to_excel(out, "DVOA_Proxy", key_cols=["Week"], deduplicate=True)
            st.success(
                f"✅ DVOA-like proxy saved for Week {wk} vs {opponent} — "
                f"Off Adj EPA/play={out.iloc[0]['Off Adj EPA/play']}, Off Adj SR%={out.iloc[0]['Off Adj SR%']}, "
                f"Def Adj EPA/play={out.iloc[0]['Def Adj EPA/play']}, Def Adj SR%={out.iloc[0]['Def Adj SR%']}"
            )
    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# Optional preview
df_dvoa_prev = _read_sheet("DVOA_Proxy")
if not df_dvoa_prev.empty:
    st.subheader("📊 DVOA-like Proxy (preview)")
    st.dataframe(df_dvoa_prev.sort_values("Week"), use_container_width=True)

# -------------------- Weekly Game Prediction --------------------
st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", 1, 25, 1, 1, key="pred_week")

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy = _read_sheet("Strategy")
        df_offense  = _read_sheet("Offense")
        df_defense  = _read_sheet("Defense")
        try:
            df_advdef = _read_sheet("Advanced_Defense")
        except Exception:
            df_advdef = pd.DataFrame()
        try:
            df_proxy = _read_sheet("DVOA_Proxy")
        except Exception:
            df_proxy = pd.DataFrame()

        row_s = df_strategy[df_strategy["Week"]==week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"]==week_to_predict]   if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"]==week_to_predict]   if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"]==week_to_predict]     if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"]==week_to_predict]       if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            rz_allowed = None
            pressures  = None
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
                reason_bits.append(f"Off+{off_adj_epa:+.2f} EPA/play vs opp D")
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

            pred_row = pd.DataFrame([{
                "Week": int(week_to_predict),
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else "",
                "Notes": reason_text
            }])
            _append_to_excel(pred_row, "Predictions", key_cols=["Week"], deduplicate=True)
        else:
            st.info("Please upload/fetch Strategy, Offense, and Defense for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# Show saved predictions
df_preds_prev = _read_sheet("Predictions")
if not df_preds_prev.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(df_preds_prev.sort_values("Week"), use_container_width=True)

# -------------------- Injuries: Quick Entry --------------------
st.markdown("### 🩹 Injuries — Quick Entry")
with st.form("injury_quick_entry"):
    q_week = st.number_input("Week", 1, 25, 1, 1, key="inj_q_week")
    q_player = st.text_input("Player")
    q_pos = st.text_input("Position")
    q_status = st.selectbox("Status", ["questionable","doubtful","out","IR","PUP","active","healthy","cleared"])
    q_part = st.text_input("Body Part / Injury")
    q_wed = st.selectbox("Practice Wed", ["DNP","LP","FP","N/A"])
    q_thu = st.selectbox("Practice Thu", ["DNP","LP","FP","N/A"])
    q_fri = st.selectbox("Practice Fri", ["DNP","LP","FP","N/A"])
    q_gstat = st.selectbox("Game Status", ["TBD","Active","Inactive","Out"])
    q_notes = st.text_area("Notes (optional)")
    q_submit = st.form_submit_button("Add/Update Injury")

if q_submit:
    inj_row = pd.DataFrame([{
        "Week": int(q_week),
        "Player": q_player,
        "Position": q_pos,
        "Status": q_status,
        "BodyPart": q_part,
        "Practice_Wed": q_wed,
        "Practice_Thu": q_thu,
        "Practice_Fri": q_fri,
        "GameStatus": q_gstat,
        "Notes": q_notes,
        "DateUpdated": pd.Timestamp.now(tz="America/Chicago")
    }])
    _append_to_excel(inj_row, "Injuries", key_cols=["Week","Player"], deduplicate=True)
    st.success(f"✅ Saved injury for {q_player} (Week {q_week}).")

# -------------------- Injuries — This Week (active only) --------------------
st.markdown("### 🩹 Injuries — This Week (Active Only)")
try:
    default_inj_week = int(week_to_predict)
except Exception:
    default_inj_week = 1

inj_view_week = st.number_input(
    "Week to view injuries",
    min_value=1, max_value=25, value=default_inj_week, step=1, key="inj_view_week"
)

df_inj_all = _read_sheet("Injuries")
if df_inj_all.empty:
    st.info("No injuries recorded yet.")
else:
    status_hide = {"healthy", "cleared", "active"}  # not shown
    # coerce Week to int safely
    df_inj_all = df_inj_all.copy()
    if "Week" in df_inj_all.columns:
        df_inj_all["Week"] = pd.to_numeric(df_inj_all["Week"], errors="coerce").astype("Int64")
    mask_week = (df_inj_all["Week"] == int(inj_view_week)) if "Week" in df_inj_all.columns else True
    mask_active = ~df_inj_all.get("Status", pd.Series([], dtype=str)).astype(str).str.strip().str.lower().isin(status_hide)
    df_inj_view = df_inj_all[mask_week & mask_active].copy()

    if df_inj_view.empty:
        st.info(f"No *active* injuries recorded for Week {inj_view_week}.")
    else:
        preferred_cols = [
            "Week","Player","Position","Status","BodyPart",
            "Practice_Wed","Practice_Thu","Practice_Fri",
            "GameStatus","Notes","DateUpdated"
        ]
        cols = [c for c in preferred_cols if c in df_inj_view.columns] + [c for c in df_inj_view.columns if c not in preferred_cols]
        sort_cols = [c for c in ["Position","Player","Status","DateUpdated"] if c in df_inj_view.columns]
        if sort_cols:
            df_inj_view = df_inj_view.sort_values(by=sort_cols)
        st.dataframe(df_inj_view[cols], use_container_width=True)

# -------------------- (Optional) Sanity Checker --------------------
with st.expander("🧪 Excel Sanity Checker (optional)"):
    st.caption("Shows file path, size, and the sheets currently in your workbook.")
    cwd = os.getcwd()
    file_path = os.path.abspath(EXCEL_FILE)
    st.write("**Current Working Directory:**", cwd)
    st.write("**Excel File Being Used:**", file_path)
    exists = os.path.exists(EXCEL_FILE)
    st.write("**Exists:**", exists)
    if exists:
        st.write("**Size (bytes):**", os.path.getsize(EXCEL_FILE))
        try:
            import openpyxl
            book = openpyxl.load_workbook(EXCEL_FILE)
            st.write("**Sheets:**", list(book.sheetnames))
        except Exception as e:
            st.warning(f"Workbook open error: {e}")
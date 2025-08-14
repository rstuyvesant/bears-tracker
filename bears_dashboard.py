import streamlit as st
import pandas as pd
import os

# Optional: heavy imports only when needed in their sections to avoid cold-start time
# from fpdf import FPDF  # Uncomment later if/when you re-enable PDF reports

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown("Track weekly stats, strategy, personnel usage, and league comparisons.")

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# ----------------------------- Utilities -----------------------------

def _read_excel_sheet(sheet_name: str) -> pd.DataFrame:
    """Safe read of a sheet; returns empty DataFrame if missing."""
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def append_to_excel(new_data: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate: bool = False, dedup_keys=None):
    """
    Replace-or-merge write:
      - If workbook/sheet exists, read sheet, merge with new_data.
      - If deduplicate=True, drop duplicates on dedup_keys (default = ['Week']).
      - Rewrite the sheet in-place.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    if new_data is None or new_data.empty:
        return

    try:
        # Normalize new_data columns to strings to avoid mixed types
        new_data = new_data.copy()
        new_data.columns = [str(c) for c in new_data.columns]

        # Read existing
        existing = _read_excel_sheet(sheet_name)
        if not existing.empty:
            existing.columns = [str(c) for c in existing.columns]
            combined = pd.concat([existing, new_data], ignore_index=True)
        else:
            combined = new_data

        # Dedup if requested
        if deduplicate:
            if dedup_keys is None:
                dedup_keys = ["Week"]
            # Ensure keys exist
            for k in dedup_keys:
                if k not in combined.columns:
                    combined[k] = None
            # Normalize Week to numeric for stable dedup
            if "Week" in combined.columns:
                combined["Week"] = pd.to_numeric(combined["Week"], errors="coerce").astype("Int64")
            combined = combined.drop_duplicates(subset=dedup_keys, keep="last")

        # Open or create workbook
        if os.path.exists(file_name):
            book = openpyxl.load_workbook(file_name)
        else:
            book = openpyxl.Workbook()
            # remove default sheet
            if book.active and book.active.title == "Sheet":
                book.remove(book.active)

        # Remove old sheet if present
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)

        # Write combined
        for r in dataframe_to_rows(combined, index=False, header=True):
            sheet.append(r)

        book.save(file_name)
    except Exception as e:
        st.error(f"Excel append error: {e}")

def _safe_float(x, default=None):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return default
        return float(x)
    except Exception:
        return default

# ----------------------------- Sidebar: Uploads -----------------------------

st.sidebar.header("📤 Upload New Weekly Data")
uploaded_offense   = st.sidebar.file_uploader("Upload Offensive Analytics (.csv)", type="csv")
uploaded_defense   = st.sidebar.file_uploader("Upload Defensive Analytics (.csv)", type="csv")
uploaded_strategy  = st.sidebar.file_uploader("Upload Weekly Strategy (.csv)", type="csv")
uploaded_personnel = st.sidebar.file_uploader("Upload Personnel Usage (.csv)", type="csv")

if uploaded_offense:
    df_offense = pd.read_csv(uploaded_offense)
    append_to_excel(df_offense, "Offense", deduplicate=True, dedup_keys=["Week"])
    st.sidebar.success("✅ Offensive data uploaded.")

if uploaded_defense:
    df_defense = pd.read_csv(uploaded_defense)
    append_to_excel(df_defense, "Defense", deduplicate=True, dedup_keys=["Week"])
    st.sidebar.success("✅ Defensive data uploaded.")

if uploaded_strategy:
    df_strategy = pd.read_csv(uploaded_strategy)
    append_to_excel(df_strategy, "Strategy", deduplicate=True, dedup_keys=["Week"])
    st.sidebar.success("✅ Strategy data uploaded.")

if uploaded_personnel:
    df_personnel = pd.read_csv(uploaded_personnel)
    append_to_excel(df_personnel, "Personnel", deduplicate=True, dedup_keys=["Week"])
    st.sidebar.success("✅ Personnel data uploaded.")

# ----------------------------- Sidebar: Auto Fetch (nfl_data_py) -----------------------------

with st.sidebar.expander("⚡ Fetch Weekly Data (nfl_data_py)"):
    st.caption("Pulls 2025 weekly team stats for CHI and saves to Excel.")
    fetch_week = st.number_input(
        "Week to fetch (2025 season)", min_value=1, max_value=25, value=1, step=1, key="fetch_week_2025"
    )

    if st.button("Fetch CHI Week via nfl_data_py"):
        try:
            import nfl_data_py as nfl

            # Optional cache refresh (best-effort)
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
                st.warning("No weekly team row found for CHI in that week. Try again later or verify the week.")
            else:
                team_week = team_week.rename(columns=str).copy()
                team_week["Week"] = wk

                # Offense fields (best-effort)
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

                # Defense (team-level; some advanced items need PBP)
                sacks_val = None
                for cand in ["sacks", "defense_sacks"]:
                    if cand in team_week.columns:
                        sacks_val = team_week[cand].values[0]
                        break

                def_row = pd.DataFrame([{
                    "Week": wk,
                    "SACK": sacks_val,
                    "RZ% Allowed": None  # requires play-by-play aggregation
                }])

                append_to_excel(off_row, "Offense", deduplicate=True, dedup_keys=["Week"])
                append_to_excel(def_row, "Defense", deduplicate=True, dedup_keys=["Week"])
                append_to_excel(team_week, "Raw_Weekly", deduplicate=False)

                st.success(f"✅ Added CHI week {wk} to Offense/Defense (available fields).")
                st.caption("Note: Red Zone % Allowed and pressures need play-by-play; use the panel below.")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

# --------------------- Sidebar: Fetch PBP-derived Defensive Metrics ---------------------

st.sidebar.markdown("### 📡 Fetch Defensive Metrics from Play-by-Play")
pbp_week = st.sidebar.number_input("Week to Fetch (2025 Season)", min_value=1, max_value=25, value=1, step=1, key="pbp_week_2025")

if st.sidebar.button("Fetch Play-by-Play Metrics"):
    try:
        import nfl_data_py as nfl

        pbp = nfl.import_pbp_data([2025], downcast=False)
        pbp_w = pbp[(pbp["week"] == int(pbp_week)) & (pbp["defteam"] == "CHI")].copy()

        if pbp_w.empty:
            st.warning("No PBP rows for CHI defense in that week yet. Try again later.")
        else:
            # Red Zone % Allowed: drives where min_yardline_100 <= 20
            dmins = (
                pbp_w.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                .min()
                .rename(columns={"yardline_100": "min_yardline_100"})
            )
            total_drives = len(dmins)
            rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
            rz_allowed = (rz_drives / total_drives * 100) if total_drives > 0 else 0.0

            # Success Rate (offense success against CHI defense)
            def play_success(row):
                if pd.isna(row.get("down")) or pd.isna(row.get("ydstogo")) or pd.isna(row.get("yards_gained")):
                    return False
                d = int(row["down"])
                togo = float(row["ydstogo"])
                gain = float(row["yards_gained"])
                if d == 1:
                    return gain >= 0.4 * togo
                elif d == 2:
                    return gain >= 0.6 * togo
                else:
                    return gain >= togo

            plays_mask = (~pbp_w["play_type"].isin(["no_play"])) & (~pbp_w["penalty"].fillna(False))
            pbp_real = pbp_w[plays_mask].copy()
            success_rate = pbp_real.apply(play_success, axis=1).mean() * 100 if len(pbp_real) else 0.0

            # Pressures approximation: qb_hit + sack
            qb_hits = pbp_w["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in pbp_w.columns else 0
            sacks = pbp_w["sack"].fillna(0).astype(int).sum() if "sack" in pbp_w.columns else 0
            pressures = int(qb_hits + sacks)

            metrics_df = pd.DataFrame([{
                "Week": int(pbp_week),
                "RZ% Allowed": round(rz_allowed, 1),
                "Success Rate% (Offense)": round(success_rate, 1),
                "Pressures": pressures
            }])
            append_to_excel(metrics_df, "Advanced_Defense", deduplicate=True, dedup_keys=["Week"])

            st.success(
                f"✅ Week {int(pbp_week)} PBP metrics saved — RZ% Allowed: {rz_allowed:.1f} | "
                f"Success Rate% (Off): {success_rate:.1f} | Pressures: {pressures}"
            )
            st.caption("Note: Hurries aren’t separately flagged in standard PBP; pressures = sacks + QB hits.")
    except Exception as e:
        st.error(f"❌ Failed to fetch metrics: {e}")

# --------------------- Sidebar: DVOA-like Proxy (Opponent-Adjusted) ---------------------

st.sidebar.markdown("### 📈 Compute DVOA-like Proxy (Opponent-Adjusted)")
proxy_week = st.sidebar.number_input("Week to Compute (2025 Season)", min_value=1, max_value=25, value=1, step=1, key="proxy_week_2025")

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
            if len(opp_def_plays):
                opp_def_success = opp_def_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                opp_def_success = None

            opp_off_plays = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off_plays["epa"].mean() if len(opp_off_plays) else None
            if len(opp_off_plays):
                opp_off_success = opp_off_plays.apply(lambda r: _success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean()
            else:
                opp_off_success = None

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

            append_to_excel(out, "DVOA_Proxy", deduplicate=True, dedup_keys=["Week"])
            st.success(
                f"✅ DVOA-like proxy saved for Week {wk} vs {opponent} "
                f"(Off Adj EPA/play={out.iloc[0]['Off Adj EPA/play']}, Off Adj SR%={out.iloc[0]['Off Adj SR%']}, "
                f"Def Adj EPA/play={out.iloc[0]['Def Adj EPA/play']}, Def Adj SR%={out.iloc[0]['Def Adj SR%']})"
            )

    except Exception as e:
        st.error(f"❌ Failed to compute proxy: {e}")

# ----------------------------- Main: Preview Uploaded Data -----------------------------

# Offense
df_offense_view = _read_excel_sheet("Offense")
if not df_offense_view.empty:
    st.subheader("📊 Offensive Analytics")
    st.dataframe(df_offense_view)

# Defense
df_defense_view = _read_excel_sheet("Defense")
if not df_defense_view.empty:
    st.subheader("🛡️ Defensive Analytics")
    st.dataframe(df_defense_view)

# Strategy
df_strategy_view = _read_excel_sheet("Strategy")
if not df_strategy_view.empty:
    st.subheader("📘 Weekly Strategy")
    st.dataframe(df_strategy_view)

# Personnel
df_personnel_view = _read_excel_sheet("Personnel")
if not df_personnel_view.empty:
    st.subheader("👥 Personnel Usage")
    st.dataframe(df_personnel_view)

# ----------------------------- Injuries Block (upload + add/edit) -----------------------------

st.markdown("### 🩺 Injuries")

uploaded_injuries = st.sidebar.file_uploader("Upload Injuries (.csv)", type="csv", key="injuries_csv_uploader")

def save_injuries_rows(new_rows: pd.DataFrame):
    """Merge new_rows into 'Injuries', dedupe by ['Week','Player']."""
    try:
        existing = _read_excel_sheet("Injuries")
        needed_cols = ["Week","Opponent","Player","Position","Status","Practice_Fri","Notes"]

        # normalize new rows
        new_rows = new_rows.copy()
        for c in needed_cols:
            if c not in new_rows.columns:
                new_rows[c] = None
        new_rows = new_rows[needed_cols]

        if not existing.empty:
            for c in needed_cols:
                if c not in existing.columns:
                    existing[c] = None
            existing = existing[needed_cols]
            combined = pd.concat([existing, new_rows], ignore_index=True)
        else:
            combined = new_rows

        combined["Week"] = pd.to_numeric(combined["Week"], errors="coerce").astype("Int64")
        combined = combined.dropna(subset=["Week", "Player"])
        combined = combined.drop_duplicates(subset=["Week","Player"], keep="last")

        append_to_excel(combined, "Injuries", deduplicate=False)
        return True, len(new_rows), len(combined)
    except Exception as e:
        st.error(f"❌ Failed to save injuries: {e}")
        return False, 0, 0

if uploaded_injuries is not None:
    try:
        inj_csv = pd.read_csv(uploaded_injuries)
        ok, added, total = save_injuries_rows(inj_csv)
        if ok:
            st.sidebar.success(f"✅ Injuries uploaded. Added/updated {added} rows. Total now {total}.")
    except Exception as e:
        st.sidebar.error(f"❌ Could not read injuries CSV: {e}")

with st.form("injury_add_form"):
    c1, c2, c3 = st.columns([1,1,1])
    week_inj = c1.number_input("Week", min_value=1, max_value=25, step=1, value=1, key="inj_week")
    opp_inj  = c2.text_input("Opponent (optional)", value="", key="inj_opp")
    pos_inj  = c3.text_input("Position (QB, WR, RB, CB...)", value="", key="inj_pos")

    player_inj = st.text_input("Player", value="", key="inj_player")
    status_inj = st.selectbox("Status", ["Questionable","Doubtful","Out","IR","Active"], index=0, key="inj_status")
    pract_inj  = st.selectbox("Practice Friday", ["Full","Limited","DNP","N/A"], index=3, key="inj_practice")
    notes_inj  = st.text_area("Notes", value="", key="inj_notes")
    submit_inj = st.form_submit_button("➕ Add / Update Injury")

if submit_inj:
    if not player_inj.strip():
        st.warning("Please enter a player name.")
    else:
        new_row = pd.DataFrame([{
            "Week": week_inj,
            "Opponent": opp_inj.strip() or None,
            "Player": player_inj.strip(),
            "Position": pos_inj.strip() or None,
            "Status": status_inj,
            "Practice_Fri": pract_inj,
            "Notes": notes_inj.strip() or None
        }])
        ok, _, total = save_injuries_rows(new_row)
        if ok:
            st.success(f"✅ Injury saved for Week {week_inj} — {player_inj} ({status_inj}). Total rows: {total}")

injuries_df = _read_excel_sheet("Injuries")
if not injuries_df.empty:
    sel_week_inj = st.number_input("Week to view (injuries)", min_value=1, max_value=25, step=1, value=1, key="inj_view_week")
    view_inj = injuries_df[injuries_df["Week"] == sel_week_inj] if "Week" in injuries_df.columns else injuries_df
    st.dataframe(view_inj if not view_inj.empty else injuries_df)

    st.download_button(
        "⬇️ Download Injuries (CSV)",
        data=injuries_df.to_csv(index=False).encode("utf-8"),
        file_name="injuries_export.csv",
        mime="text/csv"
    )
else:
    st.info("No injuries recorded yet. Upload a CSV or add one via the form above.")

# ----------------------------- Media Summaries -----------------------------

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

df_media_view = _read_excel_sheet("Media_Summaries")
if not df_media_view.empty:
    st.subheader("📰 Saved Media Summaries")
    st.dataframe(df_media_view)

# ----------------------------- Weekly Game Prediction -----------------------------

st.markdown("### 🔮 Weekly Game Prediction")
week_to_predict = st.number_input("Select Week to Predict", min_value=1, max_value=25, step=1, key="predict_week_input")

if os.path.exists(EXCEL_FILE):
    try:
        df_strategy = _read_excel_sheet("Strategy")
        df_offense  = _read_excel_sheet("Offense")
        df_defense  = _read_excel_sheet("Defense")
        df_advdef   = _read_excel_sheet("Advanced_Defense")
        df_proxy    = _read_excel_sheet("DVOA_Proxy")

        row_s = df_strategy[df_strategy["Week"] == week_to_predict] if not df_strategy.empty else pd.DataFrame()
        row_o = df_offense[df_offense["Week"] == week_to_predict] if not df_offense.empty else pd.DataFrame()
        row_d = df_defense[df_defense["Week"] == week_to_predict] if not df_defense.empty else pd.DataFrame()
        row_a = df_advdef[df_advdef["Week"] == week_to_predict] if not df_advdef.empty else pd.DataFrame()
        row_p = df_proxy[df_proxy["Week"] == week_to_predict] if not df_proxy.empty else pd.DataFrame()

        if not row_s.empty and not row_o.empty and not row_d.empty:
            # Strategy text
            strategy_text = row_s.iloc[0].astype(str).str.cat(sep=" ").lower()

            # Offense basics
            ypa = _safe_float(row_o.iloc[0].get("YPA"), default=None)

            # Prefer Advanced_Defense RZ% & Pressures when available
            rz_allowed = None
            pressures = None
            if not row_a.empty:
                rz_allowed = _safe_float(row_a.iloc[0].get("RZ% Allowed"), default=None)
                pressures  = _safe_float(row_a.iloc[0].get("Pressures"), default=None)
            if rz_allowed is None:
                rz_allowed = _safe_float(row_d.iloc[0].get("RZ% Allowed"), default=None)

            # DVOA-like proxy (opponent-adjusted EPA/SR)
            off_adj_epa = off_adj_sr = def_adj_epa = def_adj_sr = None
            if not row_p.empty:
                off_adj_epa = _safe_float(row_p.iloc[0].get("Off Adj EPA/play"), default=None)
                off_adj_sr  = _safe_float(row_p.iloc[0].get("Off Adj SR%"), default=None)
                def_adj_epa = _safe_float(row_p.iloc[0].get("Def Adj EPA/play"), default=None)
                def_adj_sr  = _safe_float(row_p.iloc[0].get("Def Adj SR%"), default=None)

            # Rule set
            reason_bits = []

            if (off_adj_epa is not None and off_adj_epa >= 0.15) and (def_adj_epa is not None and def_adj_epa <= -0.05):
                prediction = "Win - efficiency edge on both sides"
                reason_bits.append(f"Off{off_adj_epa:+.2f} EPA/play vs opp D")
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
            prediction_entry = pd.DataFrame([{
                "Week": week_to_predict,
                "Prediction": prediction.split("-")[0].strip(),
                "Reason": prediction.split("-")[1].strip() if "-" in prediction else "",
                "Notes": reason_text
            }])
            append_to_excel(prediction_entry, "Predictions", deduplicate=True, dedup_keys=["Week"])

        else:
            st.info("Please upload or fetch Strategy, Offense, and Defense data for this week first.")
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# ----------------------------- DVOA Proxy & Predictions Preview -----------------------------

df_dvoa_view = _read_excel_sheet("DVOA_Proxy")
if not df_dvoa_view.empty:
    st.subheader("📊 DVOA-like Proxy Metrics")
    st.dataframe(df_dvoa_view)

df_preds_view = _read_excel_sheet("Predictions")
if not df_preds_view.empty:
    st.subheader("📈 Saved Game Predictions")
    st.dataframe(df_preds_view)

# ----------------------------- Download Workbook -----------------------------

if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="⬇️ Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.warning(f"Prediction failed. Check uploaded/fetched data. Error: {e}")

# Show saved predictions
if os.path.exists(EXCEL_FILE):
    try:
        df_preds = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
        st.subheader("📈 Saved Game Predictions")
        st.dataframe(df_preds, use_container_width=True)
    except Exception:
        st.info("No predictions saved yet.")
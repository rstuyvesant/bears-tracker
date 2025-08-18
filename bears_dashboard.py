import os
from datetime import datetime
import streamlit as st
import pandas as pd

# External deps used at runtime in some buttons:
# - openpyxl for Excel I/O
# - nfl_data_py for free NFL data (weekly & PBP)
# - fpdf for a simple PDF

st.set_page_config(page_title="Chicago Bears 2025–26 Weekly Tracker", layout="wide")
st.title("🐻 Chicago Bears 2025–26 Weekly Tracker")
st.markdown(
    "Track weekly stats, strategy, personnel, injuries, snap counts, opponent previews, "
    "and compute an opponent-adjusted DVOA-like proxy from play-by-play."
)

EXCEL_FILE = "bears_weekly_analytics.xlsx"

# =========================
# Excel helpers
# =========================
def _ensure_openpyxl():
    import openpyxl  # noqa: F401


def append_to_excel(new_df: pd.DataFrame, sheet_name: str, file_name: str = EXCEL_FILE, deduplicate_on: list | None = None):
    """
    Append or create sheet. Optionally drop duplicates using columns in deduplicate_on.
    Uses openpyxl under the hood to preserve other sheets.
    """
    _ensure_openpyxl()
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.workbook import Workbook

    # Normalize column names for consistency
    new_df = new_df.copy()
    new_df.columns = [str(c).strip() for c in new_df.columns]

    if os.path.exists(file_name):
        book = load_workbook(file_name)
    else:
        book = Workbook()
        # remove default sheet
        if "Sheet" in book.sheetnames and len(book.sheetnames) == 1:
            book.remove(book.active)

    # Read existing into DataFrame (if any)
    if sheet_name in book.sheetnames:
        ws = book[sheet_name]
        rows = list(ws.values)
        if rows:
            header = [str(h).strip() if h is not None else "" for h in rows[0]]
            existing = pd.DataFrame(rows[1:], columns=header)
        else:
            existing = pd.DataFrame(columns=list(new_df.columns))
        # concat & dedup
        combined = pd.concat([existing, new_df], ignore_index=True)
        if deduplicate_on:
            combined = combined.drop_duplicates(subset=deduplicate_on, keep="last")
        else:
            combined = combined.drop_duplicates(keep="last")
        # clear and rewrite
        del book[sheet_name]
        ws = book.create_sheet(sheet_name)
        for r in dataframe_to_rows(combined, index=False, header=True):
            ws.append(r)
    else:
        ws = book.create_sheet(sheet_name)
        from openpyxl.utils.dataframe import dataframe_to_rows
        for r in dataframe_to_rows(new_df, index=False, header=True):
            ws.append(r)

    book.save(file_name)


def read_sheet(sheet_name: str) -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


# =========================
# Sidebar: housekeeping
# =========================
with st.sidebar:
    st.header("🗂️ Housekeeping")
    st.caption(f"Working file: **{EXCEL_FILE}**")

    # Create next-week template row (helps avoid manual CSV scaffolding)
    st.markdown("**Template Tools**")
    next_week = st.number_input("Next Week # (1–25)", min_value=1, max_value=25, value=1, step=1, key="next_week_template")
    if st.button("➕ Create Next Week Template Row"):
        base_cols = ["Week"]
        # Offense/Defense/Strategy/Personnel minimal stubs
        off_stub = pd.DataFrame([{"Week": int(next_week)}])
        def_stub = pd.DataFrame([{"Week": int(next_week)}])
        strat_stub = pd.DataFrame([{"Week": int(next_week), "Opponent": "", "Off_Strategy": "", "Def_Strategy": "", "Notes": ""}])
        per_stub = pd.DataFrame([{"Week": int(next_week)}])
        append_to_excel(off_stub, "Offense", deduplicate_on=["Week"])
        append_to_excel(def_stub, "Defense", deduplicate_on=["Week"])
        append_to_excel(strat_stub, "Strategy", deduplicate_on=["Week"])
        append_to_excel(per_stub, "Personnel", deduplicate_on=["Week"])
        st.success(f"Template rows created for Week {int(next_week)} in Offense/Defense/Strategy/Personnel.")

    # Optional: Excel sanity checker
    if st.toggle("Show Excel Sanity Checker", value=False):
        st.subheader("🧪 Excel Sanity Checker")
        st.write("Exists:", os.path.exists(EXCEL_FILE))
        if os.path.exists(EXCEL_FILE):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(EXCEL_FILE)
                st.write("Sheets:", wb.sheetnames)
                for s in wb.sheetnames[:6]:
                    df_prev = read_sheet(s)
                    st.markdown(f"**Preview – {s}**")
                    st.dataframe(df_prev.head(10))
            except Exception as e:
                st.error(f"Checker error: {e}")

# =========================
# Upload panels
# =========================
st.header("📤 Upload Weekly Data (CSV)")
col_u1, col_u2 = st.columns(2)
with col_u1:
    up_off = st.file_uploader("Offense (.csv)", type="csv", key="up_off")
    if up_off:
        df = pd.read_csv(up_off)
        append_to_excel(df, "Offense", deduplicate_on=["Week"])
        st.success("Offense uploaded.")

    up_def = st.file_uploader("Defense (.csv)", type="csv", key="up_def")
    if up_def:
        df = pd.read_csv(up_def)
        append_to_excel(df, "Defense", deduplicate_on=["Week"])
        st.success("Defense uploaded.")

    up_strat = st.file_uploader("Strategy (.csv)", type="csv", key="up_strat")
    if up_strat:
        df = pd.read_csv(up_strat)
        append_to_excel(df, "Strategy", deduplicate_on=["Week"])
        st.success("Strategy uploaded.")

    up_personnel = st.file_uploader("Personnel (.csv)", type="csv", key="up_personnel")
    if up_personnel:
        df = pd.read_csv(up_personnel)
        append_to_excel(df, "Personnel", deduplicate_on=["Week"])
        st.success("Personnel uploaded.")

with col_u2:
    up_inj = st.file_uploader("Injuries (.csv)", type="csv", key="up_inj")
    if up_inj:
        df = pd.read_csv(up_inj)
        # dedup by Week + Player so updates overwrite
        if "Week" in df.columns and "Player" in df.columns:
            append_to_excel(df, "Injuries", deduplicate_on=["Week", "Player"])
        else:
            append_to_excel(df, "Injuries")
        st.success("Injuries uploaded.")

    up_snap = st.file_uploader("Snap Counts (.csv)", type="csv", key="up_snap")
    if up_snap:
        df = pd.read_csv(up_snap)
        # If you include Player column, dedup with it; else only Week
        if "Week" in df.columns and "Player" in df.columns:
            append_to_excel(df, "SnapCounts", deduplicate_on=["Week", "Player"])
        else:
            append_to_excel(df, "SnapCounts", deduplicate_on=["Week"])
        st.success("Snap Counts uploaded.")

    up_preview = st.file_uploader("Opponent Preview (.csv)", type="csv", key="up_prev")
    if up_preview:
        df = pd.read_csv(up_preview)
        # Typical columns: Week, Opponent, Strengths, Weaknesses, Key_Players, Notes
        if "Week" in df.columns:
            append_to_excel(df, "Opponent_Preview", deduplicate_on=["Week"])
        else:
            append_to_excel(df, "Opponent_Preview")
        st.success("Opponent Preview uploaded.")

# Quick entry for injuries (optional)
st.subheader("➕ Injury Quick Entry")
with st.form("inj_quick_form", clear_on_submit=True):
    q_week = st.number_input("Week", min_value=1, max_value=25, value=1, step=1)
    q_player = st.text_input("Player")
    q_status = st.selectbox("Status", ["Questionable", "Doubtful", "Out", "IR", "Active"])
    q_injury = st.text_input("Injury Detail (e.g., hamstring)")
    q_note = st.text_area("Notes")
    submitted = st.form_submit_button("Add/Update Injury")
    if submitted:
        row = pd.DataFrame([{
            "Week": int(q_week), "Player": q_player.strip(),
            "Status": q_status, "Injury": q_injury.strip(), "Notes": q_note.strip()
        }])
        # Dedup by Week+Player to overwrite if status changes
        append_to_excel(row, "Injuries", deduplicate_on=["Week", "Player"])
        st.success(f"Injury saved for Week {int(q_week)}: {q_player} – {q_status}")

# =========================
# Fetchers (best-effort, free)
# =========================
st.header("🔄 Fetch (free best-effort)")

fcol1, fcol2 = st.columns(2)
with fcol1:
    wk_fetch = st.number_input("Week to Fetch (2025 season)", 1, 25, 1, 1, key="fetch_week")
    if st.button("Fetch Weekly Team Metrics (nfl_data_py)"):
        try:
            import nfl_data_py as nfl
            weekly = nfl.import_weekly_data([2025])
            chi = weekly[(weekly["team"] == "CHI") & (weekly["week"] == int(wk_fetch))].copy()
            if chi.empty:
                st.warning("No weekly row yet for CHI.")
            else:
                # crude mappings for Offense
                # YPA: passing_yards / attempts
                pass_yards = chi["passing_yards"].iloc[0] if "passing_yards" in chi.columns else None
                attempts = None
                for c in ["attempts", "passing_attempts", "pass_attempts"]:
                    if c in chi.columns:
                        attempts = chi[c].iloc[0]
                        break
                ypa = round(pass_yards / attempts, 2) if pass_yards is not None and attempts not in (None, 0) else None

                completions = None
                for c in ["completions", "passing_completions", "pass_completions"]:
                    if c in chi.columns:
                        completions = chi[c].iloc[0]
                        break
                cmp_pct = round((completions / attempts) * 100, 1) if completions not in (None, 0) and attempts not in (None, 0) else None

                yds_total = None
                for c in ["yards", "total_yards", "offense_yards"]:
                    if c in chi.columns:
                        yds_total = chi[c].iloc[0]
                        break

                off_row = pd.DataFrame([{"Week": int(wk_fetch), "YDS": yds_total, "YPA": ypa, "CMP%": cmp_pct}])
                append_to_excel(off_row, "Offense", deduplicate_on=["Week"])

                # crude mappings for Defense
                sacks = None
                for c in ["sacks", "defense_sacks"]:
                    if c in chi.columns:
                        sacks = chi[c].iloc[0]
                        break
                def_row = pd.DataFrame([{"Week": int(wk_fetch), "SACK": sacks}])
                append_to_excel(def_row, "Defense", deduplicate_on=["Week"])

                # keep raw
                chi2 = chi.copy()
                chi2["Week"] = int(wk_fetch)
                append_to_excel(chi2, "Raw_Weekly", deduplicate_on=["Week"])
                st.success(f"Fetched & saved CHI week {int(wk_fetch)} Offense/Defense (basic).")
        except Exception as e:
            st.error(f"Fetch failed: {e}")

with fcol2:
    pbp_wk = st.number_input("Week for PBP-derived Defense", 1, 25, 1, 1, key="pbp_week")
    if st.button("Fetch PBP → RZ% Allowed / Success Rate / Pressures"):
        try:
            import nfl_data_py as nfl
            pbp = nfl.import_pbp_data([2025], downcast=False)
            w = int(pbp_wk)
            df = pbp[(pbp["week"] == w) & (pbp["defteam"] == "CHI")].copy()
            if df.empty:
                st.warning("No CHI defensive PBP yet for that week.")
            else:
                # Red zone drives
                dmins = (
                    df.groupby(["game_id", "drive"], as_index=False)["yardline_100"]
                    .min()
                    .rename(columns={"yardline_100": "min_yardline_100"})
                )
                total_drives = len(dmins)
                rz_drives = len(dmins[dmins["min_yardline_100"] <= 20]) if total_drives > 0 else 0
                rz_allowed = round((rz_drives / total_drives) * 100, 1) if total_drives else 0.0

                # Success rate (offense success vs CHI D)
                def success_flag(row):
                    if pd.isna(row.get("down")) or pd.isna(row.get("ydstogo")) or pd.isna(row.get("yards_gained")):
                        return False
                    d = int(row["down"]); togo = float(row["ydstogo"]); gain = float(row["yards_gained"])
                    if d == 1:
                        return gain >= 0.4 * togo
                    elif d == 2:
                        return gain >= 0.6 * togo
                    else:
                        return gain >= togo

                real = df[(~df["play_type"].isin(["no_play"])) & (~df["penalty"].fillna(False))].copy()
                sr = round(real.apply(success_flag, axis=1).mean() * 100, 1) if len(real) else 0.0

                qb_hits = df["qb_hit"].fillna(0).astype(int).sum() if "qb_hit" in df.columns else 0
                sacks = df["sack"].fillna(0).astype(int).sum() if "sack" in df.columns else 0
                pressures = int(qb_hits + sacks)

                out = pd.DataFrame([{
                    "Week": w, "RZ% Allowed": rz_allowed, "Success Rate% (Offense)": sr, "Pressures": pressures
                }])
                append_to_excel(out, "Advanced_Defense", deduplicate_on=["Week"])
                st.success(f"Saved Week {w} Advanced Defense: RZ%={rz_allowed} | SR%={sr} | Pressures={pressures}")
        except Exception as e:
            st.error(f"PBP fetch failed: {e}")

    # Snap Count fetch placeholder (free sources are limited without scraping)
    if st.button("Fetch Snap Counts (best effort)"):
        st.info("Snap counts fetch is a placeholder (free sources vary). Upload CSVs to keep this current.")

    # Opponent Preview fetch placeholder
    if st.button("Fetch Opponent Preview (best effort)"):
        st.info("Opponent preview fetch is a placeholder. Upload your preview CSV for richer context.")

# =========================
# DVOA-like Proxy (Opponent-adjusted EPA/SR)
# =========================
st.header("📈 Compute DVOA-like Proxy (Opponent-Adjusted)")

proxy_wk = st.number_input("Week to compute", 1, 25, 1, 1, key="proxy_wk")
if st.button("Compute DVOA-like Proxy"):
    try:
        import nfl_data_py as nfl
        wk = int(proxy_wk)
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
            # find opponent(s)
            opps = set()
            if not bears_off.empty:
                opps.update(bears_off["defteam"].unique().tolist())
            if not bears_def.empty:
                opps.update(bears_def["posteam"].unique().tolist())
            opponent = list(opps)[0] if opps else "UNK"

            prior = plays[plays["week"] < wk].copy()

            def success_flag(down, togo, gain):
                try:
                    if pd.isna(down) or pd.isna(togo) or pd.isna(gain):
                        return False
                    d = int(down); t = float(togo); g = float(gain)
                    if d == 1:
                        return g >= 0.4 * t
                    elif d == 2:
                        return g >= 0.6 * t
                    else:
                        return g >= t
                except Exception:
                    return False

            # Opponent defense baseline (allowed)
            opp_def = prior[prior["defteam"] == opponent].copy()
            opp_def_epa = opp_def["epa"].mean() if len(opp_def) else None
            opp_def_sr = opp_def.apply(lambda r: success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_def) else None

            # Opponent offense baseline (their offense)
            opp_off = prior[prior["posteam"] == opponent].copy()
            opp_off_epa = opp_off["epa"].mean() if len(opp_off) else None
            opp_off_sr = opp_off.apply(lambda r: success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(opp_off) else None

            # CHI week values
            chi_off_epa = bears_off["epa"].mean() if len(bears_off) else None
            chi_off_sr = bears_off.apply(lambda r: success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(bears_off) else None

            chi_def_epa_allowed = bears_def["epa"].mean() if len(bears_def) else None
            chi_def_sr_allowed = bears_def.apply(lambda r: success_flag(r.get("down"), r.get("ydstogo"), r.get("yards_gained")), axis=1).mean() if len(bears_def) else None

            def safe_diff(a, b):
                if a is None or pd.isna(a) or b is None or pd.isna(b):
                    return None
                return float(a) - float(b)

            off_adj_epa = safe_diff(chi_off_epa, opp_def_epa)
            off_adj_sr = safe_diff(chi_off_sr, opp_def_sr)
            def_adj_epa = safe_diff(chi_def_epa_allowed, opp_off_epa)
            def_adj_sr = safe_diff(chi_def_sr_allowed, opp_off_sr)

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

            append_to_excel(out, "DVOA_Proxy", deduplicate_on=["Week"])
            st.success(f"Saved DVOA-like proxy for Week {wk} vs {opponent}.")
    except Exception as e:
        st.error(f"Proxy compute failed: {e}")

# Preview of proxy if present
df_proxy_prev = read_sheet("DVOA_Proxy")
if not df_proxy_prev.empty:
    st.subheader("📊 DVOA-like Proxy (preview)")
    st.dataframe(df_proxy_prev.sort_values("Week").tail(6))

# =========================
# Weekly Prediction
# =========================
st.header("🔮 Weekly Game Prediction")
pred_week = st.number_input("Select Week", 1, 25, 1, 1, key="pred_wk")

if st.button("Run Prediction"):
    try:
        df_s = read_sheet("Strategy")
        df_o = read_sheet("Offense")
        df_d = read_sheet("Defense")
        df_a = read_sheet("Advanced_Defense")
        df_p = read_sheet("DVOA_Proxy")

        row_s = df_s[df_s["Week"] == pred_week] if not df_s.empty and "Week" in df_s else pd.DataFrame()
        row_o = df_o[df_o["Week"] == pred_week] if not df_o.empty and "Week" in df_o else pd.DataFrame()
        row_d = df_d[df_d["Week"] == pred_week] if not df_d.empty and "Week" in df_d else pd.DataFrame()
        row_a = df_a[df_a["Week"] == pred_week] if not df_a.empty and "Week" in df_a else pd.DataFrame()
        row_p = df_p[df_p["Week"] == pred_week] if not df_p.empty and "Week" in df_p else pd.DataFrame()

        if row_s.empty or row_o.empty or row_d.empty:
            st.info("Please upload or fetch Strategy, Offense, and Defense for that week first.")
        else:
            # fields
            def _gf(r, k):
                try:
                    v = r.iloc[0].get(k)
                    return None if (isinstance(v, float) and pd.isna(v)) else v
                except Exception:
                    return None

            strat_text = " ".join([str(v) for v in row_s.iloc[0].astype(str).tolist()]).lower()

            ypa = _gf(row_o, "YPA")
            rz_allowed = _gf(row_a, "RZ% Allowed") if not row_a.empty else _gf(row_d, "RZ% Allowed")
            pressures = _gf(row_a, "Pressures") if not row_a.empty else None

            off_adj_epa = _gf(row_p, "Off Adj EPA/play")
            def_adj_epa = _gf(row_p, "Def Adj EPA/play")

            # Rules
            reason_bits = []
            if off_adj_epa is not None and off_adj_epa >= 0.15 and def_adj_epa is not None and def_adj_epa <= -0.05:
                prediction = "Win – efficiency edge on both sides"
                reason_bits += [f"Off+{off_adj_epa:+.2f} EPA/play vs opp D", f"Def{def_adj_epa:+.2f} EPA/play vs opp O"]
            elif pressures is not None and pressures >= 8 and any(k in strat_text for k in ["blitz", "pressure"]):
                prediction = "Win – pass rush advantage"
                reason_bits.append(f"Pressures={int(pressures)}")
                if rz_allowed is not None:
                    reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif rz_allowed is not None and rz_allowed < 50 and any(k in strat_text for k in ["zone", "two-high", "split-safety"]):
                prediction = "Win – red zone + coverage advantage"
                reason_bits.append(f"RZ% Allowed={rz_allowed:.0f}")
            elif off_adj_epa is not None and off_adj_epa <= -0.10 and rz_allowed is not None and rz_allowed > 65:
                prediction = "Loss – inefficient offense & poor red zone defense"
                reason_bits += [f"Off{off_adj_epa:+.2f} EPA/play", f"RZ% Allowed={rz_allowed:.0f}"]
            elif ypa is not None and ypa < 6 and rz_allowed is not None and rz_allowed > 65:
                prediction = "Loss – inefficient passing & weak red zone defense"
                reason_bits += [f"YPA={float(ypa):.1f}", f"RZ% Allowed={rz_allowed:.0f}"]
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

            st.success(f"**Predicted Outcome for Week {int(pred_week)}: {prediction}**")
            if reason_bits:
                st.caption(" | ".join(reason_bits))

            pred_row = pd.DataFrame([{
                "Week": int(pred_week),
                "Prediction": prediction.split("–")[0].strip(),
                "Reason": prediction.split("–")[1].strip() if "–" in prediction else "",
                "Notes": " | ".join(reason_bits)
            }])
            append_to_excel(pred_row, "Predictions", deduplicate_on=["Week"])
    except Exception as e:
        st.error(f"Prediction failed: {e}")

# Preview sections
st.subheader("📈 Saved Game Predictions")
df_preds = read_sheet("Predictions")
if not df_preds.empty:
    st.dataframe(df_preds.sort_values("Week"))

# =========================
# Beat Writer / ESPN Summary
# =========================
st.header("📰 Weekly Media Summary")
with st.form("media_form", clear_on_submit=True):
    m_week = st.number_input("Week", 1, 25, 1, 1, key="media_week")
    m_opp = st.text_input("Opponent")
    m_text = st.text_area("Summary / Notes")
    m_submit = st.form_submit_button("Save Summary")
    if m_submit:
        row = pd.DataFrame([{"Week": int(m_week), "Opponent": m_opp.strip(), "Summary": m_text.strip()}])
        append_to_excel(row, "Media_Summaries", deduplicate_on=["Week"])
        st.success(f"Summary saved for Week {int(m_week)}.")

df_media_prev = read_sheet("Media_Summaries")
if not df_media_prev.empty:
    st.subheader("🗂️ Saved Media Summaries")
    st.dataframe(df_media_prev.sort_values("Week"))

# =========================
# Download buttons
# =========================
st.sidebar.header("⬇️ Downloads")
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="Download All Data (Excel)",
            data=f,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.sidebar.caption("Tip: The Excel workbook contains all sheets: Offense, Defense, Strategy, Personnel, Injuries, SnapCounts, Advanced_Defense, DVOA_Proxy, Predictions, Media_Summaries, Opponent_Preview, Raw_Weekly.")
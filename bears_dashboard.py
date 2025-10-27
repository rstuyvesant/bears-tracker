# dashboard.py
import io
import datetime as dt
from typing import List, Tuple, Optional

import streamlit as st
import pandas as pd
import numpy as np

# External data package (nflverse)
import nfl_data_py as nfl


# =============================
# App Config & Session Storage
# =============================
st.set_page_config(
    page_title="Bears Weekly Tracker",
    page_icon="🐻",
    layout="wide"
)

if "approved_weeks" not in st.session_state:
    st.session_state.approved_weeks = pd.DataFrame()  # holds approved rows for CHI/OPP and snap counts
if "latest_fetch" not in st.session_state:
    st.session_state.latest_fetch = {}  # temp store for fetched-but-not-approved frames


# ==============
# Constants
# ==============
TEAM = "CHI"
SEASON_DEFAULT = dt.datetime.now().year


# ==============
# Helpers
# ==============
def normalize_week_col(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        if str(c).lower() == "week":
            return df.rename(columns={c: "Week"})
    return df


def get_opponent_menu(season: int, team: str) -> List[str]:
    """
    Build opponent codes for the season from schedule.
    Avoids FutureWarning by using Index.unique().
    """
    try:
        sched = nfl.import_schedules([season])
        sched = sched[sched["home_team"].notna() & sched["away_team"].notna()]
        mask = (sched["home_team"] == team) | (sched["away_team"] == team)
        opps = []
        for _, r in sched[mask].iterrows():
            if r["home_team"] == team:
                opps.append(r["away_team"])
            else:
                opps.append(r["home_team"])
        opps = pd.Index(opps).unique().tolist()
        opps = sorted(opps)
        return opps if opps else ["GB", "MIN", "DET"]
    except Exception:
        return ["GB", "MIN", "DET"]


def _safe_ratio(num: pd.Series, den: pd.Series) -> pd.Series:
    try:
        return num.astype(float) / den.replace(0, np.nan)
    except Exception:
        return pd.Series(np.nan, index=num.index if isinstance(num, pd.Series) else None)


# ------------------------------------------------------------------------------------
# ROBUST WEEKLY IMPORT (primary: weekly endpoint; fallback: play-by-play aggregation)
# ------------------------------------------------------------------------------------
def import_weekly_player_data(season: int) -> pd.DataFrame:
    """
    Try nfl_data_py weekly endpoint. If it fails (404, moved, etc.), fall back to PBP
    and return a "player-like" frame we can aggregate to team totals.
    """
    try:
        return nfl.import_weekly_data([season])
    except Exception:
        # Fallback: construct team offense from play-by-play
        pbp = nfl.import_pbp([season])
        # Keep scrimmage plays with defined type
        pbp = pbp[pbp.get("play_type").notna()]

        # Group by offense team (posteam) vs defense team (defteam)
        g = pbp.groupby(["season", "week", "posteam", "defteam"], dropna=False)
        agg = pd.DataFrame({
            "passing_yards": g["passing_yards"].sum(min_count=1),
            "pass_attempts": g["pass_attempt"].sum(min_count=1),
            "completions": g["complete_pass"].sum(min_count=1),
            "interceptions": g["interception"].sum(min_count=1),
            "rushing_yards": g["rushing_yards"].sum(min_count=1),
            "rush_attempts": g["rush_attempt"].sum(min_count=1),
            "fumbles_lost": g["fumble_lost"].sum(min_count=1),
            "epa": g["epa"].sum(min_count=1),
        }).reset_index()

        agg = agg.rename(columns={
            "posteam": "team",
            "defteam": "opponent_team"
        })
        return agg


def team_week_totals_from_players(weekly_players: pd.DataFrame, season: int, week: int) -> pd.DataFrame:
    """
    Aggregate (player-like) weekly data to team totals + derived rates.
    Works for both the true weekly table and our PBP fallback above.
    """
    df = weekly_players.copy()
    df = df[(df["season"] == season) & (df["week"] == week)]
    if df.empty:
        return pd.DataFrame()

    cols = df.columns.str.lower()

    def pick_col(*cands):
        for c in cands:
            if c.lower() in cols:
                return df.columns[list(cols).index(c.lower())]
        return None

    team_col = pick_col("team", "recent_team", "posteam")
    opp_col  = pick_col("opponent_team", "opp_team", "defteam")
    if team_col is None or opp_col is None:
        raise RuntimeError("Expected team/opponent columns not found in weekly dataset/fallback.")

    pass_yds = pick_col("passing_yards", "pass_yds")
    pass_att = pick_col("attempts", "pass_attempts", "pass_att")
    pass_cmp = pick_col("completions", "pass_completions", "cmp")
    interceptions = pick_col("interceptions", "int")
    rush_yds = pick_col("rushing_yards", "rush_yds")
    rush_att = pick_col("rush_attempts", "rushing_attempts", "ra")
    fumbles_lost = pick_col("fumbles_lost", "fum_lost")
    points = pick_col("points")  # may not exist
    epa = pick_col("epa", "total_epa")

    # Aggregate to team totals
    agg_map = {}
    for c in [pass_yds, pass_att, pass_cmp, interceptions, rush_yds, rush_att, fumbles_lost, points, epa]:
        if c is not None:
            agg_map[c] = "sum"

    grouped = df.groupby([team_col, opp_col], dropna=False).agg(agg_map).reset_index()
    grouped = grouped.rename(columns={team_col: "Team", opp_col: "Opponent"})
    grouped["Season"] = season
    grouped["Week"] = week

    # Derived rates
    if pass_yds and pass_att:
        grouped["YPA"] = _safe_ratio(grouped[pass_yds], grouped[pass_att])
    else:
        grouped["YPA"] = np.nan

    if rush_yds and rush_att:
        grouped["YPC"] = _safe_ratio(grouped[rush_yds], grouped[rush_att])
    else:
        grouped["YPC"] = np.nan

    if pass_cmp and pass_att:
        grouped["CMP%"] = _safe_ratio(grouped[pass_cmp], grouped[pass_att]) * 100.0
    else:
        grouped["CMP%"] = np.nan

    # Tidy labels
    rename_hint = {}
    if pass_yds: rename_hint[pass_yds] = "PassYds"
    if pass_att: rename_hint[pass_att] = "PassAtt"
    if pass_cmp: rename_hint[pass_cmp] = "PassCmp"
    if interceptions: rename_hint[interceptions] = "INT"
    if rush_yds: rename_hint[rush_yds] = "RushYds"
    if rush_att: rename_hint[rush_att] = "RushAtt"
    if fumbles_lost: rename_hint[fumbles_lost] = "FumLost"
    if points: rename_hint[points] = "Pts(approx)"
    if epa: rename_hint[epa] = "EPA(sum)"
    grouped = grouped.rename(columns=rename_hint)

    # Reorder
    front = ["Season", "Week", "Team", "Opponent",
             "PassYds", "RushYds", "PassAtt", "RushAtt", "PassCmp",
             "YPA", "YPC", "CMP%", "INT", "FumLost", "Pts(approx)", "EPA(sum)"]
    front = [c for c in front if c in grouped.columns]
    grouped = grouped[front + [c for c in grouped.columns if c not in front]]
    return grouped


def compute_defense_from_opponent(off_df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert offense to the opponent's defense allowed (i.e., what CHI allowed = opponent's offense).
    """
    if off_df.empty:
        return off_df
    df = off_df.copy()
    df["DefenseTeam"] = df["Opponent"]
    df["Opponent"] = df["Team"]
    df["Team"] = df["DefenseTeam"]
    df = df.drop(columns=["DefenseTeam"])
    rename_map = {}
    for col in ["PassYds", "RushYds", "YPA", "YPC", "CMP%", "INT", "FumLost", "Pts(approx)", "EPA(sum)"]:
        if col in df.columns:
            rename_map[col] = f"{col} Allowed"
    df = df.rename(columns=rename_map)
    return df


def compute_league_averages(offense_df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute per-week NFL averages from team offense totals.
    """
    if offense_df.empty:
        return pd.DataFrame()
    metric_cols = [c for c in offense_df.columns if c not in ["Season", "Week", "Team", "Opponent"]]
    avg = offense_df.groupby(["Season", "Week"], dropna=False)[metric_cols].mean(numeric_only=True).reset_index()
    avg["Team"] = "NFL_AVG"
    avg["Opponent"] = "—"
    return avg


# -----------------------------------------
# Snap Counts (try library; robust fallback)
# -----------------------------------------
def try_import_snap_counts(season: int) -> pd.DataFrame:
    """
    Try a few nfl_data_py functions; return empty if none available.
    """
    candidates = [
        "import_snap_counts",
        "import_weekly_snap_counts",
        "import_pfr_snap_counts",
        "import_participation",   # may carry participation/snap-like info
    ]
    for fn in candidates:
        try:
            f = getattr(nfl, fn)
            df = f([season])
            return df
        except Exception:
            continue
    return pd.DataFrame()


def filter_team_week(df: pd.DataFrame, team: str, season: int, week: int) -> pd.DataFrame:
    if df.empty:
        return df
    cols = df.columns.str.lower()
    # team col
    for c in ["team", "recent_team", "club", "posteam", "defteam"]:
        if c in cols:
            team_col = df.columns[list(cols).index(c)]
            break
    else:
        team_col = None

    wk_col = "week" if "week" in cols else None
    ssn_col = "season" if "season" in cols else None

    q = df.copy()
    if ssn_col:
        q = q[q[ssn_col] == season]
    if wk_col:
        q = q[q[wk_col] == week]
    if team_col:
        q = q[q[team_col] == team]
    return q


def fetch_snaps_from_participation(season: int, week: int, team: str) -> pd.DataFrame:
    """
    Best-effort snaps from nfl_data_py participation; if missing, read nflverse participation parquet (GitHub).
    Returns columns: Season, Week, Team, Player, Position, Snaps, Snap%
    """
    # 1) Try participation-like endpoints first
    part = pd.DataFrame()
    for fn_name in ["import_participation", "import_weekly_participation", "import_pfr_snap_counts", "import_snap_counts"]:
        try:
            f = getattr(nfl, fn_name)
            part = f([season])
            if isinstance(part, pd.DataFrame) and not part.empty:
                break
        except Exception:
            part = pd.DataFrame()
            continue

    # 2) If still empty, pull parquet from GitHub
    if part.empty:
        # Modern path
        for url in [
            f"https://github.com/nflverse/nflfastR-data/raw/master/data/participation/participation_{season}.parquet",
            f"https://github.com/nflverse/nflfastR-data/raw/master/data/participation_{season}.parquet",  # older
        ]:
            try:
                part = pd.read_parquet(url)
                if not part.empty:
                    break
            except Exception:
                continue

    if part.empty:
        return pd.DataFrame()

    df = part.copy()
    cols = df.columns.str.lower()

    def pick(*cands):
        for c in cands:
            if c.lower() in cols:
                return df.columns[list(cols).index(c.lower())]
        return None

    team_col = pick("team", "posteam", "recent_team", "club")
    wk_col = pick("week")
    ssn_col = pick("season")
    player_col = pick("player_name", "player", "name")
    pos_col = pick("position", "pos")
    off_snaps_col = pick("offense", "offense_snaps", "offense_play_count")
    off_pct_col = pick("offense_pct", "offense_percentage")

    if ssn_col is not None:
        df = df[df[ssn_col] == season]
    if wk_col is not None:
        df = df[df[wk_col] == week]
    if team_col is not None:
        df = df[df[team_col] == team]

    out = pd.DataFrame()
    out["Season"] = season
    out["Week"] = week
    out["Team"] = team
    out["Player"] = df[player_col] if player_col in df.columns else np.nan
    out["Position"] = df[pos_col] if pos_col in df.columns else np.nan
    out["Snaps"] = pd.to_numeric(df[off_snaps_col], errors="coerce") if off_snaps_col in df.columns else np.nan

    if off_pct_col in df.columns:
        out["Snap%"] = pd.to_numeric(df[off_pct_col], errors="coerce")
    else:
        # Compute percentage if only Snaps are present
        try:
            total = out["Snaps"].astype(float).sum()
            out["Snap%"] = (out["Snaps"].astype(float) / total * 100.0).round(1) if total else np.nan
        except Exception:
            out["Snap%"] = np.nan

    return out.dropna(subset=["Player"], how="all").reset_index(drop=True)


def to_excel_download(dataframes: List[Tuple[str, pd.DataFrame]], fname: str = "bears_tracker.xlsx") -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for sheet, df in dataframes:
            sheet_name = sheet[:31] if len(sheet) > 31 else sheet
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    bio.seek(0)
    return bio.read()


def make_blank_template(kind: str, season: int) -> pd.DataFrame:
    weeks = list(range(1, 18))
    if kind == "offense":
        cols = ["Season", "Week", "Team", "Opponent", "PassYds", "RushYds", "PassAtt", "RushAtt",
                "PassCmp", "YPA", "YPC", "CMP%", "INT", "FumLost", "Pts(approx)", "EPA(sum)"]
    elif kind == "defense":
        cols = ["Season", "Week", "Team", "Opponent", "PassYds Allowed", "RushYds Allowed", "YPA Allowed",
                "YPC Allowed", "CMP% Allowed", "INT Allowed", "FumLost Allowed", "Pts Allowed (approx)",
                "EPA Allowed (sum)"]
    elif kind == "snaps":
        cols = ["Season", "Week", "Team", "Player", "Position", "Snaps", "Snap%"]
    else:
        cols = ["Season", "Week"]
    df = pd.DataFrame([{c: None for c in cols} for _ in weeks])
    df["Season"] = season
    df["Week"] = weeks
    if "Team" in df.columns:
        df["Team"] = TEAM
    return df


# ==========================
# UI - Controls / Templates
# ==========================
st.title("🐻 Bears Weekly Tracker — Resilient Build")

with st.sidebar:
    st.header("Weekly Controls")
    season = st.number_input("Season", min_value=2012, max_value=2100, value=SEASON_DEFAULT, step=1)
    week = st.number_input("Week", min_value=1, max_value=22, value=1, step=1)
    opp_list = get_opponent_menu(season, TEAM)
    opponent = st.selectbox("Opponent", options=opp_list, index=0)

    st.divider()
    st.header("Templates (CSV)")
    t1 = make_blank_template("offense", season)
    t2 = make_blank_template("defense", season)
    t3 = make_blank_template("snaps", season)

    st.download_button(
        "⬇️ Bears Offense (17 weeks CSV)",
        data=t1.to_csv(index=False).encode("utf-8"),
        file_name=f"bears_offense_template_{season}.csv",
        mime="text/csv",
        use_container_width=True
    )
    st.download_button(
        "⬇️ Bears Defense (17 weeks CSV)",
        data=t2.to_csv(index=False).encode("utf-8"),
        file_name=f"bears_defense_template_{season}.csv",
        mime="text/csv",
        use_container_width=True
    )
    st.download_button(
        "⬇️ CHI Snap Counts (17 weeks CSV)",
        data=t3.to_csv(index=False).encode("utf-8"),
        file_name=f"bears_snaps_template_{season}.csv",
        mime="text/csv",
        use_container_width=True
    )

st.caption("Tip: Use the buttons below to auto-fetch. You’ll review & approve before saving.")


# =================
# Fetch Buttons
# =================
colA, colB, colC = st.columns([1, 1, 1])

with colA:
    fetch_btn = st.button("🔄 Fetch CHI & Opp team stats (Weekly)")
with colB:
    fetch_snaps_btn = st.button("🔄 Fetch CHI snap counts (Weekly)")
with colC:
    recompute_avg_btn = st.button("🧮 Recompute NFL weekly averages (from fetched data)")

latest = st.session_state.latest_fetch


# ==========================================
# Fetch: CHI + OPP offense + derived defense
# ==========================================
if fetch_btn:
    try:
        wp = import_weekly_player_data(season)
        team_totals = team_week_totals_from_players(wp, season, week)

        off_chi = team_totals[team_totals["Team"] == TEAM].copy()
        off_opp = team_totals[team_totals["Team"] == opponent].copy()

        def_chi = compute_defense_from_opponent(off_opp)  # what CHI allowed (opp offense)
        def_opp = compute_defense_from_opponent(off_chi)  # what OPP allowed (CHI offense)

        latest["off_chi"] = off_chi.reset_index(drop=True)
        latest["off_opp"] = off_opp.reset_index(drop=True)
        latest["def_chi"] = def_chi.reset_index(drop=True)
        latest["def_opp"] = def_opp.reset_index(drop=True)

        if off_chi.empty and off_opp.empty:
            st.warning("Fetched, but no rows for this week/team yet (bye week or data not available).")
        else:
            st.success("Fetched weekly team totals for CHI and opponent.")
    except Exception as e:
        st.error(f"Could not fetch weekly team totals: {e}")


# ===========================
# Fetch: CHI snap counts
# ===========================
if fetch_snaps_btn:
    try:
        snaps_all = try_import_snap_counts(season)
        snaps_chi = filter_team_week(snaps_all, TEAM, season, week)
        if snaps_chi.empty:
            # Robust fallback from nflverse parquet
            snaps_chi = fetch_snaps_from_participation(season, week, TEAM)

        latest["snaps_chi"] = snaps_chi.reset_index(drop=True)
        if snaps_chi.empty:
            st.warning("No snap counts found for this week from any source. Use the Snap template to upload instead.")
        else:
            st.success("Fetched weekly CHI snap counts.")
    except Exception as e:
        st.error(f"Could not fetch snap counts: {e}")


# ==========================================
# Recompute NFL per-week averages (offense)
# ==========================================
if recompute_avg_btn:
    try:
        wp = import_weekly_player_data(season)
        tt = team_week_totals_from_players(wp, season, week)  # all teams
        nfl_avg = compute_league_averages(tt)
        latest["nfl_avg_week"] = nfl_avg
        if nfl_avg.empty:
            st.warning("Computed NFL averages, but no team rows available for that week.")
        else:
            st.success("Recomputed NFL per-week averages.")
    except Exception as e:
        st.error(f"Could not compute NFL weekly averages: {e}")


# ==================
# Review & Approve
# ==================
st.subheader("Review & Approve — This Week’s Data")

with st.expander(f"👀 Preview fetched tables for Season {season}, Week {week}"):
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Bears Offense (This Week)**")
        st.dataframe(latest.get("off_chi", pd.DataFrame()))
        st.markdown("**Opponent Offense (This Week)**")
        st.dataframe(latest.get("off_opp", pd.DataFrame()))
    with col2:
        st.markdown("**Bears Defense (This Week)**")
        st.dataframe(latest.get("def_chi", pd.DataFrame()))
        st.markdown("**Opponent Defense (This Week)**")
        st.dataframe(latest.get("def_opp", pd.DataFrame()))
    st.markdown("**CHI Snap Counts (This Week)**")
    st.dataframe(latest.get("snaps_chi", pd.DataFrame()))

approve = st.button("✅ Approve & Save this Week")
if approve:
    frames = []
    for key in ["off_chi", "off_opp", "def_chi", "def_opp", "snaps_chi"]:
        df = latest.get(key, pd.DataFrame())
        if not df.empty:
            frames.append(df.copy())
    if frames:
        approved = pd.concat([st.session_state.approved_weeks] + frames, ignore_index=True)
        # Prevent duplicates by common keys (include Player for snaps if present)
        subset = [c for c in ["Season", "Week", "Team", "Opponent", "Player"] if c in approved.columns]
        if subset:
            approved = approved.drop_duplicates(subset=subset, keep="last")
        st.session_state.approved_weeks = approved
        st.success("Saved. Your YTD tables will reflect this immediately.")
    else:
        st.info("Nothing to save (no fetched tables).")


# =============
# YTD Tables
# =============
st.subheader("📊 YTD Summary (Approved Weeks Only)")

aw = st.session_state.approved_weeks.copy()

def is_offense(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    cols = df.columns
    has_off = ("YPA" in cols) or ("YPC" in cols) or ("PassYds" in cols)
    has_allowed = any("Allowed" in c for c in cols)
    return df if (has_off and not has_allowed) else pd.DataFrame()

def is_defense(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    return df if any("Allowed" in c for c in df.columns) else pd.DataFrame()

def is_snaps(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    return df if ("Player" in df.columns or "player_name" in df.columns) else pd.DataFrame()

off_ytd = is_offense(aw)
def_ytd = is_defense(aw)
snaps_ytd = is_snaps(aw)

colY1, colY2 = st.columns(2)
with colY1:
    st.markdown("**YTD — CHI Offense (rows by week)**")
    st.dataframe(off_ytd[off_ytd["Team"] == TEAM] if not off_ytd.empty else pd.DataFrame())

with colY2:
    st.markdown("**YTD — CHI Defense (rows by week)**")
    st.dataframe(def_ytd[def_ytd["Team"] == TEAM] if not def_ytd.empty else pd.DataFrame())

st.markdown("**YTD — CHI Snap Counts**")
if not snaps_ytd.empty:
    snaps_view = snaps_ytd[snaps_ytd.get("Team", TEAM) == TEAM]
    st.dataframe(snaps_view)
else:
    st.dataframe(pd.DataFrame())

# Optional: per-week NFL average from approved offense weeks
try:
    if not off_ytd.empty:
        nfl_avg_from_saved = compute_league_averages(off_ytd.dropna(subset=["Week"]))
        nfl_avg_from_saved = nfl_avg_from_saved.sort_values(["Season", "Week"])
    else:
        nfl_avg_from_saved = pd.DataFrame()
except Exception:
    nfl_avg_from_saved = pd.DataFrame()

st.markdown("**YTD — NFL Offense Averages (from approved weeks)**")
st.dataframe(nfl_avg_from_saved)


# ===============
# Export (Excel)
# ===============
st.subheader("📤 Export")

weekly_tabs = []
for label, key in [
    ("Wk Offense CHI", "off_chi"),
    ("Wk Offense OPP", "off_opp"),
    ("Wk Defense CHI", "def_chi"),
    ("Wk Defense OPP", "def_opp"),
    ("Wk CHI Snaps", "snaps_chi"),
    ("Wk NFL Avg", "nfl_avg_week")
]:
    df = latest.get(key, pd.DataFrame())
    weekly_tabs.append((label, df if not df.empty else pd.DataFrame()))

ytd_tabs = [
    ("YTD Offense", off_ytd if not off_ytd.empty else pd.DataFrame()),
    ("YTD Defense", def_ytd if not def_ytd.empty else pd.DataFrame()),
    ("YTD Snaps", snaps_ytd if not snaps_ytd.empty else pd.DataFrame()),
    ("YTD NFL Avg (from saved)", nfl_avg_from_saved if not nfl_avg_from_saved.empty else pd.DataFrame()),
]

export_bytes = to_excel_download(weekly_tabs + ytd_tabs, fname="bears_tracker.xlsx")
st.download_button(
    label="⬇️ Download Excel (Weekly + YTD)",
    data=export_bytes,
    file_name=f"bears_tracker_{season}_wk{week}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)


# ============================
# Manual Uploads (Optional)
# ============================
st.subheader("📥 Optional: Manual Uploads (if auto-fetch not available)")

up_col1, up_col2, up_col3 = st.columns(3)
with up_col1:
    up_off = st.file_uploader("Upload Bears Offense (CSV/XLSX)", type=["csv", "xlsx"], key="up_off")
with up_col2:
    up_def = st.file_uploader("Upload Bears Defense (CSV/XLSX)", type=["csv", "xlsx"], key="up_def")
with up_col3:
    up_snaps = st.file_uploader("Upload CHI Snap Counts (CSV/XLSX)", type=["csv", "xlsx"], key="up_snaps")

def read_any(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    name = file.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(file)
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"Could not read {file.name}: {e}")
        return pd.DataFrame()

if up_off or up_def or up_snaps:
    manual_frames = []
    m_off = normalize_week_col(read_any(up_off))
    m_def = normalize_week_col(read_any(up_def))
    m_snp = normalize_week_col(read_any(up_snaps))

    if not m_off.empty:
        st.markdown("**Preview — Uploaded Bears Offense**")
        st.dataframe(m_off.head(20))
        manual_frames.append(m_off)
    if not m_def.empty:
        st.markdown("**Preview — Uploaded Bears Defense**")
        st.dataframe(m_def.head(20))
        manual_frames.append(m_def)
    if not m_snp.empty:
        st.markdown("**Preview — Uploaded CHI Snap Counts**")
        st.dataframe(m_snp.head(20))
        manual_frames.append(m_snp)

    if st.button("✅ Approve & Save Uploaded"):
        if manual_frames:
            approved = pd.concat([st.session_state.approved_weeks] + manual_frames, ignore_index=True)
            subset = [c for c in ["Season", "Week", "Team", "Opponent", "Player"] if c in approved.columns]
            if subset:
                approved = approved.drop_duplicates(subset=subset, keep="last")
            st.session_state.approved_weeks = approved
            st.success("Uploaded data saved to YTD.")


# ============================
# Data Collection Notes
# ============================
with st.expander("ℹ️ What data is collected (exactly)?", expanded=False):
    st.markdown("""
- **Team Weekly Offense** (primary `import_weekly_data`; fallback from play-by-play):
  - If present: `PassYds, RushYds, PassAtt, RushAtt, PassCmp, INT, FumLost, Pts(approx), EPA(sum)`.
  - Derived: `YPA, YPC, CMP%`.
- **Team Weekly Defense**: mirrored from the opponent’s offense (e.g., `PassYds Allowed`, etc.).
- **CHI Snap Counts**:
  - Tries one of `import_snap_counts` / `import_weekly_snap_counts` / `import_pfr_snap_counts` / `import_participation`.
  - If unavailable, pulls season **participation parquet** from nflverse GitHub.
- **NFL Weekly Averages**:
  - Mean across **all teams'** weekly offense totals for each **Season, Week**.
- You always **review & approve** before data is added to YTD tables and exports.
""")

import io
import datetime as dt
from typing import List, Tuple, Optional

import streamlit as st
import pandas as pd
import numpy as np

# External data package (nflverse)
import nfl_data_py as nfl


# -----------------------------
# App Config
# -----------------------------
st.set_page_config(
    page_title="Bears Weekly Tracker",
    page_icon="🐻",
    layout="wide"
)

# Initialize session storage
if "approved_weeks" not in st.session_state:
    st.session_state.approved_weeks = pd.DataFrame()  # holds approved (weekly) rows for CHI, OPP, and snap counts
if "latest_fetch" not in st.session_state:
    st.session_state.latest_fetch = {}  # temporary store for fetched-but-not-approved frames


# -----------------------------
# Helpers
# -----------------------------
TEAM = "CHI"
SEASON_DEFAULT = dt.datetime.now().year

TEAM_ALIASES = {
    # minimal alias helper if needed
    "CHI": "CHI"
}

def normalize_week_col(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        if c.lower() == "week":
            df = df.rename(columns={c: "Week"})
            break
    return df

def _safe_float_div(n: float, d: float) -> float:
    try:
        if d == 0:
            return np.nan
        return n / d
    except Exception:
        return np.nan

def get_opponent_menu(season: int, team: str) -> List[str]:
    """Create a list of opponent codes for the season using schedules."""
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
        opps = sorted(list(pd.unique(opps)))
        return opps if opps else ["GB", "MIN", "DET"]
    except Exception:
        return ["GB", "MIN", "DET"]


def import_weekly_player_data(season: int) -> pd.DataFrame:
    """
    Pull weekly player-level data (nflverse weekly). We aggregate to team totals.
    """
    df = nfl.import_weekly_data([season])  # team, opponent_team, week, season, various stats
    # ensure expected fields exist (column naming varies with versions)
    return df


def team_week_totals_from_players(weekly_players: pd.DataFrame, season: int, week: int) -> pd.DataFrame:
    """
    Aggregate player weekly to team totals for offense perspective.
    We’ll compute some common team-level metrics robustly based on columns that exist.
    """
    df = weekly_players.copy()
    df = df[(df["season"] == season) & (df["week"] == week)]

    # Some installs use 'recent_team' instead of 'team'
    team_col = "team" if "team" in df.columns else ("recent_team" if "recent_team" in df.columns else None)
    opp_col = "opponent_team" if "opponent_team" in df.columns else ("opp_team" if "opp_team" in df.columns else None)
    if team_col is None or opp_col is None:
        raise RuntimeError("Expected columns 'team'/'recent_team' and 'opponent_team' not found in weekly data.")

    # Common stat columns in nfl_data_py weekly
    cols = df.columns.str.lower()

    def pick(*cands) -> Optional[str]:
        for c in cands:
            if c.lower() in cols:
                # return original case from df.columns
                return df.columns[list(cols).index(c.lower())]
        return None

    # Offense building blocks
    pass_yds = pick("passing_yards", "pass_yds", "py")
    pass_att = pick("attempts", "pass_attempts", "pass_att")
    pass_cmp = pick("completions", "pass_completions", "cmp")
    interceptions = pick("interceptions", "int")
    rush_yds = pick("rushing_yards", "rush_yds", "ry")
    rush_att = pick("rush_attempts", "rushing_attempts", "ra")
    fumbles_lost = pick("fumbles_lost", "fum_lost")
    points = pick("points", "fantasy_points_ppr", "fantasy_points")  # points may not exist; leave NaN if missing
    epa = pick("epa", "total_epa")  # player EPA sums to team EPA (approx)

    # Third-down conversions sometimes appear only in PBP aggregations; leave NaN if missing
    # Build a dictionary of aggregations
    agg_map = {}
    for col in [pass_yds, pass_att, pass_cmp, interceptions, rush_yds, rush_att, fumbles_lost, points, epa]:
        if col is not None:
            agg_map[col] = "sum"

    grouped = df.groupby([team_col, opp_col], dropna=False).agg(agg_map).reset_index()
    grouped = grouped.rename(columns={team_col: "Team", opp_col: "Opponent"})
    grouped["Season"] = season
    grouped["Week"] = week

    # Derived rates
    if pass_yds is not None and pass_att is not None:
        grouped["YPA"] = grouped[pass_yds].apply(float) / grouped[pass_att].replace(0, np.nan)
    else:
        grouped["YPA"] = np.nan
    if rush_yds is not None and rush_att is not None:
        grouped["YPC"] = grouped[rush_yds].apply(float) / grouped[rush_att].replace(0, np.nan)
    else:
        grouped["YPC"] = np.nan
    if pass_cmp is not None and pass_att is not None:
        grouped["CMP%"] = (grouped[pass_cmp].apply(float) / grouped[pass_att].replace(0, np.nan)) * 100.0
    else:
        grouped["CMP%"] = np.nan

    # Tidy friendly columns
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

    # Reorder key columns up front
    front_cols = ["Season", "Week", "Team", "Opponent", "PassYds", "RushYds", "PassAtt", "RushAtt", "PassCmp",
                  "YPA", "YPC", "CMP%", "INT", "FumLost", "Pts(approx)", "EPA(sum)"]
    front_cols = [c for c in front_cols if c in grouped.columns]
    other_cols = [c for c in grouped.columns if c not in front_cols]
    grouped = grouped[front_cols + other_cols]

    return grouped


def compute_defense_from_opponent(off_df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert an offense table to the corresponding defense table for the opponent.
    (i.e., what CHI allowed equals the opponent's offense)
    """
    if off_df.empty:
        return off_df
    df = off_df.copy()
    df["DefenseTeam"] = df["Opponent"]
    df["Opponent"] = df["Team"]
    df["Team"] = df["DefenseTeam"]
    df = df.drop(columns=["DefenseTeam"])
    # Rename a few columns to "Allowed" if present
    rename_map = {}
    for col in ["PassYds", "RushYds", "YPA", "YPC", "CMP%", "INT", "FumLost", "Pts(approx)", "EPA(sum)"]:
        if col in df.columns:
            rename_map[col] = f"{col} Allowed"
    df = df.rename(columns=rename_map)
    return df


def compute_league_averages(offense_df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute per-week NFL averages from team offense totals (all teams, all games) for the season.
    """
    if offense_df.empty:
        return offense_df
    # Average by Week across teams
    metric_cols = [c for c in offense_df.columns if c not in ["Season", "Week", "Team", "Opponent"]]
    avg = offense_df.groupby(["Season", "Week"], dropna=False)[metric_cols].mean(numeric_only=True).reset_index()
    avg["Team"] = "NFL_AVG"
    avg["Opponent"] = "—"
    return avg


def try_import_snap_counts(season: int) -> pd.DataFrame:
    """
    Try a few function names because nfl_data_py has added/renamed functions across versions.
    Falls back to empty DataFrame if unavailable.
    """
    candidates = [
        "import_snap_counts",
        "import_weekly_snap_counts",
        "import_pfr_snap_counts",
        "import_participation",  # sometimes has participation/snap-like info
    ]
    for fn in candidates:
        try:
            f = getattr(nfl, fn)
            df = f([season])
            # Expect team, season, week, player, snaps or pct
            return df
        except Exception:
            continue
    return pd.DataFrame()


def filter_team_week(df: pd.DataFrame, team: str, season: int, week: int) -> pd.DataFrame:
    if df.empty:
        return df
    cols = df.columns.str.lower()
    # team column guess
    team_cols = ["team", "recent_team", "club", "posteam", "defteam"]
    tc = None
    for c in team_cols:
        if c in cols:
            tc = df.columns[list(cols).index(c)]
            break
    # week column
    wc = "week" if "week" in cols else None
    sc = "season" if "season" in cols else None
    q = df.copy()
    if sc:
        q = q[q[sc] == season]
    if wc:
        q = q[q[wc] == week]
    if tc:
        q = q[q[tc] == team]
    return q


def to_excel_download(dataframes: List[Tuple[str, pd.DataFrame]], fname: str = "bears_tracker.xlsx") -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for sheet, df in dataframes:
            # cap at Excel's sheet name limits
            writer_sheet = sheet[:31] if len(sheet) > 31 else sheet
            df.to_excel(writer, index=False, sheet_name=writer_sheet)
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


# -----------------------------
# UI - Controls
# -----------------------------
st.title("🐻 Bears Weekly Tracker — Clean Build")

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


# -----------------------------
# Fetch Buttons
# -----------------------------
colA, colB, colC = st.columns([1,1,1])

with colA:
    fetch_btn = st.button("🔄 Fetch CHI & Opp team stats (Weekly)")
with colB:
    fetch_snaps_btn = st.button("🔄 Fetch CHI snap counts (Weekly)")
with colC:
    recompute_avg_btn = st.button("🧮 Recompute NFL weekly averages (from fetched data)")

# Storage convenience
latest = st.session_state.latest_fetch


# -----------------------------
# Fetch: CHI, OPP offense/defense
# -----------------------------
if fetch_btn:
    try:
        wp = import_weekly_player_data(season)
        team_totals = team_week_totals_from_players(wp, season, week)

        # Pull CHI row and OPP row
        row_chi = team_totals[team_totals["Team"] == TEAM].copy()
        row_opp = team_totals[team_totals["Team"] == opponent].copy()

        # If we didn't find rows (bye or data lag), keep empty frames
        off_chi = row_chi.reset_index(drop=True)
        off_opp = row_opp.reset_index(drop=True)

        # Derive defenses from opponent offenses (what each allowed)
        def_chi = compute_defense_from_opponent(off_opp)  # what CHI allowed = OPP offense
        def_opp = compute_defense_from_opponent(off_chi)  # what OPP allowed = CHI offense

        latest["off_chi"] = off_chi
        latest["off_opp"] = off_opp
        latest["def_chi"] = def_chi
        latest["def_opp"] = def_opp

        st.success("Fetched weekly team totals for CHI and opponent.")
    except Exception as e:
        st.error(f"Could not fetch weekly team totals: {e}")

# -----------------------------
# Fetch: CHI snap counts
# -----------------------------
if fetch_snaps_btn:
    try:
        snaps_all = try_import_snap_counts(season)
        snaps_chi = filter_team_week(snaps_all, TEAM, season, week)
        latest["snaps_chi"] = snaps_chi.reset_index(drop=True)
        if snaps_chi.empty:
            st.warning("No snap count function available (or no data for this week). Use template CSV to upload instead.")
        else:
            st.success("Fetched weekly CHI snap counts.")
    except Exception as e:
        st.error(f"Could not fetch snap counts: {e}")

# -----------------------------
# Compute NFL weekly averages from what we fetched (all teams in week)
# -----------------------------
if recompute_avg_btn:
    try:
        # For NFL averages we need all teams in the week.
        # We’ll re-pull the weekly player data and aggregate across all teams.
        wp = import_weekly_player_data(season)
        tt = team_week_totals_from_players(wp, season, week)  # all teams
        nfl_avg = compute_league_averages(tt)
        latest["nfl_avg_week"] = nfl_avg
        st.success("Recomputed NFL per-week averages.")
    except Exception as e:
        st.error(f"Could not compute NFL weekly averages: {e}")


# -----------------------------
# Review & Approve
# -----------------------------
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
        # Drop duplicate (Season, Week, Team, Opponent) rows to prevent double-approval
        # (For snap counts, Opponent may not exist; safe subset.)
        subset = [c for c in ["Season", "Week", "Team", "Opponent"] if c in approved.columns]
        approved = approved.drop_duplicates(subset=subset, keep="last")
        st.session_state.approved_weeks = approved
        st.success("Saved. Your YTD tables will reflect this immediately.")
    else:
        st.info("Nothing to save (no fetched tables).")


# -----------------------------
# YTD Tables (built from approved weeks)
# -----------------------------
st.subheader("📊 YTD Summary (Approved Weeks Only)")

aw = st.session_state.approved_weeks.copy()

# Split into offense/defense/snap subsets if present
def is_offense(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    # If it has "YPA" (not Allowed) we’ll consider it offense
    cols = df.columns
    has_off = any(c == "YPA" for c in cols) or any(c == "YPC" for c in cols) or "PassYds" in cols
    # If it has " Allowed", it's defense not offense.
    has_allowed = any("Allowed" in c for c in cols)
    return df if (has_off and not has_allowed) else pd.DataFrame()

def is_defense(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    has_allowed = any("Allowed" in c for c in df.columns)
    return df if has_allowed else pd.DataFrame()

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
st.dataframe(snaps_ytd[snaps_ytd.get("Team", TEAM) == TEAM] if not snaps_ytd.empty else pd.DataFrame())

# Optional simple per-week NFL average from approved weeks (offense only)
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


# -----------------------------
# Weekly + YTD Export (Excel)
# -----------------------------
st.subheader("📤 Export")

# Current week tables to include (even if not yet approved)
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

# YTD tabs (approved only)
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

# -----------------------------
# Manual Uploads (if auto-fetch isn’t available)
# -----------------------------
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
            approved = approved.drop_duplicates(subset=subset, keep="last")
            st.session_state.approved_weeks = approved
            st.success("Uploaded data saved to YTD.")


# -----------------------------
# Minimal Data Notes
# -----------------------------
with st.expander("ℹ️ What data is collected (exactly)?", expanded=False):
    st.markdown("""
- **Team Weekly Offense** (from `nfl_data_py.import_weekly_data` aggregated to team by week):
  - *If present in your package build*: `PassYds, RushYds, PassAtt, RushAtt, PassCmp, INT, FumLost, Pts(approx), EPA(sum)`.
  - Derived: `YPA, YPC, CMP%`.
- **Team Weekly Defense**: mirrored from the opponent’s offense (e.g., `PassYds Allowed`, etc.).
- **CHI Snap Counts** (best-effort):
  - Tries one of: `import_snap_counts`, `import_weekly_snap_counts`, `import_pfr_snap_counts`, or `import_participation`.
  - If none exist, use the provided **Snap Counts template**.
- **NFL Weekly Averages**:
  - Mean across all teams' weekly offense totals for each **Season, Week**.
- You always **review & approve** before data is added to your YTD tables.
""")


# End of app

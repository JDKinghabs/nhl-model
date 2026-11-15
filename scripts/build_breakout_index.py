"""
build_breakout_index.py  -  your Breakout Index: zone exit time + costly giveaways + transition shots
=====================================================================================================
WHAT IT DOES (in Excel terms):
    Builds three "getting out of your own end" statistics for every team from the NHL play-by-play
    (data/pbp.db, made by pull_play_by_play.py), at 5-on-5 in the regular season, and blends them
    into one Breakout Index. Then it tests the index honestly.

    THE THREE STATS (all 5-on-5; "from the team's point of view"):

      1. ZONE EXIT TIME   Every time a team wins the puck in its OWN zone (it wins a defensive-zone faceoff or
                          takes the puck away in its own zone), start a stopwatch. Stop it at the first event that
                          happens OUTSIDE the team's defensive zone (play has left the zone). If the play instead
                          ends first (a shot against, a stoppage such as icing, a penalty, or the team giving the
                          puck away in its own zone), or it takes more than 20 seconds, that attempt counts as 20.
                          Result: the team's average seconds to clear the zone. LOWER IS BETTER.

      2. COSTLY GIVEAWAYS A giveaway in the team's own zone that is followed by an opponent shot attempt within 10
                          seconds (before play stops). Result: costly giveaways per 60 minutes of 5-on-5 play.
                          LOWER IS BETTER.

      3. TRANSITION SHOTS A shot attempt (shot on goal, miss, block or goal) taken within 5 seconds of an event that
                          happened outside the offensive zone, when play was already in motion (not right after a
                          faceoff or a stoppage): the puck came up the ice and went to the net quickly.
                          Result: transition shot attempts per 60 minutes of 5-on-5 play. HIGHER IS BETTER.

    THE INDEX: each stat is converted to "how far above or below the league average that season" (like a z-score,
    flipped so higher is always better), then the three are averaged with equal weight.

    THE TESTS:
      * REPEATS?    does a team's number from odd-numbered games match its number from even-numbered games?
      * LINKS?      how does it line up with points %, goal difference, shot-attempt share (Corsi), expected goals,
                    and (2021-22 on) the NHL's real zone-time numbers (a team that exits quickly should spend
                    less time in its own end)?
      * ADDS?       does adding the index to the contender score (version 1) pick the real conference finalists
                    better, on seasons the score has never seen? (needs 8 or more seasons of play-by-play)
      * CHAMPIONS   where did champions and finalists rank?

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_breakout_index.py
    Safe to run while the download is still going; it uses the seasons that are fully downloaded so far.
    Rerun it when the download finishes to use every season.
    Options:  --seasons 20232024,20242025     only those seasons
              --force-forecast                run the hold-out test even with fewer than 8 seasons (unreliable)

WHAT YOU SHOULD SEE:
    A summary in the terminal and a new file:  outputs/breakout_index.xlsx   (close it in Excel first).

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. Shot-quality columns from MoneyPuck.com
    (non-commercial use, credit required) are used only for comparison.
"""

import argparse
import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_contender_fit as cf          # contender score v1 and the hold-out test helpers
import build_popular_stats_check as ps    # shared stat helpers

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
PBP_DB = PROJECT_FOLDER / "data" / "pbp.db"
NHL_DB = PROJECT_FOLDER / "data" / "nhl.db"
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "breakout_index.xlsx"

# ---- The settings that define the three stats (change these and rerun to try other versions) ----
EXIT_CAP = 20            # seconds: a clearance slower than this (or a failed one) counts as this many
FAST_EXIT = 8            # seconds: "quick exit" threshold, shown as extra information
GIVEAWAY_WINDOW = 10     # seconds after an own-zone giveaway in which an opponent shot attempt makes it "costly"
TRANSITION_WINDOW = 5    # seconds between an event outside the offensive zone and the shot for it to be a "transition shot"
MIN_COVERAGE = 0.98      # a season is used only if this share of its games is downloaded
MIN_SEASONS_FOR_FORECAST = 8
FIVE_ON_FIVE = "1551"    # the NHL's code for 5 skaters each and both goalies in

SHOT_TYPES = ["shot-on-goal", "missed-shot", "blocked-shot", "goal"]
STOP_TYPES = ["stoppage", "penalty", "delayed-penalty", "period-end", "game-end"]
NOT_IN_MOTION = ["faceoff", "stoppage", "period-start", "penalty", "delayed-penalty", "period-end", "game-end", "shootout-complete"]
FLIP = {"O": "D", "D": "O", "N": "N"}

COMPONENTS = [  # (column, plain name, higher is better?)
    ("exit_time", "Zone exit time (seconds, lower is better)", False),
    ("costly_pg", "Costly giveaways per 60 (lower is better)", False),
    ("trans_pg", "Transition shot attempts per 60 (higher is better)", True),
]
CP_ORDER = cf.CHECKPOINTS
PLAIN_CP = cf.PLAIN_CP
OUTCOME_ORDER = cf.OUTCOME_ORDER

READ_ME = [
    "Breakout Index: zone exit time + costly giveaways + transition shots",
    "",
    "WHAT THE THREE STATS MEAN (5-on-5, regular season, from each team's point of view; play-by-play events, not puck tracking)",
    f"  Zone exit time      Start a stopwatch when a team wins the puck in its own zone (defensive-zone faceoff win or takeaway). Stop it at the first",
    f"                      event outside the team's defensive zone. If play ends first (a shot against, a stoppage such as icing, a penalty, or the team",
    f"                      gives the puck away in its own zone) or it takes more than {EXIT_CAP} seconds, that attempt counts as {EXIT_CAP}. Average seconds; lower is better.",
    f"  Costly giveaways    An own-zone giveaway followed by an opponent shot attempt within {GIVEAWAY_WINDOW} seconds (before play stops), per 60 minutes. Lower is better.",
    f"  Transition shots    A shot attempt within {TRANSITION_WINDOW} seconds of an event outside the offensive zone, while play was in motion (not right after a faceoff",
    "                      or stoppage), per 60 minutes. Higher is better.",
    "  Breakout Index      Each stat becomes 'steps above or below the league average that season' (flipped so higher is always better); the three are averaged.",
    "",
    "HOW EVENTS ARE READ",
    "  The NHL tags every event with a zone from the point of view of the team that did it (shots, hits, giveaways, takeaways, faceoffs). Blocked shots are tagged from",
    "  the BLOCKING team's point of view: the script flips that so every event is read from the right team's side. Only events at 5-on-5 start an attempt.",
    "",
    "HOW TO READ THE TESTS",
    "  Repeats    rank correlation (within each season) between a team's number in odd-numbered games and in even-numbered games. High = a real team trait.",
    "  Links      rank correlation with other team numbers. 0.3 is modest, 0.5 strong. Differences under about 0.1 are within noise with 30 teams a season.",
    "  Zone-time check (2021-22 on): a quick-exit team should spend LESS time in its own zone at 5v5: look for a positive link between exit time and defensive-zone time.",
    "  Hold-out test: the contender score is built on all seasons but one and used on the one it has never seen; the index must improve it to count.",
    "",
    "WHAT THIS CANNOT SAY",
    "  Events are sparse (no puck possession or passes), so 'exit' means 'play was next seen outside the zone'. Giveaway and takeaway counts vary a little by arena",
    "  (scorer differences); using the 'costly' version and a league-wide comparison softens that, it does not remove it.",
    "",
    "Data: (c) NHL.",
]


# ----------------------------------------------------------------------------
# Reading the play-by-play of one season
# ----------------------------------------------------------------------------
def season_coverage(pbp, nhl):
    """Games saved vs games played, per season (regular season only)."""
    saved = pd.read_sql_query("SELECT season_id, COUNT(*) AS saved FROM pbp_games WHERE status = 'ok' AND game_type = 2 GROUP BY season_id", pbp)
    played = pd.read_sql_query("SELECT season_id, COUNT(*) AS played FROM games WHERE game_type = 2 AND season_id >= 20092010 GROUP BY season_id", nhl)
    table = played.merge(saved, on="season_id", how="left").fillna({"saved": 0})
    table["coverage"] = table["saved"] / table["played"]
    return table


def load_events(pbp, season_id):
    return pd.read_sql_query(
        """SELECT e.game_id, e.sort_order, e.period, e.time_s, e.situation_code, e.type, e.team_id, e.zone,
                  g.home_team_id, g.away_team_id
           FROM pbp_events e JOIN pbp_games g ON g.game_id = e.game_id
           WHERE g.season_id = ? AND g.game_type = 2 AND g.status = 'ok' AND e.period_type != 'SO'""", pbp, params=[season_id])


def forward_match(left, right, name):
    """For each row of `left`, find the NEXT row of `right` (same game, same team perspective) and attach its order number and time."""
    r = right[["game_id", "T", "sort_order", "abs_t"]].copy()
    r[f"{name}_so"], r[f"{name}_t"] = r["sort_order"], r["abs_t"]
    r = r.drop(columns="abs_t").sort_values("sort_order")
    return pd.merge_asof(left.sort_values("sort_order"), r, on="sort_order", by=["game_id", "T"],
                         direction="forward", allow_exact_matches=False)


def team_game_stats(ev):
    """From one season's events, one row per team per game with the counts the three stats are built from."""
    ev = ev.sort_values(["game_id", "sort_order"]).reset_index(drop=True)
    ev["abs_t"] = (ev["period"] - 1) * 1200 + ev["time_s"]

    # --- every event read from BOTH teams' points of view ---
    home = ev.assign(T=ev["home_team_id"])
    away = ev.assign(T=ev["away_team_id"])
    p = pd.concat([home, away], ignore_index=True)
    p["mine"] = p["team_id"] == p["T"]
    owner_view = p["type"] != "blocked-shot"                         # blocked shots are tagged from the BLOCKER's side
    tag_is_T = p["mine"] == owner_view                               # does the zone tag describe team T's side?
    p["zone_T"] = p["zone"].where(tag_is_T, p["zone"].map(FLIP))
    five = p["situation_code"] == FIVE_ON_FIVE

    # --- 1) zone exit time ---
    starts = p[p["mine"] & five & p["zone_T"].eq("D") & p["type"].isin(["faceoff", "takeaway"])][["game_id", "T", "sort_order", "abs_t"]]
    exits = p[p["zone_T"].isin(["N", "O"])]
    fails = p[p["type"].isin(STOP_TYPES) | (~p["mine"] & p["type"].isin(SHOT_TYPES))
              | (p["mine"] & p["type"].eq("giveaway") & p["zone_T"].eq("D")) | ~five]
    s = forward_match(forward_match(starts, exits, "exit"), fails, "fail")
    took = s["exit_t"] - s["abs_t"]
    ok = s["exit_so"].notna() & (s["fail_so"].isna() | (s["exit_so"] < s["fail_so"])) & (took <= EXIT_CAP)
    s["exit_seconds"] = np.where(ok, took, EXIT_CAP)
    s["exit_ok"], s["exit_fast"] = ok, ok & (took <= FAST_EXIT)
    exit_stats = s.groupby(["game_id", "T"]).agg(exit_starts=("exit_seconds", "size"), exit_sum=("exit_seconds", "sum"),
                                                 exit_ok=("exit_ok", "sum"), exit_fast=("exit_fast", "sum")).reset_index()

    # --- 2) costly giveaways ---
    gives = p[p["mine"] & five & p["type"].eq("giveaway") & p["zone_T"].eq("D")][["game_id", "T", "sort_order", "abs_t"]]
    opp_shots = p[~p["mine"] & p["type"].isin(SHOT_TYPES)]
    stops = p[p["type"].isin(STOP_TYPES)]
    g = forward_match(forward_match(gives, opp_shots, "shot"), stops, "stop")
    costly = g["shot_so"].notna() & ((g["shot_t"] - g["abs_t"]) <= GIVEAWAY_WINDOW) & (g["stop_so"].isna() | (g["shot_so"] < g["stop_so"]))
    g["costly"] = costly
    give_stats = g.groupby(["game_id", "T"]).agg(giveaways=("costly", "size"), costly=("costly", "sum")).reset_index()

    # --- 3) transition shots (and all 5v5 shot attempts, for comparison) ---
    ev["prev_type"] = ev.groupby("game_id")["type"].shift(1)
    ev["prev_t"] = ev.groupby("game_id")["abs_t"].shift(1)
    ev["prev_zone"] = ev.groupby("game_id")["zone"].shift(1)
    ev["prev_team"] = ev.groupby("game_id")["team_id"].shift(1)
    other = ev["home_team_id"] + ev["away_team_id"]
    prev_view = np.where(ev["prev_type"].eq("blocked-shot"), other - ev["prev_team"], ev["prev_team"])
    prev_zone_T = ev["prev_zone"].where(prev_view == ev["team_id"], ev["prev_zone"].map(FLIP))
    shots = ev["situation_code"].eq(FIVE_ON_FIVE) & ev["type"].isin(SHOT_TYPES)
    quick = (ev["abs_t"] - ev["prev_t"] <= TRANSITION_WINDOW) & ~ev["prev_type"].isin(NOT_IN_MOTION) & prev_zone_T.isin(["N", "D"])
    sh = ev[shots].assign(transition=quick[shots], opponent=(other - ev["team_id"])[shots])
    shots_for = sh.groupby(["game_id", "team_id"]).agg(shots_for=("transition", "size"), trans_for=("transition", "sum")).reset_index().rename(columns={"team_id": "T"})
    shots_against = sh.groupby(["game_id", "opponent"]).agg(shots_against=("transition", "size"), trans_against=("transition", "sum")).reset_index().rename(columns={"opponent": "T"})

    out = exit_stats.merge(give_stats, on=["game_id", "T"], how="outer").merge(shots_for, on=["game_id", "T"], how="outer") \
        .merge(shots_against, on=["game_id", "T"], how="outer").fillna(0)
    out["T"] = out["T"].astype(int)
    return out.rename(columns={"T": "team_id"})


def check_blocked_shot_view(ev):
    """Sanity check: read from the shooter's side, blocked shots should be taken in the offensive zone."""
    b = ev[ev["type"] == "blocked-shot"]
    return float(b["zone"].map(FLIP).eq("O").mean()) if len(b) else float("nan")


# ----------------------------------------------------------------------------
# From game rows to team-season numbers and the index
# ----------------------------------------------------------------------------
def add_ice_time(game_rows, nhl):
    toi = pd.read_sql_query("SELECT game_id, team_id, iceTime AS toi5 FROM mp_team_games WHERE situation = '5on5' AND playoffGame = 0", nhl)
    return game_rows.merge(toi, on=["game_id", "team_id"], how="left")


def summarize(rows):
    """One row per team-season (or per whatever groups are present) from per-game counts."""
    g = rows.groupby(["season_id", "team_id"]).agg(
        games=("game_id", "nunique"), exit_starts=("exit_starts", "sum"), exit_sum=("exit_sum", "sum"), exit_ok=("exit_ok", "sum"),
        exit_fast=("exit_fast", "sum"), giveaways=("giveaways", "sum"), costly=("costly", "sum"), trans_for=("trans_for", "sum"),
        trans_against=("trans_against", "sum"), shots_for=("shots_for", "sum"), shots_against=("shots_against", "sum"), toi5=("toi5", "sum")).reset_index()
    hours = g["toi5"] / 3600
    out = g[["season_id", "team_id", "games"]].copy()
    out["exit_time"] = g["exit_sum"] / g["exit_starts"]
    out["exit_success"] = g["exit_ok"] / g["exit_starts"]
    out["quick_exit"] = g["exit_fast"] / g["exit_starts"]
    out["costly_pg"] = g["costly"] / hours
    out["costly_share"] = g["costly"] / g["giveaways"]
    out["trans_pg"] = g["trans_for"] / hours
    out["trans_against_pg"] = g["trans_against"] / hours
    out["trans_share"] = g["trans_for"] / (g["trans_for"] + g["trans_against"])
    out["exit_starts"] = g["exit_starts"]
    return out.replace([np.inf, -np.inf], np.nan)


def add_index(t):
    """z-scores within each season (flipped so higher = better), and the equal-weight Breakout Index."""
    t = t.copy()
    for col, _, higher in COMPONENTS:
        z = ps.zscore_by_season(t, col)
        t[f"z_{col}"] = z if higher else -z
    raw = t[[f"z_{c}" for c, _, _ in COMPONENTS]].mean(axis=1)
    t["breakout"] = (raw - raw.groupby(t["season_id"]).transform("mean")) / raw.groupby(t["season_id"]).transform("std")
    t["z_breakout"] = t["breakout"].fillna(0.0)
    return t


# ----------------------------------------------------------------------------
# The tests
# ----------------------------------------------------------------------------
def benchmark_halves(nhl, seasons):
    """Familiar stats split the same way (odd vs even games), so the Breakout stats can be compared like for like."""
    ids = ",".join(str(int(s)) for s in seasons)
    mp = pd.read_sql_query(
        f"""SELECT season_id, game_id, team_id, shotAttemptsFor AS saf, shotAttemptsAgainst AS saa, xGoalsFor AS xgf, xGoalsAgainst AS xga
            FROM mp_team_games WHERE situation = '5on5' AND playoffGame = 0 AND season_id IN ({ids})""", nhl)
    pts = pd.read_sql_query(f"SELECT season_id, game_id, team_id, points FROM team_games WHERE game_type = 2 AND season_id IN ({ids})", nhl)
    out = {}
    for parity in (1, 0):
        m = mp[mp["game_id"] % 2 == parity].groupby(["season_id", "team_id"])[["saf", "saa", "xgf", "xga"]].sum()
        p = pts[pts["game_id"] % 2 == parity].groupby(["season_id", "team_id"])["points"].agg(["sum", "count"])
        d = pd.DataFrame({"corsi": m["saf"] / (m["saf"] + m["saa"]), "xg_share": m["xgf"] / (m["xgf"] + m["xga"]),
                          "pts_pct": p["sum"] / (2 * p["count"])}).reset_index()
        out[parity] = d
    return out[1].merge(out[0], on=["season_id", "team_id"], suffixes=("_a", "_b"))


def split_half(game_rows, nhl, seasons):
    """Odd vs even game ids, per team-season, for the three stats and the index (and familiar stats for comparison)."""
    a = add_index(summarize(game_rows[game_rows["game_id"] % 2 == 1]))
    b = add_index(summarize(game_rows[game_rows["game_id"] % 2 == 0]))
    m = a.merge(b, on=["season_id", "team_id"], suffixes=("_a", "_b"))
    rows = []

    def add_row(name, r):
        rows.append({"Statistic": name, "Repeats: odd vs even games": round(r.mean(), 2),
                     "Estimate for a full season": round(2 * r.mean() / (1 + r.mean()), 2),
                     "Lowest season": round(r.min(), 2), "Highest season": round(r.max(), 2)})

    for col, name in [(c, n) for c, n, _ in COMPONENTS] + [("breakout", "Breakout Index")]:
        add_row(name, ps.season_spearman(m, f"{col}_a", f"{col}_b"))
    bench = benchmark_halves(nhl, seasons)
    for col, name in [("corsi", "For comparison: Corsi (5v5 shot-attempt share)"), ("xg_share", "For comparison: expected-goals share (5v5)"),
                      ("pts_pct", "For comparison: points %")]:
        add_row(name, ps.season_spearman(bench, f"{col}_a", f"{col}_b"))
    return pd.DataFrame(rows)


def comparison_numbers(nhl, seasons):
    """Other team numbers to compare with: points %, goal difference, Corsi, expected-goals share, NHL zone time."""
    ts = pd.read_sql_query("SELECT season_id, team_id, season, team, pts_pct, goal_diff / 82.0 AS gd_pg, outcome, champion, cup_finalist, conf_finalist, include_main FROM team_seasons", nhl)
    sat = pd.read_sql_query("SELECT seasonId AS season_id, teamId AS team_id, satPct AS corsi FROM ts_percentages_regular WHERE teamId IS NOT NULL", nhl)
    sat["team_id"] = sat["team_id"].astype(int)
    xg = pd.read_sql_query(
        "SELECT season_id, team_id, xGoalsFor / (xGoalsFor + xGoalsAgainst) AS xg_share FROM mp_team_seasons WHERE situation = '5on5' AND team_id IS NOT NULL", nhl)
    xg["team_id"] = xg["team_id"].astype(int)
    try:
        edge = pd.read_sql_query("SELECT season_id, teamId AS team_id, es_dz_pct, es_oz_pct FROM edge_zone_time_regular", nhl)
        edge["team_id"] = edge["team_id"].astype(int)
    except Exception:                                                    # noqa: BLE001
        edge = pd.DataFrame(columns=["season_id", "team_id", "es_dz_pct", "es_oz_pct"])
    out = ts.merge(sat, on=["season_id", "team_id"], how="left").merge(xg, on=["season_id", "team_id"], how="left").merge(edge, on=["season_id", "team_id"], how="left")
    return out[out["season_id"].isin(seasons)]


def links_table(t):
    targets = [("pts_pct", "Points %"), ("gd_pg", "Goal difference per game"), ("corsi", "Corsi (shot-attempt share, 5v5)"),
               ("xg_share", "Expected-goals share (5v5)"), ("es_dz_pct", "NHL zone time: share in own zone (5v5, 2021-22 on)"),
               ("es_oz_pct", "NHL zone time: share in offensive zone (5v5, 2021-22 on)")]
    rows = []
    for col, name, _ in [(c, n, h) for c, n, h in COMPONENTS] + [("breakout", "Breakout Index (higher is better)", True)]:
        row = {"Breakout statistic": name}
        for target, label in targets:
            d = t.dropna(subset=[col, target])
            row[label] = round(ps.season_spearman(d, col, target).mean(), 2) if len(d) > 30 else None
        rows.append(row)
    return pd.DataFrame(rows)


def internal_links(t):
    cols = [c for c, _, _ in COMPONENTS]
    names = {c: n.split(" (")[0] for c, n, _ in COMPONENTS}
    rows = []
    for a in cols:
        row = {"": names[a]}
        for b in cols:
            row[names[b]] = round(ps.season_spearman(t, a, b).mean(), 2) if a != b else None
        rows.append(row)
    return pd.DataFrame(rows)


def level_by_season(t):
    g = t.groupby("season_id").agg(exit_time=("exit_time", "mean"), quick_exit=("quick_exit", "mean"), exit_success=("exit_success", "mean"),
                                   costly_pg=("costly_pg", "mean"), costly_share=("costly_share", "mean"), trans_pg=("trans_pg", "mean")).reset_index()
    g.insert(0, "Season", g["season_id"].map(season_label))
    g = g.drop(columns="season_id").round(3)
    g.columns = ["Season", "Avg zone exit time (s)", "Share of exits within 8 s", "Share of attempts that cleared the zone", "Costly giveaways /60",
                 "Share of own-zone giveaways that were costly", "Transition shot attempts /60"]
    return g


def season_label(season_id):
    text = str(int(season_id))
    return f"{text[:4]}-{text[6:]}"


def champion_profile(t):
    d = t[t["outcome"].isin(OUTCOME_ORDER)].copy()
    for col, name, higher in COMPONENTS + [("breakout", "Breakout Index", True)]:
        p = d.groupby("season_id")[col].rank(pct=True)
        d[f"pct_{col}"] = (p if higher else 1 - p + 1 / d.groupby("season_id")[col].transform("count")) * 100
    rows = []
    for outcome in OUTCOME_ORDER + ["ALL teams"]:
        part = d if outcome == "ALL teams" else d[d["outcome"] == outcome]
        rows.append({"How the season ended": outcome, "Team-seasons": len(part),
                     "Zone exit time (percentile)": round(part["pct_exit_time"].mean(), 0), "Costly giveaways (percentile)": round(part["pct_costly_pg"].mean(), 0),
                     "Transition shots (percentile)": round(part["pct_trans_pg"].mean(), 0), "Breakout Index (percentile)": round(part["pct_breakout"].mean(), 0)})
    return pd.DataFrame(rows)


def checkpoint_features(game_rows, nhl, checkpoints):
    """The index using only games up to each checkpoint date (no peeking), for the contender test."""
    dates = pd.read_sql_query("SELECT game_id, game_date FROM games", nhl)
    rows = game_rows.merge(dates, on="game_id")
    out = {}
    for cp in CP_ORDER:
        d = checkpoints[checkpoints["checkpoint"] == cp][["season_id", "date"]]
        r = rows.merge(d, on="season_id")
        r = r[r["game_date"] <= r["date"]]
        out[cp] = add_index(summarize(r))[["season_id", "team_id", "z_breakout", "breakout"]]
    return out


def forecast_test(connection, game_rows, covered, force):
    """Does the index add to the contender score? Returns (table, note)."""
    if len(covered) < MIN_SEASONS_FOR_FORECAST and not force:
        return pd.DataFrame(), (f"Not run yet: only {len(covered)} seasons of play-by-play are downloaded and the hold-out test needs "
                                f"{MIN_SEASONS_FOR_FORECAST} or more. Rerun this script when the download has finished.")
    games_raw, seasons, checkpoints = cf.load(connection)
    games, _ = ps.prepare_games(games_raw, seasons)
    tables = cf.build_checkpoint_tables(games, seasons, checkpoints)
    feats = checkpoint_features(game_rows, connection, checkpoints)
    v1 = cf.MODELS["Model 3: + power kill (PP% + PK%)"]
    rows = []
    for cp in CP_ORDER:
        t = tables[cp].merge(feats[cp], on=["season_id", "team_id"], how="inner")
        t = t[t["season_id"].isin(covered)].copy()
        if t["season_id"].nunique() < 3:
            continue
        base = cf.held_out_scores(t, v1)
        plus = cf.held_out_scores(t.assign(z_breakout=t["z_breakout"]), v1 + ["breakout"])
        t["v1"], t["v1+breakout"] = base, plus
        a, b = cf.per_season_metrics(t, "v1"), cf.per_season_metrics(t, "v1+breakout")
        diff = b["auc"] - a["auc"]
        rows.append({"Checkpoint": PLAIN_CP[cp], "Seasons": len(diff), "AUC v1": a["auc"].mean(), "AUC v1 + Breakout Index": b["auc"].mean(),
                     "Gain": diff.mean(), "Margin (+/-, 2 std. errors)": 2 * diff.std(ddof=1) / np.sqrt(len(diff)),
                     "Avg finalist rank v1": a["avg_rank_f4"].mean(), "Avg finalist rank with index": b["avg_rank_f4"].mean(),
                     "Seasons better": int((b["avg_rank_f4"] < a["avg_rank_f4"]).sum()), "Seasons worse": int((b["avg_rank_f4"] > a["avg_rank_f4"]).sum())})
    table = pd.DataFrame(rows).round(3)
    note = "Hold-out test run on " + f"{len(covered)} seasons" + (" (FORCED with fewer than 8 seasons: treat as a code check, not a result)." if len(covered) < MIN_SEASONS_FOR_FORECAST else ".")
    return table, note


# ----------------------------------------------------------------------------
def verdict_lines(reliab, links, level, forecast, forecast_note, champs, n_seasons, season_list):
    rel = reliab.set_index("Statistic")["Repeats: odd vs even games"]
    link = links.set_index("Breakout statistic")
    idx = link.loc["Breakout Index (higher is better)"]
    corsi_link = idx["Corsi (shot-attempt share, 5v5)"]
    pts_link = idx["Points %"]
    edge_exit = link.iloc[0]["NHL zone time: share in own zone (5v5, 2021-22 on)"]
    edge_idx = idx["NHL zone time: share in offensive zone (5v5, 2021-22 on)"]
    d = champs.set_index("How the season ended")
    lines = [
        ("What this is", f"A prototype of your Breakout Index on {n_seasons} fully downloaded seasons ({season_list}). Rerun when the download finishes to use every season."),
        ("Does it repeat?", f"Odd-game vs even-game rank correlation (half a season each): zone exit time {rel.iloc[0]:.2f}, costly giveaways {rel.iloc[1]:.2f}, transition shots {rel.iloc[2]:.2f}, "
                            f"Breakout Index {rel.iloc[3]:.2f}. For comparison, split the same way: Corsi {rel.iloc[4]:.2f}, expected-goals share {rel.iloc[5]:.2f}, points % {rel.iloc[6]:.2f}. "
                            "Closer to 1 means more of a stable team trait and less luck."),
        ("Does it line up with winning?", f"The index's link with points % is {pts_link:+.2f} and with shot-attempt share (Corsi) {corsi_link:+.2f}. "
                                          + ("The link with Corsi is modest, so it is partly a different thing." if abs(corsi_link) < 0.35 else
                                             "It overlaps with Corsi to a real degree (about 0.3 to 0.5 is moderate), so it may add little beyond it. The contender test below is the real check.")),
        ("Does it match the NHL's real zone time?", (f"Zone exit time vs share of 5v5 time in the team's own zone: {edge_exit:+.2f} (positive is what we expect: slow exits, more time in your own end). "
                                                      f"Breakout Index vs offensive-zone share: {edge_idx:+.2f}.") if pd.notna(edge_exit) else "Not testable yet: no 2021-22-or-later seasons have finished downloading."),
        ("Does it improve the contender score?", forecast_note + (" " + _forecast_sentence(forecast) if len(forecast) else "")),
        ("Champions", f"Average Breakout Index percentile: Won Cup {d.loc['Won Cup', 'Breakout Index (percentile)']:.0f}, Lost Cup final {d.loc['Lost Cup final', 'Breakout Index (percentile)']:.0f}, "
                      f"Missed playoffs {d.loc['Missed playoffs', 'Breakout Index (percentile)']:.0f} (100 = best, 50 = average)." if "Won Cup" in d.index else "No champions in the seasons downloaded so far."),
    ]
    return pd.DataFrame(lines, columns=["", "Answer"])


def _forecast_sentence(forecast):
    gain, margin = forecast["Gain"].mean(), forecast["Margin (+/-, 2 std. errors)"].mean()
    return (f"Adding the index to version 1 changes the average AUC by {gain:+.3f} (margin about +/-{margin:.3f}): "
            + ("larger than the noise." if gain > margin else "within the noise."))


def write_excel(sheets, notes):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 150
        ws["A1"].font = Font(bold=True)
        for name, frame in sheets.items():
            start = 1 if name in notes else 0
            frame.to_excel(writer, sheet_name=name, index=False, startrow=start)
            ws = writer.sheets[name]
            header_row = start + 1
            if name in notes:
                ws.cell(row=1, column=1, value=notes[name]).font = Font(bold=True)
            ws.freeze_panes = ws.cell(row=header_row + 1, column=2)
            for cell in ws[header_row]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column in ws.columns:
                values = [c.value for c in column[header_row - 1:300] if c.value is not None and (not isinstance(c.value, str) or len(c.value) < 70)]
                longest = max((len(str(v)) for v in values), default=8)
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 46)
        ws = writer.sheets["Verdict"]
        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["B"].width = 135
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(bold=True)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")


def main():
    global EXIT_CAP                                                      # so --exit-cap can change the setting for the whole run
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Build the Breakout Index from play-by-play.")
    parser.add_argument("--seasons", help="comma-separated season ids to use, e.g. 20232024,20242025")
    parser.add_argument("--force-forecast", action="store_true", help="run the hold-out test with fewer than 8 seasons (unreliable)")
    parser.add_argument("--exit-cap", type=int, default=EXIT_CAP, help=f"seconds after which a zone exit counts as failed (default {EXIT_CAP})")
    args = parser.parse_args()
    EXIT_CAP = args.exit_cap
    if not PBP_DB.exists() or not NHL_DB.exists():
        print("I need data/pbp.db (python scripts/pull_play_by_play.py) and data/nhl.db (python scripts/build_database.py).")
        return 1
    pbp = sqlite3.connect(f"file:{PBP_DB}?mode=ro", uri=True)           # read-only: safe while the download is running
    nhl = sqlite3.connect(NHL_DB)
    if "mp_team_games" not in [r[0] for r in nhl.execute("SELECT name FROM sqlite_master WHERE type='table'")]:
        print("The MoneyPuck tables are missing (they give 5-on-5 time on ice). Run: python scripts/load_moneypuck.py")
        return 1

    coverage = season_coverage(pbp, nhl)
    usable = [int(s) for s in coverage.loc[coverage["coverage"] >= MIN_COVERAGE, "season_id"]]
    if args.seasons:
        usable = [s for s in usable if s in {int(x) for x in args.seasons.split(",")}]
    flags = pd.read_csv(PROJECT_FOLDER / "reference" / "season_flags.csv")
    odd_seasons = set(int(s) for s in flags.loc[flags["include_main"] == "N", "season_id"])
    current = int(coverage["season_id"].max())
    usable = [s for s in usable if s != current]
    print("Season coverage of the play-by-play download:")
    for r in coverage.itertuples():
        mark = "USED" if r.season_id in usable else ("in progress" if r.season_id == current else "")
        print(f"  {season_label(r.season_id)}: {int(r.saved):,} of {int(r.played):,} games ({r.coverage:.0%})  {mark}")
    if not usable:
        print("No season is fully downloaded yet. Try again later.")
        return 1

    pieces, blocked_checks = [], []
    for season in usable:
        ev = load_events(pbp, season)
        blocked_checks.append(check_blocked_shot_view(ev))
        pieces.append(team_game_stats(ev).assign(season_id=season))
        print(f"  {season_label(season)}: {ev['game_id'].nunique():,} games, {len(ev):,} events processed")
    game_rows = add_ice_time(pd.concat(pieces, ignore_index=True), nhl)
    missing_toi = int(game_rows["toi5"].isna().sum())
    game_rows["toi5"] = game_rows["toi5"].fillna(game_rows["toi5"].median())
    print(f"\nCheck: blocked shots read from the shooter's side land in the offensive zone {np.nanmean(blocked_checks):.1%} of the time -> "
          f"{'PASS' if np.nanmean(blocked_checks) > 0.9 else 'CHECK'}; games without 5v5 time on ice from MoneyPuck: {missing_toi}")

    team = add_index(summarize(game_rows))
    compare = comparison_numbers(nhl, usable)
    team = team.merge(compare, on=["season_id", "team_id"], how="left")
    main_team = team[team["include_main"].fillna("Y") == "Y"]
    reliab = split_half(game_rows, nhl, usable)
    links = links_table(team)
    internal = internal_links(team)
    level = level_by_season(team)
    champs = champion_profile(team[team["include_main"].fillna("Y") == "Y"])
    forecast, note = forecast_test(nhl, game_rows[game_rows["season_id"].isin([s for s in usable if s not in odd_seasons])],
                                   [s for s in usable if s not in odd_seasons], args.force_forecast)
    verdict = verdict_lines(reliab, links, level, forecast, note, champs, len(usable), ", ".join(season_label(s) for s in usable))

    out = team.sort_values(["season_id", "breakout"], ascending=[False, False])
    teams_out = pd.DataFrame({
        "Season": out["season_id"].map(season_label), "Team": out["team"], "Breakout Index (z, higher is better)": out["breakout"].round(2),
        "Avg zone exit time (s)": out["exit_time"].round(2), "Share of exits within 8 s": out["quick_exit"].round(3),
        "Share of attempts that cleared the zone": out["exit_success"].round(3), "Costly giveaways /60": out["costly_pg"].round(2),
        "Share of own-zone giveaways that were costly": out["costly_share"].round(3), "Transition shot attempts /60": out["trans_pg"].round(2),
        "Transition shot attempts against /60": out["trans_against_pg"].round(2), "Transition shot share": out["trans_share"].round(3),
        "Own-zone starts counted": out["exit_starts"].astype(int), "Points %": out["pts_pct"].round(3), "Corsi % (5v5)": (out["corsi"] * 100).round(1),
        "How the season ended": out["outcome"]}).reset_index(drop=True)

    sheets = {"Verdict": verdict, "Repeats": reliab, "Links": links, "Between the three": internal, "Level by season": level,
              "Champions": champs, "Contender test": forecast if len(forecast) else pd.DataFrame({"": [note]}), "Team-seasons": teams_out}
    notes = {"Links": "Rank correlation (within each season, averaged) between each breakout statistic and other team numbers. Positive is expected for the index; for exit time and costly giveaways negative is 'good'."}
    try:
        write_excel(sheets, notes)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 220)
    pd.set_option("display.max_colwidth", 150)
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nLevel by season:"); print(level.to_string(index=False))
    print("\nRepeats (odd vs even games):"); print(reliab.to_string(index=False))
    print("\nLinks:"); print(links.to_string(index=False))
    print("\nBetween the three stats:"); print(internal.to_string(index=False))
    print("\nChampions:"); print(champs.to_string(index=False))
    if len(forecast):
        print("\nContender test:"); print(forecast.to_string(index=False))
    print("\nTop 5 and bottom 3 of the latest season:")
    latest = out[out["season_id"] == out["season_id"].max()]
    print(latest[["team", "breakout", "exit_time", "costly_pg", "trans_pg", "pts_pct"]].head(5).round(2).to_string(index=False))
    print(latest[["team", "breakout", "exit_time", "costly_pg", "trans_pg", "pts_pct"]].tail(3).round(2).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

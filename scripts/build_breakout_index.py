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

    FIVE MORE CANDIDATE STATS (not in the index: each has to earn its place):

      4. ICINGS             Icings per 60 minutes. The team that iced the puck is the one with the next faceoff in
                            its own end. LOWER IS BETTER.
      5. PINNED-IN SPELLS   Times the opponent gets two or more shot attempts in a row from inside your zone, each
                            within 30 seconds of the last, with no stoppage and no event outside your zone in between.
                            Per 60 minutes. LOWER IS BETTER.
      6. OWN-ZONE TURNOVERS Your own-zone giveaways plus the pucks the opponent took off you in your end (their
                            takeaways), per 60 minutes. LOWER IS BETTER.
      7. DEFENCE TO OFFENCE Share of own-zone puck wins (defensive-zone faceoff win, takeaway, blocked shot, opponent's
                            missed shot) that become your own shot attempt within 20 seconds, before play stops.
                            HIGHER IS BETTER.
      8. TRANSITION SHOT    Of your unblocked transition shots, the share taken from the slot (between the faceoff
         QUALITY            dots, from the goal line out to the top of the circles). HIGHER IS BETTER.

    THE PASS RULE (written down before any result was looked at). A stat earns a place only if all three hold:
      1. it repeats: the full-season estimate from odd vs even games is 0.50 or more;
      2. it points the right way: its link with goal difference in the same season is positive;
      3. it adds something beyond Corsi out of sample: built from one half of a season's games, its link with the
         results of the OTHER half, after taking out Corsi, is above zero by more than its noise margin.
    With eight stats tested, one narrow pass can still be luck: read a pass as "worth keeping an eye on", not proof.

    THE TESTS:
      * REPEATS?    does a team's number from odd-numbered games match its number from even-numbered games?
      * LINKS?      how does it line up with points %, goal difference, shot-attempt share (Corsi), expected goals,
                    and (2021-22 on) the NHL's real zone-time numbers (a team that exits quickly should spend
                    less time in its own end)?
      * ADDS?       does adding the index to the contender score (version 1) pick the real conference finalists
                    better, on seasons the score has never seen? (needs 8 or more seasons of play-by-play)
      * CHAMPIONS   where did champions and finalists rank?
      * WEIGHTS?    which mix of the three stats lines up best with good teams (points % and goal difference)?
                    Every mix in steps of 10% is tried. To stay honest the best mix is chosen on all seasons but one
                    and scored on the season left out, then compared with plain equal weights.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_breakout_index.py
    Safe to run while the download is still going; it uses the seasons that are fully downloaded so far.
    Five early seasons (2009-10 to 2013-14) are deliberately left out: they are kept for a one-time blind test of the
    final metric (see LOCKBOX below and build_breakout_round2.py).
    Options:  --seasons 20232024,20242025     only those seasons
              --weights 0.5,0.2,0.3           weights for zone exit time, costly giveaways, transition shots (default: equal)
              --exit-cap 30                   seconds after which a zone exit counts as failed (default 20)
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
Z_COLS = [f"z_{c}" for c, _, _ in COMPONENTS]
SHORT_NAMES = ["Zone exit time", "Costly giveaways", "Transition shots"]
# How much each stat counts in the index, in the order above (zone exit time, costly giveaways, transition shots).
# They are rescaled to add up to 1. Change them here, or for one run with  --weights 0.5,0.2,0.3
WEIGHTS = (1 / 3, 1 / 3, 1 / 3)
WEIGHT_STEP = 0.1        # the weight test tries every mix of the three weights in steps of this size

# ---- Five more candidate stats (not in the index yet: each has to earn its place, see PASS RULE below) ----
PINNED_GAP = 30          # seconds: opponent shot attempts this close together, with play never leaving the zone, are one "pinned-in spell"
D2O_WINDOW = 20          # seconds from an own-zone puck win to the team's own shot attempt for it to count as "defence to offence"
SLOT_HALF_WIDTH = 22     # feet either side of the middle of the ice (the faceoff dots are 22 feet out)
SLOT_DEPTH = 35          # feet out from the goal line (the top of the faceoff circles)
GOAL_LINE_X = 89         # the goal line is 89 feet from centre ice in the NHL's coordinates
PINNED_GAPS = [(PINNED_GAP, ""), (15, "15"), (45, "45")]     # the normal gap, plus a shorter and a longer one for round 2 of the search
D2O_WINDOWS = [(D2O_WINDOW, ""), (10, "10"), (30, "30")]     # the normal window, plus a shorter and a longer one

# ---- The blind test. These five seasons have never been used in any test. Every script leaves them alone, so that the one
#      metric the search ends up with can be tried on them ONCE (build_breakout_round2.py --blind-test). Do not remove a season
#      from this list to "have a look": a peek spends the only clean test we have. ----
LOCKBOX = [20092010, 20102011, 20112012, 20122013, 20132014]
FEATURES_VERSION = 3     # raise this when team_game_stats changes, so saved copies of its output are rebuilt

# ---- BER, "breakout efficiency rating":  100 x (controlled exits + 0.5 x safe clears) / breakout attempts ----
# The play-by-play does not say who has the puck, so each breakout attempt (an own-zone puck win at 5-on-5) is judged by the
# FIRST telling event after it, if that comes within BER_WINDOW seconds:
#   CONTROLLED EXIT  an event beyond our blue line that shows WE had the puck there: our shot attempt, our giveaway,
#                    a hit on our player, or the opponent taking the puck off us.
#   SAFE CLEAR       an event beyond our blue line that shows THEY had the puck there: their giveaway or shot, our hit on them,
#                    our takeaway, or the opponent icing it. It only counts as safe if nothing below happens within
#                    SAFE_CLEAR_WINDOW seconds (otherwise the puck came straight back: failed).
#   FAILED           something in our zone that shows we did not get out: a shot attempt against, our giveaway, the puck taken
#                    off us, our own takeaway or hit there (they had it again), or our icing.
#   CANNOT BE JUDGED a stoppage, a change of strength or nothing telling within the window. These are left OUT of the
#                    denominator: with no event to go on, calling them failures would just measure how few events were recorded.
BER_WINDOW = 20
SAFE_CLEAR_WINDOW = 10
CANDIDATES = [  # (column, plain name, higher is better?)
    ("icing_pg", "Icings per 60 (lower is better)", False),
    ("pinned_pg", "Pinned-in spells per 60 (lower is better)", False),
    ("turnover_pg", "Own-zone turnovers per 60: giveaways + pucks taken off you (lower is better)", False),
    ("d2o_rate", "Defence-to-offence rate: own-zone puck wins that become your shot attempt (higher is better)", True),
    ("trans_slot_share", "Transition shot quality: share from the slot (higher is better)", True),
]
ALL_STATS = COMPONENTS + CANDIDATES
ALL_Z = [f"z_{c}" for c, _, _ in ALL_STATS]
# PASS RULE, fixed before any result was looked at. A stat earns a place only if ALL THREE hold:
#   1. it repeats: the full-season estimate from odd vs even games is 0.50 or more;
#   2. it points the right way: its plain link with goal difference in the same season is positive;
#   3. it adds something beyond Corsi out of sample: built from one half of a season's games, its link with the RESULTS of the
#      other half (average of points % and goal difference), after taking out Corsi, is above zero by more than its noise margin.
PASS_REPEAT = 0.50
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
    "FIVE MORE CANDIDATE STATS (sheet 'New stats'; not in the index, each has to earn its place)",
    "  Icings               Icings per 60 minutes; the team that iced the puck is the one with the next faceoff in its own end. Lower is better.",
    f"  Pinned-in spells     Times the opponent gets 2+ shot attempts in a row from inside your zone, each within {PINNED_GAP} seconds of the last, with no stoppage and",
    "                       no event outside your zone in between. Per 60 minutes. Lower is better.",
    "  Own-zone turnovers   Your own-zone giveaways plus the pucks the opponent took off you in your end, per 60 minutes. Lower is better.",
    f"  Defence to offence   Share of own-zone puck wins (faceoff win, takeaway, blocked shot, opponent's missed shot) that become your own shot attempt within {D2O_WINDOW} seconds.",
    "  Transition shot quality   Of your unblocked transition shots, the share from the slot (between the faceoff dots, goal line to the top of the circles).",
    f"  PASS RULE (fixed before the results were looked at): a stat earns a place only if (1) its full-season repeat estimate is {PASS_REPEAT:.2f} or more, (2) its link with goal",
    "  difference is positive, and (3) built from one half of a season's games, its link with the results of the OTHER half, after taking out Corsi, is above zero by more",
    "  than its noise margin. With eight stats tested, one narrow pass can still be luck: read a pass as 'worth keeping an eye on', not proof.",
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
    "  Weights    every mix of the three weights (steps of 10%) is scored by its average rank correlation with points % and goal difference. Picking the best mix on the same",
    "             seasons it is scored on would flatter it, so the 'Weights by season' sheet picks the mix WITHOUT one season and scores it on that season. 'Beyond Corsi' is the",
    "             link left after taking out what shot-attempt share already explains; 'Other half' uses the stats from odd games against the results of even games (and the reverse).",
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
        """SELECT e.game_id, e.sort_order, e.period, e.time_s, e.situation_code, e.type, e.team_id, e.zone, e.x, e.y, e.reason,
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


def both(frame, sums):
    """
    Adds things up per team per game twice: over the whole game, and over CLOSE-SCORE moments only (columns ending _c).
    `sums` maps the name of each total to the column to add up; `frame` needs the columns game_id, T and close.
    """
    data = {name: frame[col].astype(float) for name, col in sums.items()}
    data.update({f"{name}_c": frame[col].astype(float) * frame["close"] for name, col in sums.items()})
    return pd.DataFrame(data).groupby([frame["game_id"], frame["T"]]).sum().reset_index()


def team_game_stats(ev):
    """From one season's events, one row per team per game with the counts every breakout stat is built from."""
    ev = ev.sort_values(["game_id", "sort_order"]).reset_index(drop=True)
    ev["abs_t"] = (ev["period"] - 1) * 1200 + ev["time_s"]
    ev["one"] = 1

    # --- the score before every event, and whether the game is CLOSE (within one goal in periods 1-2, tied after that).
    #     A team sitting on a lead lets itself be pinned in, so close-score moments are a cleaner read. ---
    goal = ev["type"].eq("goal")
    home_goal = (goal & ev["team_id"].eq(ev["home_team_id"])).astype(int)
    away_goal = (goal & ev["team_id"].eq(ev["away_team_id"])).astype(int)
    lead = (home_goal.groupby(ev["game_id"]).cumsum() - home_goal) - (away_goal.groupby(ev["game_id"]).cumsum() - away_goal)
    ev["close"] = np.where(ev["period"] <= 2, lead.abs() <= 1, lead == 0)
    # the clock: seconds of 5-on-5 play, and of close-score 5-on-5 play, counted from the time between events
    gap_to_next = (ev.groupby(["game_id", "period"])["abs_t"].shift(-1) - ev["abs_t"]).clip(lower=0).fillna(0)
    five_now = ev["situation_code"].eq(FIVE_ON_FIVE)
    clock = pd.DataFrame({"game_id": ev["game_id"], "ev5": gap_to_next.where(five_now, 0), "ev5_c": gap_to_next.where(five_now & ev["close"], 0)}
                         ).groupby("game_id").sum().reset_index()

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
    starts = p[p["mine"] & five & p["zone_T"].eq("D") & p["type"].isin(["faceoff", "takeaway"])][["game_id", "T", "sort_order", "abs_t", "close", "one"]]
    exits = p[p["zone_T"].isin(["N", "O"])]
    fails = p[p["type"].isin(STOP_TYPES) | (~p["mine"] & p["type"].isin(SHOT_TYPES))
              | (p["mine"] & p["type"].eq("giveaway") & p["zone_T"].eq("D")) | ~five]
    s = forward_match(forward_match(starts, exits, "exit"), fails, "fail")
    took = s["exit_t"] - s["abs_t"]
    ok = s["exit_so"].notna() & (s["fail_so"].isna() | (s["exit_so"] < s["fail_so"])) & (took <= EXIT_CAP)
    s["exit_seconds"] = np.where(ok, took, EXIT_CAP)
    s["exit_ok"], s["exit_fast"] = ok, ok & (took <= FAST_EXIT)
    exit_stats = both(s, {"exit_starts": "one", "exit_sum": "exit_seconds", "exit_ok": "exit_ok", "exit_fast": "exit_fast"})

    # --- 2) costly giveaways ---
    gives = p[p["mine"] & five & p["type"].eq("giveaway") & p["zone_T"].eq("D")][["game_id", "T", "sort_order", "abs_t", "close", "one"]]
    opp_shots = p[~p["mine"] & p["type"].isin(SHOT_TYPES)]
    stops = p[p["type"].isin(STOP_TYPES)]
    g = forward_match(forward_match(gives, opp_shots, "shot"), stops, "stop")
    costly = g["shot_so"].notna() & ((g["shot_t"] - g["abs_t"]) <= GIVEAWAY_WINDOW) & (g["stop_so"].isna() | (g["shot_so"] < g["stop_so"]))
    g["costly"] = costly
    give_stats = both(g, {"giveaways": "one", "costly": "costly"})

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
    # candidate 5) transition shot quality: of the unblocked transition shots, how many came from the slot?
    # (blocked shots are left out because their location is where the block happened, not where the shot was taken)
    in_slot = sh["zone"].eq("O") & (sh["y"].abs() <= SLOT_HALF_WIDTH) & sh["x"].abs().between(GOAL_LINE_X - SLOT_DEPTH, GOAL_LINE_X)
    unblocked = sh["transition"] & sh["type"].ne("blocked-shot")
    sh = sh.assign(trans_unblocked=unblocked, trans_slot=unblocked & in_slot)
    shots_for = both(sh.assign(T=sh["team_id"]), {"shots_for": "one", "trans_for": "transition", "trans_unblocked": "trans_unblocked", "trans_slot": "trans_slot"})
    shots_against = both(sh.assign(T=sh["opponent"]), {"shots_against": "one", "trans_against": "transition"})

    # --- candidate 1) icings: the team that iced the puck is the one that has the next faceoff in its own end ---
    icing = ev[ev["type"].eq("stoppage") & ev["reason"].eq("icing") & ev["situation_code"].eq(FIVE_ON_FIVE)][["game_id", "sort_order", "close", "one"]]
    faceoffs = ev[ev["type"].eq("faceoff")][["game_id", "sort_order", "team_id", "zone", "home_team_id", "away_team_id"]].rename(columns={"sort_order": "fo_so"})
    faceoffs["sort_order"] = faceoffs["fo_so"]
    ic = pd.merge_asof(icing.sort_values("sort_order"), faceoffs.sort_values("sort_order"), on="sort_order", by="game_id",
                       direction="forward", allow_exact_matches=False)
    loser = ic["home_team_id"] + ic["away_team_id"] - ic["team_id"]
    ic["T"] = np.where(ic["zone"].eq("D"), ic["team_id"], np.where(ic["zone"].eq("O"), loser, np.nan))    # faceoff zone is from the WINNER's side
    ic = ic.dropna(subset=["T"])
    ic["T"] = ic["T"].astype(int)
    icing_stats = both(ic, {"icings": "one"})

    # --- candidate 2) pinned-in spells: 2+ opponent shot attempts in a row, each within PINNED_GAP seconds, play never leaving the zone.
    #     Also counted with a shorter and a longer gap, and the DANGEROUS ones (a spell with an unblocked shot from the slot). ---
    p = p.sort_values(["game_id", "T", "sort_order"]).reset_index(drop=True)
    p = p.merge(ic[["game_id", "sort_order", "T"]].rename(columns={"T": "icing_team"}), on=["game_id", "sort_order"], how="left")   # who iced it (5-on-5 icings)
    five = p["situation_code"] == FIVE_ON_FIVE
    opp_shot_in_zone = ~p["mine"] & p["type"].isin(SHOT_TYPES) & p["zone_T"].eq("D") & five
    breaks = p["type"].isin(STOP_TYPES + ["faceoff", "period-start"]) | p["zone_T"].isin(["N", "O"]) | ~five
    p["segment"] = breaks.groupby([p["game_id"], p["T"]]).cumsum()
    p["slot"] = p["type"].ne("blocked-shot") & (p["y"].abs() <= SLOT_HALF_WIDTH) & p["x"].abs().between(GOAL_LINE_X - SLOT_DEPTH, GOAL_LINE_X)
    os_ = p[opp_shot_in_zone][["game_id", "T", "segment", "abs_t", "close", "slot"]].copy()
    gap = os_.groupby(["game_id", "T", "segment"])["abs_t"].diff()
    pinned_pieces = []
    for seconds, tag in PINNED_GAPS:
        os_["spell"] = (gap.isna() | (gap > seconds)).cumsum()
        spells = os_.groupby("spell").agg(game_id=("game_id", "first"), T=("T", "first"), close=("close", "first"),
                                          shots=("abs_t", "size"), danger=("slot", "max"))
        spells = spells[spells["shots"] >= 2].assign(one=1)
        sums = {f"pinned{tag}": "one"}
        if tag == "":
            sums.update({"pinned_shots": "shots", "pinned_danger": "danger"})
        pinned_pieces.append(both(spells, sums))

    # --- candidate 3) own-zone turnovers: the team's own-zone giveaways plus the pucks the opponent took off it there ---
    stripped = both(p[~p["mine"] & five & p["type"].eq("takeaway") & p["zone_T"].eq("D")], {"stripped": "one"})

    # --- candidate 4) defence to offence: own-zone puck wins (faceoff win, takeaway, blocked shot, opponent's missed shot)
    #     that become the team's own shot attempt within D2O_WINDOW seconds, before play stops (also with a shorter and a longer window) ---
    wins = p[five & p["zone_T"].eq("D") & ((p["mine"] & p["type"].isin(["faceoff", "takeaway"]))
                                           | (~p["mine"] & p["type"].isin(["blocked-shot", "missed-shot"])))][["game_id", "T", "sort_order", "abs_t", "close", "one"]]
    own_shots = p[p["mine"] & p["type"].isin(SHOT_TYPES)]
    d = forward_match(forward_match(wins, own_shots, "shot"), p[p["type"].isin(STOP_TYPES)], "stop")
    reached = d["shot_so"].notna() & (d["stop_so"].isna() | (d["shot_so"] < d["stop_so"]))
    sums = {"d2o_starts": "one"}
    for seconds, tag in D2O_WINDOWS:
        d[f"d2o_ok{tag}"] = reached & ((d["shot_t"] - d["abs_t"]) <= seconds)
        sums[f"d2o_ok{tag}"] = f"d2o_ok{tag}"
    d2o_stats = both(d, sums)

    # --- BER: each own-zone puck win judged by the first telling event after it (see the BER notes at the top of this file) ---
    out_of_zone, in_zone = p["zone_T"].isin(["N", "O"]), p["zone_T"].eq("D")
    mine, theirs, kind = p["mine"], ~p["mine"], p["type"]
    we_had_it_out = out_of_zone & ((mine & kind.isin(SHOT_TYPES)) | (mine & kind.eq("giveaway")) | (theirs & kind.eq("hit")) | (theirs & kind.eq("takeaway")))
    they_had_it_out = (out_of_zone & ((theirs & kind.eq("giveaway")) | (theirs & kind.isin(SHOT_TYPES)) | (mine & kind.eq("hit")) | (mine & kind.eq("takeaway")))
                       ) | (p["icing_team"].notna() & p["icing_team"].ne(p["T"]))
    still_in = (in_zone & ((theirs & kind.isin(SHOT_TYPES)) | (mine & kind.eq("giveaway")) | (theirs & kind.eq("takeaway"))
                           | (mine & kind.eq("takeaway")) | (mine & kind.eq("hit")))) | p["icing_team"].eq(p["T"])
    play_ended = kind.isin(STOP_TYPES + ["period-start"]) | ~five
    b = wins
    for name, mask in (("ctrl", we_had_it_out), ("clear", they_had_it_out), ("fail", still_in), ("end", play_ended)):
        b = forward_match(b, p[mask], name)
    order = b[["ctrl_so", "clear_so", "fail_so", "end_so"]].fillna(np.inf).to_numpy()
    first = order.argmin(axis=1)                                         # a tie (an icing is also a stoppage) goes to the earlier column
    when = b[["ctrl_t", "clear_t", "fail_t", "end_t"]].to_numpy()[np.arange(len(b)), first]
    in_time = np.isfinite(order.min(axis=1)) & ((when - b["abs_t"].to_numpy()) <= BER_WINDOW)
    came_back = (b["fail_so"].notna() & ((b["fail_t"] - b["clear_t"]) <= SAFE_CLEAR_WINDOW) & (b["end_so"].isna() | (b["fail_so"] < b["end_so"]))).to_numpy()
    b["ber_ctrl"] = in_time & (first == 0)
    b["ber_safe"] = in_time & (first == 1) & ~came_back
    b["ber_fail"] = in_time & ((first == 2) | ((first == 1) & came_back))
    ber_stats = both(b, {"ber_attempts": "one", "ber_ctrl": "ber_ctrl", "ber_safe": "ber_safe", "ber_fail": "ber_fail"})

    out = exit_stats
    for piece in [give_stats, shots_for, shots_against, icing_stats, stripped, d2o_stats, ber_stats] + pinned_pieces:
        out = out.merge(piece, on=["game_id", "T"], how="outer")
    out = out.fillna(0).merge(clock, on="game_id", how="left")
    out["T"] = out["T"].astype(int)
    return out.rename(columns={"T": "team_id"})


def load_game_rows(pbp, nhl, only=None, lockbox=False):
    """
    Reads every fully downloaded season and returns (one row per team per game, the seasons used, the flagged odd seasons),
    or None if no season is ready. `only` is an optional comma-separated list of season ids.
    The five LOCKBOX seasons are left alone unless lockbox=True, which reads ONLY those (the blind test).
    The per-game counts of each season are saved in data/breakout/ so later runs do not have to rebuild them.
    """
    coverage = season_coverage(pbp, nhl)
    flags = pd.read_csv(PROJECT_FOLDER / "reference" / "season_flags.csv")
    odd_seasons = set(int(s) for s in flags.loc[flags["include_main"] == "N", "season_id"])
    current = int(coverage["season_id"].max())
    ready = [int(s) for s in coverage.loc[coverage["coverage"] >= MIN_COVERAGE, "season_id"] if int(s) != current]
    usable = [s for s in ready if (s in LOCKBOX) == lockbox]
    if only:
        usable = [s for s in usable if s in {int(x) for x in only.split(",")}]
    print("Season coverage of the play-by-play download:")
    for r in coverage.itertuples():
        if r.season_id in usable:
            mark = "USED"
        elif r.season_id == current:
            mark = "in progress"
        elif r.season_id in LOCKBOX and not lockbox:
            mark = "kept back for the blind test"
        else:
            mark = ""
        print(f"  {season_label(r.season_id)}: {int(r.saved):,} of {int(r.played):,} games ({r.coverage:.0%})  {mark}")
    if not usable:
        print("No season is fully downloaded yet. Try again later.")
        return None

    folder = PROJECT_FOLDER / "data" / "breakout"
    folder.mkdir(parents=True, exist_ok=True)
    saved_games = {int(r.season_id): int(r.saved) for r in coverage.itertuples()}
    pieces, blocked_checks = [], []
    for season in usable:
        stamp = f"v{FEATURES_VERSION}_{EXIT_CAP}_{FAST_EXIT}_{GIVEAWAY_WINDOW}_{TRANSITION_WINDOW}_{PINNED_GAP}_{D2O_WINDOW}_{saved_games[season]}"
        saved = folder / f"game_rows_{season}_{stamp}.pkl"
        if saved.exists():
            rows, check = pd.read_pickle(saved)
            print(f"  {season_label(season)}: {rows['game_id'].nunique():,} games (saved copy)")
        else:
            ev = load_events(pbp, season)
            rows, check = team_game_stats(ev).assign(season_id=season), check_blocked_shot_view(ev)
            pd.to_pickle((rows, check), saved)
            print(f"  {season_label(season)}: {ev['game_id'].nunique():,} games, {len(ev):,} events processed")
        pieces.append(rows)
        blocked_checks.append(check)
    game_rows = add_ice_time(pd.concat(pieces, ignore_index=True), nhl)
    missing_toi = int(game_rows["toi5"].isna().sum())
    game_rows["toi5"] = game_rows["toi5"].fillna(game_rows["toi5"].median())
    # 5-on-5 ice time in close-score moments: MoneyPuck's 5-on-5 time, split by the share of the events' clock that was close
    game_rows["toi5_c"] = game_rows["toi5"] * (game_rows["ev5_c"] / game_rows["ev5"]).fillna(0)
    clock_match = float(np.corrcoef(game_rows["ev5"], game_rows["toi5"])[0, 1])
    print(f"\nCheck: blocked shots read from the shooter's side land in the offensive zone {np.nanmean(blocked_checks):.1%} of the time -> "
          f"{'PASS' if np.nanmean(blocked_checks) > 0.9 else 'CHECK'}; games without 5v5 time on ice from MoneyPuck: {missing_toi}")
    print(f"Check: 5-on-5 time counted from the events vs MoneyPuck's: correlation {clock_match:.2f}, ratio {game_rows['ev5'].sum() / game_rows['toi5'].sum():.2f}; "
          f"share of 5-on-5 time with the score close: {game_rows['toi5_c'].sum() / game_rows['toi5'].sum():.0%}")
    return game_rows, usable, odd_seasons


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
        trans_against=("trans_against", "sum"), shots_for=("shots_for", "sum"), shots_against=("shots_against", "sum"), toi5=("toi5", "sum"),
        icings=("icings", "sum"), pinned=("pinned", "sum"), stripped=("stripped", "sum"), d2o_starts=("d2o_starts", "sum"), d2o_ok=("d2o_ok", "sum"),
        trans_unblocked=("trans_unblocked", "sum"), trans_slot=("trans_slot", "sum")).reset_index()
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
    # the five candidate stats
    out["icing_pg"] = g["icings"] / hours
    out["pinned_pg"] = g["pinned"] / hours
    out["turnover_pg"] = (g["giveaways"] + g["stripped"]) / hours
    out["d2o_rate"] = g["d2o_ok"] / g["d2o_starts"]
    out["trans_slot_share"] = g["trans_slot"] / g["trans_unblocked"]
    out["d2o_starts"], out["trans_unblocked"] = g["d2o_starts"], g["trans_unblocked"]
    return out.replace([np.inf, -np.inf], np.nan)


def add_index(t):
    """z-scores within each season (flipped so higher = better), and the Breakout Index (weighted by WEIGHTS; equal by default)."""
    t = t.copy()
    for col, _, higher in ALL_STATS:                                      # the three index stats and the five candidates
        z = ps.zscore_by_season(t, col)
        t[f"z_{col}"] = z if higher else -z
    z = t[Z_COLS]
    w = pd.Series(WEIGHTS, index=Z_COLS) / sum(WEIGHTS)
    raw = ((z * w).sum(axis=1, min_count=1) / (z.notna() * w).sum(axis=1)).replace([np.inf, -np.inf], np.nan)
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
    """How the eight stats move together. Every stat is turned so that higher = better, so a positive number means 'good at one, good at the other'."""
    cols = ALL_Z
    names = {f"z_{c}": short(n) for c, n, _ in ALL_STATS}
    rows = []
    for a in cols:
        row = {"": names[a]}
        for b in cols:
            row[names[b]] = round(ps.season_spearman(t, a, b).mean(), 2) if a != b else None
        rows.append(row)
    return pd.DataFrame(rows)


def level_by_season(t):
    g = t.groupby("season_id").agg(exit_time=("exit_time", "mean"), quick_exit=("quick_exit", "mean"), exit_success=("exit_success", "mean"),
                                   costly_pg=("costly_pg", "mean"), costly_share=("costly_share", "mean"), trans_pg=("trans_pg", "mean"),
                                   icing_pg=("icing_pg", "mean"), pinned_pg=("pinned_pg", "mean"), turnover_pg=("turnover_pg", "mean"),
                                   d2o_rate=("d2o_rate", "mean"), trans_slot_share=("trans_slot_share", "mean")).reset_index()
    g.insert(0, "Season", g["season_id"].map(season_label))
    g = g.drop(columns="season_id").round(3)
    g.columns = ["Season", "Avg zone exit time (s)", "Share of exits within 8 s", "Share of attempts that cleared the zone", "Costly giveaways /60",
                 "Share of own-zone giveaways that were costly", "Transition shot attempts /60", "Icings /60", "Pinned-in spells /60",
                 "Own-zone turnovers /60", "Defence-to-offence rate", "Transition shots from the slot (share)"]
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
        out[cp] = add_index(summarize(r))[["season_id", "team_id", "z_breakout", "breakout"] + ALL_Z]
    return out


# ----------------------------------------------------------------------------
# The scorecard: the same tests for every stat, and the PASS RULE
# ----------------------------------------------------------------------------
def short(name):
    return name.split(" (")[0].split(":")[0]


def scorecard(team, game_rows, nhl, seasons, contender):
    """
    One row per stat (the three in the index and the five candidates). Every link is turned so that POSITIVE = good
    for the team (fewer icings, more defence-to-offence...). The PASS RULE at the top of this file decides the last column.
    """
    season_ids = sorted(int(s) for s in team["season_id"].unique() if s in set(seasons))
    n = len(season_ids)
    one_each = np.eye(len(ALL_Z))                                        # "mixes" that are simply each stat on its own
    targets = ["pts_pct", "gd_pg"]

    def stack(parts):
        return {k: np.vstack([p[k] for p in parts]) for k in parts[0]}

    same = stack([mix_links(team[team["season_id"] == s], one_each, targets, ALL_Z) for s in season_ids])
    halves = half_frames(game_rows, nhl, season_ids)
    parts = []
    for s in season_ids:
        a, b = (mix_links(h[h["season_id"] == s], one_each, targets, ALL_Z) for h in halves)
        parts.append({k: (a[k] + b[k]) / 2 for k in a})
    other = stack(parts)
    other_beyond = (other["pts_pct_beyond"] + other["gd_pg_beyond"]) / 2          # one row per season, one column per stat
    margin = 2 * other_beyond.std(axis=0, ddof=1) / np.sqrt(n) if n > 1 else np.full(len(ALL_Z), np.nan)

    odd = add_index(summarize(game_rows[game_rows["game_id"] % 2 == 1]))
    even = add_index(summarize(game_rows[game_rows["game_id"] % 2 == 0]))
    m = odd.merge(even, on=["season_id", "team_id"], suffixes=("_a", "_b"))
    rows = []
    for i, (col, name, _) in enumerate(ALL_STATS):
        half = float(ps.season_spearman(m, f"{col}_a", f"{col}_b").mean())
        full = 2 * half / (1 + half)
        d = team.dropna(subset=[f"z_{col}", "es_dz_pct"])
        edge = float(ps.season_spearman(d, f"z_{col}", "es_dz_pct").mean()) if len(d) > 30 else np.nan
        goal_diff_link = float(same["gd_pg"][:, i].mean())
        beyond, noise = float(other_beyond[:, i].mean()), float(margin[i])
        checks = [(full >= PASS_REPEAT, "does not repeat enough"), (goal_diff_link > 0, "points the wrong way"),
                  (beyond > noise, "adds nothing beyond Corsi that is clear of the noise")]
        gain, gain_margin = contender.get(col, (np.nan, np.nan))
        rows.append({
            "Statistic": name, "Status": "in the index" if i < len(COMPONENTS) else "candidate",
            "Repeats (odd vs even games)": round(half, 2), "Full-season estimate": round(full, 2),
            "Link with points %": round(float(same["pts_pct"][:, i].mean()), 2), "Link with goal difference": round(goal_diff_link, 2),
            "Same season, beyond Corsi": round(float(((same["pts_pct_beyond"] + same["gd_pg_beyond"]) / 2)[:, i].mean()), 2),
            "Other half of the season: link with results": round(float(((other["pts_pct"] + other["gd_pg"]) / 2)[:, i].mean()), 2),
            "Other half, beyond Corsi": round(beyond, 3), "Noise margin (+/-)": round(noise, 3),
            "Link with NHL own-zone time (negative = less time in own end)": round(edge, 2) if pd.notna(edge) else None,
            "Contender score gain (AUC)": round(gain, 3) if pd.notna(gain) else None,
            "Contender margin (+/-)": round(gain_margin, 3) if pd.notna(gain_margin) else None,
            "Passes all three rules?": "YES" if all(ok for ok, _ in checks) else "no",
            "Why not": "; ".join(text for ok, text in checks if not ok)})
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# The weight test: which mix of the three stats lines up best with good teams?
# ----------------------------------------------------------------------------
def weight_mixes():
    """Every mix of the three weights in steps of WEIGHT_STEP that adds up to 1, plus equal weights (always the last row)."""
    n = int(round(1 / WEIGHT_STEP))
    mixes = [(i / n, j / n, (n - i - j) / n) for i in range(n + 1) for j in range(n + 1 - i)]
    return np.array(mixes + [(1 / 3, 1 / 3, 1 / 3)])


def _rank_links(scores, target):
    """Rank correlation of each column of `scores` with `target` (the teams of one season)."""
    r = pd.DataFrame(np.asarray(scores)).rank().to_numpy()
    y = pd.Series(np.asarray(target)).rank().to_numpy()
    rc, yc = r - r.mean(axis=0), y - y.mean()
    return (rc * yc[:, None]).sum(axis=0) / np.sqrt((rc ** 2).sum(axis=0) * (yc ** 2).sum())


def mix_links(frame, mixes, targets, cols=None):
    """
    For one season: how every weight mix lines up with each target. Three versions per target:
      target            the plain rank correlation
      target_beyond     the same after taking out what Corsi already explains (a 'partial' correlation)
      target_corsi      Corsi's own link with the target (one number, the benchmark)
    """
    cols = cols or Z_COLS
    d = frame.dropna(subset=cols + targets + ["corsi"])
    scores = d[cols].to_numpy() @ mixes.T
    with_corsi = _rank_links(scores, d["corsi"])
    out = {}
    for t in targets:
        raw = _rank_links(scores, d[t])
        corsi_link = float(_rank_links(d[["corsi"]], d[t])[0])
        out[t] = raw
        out[f"{t}_beyond"] = (raw - with_corsi * corsi_link) / np.sqrt((1 - with_corsi ** 2) * (1 - corsi_link ** 2))
        out[f"{t}_corsi"] = np.full(len(mixes), corsi_link)
    return out


def half_frames(game_rows, nhl, seasons):
    """Two tables: the three stats from odd-numbered games next to the RESULTS of even-numbered games, and the reverse."""
    ids = ",".join(str(int(s)) for s in seasons)
    results = pd.read_sql_query(
        f"SELECT season_id, game_id, team_id, points, goalsFor - goalsAgainst AS gd FROM team_games WHERE game_type = 2 AND season_id IN ({ids})", nhl)
    bench = benchmark_halves(nhl, seasons)                               # corsi_a = odd games, corsi_b = even games
    frames = []
    for parity, suffix in ((1, "_a"), (0, "_b")):
        z = add_index(summarize(game_rows[game_rows["game_id"] % 2 == parity]))[["season_id", "team_id"] + ALL_Z]
        r = results[results["game_id"] % 2 != parity].groupby(["season_id", "team_id"]).agg(
            pts=("points", "sum"), n=("points", "size"), gd=("gd", "sum")).reset_index()
        r["pts_pct"], r["gd_pg"] = r["pts"] / (2 * r["n"]), r["gd"] / r["n"]
        c = bench[["season_id", "team_id", f"corsi{suffix}"]].rename(columns={f"corsi{suffix}": "corsi"})
        frames.append(z.merge(r[["season_id", "team_id", "pts_pct", "gd_pg"]], on=["season_id", "team_id"]).merge(c, on=["season_id", "team_id"]))
    return frames


def weight_test(team, game_rows, nhl, seasons):
    """
    Tries every mix of the three weights and asks which lines up best with good teams (points % and goal difference).
    To stay honest the 'best' mix is CHOSEN on all seasons but one and SCORED on the season left out, then compared
    with plain equal weights. Returns a dictionary of tables and a few numbers for the verdict.
    """
    mixes = weight_mixes()
    equal = len(mixes) - 1
    grid = slice(0, equal)                                               # the mixes that can be chosen (equal weights is the yardstick)
    targets = ["pts_pct", "gd_pg"]
    season_ids = sorted(int(s) for s in team["season_id"].unique() if s in set(seasons))
    n = len(season_ids)

    def stack(parts):
        return {k: np.vstack([p[k] for p in parts]) for k in parts[0]}

    same = stack([mix_links(team[team["season_id"] == s], mixes, targets + ["xg_share"]) for s in season_ids])
    halves = half_frames(game_rows, nhl, season_ids)
    other_parts = []
    for s in season_ids:
        a, b = (mix_links(h[h["season_id"] == s], mixes, targets) for h in halves)
        other_parts.append({k: (a[k] + b[k]) / 2 for k in a})
    other = stack(other_parts)

    goal = (same["pts_pct"] + same["gd_pg"]) / 2                         # what "best" means: the average link with points % and goal difference
    chosen = np.array([int(np.delete(goal[:, grid], i, axis=0).mean(axis=0).argmax()) for i in range(n)])
    rows_i = np.arange(n)
    singles = [int(np.where((np.abs(mixes[grid] - np.eye(3)[k]) < 1e-9).all(axis=1))[0][0]) for k in range(3)]
    methods = [("Equal weights (one third each)", np.full(n, equal)), ("Best mix, chosen on the OTHER seasons", chosen)] + \
              [(f"{SHORT_NAMES[k]} only", np.full(n, singles[k])) for k in range(3)]
    columns = [("Same season: points %", same, "pts_pct"), ("Same season: goal difference", same, "gd_pg"),
               ("Same season: expected-goals share", same, "xg_share"),
               ("Same season, beyond Corsi: points %", same, "pts_pct_beyond"), ("Same season, beyond Corsi: goal difference", same, "gd_pg_beyond"),
               ("Other half of the season: points %", other, "pts_pct"), ("Other half of the season: goal difference", other, "gd_pg"),
               ("Other half, beyond Corsi: points %", other, "pts_pct_beyond"), ("Other half, beyond Corsi: goal difference", other, "gd_pg_beyond")]
    summary = []
    for name, picks in methods:
        row = {"Weights used": name}
        for label, source, key in columns:
            row[label] = round(float(source[key][rows_i, picks].mean()), 2)
        summary.append(row)
    corsi_row = {"Weights used": "For comparison: Corsi alone"}
    for label, source, key in columns:
        corsi_row[label] = None if key.endswith("_beyond") else round(float(source[f"{key}_corsi"][:, 0].mean()), 2)
    summary.append(corsi_row)

    held_chosen, held_equal = goal[rows_i, chosen], goal[:, equal]
    diff = held_chosen - held_equal
    folds = pd.DataFrame({"Season left out": [season_label(s) for s in season_ids],
                          "Weight chosen: zone exit time": mixes[chosen, 0], "Weight chosen: costly giveaways": mixes[chosen, 1],
                          "Weight chosen: transition shots": mixes[chosen, 2],
                          "Score of the chosen mix on the season left out": held_chosen.round(3),
                          "Score of equal weights on that season": held_equal.round(3), "Difference": diff.round(3)})

    table = pd.DataFrame({"Weight: zone exit time": mixes[:, 0].round(2), "Weight: costly giveaways": mixes[:, 1].round(2),
                          "Weight: transition shots": mixes[:, 2].round(2),
                          "Score (avg link with points % and goal difference)": goal.mean(axis=0).round(3),
                          "Link with points %": same["pts_pct"].mean(axis=0).round(3), "Link with goal difference": same["gd_pg"].mean(axis=0).round(3),
                          "Beyond Corsi (avg of the two)": ((same["pts_pct_beyond"] + same["gd_pg_beyond"]) / 2).mean(axis=0).round(3),
                          "Other half of the season (avg of the two)": ((other["pts_pct"] + other["gd_pg"]) / 2).mean(axis=0).round(3),
                          "Seasons where it beat equal weights": (goal > goal[:, [equal]]).sum(axis=0)})
    table.insert(0, "Mix", ["equal weights" if i == equal else "" for i in range(len(mixes))])
    table = table.sort_values("Score (avg link with points % and goal difference)", ascending=False).reset_index(drop=True)
    table.insert(0, "Rank", np.arange(1, len(table) + 1))

    all_mean = goal.mean(axis=0)
    best = int(all_mean[grid].argmax())
    return {"summary": pd.DataFrame(summary), "folds": folds, "grid": table, "n_mixes": equal, "n_seasons": n,
            "best_weights": mixes[best], "best_score": float(all_mean[best]), "equal_score": float(all_mean[equal]),
            "worst_score": float(all_mean[grid].min()), "held_chosen": float(held_chosen.mean()), "held_equal": float(held_equal.mean()),
            "held_diff": float(diff.mean()), "held_margin": float(2 * diff.std(ddof=1) / np.sqrt(n)) if n > 1 else float("nan"),
            "seasons_better": int((diff > 0).sum()), "avg_chosen": mixes[chosen].mean(axis=0),
            "beyond_equal": float(((same["pts_pct_beyond"] + same["gd_pg_beyond"]) / 2)[:, equal].mean()),
            "beyond_best": float(((same["pts_pct_beyond"] + same["gd_pg_beyond"]) / 2)[rows_i, chosen].mean()),
            "other_equal": float(((other["pts_pct"] + other["gd_pg"]) / 2)[:, equal].mean()),
            "other_best": float(((other["pts_pct"] + other["gd_pg"]) / 2)[rows_i, chosen].mean()),
            "other_beyond_equal": float(((other["pts_pct_beyond"] + other["gd_pg_beyond"]) / 2)[:, equal].mean()),
            "other_beyond_best": float(((other["pts_pct_beyond"] + other["gd_pg_beyond"]) / 2)[rows_i, chosen].mean())}


def forecast_test(connection, game_rows, covered, force):
    """Does the index add to the contender score? Returns (table, note)."""
    if len(covered) < MIN_SEASONS_FOR_FORECAST and not force:
        return pd.DataFrame(), (f"Not run yet: only {len(covered)} seasons of play-by-play are downloaded and the hold-out test needs "
                                f"{MIN_SEASONS_FOR_FORECAST} or more. Rerun this script when the download has finished."), {}
    games_raw, seasons, checkpoints = cf.load(connection)
    games, _ = ps.prepare_games(games_raw, seasons)
    tables = cf.build_checkpoint_tables(games, seasons, checkpoints)
    feats = checkpoint_features(game_rows, connection, checkpoints)
    v1 = cf.MODELS["Model 3: + power kill (PP% + PK%)"]
    rows = []
    one_at_a_time = {c: [] for c, _, _ in ALL_STATS}
    for cp in CP_ORDER:
        t = tables[cp].merge(feats[cp], on=["season_id", "team_id"], how="inner")
        t = t[t["season_id"].isin(covered)].copy()
        if t["season_id"].nunique() < 3:
            continue
        t[ALL_Z] = t[ALL_Z].fillna(0.0)
        t["v1"] = cf.held_out_scores(t, v1)
        t["v1+breakout"] = cf.held_out_scores(t, v1 + ["breakout"])
        t["v1+three"] = cf.held_out_scores(t, v1 + [c for c, _, _ in COMPONENTS])     # the model picks its own weight for each stat
        a, b, c3 = cf.per_season_metrics(t, "v1"), cf.per_season_metrics(t, "v1+breakout"), cf.per_season_metrics(t, "v1+three")
        diff, diff3 = b["auc"] - a["auc"], c3["auc"] - a["auc"]
        rows.append({"Checkpoint": PLAIN_CP[cp], "Seasons": len(diff), "AUC v1": a["auc"].mean(), "AUC v1 + Breakout Index": b["auc"].mean(),
                     "Gain": diff.mean(), "Margin (+/-, 2 std. errors)": 2 * diff.std(ddof=1) / np.sqrt(len(diff)),
                     "AUC v1 + the three stats separately (model picks the weights)": c3["auc"].mean(),
                     "Gain (separately)": diff3.mean(), "Margin (separately)": 2 * diff3.std(ddof=1) / np.sqrt(len(diff3)),
                     "Avg finalist rank v1": a["avg_rank_f4"].mean(), "Avg finalist rank with index": b["avg_rank_f4"].mean(),
                     "Seasons better": int((b["avg_rank_f4"] < a["avg_rank_f4"]).sum()), "Seasons worse": int((b["avg_rank_f4"] > a["avg_rank_f4"]).sum())})
        for col, _, _ in ALL_STATS:                                      # each stat added to version 1 on its own
            t["one"] = cf.held_out_scores(t, v1 + [col])
            one_at_a_time[col].append(cf.per_season_metrics(t, "one")["auc"] - a["auc"])
    table = pd.DataFrame(rows).round(3)
    note = "Hold-out test run on " + f"{len(covered)} seasons" + (" (FORCED with fewer than 8 seasons: treat as a code check, not a result)." if len(covered) < MIN_SEASONS_FOR_FORECAST else ".")
    contender = {}
    for col, gains in one_at_a_time.items():
        if gains:
            by_season = pd.concat(gains, axis=1).mean(axis=1)            # one number per season: the gain averaged over the five checkpoints
            contender[col] = (float(by_season.mean()), float(2 * by_season.std(ddof=1) / np.sqrt(len(by_season))))
    return table, note, contender


# ----------------------------------------------------------------------------
def weights_sentence(w):
    pct = lambda x: f"{round(float(x) * 100):.0f}%"                       # noqa: E731
    best, avg = w["best_weights"], w["avg_chosen"]
    better = w["held_diff"] > w["held_margin"]
    return (f"Tried all {w['n_mixes']} mixes of the three weights in steps of {round(WEIGHT_STEP * 100)}%, scoring each by its average link with points % and goal difference. "
            f"Across all {w['n_seasons']} seasons the top mix is zone exit time {pct(best[0])}, costly giveaways {pct(best[1])}, transition shots {pct(best[2])} "
            f"(score {w['best_score']:.2f}; equal weights {w['equal_score']:.2f}; the worst mix {w['worst_score']:.2f}). "
            f"The honest check: choose the mix on all seasons but one and score it on the season left out. That gives {w['held_chosen']:.2f} against {w['held_equal']:.2f} for equal weights "
            f"(difference {w['held_diff']:+.2f}, noise margin about +/-{w['held_margin']:.2f}, better in {w['seasons_better']} of {w['n_seasons']} seasons); "
            f"the mix chosen averaged {pct(avg[0])} / {pct(avg[1])} / {pct(avg[2])}. "
            + ("The chosen weights beat equal weights by more than the noise." if better else
               "That is within the noise: re-weighting does not reliably beat equal weights.")
            + f" After taking out what Corsi already explains, the link is {w['beyond_best']:+.2f} with the chosen mix and {w['beyond_equal']:+.2f} with equal weights. "
              f"Predicting the OTHER half of the same season: {w['other_best']:+.2f} (chosen) and {w['other_equal']:+.2f} (equal); beyond Corsi {w['other_beyond_best']:+.2f} and {w['other_beyond_equal']:+.2f}.")


def card_sentence(card):
    yes = card["Passes all three rules?"] == "YES"
    candidates, current = card[card["Status"] == "candidate"], card[card["Status"] == "in the index"]
    parts = []
    for _, r in candidates.iterrows():
        result = "PASSES" if r["Passes all three rules?"] == "YES" else "fails (" + r["Why not"] + ")"
        parts.append(f"{short(r['Statistic'])}: repeats {r['Full-season estimate']:.2f}, link with goal difference {r['Link with goal difference']:+.2f}, "
                     f"beyond Corsi in the other half {r['Other half, beyond Corsi']:+.3f} (noise +/-{r['Noise margin (+/-)']:.3f}): {result}")
    return (f"The rule was fixed before looking: a stat must repeat ({PASS_REPEAT:.2f} or more over a full season), point the right way, and add something beyond Corsi "
            f"in the OTHER half of the season that is clear of the noise. {int(yes[candidates.index].sum())} of {len(candidates)} new stats pass. "
            + " | ".join(parts)
            + f" | Held to the same rule, {int(yes[current.index].sum())} of the {len(current)} stats already in the index pass."
            + " With eight stats tested, one narrow pass can still be luck.")


def verdict_lines(reliab, links, level, forecast, forecast_note, champs, n_seasons, season_list, weights=None, card=None):
    rel = reliab.set_index("Statistic")["Repeats: odd vs even games"]
    link = links.set_index("Breakout statistic")
    idx = link.loc["Breakout Index (higher is better)"]
    corsi_link = idx["Corsi (shot-attempt share, 5v5)"]
    pts_link = idx["Points %"]
    edge_exit = link.iloc[0]["NHL zone time: share in own zone (5v5, 2021-22 on)"]
    edge_idx = idx["NHL zone time: share in offensive zone (5v5, 2021-22 on)"]
    d = champs.set_index("How the season ended")
    lines = [
        ("What this is", f"A prototype of your Breakout Index on {n_seasons} fully downloaded seasons ({season_list}). Seasons 2009-10 to 2013-14 are kept back for a one-time blind test of the final metric."),
        ("Does it repeat?", f"Odd-game vs even-game rank correlation (half a season each): zone exit time {rel.iloc[0]:.2f}, costly giveaways {rel.iloc[1]:.2f}, transition shots {rel.iloc[2]:.2f}, "
                            f"Breakout Index {rel.iloc[3]:.2f}. For comparison, split the same way: Corsi {rel.iloc[4]:.2f}, expected-goals share {rel.iloc[5]:.2f}, points % {rel.iloc[6]:.2f}. "
                            "Closer to 1 means more of a stable team trait and less luck."),
        ("Does it line up with winning?", f"The index's link with points % is {pts_link:+.2f} and with shot-attempt share (Corsi) {corsi_link:+.2f}. "
                                          + ("The link with Corsi is modest, so it is partly a different thing." if abs(corsi_link) < 0.35 else
                                             "It overlaps with Corsi to a real degree (about 0.3 to 0.5 is moderate), so it may add little beyond it. The contender test below is the real check.")),
        ("Does it match the NHL's real zone time?", (f"Zone exit time vs share of 5v5 time in the team's own zone: {edge_exit:+.2f} (positive is what we expect: slow exits, more time in your own end). "
                                                      f"Breakout Index vs offensive-zone share: {edge_idx:+.2f}.") if pd.notna(edge_exit) else "Not testable yet: no 2021-22-or-later seasons have finished downloading."),
        ("Which weights work best?", weights_sentence(weights) if weights else "Not run."),
        ("The five new stats", card_sentence(card) if card is not None else "Not run."),
        ("Does it improve the contender score?", forecast_note + (" " + _forecast_sentence(forecast) if len(forecast) else "")),
        ("Champions", f"Average Breakout Index percentile: Won Cup {d.loc['Won Cup', 'Breakout Index (percentile)']:.0f}, Lost Cup final {d.loc['Lost Cup final', 'Breakout Index (percentile)']:.0f}, "
                      f"Missed playoffs {d.loc['Missed playoffs', 'Breakout Index (percentile)']:.0f} (100 = best, 50 = average)." if "Won Cup" in d.index else "No champions in the seasons downloaded so far."),
    ]
    return pd.DataFrame(lines, columns=["", "Answer"])


def _forecast_sentence(forecast):
    gain, margin = forecast["Gain"].mean(), forecast["Margin (+/-, 2 std. errors)"].mean()
    gain3, margin3 = forecast["Gain (separately)"].mean(), forecast["Margin (separately)"].mean()
    return (f"Adding the index to version 1 changes the average AUC by {gain:+.3f} (margin about +/-{margin:.3f}): "
            + ("larger than the noise." if gain > margin else "within the noise.")
            + f" Letting the model pick its own weight for each of the three stats changes it by {gain3:+.3f} (margin about +/-{margin3:.3f}): "
            + ("larger than the noise." if gain3 > margin3 else "within the noise."))


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
    global EXIT_CAP, WEIGHTS                                             # so --exit-cap and --weights can change the settings for the whole run
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Build the Breakout Index from play-by-play.")
    parser.add_argument("--seasons", help="comma-separated season ids to use, e.g. 20232024,20242025")
    parser.add_argument("--force-forecast", action="store_true", help="run the hold-out test with fewer than 8 seasons (unreliable)")
    parser.add_argument("--exit-cap", type=int, default=EXIT_CAP, help=f"seconds after which a zone exit counts as failed (default {EXIT_CAP})")
    parser.add_argument("--weights", help="weights for zone exit time, costly giveaways, transition shots, e.g. 0.5,0.2,0.3 (default: equal)")
    args = parser.parse_args()
    EXIT_CAP = args.exit_cap
    if args.weights:
        try:
            WEIGHTS = tuple(float(x) for x in args.weights.split(","))
            assert len(WEIGHTS) == 3 and min(WEIGHTS) >= 0 and sum(WEIGHTS) > 0
        except (ValueError, AssertionError):
            print("--weights needs three numbers that are zero or more, separated by commas, for example  --weights 0.5,0.2,0.3")
            return 1
    if not PBP_DB.exists() or not NHL_DB.exists():
        print("I need data/pbp.db (python scripts/pull_play_by_play.py) and data/nhl.db (python scripts/build_database.py).")
        return 1
    pbp = sqlite3.connect(f"file:{PBP_DB}?mode=ro", uri=True)           # read-only: safe while the download is running
    nhl = sqlite3.connect(NHL_DB)
    if "mp_team_games" not in [r[0] for r in nhl.execute("SELECT name FROM sqlite_master WHERE type='table'")]:
        print("The MoneyPuck tables are missing (they give 5-on-5 time on ice). Run: python scripts/load_moneypuck.py")
        return 1

    loaded = load_game_rows(pbp, nhl, args.seasons)
    if loaded is None:
        return 1
    game_rows, usable, odd_seasons = loaded

    team = add_index(summarize(game_rows))
    compare = comparison_numbers(nhl, usable)
    team = team.merge(compare, on=["season_id", "team_id"], how="left")
    main_team = team[team["include_main"].fillna("Y") == "Y"]
    reliab = split_half(game_rows, nhl, usable)
    links = links_table(team)
    internal = internal_links(team)
    level = level_by_season(team)
    champs = champion_profile(team[team["include_main"].fillna("Y") == "Y"])
    normal = [s for s in usable if s not in odd_seasons]
    forecast, note, contender = forecast_test(nhl, game_rows[game_rows["season_id"].isin(normal)], normal, args.force_forecast)
    weights = weight_test(main_team, game_rows, nhl, normal) if len(usable) >= 3 else None
    card = scorecard(main_team, game_rows, nhl, normal, contender) if len(usable) >= 3 else None
    verdict = verdict_lines(reliab, links, level, forecast, note, champs, len(usable), ", ".join(season_label(s) for s in usable), weights, card)
    used = " / ".join(f"{w / sum(WEIGHTS):.0%}" for w in WEIGHTS)
    print(f"Weights used for the index in this run (zone exit time / costly giveaways / transition shots): {used}")

    out = team.sort_values(["season_id", "breakout"], ascending=[False, False])
    teams_out = pd.DataFrame({
        "Season": out["season_id"].map(season_label), "Team": out["team"], "Breakout Index (z, higher is better)": out["breakout"].round(2),
        "Avg zone exit time (s)": out["exit_time"].round(2), "Share of exits within 8 s": out["quick_exit"].round(3),
        "Share of attempts that cleared the zone": out["exit_success"].round(3), "Costly giveaways /60": out["costly_pg"].round(2),
        "Share of own-zone giveaways that were costly": out["costly_share"].round(3), "Transition shot attempts /60": out["trans_pg"].round(2),
        "Transition shot attempts against /60": out["trans_against_pg"].round(2), "Transition shot share": out["trans_share"].round(3),
        "Own-zone starts counted": out["exit_starts"].astype(int),
        "Icings /60": out["icing_pg"].round(2), "Pinned-in spells /60": out["pinned_pg"].round(2), "Own-zone turnovers /60": out["turnover_pg"].round(2),
        "Defence-to-offence rate": out["d2o_rate"].round(3), "Transition shots from the slot (share)": out["trans_slot_share"].round(3),
        "Points %": out["pts_pct"].round(3), "Corsi % (5v5)": (out["corsi"] * 100).round(1),
        "How the season ended": out["outcome"]}).reset_index(drop=True)

    sheets = {"Verdict": verdict, "Repeats": reliab, "Links": links, "Between the stats": internal, "Level by season": level,
              "Champions": champs, "Contender test": forecast if len(forecast) else pd.DataFrame({"": [note]}), "Team-seasons": teams_out}
    notes = {"Links": "Rank correlation (within each season, averaged) between each breakout statistic and other team numbers. Positive is expected for the index; for exit time and costly giveaways negative is 'good'.",
             "Between the stats": "How the eight stats move together (rank correlation within each season, averaged). Each is turned so higher = better: a positive number means teams good at one tend to be good at the other."}
    if card is not None:
        sheets = {"Verdict": verdict, "New stats": card, **{k: v for k, v in sheets.items() if k != "Verdict"}}
        notes["New stats"] = ("The same tests for every stat. Every link is turned so that POSITIVE = good for the team. A stat passes only if it repeats (full-season estimate 0.50+), "
                              "points the right way, and its 'Other half, beyond Corsi' number is bigger than its noise margin. The rule was fixed before the results were looked at.")
    if weights:
        sheets = {"Verdict": verdict, "Weights": weights["summary"], "Weights by season": weights["folds"], "Weight grid": weights["grid"],
                  **{k: v for k, v in sheets.items() if k != "Verdict"}}
        notes["Weights"] = ("How well each way of weighting the three stats lines up with good teams (rank correlation, averaged over seasons; higher is better). "
                            "'Beyond Corsi' = after taking out what shot-attempt share already explains. 'Other half' = stats from odd games vs results in even games, and the reverse.")
        notes["Weights by season"] = "The honest check: the best mix is chosen WITHOUT the season on that row, then scored on it. If the chosen weights jump around from row to row, no mix is clearly best."
        notes["Weight grid"] = "Every mix tried, best first, using all seasons at once (this flatters the top rows a little: the 'Weights by season' sheet is the fair test)."
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
    print("\nBetween the stats (each turned so higher = better):"); print(internal.to_string(index=False))
    if card is not None:
        print("\nScorecard, one column per stat:"); print(card.set_index("Statistic").T.rename(columns=short).to_string())
    print("\nChampions:"); print(champs.to_string(index=False))
    if weights:
        print("\nWeights: how each weighting lines up with good teams:"); print(weights["summary"].T.to_string(header=False))
        print("\nWeights chosen when each season is left out:"); print(weights["folds"].to_string(index=False))
        print("\nTop 8 mixes on all seasons (and where equal weights ranks):")
        g = weights["grid"]
        print(pd.concat([g.head(8), g[g["Mix"] == "equal weights"], g.tail(3)]).drop_duplicates().to_string(index=False))
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

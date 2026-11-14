"""
build_goaltending_study.py  -  goaltending in context: goalie bailing out the defense, or defense let down by the goalie?
===========================================================================================================================
WHAT IT DOES (in Excel terms):
    Uses MoneyPuck's shot-quality numbers to split every team's goals against into two parts:

        DEFENSE   how many goals the team's defense should allow, from the shots it lets through and how
                  dangerous they are:  "expected goals against per 60 minutes at 5-on-5" (xGA/60). Lower is better.
        GOALIE    how many goals the goalies saved or let in compared with what those shots should produce:
                  "goals saved above expected" (GSAx), per 60 minutes, all situations. Higher is better.

        (goals against = expected goals against - goals saved above expected. This is exact arithmetic.)

    Each team-season gets a label (above / below the league median on each part):
        Both good  |  Goalie bailing out the defense  |  Defense let down by the goalie  |  Both weak

    Then it asks the questions that matter:
      1. WHICH PART IS REAL?  How much of the difference between teams' goals against comes from defense and how
         much from goaltending, and which part REPEATS (before Jan 1 vs after)?
      2. DO THE LABELS PREDICT?  Do teams "bailed out" by their goalie slip later, and teams "let down" recover?
      3. THE HOT GOALIE: in the playoffs, does any regular-season signal (goaltending results, the last 20 games,
         raw save %, the starter's previous season, defense quality) foresee which team's goalie gets hot?
         And did the champions have hot goalies?

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py
    python scripts/load_moneypuck.py          (adds the MoneyPuck tables)
    python scripts/build_goaltending_study.py

WHAT YOU SHOULD SEE:
    A summary in the terminal and a new file:  outputs/goaltending_in_context.xlsx  (close it in Excel first).
    Start with its "Verdict" sheet.

DATA CREDIT:
    Shot-quality data is from MoneyPuck.com (non-commercial use; credit MoneyPuck.com wherever it is shown).
    Other data is (c) NHL, used for personal, non-commercial purposes.
"""

import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_popular_stats_check as ps    # stat helpers and the hold-out forecast test

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
DB_FILE = PROJECT_FOLDER / "data" / "nhl.db"
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "goaltending_in_context.xlsx"

FIRST_SEASON = ps.FIRST_SEASON          # 2009-10: the 15 main seasons used in the other studies
SPLIT = "jan1"                          # about the middle of the season
LABELS = ["Both good", "Goalie bailing out the defense", "Defense let down by the goalie", "Both weak"]
ROUNDS = {"Lost round 1": 0, "Lost round 2": 1, "Lost conference final": 2, "Lost Cup final": 3, "Won Cup": 4}

READ_ME = [
    "Goaltending in context: goalie bailing out the defense, or defense let down by the goalie?",
    "",
    "THE TWO PARTS OF GOALS AGAINST (data: MoneyPuck shot quality; seasons 2009-10 on, 15 main seasons)",
    "  Defense = expected goals against per 60 minutes at 5-on-5 (xGA/60). It counts every unblocked shot allowed, weighted by how",
    "            likely it was to be a goal (location, shot type, rebounds ...). Lower = better defense.",
    "  Goalie  = goals saved above expected per 60 minutes, all situations (GSAx/60) = (expected goals against - goals against).",
    "            Higher = the goalies stopped more than the shots should have allowed.",
    "  Goals against = expected goals against - GSAx. So differences between teams' goals against are exactly the sum of",
    "  differences in defense and in goaltending.",
    "",
    "THE LABELS (above / below the league median that season, on each part)",
    "  Both good | Goalie bailing out the defense (weak defense, strong goalie) | Defense let down by the goalie (strong defense, weak goalie) | Both weak.",
    "",
    "HOW TO READ THE TESTS",
    "  Repeats: correlation (on ranks, within each season) between the stat before Jan 1 and after Jan 1. Low = mostly luck / noise.",
    "  After-Jan-1 results by label: how a team's points % after Jan 1 compares with what its record before Jan 1 predicted.",
    "  Forecast gain: same hold-out test as before (the rule is built on the other seasons, then used on the one it has never seen).",
    "  Playoffs: playoff GSAx/60 is the team's goaltending result over the playoff games it played (all situations).",
    "  Correlations on ranks: about 0.3 is modest, 0.5 strong. With about 240 playoff team-runs, differences under about 0.13 are noise.",
    "",
    "WHAT THIS CANNOT SAY",
    "  Team-level only: a team's goalie numbers blend all of its goalies (including a backup). Playoff goalie-level data is not in the",
    "  free files. MoneyPuck's expected-goals model is one reasonable model, not the truth: shot quality it cannot see (screens, passes",
    "  before the shot) shows up as 'goalie' results. 'Defense' here is shot quality allowed, not individual defenders.",
    "",
    "Data: MoneyPuck.com (non-commercial use) and (c) NHL.",
]


# ----------------------------------------------------------------------------
# Loading
# ----------------------------------------------------------------------------
def load(connection):
    games = pd.read_sql_query(
        """SELECT season_id, team_id, game_id, game_date, playoffGame, situation, iceTime, xGoalsAgainst, goalsAgainst,
                  shotsOnGoalAgainst, highDangerShotsAgainst, highDangerxGoalsAgainst, highDangerGoalsAgainst
           FROM mp_team_games WHERE situation IN ('all', '5on5') AND season_id >= ?""", connection, params=[FIRST_SEASON])
    keys = ["season_id", "team_id", "game_id", "game_date", "playoffGame"]
    a = games[games["situation"] == "all"].drop(columns="situation")
    f = games[games["situation"] == "5on5"].drop(columns="situation")
    wide = a.merge(f, on=keys, suffixes=("", "_5"))
    seasons = pd.read_sql_query(
        """SELECT season_id, season, team_id, team, include_main, outcome, made_playoffs, conf_finalist, cup_finalist, champion
           FROM team_seasons WHERE season_id >= ?""", connection, params=[FIRST_SEASON])
    checkpoints = pd.read_sql_query("SELECT season_id, checkpoint, date FROM checkpoints WHERE reached = 1", connection)
    goalies = pd.read_sql_query(
        """SELECT season, season_id, team_id, playerId, name, games_played, icetime, xGoals, goals, ongoal
           FROM mp_goalie_seasons WHERE situation = 'all'""", connection)
    return wide, seasons, checkpoints, goalies


def team_stats(w):
    """Defense and goaltending numbers for every team-season, from whichever games are passed in."""
    g = w.groupby(["season_id", "team_id"]).agg(
        games=("game_id", "count"), ice=("iceTime", "sum"), xga=("xGoalsAgainst", "sum"), ga=("goalsAgainst", "sum"),
        soga=("shotsOnGoalAgainst", "sum"), ice5=("iceTime_5", "sum"), xga5=("xGoalsAgainst_5", "sum"),
        ga5=("goalsAgainst_5", "sum"), soga5=("shotsOnGoalAgainst_5", "sum"), hda5=("highDangerShotsAgainst_5", "sum"),
        hdxga5=("highDangerxGoalsAgainst_5", "sum")).reset_index()
    out = g[["season_id", "team_id", "games"]].copy()
    out["xga60_5"] = g["xga5"] / (g["ice5"] / 3600)             # defense: expected goals against per 60, 5v5
    out["hda60_5"] = g["hda5"] / (g["ice5"] / 3600)             # defense: high-danger shots against per 60, 5v5
    out["gsax60"] = (g["xga"] - g["ga"]) / (g["ice"] / 3600)    # goalie: goals saved above expected per 60, all situations
    out["gsax"] = g["xga"] - g["ga"]
    out["sv"] = 1 - g["ga"] / g["soga"]                         # raw save %, all situations
    out["sv5"] = 1 - g["ga5"] / g["soga5"]
    out["ga_pg"] = g["ga"] / g["games"]
    out["xga_pg"] = g["xga"] / g["games"]
    out["gsax_pg"] = (g["xga"] - g["ga"]) / g["games"]
    return out.replace([np.inf, -np.inf], np.nan)


def label_teams(frame, defense="xga60_5", goalie="gsax60"):
    """Add the four labels (above / below the league median that season on each part)."""
    f = frame.copy()
    d_good = f[defense] <= f.groupby("season_id")[defense].transform("median")
    g_good = f[goalie] >= f.groupby("season_id")[goalie].transform("median")
    f["label"] = np.select([d_good & g_good, ~d_good & g_good, d_good & ~g_good], LABELS[:3], default=LABELS[3])
    return f


# ----------------------------------------------------------------------------
# Question 1: which part is real, and which repeats?
# ----------------------------------------------------------------------------
def split_frames(wide, checkpoints, cp):
    """Stats before and after a checkpoint date (regular season only)."""
    reg = wide[wide["playoffGame"] == 0]
    dates = checkpoints[checkpoints["checkpoint"] == cp][["season_id", "date"]]
    t = reg.merge(dates, on="season_id")
    before = team_stats(t[t["game_date"] <= t["date"]]).add_suffix("_td").rename(columns={"season_id_td": "season_id", "team_id_td": "team_id"})
    after = team_stats(t[t["game_date"] > t["date"]]).add_suffix("_rest").rename(columns={"season_id_rest": "season_id", "team_id_rest": "team_id"})
    return before.merge(after, on=["season_id", "team_id"])


STAT_NAMES = [
    ("xga60_5", "Defense: expected goals against per 60 (5v5)", "lower is better"),
    ("hda60_5", "Defense: high-danger shots against per 60 (5v5)", "lower is better"),
    ("gsax60", "Goalie: goals saved above expected per 60", "higher is better"),
    ("sv", "Goalie: raw save % (all situations)", "higher is better"),
    ("sv5", "Goalie: raw save % (5v5)", "higher is better"),
    ("ga_pg", "Result: goals against per game", "lower is better"),
]


def stability_table(frame):
    rows = []
    for key, name, direction in STAT_NAMES:
        repeats = ps.season_spearman(frame, f"{key}_td", f"{key}_rest")
        rows.append({"Statistic": name, "Better when": direction, "Repeats (before vs after Jan 1)": round(repeats.mean(), 2),
                     "Lowest season": round(repeats.min(), 2), "Highest season": round(repeats.max(), 2),
                     "Carry-over (share of the gap still there)": f"{ps.carry_over(frame, key):.0%}"})
    return pd.DataFrame(rows)


def goals_against_split(full):
    """How much of the spread in goals against comes from defense, how much from goaltending? (full seasons)"""
    d = full.dropna(subset=["xga_pg", "gsax_pg", "ga_pg"]).copy()
    for col in ("xga_pg", "gsax_pg", "ga_pg"):
        d[f"{col}_dev"] = d[col] - d.groupby("season_id")[col].transform("mean")
    var_ga = (d["ga_pg_dev"] ** 2).mean()
    var_x, var_g = (d["xga_pg_dev"] ** 2).mean(), (d["gsax_pg_dev"] ** 2).mean()
    cov = (d["xga_pg_dev"] * (-d["gsax_pg_dev"])).mean()                     # goals against = xGA - GSAx
    rows = [
        ("Spread of goals against per game (std. dev.)", round(np.sqrt(var_ga), 3), "how much teams differ in goals against"),
        ("  from defense (expected goals against)", round(np.sqrt(var_x), 3), f"{var_x / var_ga:.0%} of the variance"),
        ("  from goaltending (goals saved above expected)", round(np.sqrt(var_g), 3), f"{var_g / var_ga:.0%} of the variance"),
        ("  overlap between the two", round(var_x / var_ga + var_g / var_ga - 1, 3),
         "the two shares add to more than 100% by this much, because teams that allow more danger also get a little more help from goaltending"),
        ("Correlation: defense vs goaltending across teams", round(d["xga_pg"].corr(d["gsax_pg"]), 3),
         "positive = teams that allow more danger get MORE help from goaltending, as if goalies 'cover' for the defense"),
    ]
    return pd.DataFrame(rows, columns=["", "Value", "Meaning"]), var_x / var_ga, var_g / var_ga


# ----------------------------------------------------------------------------
# Question 2: do the labels predict what happens after Jan 1?
# ----------------------------------------------------------------------------
def label_followup(frame):
    """Compare each label's points % after Jan 1 with what the record before Jan 1 predicted."""
    d = frame.dropna(subset=["xga60_5_td", "gsax60_td", "pts_pct_td", "pts_pct_rest"]).copy()
    d = label_teams(d.rename(columns={"xga60_5_td": "xga60_5", "gsax60_td": "gsax60"}))
    d["x"] = d["pts_pct_td"] - d.groupby("season_id")["pts_pct_td"].transform("mean")
    d["y"] = d["pts_pct_rest"] - d.groupby("season_id")["pts_pct_rest"].transform("mean")
    slope = float((d["x"] * d["y"]).sum() / (d["x"] ** 2).sum())
    d["resid"] = d["y"] - slope * d["x"]
    pts = 2 * d["gp_rest"].mean()                                           # points % -> standings points over the rest
    rows = []
    for label in LABELS:
        g = d[d["label"] == label]
        se = g["resid"].std() / np.sqrt(len(g))
        rows.append({"Label (as of Jan 1)": label, "Team-seasons": len(g),
                     "Points % before Jan 1": round(g["pts_pct_td"].mean(), 3),
                     "Points % after Jan 1": round(g["pts_pct_rest"].mean(), 3),
                     "Points % after, vs what the record predicted": round(g["resid"].mean(), 4),
                     "= standings points over the rest of the season": round(g["resid"].mean() * pts, 1),
                     "Margin (+/-, 2 std. errors, in points)": round(2 * se * pts, 1)})
    return pd.DataFrame(rows)


def forecast_gain(frame):
    """Does each defense / goalie stat improve the forecast of points % after Jan 1, beyond record and goal difference?"""
    rows = []
    for key, name, _ in STAT_NAMES[:4]:
        d = frame.copy()
        result = ps.predict_gain(d, key)
        rows.append({"Statistic (before Jan 1)": name, "Forecast without it (R-sq)": round(result["base"], 3),
                     "Forecast with it (R-sq)": round(result["model"], 3), "Gain": round(result["gain"], 3),
                     "Seasons it helped (of 15)": result["improved"],
                     "Effect of +1 standard deviation (standings points, rest of season)": round(result["effect_pts"], 1)})
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# Question 3: the hot goalie
# ----------------------------------------------------------------------------
def starter_prior(goalies):
    """Each team's starter (most games) and his goals saved above expected per 60 the season before."""
    g = goalies.copy()
    g["gsax60_goalie"] = (g["xGoals"] - g["goals"]) / (g["icetime"] / 3600)
    starters = g.sort_values("games_played", ascending=False).drop_duplicates(["season_id", "team_id"])
    prior = g[g["games_played"] >= 10][["season", "playerId", "gsax60_goalie"]].copy()
    prior["season"] = prior["season"] + 1
    out = starters[["season_id", "team_id", "season", "playerId", "name", "games_played"]].merge(
        prior, on=["season", "playerId"], how="left").rename(columns={"gsax60_goalie": "starter_prior_gsax60", "name": "starter"})
    return out[["season_id", "team_id", "starter", "games_played", "starter_prior_gsax60"]]


def playoff_table(wide, seasons, goalies):
    reg, po = wide[wide["playoffGame"] == 0], wide[wide["playoffGame"] == 1]
    full = team_stats(reg)
    last20 = team_stats(reg.sort_values("game_date").groupby(["season_id", "team_id"]).tail(20))[["season_id", "team_id", "gsax60"]].rename(
        columns={"gsax60": "last20_gsax60"})
    playoff = team_stats(po)[["season_id", "team_id", "games", "gsax60", "gsax", "sv", "xga60_5"]].rename(
        columns={"games": "playoff_games", "gsax60": "playoff_gsax60", "gsax": "playoff_gsax", "sv": "playoff_sv", "xga60_5": "playoff_xga60_5"})
    t = full.merge(last20, on=["season_id", "team_id"]).merge(playoff, on=["season_id", "team_id"]).merge(
        starter_prior(goalies), on=["season_id", "team_id"], how="left")
    t = t.merge(seasons[["season_id", "team_id", "season", "team", "outcome", "include_main"]], on=["season_id", "team_id"])
    t = t[(t["include_main"] == "Y") & t["outcome"].isin(ROUNDS)]
    t["rounds_won"] = t["outcome"].map(ROUNDS)
    return t


def playoff_predictors(t):
    """Pooled rank correlation (each season ranked on its own) between a regular-season signal and playoff GSAx/60."""
    signals = [("gsax60", "Regular-season goals saved above expected / 60"),
               ("last20_gsax60", "Same, last 20 regular-season games"),
               ("sv", "Regular-season raw save %"),
               ("starter_prior_gsax60", "Starter's goals saved above expected / 60, season before"),
               ("xga60_5", "Defense: expected goals against / 60 (5v5)")]
    rows = []
    y = t.groupby("season_id")["playoff_gsax60"].rank(pct=True)
    for key, name in signals:
        x = t.groupby("season_id")[key].rank(pct=True)
        ok = x.notna() & y.notna()
        rows.append({"Regular-season signal": name, "Playoff team-runs": int(ok.sum()),
                     "Link with playoff goaltending (rank correlation)": round(x[ok].rank().corr(y[ok].rank()), 2)})
    return pd.DataFrame(rows)


def champion_goaltending(t):
    t = t.copy()
    t["playoff_pct"] = t.groupby("season_id")["playoff_gsax60"].rank(pct=True) * 100
    t["reg_pct"] = t.groupby("season_id")["gsax60"].rank(pct=True) * 100
    rows = []
    for outcome in ["Won Cup", "Lost Cup final", "Lost conference final", "Lost round 2", "Lost round 1"]:
        g = t[t["outcome"] == outcome]
        rows.append({"How the playoffs ended": outcome, "Teams": len(g),
                     "Playoff goaltending (avg percentile among playoff teams)": round(g["playoff_pct"].mean(), 0),
                     "Regular-season goaltending (avg percentile among playoff teams)": round(g["reg_pct"].mean(), 0),
                     "Share with positive playoff GSAx": f"{(g['playoff_gsax'] > 0).mean():.0%}",
                     "Average playoff GSAx (goals)": round(g["playoff_gsax"].mean(), 1)})
    link = t["playoff_gsax60"].rank().corr(t["rounds_won"].rank())
    return pd.DataFrame(rows), float(link)


def hot_goalie_rule(t, min_games=7):
    """If a team's regular-season goaltending ranked top N among the playoff teams, how often was it top N in the playoffs?"""
    d = t[t["playoff_games"] >= min_games].copy()
    d["reg_rank"] = d.groupby("season_id")["gsax60"].rank(ascending=False)
    d["po_rank"] = d.groupby("season_id")["playoff_gsax60"].rank(ascending=False)
    d["n_teams"] = d.groupby("season_id")["gsax60"].transform("count")
    rows = []
    for n_top in (3, 5):
        flagged = d[d["reg_rank"] <= n_top]
        others = d[d["reg_rank"] > n_top]
        chance = float((n_top / flagged["n_teams"]).mean())
        hit = float((flagged["po_rank"] <= n_top).mean())
        rows.append({"Rule": f"Regular-season goaltending ranked top {n_top} (among playoff teams that played {min_games}+ games)",
                     "Team-runs that met it": len(flagged), "Of those, also top-N in the playoffs (count)": int((flagged["po_rank"] <= n_top).sum()),
                     "Rate": round(hit, 3), "Rate if it meant nothing (chance)": round(chance, 3),
                     "Lift": round(hit / chance, 2), "Rate for all other teams": round(float((others["po_rank"] <= n_top).mean()), 3)})
    return pd.DataFrame(rows), len(d)


def hot_goalies(t, n=15):
    d = t.copy()
    d["reg_rank"] = d.groupby("season_id")["gsax60"].rank(ascending=False).astype(int)          # among ALL playoff teams
    d["playoff_teams"] = d.groupby("season_id")["gsax60"].transform("count")
    d = d[d["playoff_games"] >= 8]
    top = d.sort_values("playoff_gsax", ascending=False).head(n)
    return pd.DataFrame({
        "Season": top["season"], "Team": top["team"], "Starter": top["starter"], "How the playoffs ended": top["outcome"],
        "Playoff games": top["playoff_games"].astype(int), "Playoff GSAx (goals)": top["playoff_gsax"].round(1),
        "Regular-season GSAx/60": top["gsax60"].round(2),
        "Regular-season rank among playoff teams": top["reg_rank"].astype(str) + " of " + top["playoff_teams"].astype(int).astype(str),
        "Starter's previous-season GSAx/60": top["starter_prior_gsax60"].round(2)})


# ----------------------------------------------------------------------------
# The current season and the goalies
# ----------------------------------------------------------------------------
def right_now(connection):
    ts = pd.read_sql_query(
        """SELECT s.team, s.team_id, s.situation, s.games_played, s.iceTime, s.xGoalsAgainst, s.goalsAgainst
           FROM mp_team_seasons s WHERE s.season = (SELECT MAX(season) FROM mp_team_seasons)
                                    AND s.situation IN ('all', '5on5')""", connection)
    a, f = ts[ts["situation"] == "all"], ts[ts["situation"] == "5on5"]
    w = a.merge(f, on=["team", "team_id"], suffixes=("", "_5"))
    d = pd.DataFrame({
        "team_id": w["team_id"], "Games played": w["games_played"],
        "xga60_5": w["xGoalsAgainst_5"] / (w["iceTime_5"] / 3600),
        "gsax60": (w["xGoalsAgainst"] - w["goalsAgainst"]) / (w["iceTime"] / 3600),
        "GSAx (goals)": (w["xGoalsAgainst"] - w["goalsAgainst"]).round(1), "Goals against": w["goalsAgainst"],
        "Expected goals against": w["xGoalsAgainst"].round(1)})
    d["season_id"] = int(connection.execute("SELECT MAX(season_id) FROM team_seasons").fetchone()[0])
    d = label_teams(d)
    names = pd.read_sql_query("SELECT DISTINCT team_id, team FROM team_seasons WHERE season_id = (SELECT MAX(season_id) FROM team_seasons)", connection)
    d = d.merge(names, on="team_id")
    current_labels = d[["season_id", "team_id", "label", "xga60_5"]].copy()
    d["Defense rank (1 = best)"] = d["xga60_5"].rank().astype(int)
    d["Goaltending rank (1 = best)"] = d["gsax60"].rank(ascending=False).astype(int)
    out = d.sort_values("gsax60", ascending=False)[["team", "Games played", "label", "Defense rank (1 = best)", "Goaltending rank (1 = best)",
                                                    "Goals against", "Expected goals against", "GSAx (goals)"]]
    return out.rename(columns={"team": "Team", "label": "Label"}).reset_index(drop=True), current_labels


def goalies_table(goalies, seasons, label_frame):
    """Goalies of the latest finished season and of the current one, with their team's defense."""
    out = []
    latest = int(goalies["season"].max())
    for season, min_games, tag in ((latest - 1, 25, "last season"), (latest, 8, "this season so far")):
        g = goalies[(goalies["season"] == season) & (goalies["games_played"] >= min_games)].copy()
        g["GSAx"] = g["xGoals"] - g["goals"]
        g["GSAx per 60"] = g["GSAx"] / (g["icetime"] / 3600)
        g["Save %"] = 1 - g["goals"] / g["ongoal"]
        g["season"] = tag
        out.append(g)
    g = pd.concat(out, ignore_index=True)
    g = g.merge(label_frame[["season_id", "team_id", "label", "xga60_5"]], on=["season_id", "team_id"], how="left")
    g = g.merge(seasons[["season_id", "team_id", "team"]], on=["season_id", "team_id"], how="left")
    cols = ["season", "name", "team", "games_played", "GSAx", "GSAx per 60", "Save %", "xga60_5", "label"]
    g = g[cols].rename(columns={"season": "Season", "name": "Goalie", "team": "Team", "games_played": "Games", "xga60_5": "Team's expected goals against / 60 (5v5)",
                                "label": "Team label"})
    for c in ("GSAx", "GSAx per 60", "Team's expected goals against / 60 (5v5)"):
        g[c] = g[c].round(2)
    g["Save %"] = g["Save %"].round(3)
    return g.sort_values(["Season", "GSAx"], ascending=[False, False]).reset_index(drop=True)


# ----------------------------------------------------------------------------
# Verdict
# ----------------------------------------------------------------------------
def verdict_lines(stability, split_shares, follow, gain, predictors, champs, champ_link, corr_def_goalie, n_seasons, rule):
    rep = stability.set_index("Statistic")["Repeats (before vs after Jan 1)"]
    r_def, r_gs, r_sv = rep.iloc[0], rep.iloc[2], rep.iloc[3]
    var_def, var_goal = split_shares
    bail = follow[follow["Label (as of Jan 1)"] == "Goalie bailing out the defense"].iloc[0]
    let = follow[follow["Label (as of Jan 1)"] == "Defense let down by the goalie"].iloc[0]
    both_good = follow[follow["Label (as of Jan 1)"] == "Both good"].iloc[0]
    gains = gain.set_index("Statistic (before Jan 1)")["Gain"]
    best_pred = predictors.sort_values("Link with playoff goaltending (rank correlation)", ascending=False).iloc[0]
    won = champs[champs["How the playoffs ended"] == "Won Cup"].iloc[0]
    best_link = best_pred["Link with playoff goaltending (rank correlation)"]
    if best_link >= 0.2:
        hot_text = (f"A regular-season signal does carry some information: '{best_pred['Regular-season signal']}' has a rank correlation of {best_link:.2f} "
                    f"with playoff goaltending (about 240 playoff runs; differences under about 0.13 are noise).")
    else:
        hot_text = (f"No regular-season signal clearly foresees a hot playoff goalie: every link is {best_link:.2f} or lower (goaltending results, the last 20 games, "
                    f"raw save %, the starter's previous season, defense quality), and about 0.13 is the noise level with about 240 playoff runs. "
                    f"Hot playoff goaltending looks mostly unforeseeable from the regular season.")
    r3 = rule.iloc[0]
    rule_text = (f" One more way to look: {int(r3.iloc[2])} of the {int(r3.iloc[1])} teams whose regular-season goaltending ranked top 3 among the playoff teams that played 7+ games "
                 f"also had top-3 goaltending in the playoffs ({r3['Rate']:.0%}, against {r3['Rate if it meant nothing (chance)']:.0%} by chance: a lift of {r3['Lift']:.1f}x), "
                 f"so being good helps a little, but most good regular-season goalies do not get hot and most hot ones were not the best in the regular season.")
    percentile = int(won["Playoff goaltending (avg percentile among playoff teams)"])
    suffix = "th" if 10 <= percentile % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(percentile % 10, "th")
    lines = [
        ("Question", "Is a team's goals against mostly its defense or its goalie? Which part repeats? Is a goalie 'bailing out' a weak defense (or being let down by a strong one) a lasting situation? And can anything foresee a hot playoff goalie?"),
        ("What drives goals against",
         f"Across teams, defense (the quality and quantity of shots allowed) accounts for about {var_def:.0%} of the differences in goals against per game, goaltending about {var_goal:.0%} "
         f"(the shares add to more than 100% because the two are slightly positively related, correlation {corr_def_goalie:+.2f}: teams that allow more danger get a little more help from their goalies)."),
        ("What repeats",
         f"Defense repeats from before Jan 1 to after (correlation {r_def:.2f} for expected goals against); goaltending results repeat much less "
         f"({r_gs:.2f} for goals saved above expected, {r_sv:.2f} for raw save %). So the defensive part of goals against is the stable part, and most of what looks like goalie skill over a half season is noise."),
        ("Bailed out vs let down (after Jan 1)",
         f"Teams labeled 'goalie bailing out the defense' earned {bail['= standings points over the rest of the season']:+.1f} standings points over the rest of the season relative to what their record predicted "
         f"(margin +/-{bail['Margin (+/-, 2 std. errors, in points)']:.1f}); teams 'let down by the goalie' {let['= standings points over the rest of the season']:+.1f} (+/-{let['Margin (+/-, 2 std. errors, in points)']:.1f}); "
         f"teams good at both {both_good['= standings points over the rest of the season']:+.1f}. A sign of regression to the mean, if the first is negative and the second positive."),
        ("Do the numbers improve the forecast?",
         "Gain in forecasting points % after Jan 1 beyond record and goal difference (R-squared): " +
         "; ".join(f"{name.split(':')[0]} - {name.split(': ', 1)[1]}: {value:+.3f}" for name, value in gains.items()) + ". Same hold-out test and noise level as the popular-stats check."),
        ("The hot playoff goalie",
         hot_text + rule_text + f" Champions' playoff goaltending averaged the {percentile}{suffix} percentile among playoff teams "
         f"(and {won['Share with positive playoff GSAx']} of champions had positive playoff GSAx), and playoff goaltending is linked to rounds won at {champ_link:.2f}: "
         f"a hot goalie matters in the playoffs, it just is not visible in advance."),
        ("Not covered yet", "The free files have no playoff data for individual goalies, and the team numbers blend starter and backup. Shot-level data (also on MoneyPuck's page) would let us do the goalie-by-goalie playoff version."),
    ]
    return pd.DataFrame(lines, columns=["", "Answer"])


# ----------------------------------------------------------------------------
def write_excel(sheets, notes):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 124
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
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 44)
        ws = writer.sheets["Verdict"]
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 130
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(bold=True)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws = writer.sheets["Goals against split"]
        ws.column_dimensions["A"].width = 52
        ws.column_dimensions["C"].width = 90


def main():
    sys.stdout.reconfigure(errors="replace")
    if not DB_FILE.exists():
        print("I can't find the database yet. Run: python scripts/build_database.py")
        return 1
    connection = sqlite3.connect(DB_FILE)
    tables = [r[0] for r in connection.execute("SELECT name FROM sqlite_master WHERE type='table'")]
    if "mp_team_games" not in tables:
        print("The MoneyPuck tables are not in the database yet. Run: python scripts/load_moneypuck.py")
        return 1
    wide, seasons, checkpoints, goalies = load(connection)
    games_nhl, seasons_nhl, checkpoints_nhl = ps.load(connection)
    games_nhl, main_seasons = ps.prepare_games(games_nhl, seasons_nhl)
    main_ids = set(games_nhl["season_id"].unique())
    wide = wide[wide["season_id"].isin(main_ids)]
    n_seasons = len(main_ids)

    # Split before / after the checkpoint, joined with the NHL results (record, goal difference)
    nhl_split = ps.split_tables(games_nhl, checkpoints_nhl)[SPLIT]
    mp_split = split_frames(wide, checkpoints, SPLIT)
    frame = nhl_split.merge(mp_split, on=["season_id", "team_id"])
    print(f"Using {n_seasons} main seasons from 2009-10 on; {len(frame)} team-seasons with both NHL and MoneyPuck data.")

    stability = stability_table(frame)
    full = team_stats(wide[wide["playoffGame"] == 0])
    split_df, var_def, var_goal = goals_against_split(full)
    corr_def_goalie = float(full["xga_pg"].corr(full["gsax_pg"]))
    follow = label_followup(frame)
    gain = forecast_gain(frame)

    t = playoff_table(wide, seasons, goalies)
    predictors = playoff_predictors(t)
    champs, champ_link = champion_goaltending(t)
    hot = hot_goalies(t)
    rule, rule_n = hot_goalie_rule(t)
    verdict = verdict_lines(stability, (var_def, var_goal), follow, gain, predictors, champs, champ_link, corr_def_goalie, n_seasons, rule)

    # Labels for every finished team-season (full season) and for the current season
    labeled = label_teams(full)
    labeled = labeled.merge(seasons[["season_id", "team_id", "season", "team", "outcome", "include_main"]], on=["season_id", "team_id"])
    label_table = labeled[labeled["include_main"] == "Y"]
    summary_rows = []
    for label in LABELS:
        g = label_table[label_table["label"] == label]
        playoffs = g[g["outcome"].isin(ROUNDS)]
        summary_rows.append({"Label (full season)": label, "Team-seasons": len(g), "Made the playoffs": f"{(g['outcome'] != 'Missed playoffs').mean():.0%}",
                             "Reached the conference finals": f"{g['outcome'].isin(['Lost conference final', 'Lost Cup final', 'Won Cup']).mean():.1%}",
                             "Won the Cup": int((g["outcome"] == "Won Cup").sum()), "Goals against / game": round(g["ga_pg"].mean(), 2),
                             "Expected goals against / game": round(g["xga_pg"].mean(), 2)})
    label_summary = pd.DataFrame(summary_rows)

    teams_out = label_table[["season", "team", "label", "outcome", "xga60_5", "hda60_5", "gsax60", "gsax", "sv", "ga_pg", "xga_pg"]].rename(columns={
        "season": "Season", "team": "Team", "label": "Label", "outcome": "How the season ended", "xga60_5": "Expected goals against /60 (5v5)",
        "hda60_5": "High-danger shots against /60 (5v5)", "gsax60": "GSAx /60", "gsax": "GSAx (goals)", "sv": "Raw save %", "ga_pg": "Goals against /game",
        "xga_pg": "Expected goals against /game"}).round(3).sort_values(["Season", "GSAx (goals)"], ascending=[True, False])

    sheets = {"Verdict": verdict, "What repeats": stability, "Goals against split": split_df,
              "Labels after Jan 1": follow, "Forecast gain": gain, "Labels over a season": label_summary,
              "Playoff signals": predictors, "Hot goalie rule": rule, "Playoff goaltending": champs, "Hot goalies": hot}
    notes = {}
    label_source = labeled[["season_id", "team_id", "label", "xga60_5"]]
    try:
        sheets["Right now"], current_labels = right_now(connection)
        label_source = pd.concat([label_source, current_labels], ignore_index=True)
        notes["Right now"] = "Current season so far: teams labeled on defense (expected goals against /60, 5v5) and goaltending (GSAx /60). Early in the season these are noisy."
    except Exception as problem:                                              # noqa: BLE001
        print(f"  (Right now list skipped: {problem})")
    sheets["Goalies"] = goalies_table(goalies, seasons, label_source)
    sheets["Team-seasons"] = teams_out
    connection.close()
    try:
        write_excel(sheets, notes)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 230)
    pd.set_option("display.max_colwidth", 140)
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nWhat repeats:"); print(stability.to_string(index=False))
    print("\nGoals against split:"); print(split_df.to_string(index=False))
    print("\nLabels after Jan 1:"); print(follow.to_string(index=False))
    print("\nForecast gain:"); print(gain.to_string(index=False))
    print("\nLabels over a season:"); print(label_summary.to_string(index=False))
    print("\nPlayoff signals:"); print(predictors.to_string(index=False))
    print("\nHot goalie rule:"); print(rule.to_string(index=False))
    print("\nPlayoff goaltending by result:"); print(champs.to_string(index=False)); print(f"  playoff GSAx/60 vs rounds won: {champ_link:.2f}")
    print("\nHot goalies:"); print(hot.to_string(index=False))
    if "Right now" in sheets:
        print("\nRight now (top/bottom):"); print(sheets["Right now"].head(6).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: MoneyPuck.com (non-commercial use) and (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""
build_contender_fit.py  -  contender-fit score v1 (NHL stats only) and an honest test of it
============================================================================================
WHAT IT DOES (in Excel terms):
    Builds a "contender-fit" score for every team at every checkpoint (Nov 1, Thanksgiving, Jan 1,
    trade deadline, end of season), using ONLY games played up to that date (no peeking), and then
    asks the one question that matters: does it pick the eventual conference finalists ("final four")
    better than just ranking teams by points %?

    The score is the model's estimate of a team's chance to be one of the final four, built from:
        record (points %), goal difference, Corsi (shot-attempt share), and your factor 5:
        the "power kill" (PP% + PK%), special-teams volume and net special-teams goals.
    Each version adds one group of stats, so you can see exactly what each idea adds.

    THE HONEST TEST: for each season, the score is built on the OTHER 14 seasons and then used to rank
    the teams of the season it has never seen (like hiding one column of answers, then checking).
    Success = the real final-four teams land higher in our list than in the points % list.

    It also writes:
      * Profile    how champions, finalists and everyone else looked at the end of the season,
      * Rule test  "a team that is top 5 / top 8 in K of these factors": how often is it the champion?
      * Right now  the current season's ranking, with a plain-English reason for each team.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py          (first, so the database is up to date)
    python scripts/build_contender_fit.py

WHAT YOU SHOULD SEE:
    A scorecard in the terminal and a new file:  outputs/contender_fit.xlsx
    (close it in Excel before running again). Start with its "Scorecard" sheet.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_popular_stats_check as ps        # re-use the stat calculations from the previous step

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "contender_fit.xlsx"

CHECKPOINTS = ["nov1", "thanksgiving", "jan1", "deadline", "season_end"]
PLAIN_CP = {"nov1": "Nov 1", "thanksgiving": "Thanksgiving", "jan1": "Jan 1",
            "deadline": "Trade deadline", "season_end": "End of season"}

# Ingredient names in plain English (z = how far above/below the league average this season, in steps)
FEATURE_NAME = {
    "pts_pct": "Record (points %)", "gd": "Goal difference", "corsi": "Shot-attempt share (Corsi)",
    "power_kill": "Power kill (PP% + PK%)", "st_volume": "Special-teams volume", "st_net": "Net special-teams goals",
}
# Each model adds one group of ingredients to the one before
MODELS = {
    "Model 1: record + goal difference": ["pts_pct", "gd"],
    "Model 2: + Corsi": ["pts_pct", "gd", "corsi"],
    "Model 3: + power kill (PP% + PK%)": ["pts_pct", "gd", "corsi", "power_kill"],
    "Model 4: + special-teams volume and net goals": ["pts_pct", "gd", "corsi", "power_kill", "st_volume", "st_net"],
}
# Simple rankings that need no fitting at all
BASELINES = {"Baseline: points %": "pts_pct", "Baseline: goal difference": "gd", "Baseline: Corsi alone": "corsi"}
MAIN_BASELINE = "Baseline: points %"

PROFILE_STATS = [  # (key, plain name, higher is better?)
    ("pts_pct", "Points %", True), ("gd", "Goal difference per game", True), ("pm5", "Plus/minus (team, 5v5)", True),
    ("corsi", "Corsi (shot-attempt share)", True), ("sog_share", "Shots-on-goal share", True),
    ("sh_pct", "Shooting %", True), ("sv_pct", "Save % (raw)", True), ("pdo", "PDO", True), ("hits", "Hits per game", True),
    ("pp_pct", "Power-play %", True), ("pk_pct", "Penalty-kill %", True),
    ("ppopp_pg", "Power plays per game", True), ("tsh_pg", "Times short-handed per game (fewer = higher)", False),
    ("st_net", "Net special-teams goals per game", True),
]
OUTCOME_ORDER = ["Won Cup", "Lost Cup final", "Lost conference final", "Lost round 2", "Lost round 1", "Missed playoffs"]

READ_ME = [
    "Contender fit v1: who is a real contender, and does our score beat 'just rank by points %'?",
    "",
    "WHAT THE SCORE IS",
    "  For each team at a checkpoint the model estimates the chance of being one of the four conference finalists,",
    "  using only games played up to that date. Ingredients (each measured against the league that season):",
    "    Record (points %), Goal difference, Shot-attempt share (Corsi), Power kill (PP% + PK%),",
    "    Special-teams volume (power plays drawn minus times short-handed), Net special-teams goals.",
    "  Four versions, each adding to the last: Model 1 (record + goal difference), 2 (+ Corsi), 3 (+ power kill),",
    "  4 (+ special-teams volume and net goals). Your other factors (slot shots, zone time, breakouts, goalie in",
    "  context) need shot-quality and play-by-play data we do not have yet: they are the next versions.",
    "",
    "THE HONEST TEST",
    "  Each season is scored by a model built on the OTHER seasons only (leave-one-season-out), then the teams are ranked.",
    "  We then look at where the four real conference finalists landed in that ranking.",
    "    AUC             1.0 = every finalist ranked above every non-finalist; 0.5 = no better than a coin flip.",
    "    Avg finalist rank   lower is better (a random list would average about 16).",
    "    In top 8        share of the real finalists the list put in its top 8.",
    "  The baselines need no model: just rank by points %, by goal difference, or by Corsi.",
    "  'Gain vs points %' is the difference in AUC; the +/- is about two standard errors across 15 seasons: a gain smaller",
    "  than its +/- is within noise.",
    "",
    "SEASONS",
    "  The 15 main seasons from 2009-10 on (five-on-five and Corsi data start there; 2019-20 and 2020-21 left out).",
    "  That is 60 conference finalists, 30 Cup finalists and 15 champions: small samples, so differences are fragile.",
    "  Checkpoint stats use games up to and including that date; 2012-13 uses share-of-season dates.",
    "",
    "SHEETS",
    "  Scorecard   every method at every checkpoint (START HERE).   Weights  what each model rewards.",
    "  Profile     how champions, finalists and the rest looked on each stat at the end of the season.",
    "  Rule test   'top 5 / top 8 in K of 4 factors': how often such teams win it all, against the base rate.",
    "  Right now   the current season's ranking from the best model, with the reasons.   Scores  every team-season.",
    "",
    "WHAT THIS CANNOT SAY YET",
    "  The score uses team stats only: no slot-shot share, no zone time, no shot-quality-adjusted goaltending,",
    "  no depth chart (top-15 center, elite D pair). A hot goalie in the playoffs is mostly unforeseeable noise.",
    "  The 'Right now' list is a probability-style ranking of final-four fit, not a prediction of the Cup winner.",
    "",
    "Data: (c) NHL.",
]


# ----------------------------------------------------------------------------
# Loading and building the checkpoint tables
# ----------------------------------------------------------------------------
def load(connection):
    games = pd.read_sql_query(
        """SELECT season_id, game_id, game_date, team_id, goalsFor, goalsAgainst, points,
                  goalsFor5On5, goalsAgainst5On5, shots5v5, satFor, satAgainst, hits,
                  ppOpportunities, powerPlayGoalsFor, timesShorthanded, ppGoalsAgainst,
                  shGoalsFor, shGoalsAgainst
           FROM team_games WHERE game_type = 2 AND season_id >= ?""", connection, params=[ps.FIRST_SEASON])
    seasons = pd.read_sql_query(
        """SELECT season_id, season, team_id, team, include_main, outcome, made_playoffs,
                  conf_finalist, cup_finalist, champion
           FROM team_seasons WHERE season_id >= ?""", connection, params=[ps.FIRST_SEASON])
    checkpoints = pd.read_sql_query(
        "SELECT season_id, checkpoint, date FROM checkpoints WHERE reached = 1", connection)
    return games, seasons, checkpoints


def summarize_all(games):
    """The popular-stats table plus the special-teams extras, one row per team per season."""
    base = ps.summarize(games)
    g = games.groupby(["season_id", "team_id"]).agg(
        gp=("game_id", "count"), ppopp=("ppOpportunities", "sum"), tsh=("timesShorthanded", "sum"),
        ppgf=("powerPlayGoalsFor", "sum"), ppga=("ppGoalsAgainst", "sum"),
        shgf=("shGoalsFor", "sum"), shga=("shGoalsAgainst", "sum")).reset_index()
    extra = g[["season_id", "team_id"]].copy()
    extra["ppopp_pg"] = g["ppopp"] / g["gp"]
    extra["tsh_pg"] = g["tsh"] / g["gp"]
    extra["st_net"] = (g["ppgf"] + g["shgf"] - g["ppga"] - g["shga"]) / g["gp"]
    return base.merge(extra, on=["season_id", "team_id"])


def add_features(table):
    """Turn raw stats into 'how far above/below this season's league average' ingredients."""
    t = table.copy()
    z = lambda col: ps.zscore_by_season(t, col)                       # noqa: E731
    t["z_pts_pct"], t["z_gd"], t["z_corsi"] = z("pts_pct"), z("gd"), z("corsi")
    t["z_st_net"] = z("st_net")
    t["_pk"] = z("pp_pct") + z("pk_pct")                                # power kill = PP% and PK% together
    t["_vol"] = z("ppopp_pg") - z("tsh_pg")                             # more power plays, fewer penalties
    t["z_power_kill"] = ps.zscore_by_season(t.assign(x=t["_pk"]), "x")
    t["z_st_volume"] = ps.zscore_by_season(t.assign(x=t["_vol"]), "x")
    t = t.drop(columns=["_pk", "_vol"])
    zcols = [c for c in t.columns if c.startswith("z_")]
    t[zcols] = t[zcols].fillna(0.0)                                     # a missing stat counts as average
    return t


def build_checkpoint_tables(games, seasons, checkpoints):
    """For each checkpoint: every team's ingredients using only games up to then, with how its season ended."""
    outcomes = seasons[["season_id", "team_id", "season", "team", "outcome", "made_playoffs",
                        "conf_finalist", "cup_finalist", "champion"]]
    tables = {}
    for cp in CHECKPOINTS:
        dates = checkpoints[checkpoints["checkpoint"] == cp][["season_id", "date"]]
        t = games.merge(dates, on="season_id")
        t = t[t["game_date"] <= t["date"]]
        tables[cp] = add_features(summarize_all(t)).merge(outcomes, on=["season_id", "team_id"])
    return tables


# ----------------------------------------------------------------------------
# The model (a simple, regularised logistic regression written out by hand)
# ----------------------------------------------------------------------------
def fit_logistic(x, y, l2=1.0, iterations=50):
    """
    Logistic regression: a weighted sum of the ingredients turned into a 0-100% chance.
    l2 gently shrinks the weights so a handful of finalists cannot push them to extremes.
    """
    xb = np.column_stack([np.ones(len(x)), x])
    penalty = np.eye(xb.shape[1]) * l2
    penalty[0, 0] = 0                                  # no shrinking of the intercept
    beta = np.zeros(xb.shape[1])
    for _ in range(iterations):
        p = 1 / (1 + np.exp(-xb @ beta))
        gradient = xb.T @ (y - p) - penalty @ beta
        hessian = -(xb.T * (p * (1 - p))) @ xb - penalty
        step = np.linalg.solve(hessian, gradient)
        beta = beta - step
        if np.abs(step).max() < 1e-8:
            break
    return beta


def predict_logistic(beta, x):
    return 1 / (1 + np.exp(-(np.column_stack([np.ones(len(x)), x]) @ beta)))


def calibrate_intercept(beta, x, total):
    """Nudge every team's chance up or down together so that the chances add up to `total` (4 final-four spots)."""
    low, high = -10.0, 10.0
    for _ in range(60):
        shift = (low + high) / 2
        adjusted = beta.copy()
        adjusted[0] += shift
        if predict_logistic(adjusted, x).sum() > total:
            high = shift
        else:
            low = shift
    adjusted = beta.copy()
    adjusted[0] += (low + high) / 2
    return adjusted


def held_out_scores(table, features):
    """Score each season with a model built on all the OTHER seasons. Returns a Series aligned to `table`."""
    cols = [f"z_{f}" for f in features]
    out = pd.Series(index=table.index, dtype=float)
    for season in sorted(table["season_id"].unique()):
        train, test = table[table["season_id"] != season], table[table["season_id"] == season]
        beta = fit_logistic(train[cols].to_numpy(), train["conf_finalist"].to_numpy(dtype=float))
        out.loc[test.index] = predict_logistic(beta, test[cols].to_numpy())
    return out


# ----------------------------------------------------------------------------
# Scoring a ranking: where did the real final-four teams land?
# ----------------------------------------------------------------------------
def auc(scores, positives):
    """Chance that a random finalist is scored above a random non-finalist (ties count half)."""
    ranks = scores.rank()
    n_pos = int(positives.sum())
    n_neg = len(scores) - n_pos
    return float((ranks[positives].sum() - n_pos * (n_pos + 1) / 2) / (n_pos * n_neg))


def per_season_metrics(table, score_col):
    rows = []
    for season, d in table.groupby("season_id"):
        rank = d[score_col].rank(ascending=False, method="average")
        f4 = d["conf_finalist"] == 1
        rows.append({"season_id": season, "auc": auc(d[score_col], f4),
                     "avg_rank_f4": float(rank[f4].mean()), "top8_f4": float((rank[f4] <= 8).mean()),
                     "champ_rank": float(rank[d["champion"] == 1].mean()),
                     "cup_rank": float(rank[d["cup_finalist"] == 1].mean())})
    return pd.DataFrame(rows).set_index("season_id")


def score_everything(tables):
    """Run every method at every checkpoint. Returns (summary rows, per-season metrics, scored tables)."""
    summary, per_season, scored = [], {}, {}
    for cp, table in tables.items():
        t = table.copy()
        for name, key in BASELINES.items():
            t[name] = t[f"z_{key}"]
        for name, features in MODELS.items():
            t[name] = held_out_scores(t, features)
        scored[cp] = t
        base = per_season_metrics(t, MAIN_BASELINE)
        for name in list(BASELINES) + list(MODELS):
            m = per_season_metrics(t, name)
            per_season[(cp, name)] = m
            diff = m["auc"] - base["auc"]
            summary.append({
                "Method": name, "Checkpoint": cp,
                "AUC": m["auc"].mean(), "Avg rank of final-four teams": m["avg_rank_f4"].mean(),
                "Final-four teams in top 8": m["top8_f4"].mean(), "Avg rank of champion": m["champ_rank"].mean(),
                "Avg rank of Cup finalists": m["cup_rank"].mean(),
                "Gain vs points % (AUC)": diff.mean(),
                "Gain margin (+/-, 2 std. errors)": 2 * diff.std(ddof=1) / np.sqrt(len(diff)),
                "Seasons better than points %": int((m["avg_rank_f4"] < base["avg_rank_f4"]).sum()),
                "Seasons worse than points %": int((m["avg_rank_f4"] > base["avg_rank_f4"]).sum()),
            })
    return pd.DataFrame(summary), per_season, scored


# ----------------------------------------------------------------------------
# Profile and rule test
# ----------------------------------------------------------------------------
def percentile_by_season(table, key, higher_better):
    p = table.groupby("season_id")[key].rank(pct=True)
    return p if higher_better else 1 - p + 1 / table.groupby("season_id")[key].transform("count")


def profile(end_table):
    t = end_table[end_table["outcome"].isin(OUTCOME_ORDER)].copy()
    rows = []
    for key, name, higher in PROFILE_STATS:
        t[f"pct_{key}"] = percentile_by_season(t, key, higher) * 100
    for outcome in OUTCOME_ORDER:
        part = t[t["outcome"] == outcome]
        row = {"How the season ended": outcome, "Team-seasons": len(part)}
        for key, name, _ in PROFILE_STATS:
            row[name] = round(part[f"pct_{key}"].mean(), 0)
        rows.append(row)
    row = {"How the season ended": "ALL teams", "Team-seasons": len(t)}
    for key, name, _ in PROFILE_STATS:
        row[name] = round(t[f"pct_{key}"].mean(), 0)
    rows.append(row)
    return pd.DataFrame(rows)


def rule_test(end_table):
    """'Top N in K of 4 factors': how often is such a team the champion / a finalist? (end of season)"""
    t = end_table[end_table["outcome"].isin(OUTCOME_ORDER)].copy()
    factor_cols = {"Record (points %)": "z_pts_pct", "Goal difference": "z_gd",
                   "Shot-attempt share (Corsi)": "z_corsi", "Power kill (PP% + PK%)": "z_power_kill"}
    base_champ = t["champion"].mean()
    base_f4 = t["conf_finalist"].mean()
    rows = []
    for top_n in (5, 8):
        strong = pd.DataFrame({name: t.groupby("season_id")[col].rank(ascending=False, method="min") <= top_n
                               for name, col in factor_cols.items()})
        t["k_strong"] = strong.sum(axis=1)
        for k in (2, 3, 4):
            part = t[t["k_strong"] >= k]
            n_champ = int(part["champion"].sum())
            rows.append({
                "Rule": f"Top {top_n} in at least {k} of 4 factors",
                "Teams meeting the rule": len(part),
                "Share of all teams": len(part) / len(t),
                "Champions among them": n_champ,
                "Share of champions that met it": n_champ / t["champion"].sum(),
                "Champion rate for such teams": n_champ / len(part) if len(part) else np.nan,
                "Lift over a random team": (n_champ / len(part)) / base_champ if len(part) else np.nan,
                "Final-four rate for such teams": part["conf_finalist"].mean() if len(part) else np.nan,
                "Final-four lift": part["conf_finalist"].mean() / base_f4 if len(part) else np.nan,
            })
    return pd.DataFrame(rows)


def calibration(scored, model):
    """When the (held-out) model said 'X% chance', how often was the team really a final-four team?"""
    edges = [0, 0.05, 0.10, 0.20, 0.35, 1.01]
    names = ["under 5%", "5-10%", "10-20%", "20-35%", "35% or more"]
    rows = []
    for cp, t in scored.items():
        bucket = pd.cut(t[model], bins=edges, labels=names, right=False)
        g = t.groupby(bucket, observed=False).agg(n=("conf_finalist", "size"), said=(model, "mean"),
                                                  actual=("conf_finalist", "mean"))
        for name, r in g.iterrows():
            rows.append({"Checkpoint": PLAIN_CP[cp], "The model said": name, "Team-seasons": int(r["n"]),
                         "Average chance it gave (%)": None if r["n"] == 0 else round(100 * r["said"], 1),
                         "How often they really were final four (%)": None if r["n"] == 0 else round(100 * r["actual"], 1)})
    return pd.DataFrame(rows)


def single_stats_text(headline, base_auc):
    corsi = headline.loc[headline["Method"] == "Baseline: Corsi alone", "AUC"].iloc[0]
    gd = headline.loc[headline["Method"] == "Baseline: goal difference", "AUC"].iloc[0]
    parts = [f"Ranking by Corsi alone gives AUC {corsi:.3f} against {base_auc:.3f} for points %: "
             + ("clearly worse, so controlling shot attempts does not by itself make a contender." if corsi < base_auc - 0.03
                else "not clearly different from the record.")]
    parts.append(f"Goal difference alone gives {gd:.3f}: " + ("no better than points % for picking the final four."
                                                               if gd <= base_auc + 0.01 else "better than points %."))
    return " ".join(parts)


def verdict_lines(headline, summary, best_model, n_seasons):
    """A few plain-English sentences, written from the numbers."""
    base = headline[headline["Method"] == MAIN_BASELINE].iloc[0]
    best = headline[headline["Method"] == best_model].iloc[0]
    rows = summary[summary["Method"] == best_model]
    margin = rows["Gain margin (+/-, 2 std. errors)"].mean()
    wins, losses = int(rows["Seasons better than points %"].sum()), int(rows["Seasons worse than points %"].sum())
    gain = best["AUC"] - base["AUC"]
    proven = gain > margin
    adds = {}
    order = list(MODELS)
    for i in range(1, len(order)):
        adds[order[i]] = (headline.loc[headline["Method"] == order[i], "AUC"].iloc[0]
                          - headline.loc[headline["Method"] == order[i - 1], "AUC"].iloc[0])
    early = summary[(summary["Method"] == best_model) & (summary["Checkpoint"] == "nov1")].iloc[0]
    late = summary[(summary["Method"] == best_model) & (summary["Checkpoint"] == "season_end")].iloc[0]
    random_rank = 16.0
    lines = [
        ("Question", "Does a contender-fit score built from NHL team stats pick the real conference finalists better than just ranking by points %?"),
        ("Short answer",
         ("Slightly, but not provably." if not proven else "Yes, by a margin that is bigger than the noise.") +
         f" Best version: {best_model}. Averaged over five checkpoints its AUC is {best['AUC']:.3f} against {base['AUC']:.3f} for points % "
         f"(a gain of {gain:+.3f}; the noise margin is about +/-{margin:.3f})."),
        ("Season by season", f"It ranked the real final four better than points % did in {wins} season-checkpoints and worse in {losses} "
                             f"(out of {n_seasons * len(CHECKPOINTS)}; the rest tied)."),
        ("When it helps", f"Early: at Nov 1 the gain is {early['Gain vs points % (AUC)']:+.3f} (+/-{early['Gain margin (+/-, 2 std. errors)']:.3f}). "
                          f"By the end of the season it is {late['Gain vs points % (AUC)']:+.3f} (+/-{late['Gain margin (+/-, 2 std. errors)']:.3f}): "
                          f"a plausible reason is that a long record already reflects what Corsi and special teams add."),
        ("What each idea adds",
         "; ".join(f"{name.split(':')[1].strip()}: {d:+.3f}" for name, d in adds.items()) +
         f" (change in average AUC from the previous version; compare with the noise margin of about +/-{margin:.3f})."),
        ("How predictable is it at all?",
         f"At the end of the regular season the real final-four teams average rank {summary[(summary['Method'] == MAIN_BASELINE) & (summary['Checkpoint'] == 'season_end')].iloc[0]['Avg rank of final-four teams']:.1f} "
         f"by points % (a random list would average about {random_rank:.0f}); at Nov 1 it is "
         f"{summary[(summary['Method'] == MAIN_BASELINE) & (summary['Checkpoint'] == 'nov1')].iloc[0]['Avg rank of final-four teams']:.1f}. "
         f"Many final-four teams are simply not obvious from the standings: that is the nature of the playoffs, not a flaw in the stats."),
        ("Single stats", single_stats_text(headline, base["AUC"])),
        ("Not tested yet", "Slot / high-danger shot share, zone time, breakouts, goaltending in context, depth (top-15 center, elite D pair). "
                           "Those need MoneyPuck and play-by-play data and are the next versions."),
    ]
    return pd.DataFrame(lines, columns=["", "Answer"])


# ----------------------------------------------------------------------------
# Weights and the live view
# ----------------------------------------------------------------------------
def weights_table(tables):
    rows = []
    for name, features in MODELS.items():
        cols = [f"z_{f}" for f in features]
        for cp, t in tables.items():
            beta = fit_logistic(t[cols].to_numpy(), t["conf_finalist"].to_numpy(dtype=float))
            for f, b in zip(features, beta[1:]):
                rows.append({"Model": name, "Checkpoint": PLAIN_CP[cp], "Ingredient": FEATURE_NAME[f], "Weight": round(b, 2)})
    return pd.DataFrame(rows).pivot_table(index=["Model", "Ingredient"], columns="Checkpoint", values="Weight",
                                         sort=False)[[PLAIN_CP[c] for c in CHECKPOINTS]].reset_index()


def right_now(connection, tables, games_all, seasons_all, best_model):
    """Rank the CURRENT season's teams with the best model (trained on every finished season)."""
    current = int(games_all["season_id"].max())
    cur_games = games_all[games_all["season_id"] == current]
    fake_main = seasons_all[seasons_all["season_id"] == current].assign(include_main="Y", outcome="in progress")
    cur_games, _ = ps.prepare_games(cur_games, fake_main)
    now = add_features(summarize_all(cur_games))
    names = seasons_all[seasons_all["season_id"] == current][["team_id", "team"]]
    now = now.merge(names, on="team_id")

    features = MODELS[best_model]
    cols = [f"z_{f}" for f in features]
    # Use the model built for the checkpoint closest to 'right now' (by games played): Nov 1 is about 11
    # games in, Thanksgiving about 21, Jan 1 about 38, the deadline about 63.
    median_gp = int(now["gp"].median())
    checkpoint = min(tables, key=lambda cp: abs(float(tables[cp]["gp"].median()) - median_gp))
    train = tables[checkpoint]
    beta = fit_logistic(train[cols].to_numpy(), train["conf_finalist"].to_numpy(dtype=float))
    beta = calibrate_intercept(beta, now[cols].to_numpy(), total=4.0)     # four final-four spots to hand out
    now["chance"] = predict_logistic(beta, now[cols].to_numpy())
    contribution = pd.DataFrame({f: beta[1 + i] * now[f"z_{f}"] for i, f in enumerate(features)})

    # Record and goal difference move together, so the model gives them opposite-looking weights.
    # For the explanation, treat them as ONE ingredient: "Results".
    grouped = {}
    for f in features:
        key = "results" if f in ("pts_pct", "gd") else f
        grouped[key] = grouped.get(key, 0) + contribution[f]
    contribution = pd.DataFrame(grouped)
    label = {**FEATURE_NAME, "results": "Results (record and goal difference)"}

    def reason(i):
        c = contribution.loc[i].sort_values(ascending=False)
        up = [label[f] for f, v in c.items() if v >= 0.15][:2]
        down = [label[f] for f, v in c.iloc[::-1].items() if v <= -0.15][:2]
        parts = []
        if up:
            parts.append("Helped by: " + ", ".join(up))
        if down:
            parts.append("Held back by: " + ", ".join(down))
        return ". ".join(parts) if parts else "Close to average on everything the model looks at"

    now["Why"] = [reason(i) for i in now.index]
    now = now.sort_values("chance", ascending=False).reset_index(drop=True)
    now.insert(0, "Rank", np.arange(1, len(now) + 1))
    out = pd.DataFrame({
        "Rank": now["Rank"], "Team": now["team"], "Games played": now["gp"],
        "Points %": now["pts_pct"].round(3), "Goal difference / game": now["gd"].round(2),
        "Corsi": (now["corsi"] * 100).round(1), "PP%": (now["pp_pct"] * 100).round(1), "PK%": (now["pk_pct"] * 100).round(1),
        "Final-four chance (model)": (now["chance"] * 100).round(1), "Why": now["Why"]})
    out["Rank by points %"] = now["pts_pct"].rank(ascending=False, method="min").astype(int).to_numpy()
    return out, current, median_gp, checkpoint, str(cur_games["game_date"].max())


# ----------------------------------------------------------------------------
def write_excel(sheets, notes):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 122
        ws["A1"].font = Font(bold=True)
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False, startrow=1 if name in notes else 0)
            ws = writer.sheets[name]
            header_row = 2 if name in notes else 1
            if name in notes:
                ws.cell(row=1, column=1, value=notes[name]).font = Font(bold=True)
            ws.freeze_panes = ws.cell(row=header_row + 1, column=2)
            for cell in ws[header_row]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column in ws.columns:
                values = [c.value for c in column[header_row - 1:300] if c.value is not None and not isinstance(c.value, str) or
                          (isinstance(c.value, str) and len(c.value) < 80)]
                longest = max((len(str(v)) for v in values), default=8)
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 46)
        if "Right now" in sheets:
            ws = writer.sheets["Right now"]
            ws.column_dimensions["J"].width = 70
            for row in ws.iter_rows(min_row=3):
                row[9].alignment = Alignment(wrap_text=True, vertical="top")
        if "Verdict" in sheets:
            ws = writer.sheets["Verdict"]
            ws.column_dimensions["A"].width = 26
            ws.column_dimensions["B"].width = 130
            for row in ws.iter_rows(min_row=2):
                row[0].font = Font(bold=True)
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")


def main():
    sys.stdout.reconfigure(errors="replace")
    if not ps.DB_FILE.exists():
        print("I can't find the database yet. Run this first:")
        print("    python scripts/build_database.py")
        return 1
    connection = sqlite3.connect(ps.DB_FILE)
    games_raw, seasons, checkpoints = load(connection)
    games, main_seasons = ps.prepare_games(games_raw, seasons)
    tables = build_checkpoint_tables(games, seasons, checkpoints)
    n_seasons = games["season_id"].nunique()
    n_f4 = int(tables["season_end"]["conf_finalist"].sum())
    print(f"Using {n_seasons} main seasons from 2009-10 on ({n_f4} conference finalists, "
          f"{int(tables['season_end']['champion'].sum())} champions).")

    summary, per_season, scored = score_everything(tables)

    # Headline: each method averaged over the five checkpoints
    headline = (summary.groupby("Method", sort=False)
                .agg(AUC=("AUC", "mean"), rank_f4=("Avg rank of final-four teams", "mean"),
                     top8=("Final-four teams in top 8", "mean"), gain=("Gain vs points % (AUC)", "mean"))
                .reset_index())
    best_model = max(MODELS, key=lambda m: headline.loc[headline["Method"] == m, "AUC"].iloc[0])
    base_auc = headline.loc[headline["Method"] == MAIN_BASELINE, "AUC"].iloc[0]
    best_row = summary[summary["Method"] == best_model]
    wins = int(best_row["Seasons better than points %"].sum())
    losses = int(best_row["Seasons worse than points %"].sum())
    print(f"\nAverage over the five checkpoints:")
    print(headline.round(3).to_string(index=False))
    print(f"\nBest model: {best_model}. Vs points %: better in {wins} season-checkpoints, worse in {losses} "
          f"(out of {n_seasons * len(CHECKPOINTS)}).")

    right, current, median_gp, used_checkpoint, through = right_now(connection, tables, games_raw, seasons, best_model)
    connection.close()

    cp_table = summary.copy()
    cp_table["Checkpoint"] = cp_table["Checkpoint"].map(PLAIN_CP)
    cp_table["order"] = cp_table["Method"].map({m: i for i, m in enumerate(list(BASELINES) + list(MODELS))})
    cp_table["cp_order"] = cp_table["Checkpoint"].map({PLAIN_CP[c]: i for i, c in enumerate(CHECKPOINTS)})
    cp_table = cp_table.sort_values(["cp_order", "order"]).drop(columns=["order", "cp_order"]).round(3)
    head_out = headline.rename(columns={"rank_f4": "Avg rank of final-four teams", "top8": "Final-four teams in top 8",
                                        "gain": "Gain vs points % (AUC)"}).round(3)
    scorecard = pd.concat([head_out.assign(Checkpoint="ALL FIVE (average)"), cp_table], ignore_index=True)
    scorecard = scorecard[["Method", "Checkpoint"] + [c for c in scorecard.columns if c not in ("Method", "Checkpoint")]]

    end = scored["season_end"]
    score_cols = list(BASELINES) + list(MODELS)
    scores_out = pd.concat([
        t[["season", "team", "outcome"] + score_cols].assign(Checkpoint=PLAIN_CP[cp]) for cp, t in scored.items()
    ], ignore_index=True)
    scores_out[score_cols] = scores_out[score_cols].round(4)
    scores_out = scores_out[["season", "Checkpoint", "team", "outcome"] + score_cols]

    sheets = {
        "Verdict": verdict_lines(headline, summary, best_model, n_seasons),
        "Scorecard": scorecard,
        "Calibration": calibration(scored, best_model),
        "Weights": weights_table(tables),
        "Profile": profile(end),
        "Rule test": rule_test(end).round(3),
        "Right now": right,
        "Scores": scores_out,
    }
    notes = {"Right now": (f"Season {str(current)[:4]}-{str(current)[6:]}: games through {through} (median {median_gp} played). Scored with "
                           f"'{best_model}', built on the {PLAIN_CP[used_checkpoint]} pattern of 15 finished seasons (the checkpoint "
                           f"closest to now), with chances adjusted to add up to the four final-four spots. "
                           f"A rough guide to final-four fit, not a Cup prediction."),
             "Weights": ("What each model rewards, in steps above/below the league average. Record and goal difference move together, "
                         "so their weights partly cancel (one can look negative): read the two as a pair."),
             "Calibration": (f"Honesty check for '{best_model}': when it said X%, how often was the team really one of the four? "
                             f"The two columns should be close.")}
    try:
        write_excel(sheets, notes)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 220)
    print("\nScorecard at each checkpoint (AUC | avg rank of final-four teams | gain vs points % +/- margin):")
    show = cp_table[["Checkpoint", "Method", "AUC", "Avg rank of final-four teams", "Gain vs points % (AUC)",
                     "Gain margin (+/-, 2 std. errors)", "Seasons better than points %"]]
    print(show.to_string(index=False))
    print("\nRight now (top 10):")
    print(right[["Rank", "Team", "Games played", "Points %", "Final-four chance (model)", "Rank by points %"]].head(10).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

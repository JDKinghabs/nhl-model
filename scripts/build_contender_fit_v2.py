"""
build_contender_fit_v2.py  -  contender fit v2: does SHOT QUALITY and GOALTENDING IN CONTEXT beat version 1?
===========================================================================================================
WHAT IT DOES (in Excel terms):
    Version 1 (build_contender_fit.py) scored every team at every checkpoint from NHL stats only: record, goal
    difference, Corsi (shot-attempt share) and the power kill. It was only slightly better than ranking by points %.

    Version 2 adds the MoneyPuck shot-quality numbers, all computed from games up to each checkpoint date:
        expected-goals share (5v5)      who gets the better shots, not just more shots
        high-danger shot share (5v5)    your factor 1: the share of the dangerous (slot-area) shots a team takes
        goaltending in context          your factor 4: goals saved above expected, per 60 minutes
    Each version adds one idea to the last, and each is tested the same honest way as before: built on 14 seasons,
    used on the 15th it has never seen, to see where the REAL conference finalists land in the ranking.

    Then it compares: v2 against v1, against ranking by points %, and each of your ideas on its own.
    It also shows how champions looked on the shot-quality numbers, tests "top 5 in most factors" again with
    the new factors, and ranks today's teams.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py
    python scripts/load_moneypuck.py
    python scripts/build_contender_fit_v2.py

WHAT YOU SHOULD SEE:
    A scorecard in the terminal and a new file:  outputs/contender_fit_v2.xlsx   (close it in Excel first).
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

import build_contender_fit as cf          # v1: checkpoint tables, logistic model, test helpers
import build_popular_stats_check as ps    # shared stat helpers

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "contender_fit_v2.xlsx"

CHECKPOINTS = cf.CHECKPOINTS
PLAIN_CP = cf.PLAIN_CP
OUTCOME_ORDER = cf.OUTCOME_ORDER

FEATURE_NAME = {**cf.FEATURE_NAME,
                "xg_share": "Expected-goals share (5v5)", "hd_share": "High-danger shot share (5v5)",
                "gsax": "Goaltending in context (goals saved above expected)"}
V1 = ["pts_pct", "gd", "corsi", "power_kill"]
V1_NAME = "v1: record + goal difference + Corsi + power kill"
MODELS = {
    V1_NAME: V1,
    "v2a: v1 + expected-goals share": V1 + ["xg_share"],
    "v2b: v1 + high-danger shot share (your factor 1)": V1 + ["hd_share"],
    "v2c: v1 + both shot-quality shares": V1 + ["xg_share", "hd_share"],
    "v2d: shot quality instead of Corsi": ["pts_pct", "gd", "power_kill", "xg_share", "hd_share"],
    "v2e: v2c + goaltending in context (your factor 4)": V1 + ["xg_share", "hd_share", "gsax"],
}
BASELINES = {"Baseline: points %": "pts_pct", "Baseline: Corsi alone": "corsi",
             "Baseline: expected-goals share alone": "xg_share", "Baseline: high-danger share alone": "hd_share"}
MAIN_BASELINE = "Baseline: points %"
# (new model, what it is compared with, plain description of the idea tested)
COMPARISONS = [
    ("v2a: v1 + expected-goals share", V1_NAME, "Shot quality (expected goals) added to v1"),
    ("v2b: v1 + high-danger shot share (your factor 1)", V1_NAME, "Your factor 1: high-danger shot share added to v1"),
    ("v2c: v1 + both shot-quality shares", V1_NAME, "Both shot-quality shares added to v1"),
    ("v2d: shot quality instead of Corsi", V1_NAME, "Shot QUALITY in place of shot QUANTITY (Corsi)"),
    ("v2e: v2c + goaltending in context (your factor 4)", "v2c: v1 + both shot-quality shares", "Your factor 4: goaltending in context added to v2c"),
]

PROFILE_STATS = [  # (column, plain name, higher is better?)
    ("pts_pct", "Points %", True), ("gd", "Goal difference per game", True), ("corsi", "Corsi (shot-attempt share)", True),
    ("xg_share", "Expected-goals share (5v5)", True), ("hd_share", "High-danger shot share (5v5)", True),
    ("xga60_5", "Expected goals against /60 (5v5)", False), ("gsax60", "Goaltending (goals saved above expected /60)", True),
    ("sv_pct", "Raw save % (5v5)", True), ("pdo", "PDO", True), ("pp_pct", "Power-play %", True), ("pk_pct", "Penalty-kill %", True),
]

READ_ME = [
    "Contender fit v2: does shot quality and goaltending in context beat version 1?",
    "",
    "WHAT IS NEW (MoneyPuck shot-quality data, computed from games up to each checkpoint date, regular season only)",
    "  Expected-goals share (5v5)    expected goals for / (for + against): who gets the better shots, not just more.",
    "  High-danger shot share (5v5)  high-danger (slot-area) shots for / (for + against). Your factor 1.",
    "  Goaltending in context        goals saved above expected per 60 minutes, all situations. Your factor 4.",
    "",
    "THE MODELS (each adds to the one before; v1 is the best model from the first version)",
] + [f"  {name}: {', '.join(FEATURE_NAME[f] for f in feats)}" for name, feats in MODELS.items()] + [
    "",
    "THE HONEST TEST (same as version 1)",
    "  Each season is scored by a model built on the OTHER 14 seasons (leave-one-season-out), then teams are ranked and we look at where the four",
    "  real conference finalists landed. AUC: 1.0 = every finalist ranked above every non-finalist; 0.5 = a coin flip. 'Avg finalist rank': lower is",
    "  better (a random list averages about 16). Margins are about two standard errors across 15 seasons: a gain smaller than its margin is within noise.",
    "  Checkpoints: Nov 1, Thanksgiving, Jan 1, trade deadline, end of the regular season (only games up to that date).",
    "",
    "HOW THE VERDICT WORDS ARE CHOSEN",
    "  'Clear improvement' = average gain bigger than its margin. 'Leans positive' = positive but within the margin. 'No help' = zero or negative.",
    "",
    "WHAT THIS CANNOT SAY",
    "  15 seasons, 60 conference finalists: small samples. MoneyPuck's expected-goals model is one reasonable model, not the truth. Team-level only:",
    "  no depth (top-15 center, elite D pair), no injuries, and no breakout or zone-time numbers yet (those need the play-by-play pull).",
    "",
    "Data: MoneyPuck.com (non-commercial use) and (c) NHL.",
]


# ----------------------------------------------------------------------------
# MoneyPuck numbers as of each checkpoint
# ----------------------------------------------------------------------------
def load_mp_games(connection, first_season):
    games = pd.read_sql_query(
        """SELECT season_id, team_id, game_id, game_date, playoffGame, situation, iceTime, xGoalsFor, xGoalsAgainst,
                  goalsFor, goalsAgainst, shotsOnGoalAgainst, highDangerShotsFor, highDangerShotsAgainst
           FROM mp_team_games WHERE situation IN ('all', '5on5') AND season_id >= ? AND playoffGame = 0""",
        connection, params=[first_season])
    keys = ["season_id", "team_id", "game_id", "game_date", "playoffGame"]
    a = games[games["situation"] == "all"].drop(columns="situation")
    f = games[games["situation"] == "5on5"].drop(columns="situation")
    return a.merge(f, on=keys, suffixes=("", "_5"))


def mp_stats(games):
    """Shot-quality and goaltending numbers for every team-season from whichever games are passed in."""
    g = games.groupby(["season_id", "team_id"]).agg(
        ice=("iceTime", "sum"), xga=("xGoalsAgainst", "sum"), ga=("goalsAgainst", "sum"), ice5=("iceTime_5", "sum"),
        xgf5=("xGoalsFor_5", "sum"), xga5=("xGoalsAgainst_5", "sum"), hdf5=("highDangerShotsFor_5", "sum"),
        hda5=("highDangerShotsAgainst_5", "sum"), gf5=("goalsFor_5", "sum"), ga5=("goalsAgainst_5", "sum"),
        soga5=("shotsOnGoalAgainst_5", "sum")).reset_index()
    out = g[["season_id", "team_id"]].copy()
    out["xg_share"] = g["xgf5"] / (g["xgf5"] + g["xga5"])
    out["hd_share"] = g["hdf5"] / (g["hdf5"] + g["hda5"])
    out["xga60_5"] = g["xga5"] / (g["ice5"] / 3600)
    out["gsax60"] = (g["xga"] - g["ga"]) / (g["ice"] / 3600)
    return out.replace([np.inf, -np.inf], np.nan)


def add_mp_features(table):
    """Ingredients in 'steps above/below the league average this season' (positive = good)."""
    t = table.copy()
    t["z_xg_share"] = ps.zscore_by_season(t, "xg_share")
    t["z_hd_share"] = ps.zscore_by_season(t, "hd_share")
    t["z_gsax"] = ps.zscore_by_season(t, "gsax60")
    for col in ("z_xg_share", "z_hd_share", "z_gsax"):
        t[col] = t[col].fillna(0.0)
    return t


def build_v2_tables(tables, mp_wide, checkpoints):
    """Add the MoneyPuck ingredients to each checkpoint table from v1."""
    out = {}
    for cp, t in tables.items():
        dates = checkpoints[checkpoints["checkpoint"] == cp][["season_id", "date"]]
        w = mp_wide.merge(dates, on="season_id")
        w = w[w["game_date"] <= w["date"]]
        merged = t.merge(mp_stats(w), on=["season_id", "team_id"], how="left")
        out[cp] = add_mp_features(merged)
    return out


# ----------------------------------------------------------------------------
# Scoring everything
# ----------------------------------------------------------------------------
def evaluate(tables):
    """Run every ranking at every checkpoint. Returns (summary rows, per-season metrics, scored tables)."""
    summary, per_season, scored = [], {}, {}
    for cp, table in tables.items():
        t = table.copy()
        for name, key in BASELINES.items():
            t[name] = t[f"z_{key}"]
        for name, features in MODELS.items():
            t[name] = cf.held_out_scores(t, features)
        scored[cp] = t
        for name in list(BASELINES) + list(MODELS):
            per_season[(cp, name)] = cf.per_season_metrics(t, name)
        for name in list(BASELINES) + list(MODELS):
            m = per_season[(cp, name)]
            vs_base = m["auc"] - per_season[(cp, MAIN_BASELINE)]["auc"]
            vs_v1 = m["auc"] - per_season[(cp, V1_NAME)]["auc"]
            summary.append({
                "Method": name, "Checkpoint": cp, "AUC": m["auc"].mean(), "Avg rank of final-four teams": m["avg_rank_f4"].mean(),
                "Final-four teams in top 8": m["top8_f4"].mean(), "Avg rank of champion": m["champ_rank"].mean(),
                "Avg rank of Cup finalists": m["cup_rank"].mean(),
                "Gain vs points % (AUC)": vs_base.mean(), "Margin vs points % (+/-)": 2 * vs_base.std(ddof=1) / np.sqrt(len(vs_base)),
                "Gain vs v1 (AUC)": vs_v1.mean(), "Margin vs v1 (+/-)": 2 * vs_v1.std(ddof=1) / np.sqrt(len(vs_v1)),
                "Seasons better than v1": int((m["avg_rank_f4"] < per_season[(cp, V1_NAME)]["avg_rank_f4"]).sum()),
                "Seasons worse than v1": int((m["avg_rank_f4"] > per_season[(cp, V1_NAME)]["avg_rank_f4"]).sum()),
            })
    return pd.DataFrame(summary), per_season, scored


def increments(per_season):
    """What each idea adds, against the version it builds on (average over checkpoints, and early vs late)."""
    rows = []
    early, late = CHECKPOINTS[:3], CHECKPOINTS[3:]
    for new, old, idea in COMPARISONS:
        diffs = {cp: per_season[(cp, new)]["auc"] - per_season[(cp, old)]["auc"] for cp in CHECKPOINTS}
        gain = np.mean([d.mean() for d in diffs.values()])
        margin = np.mean([2 * d.std(ddof=1) / np.sqrt(len(d)) for d in diffs.values()])
        gain_early = np.mean([diffs[cp].mean() for cp in early])
        gain_late = np.mean([diffs[cp].mean() for cp in late])
        better = sum(int((per_season[(cp, new)]["avg_rank_f4"] < per_season[(cp, old)]["avg_rank_f4"]).sum()) for cp in CHECKPOINTS)
        worse = sum(int((per_season[(cp, new)]["avg_rank_f4"] > per_season[(cp, old)]["avg_rank_f4"]).sum()) for cp in CHECKPOINTS)
        word = "Clear improvement" if gain > margin else "Leans positive (within noise)" if gain > 0 else "No help"
        rows.append({"Idea tested": idea, "Compared with": old, "Average gain (AUC)": gain, "Margin (+/-)": margin,
                     "Early checkpoints (Nov 1 to Jan 1)": gain_early, "Late checkpoints (deadline, end)": gain_late,
                     "Season-checkpoints better": better, "Worse": worse, "Verdict": word})
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# Profile and rule test with the new factors
# ----------------------------------------------------------------------------
def profile_v2(end_table):
    t = end_table[end_table["outcome"].isin(OUTCOME_ORDER)].copy()
    for key, _, higher in PROFILE_STATS:
        p = t.groupby("season_id")[key].rank(pct=True)
        t[f"pct_{key}"] = (p if higher else 1 - p + 1 / t.groupby("season_id")[key].transform("count")) * 100
    rows = []
    for outcome in OUTCOME_ORDER + ["ALL teams"]:
        part = t if outcome == "ALL teams" else t[t["outcome"] == outcome]
        row = {"How the season ended": outcome, "Team-seasons": len(part)}
        for key, name, _ in PROFILE_STATS:
            row[name] = round(part[f"pct_{key}"].mean(), 0)
        rows.append(row)
    return pd.DataFrame(rows)


def rule_test_v2(end_table):
    """'Top N in at least K of the factors': how often is such a team the champion / a finalist? (end of season)"""
    t = end_table[end_table["outcome"].isin(OUTCOME_ORDER)].copy()
    factors = {"Record (points %)": "z_pts_pct", "Goal difference": "z_gd", "Expected-goals share": "z_xg_share",
               "High-danger shot share": "z_hd_share", "Power kill (PP% + PK%)": "z_power_kill"}
    base_champ, base_f4 = t["champion"].mean(), t["conf_finalist"].mean()
    rows = []
    for top_n in (5, 8):
        strong = pd.DataFrame({name: t.groupby("season_id")[col].rank(ascending=False, method="min") <= top_n
                               for name, col in factors.items()})
        t["k"] = strong.sum(axis=1)
        for k in (3, 4, 5):
            part = t[t["k"] >= k]
            n_champ = int(part["champion"].sum())
            rows.append({
                "Rule": f"Top {top_n} in at least {k} of 5 factors", "Teams meeting the rule": len(part),
                "Share of all teams": len(part) / len(t), "Champions among them": n_champ,
                "Share of champions that met it": n_champ / t["champion"].sum(),
                "Champion rate for such teams": n_champ / len(part) if len(part) else np.nan,
                "Lift over a random team": (n_champ / len(part)) / base_champ if len(part) else np.nan,
                "Final-four rate for such teams": part["conf_finalist"].mean() if len(part) else np.nan,
                "Final-four lift": part["conf_finalist"].mean() / base_f4 if len(part) else np.nan})
    return pd.DataFrame(rows)


def weights_v2(tables, names):
    rows = []
    for name in names:
        features = MODELS[name]
        cols = [f"z_{f}" for f in features]
        for cp, t in tables.items():
            beta = cf.fit_logistic(t[cols].to_numpy(), t["conf_finalist"].to_numpy(dtype=float))
            for f, b in zip(features, beta[1:]):
                rows.append({"Model": name, "Checkpoint": PLAIN_CP[cp], "Ingredient": FEATURE_NAME[f], "Weight": round(b, 2)})
    return pd.DataFrame(rows).pivot_table(index=["Model", "Ingredient"], columns="Checkpoint", values="Weight", sort=False)[
        [PLAIN_CP[c] for c in CHECKPOINTS]].reset_index()


# ----------------------------------------------------------------------------
# Right now
# ----------------------------------------------------------------------------
def right_now(tables, games_raw, seasons, mp_wide, best_model):
    current = int(games_raw["season_id"].max())
    cur = games_raw[games_raw["season_id"] == current]
    fake_main = seasons[seasons["season_id"] == current].assign(include_main="Y", outcome="in progress")
    cur, _ = ps.prepare_games(cur, fake_main)
    through = str(cur["game_date"].max())
    now = cf.add_features(cf.summarize_all(cur))
    mp_cur = mp_wide[(mp_wide["season_id"] == current) & (mp_wide["game_date"] <= through)]
    now = add_mp_features(now.merge(mp_stats(mp_cur), on="team_id", how="left", suffixes=("", "_mp")).assign(season_id=current)) \
        if "season_id" in now.columns else None
    now = now.merge(seasons[seasons["season_id"] == current][["team_id", "team"]], on="team_id")

    features = MODELS[best_model]
    cols = [f"z_{f}" for f in features]
    median_gp = int(now["gp"].median())
    checkpoint = min(tables, key=lambda cp: abs(float(tables[cp]["gp"].median()) - median_gp))
    train = tables[checkpoint]
    beta = cf.fit_logistic(train[cols].to_numpy(), train["conf_finalist"].to_numpy(dtype=float))
    beta = cf.calibrate_intercept(beta, now[cols].to_numpy(), total=4.0)
    now["chance"] = cf.predict_logistic(beta, now[cols].to_numpy())
    contribution = pd.DataFrame({f: beta[1 + i] * now[f"z_{f}"] for i, f in enumerate(features)})
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
        parts = (["Helped by: " + ", ".join(up)] if up else []) + (["Held back by: " + ", ".join(down)] if down else [])
        return ". ".join(parts) if parts else "Close to average on everything the model looks at"

    now["Why"] = [reason(i) for i in now.index]
    now = now.sort_values("chance", ascending=False).reset_index(drop=True)
    out = pd.DataFrame({
        "Rank": np.arange(1, len(now) + 1), "Team": now["team"], "Games played": now["gp"], "Points %": now["pts_pct"].round(3),
        "Goal diff. / game": now["gd"].round(2), "Corsi %": (now["corsi"] * 100).round(1),
        "Expected-goals share %": (now["xg_share"] * 100).round(1), "High-danger share %": (now["hd_share"] * 100).round(1),
        "Final-four chance (model)": (now["chance"] * 100).round(1), "Why": now["Why"]})
    out["Rank by points %"] = now["pts_pct"].rank(ascending=False, method="min").astype(int).to_numpy()
    return out, current, median_gp, checkpoint, through


# ----------------------------------------------------------------------------
# Why does high-danger shot share disappoint? (a diagnostic)
# ----------------------------------------------------------------------------
def shot_quality_check(connection, headline):
    """For each shot-share stat: how well does it line up with results, how reliable is it, and how many shots is it built on?"""
    ts = pd.read_sql_query(
        """SELECT season_id, team_id, highDangerShotsFor hdf, highDangerShotsAgainst hda, mediumDangerShotsFor mdf,
                  mediumDangerShotsAgainst mda, lowDangerShotsFor ldf, lowDangerShotsAgainst lda, xGoalsFor xgf,
                  xGoalsAgainst xga, shotAttemptsFor sf, shotAttemptsAgainst sa
           FROM mp_team_seasons WHERE situation = '5on5' AND season_id >= ?""", connection, params=[ps.FIRST_SEASON])
    pts = pd.read_sql_query("SELECT season_id, team_id, pts_pct FROM team_seasons WHERE include_main = 'Y' AND outcome != 'Season in progress'", connection)
    d = ts.merge(pts, on=["season_id", "team_id"])
    shares = {"Corsi (all shot attempts)": ("sf", "sa"), "Expected-goals share": ("xgf", "xga"), "High-danger shot share": ("hdf", "hda"),
              "Medium-danger shot share": ("mdf", "mda"), "Low-danger shot share": ("ldf", "lda")}
    for name, (f, a) in shares.items():
        d[name] = d[f] / (d[f] + d[a])
    g = pd.read_sql_query(
        """SELECT season_id, team_id, game_id, highDangerShotsFor hdf, highDangerShotsAgainst hda, mediumDangerShotsFor mdf,
                  mediumDangerShotsAgainst mda, lowDangerShotsFor ldf, lowDangerShotsAgainst lda, xGoalsFor xgf, xGoalsAgainst xga,
                  shotAttemptsFor sf, shotAttemptsAgainst sa
           FROM mp_team_games WHERE situation = '5on5' AND playoffGame = 0 AND season_id >= ?""", connection, params=[ps.FIRST_SEASON])
    g = g[g["season_id"].isin(pts["season_id"].unique())]
    g["half"] = g["game_id"] % 2
    sums = g.groupby(["season_id", "team_id", "half"]).sum(numeric_only=True).reset_index()
    alone = {"Corsi (all shot attempts)": "Baseline: Corsi alone", "Expected-goals share": "Baseline: expected-goals share alone",
             "High-danger shot share": "Baseline: high-danger share alone"}
    rows = []
    for name, (f, a) in shares.items():
        explains = d.groupby("season_id").apply(lambda x: x[name].rank().corr(x["pts_pct"].rank()), include_groups=False).mean()
        s = sums.assign(v=sums[f] / (sums[f] + sums[a]))
        wide = s.pivot_table(index=["season_id", "team_id"], columns="half", values="v").reset_index()
        reliability = wide.groupby("season_id").apply(lambda x: x[0].rank().corr(x[1].rank()), include_groups=False).mean()
        count = (d[f] + d[a]).mean() if f not in ("xgf",) else None
        auc = headline.loc[headline["Method"] == alone.get(name, ""), "AUC"]
        rows.append({"Shot-share statistic (5v5)": name, "Link with points % (rank correlation)": round(explains, 2),
                     "Reliability: odd vs even games (rank correlation)": round(reliability, 2),
                     "Shots counted per team-season (for + against)": None if count is None else round(count, 0),
                     "AUC alone for picking final-four teams": None if auc.empty else round(auc.iloc[0], 3)})
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# Verdict
# ----------------------------------------------------------------------------
def verdict_lines(headline, summary, incr, best_model, n_seasons, sq):
    base = headline[headline["Method"] == MAIN_BASELINE].iloc[0]
    v1 = headline[headline["Method"] == V1_NAME].iloc[0]
    best = headline[headline["Method"] == best_model].iloc[0]
    rows = summary[summary["Method"] == best_model]
    margin_v1 = rows["Margin vs v1 (+/-)"].mean()
    margin_base = rows["Margin vs points % (+/-)"].mean()
    wins, losses = int(rows["Seasons better than v1"].sum()), int(rows["Seasons worse than v1"].sum())
    by_name = {r["Idea tested"]: r for _, r in incr.iterrows()}
    f1 = by_name["Your factor 1: high-danger shot share added to v1"]
    f4 = by_name["Your factor 4: goaltending in context added to v2c"]
    quality = by_name["Shot QUALITY in place of shot QUANTITY (Corsi)"]
    corsi_alone = headline[headline["Method"] == "Baseline: Corsi alone"]["AUC"].iloc[0]
    xg_alone = headline[headline["Method"] == "Baseline: expected-goals share alone"]["AUC"].iloc[0]
    hd_alone = headline[headline["Method"] == "Baseline: high-danger share alone"]["AUC"].iloc[0]
    early = summary[(summary["Method"] == best_model) & summary["Checkpoint"].isin(CHECKPOINTS[:3])]
    s = sq.set_index("Shot-share statistic (5v5)")
    hd_n = s.loc["High-danger shot share", "Shots counted per team-season (for + against)"]
    hd_rel, cf_rel = s.loc["High-danger shot share", "Reliability: odd vs even games (rank correlation)"], s.loc["Corsi (all shot attempts)", "Reliability: odd vs even games (rank correlation)"]
    xg_rel = s.loc["Expected-goals share", "Reliability: odd vs even games (rank correlation)"]
    xg_link, cf_link = s.loc["Expected-goals share", "Link with points % (rank correlation)"], s.loc["Corsi (all shot attempts)", "Link with points % (rank correlation)"]
    hd_link = s.loc["High-danger shot share", "Link with points % (rank correlation)"]
    lines = [
        ("Question", "Do shot quality (expected goals, high-danger shots) and goaltending in context pick the real conference finalists better than version 1 (record, goal difference, Corsi, power kill), and better than just ranking by points %?"),
        ("Short answer",
         f"Best version: {best_model}. Averaged over five checkpoints its AUC is {best['AUC']:.3f}, against {v1['AUC']:.3f} for v1 and {base['AUC']:.3f} for points %. "
         f"Gain over v1: {best['AUC'] - v1['AUC']:+.3f} (margin about +/-{margin_v1:.3f}); over points %: {best['AUC'] - base['AUC']:+.3f} (+/-{margin_base:.3f}). "
         + ("That is bigger than the noise." if best["AUC"] - base["AUC"] > margin_base else "Still within the noise.")),
        ("Season by season", f"Against v1 it ranked the real final four better in {wins} season-checkpoints and worse in {losses} (out of {n_seasons * len(CHECKPOINTS)})."),
        ("Your factor 1: high-danger shot share", f"Added to v1: {f1['Average gain (AUC)']:+.3f} (margin +/-{f1['Margin (+/-)']:.3f}); early in the season {f1['Early checkpoints (Nov 1 to Jan 1)']:+.3f}, late {f1['Late checkpoints (deadline, end)']:+.3f}. "
                                                  f"Verdict: {f1['Verdict']}. On its own, ranking by high-danger share gives AUC {hd_alone:.3f}."),
        ("Why high-danger share disappoints", f"A team takes and allows only about {hd_n:,.0f} high-danger shots a season at 5v5, so its share is mostly noise: reliability (odd games vs even games) is "
                                               f"{hd_rel:.2f}, against {cf_rel:.2f} for Corsi and {xg_rel:.2f} for expected-goals share. Its link with points % is {hd_link:.2f}, against {cf_link:.2f} for Corsi and {xg_link:.2f} for expected-goals share. "
                                               "A 'danger bucket' is a coarse way to measure shot quality; expected goals gives every shot its own danger weight and is the better measure."),
        ("Shot quality vs shot quantity", f"Expected-goals share beats Corsi on its own (AUC {xg_alone:.3f} against {corsi_alone:.3f}). Replacing Corsi with the two shot-quality shares changes the v1 AUC by "
                                          f"{quality['Average gain (AUC)']:+.3f} (+/-{quality['Margin (+/-)']:.3f}): {quality['Verdict']}. Once record, goal difference and the power kill are in, shot quality adds little for the final-four question."),
        ("Your factor 4: goaltending in context", f"Added on top of the shot-quality version: {f4['Average gain (AUC)']:+.3f} (margin +/-{f4['Margin (+/-)']:.3f}): {f4['Verdict']}. "
                                                   "Consistent with the goaltending study: goalie results repeat very little, so they add little to a forecast."),
        ("When it helps", f"Best model at Nov 1 / Thanksgiving / Jan 1: average gain over points % of {early['Gain vs points % (AUC)'].mean():+.3f}; at the deadline and end of season the record already tells most of the story."),
        ("Not tested yet", "Zone time and breakouts (factors 2 and 3: they need the play-by-play pull) and depth (top-15 center, elite D pair: they need shift data)."),
    ]
    return pd.DataFrame(lines, columns=["", "Answer"])


def write_excel(sheets, notes):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 126
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
        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 130
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(bold=True)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        if "Right now" in sheets:
            ws = writer.sheets["Right now"]
            ws.column_dimensions["J"].width = 70
            for row in ws.iter_rows(min_row=3):
                row[9].alignment = Alignment(wrap_text=True, vertical="top")


def main():
    sys.stdout.reconfigure(errors="replace")
    if not cf.ps.DB_FILE.exists():
        print("I can't find the database yet. Run: python scripts/build_database.py")
        return 1
    connection = sqlite3.connect(cf.ps.DB_FILE)
    if "mp_team_games" not in [r[0] for r in connection.execute("SELECT name FROM sqlite_master WHERE type='table'")]:
        print("The MoneyPuck tables are not in the database yet. Run: python scripts/load_moneypuck.py")
        return 1
    games_raw, seasons, checkpoints = cf.load(connection)
    games, _ = ps.prepare_games(games_raw, seasons)
    v1_tables = cf.build_checkpoint_tables(games, seasons, checkpoints)
    mp_wide = load_mp_games(connection, ps.FIRST_SEASON)
    tables = build_v2_tables(v1_tables, mp_wide, checkpoints)
    n_seasons = games["season_id"].nunique()
    n_missing = int(tables["season_end"]["xg_share"].isna().sum())
    print(f"Using {n_seasons} main seasons from 2009-10 on ({len(tables['season_end'])} team-seasons; {n_missing} without MoneyPuck data).")

    summary, per_season, scored = evaluate(tables)
    headline = (summary.groupby("Method", sort=False)
                .agg(AUC=("AUC", "mean"), rank_f4=("Avg rank of final-four teams", "mean"), top8=("Final-four teams in top 8", "mean"),
                     gain_base=("Gain vs points % (AUC)", "mean"), gain_v1=("Gain vs v1 (AUC)", "mean")).reset_index())
    best_model = max(MODELS, key=lambda m: headline.loc[headline["Method"] == m, "AUC"].iloc[0])
    incr = increments(per_season)
    sq = shot_quality_check(connection, headline)
    verdict = verdict_lines(headline, summary, incr, best_model, n_seasons, sq)

    right, current, median_gp, used_cp, through = right_now(tables, games_raw, seasons, mp_wide, best_model)
    connection.close()

    cp_table = summary.copy()
    cp_table["Checkpoint"] = cp_table["Checkpoint"].map(PLAIN_CP)
    order = {m: i for i, m in enumerate(list(BASELINES) + list(MODELS))}
    cp_table["m_order"], cp_table["c_order"] = cp_table["Method"].map(order), cp_table["Checkpoint"].map({PLAIN_CP[c]: i for i, c in enumerate(CHECKPOINTS)})
    cp_table = cp_table.sort_values(["c_order", "m_order"]).drop(columns=["m_order", "c_order"]).round(3)
    head_out = headline.rename(columns={"rank_f4": "Avg rank of final-four teams", "top8": "Final-four teams in top 8",
                                        "gain_base": "Gain vs points % (AUC)", "gain_v1": "Gain vs v1 (AUC)"}).round(3)
    scorecard = pd.concat([head_out.assign(Checkpoint="ALL FIVE (average)"), cp_table], ignore_index=True)
    scorecard = scorecard[["Method", "Checkpoint"] + [c for c in scorecard.columns if c not in ("Method", "Checkpoint")]]

    end = scored["season_end"]
    score_cols = [MAIN_BASELINE, V1_NAME, best_model]
    scores_out = pd.concat([t[["season", "team", "outcome"] + list(dict.fromkeys(score_cols))].assign(Checkpoint=PLAIN_CP[cp]) for cp, t in scored.items()], ignore_index=True)
    scores_out[list(dict.fromkeys(score_cols))] = scores_out[list(dict.fromkeys(score_cols))].round(4)

    show_models = list(dict.fromkeys([V1_NAME, best_model, "v2c: v1 + both shot-quality shares"]))
    sheets = {"Verdict": verdict, "Scorecard": scorecard, "What each idea adds": incr.round(3), "Shot quality check": sq,
              "Weights": weights_v2(tables, show_models), "Profile": profile_v2(end), "Rule test": rule_test_v2(end).round(3),
              "Right now": right, "Scores": scores_out}
    notes = {"Right now": (f"Season {str(current)[:4]}-{str(current)[6:]}: games through {through} (median {median_gp} played). Scored with '{best_model}', built on the "
                           f"{PLAIN_CP[used_cp]} pattern of 15 finished seasons, chances adjusted to add to the four final-four spots. A rough guide, not a Cup prediction."),
             "Weights": "What each model rewards, in steps above/below the league average. Record and goal difference move together, so their weights partly cancel."}
    try:
        write_excel(sheets, notes)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 240)
    pd.set_option("display.max_colwidth", 140)
    print("\nAverage over the five checkpoints:")
    print(headline.round(3).to_string(index=False))
    print(f"\nBest model: {best_model}")
    print("\nWhat each idea adds:")
    print(incr.round(3).to_string(index=False))
    print("\nShot quality check:")
    print(sq.to_string(index=False))
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nRight now (top 8):")
    print(right[["Rank", "Team", "Games played", "Points %", "Final-four chance (model)", "Rank by points %"]].head(8).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: MoneyPuck.com (non-commercial use) and (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

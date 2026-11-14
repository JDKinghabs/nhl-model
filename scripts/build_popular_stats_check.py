"""
build_popular_stats_check.py  -  "which popular stats explain, repeat and predict?"
=================================================================================
WHAT IT DOES (in Excel terms):
    For each popular team stat (plus/minus, PDO, Corsi, hits, shooting %, save %, power-play %,
    penalty-kill % ...) it asks three plain questions, season by season, and then averages the answers:

      1. EXPLAINS?   Over a full season, does the stat line up with how many points the team earned?
                     (the correlation between the stat and points %, like =CORREL in Excel, but on ranks)
      2. REPEATS?    Does a team's stat in the FIRST part of the season (up to Jan 1) look like its stat in
                     the REST of the season? If not, the first part was mostly luck, not a team trait.
      3. PREDICTS?   Does the first-part stat predict the points % in the REST of the season, beyond what the
                     record and goal difference already tell you? This is tested honestly: the prediction
                     rule is built on all seasons EXCEPT one, then tried on the one it has never seen,
                     and this is repeated for each season.

    A fourth, rougher question: among playoff teams, does the regular-season stat line up with how many
    rounds the team won?

    The same checks are also run at Thanksgiving and at the trade deadline (see the "By checkpoint" sheet).

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py        (first, so the database is up to date)
    python scripts/build_popular_stats_check.py

WHAT YOU SHOULD SEE:
    A table of verdicts in the terminal and a new file:  outputs/popular_stats_check.xlsx
    (close it in Excel before running again). Start with its "Verdicts" sheet.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
DB_FILE = PROJECT_FOLDER / "data" / "nhl.db"
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "popular_stats_check.xlsx"

FIRST_SEASON = 20092010       # five-on-five and shot-attempt numbers start here
CHECKPOINTS = ["thanksgiving", "jan1", "deadline"]
HEADLINE = "jan1"             # about the middle of the season: the one the Verdicts sheet uses
PLAYOFF_HINT = 0.13           # playoff link must be about twice its margin of error to be mentioned
PLAIN_CP = {"thanksgiving": "Thanksgiving", "jan1": "Jan 1", "deadline": "Trade deadline"}

# (key, short name, what it is in plain English)
STATS = [
    ("pts_pct", "Points % (the record itself)", "Benchmark: points earned out of points possible."),
    ("gd", "Goal difference per game", "Goals for minus goals against per game, all situations. Benchmark."),
    ("pm5", "Plus/minus (team, 5v5)", "Team goals for minus against at 5-on-5 per game: the team version of +/-."),
    ("corsi", "Corsi (shot-attempt share, 5v5)", "Share of all shot attempts (on goal, missed, blocked) the team takes at 5-on-5."),
    ("sog_share", "Shots-on-goal share (5v5)", "Share of shots on goal the team takes at 5-on-5."),
    ("sh_pct", "Shooting % (5v5)", "Goals divided by shots on goal at 5-on-5."),
    ("sv_pct", "Save % (5v5)", "Share of shots on goal the team's goalies stop at 5-on-5 (raw, not adjusted for shot quality)."),
    ("pdo", "PDO (shooting % + save %)", "Shooting % plus save %: a common 'luck' meter (the league average is about 1.000)."),
    ("hits", "Hits per game", "Hits the team throws per game."),
    ("pp_pct", "Power-play %", "Power-play goals divided by power-play chances."),
    ("pk_pct", "Penalty-kill %", "Share of times short-handed that the team does not allow a goal."),
]
KEYS = [s[0] for s in STATS]
NAME = {s[0]: s[1] for s in STATS}
PLAIN = {s[0]: s[2] for s in STATS}

READ_ME = [
    "Popular-stats reality check: does it explain, repeat, and predict?",
    "",
    "THE THREE QUESTIONS (every number is an average over the seasons used, one season at a time)",
    "  Explains   the link between the stat and points % over a full season. Near 0 = unrelated; 1 = perfectly lined up.",
    "  Repeats    the link between the stat before Jan 1 and the same stat after Jan 1. Low = it was mostly luck or",
    "             the team changed. High = a real, stable feature of the team.",
    "  Predicts   does the stat before Jan 1 help forecast points % after Jan 1 BEYOND what the record and goal",
    "             difference already say? Tested on seasons the rule was not built on (leave-one-season-out).",
    "  Playoffs   among playoff teams, the link between the regular-season stat and rounds won, with the regular-season",
    "             record held constant. Small sample (about 16 teams x 15 seasons) so treat as a rough hint.",
    "",
    "HOW TO READ THE NUMBERS",
    "  'Correlation' runs from -1 to +1 and is calculated on ranks within each season (best team = 1 ...), so seasons",
    "  with different scoring levels compare fairly.  About 0.3 is a modest link, 0.5 is strong.",
    "  NOISE BAND: with about 30 teams a season and 15 seasons, differences smaller than roughly 0.1 are within noise.",
    "  'Carry-over' = if a team was 10 units above average before Jan 1, how many units above average was it after?",
    "  'Adds beyond record' = how much better the forecast gets (in out-of-sample R-squared; 0.02 = 2 percentage points",
    "  of the spread in rest-of-season points % explained) and in how many of the seasons it helped.",
    "",
    "HOW THE VERDICT WORDS ARE CHOSEN (rules of thumb: judge by the numbers too)",
    "  Explains: 0.25+ tracks winning, -0.25 or lower runs against it.  Repeats: 0.50+ high, 0.25-0.50 partly, below 0.25 low.",
    "  'Adds' needs a gain of at least 0.01 AND to help in at least 60% of seasons, at Jan 1 and at one other checkpoint.",
    "  'Effect of +1 standard deviation' = at the same record and goal difference, how many standings points over the rest of",
    "  the season a team that is 'one step better' on the stat is expected to earn. A NEGATIVE number means the stat works in",
    "  reverse: the record was flattered by luck and evened out afterwards.",
    "  The playoff hint is only mentioned when the playoff link is at least 0.13 (about twice its margin of error).",
    "  Even a perfectly stable team trait cannot score 1.0 on 'repeats': about 40 games per half is a short sample.",
    "  Hits (0.75) shows what 'almost perfectly stable' looks like here.",
    "  About 10 stats are tested at once, so one or two will look interesting by chance: look for results that hold",
    "  across checkpoints (see the By checkpoint sheet) and across seasons (By season sheet), not for a single number.",
    "",
    "WHAT IT USES",
    "  Regular-season games only. Seasons: main seasons from 2009-10 on (2019-20 and 2020-21 left out, see",
    "  reference/season_flags.csv). Five-on-five stats start in 2009-10, so every stat uses the same seasons.",
    "  Shooting %, save %, PDO and shots-on-goal share use 5-on-5 shots; the one 2009-10 game with no shot data is left out of those.",
    "  Checked: game-by-game Corsi matches the NHL's season Corsi to within 0.0003 for every team-season.",
    "",
    "WHAT IT CANNOT SAY",
    "  It tests team-level stats only (no player stats yet, no shot quality: those come with MoneyPuck). 'Raw' save %",
    "  is not adjusted for shot quality, which is exactly why goalies in front of weak defences look worse.",
    "  Correlation is not cause: a stat that 'repeats and predicts' is a good signal, not proof of why teams win.",
    "",
    "Data: (c) NHL.",
]


# ----------------------------------------------------------------------------
# Loading and building the stats
# ----------------------------------------------------------------------------
def load(connection):
    games = pd.read_sql_query(
        """SELECT season_id, game_id, game_date, team_id, goalsFor, goalsAgainst, points,
                  goalsFor5On5, goalsAgainst5On5, shots5v5, satFor, satAgainst, hits,
                  ppOpportunities, powerPlayGoalsFor, timesShorthanded, ppGoalsAgainst
           FROM team_games WHERE game_type = 2 AND season_id >= ?""", connection, params=[FIRST_SEASON])
    seasons = pd.read_sql_query(
        """SELECT season_id, season, team_id, team, include_main, outcome, made_playoffs
           FROM team_seasons WHERE season_id >= ?""", connection, params=[FIRST_SEASON])
    checkpoints = pd.read_sql_query(
        "SELECT season_id, checkpoint, date FROM checkpoints WHERE reached = 1", connection)
    return games, seasons, checkpoints


def prepare_games(games, seasons):
    """Keep main, finished seasons; add the opponent's shots (our shots against)."""
    main = seasons[(seasons["include_main"] == "Y") & (seasons["outcome"] != "Season in progress")]
    games = games[games["season_id"].isin(main["season_id"].unique())].copy()
    opponent = games[["game_id", "team_id", "shots5v5"]].rename(
        columns={"team_id": "opp_id", "shots5v5": "opp_shots5v5"})
    games = games.merge(opponent, on="game_id")
    games = games[games["team_id"] != games["opp_id"]].copy()
    # One 2009-10 game has no shot data: leave it out of the shot-based stats (not the goal-based ones)
    ok = games["shots5v5"].notna() & games["opp_shots5v5"].notna() & games["satFor"].notna()
    games["gf5_ok"] = games["goalsFor5On5"].where(ok)
    games["ga5_ok"] = games["goalsAgainst5On5"].where(ok)
    games["sf5_ok"] = games["shots5v5"].where(ok)
    games["sa5_ok"] = games["opp_shots5v5"].where(ok)
    games["satf_ok"] = games["satFor"].where(ok)
    games["sata_ok"] = games["satAgainst"].where(ok)
    return games, main


def summarize(games):
    """One row per team per season: every stat, calculated from whichever games are passed in."""
    g = games.groupby(["season_id", "team_id"]).agg(
        gp=("game_id", "count"), pts=("points", "sum"), gf=("goalsFor", "sum"), ga=("goalsAgainst", "sum"),
        gf5=("goalsFor5On5", "sum"), ga5=("goalsAgainst5On5", "sum"),
        gf5_ok=("gf5_ok", "sum"), ga5_ok=("ga5_ok", "sum"), sf5=("sf5_ok", "sum"), sa5=("sa5_ok", "sum"),
        satf=("satf_ok", "sum"), sata=("sata_ok", "sum"), hits=("hits", "sum"),
        ppgf=("powerPlayGoalsFor", "sum"), ppopp=("ppOpportunities", "sum"),
        ppga=("ppGoalsAgainst", "sum"), tsh=("timesShorthanded", "sum")).reset_index()
    out = g[["season_id", "team_id", "gp"]].copy()
    out["pts_pct"] = g["pts"] / (2 * g["gp"])
    out["gd"] = (g["gf"] - g["ga"]) / g["gp"]
    out["pm5"] = (g["gf5"] - g["ga5"]) / g["gp"]
    out["corsi"] = g["satf"] / (g["satf"] + g["sata"])
    out["sog_share"] = g["sf5"] / (g["sf5"] + g["sa5"])
    out["sh_pct"] = g["gf5_ok"] / g["sf5"]
    out["sv_pct"] = 1 - g["ga5_ok"] / g["sa5"]
    out["pdo"] = out["sh_pct"] + out["sv_pct"]
    out["hits"] = g["hits"] / g["gp"]
    out["pp_pct"] = g["ppgf"] / g["ppopp"]
    out["pk_pct"] = 1 - g["ppga"] / g["tsh"]
    return out.replace([np.inf, -np.inf], np.nan)


def split_tables(games, checkpoints):
    """For each checkpoint: a table with every stat 'to date', for 'the rest' and for the 'full' season."""
    full = summarize(games).rename(columns={k: f"{k}_full" for k in KEYS + ["gp"]})
    tables = {}
    for cp in CHECKPOINTS:
        dates = checkpoints[checkpoints["checkpoint"] == cp][["season_id", "date"]]
        t = games.merge(dates, on="season_id")
        before = summarize(t[t["game_date"] <= t["date"]]).rename(columns={k: f"{k}_td" for k in KEYS + ["gp"]})
        after = summarize(t[t["game_date"] > t["date"]]).rename(columns={k: f"{k}_rest" for k in KEYS + ["gp"]})
        merged = before.merge(after, on=["season_id", "team_id"]).merge(full, on=["season_id", "team_id"])
        tables[cp] = merged
    return tables


# ----------------------------------------------------------------------------
# The tests
# ----------------------------------------------------------------------------
def season_spearman(frame, a, b):
    """Rank correlation of two columns, computed inside each season. Returns a Series by season."""
    results = {}
    if a == b:                       # a stat compared with itself says nothing (the record vs the record)
        return pd.Series({season: np.nan for season in frame["season_id"].unique()})
    for season, part in frame.groupby("season_id"):
        part = part[[a, b]].dropna()
        results[season] = part[a].rank().corr(part[b].rank()) if len(part) >= 10 else np.nan
    return pd.Series(results)


def carry_over(frame, key):
    """Share of the 'gap from average' before the checkpoint that is still there after it."""
    d = frame[["season_id", f"{key}_td", f"{key}_rest"]].dropna().copy()
    d["x"] = d[f"{key}_td"] - d.groupby("season_id")[f"{key}_td"].transform("mean")
    d["y"] = d[f"{key}_rest"] - d.groupby("season_id")[f"{key}_rest"].transform("mean")
    return float((d["x"] * d["y"]).sum() / (d["x"] ** 2).sum())


def zscore_by_season(frame, column):
    g = frame.groupby("season_id")[column]
    return (frame[column] - g.transform("mean")) / g.transform("std")


def leave_one_season_out(frame, features, target):
    """
    Build the forecast rule on every season but one, score it on the held-out season, repeat.
    Returns the R-squared (share of the spread explained) for each held-out season.
    """
    scores = {}
    for season in sorted(frame["season_id"].unique()):
        train, test = frame[frame["season_id"] != season], frame[frame["season_id"] == season]
        x_train = np.column_stack([np.ones(len(train))] + [train[f] for f in features])
        beta = np.linalg.lstsq(x_train, train[target], rcond=None)[0]
        x_test = np.column_stack([np.ones(len(test))] + [test[f] for f in features])
        errors = test[target] - x_test @ beta
        spread = ((test[target] - test[target].mean()) ** 2).sum()
        scores[season] = 1 - float((errors ** 2).sum() / spread)
    return pd.Series(scores)


def predict_gain(frame, key):
    """How much does adding this stat improve the forecast of the rest-of-season points %?"""
    d = frame.copy()
    for k in {"pts_pct", "gd", key}:
        d[f"z_{k}_td"] = zscore_by_season(d, f"{k}_td")
    d["z_target"] = zscore_by_season(d, "pts_pct_rest")
    d = d.dropna(subset=[f"z_{k}_td" for k in {"pts_pct", "gd", key}] + ["z_target"])
    if key == "pts_pct":
        return {"base": np.nan, "model": np.nan, "gain": np.nan, "improved": 0, "effect_pts": np.nan}
    base_features = ["z_pts_pct_td"] if key == "gd" else ["z_pts_pct_td", "z_gd_td"]
    base = leave_one_season_out(d, base_features, "z_target")
    model = leave_one_season_out(d, base_features + [f"z_{key}_td"], "z_target")
    gain = model - base

    # Direction and size: fit once on every season. The coefficient is "per 1 standard deviation of the
    # stat, in standard deviations of rest-of-season points %"; convert it to standings points.
    features = base_features + [f"z_{key}_td"]
    x_all = np.column_stack([np.ones(len(d))] + [d[f] for f in features])
    beta = np.linalg.lstsq(x_all, d["z_target"], rcond=None)[0][-1]
    spread = d.groupby("season_id")["pts_pct_rest"].std().mean()          # typical spread of points % after
    effect_pts = float(beta * spread * 2 * d["gp_rest"].mean())            # points % -> standings points
    return {"base": float(base.mean()), "model": float(model.mean()), "gain": float(gain.mean()),
            "improved": int((gain > 0).sum()), "effect_pts": effect_pts}


def playoff_link(full_table, seasons):
    """Among playoff teams: regular-season stat (percentile in its season) vs rounds won, record held constant."""
    rounds = {"Lost round 1": 0, "Lost round 2": 1, "Lost conference final": 2, "Lost Cup final": 3, "Won Cup": 4}
    d = full_table.merge(seasons[["season_id", "team_id", "outcome"]], on=["season_id", "team_id"])
    d = d[d["outcome"].isin(rounds)].copy()
    d["rounds_won"] = d["outcome"].map(rounds)
    z = d["pts_pct_full"].groupby(d["season_id"]).rank(pct=True)
    y = d["rounds_won"].rank()
    out = {}
    for key in KEYS:
        x = d[f"{key}_full"].groupby(d["season_id"]).rank(pct=True)
        ok = x.notna() & z.notna()
        rx, ry, rz = x[ok].rank(), y[ok], z[ok].rank()
        r_xy, r_xz, r_yz = rx.corr(ry), rx.corr(rz), ry.corr(rz)
        partial = (r_xy - r_xz * r_yz) / np.sqrt((1 - r_xz ** 2) * (1 - r_yz ** 2)) if key != "pts_pct" else np.nan
        out[key] = {"raw": float(r_xy), "partial": float(partial), "n": int(ok.sum())}
    return out


def adds_to_forecast(gain, improved, n_seasons):
    """Does the stat improve the held-out forecast: by at least 0.01 and in at least 60% of seasons?"""
    return (not np.isnan(gain)) and gain >= 0.01 and improved >= 0.6 * n_seasons


def verdict_text(key, explains, repeats, forecast, n_seasons, playoff_partial, checkpoints_adding):
    """Pick the plain-English verdict from the numbers (rules of thumb described in the Read me)."""
    if key == "pts_pct":
        return "Benchmark: this is the record itself. Its 'repeats' number is how much a team's record before Jan 1 carries over."
    effect = forecast["effect_pts"]
    # "Adds" must hold at the headline checkpoint AND at least one other, so one lucky split doesn't count
    adds = adds_to_forecast(forecast["gain"], forecast["improved"], n_seasons) and checkpoints_adding >= 2
    tracks, against = explains >= 0.25, explains <= -0.25
    base_name = "record" if key == "gd" else "record and goal difference"
    if key == "gd":
        text = ("Benchmark. Goal difference forecasts the rest of the season better than the record alone."
                if adds else "Benchmark. Goal difference adds little to the record as a forecast.")
    elif adds and effect < 0:
        text = (f"Works in REVERSE: at the same {base_name}, teams that looked better on this stat did worse "
                f"afterwards. That is the signature of luck evening out.")
    elif adds and repeats >= 0.5:
        text = f"Real and repeatable: a stable team trait that tells you something the {base_name} do not."
    elif adds and repeats >= 0.25:
        text = f"A modest signal: partly repeatable, and adds a little beyond the {base_name}."
    elif adds:
        text = "Helps the forecast a little, but does not repeat on its own: be careful with it."
    elif against:
        text = "Runs against winning in this data and adds nothing to the forecast."
    elif tracks and repeats >= 0.25:
        text = f"Describes good teams, but adds nothing beyond the {base_name}."
    elif tracks:
        text = "Tracks results but is mostly luck: it does not repeat, so it does not help predict."
    elif repeats >= 0.25:
        text = "A stable team trait, but unrelated to winning, and it adds nothing to the forecast."
    else:
        text = "Mostly noise: does not line up with results, does not repeat, and adds nothing to the forecast."
    if adds and key != "gd":      # only quote a size when the stat reliably helps; otherwise it is noise
        text += (f" At the same {base_name}, +1 standard deviation on this stat = {effect:+.1f} "
                 f"standings points over the rest of the season.")
    if not np.isnan(playoff_partial) and abs(playoff_partial) >= PLAYOFF_HINT:
        better = "deeper" if playoff_partial > 0 else "less deep"
        text += f" Playoff hint (rough): teams stronger on this stat went {better} than their record alone suggests."
    return text


def label(value, strong=0.5, partly=0.25):
    if np.isnan(value):
        return ""
    a = abs(value)
    return "strong" if a >= strong else "modest" if a >= partly else "weak"


# ----------------------------------------------------------------------------
# Writing the workbook
# ----------------------------------------------------------------------------
def write_excel(sheets, blocks):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 122
        ws["A1"].font = Font(bold=True)
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column in ws.columns:
                longest = max((len(str(c.value)) for c in column[:300] if c.value is not None), default=8)
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 46)
        ws = writer.sheets["Verdicts"]
        for letter, width in (("A", 34), ("B", 60), ("N", 100)):
            ws.column_dimensions[letter].width = width
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[1].height = 62
        # "By season": three small tables one under the other
        row = 0
        for title, frame in blocks.items():
            frame.to_excel(writer, sheet_name="By season", startrow=row + 1)
            ws = writer.sheets["By season"]
            ws.cell(row=row + 1, column=1, value=title).font = Font(bold=True)
            row += len(frame) + 4
        ws = writer.sheets["By season"]
        ws.column_dimensions["A"].width = 12
        for column in list(ws.columns)[1:]:
            ws.column_dimensions[column[0].column_letter].width = 14


def main():
    sys.stdout.reconfigure(errors="replace")
    if not DB_FILE.exists():
        print("I can't find the database yet. Run this first:")
        print("    python scripts/build_database.py")
        return 1
    connection = sqlite3.connect(DB_FILE)
    games, seasons, checkpoints = load(connection)
    connection.close()
    games, main_seasons = prepare_games(games, seasons)
    tables = split_tables(games, checkpoints)
    n_seasons = games["season_id"].nunique()
    print(f"Using {n_seasons} main seasons from 2009-10 on, {len(tables[HEADLINE])} team-seasons.")

    # Quick self-check: each team's games before + after the checkpoint equal the full season
    head = tables[HEADLINE]
    assert ((head["gp_td"] + head["gp_rest"]) == head["gp_full"]).all(), "games before + after should equal the full season"

    # --- Every test, at every checkpoint ---
    detail_rows, headline = [], {}
    per_season = {"Explains: stat vs points % over the full season": {}, "Repeats: stat before Jan 1 vs after": {},
                  "Predicts: stat before Jan 1 vs points % after": {}}
    for cp, frame in tables.items():
        for key in KEYS:
            explains_s = season_spearman(frame, f"{key}_full", "pts_pct_full")
            repeats_s = season_spearman(frame, f"{key}_td", f"{key}_rest")
            predicts_s = season_spearman(frame, f"{key}_td", "pts_pct_rest")
            forecast = predict_gain(frame, key)
            row = {"checkpoint": cp, "key": key, "explains": explains_s.mean(), "repeats": repeats_s.mean(),
                   "predicts": predicts_s.mean(), "carry_over": carry_over(frame, key),
                   "forecast_base": forecast["base"], "forecast_model": forecast["model"],
                   "gain": forecast["gain"], "improved": forecast["improved"], "effect_pts": forecast["effect_pts"]}
            detail_rows.append(row)
            if cp == HEADLINE:
                headline[key] = row
                per_season["Explains: stat vs points % over the full season"][NAME[key]] = explains_s
                per_season["Repeats: stat before Jan 1 vs after"][NAME[key]] = repeats_s
                per_season["Predicts: stat before Jan 1 vs points % after"][NAME[key]] = predicts_s
    detail = pd.DataFrame(detail_rows)
    detail["adds"] = [adds_to_forecast(r.gain, r.improved, n_seasons) for r in detail.itertuples()]
    checkpoints_adding = detail.groupby("key")["adds"].sum().to_dict()

    playoffs = playoff_link(tables[HEADLINE], seasons)

    # --- The Verdicts sheet ---
    verdict_rows = []
    for key in KEYS:
        h, p = headline[key], playoffs[key]
        verdict_rows.append({
            "Statistic": NAME[key],
            "What it is": PLAIN[key],
            "Seasons used": n_seasons,
            "EXPLAINS results (full-season link with points %)": round(h["explains"], 2),
            "REPEATS (before Jan 1 vs after)": round(h["repeats"], 2),
            "Carry-over (share of the gap that is still there)": f"{h['carry_over']:.0%}",
            "PREDICTS points % after Jan 1 (link)": round(h["predicts"], 2),
            "Forecast without the stat (R-squared)": None if np.isnan(h["forecast_base"]) else round(h["forecast_base"], 3),
            "Forecast with the stat (R-squared)": None if np.isnan(h["forecast_model"]) else round(h["forecast_model"], 3),
            "Helps the forecast in how many seasons": "" if np.isnan(h["gain"]) else f"{h['improved']} of {n_seasons}",
            "Checkpoints where it helps (of 3: Thanksgiving, Jan 1, deadline)":
                "" if key == "pts_pct" else int(checkpoints_adding[key]),
            "Effect of +1 standard deviation (standings points over the rest of the season)":
                None if np.isnan(h["effect_pts"]) else round(h["effect_pts"], 1),
            "Playoff rounds won (link, record held constant)": None if np.isnan(p["partial"]) else round(p["partial"], 2),
            "VERDICT": verdict_text(key, h["explains"], h["repeats"], h, n_seasons, p["partial"],
                                    int(checkpoints_adding[key])),
        })
    verdicts = pd.DataFrame(verdict_rows)

    by_checkpoint = detail.copy()
    by_checkpoint["Statistic"] = by_checkpoint["key"].map(NAME)
    by_checkpoint["Checkpoint"] = by_checkpoint["checkpoint"].map(PLAIN_CP)
    by_checkpoint = by_checkpoint[["Statistic", "Checkpoint", "repeats", "carry_over", "predicts", "forecast_base",
                                   "forecast_model", "gain", "improved", "effect_pts"]].rename(columns={
        "repeats": "Repeats (before vs after)", "carry_over": "Carry-over", "predicts": "Predicts points % after (link)",
        "forecast_base": "Forecast R-sq without stat", "forecast_model": "Forecast R-sq with stat",
        "gain": "Gain", "improved": f"Seasons helped (of {n_seasons})",
        "effect_pts": "Effect of +1 SD (standings points, rest of season)"}).round(3)
    by_checkpoint = by_checkpoint.sort_values(["Statistic", "Checkpoint"], key=lambda c: c.map(
        {v: i for i, v in enumerate([NAME[k] for k in KEYS] + ["Thanksgiving", "Jan 1", "Trade deadline"])})).reset_index(drop=True)

    playoff_rows = [{"Statistic": NAME[k], "Playoff teams": playoffs[k]["n"],
                     "Link with rounds won (raw)": round(playoffs[k]["raw"], 2),
                     "Link, record held constant": None if np.isnan(playoffs[k]["partial"]) else round(playoffs[k]["partial"], 2)}
                    for k in KEYS]

    names = seasons.drop_duplicates(["season_id", "team_id"])[["season_id", "season", "team_id", "team"]]
    data = tables[HEADLINE].merge(names, on=["season_id", "team_id"]).round(4)
    first = ["season", "team", "gp_td", "gp_rest"]
    data = data[first + [c for c in data.columns if c not in first + ["season_id", "team_id", "gp_full"]]]

    sheets = {"Verdicts": verdicts, "By checkpoint": by_checkpoint, "Playoffs": pd.DataFrame(playoff_rows),
              "Team-seasons (Jan 1)": data}
    blocks = {title: pd.DataFrame(cols).round(2).rename_axis("season_id") for title, cols in per_season.items()}
    try:
        write_excel(sheets, blocks)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 250)
    pd.set_option("display.max_colwidth", 100)
    print("\nVERDICTS (checkpoint: Jan 1)")
    show = verdicts.copy()
    show.columns = ["stat", "what", "n", "explains", "repeats", "carry", "predicts", "R2 base", "R2 with", "helps",
                    "cps", "effect", "playoffs", "verdict"]
    print(show.drop(columns=["what", "n"]).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

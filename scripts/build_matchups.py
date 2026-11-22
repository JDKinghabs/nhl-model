"""
build_matchups.py  -  every team's profile, for the matchup panels under tonight's games
========================================================================================
WHAT IT DOES (in Excel terms):
    Makes one row per team for the season so far, with the numbers the home page shows when you click a game:
    record (overall, at home, on the road, last 10, streak), goals and shots for and against per game (overall, at
    home and on the road), shooting % and save %, shot-attempt share (Corsi), the scoring-chance numbers from the
    Zone time study, BER and BEI from the Breakouts study, the power play and penalty kill, and the contender-fit
    rank, and from the Goaltending in context study goals saved above expected (GSAx), the defense rank (shot
    quality allowed) and the goaltending rank. Every number also gets a league rank (1 = best). A second sheet lists every game this season, so the page
    can show each team's last 5 results and the season series between the two teams.

    It shows the season so far. It does NOT predict tonight's score, a total or a winner: projections like that are
    betting tools and stay out of the public site.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_matchups.py
    Run it after the daily data steps (the daily update does this): it reads data/nhl.db, data/standings.csv and the
    workbooks of the Zone time, Breakouts and Contender fit studies (any that are missing are simply left out).

WHAT YOU SHOULD SEE:
    A short table and a new file:  outputs/team_profiles.xlsx   (close it in Excel first).

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import re
import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
NHL_DB = PROJECT_FOLDER / "data" / "nhl.db"
STANDINGS = PROJECT_FOLDER / "data" / "standings.csv"
OUTPUTS = PROJECT_FOLDER / "outputs"
OUTPUT_FILE = OUTPUTS / "team_profiles.xlsx"

# Every number that gets a league rank: (column, higher is better?)
RANKED = {
    "pts_pct": True, "gd_pg": True, "gf_pg": True, "ga_pg": False,
    "sf_pg": True, "sa_pg": False, "sf_pg_home": True, "sf_pg_road": True, "sa_pg_home": False, "sa_pg_road": False,
    "sh_pct": True, "sv_pct": True, "sv_pct_home": True, "sv_pct_road": True, "pdo": True,
    "corsi": True, "chance_share": True, "cf60": True, "ca60": False, "oz_min60": True, "ber": True, "bei": True,
    "pp_pct": True, "pk_pct": True, "power_kill": True, "fit": True, "gsax": True,
}


def read_optional(path, sheet, header=1):
    try:
        return pd.read_excel(path, sheet_name=sheet, header=header)
    except (FileNotFoundError, ValueError):
        print(f"  (left out: {path.name} / {sheet} not found)")
        return None


def main():
    sys.stdout.reconfigure(errors="replace")
    if not NHL_DB.exists() or not STANDINGS.exists():
        print("I need data/nhl.db (python scripts/build_database.py) and data/standings.csv (python scripts/pull_schedule_standings.py).")
        return 1
    nhl = sqlite3.connect(NHL_DB)
    season = int(pd.read_sql_query("SELECT MAX(season_id) AS s FROM team_games WHERE game_type = 2", nhl)["s"][0])
    standings = pd.read_csv(STANDINGS)
    ids = pd.read_sql_query("SELECT team_id, team AS team_name FROM team_seasons WHERE season_id = ?", nhl, params=[season])
    teams = standings.merge(ids, on="team_name", how="left")
    if teams["team_id"].isna().any():
        print("Could not match these standings names to team ids:", ", ".join(teams.loc[teams["team_id"].isna(), "team_name"]))
        return 1
    teams["team_id"] = teams["team_id"].astype(int)

    g = pd.read_sql_query(
        """SELECT t.game_id, t.game_date, t.team_id, t.opponent_abbrev, t.home_road, t.goalsFor AS gf, t.goalsAgainst AS ga,
                  t.wins, t.otLosses AS otl, t.shotsForPerGame AS sf, t.shotsAgainstPerGame AS sa, t.satFor AS saf, t.satAgainst AS saa,
                  t.powerPlayGoalsFor AS ppg, t.ppOpportunities AS ppo, t.ppGoalsAgainst AS ppga, t.timesShorthanded AS tsh,
                  g.overtime, g.shootout
           FROM team_games t LEFT JOIN games g ON g.game_id = t.game_id
           WHERE t.season_id = ? AND t.game_type = 2""", nhl, params=[season])
    g = g.merge(teams[["team_id", "team"]], on="team_id", how="left")

    def per_game(frame, suffix=""):
        s = frame.groupby("team_id").agg(n=("game_id", "size"), gf=("gf", "sum"), ga=("ga", "sum"), sf=("sf", "sum"), sa=("sa", "sum"))
        return pd.DataFrame({f"gf_pg{suffix}": s["gf"] / s["n"], f"ga_pg{suffix}": s["ga"] / s["n"], f"sf_pg{suffix}": s["sf"] / s["n"],
                             f"sa_pg{suffix}": s["sa"] / s["n"], f"sh_pct{suffix}": s["gf"] / s["sf"], f"sv_pct{suffix}": 1 - s["ga"] / s["sa"]})

    totals = g.groupby("team_id").agg(saf=("saf", "sum"), saa=("saa", "sum"), ppg=("ppg", "sum"), ppo=("ppo", "sum"), ppga=("ppga", "sum"), tsh=("tsh", "sum"))
    stats = per_game(g).join(per_game(g[g["home_road"] == "H"], "_home")).join(per_game(g[g["home_road"] == "R"], "_road"))
    stats["pdo"] = stats["sh_pct"] + stats["sv_pct"]
    stats["corsi"] = totals["saf"] / (totals["saf"] + totals["saa"])
    stats["pp_pct"] = totals["ppg"] / totals["ppo"]
    stats["pk_pct"] = 1 - totals["ppga"] / totals["tsh"]
    stats["power_kill"] = (stats["pp_pct"] + stats["pk_pct"]) * 100
    t = teams.merge(stats.reset_index(), on="team_id", how="left")
    t["gd_pg"] = t["gd"] / t["gp"]

    # numbers from the research pages (each optional)
    zone = read_optional(OUTPUTS / "zone_chances.xlsx", "This season")
    if zone is not None:
        t = t.merge(zone.rename(columns={"Team": "team_name", "Chance share": "chance_share", "Chances for per 60": "cf60",
                                         "Chances against per 60": "ca60", "Offensive-zone minutes per 60": "oz_min60"})[
            ["team_name", "chance_share", "cf60", "ca60", "oz_min60"]], on="team_name", how="left")
    ber = read_optional(OUTPUTS / "breakout_now.xlsx", "Right now")
    if ber is not None:
        t = t.merge(ber.rename(columns={"Team": "team_name", "BER": "ber"})[["team_name", "ber"]], on="team_name", how="left")
    bei = read_optional(OUTPUTS / "breakout_bei.xlsx", "This season")
    if bei is not None:
        t = t.merge(bei.rename(columns={"Team": "team_name", "BEI": "bei"})[["team_name", "bei"]], on="team_name", how="left")
    fit = read_optional(OUTPUTS / "contender_fit.xlsx", "Right now")
    if fit is not None:
        t = t.merge(fit.rename(columns={"Team": "team_name", "Final-four chance (model)": "fit"})[["team_name", "fit"]], on="team_name", how="left")

    goal = read_optional(OUTPUTS / "goaltending_in_context.xlsx", "Right now")
    goalie_through = ""
    if goal is not None:
        t = t.merge(goal.rename(columns={"Team": "team_name", "GSAx (goals)": "gsax", "Defense rank (1 = best)": "defense_rank",
                                         "Goaltending rank (1 = best)": "goalie_rank", "Label": "goalie_label"})[
            ["team_name", "gsax", "defense_rank", "goalie_rank", "goalie_label"]], on="team_name", how="left")
        t["rank_defense_rank"], t["rank_goalie_rank"] = t["defense_rank"], t["goalie_rank"]       # these two are ranks already
        note = str(pd.read_excel(OUTPUTS / "goaltending_in_context.xlsx", sheet_name="Right now", header=None, nrows=1).iloc[0, 0])
        found = re.search(r"games through (\d{4}-\d{2}-\d{2})", note)
        goalie_through = found.group(1) if found else ""

    for col, higher in RANKED.items():
        if col in t:
            t[f"rank_{col}"] = t[col].rank(ascending=not higher, method="min")

    # every game this season, newest first: for each team's last 5 and the season series
    def result(r):
        if r["wins"] == 1:
            letter = "W"
        elif r["otl"] == 1:
            letter = "OTL"
        else:
            letter = "L"
        extra = " SO" if r["shootout"] == 1 else (" OT" if r["overtime"] == 1 else "")
        return letter + extra

    g["result"] = g.apply(result, axis=1)
    # The game-by-game goals leave out the shootout; the official final score gives the shootout winner one more goal
    so = g["shootout"].eq(1)
    gf = g["gf"] + (so & g["wins"].eq(1)).astype(int)
    ga = g["ga"] + (so & g["wins"].ne(1)).astype(int)
    g["score"] = gf.astype(int).astype(str) + "-" + ga.astype(int).astype(str)
    games = g.sort_values(["game_date", "game_id"], ascending=False)[["team", "game_date", "opponent_abbrev", "home_road", "result", "score", "game_id"]]
    games.columns = ["Team", "Date", "Opponent", "Home or road", "Result", "Score", "Game id"]

    through = g["game_date"].max()
    cols = ["team", "team_name", "conference", "gp", "w", "l", "otl", "home", "road", "l10", "streak", "gd"] + \
           [c for c in RANKED if c in t] + [f"rank_{c}" for c in RANKED if c in t] + \
           [c for c in ("gf_pg_home", "gf_pg_road", "ga_pg_home", "ga_pg_road", "sh_pct_home", "sh_pct_road",
                        "defense_rank", "goalie_rank", "rank_defense_rank", "rank_goalie_rank", "goalie_label") if c in t]
    out = t[cols].sort_values("pts_pct", ascending=False)
    try:
        OUTPUTS.mkdir(exist_ok=True)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            pd.DataFrame({"": [f"Season {str(season)[:4]}-{str(season)[6:]}: games through {through}. One row per team; rank_ columns are league ranks (1 = best)."]}).to_excel(
                writer, sheet_name="Teams", index=False, header=False)
            out.round(4).to_excel(writer, sheet_name="Teams", index=False, startrow=1)
            games.to_excel(writer, sheet_name="Games", index=False)
            pd.DataFrame({"Item": ["games_through", "goalie_through"], "Value": [through, goalie_through]}).to_excel(writer, sheet_name="Notes", index=False)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1
    pd.set_option("display.width", 200)
    print(f"Team profiles for {len(out)} teams, games through {through}.")
    print(out[["team", "pts_pct", "gf_pg", "ga_pg", "sv_pct", "corsi"] + [c for c in ("chance_share", "ber", "bei", "fit") if c in out]].head(5).round(3).to_string(index=False))
    print(f"Saved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""
load_moneypuck.py  -  put the MoneyPuck files into the database and prove they match the NHL's numbers
========================================================================================================
WHAT IT DOES (in Excel terms):
    Reads the files you downloaded (data/moneypuck/), links every MoneyPuck team to the NHL team ids we
    already use, adds three tables to the database (data/nhl.db), and then CHECKS the new numbers against
    the NHL numbers we trust, so we know the data is clean before building on it:

      mp_team_games     every team in every game (regular season AND playoffs), split by situation:
                        all, 5on5, 5on4 (power play), 4on5 (penalty kill), other. Expected goals, shots by
                        danger (low / medium / high), shot attempts, and more.
      mp_team_seasons   the same, one row per team per season (regular season), 2008-09 to now
      mp_goalie_seasons one row per goalie per season and situation: goals, expected goals, shots by danger

    THE CHECKS (each prints a plain pass/fail line):
      * every MoneyPuck team found its NHL team id
      * every NHL game has its MoneyPuck game (and the other way round)
      * goals per game and per season match the NHL's goals
      * 5-on-5 goals match the NHL's 5v5 goals
      * shot attempts match the NHL's shot attempts (Corsi)
      * MoneyPuck's game rows add up to its own season rows

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py          (first: this adds tables to that database)
    python scripts/load_moneypuck.py
    Run it again after build_database.py (which rebuilds the database from scratch) or after new downloads.

WHAT YOU SHOULD SEE:
    Row counts and a list of checks. The MoneyPuck tables appear in data/nhl.db.

DATA CREDIT:
    Data is from MoneyPuck.com, used here for non-commercial purposes. Please credit MoneyPuck.com wherever
    anything built from it is shown. NHL data is (c) NHL.
"""

import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
MP = PROJECT_FOLDER / "data" / "moneypuck"
DB_FILE = PROJECT_FOLDER / "data" / "nhl.db"

# The columns of the game-by-game file worth keeping (the file has 111; most are variations of these)
GAME_COLUMNS = [
    "team", "season", "gameId", "gameDate", "home_or_away", "situation", "playoffGame", "iceTime",
    "xGoalsFor", "xGoalsAgainst", "flurryAdjustedxGoalsFor", "flurryAdjustedxGoalsAgainst",
    "scoreVenueAdjustedxGoalsFor", "scoreVenueAdjustedxGoalsAgainst",
    "goalsFor", "goalsAgainst", "shotsOnGoalFor", "shotsOnGoalAgainst", "missedShotsFor", "missedShotsAgainst",
    "blockedShotAttemptsFor", "blockedShotAttemptsAgainst", "shotAttemptsFor", "shotAttemptsAgainst",
    "unblockedShotAttemptsFor", "unblockedShotAttemptsAgainst",
    "lowDangerShotsFor", "lowDangerShotsAgainst", "mediumDangerShotsFor", "mediumDangerShotsAgainst",
    "highDangerShotsFor", "highDangerShotsAgainst",
    "lowDangerxGoalsFor", "lowDangerxGoalsAgainst", "mediumDangerxGoalsFor", "mediumDangerxGoalsAgainst",
    "highDangerxGoalsFor", "highDangerxGoalsAgainst",
    "lowDangerGoalsFor", "lowDangerGoalsAgainst", "mediumDangerGoalsFor", "mediumDangerGoalsAgainst",
    "highDangerGoalsFor", "highDangerGoalsAgainst",
    "penaltiesFor", "penaltiesAgainst", "hitsFor", "hitsAgainst", "takeawaysFor", "takeawaysAgainst",
    "giveawaysFor", "giveawaysAgainst", "dZoneGiveawaysFor", "dZoneGiveawaysAgainst", "faceOffsWonFor", "faceOffsWonAgainst",
]


def ok(passed):
    return "PASS" if passed else "CHECK"


# ----------------------------------------------------------------------------
# Reading the files
# ----------------------------------------------------------------------------
def read_season_files(kind):
    """History zip + 2025-26 + 2026-27 -> one table. kind = 'teams' or 'goalies'."""
    frames = []
    history = MP / f"{kind}_2008_to_2024" / f"{kind}_2008_to_2024.csv"
    paths = [history] + [MP / f"{kind}_{year}.csv" for year in (2025, 2026)]
    for path in paths:
        if not path.exists():
            print(f"  MISSING {path.name}: run python scripts/import_moneypuck_downloads.py")
            continue
        frame = pd.read_csv(path)
        frame = frame.loc[:, ~frame.columns.str.match(r"^team\.\d+$")]     # the files repeat the 'team' column
        frames.append(frame)
    return pd.concat(frames, ignore_index=True)


def read_game_file():
    path = MP / "all_teams.csv"
    if not path.exists():
        print("  MISSING all_teams.csv: run python scripts/import_moneypuck_downloads.py")
        return None
    return pd.read_csv(path, usecols=GAME_COLUMNS)


# ----------------------------------------------------------------------------
# Linking MoneyPuck teams to NHL team ids
# ----------------------------------------------------------------------------
def team_id_map(mp_games, connection):
    """
    MoneyPuck writes abbreviations its own way (T.B, N.J, ARI ...). Rather than trusting a hand-made table, learn
    each team's NHL id from the games: the NHL game table says who was home and who was away in each game id,
    and MoneyPuck says whether its row is the HOME or AWAY side of that same game id.
    Returns a DataFrame with columns team, season, team_id.
    """
    rows = mp_games[mp_games["situation"] == "all"][["team", "season", "gameId", "home_or_away"]].drop_duplicates()
    nhl = pd.read_sql_query("SELECT game_id, home_team_id, away_team_id FROM games", connection)
    merged = rows.merge(nhl, left_on="gameId", right_on="game_id")
    merged["team_id"] = np.where(merged["home_or_away"].str.upper() == "HOME", merged["home_team_id"], merged["away_team_id"])
    counts = merged.groupby(["team", "season"])["team_id"].agg(lambda s: s.value_counts().index[0]).reset_index()
    purity = merged.groupby(["team", "season"])["team_id"].agg(lambda s: s.value_counts(normalize=True).iloc[0])
    odd = sorted(f"{team} {season}" for (team, season), share in purity.items() if share < 0.99)
    return counts, odd


def check_team_in_game(mp_games, connection):
    """The strong check: each MoneyPuck row's team must be one of the two teams the NHL says played that game."""
    nhl = pd.read_sql_query("SELECT game_id, home_team_id, away_team_id FROM games", connection)
    rows = mp_games[mp_games["situation"] == "all"][["gameId", "team_id"]].drop_duplicates()
    m = rows.merge(nhl, left_on="gameId", right_on="game_id")
    inside = (m["team_id"] == m["home_team_id"]) | (m["team_id"] == m["away_team_id"])
    print(f"\nCheck: every MoneyPuck row's team really played in that game (per the NHL game list)")
    print(f"  {len(m):,} team-games checked; team is one of the two NHL teams in {inside.mean():.2%}  -> {ok(inside.mean() > 0.9999)}")


# ----------------------------------------------------------------------------
# Checks
# ----------------------------------------------------------------------------
def check_games(mp_games, connection):
    print("\nCheck: every NHL game has its MoneyPuck game, and the other way round")
    nhl = pd.read_sql_query("SELECT game_id, season_id, game_type FROM games", connection)
    nhl = nhl[nhl["season_id"] >= 20082009]
    mp_ids = set(mp_games["gameId"].unique())
    missing = nhl[~nhl["game_id"].isin(mp_ids)]
    extra = mp_ids - set(nhl["game_id"])
    print(f"  NHL games from 2008-09 on: {len(nhl):,}; found in MoneyPuck: {len(nhl) - len(missing):,}; "
          f"MoneyPuck games not in NHL list: {len(extra)}  -> {ok(len(missing) <= 10 and not extra)}")
    if len(missing):
        by_season = missing.groupby("season_id").size()
        print(f"    NHL games missing from MoneyPuck, by season: {by_season.to_dict()}")


def check_game_goals(mp_games, connection):
    print("\nCheck: goals in each game (MoneyPuck 'all' situation) vs the NHL's goals")
    tg = pd.read_sql_query("SELECT game_id, team_id, season_id, goalsFor, goalsAgainst FROM team_games", connection)
    allsit = mp_games[mp_games["situation"] == "all"][["gameId", "team_id", "goalsFor", "goalsAgainst"]]
    m = allsit.merge(tg, left_on=["gameId", "team_id"], right_on=["game_id", "team_id"], suffixes=("_mp", "_nhl"))
    same = (m["goalsFor_mp"] == m["goalsFor_nhl"]) & (m["goalsAgainst_mp"] == m["goalsAgainst_nhl"])
    print(f"  {len(m):,} team-games compared; identical goals for and against in {same.mean():.2%}  -> {ok(same.mean() > 0.97)}")
    if not same.all():
        bad = m[~same]
        diff = (bad["goalsFor_mp"] - bad["goalsFor_nhl"]).abs() + (bad["goalsAgainst_mp"] - bad["goalsAgainst_nhl"]).abs()
        so = bad[(bad["goalsFor_nhl"] - bad["goalsFor_mp"]).abs() <= 1]
        print(f"    {len(bad):,} differ; {int((diff <= 2).sum()):,} by 1 or 2 goals in total (typically a shootout winner "
              f"that the NHL counts and MoneyPuck does not); {int((diff > 2).sum())} by more")


def check_season_goals(team_seasons, connection):
    print("\nCheck: season totals vs the NHL (regular season)")
    nhl = pd.read_sql_query(
        "SELECT s.seasonId AS season_id, s.teamId AS team_id, s.goalsFor AS nhl_gf, s.goalsAgainst AS nhl_ga, "
        "g.goalsFor5On5 AS nhl_gf5, p.satFor AS nhl_satf, p.satAgainst AS nhl_sata "
        "FROM ts_summary_regular s "
        "LEFT JOIN ts_goalsforbystrength_regular g ON g.seasonId = s.seasonId AND g.teamId = s.teamId "
        "LEFT JOIN ts_summaryshooting_regular p ON p.seasonId = s.seasonId AND p.teamId = s.teamId", connection)
    nhl = nhl[nhl["team_id"].notna()]
    nhl["team_id"] = nhl["team_id"].astype(int)
    # goals against at 5v5 comes from another table
    ga5 = pd.read_sql_query("SELECT seasonId AS season_id, teamId AS team_id, goalsAgainst5On5 AS nhl_ga5b "
                            "FROM ts_goalsagainstbystrength_regular WHERE teamId IS NOT NULL", connection)
    ga5["team_id"] = ga5["team_id"].astype(int)
    nhl = nhl.merge(ga5, on=["season_id", "team_id"], how="left")
    pivots = {}
    for sit in ("all", "5on5"):
        part = team_seasons[team_seasons["situation"] == sit]
        pivots[sit] = part[["season_id", "team_id", "goalsFor", "goalsAgainst", "shotAttemptsFor", "shotAttemptsAgainst"]].rename(
            columns=lambda c: c if c in ("season_id", "team_id") else f"{c}_{sit}")
    m = nhl.merge(pivots["all"], on=["season_id", "team_id"]).merge(pivots["5on5"], on=["season_id", "team_id"])
    finished = m[m["season_id"] != m["season_id"].max()]

    def line(label, mine, theirs, tolerance):
        d = (finished[mine] - finished[theirs]).abs()
        exact = (d == 0).mean()
        close = (d <= tolerance).mean()
        print(f"  {label:<34} {len(finished)} team-seasons: exact {exact:.1%}, within {tolerance} {close:.1%}, "
              f"largest gap {d.max():.0f}  -> {ok(close > 0.97)}")

    line("goals for (all situations)", "goalsFor_all", "nhl_gf", 3)
    line("goals against (all situations)", "goalsAgainst_all", "nhl_ga", 3)
    seasons_5v5 = finished[finished["nhl_gf5"].notna()]
    for label, mine, theirs in (("5v5 goals for", "goalsFor_5on5", "nhl_gf5"), ("5v5 goals against", "goalsAgainst_5on5", "nhl_ga5b")):
        d = (seasons_5v5[mine] - seasons_5v5[theirs]).abs()
        print(f"  {label:<34} {len(seasons_5v5)} team-seasons (2009-10 on): exact {(d == 0).mean():.1%}, "
              f"within 3 {(d <= 3).mean():.1%}, largest gap {d.max():.0f}  -> {ok((d <= 3).mean() > 0.9)}")
    sat = finished[finished["nhl_satf"].notna() & (finished["nhl_satf"] > 0)]
    corr = np.corrcoef(sat["shotAttemptsFor_5on5"], sat["nhl_satf"])[0, 1]
    ratio = (sat["shotAttemptsFor_5on5"] / sat["nhl_satf"]).median()
    print(f"  {'5v5 shot attempts for':<34} {len(sat)} team-seasons: correlation {corr:.3f}, median ratio MoneyPuck/NHL {ratio:.3f}  -> {ok(corr > 0.98)}")


def check_internal(mp_games, team_seasons):
    print("\nCheck: MoneyPuck's game rows add up to its own season rows (expected goals, regular season)")
    games = mp_games[(mp_games["playoffGame"] == 0) & (mp_games["situation"] == "all")]
    summed = games.groupby(["season", "team"])["xGoalsFor"].sum().reset_index().rename(columns={"xGoalsFor": "from_games"})
    seasons = team_seasons[team_seasons["situation"] == "all"][["season", "team", "xGoalsFor"]].rename(columns={"xGoalsFor": "from_season"})
    m = summed.merge(seasons, on=["season", "team"])
    d = (m["from_games"] - m["from_season"]).abs()
    print(f"  {len(m)} team-seasons: median gap {d.median():.2f} xGoals, largest {d.max():.2f}  -> {ok(d.median() < 1.0)}")


def main():
    sys.stdout.reconfigure(errors="replace")
    if not DB_FILE.exists():
        print("I can't find the database yet. Run this first:")
        print("    python scripts/build_database.py")
        return 1
    print("Reading the MoneyPuck files...")
    mp_games = read_game_file()
    if mp_games is None:
        return 1
    team_seasons = read_season_files("teams")
    goalie_seasons = read_season_files("goalies")
    print(f"  game-by-game rows: {len(mp_games):,}; team-season rows: {len(team_seasons):,}; goalie-season rows: {len(goalie_seasons):,}")

    connection = sqlite3.connect(DB_FILE)
    mapping, odd = team_id_map(mp_games, connection)
    print(f"\nLinking MoneyPuck teams to NHL team ids: {len(mapping)} (abbreviation, season) pairs.")
    if odd:
        print(f"  Note: for {', '.join(odd)} MoneyPuck's HOME/AWAY labels disagree with the NHL's (it labels every game AWAY). "
              f"The team is taken from the majority of games, and the check below confirms it. Do not trust MoneyPuck's "
              f"home_or_away column for those seasons.")

    mp_games = mp_games.merge(mapping, on=["team", "season"], how="left")
    team_seasons = team_seasons.merge(mapping, on=["team", "season"], how="left")
    goalie_seasons = goalie_seasons.merge(mapping, on=["team", "season"], how="left")
    for name, frame in (("game rows", mp_games), ("team-season rows", team_seasons)):
        unmatched = frame[frame["team_id"].isna()]
        print(f"  {name} with no NHL team id: {len(unmatched)}"
              + (f" ({sorted(unmatched['team'].unique())})" if len(unmatched) else "") + f"  -> {ok(len(unmatched) == 0)}")
    goalie_unmatched = goalie_seasons[goalie_seasons["team_id"].isna()]
    print(f"  goalie-season rows with no NHL team id: {len(goalie_unmatched)} (goalies on a team that played no game that season, "
          f"or multi-team totals)")

    for frame in (mp_games, team_seasons, goalie_seasons):
        frame["season_id"] = (frame["season"].astype(int) * 10000 + frame["season"].astype(int) + 1).astype(int)
    mp_games["team_id"] = mp_games["team_id"].astype("Int64")
    team_seasons["team_id"] = team_seasons["team_id"].astype("Int64")
    mp_games = mp_games.rename(columns={"gameId": "game_id", "gameDate": "game_date", "home_or_away": "home_or_away"})
    mp_games["game_date"] = pd.to_datetime(mp_games["game_date"].astype(str), format="%Y%m%d").dt.strftime("%Y-%m-%d")

    # --- checks (use the NHL tables already in the database) ---
    check_team_in_game(mp_games.rename(columns={"game_id": "gameId"}), connection)
    check_games(mp_games.rename(columns={"game_id": "gameId"}), connection)
    check_game_goals(mp_games.rename(columns={"game_id": "gameId"}), connection)
    check_season_goals(team_seasons.dropna(subset=["team_id"]).assign(team_id=lambda d: d["team_id"].astype(int)), connection)
    check_internal(mp_games.rename(columns={"game_id": "gameId"}), team_seasons)

    # --- save to the database ---
    with connection:
        mp_games.to_sql("mp_team_games", connection, index=False, if_exists="replace", chunksize=50000)
        team_seasons.to_sql("mp_team_seasons", connection, index=False, if_exists="replace")
        goalie_seasons.to_sql("mp_goalie_seasons", connection, index=False, if_exists="replace")
        connection.execute("CREATE INDEX IF NOT EXISTS ix_mp_games ON mp_team_games (season_id, team_id, game_date, situation)")
        connection.execute("CREATE INDEX IF NOT EXISTS ix_mp_game_id ON mp_team_games (game_id, team_id, situation)")
        connection.execute("CREATE INDEX IF NOT EXISTS ix_mp_ts ON mp_team_seasons (season_id, team_id, situation)")
    connection.close()
    print(f"\nSaved to {DB_FILE}: mp_team_games ({len(mp_games):,} rows), mp_team_seasons ({len(team_seasons):,}), "
          f"mp_goalie_seasons ({len(goalie_seasons):,})")
    print("Data: MoneyPuck.com (non-commercial use, credit required); NHL data (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

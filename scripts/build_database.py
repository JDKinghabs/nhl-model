"""
build_database.py  -  put everything we downloaded into ONE database file
=========================================================================
WHAT IT DOES (in Excel terms):
    A SQLite database is like a workbook that can hold millions of rows across many sheets
    ("tables") and answer questions about them quickly, all in one file: data/nhl.db.
    This script rebuilds that file from the CSV files the other scripts saved:

      games                 one row per game (date, home/away, score, overtime?)
      team_games            one row per team per game: every game-by-game statistic we pulled
      standings_snapshots   standings on each checkpoint date (Nov 1, Thanksgiving ...)
      checkpoints           the checkpoint dates and how each was chosen
      team_seasons          one row per team per season, with how far they went in the playoffs
      playoff_series        every playoff series
      seasons, season_flags, franchises, trade_deadlines    small lookup tables
      ts_<report>_<regular|playoffs>   the team SEASON tables from the NHL (power play, etc.)
      edge_zone_time_<regular|playoffs>   the NHL's real zone-time data (2021-22 on)

    It is safe to run again any time: the file is just rebuilt from the CSV files. Nothing here
    is downloaded; it only reads what is already on your computer.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py

WHAT YOU SHOULD SEE:
    Row counts for every table and three checks that the game-by-game rows add up to the
    season totals the NHL reports. The file  data/nhl.db  appears (close any program that has it open).

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import re
import sqlite3
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
DATA = PROJECT_FOLDER / "data"
REFERENCE = PROJECT_FOLDER / "reference"
DB_FILE = DATA / "nhl.db"

KEY_RENAMES = {"seasonId": "season_id", "gameTypeId": "game_type", "gameId": "game_id",
               "gameDate": "game_date", "teamId": "team_id", "teamFullName": "team",
               "opponentTeamAbbrev": "opponent_abbrev", "homeRoad": "home_road"}


def read(path):
    return pd.read_csv(path, encoding="utf-8-sig") if Path(path).exists() else None


# ----------------------------------------------------------------------------
# Game tables
# ----------------------------------------------------------------------------
def build_team_games():
    """Stack regular season + playoffs and attach every game-by-game report to the summary rows."""
    folder = DATA / "nhl" / "team_games"
    parts = []
    for kind in ("regular", "playoffs"):
        base = read(folder / f"summary_{kind}.csv")
        if base is None:
            continue
        base = base[base["teamId"].notna()].copy()
        for path in sorted(folder.glob(f"*_{kind}.csv")):
            report = path.stem[: -(len(kind) + 1)]
            if report == "summary":
                continue
            extra = pd.read_csv(path, encoding="utf-8-sig")
            extra = extra[extra["teamId"].notna()]
            # Keep only columns the summary does not already have (plus the two that identify the row)
            new_cols = [c for c in extra.columns if c not in base.columns]
            base = base.merge(extra[["gameId", "teamId"] + new_cols].drop_duplicates(["gameId", "teamId"]),
                              on=["gameId", "teamId"], how="left")
        parts.append(base)
    if not parts:
        return None
    team_games = pd.concat(parts, ignore_index=True).rename(columns=KEY_RENAMES)
    team_games["team_id"] = team_games["team_id"].astype(int)

    # The NHL leaves the goals-by-strength cells EMPTY when a team scored (or allowed) none in that
    # situation, instead of writing 0 (about 7% of team-games: the shutouts). In seasons where that
    # report has numbers at all, an empty cell means zero goals. (Check 4 below proves it: with this
    # fix the game-by-game 5v5 goals add up to the NHL's season totals for every team.)
    strength_columns = [c for c in team_games.columns
                        if re.match(r"goals(For|Against)(\dOn\d|EmptyNet|ExtraAttacker|PenaltyShots)$", c)]
    for column in strength_columns:
        season_has_numbers = team_games.groupby("season_id")[column].transform(lambda s: s.notna().any())
        team_games.loc[season_has_numbers & team_games[column].isna(), column] = 0
    return team_games.sort_values(["season_id", "game_date", "game_id", "team_id"]).reset_index(drop=True)


def build_games(team_games):
    """Turn two team rows per game into one row per game."""
    home = team_games[team_games["home_road"] == "H"]
    road = team_games[team_games["home_road"] == "R"]
    games = home[["season_id", "game_type", "game_id", "game_date", "team_id", "goalsFor"]].merge(
        road[["game_id", "team_id", "goalsFor"]], on="game_id", suffixes=("_home", "_away"))
    games = games.rename(columns={"team_id_home": "home_team_id", "team_id_away": "away_team_id",
                                  "goalsFor_home": "home_goals", "goalsFor_away": "away_goals"})
    went_long = team_games.groupby("game_id").agg(
        overtime=("otLosses", lambda s: bool((s == 1).any())),
        shootout=("winsInShootout", lambda s: bool((s == 1).any())))
    games = games.merge(went_long, on="game_id", how="left")
    return games.sort_values(["season_id", "game_date", "game_id"]).reset_index(drop=True)


# ----------------------------------------------------------------------------
# Checks
# ----------------------------------------------------------------------------
def run_checks(team_games, games):
    print("\nCheck 1: every game has exactly two team rows (one per side)")
    per_game = team_games.groupby("game_id").size()
    odd = per_game[per_game != 2]
    print(f"  {len(per_game):,} games, {len(odd)} with the wrong number of rows" + ("" if odd.empty else f": {list(odd.index[:5])}"))

    print("\nCheck 2: game-by-game totals equal the NHL's season totals (regular season)")
    season = read(DATA / "nhl" / "team_season" / "summary_regular.csv")
    if season is None:
        print("  (season summary file not found; skipped)")
        return
    season = season[season["teamId"].notna()].rename(columns={"seasonId": "season_id", "teamId": "team_id"})
    season["team_id"] = season["team_id"].astype(int)
    summed = (team_games[team_games["game_type"] == 2]
              .groupby(["season_id", "team_id"])
              .agg(gp_g=("gamesPlayed", "sum"), pts_g=("points", "sum"), gf_g=("goalsFor", "sum"),
                   ga_g=("goalsAgainst", "sum"), wins_g=("wins", "sum")).reset_index())
    both = season.merge(summed, on=["season_id", "team_id"], how="outer", indicator=True)
    missing = both[both["_merge"] != "both"]
    both = both[both["_merge"] == "both"]
    latest = int(both["season_id"].max())
    for label, mine, theirs in [("games played", "gp_g", "gamesPlayed"), ("points", "pts_g", "points"),
                                ("goals for", "gf_g", "goalsFor"), ("goals against", "ga_g", "goalsAgainst"),
                                ("wins", "wins_g", "wins")]:
        bad = both[both[mine] != both[theirs]]
        old_bad = bad[bad["season_id"] != latest]
        line = f"  {label:<13} {len(both) - len(bad):>4} of {len(both)} team-seasons match"
        if not old_bad.empty:
            line += f"   DIFFERENCES in finished seasons: {sorted(int(s) for s in old_bad['season_id'].unique())[:6]}"
        elif not bad.empty:
            line += f"   (only the current season differs: the two tables were downloaded on different days)"
        print(line)
    if len(missing):
        print(f"  {len(missing)} team-seasons exist in only one of the two tables")

    print("\nCheck 3: one game per day shows up for every team at most once")
    dupes = team_games.duplicated(["season_id", "game_type", "team_id", "game_date"]).sum()
    print(f"  {dupes} team-days with two games" + ("" if dupes == 0 else " (possible doubleheader or data error)"))

    print("\nCheck 4: game-by-game 5v5 goals add up to the NHL's season totals")
    for label, mine, season_file in [("5v5 goals for", "goalsFor5On5", "goalsforbystrength_regular"),
                                     ("5v5 goals against", "goalsAgainst5On5", "goalsagainstbystrength_regular")]:
        theirs = read(DATA / "nhl" / "team_season" / f"{season_file}.csv")
        if theirs is None or mine not in theirs.columns:
            continue
        theirs = theirs[theirs["teamId"].notna() & theirs[mine].notna()].rename(
            columns={"seasonId": "season_id", "teamId": "team_id", mine: "season_total"})
        theirs["team_id"] = theirs["team_id"].astype(int)
        mine_sum = (team_games[team_games["game_type"] == 2].groupby(["season_id", "team_id"])[mine]
                    .sum().reset_index())
        both = theirs.merge(mine_sum, on=["season_id", "team_id"])
        both = both[both["season_id"] != both["season_id"].max()]
        ok = int((both["season_total"] == both[mine]).sum())
        print(f"  {label:<18} {ok} of {len(both)} finished team-seasons match")


def main():
    sys.stdout.reconfigure(errors="replace")
    team_games = build_team_games()
    if team_games is None:
        print("I can't find the game-by-game files yet. Run this first:")
        print("    python scripts/pull_game_results.py")
        return 1
    games = build_games(team_games)

    tables = {"games": games, "team_games": team_games}
    simple = {
        "standings_snapshots": DATA / "nhl" / "standings_snapshots.csv",
        "checkpoints": DATA / "reference" / "checkpoints.csv",
        "team_seasons": DATA / "reference" / "team_seasons.csv",
        "seasons": DATA / "reference" / "seasons.csv",
        "playoff_series": DATA / "reference" / "playoff_series.csv",
        "season_flags": REFERENCE / "season_flags.csv",
        "franchises": REFERENCE / "franchises.csv",
        "trade_deadlines": REFERENCE / "trade_deadlines.csv",
    }
    for name, path in simple.items():
        frame = read(path)
        if frame is None:
            print(f"  (skipping {name}: {path.name} not found yet)")
        else:
            tables[name] = frame
    for kind in ("regular", "playoffs"):
        for path in sorted((DATA / "nhl" / "team_season").glob(f"*_{kind}.csv")):
            tables[f"ts_{path.stem}"] = pd.read_csv(path, encoding="utf-8-sig")
        edge = read(DATA / "nhl" / "edge" / f"zone_time_{kind}.csv")
        if edge is not None:
            tables[f"edge_zone_time_{kind}"] = edge

    DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    try:
        if DB_FILE.exists():
            DB_FILE.unlink()
        connection = sqlite3.connect(DB_FILE)
    except (PermissionError, sqlite3.Error) as problem:
        print(f"Could not rebuild {DB_FILE.name}: {problem}. Is another program using it?")
        return 1

    with connection:
        for name, frame in tables.items():
            for column in frame.columns:                      # dates as plain text 2025-01-12
                if frame[column].dtype == object and len(frame) and hasattr(frame[column].iloc[0], "isoformat"):
                    frame[column] = frame[column].astype(str)
            frame.to_sql(name, connection, index=False, if_exists="replace")
        # Indexes make the common questions fast ("this team, this season, in date order")
        connection.execute("CREATE INDEX ix_team_games ON team_games (season_id, team_id, game_date)")
        connection.execute("CREATE INDEX ix_games ON games (season_id, game_date)")
        if "standings_snapshots" in tables:
            connection.execute("CREATE INDEX ix_snap ON standings_snapshots (season_id, checkpoint, team_id)")

    print(f"Built {DB_FILE}  ({DB_FILE.stat().st_size / 1e6:.1f} MB)\n")
    print(f"{'table':<36}{'rows':>10}")
    for name, frame in tables.items():
        print(f"{name:<36}{len(frame):>10,}")

    run_checks(team_games, games)
    connection.close()
    if (PROJECT_FOLDER / "data" / "moneypuck").exists() and any((PROJECT_FOLDER / "data" / "moneypuck").glob("*.csv")):
        print("\nNote: this rebuild starts the database from scratch, so the MoneyPuck tables are gone.")
        print("      Run  python scripts/load_moneypuck.py  to add them back.")
    print("\nData: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

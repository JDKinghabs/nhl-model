"""
pull_edge_zone_time.py  -  how much of the game each team spends in each zone (NHL "EDGE" data)
================================================================================================
WHAT IT DOES (in Excel terms):
    For every team and season from 2021-22 on, asks the NHL's EDGE tracking pages: "what share of
    the time was this team's play in the offensive zone, the neutral zone and its own defensive
    zone?" It gets four versions of the answer: all situations, 5-on-5 ("es"), power play ("pp")
    and penalty kill ("pk"), each with the team's rank in the league and the league average.
    It also gets each team's shot-attempt differential (shots for minus against, per game).

    This is the only REAL zone-time data that exists (it comes from the NHL's puck and player
    tracking). Before 2021-22 we can only estimate it, and this file is what we will check that
    estimate against.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/pull_edge_zone_time.py

    Run  python scripts/pull_team_season_reports.py  first: it saves the list of team ids this
    script needs.

WHAT YOU SHOULD SEE:
    One line per season, then a short summary. Files appear in  data/nhl/edge/ :
      zone_time_regular.csv   (one row per team per season)
      zone_time_playoffs.csv
      zone_time_log.csv       (what was asked and what came back: used by the coverage report)
    It remembers finished seasons, so running it again only asks for the current season.

DATA CREDIT:
    Data is (c) NHL EDGE, used here for personal, non-commercial purposes. Unofficial, can change.
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

import nhl_api          # our shared, polite downloader (scripts/nhl_api.py)

# ----------------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------------
FIRST_START_YEAR = 2021          # EDGE zone time starts with the 2021-22 season
GAME_TYPES = {2: "regular", 3: "playoffs"}

URL = "https://api-web.nhle.com/v1/edge/team-zone-time-details/{team_id}/{season}/{game_type}"

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
TEAM_FOLDER = PROJECT_FOLDER / "data" / "nhl" / "team_season"
SAVE_FOLDER = PROJECT_FOLDER / "data" / "nhl" / "edge"      # "data" is ignored by git
LOG_FILE = SAVE_FOLDER / "zone_time_log.csv"

ZONE_NAMES = {"offensiveZone": "oz", "neutralZone": "nz", "defensiveZone": "dz"}


def current_start_year():
    """The season that is under way now (seasons start in the autumn)."""
    today = datetime.now()
    return today.year if today.month >= 9 else today.year - 1


def flatten(reply):
    """
    The NHL sends the answer as nested boxes. Flatten it into one row, like this:
      all_oz_pct, all_oz_rank, all_oz_league_avg, all_nz_pct, ... es_oz_pct ... pk_dz_pct,
      sa_diff, sa_diff_rank, sog_diff, sog_diff_rank
    (oz = offensive zone, nz = neutral zone, dz = defensive zone)
    """
    row = {}
    for block in reply.get("zoneTimeDetails", []):
        code = block.get("strengthCode")
        for long_name, short in ZONE_NAMES.items():
            row[f"{code}_{short}_pct"] = block.get(f"{long_name}Pctg")
            row[f"{code}_{short}_rank"] = block.get(f"{long_name}Rank")
            row[f"{code}_{short}_league_avg"] = block.get(f"{long_name}LeagueAvg")
    shots = reply.get("shotDifferential") or {}
    row["sa_diff"] = shots.get("shotAttemptDifferential")
    row["sa_diff_rank"] = shots.get("shotAttemptDifferentialRank")
    row["sog_diff"] = shots.get("sogDifferential")
    row["sog_diff_rank"] = shots.get("sogDifferentialRank")
    return row


def load_team_ids():
    """
    Which teams to ask about, from the team summary files the other script saved:
    every team for the regular season, only the teams that played for the playoffs.
    Returns {game type: table}, or None if those files are not there yet.
    """
    lists = {}
    for game_type, name in GAME_TYPES.items():
        path = TEAM_FOLDER / f"summary_{name}.csv"
        if not path.exists():
            print("I need the list of teams first. Run this, then run me again:")
            print("    python scripts/pull_team_season_reports.py")
            return None
        teams = pd.read_csv(path, encoding="utf-8-sig").rename(columns={"seasonId": "season_id"})
        lists[game_type] = teams[["season_id", "teamId", "teamFullName"]].drop_duplicates()
    return lists


def main():
    sys.stdout.reconfigure(errors="replace")
    team_lists = load_team_ids()
    if team_lists is None:
        return 1

    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)
    this_year = current_start_year()

    # Our memory of what we already asked (so finished seasons are not asked twice).
    log = pd.read_csv(LOG_FILE) if LOG_FILE.exists() else pd.DataFrame(
        columns=["season_id", "game_type", "teams_asked", "teams_with_data", "pulled_at"])
    stored = {name: [] for name in GAME_TYPES.values()}
    for name in GAME_TYPES.values():
        path = SAVE_FOLDER / f"zone_time_{name}.csv"
        if path.exists():
            stored[name].append(pd.read_csv(path, encoding="utf-8-sig"))

    new_rows = {name: [] for name in GAME_TYPES.values()}
    new_log = []

    for start_year in range(FIRST_START_YEAR, this_year + 1):
        season_id = int(f"{start_year}{start_year + 1}")
        for game_type, name in GAME_TYPES.items():
            season_teams = team_lists[game_type][team_lists[game_type]["season_id"] == season_id]
            if season_teams.empty:
                continue                      # no team list for this season/type (e.g. playoffs not played yet)
            already = log[(log["season_id"] == season_id) & (log["game_type"] == game_type)]
            if not already.empty and start_year < this_year:
                continue                      # a finished season we already have: skip
            got = 0
            for _, team in season_teams.iterrows():
                reply = nhl_api.get_json(
                    URL.format(team_id=int(team["teamId"]), season=season_id, game_type=game_type),
                    quiet_404=True)
                if not reply or not reply.get("zoneTimeDetails"):
                    continue                  # no data for this team (e.g. missed the playoffs)
                row = {"season_id": season_id, "teamId": int(team["teamId"]),
                       "team": team["teamFullName"]}
                row.update(flatten(reply))
                new_rows[name].append(row)
                got += 1
            print(f"{start_year}-{str(start_year + 1)[2:]} {name:<8}: {got} of {len(season_teams)} teams had data")
            new_log.append({"season_id": season_id, "game_type": game_type,
                            "teams_asked": len(season_teams), "teams_with_data": got,
                            "pulled_at": datetime.now().strftime("%Y-%m-%d %H:%M")})

    # Merge new rows into what we already had (a re-pulled season replaces its old rows).
    for name in GAME_TYPES.values():
        if not new_rows[name]:
            continue
        fresh = pd.DataFrame(new_rows[name])
        old = stored[name][0] if stored[name] else pd.DataFrame()
        if not old.empty:
            old = old[~old["season_id"].isin(fresh["season_id"].unique())]
        combined = pd.concat([old, fresh], ignore_index=True).sort_values(["season_id", "team"])
        combined.to_csv(SAVE_FOLDER / f"zone_time_{name}.csv", index=False, encoding="utf-8-sig")
        print(f"Saved zone_time_{name}.csv  ({len(combined)} rows)")

    if new_log:
        fresh_log = pd.DataFrame(new_log)
        keep = log.merge(fresh_log[["season_id", "game_type"]], on=["season_id", "game_type"],
                         how="left", indicator=True)
        keep = keep[keep["_merge"] == "left_only"].drop(columns="_merge")
        pd.concat([keep, fresh_log], ignore_index=True).sort_values(["season_id", "game_type"]) \
            .to_csv(LOG_FILE, index=False)

    print("Data: (c) NHL EDGE")
    return 0


if __name__ == "__main__":
    sys.exit(main())

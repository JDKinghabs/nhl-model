"""
pull_team_season_reports.py  -  every team's season totals, 2005-06 to now (NHL stats site)
==========================================================================================
WHAT IT DOES (in Excel terms):
    The NHL stats site has many "reports", each one a table with one row per team for a season
    (summary, power play, penalty kill, goals by strength, shot attempts, hits, faceoffs ...).
    This script downloads each report for every season since 2005-06, for both the regular
    season and the playoffs, and stacks the seasons into ONE table per report (like pasting
    each season's sheet underneath the last one).

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/pull_team_season_reports.py

    It takes about 10-15 minutes the first time because it waits one second between requests
    (about 600 of them). It is safe to stop and run again: seasons it already has are skipped,
    and the current season is refreshed every time.

    Optional:
      --only summary,realtime   only pull those reports
      --force                   ask again for everything, even seasons we already have

WHAT YOU SHOULD SEE:
    A line per report as it goes, then a summary table. New files appear in
    data/nhl/team_season/ :   summary_regular.csv, summary_playoffs.csv, powerplay_regular.csv ...
    plus  pull_log.csv  (what was asked and what came back; the coverage report reads it).

    Some reports do not exist for the oldest seasons (the NHL did not record shot attempts
    before 2009-10, for example). Those show as "empty": that is a finding, not a mistake.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. Unofficial, can change.
"""

import argparse
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

import nhl_api          # our shared, polite downloader (scripts/nhl_api.py)

# pandas warns when it stacks tables where a column is empty for some seasons. That is expected
# here (old seasons have empty columns) and harmless, so keep the screen quiet.
warnings.filterwarnings("ignore", category=FutureWarning, message=".*concatenation.*")

# ----------------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------------
FIRST_START_YEAR = 2005          # 2005-06: first season after the lockout, start of the salary cap

# The reports worth having for our questions. More exist (shot types, shootouts, days of rest ...):
# add a name to this list and run again to pull it.
REPORTS = [
    "summary",                   # points, goals, shots per game, faceoff %, PP% and PK%
    "powerplay",                 # PP goals, opportunities, time on ice, short-handed goals against
    "penaltykill",               # PK goals against, times short-handed, PK time, short-handed goals
    "powerplaytime",             # PP split by 5v4 / 4v3 / 5v3
    "penaltykilltime",           # PK split by 4v5 / 3v5 / 3v4
    "goalsforbystrength",        # goals for at 5v5, 5v4, 4v5, empty net ...
    "goalsagainstbystrength",    # goals against at 5v5, 5v4 ...
    "percentages",               # shot-attempt share (Corsi), 5v5 shooting% and save%, PDO, zone starts
    "summaryshooting",           # shot attempts for/against, in different score situations
    "realtime",                  # hits, blocks, giveaways, takeaways, shot attempts
    "faceoffpercentages",        # faceoff % by zone and by strength
    "penalties",                 # penalties taken and drawn
    "scoretrailfirst",           # record when scoring first or trailing first
    "outshootoutshotby",         # record when out-shooting or out-shot by the opponent
]

GAME_TYPES = {2: "regular", 3: "playoffs"}

URL = "https://api.nhle.com/stats/rest/en/team/{report}"

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
SAVE_FOLDER = PROJECT_FOLDER / "data" / "nhl" / "team_season"      # "data" is ignored by git
LOG_FILE = SAVE_FOLDER / "pull_log.csv"
LOG_COLUMNS = ["season_id", "report", "game_type", "rows", "columns", "status", "pulled_at"]


def current_start_year():
    """The season that is under way now (seasons start in the autumn)."""
    today = datetime.now()
    return today.year if today.month >= 9 else today.year - 1


def download(report, season_id, game_type):
    """
    One report, one season, one game type -> list of rows (one per team), or None if the
    download failed (so we know to try again later, and don't confuse it with "no such data").
    """
    params = {
        "isAggregate": "false",
        "isGame": "false",       # false = one row per team for the whole season
        "start": 0,
        "limit": -1,             # -1 = every row
        "cayenneExp": f"gameTypeId={game_type} and seasonId={season_id}",
    }
    raw = nhl_api.get_json(URL.format(report=report), params=params)
    if raw is None:
        return None
    return raw.get("data", [])


def pull_one_report(report, game_type, this_year, log, force):
    """Bring one report's stacked file up to date. Returns (new log rows, number of seasons asked)."""
    name = GAME_TYPES[game_type]
    path = SAVE_FOLDER / f"{report}_{name}.csv"
    stored = pd.read_csv(path, encoding="utf-8-sig") if path.exists() else pd.DataFrame()
    new_log, kept, asked = [], [], 0

    for start_year in range(FIRST_START_YEAR, this_year + 1):
        season_id = int(f"{start_year}{start_year + 1}")
        if game_type == 3 and start_year == this_year:
            continue                                  # this season's playoffs have not happened yet
        done = ((log["season_id"] == season_id) & (log["report"] == report) &
                (log["game_type"] == game_type)).any()
        if done and not force and start_year < this_year:
            continue                                  # a finished season we already have

        rows = download(report, season_id, game_type)
        asked += 1
        if rows is None:
            continue                                  # download failed: try again next run
        new_log.append({"season_id": season_id, "report": report, "game_type": game_type,
                        "rows": len(rows), "columns": len(rows[0]) if rows else 0,
                        "status": "ok" if rows else "empty",
                        "pulled_at": datetime.now().strftime("%Y-%m-%d %H:%M")})
        if rows:
            fresh = pd.DataFrame(rows)
            fresh["gameTypeId"] = game_type
            kept.append(fresh)
        if not stored.empty:                          # drop the old copy of a season we just re-pulled
            stored = stored[stored["seasonId"] != season_id]

    if kept:
        combined = pd.concat([stored] + kept, ignore_index=True) if not stored.empty \
            else pd.concat(kept, ignore_index=True)
        combined = combined.sort_values(["seasonId", "teamFullName"]).reset_index(drop=True)
        combined.to_csv(path, index=False, encoding="utf-8-sig")
    return new_log, asked


def main():
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Download NHL team season reports.")
    parser.add_argument("--only", help="comma-separated report names to pull")
    parser.add_argument("--force", action="store_true", help="ask again even for seasons we already have")
    args = parser.parse_args()

    reports = [r.strip() for r in args.only.split(",")] if args.only else REPORTS
    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)
    this_year = current_start_year()
    log = pd.read_csv(LOG_FILE) if LOG_FILE.exists() else pd.DataFrame(columns=LOG_COLUMNS)

    print(f"Pulling {len(reports)} reports for 2005-06 to {this_year}-{str(this_year + 1)[2:]} "
          f"(regular season and playoffs).")
    for report in reports:
        for game_type, name in GAME_TYPES.items():
            new_log, asked = pull_one_report(report, game_type, this_year, log, args.force)
            if new_log:
                fresh = pd.DataFrame(new_log)
                keys = fresh[["season_id", "report", "game_type"]]
                old = log.merge(keys, on=["season_id", "report", "game_type"], how="left", indicator=True)
                old = old[old["_merge"] == "left_only"].drop(columns="_merge")
                log = pd.concat([old, fresh], ignore_index=True)
                log.sort_values(["report", "game_type", "season_id"]).to_csv(LOG_FILE, index=False)
            ok = sum(1 for r in new_log if r["status"] == "ok")
            empty = sum(1 for r in new_log if r["status"] == "empty")
            if asked == 0:
                print(f"  {report:<24} {name:<9} already have it")
            else:
                print(f"  {report:<24} {name:<9} asked {asked} seasons: {ok} with data, {empty} empty"
                      + ("" if ok + empty == asked else f", {asked - ok - empty} FAILED (run again)"))

    print()
    print("Every report has a row per team for every season, BUT old seasons may have empty columns.")
    print("To see which statistic is really filled in from which season, run:")
    print("    python scripts/build_coverage_report.py")
    print()
    print(f"Files are in {SAVE_FOLDER}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

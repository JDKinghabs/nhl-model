"""
pull_game_results.py  -  every team's result in every game, 2005-06 to now (NHL stats site)
==========================================================================================
WHAT IT DOES (in Excel terms):
    The team season tables we already have are one row per team per SEASON. This script gets the
    same kind of numbers one row per team per GAME (about 2,500 rows a season: every game appears
    twice, once from each team's side). With those rows we can add up "everything through
    Thanksgiving" or "everything through the trade deadline" for any team, which is how the
    checkpoint snapshots work.

    Each report is one table per season (summary = score, shots, points; powerplay; penaltykill;
    goals by strength; shot attempts; hits and giveaways). Seasons are stacked into one file per
    report, like pasting each season's sheet underneath the last.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/pull_game_results.py

    About 300 requests, one per second, so around 5-6 minutes the first time. It is safe to stop
    and run again: finished seasons are skipped, and the current season is refreshed every time.
    Optional:  --only summary,realtime     --force (ask again for everything)

WHAT YOU SHOULD SEE:
    One line per report and game type, then the files in  data/nhl/team_games/ :
    summary_regular.csv, summary_playoffs.csv, powerplay_regular.csv ... and pull_log.csv

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

# Old seasons have empty columns, so stacking tables makes pandas warn. Harmless, so stay quiet.
warnings.filterwarnings("ignore", category=FutureWarning, message=".*concatenation.*")

# ----------------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------------
FIRST_START_YEAR = 2005          # 2005-06: first season after the lockout, start of the salary cap

REPORTS = [
    "summary",                   # goals, shots, points, faceoff %, PP% and PK% for that game
    "powerplay",                 # power-play goals and opportunities in that game
    "penaltykill",               # times short-handed and goals against in that game
    "goalsforbystrength",        # goals for at 5v5, 5v4 ... in that game
    "goalsagainstbystrength",    # goals against at 5v5, 5v4 ... in that game
    "summaryshooting",           # shot attempts for/against in that game
    "realtime",                  # hits, blocks, giveaways, takeaways in that game
]
GAME_TYPES = {2: "regular", 3: "playoffs"}

URL = "https://api.nhle.com/stats/rest/en/team/{report}"

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
SAVE_FOLDER = PROJECT_FOLDER / "data" / "nhl" / "team_games"      # "data" is ignored by git
LOG_FILE = SAVE_FOLDER / "pull_log.csv"
LOG_COLUMNS = ["season_id", "report", "game_type", "rows", "status", "pulled_at"]


def current_start_year():
    """The season that is under way now (seasons start in the autumn)."""
    today = datetime.now()
    return today.year if today.month >= 9 else today.year - 1


def download(report, season_id, game_type):
    """One report, one season, one game type, one row per team per game. None = download failed."""
    params = {
        "isAggregate": "false",
        "isGame": "true",        # true = one row per team per GAME (the season tables used false)
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
        done = ((log["season_id"] == season_id) & (log["report"] == report) &
                (log["game_type"] == game_type)).any()
        if done and not force and start_year < this_year:
            continue                                  # a finished season we already have

        rows = download(report, season_id, game_type)
        asked += 1
        if rows is None:
            continue                                  # download failed: try again next run
        new_log.append({"season_id": season_id, "report": report, "game_type": game_type,
                        "rows": len(rows), "status": "ok" if rows else "empty",
                        "pulled_at": datetime.now().strftime("%Y-%m-%d %H:%M")})
        if rows:
            fresh = pd.DataFrame(rows)
            fresh["seasonId"] = season_id             # the game rows do not say which season they are
            fresh["gameTypeId"] = game_type
            kept.append(fresh)
        if not stored.empty:                          # drop the old copy of a season we just re-pulled
            stored = stored[stored["seasonId"] != season_id]

    if kept:
        combined = pd.concat(([stored] if not stored.empty else []) + kept, ignore_index=True)
        combined = combined.sort_values(["seasonId", "gameDate", "gameId", "teamId"]).reset_index(drop=True)
        combined.to_csv(path, index=False, encoding="utf-8-sig")
    return new_log, asked


def main():
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Download NHL team results, one row per team per game.")
    parser.add_argument("--only", help="comma-separated report names to pull")
    parser.add_argument("--force", action="store_true", help="ask again even for seasons we already have")
    args = parser.parse_args()

    reports = [r.strip() for r in args.only.split(",")] if args.only else REPORTS
    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)
    this_year = current_start_year()
    log = pd.read_csv(LOG_FILE) if LOG_FILE.exists() else pd.DataFrame(columns=LOG_COLUMNS)

    print(f"Pulling {len(reports)} game-by-game reports for 2005-06 to {this_year}-{str(this_year + 1)[2:]}.")
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
            rows = sum(r["rows"] for r in new_log)
            if asked == 0:
                print(f"  {report:<24} {name:<9} already have it")
            else:
                print(f"  {report:<24} {name:<9} asked {asked} seasons: {ok} with data ({rows:,} rows), {empty} empty"
                      + ("" if ok + empty == asked else f", {asked - ok - empty} FAILED (run again)"))

    print()
    print(f"Files are in {SAVE_FOLDER}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""
pull_play_by_play.py  -  every event of every NHL game (the overnight download)
==============================================================================
WHAT IT DOES (in Excel terms):
    The NHL publishes a "play-by-play" for every game: one row per event (faceoff, shot, hit, giveaway,
    takeaway, penalty, goal ...) with the time, the team, the players involved and, for most events,
    WHERE on the ice it happened (x, y) and which zone (offensive, neutral, defensive).
    This script downloads that for every game since 2009-10 and stores it in one database file,
    data/pbp.db (about 1-2 GB). Later scripts use it to estimate zone time, breakouts and shot locations.

    Roughly 27,000 games, one request per second, so about 8 hours. It is built to run overnight:
      * RESUMABLE: games already saved are skipped, so you can stop and start it any time.
      * POLITE: at least one second between requests; if the NHL starts refusing, it backs off
        (waits 5 minutes) and gives up after three tries instead of hammering the server.
      * ORDER: the most recent seasons first (so the checks against the NHL's real zone-time data
        for 2021-22 on can start early), then older ones, then the two odd seasons (2019-20, 2020-21),
        and the current season last.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/pull_play_by_play.py                 the whole thing (leave it running overnight)
    python scripts/pull_play_by_play.py --limit 50      just 50 games (a test)
    python scripts/pull_play_by_play.py --status        how far along is it? (safe to run any time)

    TO STOP CLEANLY: create an empty file named STOP inside the folder  data/pbp/  (the script notices it
    within a second, finishes the game it is on, and exits). Or press Ctrl+C. Nothing is lost either way.

WHAT YOU SHOULD SEE:
    A progress line every 100 games (how many done, speed, estimated time left), also saved in
    data/pbp/progress.log. At the end, a summary.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. The interface is unofficial and can change.
"""

import argparse
import sqlite3
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

import nhl_api          # our shared, polite downloader (scripts/nhl_api.py)

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
NHL_DB = PROJECT_FOLDER / "data" / "nhl.db"                 # has the list of games (from build_database.py)
PBP_FOLDER = PROJECT_FOLDER / "data" / "pbp"               # "data" is ignored by git
PBP_DB = PROJECT_FOLDER / "data" / "pbp.db"
STOP_FILE = PBP_FOLDER / "STOP"
PROGRESS_LOG = PBP_FOLDER / "progress.log"
FLAGS_FILE = PROJECT_FOLDER / "reference" / "season_flags.csv"

URL = "https://api-web.nhle.com/v1/gamecenter/{game_id}/play-by-play"
FIRST_SEASON = 20092010            # coordinates and zone tags exist from 2009-10
PROGRESS_EVERY = 100
PAUSE_AFTER_ERRORS = 300           # seconds to wait after a run of failures
ERRORS_BEFORE_PAUSE = 10
PAUSES_BEFORE_GIVING_UP = 3

# How the NHL's field names map to our column names
PLAY_FIELDS = ["eventId", "sortOrder", "situationCode", "homeTeamDefendingSide", "typeDescKey"]
DETAIL_FIELDS = {
    "eventOwnerTeamId": "team_id", "xCoord": "x", "yCoord": "y", "zoneCode": "zone", "shotType": "shot_type",
    "shootingPlayerId": "shooter_id", "goalieInNetId": "goalie_id", "blockingPlayerId": "blocker_id",
    "winningPlayerId": "winner_id", "losingPlayerId": "loser_id", "hittingPlayerId": "hitter_id",
    "hitteePlayerId": "hittee_id", "playerId": "player_id", "reason": "reason", "secondaryReason": "secondary_reason",
    "typeCode": "pen_type", "descKey": "pen_desc", "duration": "pen_minutes", "committedByPlayerId": "committed_by",
    "drawnByPlayerId": "drawn_by", "servedByPlayerId": "served_by", "scoringPlayerId": "scorer_id",
    "assist1PlayerId": "assist1_id", "assist2PlayerId": "assist2_id", "awaySOG": "away_sog", "homeSOG": "home_sog",
    "awayScore": "away_score", "homeScore": "home_score",
}
EVENT_COLUMNS = (["game_id", "event_id", "sort_order", "period", "period_type", "time_s", "situation_code",
                  "home_def_side", "type"] + list(DETAIL_FIELDS.values()))
ROSTER_COLUMNS = ["game_id", "team_id", "player_id", "first_name", "last_name", "sweater", "position"]

SCHEMA = f"""
CREATE TABLE IF NOT EXISTS pbp_games (
    game_id INTEGER PRIMARY KEY, season_id INTEGER, game_type INTEGER, game_date TEXT,
    home_team_id INTEGER, away_team_id INTEGER, home_score INTEGER, away_score INTEGER,
    status TEXT, n_events INTEGER, pulled_at TEXT);
CREATE TABLE IF NOT EXISTS pbp_events ({', '.join(c + (' TEXT' if c in ('period_type', 'situation_code', 'home_def_side', 'type', 'zone', 'shot_type', 'reason', 'secondary_reason', 'pen_type', 'pen_desc') else ' INTEGER' if c not in ('x', 'y') else ' REAL') for c in EVENT_COLUMNS)});
CREATE TABLE IF NOT EXISTS pbp_rosters (game_id INTEGER, team_id INTEGER, player_id INTEGER, first_name TEXT, last_name TEXT, sweater INTEGER, position TEXT);
CREATE INDEX IF NOT EXISTS ix_pbp_events_game ON pbp_events (game_id, sort_order);
CREATE INDEX IF NOT EXISTS ix_pbp_rosters_game ON pbp_rosters (game_id);
"""


def clock_to_seconds(text):
    """'01:01' -> 61"""
    try:
        minutes, seconds = str(text).split(":")
        return int(minutes) * 60 + int(seconds)
    except (ValueError, AttributeError):
        return None


def flatten_game(game_id, data):
    """Turn one game's play-by-play into (event rows, roster rows, a summary row)."""
    events = []
    for play in data.get("plays", []):
        details = play.get("details") or {}
        period = play.get("periodDescriptor") or {}
        row = {"game_id": game_id, "event_id": play.get("eventId"), "sort_order": play.get("sortOrder"),
               "period": period.get("number"), "period_type": period.get("periodType"),
               "time_s": clock_to_seconds(play.get("timeInPeriod")), "situation_code": play.get("situationCode"),
               "home_def_side": play.get("homeTeamDefendingSide"), "type": play.get("typeDescKey")}
        for source, name in DETAIL_FIELDS.items():
            row[name] = details.get(source)
        events.append(tuple(row[c] for c in EVENT_COLUMNS))

    rosters = []
    for spot in data.get("rosterSpots", []):
        rosters.append((game_id, spot.get("teamId"), spot.get("playerId"), (spot.get("firstName") or {}).get("default"),
                        (spot.get("lastName") or {}).get("default"), spot.get("sweaterNumber"), spot.get("positionCode")))

    summary = {"home_score": (data.get("homeTeam") or {}).get("score"), "away_score": (data.get("awayTeam") or {}).get("score"),
               "n_events": len(events)}
    return events, rosters, summary


def season_order():
    """Season ids in the order we want to download them."""
    connection = sqlite3.connect(NHL_DB)
    seasons = [r[0] for r in connection.execute("SELECT DISTINCT season_id FROM games WHERE season_id >= ? ORDER BY season_id DESC", (FIRST_SEASON,))]
    connection.close()
    flagged = set()
    if FLAGS_FILE.exists():
        flags = pd.read_csv(FLAGS_FILE)
        flagged = set(int(s) for s in flags.loc[flags["include_main"] == "N", "season_id"])
    current = max(seasons) if seasons else None
    normal = [s for s in seasons if s not in flagged and s != current]
    odd = [s for s in seasons if s in flagged]
    return normal + odd + ([current] if current in seasons and current not in flagged else [])


def game_queue(done_ids, season_filter=None):
    """Games still to download, in download order: per season, regular season then playoffs, by date."""
    connection = sqlite3.connect(NHL_DB)
    games = pd.read_sql_query(
        "SELECT game_id, season_id, game_type, game_date, home_team_id, away_team_id FROM games WHERE season_id >= ?",
        connection, params=[FIRST_SEASON])
    connection.close()
    rank = {season: i for i, season in enumerate(season_order())}
    games["order"] = games["season_id"].map(rank)
    games = games.sort_values(["order", "game_type", "game_date", "game_id"])
    if season_filter:
        games = games[games["season_id"] == int(season_filter)]
    return games[~games["game_id"].isin(done_ids)]


def open_db():
    PBP_FOLDER.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(PBP_DB, timeout=60)
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    connection.executescript(SCHEMA)
    return connection


def log(message):
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {message}"
    print(line, flush=True)
    try:
        with open(PROGRESS_LOG, "a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        pass


def save_game(connection, game, data):
    events, rosters, summary = flatten_game(int(game["game_id"]), data)
    with connection:
        connection.execute("DELETE FROM pbp_events WHERE game_id = ?", (int(game["game_id"]),))
        connection.execute("DELETE FROM pbp_rosters WHERE game_id = ?", (int(game["game_id"]),))
        connection.executemany(f"INSERT INTO pbp_events VALUES ({','.join('?' * len(EVENT_COLUMNS))})", events)
        connection.executemany(f"INSERT INTO pbp_rosters VALUES ({','.join('?' * len(ROSTER_COLUMNS))})", rosters)
        connection.execute(
            "INSERT OR REPLACE INTO pbp_games VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (int(game["game_id"]), int(game["season_id"]), int(game["game_type"]), game["game_date"], int(game["home_team_id"]),
             int(game["away_team_id"]), summary["home_score"], summary["away_score"], "ok", summary["n_events"],
             datetime.now().strftime("%Y-%m-%d %H:%M:%S")))


def mark_missing(connection, game, status):
    with connection:
        connection.execute(
            "INSERT OR REPLACE INTO pbp_games VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (int(game["game_id"]), int(game["season_id"]), int(game["game_type"]), game["game_date"], int(game["home_team_id"]),
             int(game["away_team_id"]), None, None, status, 0, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))


# ----------------------------------------------------------------------------
def show_status():
    if not PBP_DB.exists():
        print("Nothing downloaded yet (data/pbp.db does not exist).")
        return 0
    connection = open_db()
    total = len(game_queue(set()))
    done = pd.read_sql_query("SELECT status, COUNT(*) AS n FROM pbp_games GROUP BY status", connection)
    have = int(done["n"].sum()) if len(done) else 0
    print(f"Games saved: {have:,} of {total:,} in the plan ({have / total:.1%}).")
    if len(done):
        print("  " + ", ".join(f"{r.status}: {r.n:,}" for r in done.itertuples()))
    recent = pd.read_sql_query(
        "SELECT pulled_at FROM pbp_games WHERE pulled_at >= ? ORDER BY pulled_at",
        connection, params=[(datetime.now() - timedelta(minutes=10)).strftime("%Y-%m-%d %H:%M:%S")])
    if len(recent) >= 2:
        span = (pd.to_datetime(recent["pulled_at"].iloc[-1]) - pd.to_datetime(recent["pulled_at"].iloc[0])).total_seconds()
        rate = (len(recent) - 1) / span if span > 0 else 0
        remaining = total - have
        if rate > 0:
            hours = remaining / rate / 3600
            print(f"  Speed in the last 10 minutes: {rate * 60:.0f} games a minute; about {hours:.1f} hours left "
                  f"(finishing around {datetime.now() + timedelta(hours=hours):%a %I:%M %p}).")
    else:
        print("  Not running right now (nothing saved in the last 10 minutes).")
    by_season = pd.read_sql_query(
        "SELECT season_id, COUNT(*) AS saved, SUM(n_events) AS events FROM pbp_games WHERE status = 'ok' GROUP BY season_id ORDER BY season_id DESC", connection)
    if len(by_season):
        print("  Saved per season (games / events):")
        for r in by_season.itertuples():
            print(f"    {str(r.season_id)[:4]}-{str(r.season_id)[6:]}: {r.saved:,} games, {int(r.events):,} events")
    size = PBP_DB.stat().st_size / 1e9
    print(f"  Database size: {size:.2f} GB")
    connection.close()
    return 0


def stay_awake():
    """
    Ask Windows not to put the computer to sleep because it looks idle while this script runs (the same request a
    video player makes). Nothing in your settings changes, and the request ends when the script ends.
    It does NOT stop sleep caused by closing the laptop lid or choosing Sleep.
    """
    try:
        import ctypes
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
        return True
    except (AttributeError, OSError):
        return False                                  # not Windows: nothing to do


def main():
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Download NHL play-by-play for every game since 2009-10.")
    parser.add_argument("--limit", type=int, help="stop after this many games (a test)")
    parser.add_argument("--status", action="store_true", help="show progress and exit")
    parser.add_argument("--season", help="only this season id, e.g. 20232024")
    parser.add_argument("--allow-sleep", action="store_true", help="do not ask Windows to stay awake")
    args = parser.parse_args()
    if args.status:
        return show_status()
    awake = False if args.allow_sleep else stay_awake()

    if not NHL_DB.exists():
        print("I can't find the list of games yet. Run this first:  python scripts/build_database.py")
        return 1
    connection = open_db()
    done = {r[0] for r in connection.execute("SELECT game_id FROM pbp_games WHERE status IN ('ok', 'not_found')")}
    queue = game_queue(done, args.season)
    total_plan = len(queue) + len(done)
    if STOP_FILE.exists():
        STOP_FILE.unlink()                            # an old stop request should not stop a fresh start
    log(f"Starting: {len(done):,} games already saved, {len(queue):,} to go"
        + (f" (limited to {args.limit})" if args.limit else "") + f". About {len(queue) / 3600:.1f} hours at one game a second."
        + (" Asked Windows not to sleep while this runs." if awake else ""))

    started, saved, missing = time.time(), 0, 0
    consecutive_errors, pauses = 0, 0
    for game in queue.itertuples():
        if STOP_FILE.exists():
            log("STOP file found: stopping cleanly.")
            break
        if args.limit and saved + missing >= args.limit:
            break
        row = {"game_id": game.game_id, "season_id": game.season_id, "game_type": game.game_type, "game_date": game.game_date,
               "home_team_id": game.home_team_id, "away_team_id": game.away_team_id}
        data, status = nhl_api.get_json_with_status(URL.format(game_id=int(game.game_id)))
        if status == "ok":
            save_game(connection, row, data)
            saved += 1
            consecutive_errors = 0
        elif status == "not_found":
            mark_missing(connection, row, "not_found")
            missing += 1
            consecutive_errors = 0
        else:
            consecutive_errors += 1
            log(f"  Problem with game {int(game.game_id)}: {status}")
            if consecutive_errors >= ERRORS_BEFORE_PAUSE:
                pauses += 1
                if pauses > PAUSES_BEFORE_GIVING_UP:
                    log(f"Giving up after {pauses - 1} long pauses: the NHL keeps refusing. Run the script again later; it will resume.")
                    break
                log(f"{consecutive_errors} problems in a row. Waiting {PAUSE_AFTER_ERRORS // 60} minutes, then trying again (pause {pauses} of {PAUSES_BEFORE_GIVING_UP}).")
                time.sleep(PAUSE_AFTER_ERRORS)
                consecutive_errors = 0
            continue
        finished = saved + missing
        if finished % PROGRESS_EVERY == 0:
            elapsed = time.time() - started
            rate = finished / elapsed
            left = len(queue) - finished if not args.limit else args.limit - finished
            log(f"{len(done) + finished:,} of {total_plan:,} games saved ({(len(done) + finished) / total_plan:.1%}); "
                f"{rate * 60:.0f} a minute; about {left / rate / 3600:.1f} hours left; now on season {str(game.season_id)[:4]}-{str(game.season_id)[6:]}")

    connection.close()
    log(f"Finished this run: {saved:,} games saved, {missing:,} not available, in {(time.time() - started) / 60:.1f} minutes.")
    print("Run  python scripts/pull_play_by_play.py --status  any time to see the total.")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

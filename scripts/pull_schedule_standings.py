"""
pull_schedule_standings.py  -  tonight's NHL games and the current standings
============================================================================
WHAT IT DOES (in Excel terms):
    Fetches two "tables" from the NHL and saves each as a CSV you can open in Excel:
      1. Tonight's games (or the next day that has games)  ->  data/schedule_today.csv
      2. The current standings, one row per team           ->  data/standings.csv

HOW TO RUN IT (in the VS Code terminal, from the "Model 2.0" folder):
    python scripts/pull_schedule_standings.py

WHAT YOU SHOULD SEE:
    A summary of tonight's games, the top 5 teams, and two files in the data folder.
    (If a CSV is open in Excel, close it first. Excel locks open files.)

DATA CREDIT:
    Data is (c) NHL, used for personal, non-commercial purposes. This is an unofficial
    interface that can change without notice, so problems are explained in plain English.
"""

import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd      # pandas = "Excel for Python": a DataFrame is a worksheet
import requests          # requests = fetches things from the internet

try:
    from zoneinfo import ZoneInfo                    # knows about daylight saving time
    EASTERN = ZoneInfo("America/New_York")
except Exception:                                    # if the time-zone list is missing, use a fixed offset
    EASTERN = timezone(timedelta(hours=-5))

# ----------------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------------
WEB_API = "https://api-web.nhle.com/v1"
PROJECT_FOLDER = Path(__file__).resolve().parent.parent   # the "Model 2.0" folder
DATA_FOLDER = PROJECT_FOLDER / "data"                     # ignored by git: never goes to GitHub
HEADERS = {"User-Agent": "nhl-model-hobby-project (personal, non-commercial)"}

# What the NHL calls a game's status, in words a person would use.
STATUS_WORDS = {"FUT": "Scheduled", "PRE": "Scheduled", "LIVE": "In progress", "CRIT": "In progress",
                "FINAL": "Final", "OFF": "Final"}


# ----------------------------------------------------------------------------
# A "function" is a named recipe, like a custom Excel formula.
# ----------------------------------------------------------------------------
def get_json(url):
    """Download a web address that returns data. Tries twice, then explains what went wrong."""
    for attempt in (1, 2):
        try:
            response = requests.get(url, headers=HEADERS, timeout=60)
            response.raise_for_status()
            return response.json()
        except (requests.RequestException, ValueError) as problem:
            if attempt == 2:
                print(f"  Could not get {url}")
                print(f"  Reason: {problem}")
                print("  Check your internet connection. The NHL may also have changed this address.")
                return None
            time.sleep(2)


def build_standings(raw):
    """Turn the NHL's standings into one tidy row per team."""
    rows = []
    for t in raw.get("standings", []):
        # t is one team. Notice t["teamAbbrev"]["default"]: a value inside a value, like a cell
        # that holds a small table. "default" is the English version of the name.
        gp = t["gamesPlayed"]
        rows.append({
            "league_rank": t["leagueSequence"],
            "team": t["teamAbbrev"]["default"],
            "team_name": t["teamName"]["default"],
            "conference": t["conferenceName"],
            "division": t["divisionName"],
            "conf_rank": t["conferenceSequence"],
            "div_rank": t["divisionSequence"],
            "gp": gp,
            "w": t["wins"],
            "l": t["losses"],
            "otl": t["otLosses"],
            "pts": t["points"],
            "pts_pct": round(t["points"] / (2 * gp), 3) if gp else 0,   # points won / points available
            "rw": t["regulationWins"],
            "row": t["regulationPlusOtWins"],            # "ROW": regulation + overtime wins
            "gf": t["goalFor"],
            "ga": t["goalAgainst"],
            "gd": t["goalDifferential"],
            "home": f"{t['homeWins']}-{t['homeLosses']}-{t['homeOtLosses']}",
            "road": f"{t['roadWins']}-{t['roadLosses']}-{t['roadOtLosses']}",
            "l10": f"{t['l10Wins']}-{t['l10Losses']}-{t['l10OtLosses']}",
            "streak": f"{t['streakCode']}{t['streakCount']}" if t.get("streakCode") else "",
            "as_of": t["date"],
        })
    return pd.DataFrame(rows).sort_values("league_rank")


def build_schedule(raw, today_et):
    """Pick tonight's games (or the next day with games) and tidy them up."""
    days_with_games = [d for d in raw.get("gameWeek", []) if d.get("games")]
    chosen = next((d for d in days_with_games if d["date"] >= today_et), None)
    if chosen is None:
        return pd.DataFrame()

    rows = []
    for g in chosen["games"]:
        start_utc = datetime.fromisoformat(g["startTimeUTC"].replace("Z", "+00:00"))
        start_et = start_utc.astimezone(EASTERN)
        state = g.get("gameState", "")
        started = state not in ("FUT", "PRE")
        rows.append({
            "game_id": g["id"],
            "game_date": chosen["date"],
            "is_today": chosen["date"] == today_et,
            "start_utc": g["startTimeUTC"],
            "start_et": start_et.strftime("%I:%M %p").lstrip("0") + " ET",
            "away": g["awayTeam"]["abbrev"],
            "home": g["homeTeam"]["abbrev"],
            "venue": g.get("venue", {}).get("default", ""),
            "status": STATUS_WORDS.get(state, state),
            "away_score": g["awayTeam"].get("score") if started else None,
            "home_score": g["homeTeam"].get("score") if started else None,
        })
    return pd.DataFrame(rows).sort_values("start_utc")


def save_csv(sheet, name):
    """Save a worksheet as a CSV. 'utf-8-sig' makes Excel show accented letters correctly."""
    DATA_FOLDER.mkdir(parents=True, exist_ok=True)
    path = DATA_FOLDER / name
    try:
        sheet.to_csv(path, index=False, encoding="utf-8-sig")
    except PermissionError:
        print(f"  Could not save {name}. Is it open in Excel? Close it and run again.")
        return False
    print(f"  Saved: {path}  ({len(sheet)} rows)")
    return True


# ----------------------------------------------------------------------------
# main() is where the script starts, like the first line of a macro.
# ----------------------------------------------------------------------------
def main():
    sys.stdout.reconfigure(errors="replace")           # never crash on an unusual character
    today_et = datetime.now(EASTERN).strftime("%Y-%m-%d")
    print(f"Today (Eastern time): {today_et}")

    print("Getting the standings...")
    raw_standings = get_json(f"{WEB_API}/standings/now")
    if raw_standings is None:
        return 1
    time.sleep(1)                                      # be polite: pause between requests

    print("Getting the schedule...")
    raw_schedule = get_json(f"{WEB_API}/schedule/{today_et}")
    if raw_schedule is None:
        return 1

    try:
        standings = build_standings(raw_standings)
    except KeyError as missing:
        print(f"  The NHL standings no longer contain the field {missing}. They may have changed the format.")
        return 1
    schedule = build_schedule(raw_schedule, today_et)

    ok = save_csv(standings, "standings.csv")
    if schedule.empty:
        print("  No games found this week (off-season?). Saving an empty schedule.")
    ok = save_csv(schedule, "schedule_today.csv") and ok

    # A quick look, so you can compare with nhl.com.
    if not schedule.empty:
        label = "Tonight" if schedule["is_today"].iloc[0] else f"Next games ({schedule['game_date'].iloc[0]})"
        print(f"\n{label}: {len(schedule)} game(s)")
        for _, game in schedule.iterrows():
            score = "" if pd.isna(game["away_score"]) else f"  [{int(game['away_score'])}-{int(game['home_score'])}]"
            print(f"  {game['away']} at {game['home']}  {game['start_et']}  {game['status']}{score}")
    print("\nTop 5 in the standings:")
    for _, team in standings.head(5).iterrows():
        print(f"  {team['team_name']:<24} {team['w']}-{team['l']}-{team['otl']}  {team['pts']} pts in {team['gp']} games")
    print("\nData: (c) NHL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

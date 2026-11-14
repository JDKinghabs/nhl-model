"""
build_checkpoints.py  -  "where did every team stand on Nov 1, Thanksgiving, Jan 1 ... ?"
==========================================================================================
WHAT IT DOES (in Excel terms):
    1. Works out the CHECKPOINT DATES for every season: Nov 1, US Thanksgiving, Jan 1, the trade
       deadline and the last day of the regular season.
         * Normal seasons use the real calendar date.
         * Seasons that started late (2012-13 after the lockout, 2020-21) have no Nov 1 or Jan 1,
           so each checkpoint is placed at the same SHARE of the season instead (for example "when
           the typical team had played 15% of its games").
    2. Asks the NHL "what did the standings look like on that date?" for each checkpoint, and
       keeps one row per team per checkpoint: games played, points, rank in the conference and
       division, and whether the team was IN a playoff spot at that moment.
    3. Checks the idea against reality: at the end of each season, "in a playoff spot" should match
       "actually made the playoffs".

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/pull_game_results.py        (first: it provides the game dates this needs)
    python scripts/build_checkpoints.py

    About 100 requests (a couple of minutes). Finished seasons are remembered; the current season
    is refreshed each time. Option:  --force  (ask again for everything)

WHAT YOU SHOULD SEE:
    The target shares, a check of the standings dates, a check of "in a spot" against reality, and
    two new files:  data/reference/checkpoints.csv  and  data/nhl/standings_snapshots.csv

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. Unofficial, can change.
"""

import argparse
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

import nhl_api          # our shared, polite downloader (scripts/nhl_api.py)

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
GAMES_FILE = PROJECT_FOLDER / "data" / "nhl" / "team_games" / "summary_regular.csv"
DEADLINES_FILE = PROJECT_FOLDER / "reference" / "trade_deadlines.csv"
FLAGS_FILE = PROJECT_FOLDER / "reference" / "season_flags.csv"
TEAM_SEASONS_FILE = PROJECT_FOLDER / "data" / "reference" / "team_seasons.csv"
CHECKPOINTS_FILE = PROJECT_FOLDER / "data" / "reference" / "checkpoints.csv"
SNAPSHOTS_FILE = PROJECT_FOLDER / "data" / "nhl" / "standings_snapshots.csv"

STANDINGS_URL = "https://api-web.nhle.com/v1/standings/{day}"

CHECKPOINT_ORDER = ["nov1", "thanksgiving", "jan1", "deadline", "season_end"]
SHARE_BASED = ["nov1", "thanksgiving", "jan1"]      # the ones that can move to a share of the season
PLAIN_NAMES = {"nov1": "Nov 1", "thanksgiving": "Thanksgiving", "jan1": "Jan 1",
               "deadline": "Trade deadline", "season_end": "End of regular season"}


def current_start_year():
    today = datetime.now()
    return today.year if today.month >= 9 else today.year - 1


def us_thanksgiving(year):
    """The fourth Thursday of November."""
    first = date(year, 11, 1)
    first_thursday = first + timedelta(days=(3 - first.weekday()) % 7)
    return first_thursday + timedelta(days=21)


# ----------------------------------------------------------------------------
# Game data: who played when, and each team's running games-played count
# ----------------------------------------------------------------------------
def load_games():
    if not GAMES_FILE.exists():
        print("I need the game-by-game results first. Run this, then run me again:")
        print("    python scripts/pull_game_results.py")
        return None
    games = pd.read_csv(GAMES_FILE, encoding="utf-8-sig")
    games = games[games["teamId"].notna()].copy()
    games["teamId"] = games["teamId"].astype(int)
    games["gameDate"] = pd.to_datetime(games["gameDate"]).dt.date
    return games


def team_abbreviations(games):
    """
    The standings list teams by abbreviation, not id. Each game row tells us the OPPONENT's
    abbreviation, so look at the other team's row of the same game to learn each team's own.
    Returns {(season_id, abbreviation): team_id}.
    """
    pairs = games[["seasonId", "gameId", "teamId", "opponentTeamAbbrev"]]
    other = pairs.merge(pairs, on=["seasonId", "gameId"], suffixes=("", "_other"))
    other = other[other["teamId"] != other["teamId_other"]]
    # row of team A says opponent = "X"; the other row is team B whose id belongs to abbreviation X
    table = other[["seasonId", "opponentTeamAbbrev", "teamId_other"]].drop_duplicates()
    return {(int(r.seasonId), r.opponentTeamAbbrev): int(r.teamId_other) for r in table.itertuples()}


def median_games_played(season_games):
    """A table: for each game date, the typical (median) number of games played by a team so far."""
    counts = season_games.groupby(["gameDate", "teamId"]).size().unstack(fill_value=0).cumsum()
    return counts.median(axis=1)


def median_at(series, day):
    """The median games played as of the end of `day` (works for any calendar date)."""
    earlier = series[series.index <= day]
    return float(earlier.iloc[-1]) if len(earlier) else 0.0


# ----------------------------------------------------------------------------
# Checkpoint dates
# ----------------------------------------------------------------------------
def build_checkpoints(games, deadlines, today):
    """One row per season per checkpoint: the date and how we picked it."""
    this_year = current_start_year()
    seasons = sorted(games["seasonId"].unique())
    info = {}
    for season in seasons:
        sg = games[games["seasonId"] == season]
        start_year = int(str(season)[:4])
        first, last = sg["gameDate"].min(), sg["gameDate"].max()
        total = int(sg.groupby("teamId").size().median())
        info[season] = {"start_year": start_year, "first": first, "last": last, "total_games": total,
                        "late_start": first > date(start_year, 10, 31), "median": median_games_played(sg)}

    # What share of a normal 82-game season has been played by each calendar checkpoint?
    normal = [s for s in seasons if not info[s]["late_start"] and info[s]["total_games"] >= 80
              and info[s]["start_year"] < this_year]
    calendar_day = {
        "nov1": lambda y: date(y, 11, 1),
        "thanksgiving": lambda y: us_thanksgiving(y),
        "jan1": lambda y: date(y + 1, 1, 1),
    }
    target_share = {}
    for name, day_for in calendar_day.items():
        shares = [median_at(info[s]["median"], day_for(info[s]["start_year"])) / info[s]["total_games"]
                  for s in normal]
        target_share[name] = float(pd.Series(shares).median())

    rows = []
    for season in seasons:
        s = info[season]
        y = s["start_year"]
        for name in CHECKPOINT_ORDER:
            when, method, note = None, "", ""
            if name in calendar_day and not s["late_start"]:
                when, method = calendar_day[name](y), "calendar date"
            elif name in calendar_day:
                need = target_share[name] * s["total_games"]
                med = s["median"]
                hit = med[med >= need]
                if len(hit):
                    when = hit.index[0]
                    method = "share of season"
                    note = (f"Season started {s['first']}, so no {PLAIN_NAMES[name]}. Placed on the first day the typical "
                            f"team had played {target_share[name]:.0%} of its {s['total_games']} games.")
                else:
                    method, note = "share of season", "Not reached yet"
            elif name == "deadline":
                row = deadlines[deadlines["season_id"] == season]
                if len(row) and pd.notna(row.iloc[0]["deadline_date"]) and row.iloc[0]["deadline_date"]:
                    when = pd.to_datetime(row.iloc[0]["deadline_date"]).date()
                    method = "calendar date"
                    note = "" if str(row.iloc[0]["status"]).startswith("confirmed") else "deadline date from memory: check"
                else:
                    method, note = "calendar date", "Deadline not announced yet"
            elif name == "season_end":
                if y < this_year:
                    when, method = s["last"], "last game of the regular season"
                else:
                    method, note = "last game of the regular season", "Season in progress"
            reached = when is not None and when <= today
            rows.append({"season_id": season, "checkpoint": name, "date": when, "method": method,
                         "reached": reached, "note": note,
                         "games_played_typical": median_at(s["median"], when) if reached else None,
                         "season_games": s["total_games"]})
    return pd.DataFrame(rows), target_share


# ----------------------------------------------------------------------------
# Standings snapshots
# ----------------------------------------------------------------------------
KEEP = {  # name in the NHL data -> name we use
    "conferenceName": "conference", "divisionName": "division",
    "leagueSequence": "league_seq", "conferenceSequence": "conf_seq", "divisionSequence": "div_seq",
    "wildcardSequence": "wildcard_seq", "gamesPlayed": "gp", "wins": "w", "losses": "l", "otLosses": "otl",
    "ties": "ties", "points": "pts", "pointPctg": "pts_pct", "regulationWins": "reg_wins",
    "regulationPlusOtWins": "row", "goalFor": "gf", "goalAgainst": "ga", "goalDifferential": "gd",
    "homePoints": "home_pts", "roadPoints": "road_pts", "l10GamesPlayed": "l10_gp", "l10Points": "l10_pts",
}


def fetch_snapshot(season, checkpoint, day, abbrev_to_id):
    raw = nhl_api.get_json(STANDINGS_URL.format(day=day), quiet_404=True)
    if raw is None:
        return None
    rows = []
    for t in raw.get("standings", []):
        abbrev = (t.get("teamAbbrev") or {}).get("default")
        row = {"season_id": season, "checkpoint": checkpoint, "requested_date": str(day),
               "snapshot_date": t.get("date"), "team_abbrev": abbrev,
               "team": (t.get("teamName") or {}).get("default"),
               "team_id": abbrev_to_id.get((int(season), abbrev))}
        for old, new in KEEP.items():
            row[new] = t.get(old)
        rows.append(row)
    return rows


def spot_rule(season, checkpoint):
    if season == 20202021:
        return "top 4 in each division"
    if season == 20192020 and checkpoint == "season_end":
        return "top 12 in each conference (qualifying round)"
    if season >= 20132014:
        return "top 3 in each division + 2 wild cards per conference"
    return "top 8 in each conference (division leaders seeded 1-3)"


def mark_playoff_spots(snap):
    """Add in_playoff_spot (True/False) and the rule used, for every row of one snapshot."""
    rule = spot_rule(int(snap["season_id"].iloc[0]), snap["checkpoint"].iloc[0])
    spot = pd.Series(False, index=snap.index)
    if rule.startswith("top 4 in each division"):
        spot = snap["div_seq"] <= 4
    elif rule.startswith("top 12"):
        spot = snap["conf_seq"] <= 12
    elif rule.startswith("top 3"):
        spot = (snap["div_seq"] <= 3) | (snap["wildcard_seq"].isin([1, 2]))
    else:   # old format: division leaders are in, then the best of the rest fill the 8 places
        for _, conf in snap.groupby("conference"):
            leaders = conf[conf["div_seq"] == 1]
            others = conf[conf["div_seq"] != 1].sort_values("conf_seq")
            chosen = list(leaders.index) + list(others.index[: max(0, 8 - len(leaders))])
            spot.loc[chosen] = True
    snap = snap.copy()
    snap["in_playoff_spot"] = spot
    snap["spot_rule"] = rule
    return snap


def main():
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Build checkpoint dates and standings snapshots.")
    parser.add_argument("--force", action="store_true", help="ask again even for snapshots we already have")
    args = parser.parse_args()

    games = load_games()
    if games is None:
        return 1
    today = datetime.now().date()
    deadlines = pd.read_csv(DEADLINES_FILE, dtype={"season_id": int}).fillna("")
    checkpoints, target_share = build_checkpoints(games, deadlines, today)
    abbrev_to_id = team_abbreviations(games)

    print("Target shares used for seasons that start late (typical team's games played / season length):")
    for name, share in target_share.items():
        print(f"  {PLAIN_NAMES[name]:<14} {share:.1%}  (about {share * 82:.0f} of 82 games)")

    # --- Ask for the standings on each checkpoint date (skip ones we already have) ---
    have = pd.read_csv(SNAPSHOTS_FILE, encoding="utf-8-sig") if SNAPSHOTS_FILE.exists() else pd.DataFrame()
    this_year = current_start_year()
    new_frames, asked = [], 0
    for cp in checkpoints[checkpoints["reached"]].itertuples():
        if not have.empty and not args.force and int(str(cp.season_id)[:4]) < this_year:
            if ((have["season_id"] == cp.season_id) & (have["checkpoint"] == cp.checkpoint) &
                    (have["requested_date"] == str(cp.date))).any():
                continue
        rows = fetch_snapshot(cp.season_id, cp.checkpoint, cp.date, abbrev_to_id)
        asked += 1
        if rows:
            snap = mark_playoff_spots(pd.DataFrame(rows))
            new_frames.append(snap)
        else:
            print(f"  No standings returned for {cp.season_id} {cp.checkpoint} ({cp.date})")
    if new_frames:
        fresh = pd.concat(new_frames, ignore_index=True)
        if not have.empty:   # replace any old copy of a snapshot we just asked for again
            key = fresh[["season_id", "checkpoint"]].drop_duplicates()
            have = have.merge(key, on=["season_id", "checkpoint"], how="left", indicator=True)
            have = have[have["_merge"] == "left_only"].drop(columns="_merge")
        snapshots = pd.concat([have, fresh], ignore_index=True)
    else:
        snapshots = have
    print(f"\nAsked the NHL for {asked} standings snapshots.")
    if snapshots.empty:
        print("No snapshots yet.")
        return 1

    snapshots = snapshots.sort_values(["season_id", "checkpoint", "league_seq"]).reset_index(drop=True)
    CHECKPOINTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SNAPSHOTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    checkpoints.to_csv(CHECKPOINTS_FILE, index=False, encoding="utf-8-sig")
    snapshots.to_csv(SNAPSHOTS_FILE, index=False, encoding="utf-8-sig")

    # --- Check 1: do the standings on date D already count the games played ON date D? ---
    print("\nCheck 1: does 'standings on date D' include games played on D? (games played vs our own count)")
    inc = exc = total = 0
    for (season, cp), snap in snapshots.groupby(["season_id", "checkpoint"]):
        day = pd.to_datetime(snap["requested_date"].iloc[0]).date()
        sg = games[(games["seasonId"] == season)]
        through = sg[sg["gameDate"] <= day].groupby("teamId").size()
        before = sg[sg["gameDate"] < day].groupby("teamId").size()
        for r in snap.itertuples():
            if pd.isna(r.team_id):
                continue
            total += 1
            inc += int(through.get(int(r.team_id), 0) == r.gp)
            exc += int(before.get(int(r.team_id), 0) == r.gp)
    print(f"  Games played match our count INCLUDING that day: {inc / total:.1%}   EXCLUDING that day: {exc / total:.1%}")
    print("  (the higher one tells us how to read a snapshot; the rest is games that day with no team playing)")

    # --- Check 2: at the end of each season, 'in a playoff spot' should equal 'made the playoffs' ---
    print("\nCheck 2: end-of-season 'in a playoff spot' vs teams that really made the playoffs")
    if TEAM_SEASONS_FILE.exists():
        actual = pd.read_csv(TEAM_SEASONS_FILE)[["season_id", "team_id", "made_playoffs"]]
        end = snapshots[snapshots["checkpoint"] == "season_end"].merge(
            actual, on=["season_id", "team_id"], how="left")
        bad = end[end["in_playoff_spot"] != end["made_playoffs"]]
        print(f"  {len(end)} team-seasons checked, {len(bad)} disagree.")
        for season, part in bad.groupby("season_id"):
            print(f"    {season}: " + ", ".join(part["team_abbrev"]))
        unmatched = snapshots["team_id"].isna().sum()
        if unmatched:
            print(f"  WARNING: {unmatched} snapshot rows could not be matched to a team id.")
    else:
        print("  (run python scripts/build_reference_tables.py first to enable this check)")

    print(f"\nSaved: {CHECKPOINTS_FILE}")
    print(f"Saved: {SNAPSHOTS_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

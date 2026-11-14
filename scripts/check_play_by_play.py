"""
check_play_by_play.py  -  did the play-by-play download come out right?
=======================================================================
WHAT IT DOES (in Excel terms):
    Looks at what pull_play_by_play.py saved in data/pbp.db and checks it against numbers we already trust:

      * COVERAGE       how many games are saved per season, against how many were played
      * GOALS          for each game, do the goals in the play-by-play match the NHL's goals for both teams?
      * LOCATIONS      what share of shots, hits, faceoffs, giveaways and takeaways have a position (x, y) and a zone?
      * SIZE           events per game by season (a season with far fewer events would be incomplete)
      * ROSTERS        about 40 players per game, with positions

    It is safe to run at any time, even while the download is still going: it checks whatever is saved so far.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/check_play_by_play.py

WHAT YOU SHOULD SEE:
    A season-by-season table and a few PASS / CHECK lines.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import sqlite3
import sys
from pathlib import Path

import pandas as pd

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
PBP_DB = PROJECT_FOLDER / "data" / "pbp.db"
NHL_DB = PROJECT_FOLDER / "data" / "nhl.db"
LOCATED_TYPES = ("shot-on-goal", "missed-shot", "blocked-shot", "goal", "hit", "faceoff", "giveaway", "takeaway")


def season_label(season_id):
    text = str(int(season_id))
    return f"{text[:4]}-{text[6:]}"


def ok(passed):
    return "PASS" if passed else "CHECK"


def main():
    sys.stdout.reconfigure(errors="replace")
    if not PBP_DB.exists() or not NHL_DB.exists():
        print("I can't find data/pbp.db or data/nhl.db yet. Run pull_play_by_play.py first.")
        return 1
    pbp = sqlite3.connect(PBP_DB)
    nhl = sqlite3.connect(NHL_DB)
    nhl_games = pd.read_sql_query(
        "SELECT game_id, season_id, game_type, home_team_id, away_team_id, home_goals, away_goals FROM games WHERE season_id >= 20092010", nhl)
    saved = pd.read_sql_query("SELECT game_id, season_id, status, n_events FROM pbp_games", pbp)

    print("COVERAGE (games saved vs games played)")
    played = nhl_games.groupby("season_id").size().rename("played")
    ok_games = saved[saved["status"] == "ok"].groupby("season_id").size().rename("saved")
    missing = saved[saved["status"] == "not_found"].groupby("season_id").size().rename("not_available")
    events = saved[saved["status"] == "ok"].groupby("season_id")["n_events"].mean().rename("events_per_game")
    table = pd.concat([played, ok_games, missing, events], axis=1).fillna(0)
    table.insert(0, "season", [season_label(s) for s in table.index])
    table["share_saved"] = (table["saved"] / table["played"]).round(3)
    table["events_per_game"] = table["events_per_game"].round(0)
    print(table.to_string(index=False))
    total_saved, total_played = int(table["saved"].sum()), int(table["played"].sum())
    print(f"\n  Total: {total_saved:,} of {total_played:,} games saved ({total_saved / total_played:.1%}); "
          f"{int(table['not_available'].sum())} not available from the NHL.")

    print("\nGOALS (play-by-play goals vs the NHL's goals, shootout goals excluded)")
    goals = pd.read_sql_query(
        "SELECT game_id, team_id, COUNT(*) AS goals FROM pbp_events WHERE type = 'goal' AND period_type != 'SO' GROUP BY game_id, team_id", pbp)
    m = saved[saved["status"] == "ok"][["game_id"]].merge(nhl_games, on="game_id")
    home = goals.rename(columns={"team_id": "home_team_id", "goals": "pbp_home"})
    away = goals.rename(columns={"team_id": "away_team_id", "goals": "pbp_away"})
    m = m.merge(home, on=["game_id", "home_team_id"], how="left").merge(away, on=["game_id", "away_team_id"], how="left").fillna({"pbp_home": 0, "pbp_away": 0})
    same = (m["pbp_home"] == m["home_goals"]) & (m["pbp_away"] == m["away_goals"])
    print(f"  {len(m):,} games compared; both teams' goals identical in {same.mean():.2%}  -> {ok(same.mean() > 0.995)}")
    if not same.all():
        bad = m[~same]
        print(f"  {len(bad)} games differ. First few (game, play-by-play home/away, NHL home/away):")
        for r in bad.head(8).itertuples():
            print(f"    {int(r.game_id)}: {int(r.pbp_home)}-{int(r.pbp_away)} vs {int(r.home_goals)}-{int(r.away_goals)}")

    print("\nLOCATIONS (share of located event types that have x, y and a zone)")
    placeholders = ",".join("?" * len(LOCATED_TYPES))
    loc = pd.read_sql_query(
        f"SELECT g.season_id, COUNT(*) AS n, SUM(e.x IS NOT NULL AND e.y IS NOT NULL) AS with_xy, SUM(e.zone IS NOT NULL) AS with_zone "
        f"FROM pbp_events e JOIN pbp_games g ON g.game_id = e.game_id WHERE e.type IN ({placeholders}) GROUP BY g.season_id ORDER BY g.season_id",
        pbp, params=list(LOCATED_TYPES))
    loc["season"] = loc["season_id"].map(season_label)
    loc["x,y share"] = (loc["with_xy"] / loc["n"]).round(4)
    loc["zone share"] = (loc["with_zone"] / loc["n"]).round(4)
    print(loc[["season", "n", "x,y share", "zone share"]].to_string(index=False))
    worst = float(loc["x,y share"].min()) if len(loc) else 0
    print(f"  Lowest share with a position in any season: {worst:.1%}  -> {ok(worst > 0.98)}")

    print("\nROSTERS")
    r = pd.read_sql_query("SELECT COUNT(*) * 1.0 / COUNT(DISTINCT game_id) AS per_game FROM pbp_rosters", pbp)
    positions = dict(pbp.execute("SELECT position, COUNT(*) FROM pbp_rosters GROUP BY position").fetchall())
    print(f"  {r['per_game'].iloc[0]:.1f} players per game; positions: {positions}")

    dupes = pbp.execute("SELECT COUNT(*) FROM (SELECT game_id, sort_order FROM pbp_events GROUP BY game_id, sort_order HAVING COUNT(*) > 1)").fetchone()[0]
    print(f"\nDUPLICATES: {dupes} repeated (game, event) rows  -> {ok(dupes == 0)}")
    print(f"\nDatabase size: {PBP_DB.stat().st_size / 1e9:.2f} GB")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

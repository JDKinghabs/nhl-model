"""
build_breakout_bei.py  -  BEI, the Breakout Efficiency Index (version 1)
========================================================================
WHAT IT DOES (in Excel terms):
    A second way of rating breakouts, next to BER. BEI blends three numbers, each turned into a league percentile
    within the season (100 = best team, 0 = worst), with fixed weights:

      45%  Own-zone giveaways per hour of defensive-zone time   (lower is better)
           = giveaways in the team's own zone at 5-on-5 (NHL play-by-play)
             divided by hours of 5-on-5 play spent in its own zone (5-on-5 ice time x the NHL EDGE own-zone share)
      30%  Share of even-strength time spent in its own zone    (lower is better; NHL EDGE puck tracking)
      25%  Transition shot attempts per 60 minutes at 5-on-5    (higher is better; a shot attempt within 5 seconds of
           play coming up the ice, read from the play-by-play)

    BEI = 0.45 x giveaway percentile + 0.30 x zone-time percentile + 0.25 x transition percentile, from 0 to 100.

    NHL EDGE zone time starts in 2021-22, so BEI exists from then on. The script also checks how BEI lines up with
    winning (points % and goal difference in the same season) next to BER and Corsi, on the finished seasons.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_breakout_bei.py
    It needs the play-by-play (pull_play_by_play.py), the zone time (pull_edge_zone_time.py) and the database.

WHAT YOU SHOULD SEE:
    A short table and a new file:  outputs/breakout_bei.xlsx   (close it in Excel first).

DATA CREDIT:
    Data is (c) NHL and NHL EDGE, used here for personal, non-commercial purposes. 5-on-5 ice time from MoneyPuck.com
    (non-commercial use, credit required); for the newest games, 5-on-5 time counted from the play-by-play.
"""

import sqlite3
import sys

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_breakout_index as bi          # the per-game counts from the play-by-play
import build_breakout_round2 as r2         # BER, for comparison
import build_popular_stats_check as ps     # season_spearman

OUTPUT_FILE = bi.PROJECT_FOLDER / "outputs" / "breakout_bei.xlsx"
EDGE_FILE = bi.PROJECT_FOLDER / "data" / "nhl" / "edge" / "zone_time_regular.csv"
FIRST_SEASON = 20212022                    # NHL EDGE zone time starts here

COMPONENTS = [  # (column, plain name, weight, higher is better?)
    ("dz_giveaways_ph", "Own-zone giveaways per hour of own-zone time", 0.45, False),
    ("dz_share", "Share of even-strength time in own zone", 0.30, False),
    ("transition_p60", "Transition shot attempts per 60", 0.25, True),
]


def percentile(values, higher_is_better):
    """Within one season: 100 for the best team, 0 for the worst, evenly spaced (ties share a value)."""
    r = values.rank(ascending=higher_is_better, method="average")
    return (r - 1) / (r.count() - 1) * 100


def bei_table(rows, edge):
    """One row per team-season: the three parts, their percentiles and BEI."""
    g = rows.groupby(["season_id", "team_id"]).agg(giveaways=("giveaways", "sum"), trans_for=("trans_for", "sum"), toi5=("toi5", "sum"),
                                                   games=("game_id", "nunique")).reset_index()
    t = g.merge(edge, on=["season_id", "team_id"], how="inner")
    t["dz_giveaways_ph"] = t["giveaways"] / (t["toi5"] / 3600 * t["dz_share"])
    t["transition_p60"] = t["trans_for"] / (t["toi5"] / 3600)
    t["BEI"] = 0.0
    for col, _, weight, higher in COMPONENTS:
        t[f"pct_{col}"] = t.groupby("season_id")[col].transform(lambda v, h=higher: percentile(v, h))
        t["BEI"] += weight * t[f"pct_{col}"]
    return t


def ber_table(rows):
    with_opp = r2.with_opponent(rows)
    g = with_opp.groupby(["season_id", "team_id"]).sum(numeric_only=True)
    return pd.DataFrame({"BER": r2.raw_stats(g, "", "")["ber"]}).reset_index()


def main():
    sys.stdout.reconfigure(errors="replace")
    if not bi.PBP_DB.exists() or not bi.NHL_DB.exists() or not EDGE_FILE.exists():
        print("I need data/pbp.db, data/nhl.db and data/nhl/edge/zone_time_regular.csv (pull_edge_zone_time.py).")
        return 1
    pbp = sqlite3.connect(f"file:{bi.PBP_DB}?mode=ro", uri=True)
    nhl = sqlite3.connect(bi.NHL_DB)

    finished = bi.load_game_rows(pbp, nhl)
    current = bi.load_game_rows(pbp, nhl, in_progress=True)
    if finished is None or current is None:
        return 1
    rows = pd.concat([finished[0][finished[0]["season_id"] >= FIRST_SEASON], current[0]], ignore_index=True)
    this_season = int(current[1][0])

    edge = pd.read_csv(EDGE_FILE, encoding="utf-8-sig").rename(columns={"teamId": "team_id", "es_dz_pct": "dz_share"})
    edge = edge[["season_id", "team_id", "dz_share"]].dropna()
    edge["team_id"] = edge["team_id"].astype(int)

    t = bei_table(rows, edge).merge(ber_table(rows), on=["season_id", "team_id"], how="left")
    seasons = sorted(t["season_id"].unique())
    compare = bi.comparison_numbers(nhl, seasons)[["season_id", "team_id", "team", "pts_pct", "gd_pg", "corsi", "outcome"]]
    t = t.merge(compare, on=["season_id", "team_id"], how="left")
    missing = 32 * len(seasons) - len(t)
    if missing > 0:
        print(f"Note: {missing} team-seasons have no zone-time row and are left out.")

    # --- the test: how each lines up with winning, on finished seasons only ---
    done = t[t["season_id"] < this_season]
    tests = []
    for col, name in [("BEI", "BEI (breakout efficiency index)"), ("BER", "BER (breakout efficiency rating)"),
                      ("pct_dz_giveaways_ph", "BEI part: own-zone giveaways per hour of own-zone time"), ("pct_dz_share", "BEI part: share of time in own zone"),
                      ("pct_transition_p60", "BEI part: transition shot attempts per 60"), ("corsi", "Yardstick: Corsi (shot-attempt share)")]:
        pts, gd = ps.season_spearman(done, col, "pts_pct"), ps.season_spearman(done, col, "gd_pg")
        tests.append({"Statistic": name, "Link with points %": round(pts.mean(), 2), "Link with goal difference": round(gd.mean(), 2),
                      "Lowest season (points %)": round(pts.min(), 2), "Highest season (points %)": round(pts.max(), 2), "Seasons": len(pts)})
    tests = pd.DataFrame(tests)
    between = round(float(ps.season_spearman(done, "BEI", "BER").mean()), 2)

    def table(season):
        d = t[t["season_id"] == season].sort_values("BEI", ascending=False)
        return pd.DataFrame({"Team": d["team"], "Games": d["games"].astype(int), "BEI": d["BEI"].round(1),
                             "Own-zone giveaways per hour of own-zone time": d["dz_giveaways_ph"].round(1),
                             "Share of even-strength time in own zone": d["dz_share"].round(3),
                             "Transition shot attempts per 60": d["transition_p60"].round(2),
                             "BER": d["BER"].round(1), "Points %": d["pts_pct"].round(3), "How the season ended": d["outcome"]}).reset_index(drop=True)

    last_season = max(s for s in seasons if s < this_season)
    dates = pd.read_sql_query("SELECT game_id, game_date FROM games WHERE season_id = ?", nhl, params=[this_season])
    through = dates[dates["game_id"].isin(current[0]["game_id"])]["game_date"].max()
    note = f"Season {bi.season_label(this_season)}: games through {through}. Early in a season these numbers are noisy."
    read_me = ["BEI, the Breakout Efficiency Index (version 1)", "",
               "Three numbers, each turned into a league percentile within the season (100 = best, 0 = worst), weighted 45 / 30 / 25:",
               "  45%  own-zone giveaways at 5-on-5 per hour of 5-on-5 time spent in the own zone (lower is better)",
               "  30%  share of even-strength time spent in the own zone, from NHL EDGE puck tracking (lower is better)",
               "  25%  transition shot attempts per 60 minutes at 5-on-5 (higher is better)",
               f"Seasons {bi.season_label(FIRST_SEASON)} on (NHL EDGE zone time starts then). 'Test' compares BEI, BER and Corsi with winning on the finished seasons.",
               f"BEI and BER rank teams alike at {between:+.2f} (rank correlation within each season, averaged).", "",
               "Data: (c) NHL and NHL EDGE. 5-on-5 ice time: MoneyPuck.com."]
    try:
        OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            pd.DataFrame({"Read me": read_me}).to_excel(writer, sheet_name="Read me", index=False)
            writer.sheets["Read me"].column_dimensions["A"].width = 140
            sheets = [("This season", table(this_season), note), ("Last season", table(last_season), f"Season {bi.season_label(last_season)}"),
                      ("Test", tests, f"Finished seasons {bi.season_label(min(done['season_id']))} to {bi.season_label(max(done['season_id']))}: rank correlation within each season, averaged"),
                      ("All seasons", t.sort_values(["season_id", "BEI"], ascending=[False, False]).round(3), "Every team-season with all the parts")]
            for name, frame, top in sheets:
                frame.to_excel(writer, sheet_name=name, index=False, startrow=1)
                ws = writer.sheets[name]
                ws.cell(row=1, column=1, value=top).font = Font(bold=True)
                ws.freeze_panes = "B3"
                for cell in ws[2]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                ws.column_dimensions["A"].width = 30
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 220)
    print(f"\n{note}")
    show = table(this_season)
    show.columns = ["Team", "GP", "BEI", "DZ GA/hr", "DZ share", "Trans/60", "BER", "P%", "Ended"]
    print(show.drop(columns="Ended").to_string(index=False))
    print("\nTest (finished seasons):"); print(tests.to_string(index=False))
    print(f"BEI vs BER, same team-seasons: {between:+.2f}")
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL and NHL EDGE; 5-on-5 ice time: MoneyPuck.com")
    return 0


if __name__ == "__main__":
    sys.exit(main())

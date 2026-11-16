"""
build_breakout_now.py  -  this season's breakout numbers for every team
=======================================================================
WHAT IT DOES (in Excel terms):
    Works out, for the season being played now, the same breakout numbers the Breakouts page shows for last season:
    BER (controlled exits, safe clears, failed), pinned-in spells, and each team's score on the two FROZEN recipes
    ("our own breakouts" and "breakout battle"). Nothing is searched or re-fitted here: the recipes are read from
    reference/breakout_metric_frozen.json exactly as they were frozen before the blind test.

    Each number is compared with the other teams THIS season, so a score of 0 is the league average so far.

    EARLY IN A SEASON THESE NUMBERS ARE NOISY. BER and the rest settle over a full season; after 20 games they are a
    first impression, not a verdict.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_breakout_now.py
    It uses the games of this season whose play-by-play has been downloaded. To add the latest games first:
        python scripts/pull_game_results.py
        python scripts/build_database.py
        python scripts/load_moneypuck.py
        python scripts/pull_play_by_play.py
    (A game that MoneyPuck's file does not have yet uses 5-on-5 time counted from the events instead.)

WHAT YOU SHOULD SEE:
    A short table in the terminal and a new file:  outputs/breakout_now.xlsx   (close it in Excel first).
    Then  python scripts/build_site.py  puts it on the Breakouts page.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. 5-on-5 ice time comes from MoneyPuck.com
    (non-commercial use, credit required).
"""

import json
import sqlite3
import sys

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_breakout_index as bi          # the per-game counts from the play-by-play
import build_breakout_round2 as r2         # the 68 building blocks and the frozen recipes

OUTPUT_FILE = bi.PROJECT_FOLDER / "outputs" / "breakout_now.xlsx"


def main():
    sys.stdout.reconfigure(errors="replace")
    if not bi.PBP_DB.exists() or not bi.NHL_DB.exists():
        print("I need data/pbp.db (python scripts/pull_play_by_play.py) and data/nhl.db (python scripts/build_database.py).")
        return 1
    if not r2.FROZEN_FILE.exists():
        print("There is no frozen recipe file (reference/breakout_metric_frozen.json).")
        return 1
    frozen = json.loads(r2.FROZEN_FILE.read_text(encoding="utf-8"))
    if frozen["settings"] != r2.settings_now():
        print("The stat settings at the top of build_breakout_index.py are not the ones the recipes were frozen with. Put them back and run again.")
        return 1
    pbp = sqlite3.connect(f"file:{bi.PBP_DB}?mode=ro", uri=True)
    nhl = sqlite3.connect(bi.NHL_DB)
    loaded = bi.load_game_rows(pbp, nhl, in_progress=True)
    if loaded is None:
        return 1
    rows, seasons, _ = loaded
    season = seasons[0]
    dates = pd.read_sql_query("SELECT game_id, game_date FROM games WHERE season_id = ? AND game_type = 2", nhl, params=[season])
    through = dates[dates["game_id"].isin(rows["game_id"])]["game_date"].max()
    not_yet = int((~dates["game_id"].isin(rows["game_id"])).sum())

    with_opp = r2.with_opponent(rows)
    g = with_opp.groupby(["season_id", "team_id"]).sum(numeric_only=True)
    own = r2.raw_stats(g, "", "")
    judged = g["ber_ctrl"] + g["ber_safe"] + g["ber_fail"]
    blocks = r2.block_table(with_opp).set_index(["season_id", "team_id"])            # each block against the other teams THIS season
    t = pd.DataFrame({"Games": rows.groupby(["season_id", "team_id"])["game_id"].nunique(), "BER": own["ber"],
                      "Controlled exits (share of judged attempts)": g["ber_ctrl"] / judged, "Safe clears (share)": g["ber_safe"] / judged,
                      "Failed (share)": g["ber_fail"] / judged, "Pinned-in spells per 60": own["pinned_pg"], "Zone-clear rate": own["exit_success"],
                      "Defence-to-offence rate": own["d2o_rate"], "Own-zone turnovers per 60": own["turnover_pg"]})
    for tier, label in r2.TIERS:
        w = np.array([frozen["metrics"][tier]["weights"].get(b, 0.0) for b in r2.BLOCKS])
        raw = pd.Series(blocks[r2.BLOCKS].to_numpy() @ w, index=blocks.index)
        t[f"{label} score (0 = league average)"] = (raw - raw.mean()) / raw.std()
    names = pd.read_sql_query("SELECT team_id, team, pts_pct FROM team_seasons WHERE season_id = ?", nhl, params=[season])
    t = t.reset_index().merge(names, on="team_id", how="left").sort_values("BER", ascending=False)
    out = pd.DataFrame({"Team": t["team"], "Games": t["Games"].astype(int), "BER": t["BER"].round(1),
                        **{c: t[c].round(3) for c in ["Controlled exits (share of judged attempts)", "Safe clears (share)", "Failed (share)"]},
                        "Pinned-in spells per 60": t["Pinned-in spells per 60"].round(2), "Zone-clear rate": t["Zone-clear rate"].round(3),
                        "Defence-to-offence rate": t["Defence-to-offence rate"].round(3), "Own-zone turnovers per 60": t["Own-zone turnovers per 60"].round(2),
                        **{f"{label} score (0 = league average)": t[f"{label} score (0 = league average)"].round(2) for _, label in r2.TIERS},
                        "Points %": t["pts_pct"].round(3)}).reset_index(drop=True)
    note = (f"Season {bi.season_label(season)}: games through {through} ({rows['game_id'].nunique():,} games; {int(out['Games'].min())} to {int(out['Games'].max())} per team"
            + (f"; {not_yet} more played but not downloaded yet" if not_yet else "") + "). Early in a season these numbers are noisy.")

    read_me = ["This season's breakout numbers", "",
               "The same numbers as the Breakouts study, for the season being played now. The two scores use the recipes frozen in reference/breakout_metric_frozen.json:",
               "nothing is searched or re-fitted. Each score compares a team with the other teams THIS season: 0 = league average, +1 = clearly good, -1 = clearly poor.",
               "BER = 100 x (controlled exits + 0.5 x safe clears) / breakout attempts that could be judged, read from play-by-play events (an approximation).",
               "Early in a season these numbers are noisy: after about 20 games they are a first impression, not a verdict.", "",
               "Data: (c) NHL. 5-on-5 ice time: MoneyPuck.com."]
    try:
        OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            pd.DataFrame({"Read me": read_me}).to_excel(writer, sheet_name="Read me", index=False)
            writer.sheets["Read me"].column_dimensions["A"].width = 150
            out.to_excel(writer, sheet_name="Right now", index=False, startrow=1)
            ws = writer.sheets["Right now"]
            ws.cell(row=1, column=1, value=note).font = Font(bold=True)
            ws.freeze_panes = "B3"
            for cell in ws[2]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            ws.column_dimensions["A"].width = 26
            for column in list(ws.columns)[1:]:
                ws.column_dimensions[column[0].column_letter].width = 16
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 220)
    print(f"\n{note}")
    show = out[["Team", "Games", "BER", "Pinned-in spells per 60"] + [f"{label} score (0 = league average)" for _, label in r2.TIERS] + ["Points %"]]
    show.columns = ["Team", "Games", "BER", "Pinned /60", "Own", "Battle", "P%"]
    print(show.to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL; 5-on-5 ice time: MoneyPuck.com")
    return 0


if __name__ == "__main__":
    sys.exit(main())

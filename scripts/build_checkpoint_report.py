"""
build_checkpoint_report.py  -  "who was in a playoff spot at each checkpoint, and did they make it?"
==================================================================================================
WHAT IT DOES (in Excel terms):
    Reads the database (data/nhl.db) and writes an Excel workbook with:
      * the checkpoint dates for every season,
      * every team's standing at every checkpoint next to how its season ended,
      * a FIRST LOOK table: of the teams that were in a playoff spot at Thanksgiving, what share made
        the playoffs, and the same for teams that were out, split by how close they were to the line,
      * lists of teams that were out at Thanksgiving but made it, and teams that were in but missed.

    It is descriptive only: it counts what happened. It does not yet say WHY, or which teams were
    likely to recover. That comes next.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py          (first, so the database is up to date)
    python scripts/build_checkpoint_report.py

WHAT YOU SHOULD SEE:
    A printed headline table and a new file:  outputs/checkpoint_standings.xlsx
    (close it in Excel before running again).

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import sqlite3
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
DB_FILE = PROJECT_FOLDER / "data" / "nhl.db"
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "checkpoint_standings.xlsx"

CHECKPOINT_ORDER = ["nov1", "thanksgiving", "jan1", "deadline", "season_end"]
PLAIN = {"nov1": "Nov 1", "thanksgiving": "Thanksgiving", "jan1": "Jan 1",
         "deadline": "Trade deadline", "season_end": "End of season"}

# How far from the cut line (in points percentage; 0.03 is about 5 points over a full season)
BAND_ORDER = ["In: far clear (more than 0.08 above the line)", "In: comfortable (0.03 to 0.08 above)",
              "In: bubble (within 0.03 of the line)", "Out: just out (within 0.03 below)",
              "Out: behind (0.03 to 0.08 below)", "Out: far behind (more than 0.08 below)"]

READ_ME = [
    "Checkpoint standings: who was in a playoff spot, and did they make it?",
    "",
    "HOW 'IN A PLAYOFF SPOT' IS DECIDED",
    "  2013-14 on:   top 3 in each division + 2 wild cards per conference.",
    "  2005-06 to 2012-13:  top 8 in each conference, division leaders seeded 1-3.",
    "  2019-20 (end of season): top 12 per conference (qualifying round). 2020-21: top 4 per division.",
    "  I tested this on the last day of every season: it matches who really made the playoffs in all 644 team-seasons.",
    "",
    "CHECKPOINT DATES",
    "  Nov 1, US Thanksgiving, Jan 1, the trade deadline and the last day of the regular season.",
    "  A standing 'on date D' counts the games played on D (tested: 100% match).",
    "  Seasons that started late (2012-13 lockout, 2020-21) have no Nov 1 or Jan 1: each checkpoint is placed",
    "  where the typical team had played the same SHARE of its season as in a normal year.",
    "  Some trade-deadline dates are from memory and flagged in reference/trade_deadlines.csv: please check them.",
    "",
    "SHEETS",
    "  First look           make-the-playoffs rate by position at each checkpoint (START HERE)",
    "  By season            the same, one row per season (shows how much it bounces around)",
    "  Out then in          teams that were OUT of a spot at Thanksgiving but made the playoffs",
    "  In then out          teams that were IN a spot at Thanksgiving but missed",
    "  Snapshots            every team at every checkpoint, with how their season ended",
    "  Checkpoint dates     the dates, and how each one was chosen",
    "",
    "WHAT THIS DOES NOT SAY",
    "  These are plain counts over about 19 seasons (about 600 team-seasons at each checkpoint). Small groups",
    "  (like 'far behind at Nov 1') are small, so do not over-read a single percentage.",
    "  2019-20 and 2020-21 are left out of the First look and By season sheets (see reference/season_flags.csv).",
    "",
    "Data: (c) NHL.",
]


def load(connection):
    snaps = pd.read_sql_query(
        """SELECT s.season_id, s.checkpoint, s.requested_date AS date, s.team_id, s.team_abbrev, s.team,
                  s.conference, s.division, s.gp, s.w, s.l, s.otl, s.pts, s.pts_pct, s.conf_seq, s.div_seq,
                  s.wildcard_seq, s.in_playoff_spot,
                  t.season, t.outcome, t.made_playoffs, t.include_main
           FROM standings_snapshots s
           LEFT JOIN team_seasons t ON t.season_id = s.season_id AND t.team_id = s.team_id""", connection)
    dates = pd.read_sql_query("SELECT * FROM checkpoints WHERE reached = 1", connection)
    return snaps, dates


def add_cushion_bands(snaps):
    """How far above/below the cut line was each team? (cut line = the weakest team that is in a spot, per conference)"""
    cut = (snaps[snaps["in_playoff_spot"] == 1]
           .groupby(["season_id", "checkpoint", "conference"])["pts_pct"].min().rename("cut_pct").reset_index())
    snaps = snaps.merge(cut, on=["season_id", "checkpoint", "conference"], how="left")
    snaps["cushion"] = (snaps["pts_pct"] - snaps["cut_pct"]).round(3)

    def band(row):
        c = row["cushion"]
        if pd.isna(c):
            return None
        if row["in_playoff_spot"] == 1:
            return BAND_ORDER[0] if c > 0.08 else BAND_ORDER[1] if c > 0.03 else BAND_ORDER[2]
        return BAND_ORDER[3] if c >= -0.03 else BAND_ORDER[4] if c >= -0.08 else BAND_ORDER[5]

    snaps["band"] = snaps.apply(band, axis=1)
    return snaps


def first_look(main):
    rows = []
    for cp in CHECKPOINT_ORDER:
        part = main[main["checkpoint"] == cp]
        if part.empty:
            continue
        for label, mask in (("ALL teams in a spot", part["in_playoff_spot"] == 1),
                            ("ALL teams out of a spot", part["in_playoff_spot"] == 0)):
            g = part[mask]
            rows.append({"Checkpoint": PLAIN[cp], "Group": label, "Team-seasons": len(g),
                         "Made the playoffs": int(g["made_playoffs"].sum()),
                         "% made it": round(100 * g["made_playoffs"].mean(), 1) if len(g) else None})
        for b in BAND_ORDER:
            g = part[part["band"] == b]
            rows.append({"Checkpoint": PLAIN[cp], "Group": b, "Team-seasons": len(g),
                         "Made the playoffs": int(g["made_playoffs"].sum()),
                         "% made it": round(100 * g["made_playoffs"].mean(), 1) if len(g) else None})
    return pd.DataFrame(rows)


def by_season(main):
    rows = []
    for season, sp in main.groupby("season"):
        row = {"Season": season}
        for cp in CHECKPOINT_ORDER[:-1]:        # the last one is the answer itself
            part = sp[sp["checkpoint"] == cp]
            if part.empty:
                continue
            ins, outs = part[part["in_playoff_spot"] == 1], part[part["in_playoff_spot"] == 0]
            row[f"{PLAIN[cp]}: in a spot, made it"] = f"{int(ins['made_playoffs'].sum())} of {len(ins)}"
            row[f"{PLAIN[cp]}: out, made it"] = f"{int(outs['made_playoffs'].sum())} of {len(outs)}"
        rows.append(row)
    return pd.DataFrame(rows)


def swing_list(main, checkpoint, was_in, did_make):
    part = main[(main["checkpoint"] == checkpoint) & (main["in_playoff_spot"] == was_in) &
                (main["made_playoffs"] == did_make)]
    out = part[["season", "team", "date", "gp", "pts", "pts_pct", "conf_seq", "cushion", "outcome"]].copy()
    out.columns = ["Season", "Team", "Date", "GP", "PTS", "PTS%", "Conf. rank", "Cushion vs line", "How the season ended"]
    return out.sort_values(["Season", "Conf. rank"]).reset_index(drop=True)


def write_excel(sheets):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 118
        ws["A1"].font = Font(bold=True)
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column in ws.columns:
                longest = max((len(str(c.value)) for c in column[:300] if c.value is not None), default=8)
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 56)


def main():
    sys.stdout.reconfigure(errors="replace")
    if not DB_FILE.exists():
        print("I can't find the database yet. Run this first:")
        print("    python scripts/build_database.py")
        return 1
    connection = sqlite3.connect(DB_FILE)
    snaps, dates = load(connection)
    connection.close()
    snaps = add_cushion_bands(snaps)

    finished = snaps[snaps["outcome"].notna() & (snaps["outcome"] != "Season in progress")]
    main_seasons = finished[finished["include_main"] == "Y"]

    look = first_look(main_seasons)
    sheets = {
        "First look": look,
        "By season": by_season(main_seasons),
        "Out then in": swing_list(main_seasons, "thanksgiving", 0, 1),
        "In then out": swing_list(main_seasons, "thanksgiving", 1, 0),
        "Snapshots": snaps.drop(columns=["include_main"]),
        "Checkpoint dates": dates[["season_id", "checkpoint", "date", "method", "games_played_typical", "season_games", "note"]],
    }
    try:
        write_excel(sheets)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    n_seasons = main_seasons["season_id"].nunique()
    print(f"Main seasons used: {n_seasons} (2019-20 and 2020-21 left out)\n")
    print("Share of teams that made the playoffs, by whether they were in a playoff spot at each checkpoint:")
    top = look[look["Group"].str.startswith("ALL")]
    print(top.to_string(index=False))
    print()
    print("Same, by how close to the line (Thanksgiving):")
    print(look[(look["Checkpoint"] == "Thanksgiving") & ~look["Group"].str.startswith("ALL")].to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

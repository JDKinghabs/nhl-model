"""
build_coverage_report.py  -  "what data do we actually have, and from which season?"
====================================================================================
WHAT IT DOES (in Excel terms):
    Looks at every team-season table we downloaded and, for each statistic and each season,
    checks whether the numbers are really there. Some old seasons come back with the column
    present but empty (or filled with zeros), which looks fine until you try to use it.
    The result is a workbook with a colour-coded grid: green = real data, yellow = partly
    there, orange = only zeros, red = nothing.

    This tells us which questions we can answer for which seasons BEFORE we build anything on
    top, so that gaps are known facts and not surprises.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_coverage_report.py

    Run these first (they download the data it reads):
        python scripts/pull_team_season_reports.py
        python scripts/pull_edge_zone_time.py

WHAT YOU SHOULD SEE:
    A short plain-English list of where the data starts, and a new file:
        outputs/data_coverage.xlsx      (close it in Excel before re-running)
    Start with the "Key metrics" sheet.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
TEAM_FOLDER = PROJECT_FOLDER / "data" / "nhl" / "team_season"
EDGE_FOLDER = PROJECT_FOLDER / "data" / "nhl" / "edge"
FLAGS_FILE = PROJECT_FOLDER / "reference" / "season_flags.csv"
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "data_coverage.xlsx"

ENOUGH = 95            # "real data" means at least this % of teams have a value
SKIP_COLUMNS = {"seasonId", "teamId", "teamFullName", "gameTypeId"}

# The statistics that matter for our questions: (plain name, report, column, why it matters).
KEY_METRICS = [
    ("Points %",                         "summary", "pointPct",                    "Baseline: how good the team's record was"),
    ("Goals for per game",               "summary", "goalsForPerGame",             "Baseline"),
    ("Goals against per game",           "summary", "goalsAgainstPerGame",         "Baseline"),
    ("Shots for per game",               "summary", "shotsForPerGame",             "Shot volume"),
    ("Shots against per game",           "summary", "shotsAgainstPerGame",         "Shot volume"),
    ("Faceoff win %",                    "summary", "faceoffWinPct",               "Possession starts"),
    ("Power-play %",                     "summary", "powerPlayPct",                "Factor 5: power kill"),
    ("Penalty-kill %",                   "summary", "penaltyKillPct",              "Factor 5: power kill"),
    ("Power-play opportunities",         "powerplay", "ppOpportunities",           "Factor 5: special-teams volume"),
    ("Times short-handed",               "penaltykill", "timesShorthanded",        "Factor 5: special-teams volume"),
    ("Power-play time per game",         "powerplay", "ppTimeOnIcePerGame",        "Factor 5: special-teams volume"),
    ("Net power-play goals",             "powerplay", "ppNetGoals",                "Factor 5: net special teams"),
    ("Net penalty-kill goals",           "penaltykill", "pkNetGoals",              "Factor 5: net special teams"),
    ("5v5 goals for",                    "goalsforbystrength", "goalsFor5On5",     "Even-strength scoring"),
    ("5v5 goals against",                "goalsagainstbystrength", "goalsAgainst5On5", "Even-strength defence"),
    ("Shot-attempt share (Corsi)",       "percentages", "satPct",                  "Factors 1-2: who controls play"),
    ("Unblocked-attempt share (Fenwick)", "percentages", "usatPct",                "Factors 1-2: who controls play"),
    ("5v5 shooting %",                   "percentages", "shootingPct5v5",          "Luck check; factor 4"),
    ("5v5 save %",                       "percentages", "savePct5v5",              "Factor 4: goaltending"),
    ("PDO (shooting % + save %)",        "percentages", "shootingPlusSavePct5v5",  "Luck check"),
    ("Zone starts % (5v5)",              "percentages", "zoneStartPct5v5",         "Factor 2: offensive-zone time (rough)"),
    ("Shot attempts for (5v5)",          "summaryshooting", "satFor",              "Factors 1-2"),
    ("Shot attempts against (5v5)",      "summaryshooting", "satAgainst",          "Factors 1-2"),
    ("Hits",                             "realtime", "hits",                       "Popular stat test"),
    ("Blocked shots",                    "realtime", "blockedShots",               "Defence style"),
    ("Giveaways",                        "realtime", "giveaways",                  "Puck management"),
    ("Takeaways",                        "realtime", "takeaways",                  "Puck management"),
    ("Penalties drawn",                  "penalties", "totalPenaltiesDrawn",       "Factor 5: special-teams volume"),
    ("Offensive-zone faceoff %",         "faceoffpercentages", "offensiveZoneFaceoffPct", "Possession starts"),
    ("Defensive-zone faceoff %",         "faceoffpercentages", "defensiveZoneFaceoffPct", "Possession starts"),
    ("Win % when scoring first",         "scoretrailfirst", "winPctScoringFirst",  "Game-state context"),
    ("Wins when out-shooting opponent",  "outshootoutshotby", "winsOutshootOpponent", "Luck check"),
]

READ_ME = [
    "Data coverage report: which statistics exist, and from which season?",
    "",
    "WHY IT EXISTS",
    "  The NHL returns a row for every team in every season back to 2005-06, but for the older seasons many",
    "  statistics come back empty or as zeros. This workbook checks the real numbers, season by season, so",
    "  we know which questions can be answered for which seasons before we build anything on top.",
    "",
    "HOW TO READ THE COLOURS",
    f"  Green  (OK)       real numbers for at least {ENOUGH}% of teams",
    "  Yellow (partial)  numbers for some teams but not most",
    "  Orange (zeros)    every value is 0. That is real for a few things (no ties after 2004-05) and means",
    "                    'not recorded' for others. Treat orange as missing unless the stat can really be zero.",
    "  Red    (none)     nothing there",
    "",
    "SHEETS",
    "  Key metrics        the statistics we care about, with the first season each one is really available.",
    "                     START HERE.",
    "  Cautions           things I found in the data that you should know before trusting a number.",
    "  League averages    league-wide averages per game by season: shows drifts and breaks in the numbers.",
    "  Zone time (EDGE)   the NHL's tracking-based zone time (2021-22 on): teams with data per season.",
    "  Other sources      data that is not in these tables (play-by-play, shifts, MoneyPuck ...): where it",
    "                     starts and whether we have pulled it yet.",
    "  Row counts         teams per report per season (should be 30-32 for the regular season).",
    "  All columns (regular) / All columns (playoffs)   every column of every report, same colours.",
    "",
    "SEASONS TO HANDLE WITH CARE (see reference/season_flags.csv)",
    "  2005-06 and 2012-13 are included but flagged. 2019-20 and 2020-21 are excluded from main results.",
    "",
    "Data: (c) NHL. This report says what data exists, not what it means.",
]

# What is NOT in the team-season tables. Status words: Have / Not pulled yet / Manual download.
OTHER_SOURCES = [
    ("Playoff series, champions, conference finalists", "NHL web API", "2005-06", "Have",
     "outputs/playoff_history.xlsx. All 21 seasons checked; 20 champions match my known list."),
    ("Team season totals (these tables)", "NHL stats API", "2005-06", "Have",
     "Results, special teams, hits etc. from 2005-06. Everything at 5v5 and all shot-attempt statistics start 2009-10."),
    ("Game-by-game team results", "NHL stats API", "2005-06", "Not pulled yet",
     "Next step. Needed for checkpoints (Nov 1, Thanksgiving, Jan 1, deadline) and game-level betting later."),
    ("Standings on any past date", "NHL web API", "2005-06", "Not pulled yet",
     "Next step. Shows who was in a playoff spot on a given date."),
    ("Trade-deadline dates by season", "Hand-checked list", "2005-06", "Not pulled yet",
     "A small reference table I will build and you can check."),
    ("Shot coordinates and zone tags (play-by-play)", "NHL web API", "2009-10", "Not pulled yet",
     "None in 2007-08, partial in 2006-07. About 22,000 requests (6+ hours, run overnight). I will ask first."),
    ("Shift charts (who was on the ice)", "NHL stats API", "2010-11", "Not pulled yet",
     "Needed for defence-pair minutes. Heavy; planned with the play-by-play pull."),
    ("Real zone time (offensive / neutral / defensive)", "NHL EDGE", "2021-22", "see Zone time (EDGE)",
     "Only 2021-22 on. Before that we can only estimate it, and EDGE is how we check the estimate."),
    ("High-danger shots, expected goals, goals saved above expected, defence pairings", "MoneyPuck.com",
     "2008-09", "Manual download",
     "MoneyPuck blocks scripts, so you download about 15 files by hand (I will walk you through it)."),
    ("Clean breakouts (defence to offence)", "Does not exist in free data", "-", "Proxy only",
     "I will estimate it from play-by-play (defensive-zone event followed quickly by a shot attempt)."),
    ("Player statistics (skaters and goalies)", "NHL stats API", "2005-06 (to check)", "Not pulled yet",
     "Same idea as the team tables; comes after the team work."),
]

# League-wide averages per team per game, to show how the numbers drift over the years:
# (plain name, report, column, divide by games played?)
LEAGUE_AVERAGES = [
    ("Goals for per game",              "summary", "goalsForPerGame", False),
    ("Shots for per game",              "summary", "shotsForPerGame", False),
    ("Power-play opportunities per game", "powerplay", "ppOpportunitiesPerGame", False),
    ("Hits per game",                   "realtime", "hits", True),
    ("Blocked shots per game",          "realtime", "blockedShots", True),
    ("Giveaways per game",              "realtime", "giveaways", True),
    ("Takeaways per game",              "realtime", "takeaways", True),
    ("Total shot attempts per game",    "realtime", "totalShotAttempts", True),
    ("5v5 shooting %",                  "percentages", "shootingPct5v5", False),
    ("5v5 save %",                      "percentages", "savePct5v5", False),
]

# Things I found in the data that you should know before trusting a number.
CAUTIONS = [
    ("Giveaways and takeaways changed in 2024-25",
     "League-average giveaways per game roughly doubled and takeaways fell by about a third, "
     "from one season to the next. That looks like a change in how they are recorded, not in how "
     "teams play. Do not compare these two across that break; compare teams only within a season."),
    ("Hits climbed for years",
     "Average hits per team went from about 1,300 (2005-06) to about 1,900 (2010s). Part of that is "
     "how arenas count them. Compare teams within a season (rank or percentile), never raw totals across seasons."),
    ("Total shot attempts are missing for five seasons",
     "The 'totalShotAttempts' column in the realtime table is blank for 2017-18 to 2021-22. The shot-attempt "
     "share (Corsi) in the percentages table is filled for all of those seasons, so use that one."),
    ("No five-on-five detail before 2009-10",
     "Goals at 5v5, shot-attempt share, shooting %, save % and zone starts all begin in 2009-10. "
     "2005-06 to 2008-09 only have results, special teams, shots on goal, hits and similar totals."),
    ("Two odd rows in 2009-10",
     "Florida and Tampa Bay each have one game with no shot data in the shot-attempt table, so their "
     "season totals there cover 81 games. Shares and percentages are not affected."),
    ("Power plays became much rarer after 2005-06",
     "Power-play opportunities per game fell from about 5.9 in 2005-06 (the post-lockout rules crackdown) to "
     "about 3.1 by 2014-15. For the 'power kill' factor, special-teams volume has to be compared within a season, "
     "and 2005-06 is an unusual year."),
    ("Short seasons have small totals",
     "2012-13 (48 games), 2019-20 (about 70) and 2020-21 (56) have smaller totals. Use per-game numbers "
     "or within-season ranks."),
    ("The current season is still being played",
     "The newest season only has the games played so far. It is meant for the live view, not for the history."),
    ("Team ids change when a franchise moves or renames",
     "Atlanta (id 11) became Winnipeg (id 52) in 2011. Phoenix (27) became Arizona (53), then Utah Hockey Club (59), "
     "then Utah Mammoth (68). To follow one franchise through time we need a small franchise lookup table; "
     "I will add one before doing any team-over-time work."),
    ("The scoring environment has drifted",
     "5v5 shooting % rose from about 7.5% (2015-16) to about 9.5% (2025-26). That is why every comparison "
     "will be made within a season (rank or percentile), never on raw numbers across seasons."),
]

GREEN, YELLOW, ORANGE, RED = "C6EFCE", "FFEB9C", "F8CBAD", "FFC7CE"
FILL = {"OK": GREEN, "partial": YELLOW, "zeros": ORANGE, "none": RED}
SYMBOL = {"OK": "OK", "partial": "partial", "zeros": "zeros", "none": "none"}


def season_label(season_id):
    text = str(int(season_id))
    return f"{text[:4]}-{text[6:]}"


def load_tables(kind):
    """Load every stacked report file of one kind ('regular' or 'playoffs') -> {report: DataFrame}."""
    tables = {}
    for path in sorted(TEAM_FOLDER.glob(f"*_{kind}.csv")):
        frame = pd.read_csv(path, encoding="utf-8-sig")
        # A few rows have no team id at all (2009-10: one game each for Florida and Tampa Bay with
        # no shot data). They are not real teams, so leave them out of the checks.
        tables[path.stem[: -(len(kind) + 1)]] = frame[frame["teamId"].notna()]
    return tables


def status_for(series):
    """Look at one column for one season and say OK / partial / zeros / none."""
    if len(series) == 0:
        return "none"
    filled = series.notna().mean() * 100
    if filled == 0:
        return "none"
    if filled < ENOUGH:
        return "partial"
    values = series.dropna()
    if (values == 0).all():
        return "zeros"
    return "OK"


def build_status_grid(tables, seasons):
    """One row per (report, column); one column per season; each cell is a status word."""
    rows = []
    for report, frame in tables.items():
        for column in frame.columns:
            if column in SKIP_COLUMNS:
                continue
            row = {"report": report, "column": column}
            for season in seasons:
                part = frame.loc[frame["seasonId"] == season, column]
                row[season_label(season)] = status_for(part)
            rows.append(row)
    return pd.DataFrame(rows)


def first_ok(row, labels):
    for label in labels:
        if row[label] == "OK":
            return label
    return "never"


def gaps_after_first(row, labels):
    first = first_ok(row, labels)
    if first == "never":
        return ""
    later = labels[labels.index(first):]
    bad = [label for label in later if row[label] != "OK"]
    return ", ".join(bad)


def build_key_metrics(grid_reg, grid_po, labels):
    lookup_reg = grid_reg.set_index(["report", "column"])
    lookup_po = grid_po.set_index(["report", "column"]) if not grid_po.empty else None
    rows = []
    for name, report, column, why in KEY_METRICS:
        key = (report, column)
        if key not in lookup_reg.index:
            rows.append({"Statistic": name, "Used for": why, "First season (regular)": "not downloaded",
                         "Gaps after that": "", "First season (playoffs)": ""})
            continue
        reg_row = lookup_reg.loc[key]
        row = {"Statistic": name, "Used for": why,
               "First season (regular)": first_ok(reg_row, labels),
               "Gaps after that": gaps_after_first(reg_row, labels),
               "First season (playoffs)": first_ok(lookup_po.loc[key], labels)
               if lookup_po is not None and key in lookup_po.index else "not downloaded"}
        for label in labels:
            row[label] = reg_row[label]
        rows.append(row)
    return pd.DataFrame(rows)


def row_counts(tables, seasons):
    rows = []
    for report, frame in tables.items():
        counts = frame.groupby("seasonId").size()
        row = {"report": report}
        for season in seasons:
            row[season_label(season)] = int(counts.get(season, 0))
        rows.append(row)
    return pd.DataFrame(rows)


def league_averages(tables, seasons):
    """The league-wide average of a few statistics, per game, by season (shows drifts and breaks)."""
    rows = []
    for name, report, column, per_game in LEAGUE_AVERAGES:
        row = {"Statistic": name}
        frame = tables.get(report)
        for season in seasons:
            value = None
            if frame is not None and column in frame.columns:
                part = frame[frame["seasonId"] == season]
                if per_game:
                    value = (part[column] / part["gamesPlayed"]).mean()
                else:
                    value = part[column].mean()
            row[season_label(season)] = None if pd.isna(value) else round(float(value), 3)
        rows.append(row)
    return pd.DataFrame(rows)


def edge_summary():
    log_file = EDGE_FOLDER / "zone_time_log.csv"
    if not log_file.exists():
        return pd.DataFrame({"note": ["EDGE zone time has not been pulled yet. Run: python scripts/pull_edge_zone_time.py"]})
    log = pd.read_csv(log_file)
    log["season"] = log["season_id"].map(season_label)
    log["game type"] = log["game_type"].map({2: "regular season", 3: "playoffs"})
    return log[["season", "game type", "teams_asked", "teams_with_data", "pulled_at"]]


def other_sources(edge):
    frame = pd.DataFrame(OTHER_SOURCES, columns=["Data", "Source", "Earliest season", "Status", "Note"])
    if "teams_with_data" in edge.columns:
        have = edge[edge["teams_with_data"] > 0]
        if not have.empty:
            frame.loc[frame["Data"].str.startswith("Real zone time"), "Status"] = "Have"
    return frame


# ----------------------------------------------------------------------------
# Writing the Excel file
# ----------------------------------------------------------------------------
def colour_cells(ws, first_col_index):
    """Colour every status cell from first_col_index (1-based) rightwards."""
    for row in ws.iter_rows(min_row=2):
        for cell in row[first_col_index - 1:]:
            word = cell.value
            if word in FILL:
                cell.fill = PatternFill("solid", fgColor=FILL[word])
                cell.alignment = Alignment(horizontal="center")


def style_sheet(ws, wrap_cols=(), width_cap=48):
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    for column in ws.columns:
        longest = max((len(str(c.value)) for c in column[:300] if c.value is not None), default=8)
        ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), width_cap)
    for letter in wrap_cols:
        for cell in ws[letter][1:]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def write_excel(key, averages, edge, others, counts_reg, counts_po, grid_reg, grid_po):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 118
        ws["A1"].font = Font(bold=True)

        key.to_excel(writer, sheet_name="Key metrics", index=False)
        ws = writer.sheets["Key metrics"]
        style_sheet(ws)
        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["B"].width = 38
        ws.column_dimensions["D"].width = 30
        colour_cells(ws, 6)
        ws.freeze_panes = "B2"

        pd.DataFrame(CAUTIONS, columns=["Caution", "What to know"]).to_excel(
            writer, sheet_name="Cautions", index=False)
        ws = writer.sheets["Cautions"]
        style_sheet(ws, wrap_cols=("A", "B"))
        ws.column_dimensions["A"].width = 44
        ws.column_dimensions["B"].width = 120

        averages.to_excel(writer, sheet_name="League averages", index=False)
        ws = writer.sheets["League averages"]
        style_sheet(ws)
        ws.column_dimensions["A"].width = 34
        for column in list(ws.columns)[1:]:
            ws.column_dimensions[column[0].column_letter].width = 9
        ws.freeze_panes = "B2"

        edge.to_excel(writer, sheet_name="Zone time (EDGE)", index=False)
        style_sheet(writer.sheets["Zone time (EDGE)"])

        others.to_excel(writer, sheet_name="Other sources", index=False)
        ws = writer.sheets["Other sources"]
        style_sheet(ws, wrap_cols=("A", "E"))
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["E"].width = 80

        counts_reg.to_excel(writer, sheet_name="Row counts", index=False, startrow=1)
        ws = writer.sheets["Row counts"]
        ws["A1"] = "Regular season: teams per report per season"
        ws["A1"].font = Font(bold=True)
        start = len(counts_reg) + 5
        ws.cell(row=start - 1, column=1, value="Playoffs: teams per report per season (16 normally; 24 in 2019-20)").font = Font(bold=True)
        counts_po.to_excel(writer, sheet_name="Row counts", index=False, startrow=start - 1)
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width = 10
        ws.column_dimensions["A"].width = 26

        for name, grid in (("All columns (regular)", grid_reg), ("All columns (playoffs)", grid_po)):
            grid.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            style_sheet(ws)
            ws.auto_filter.ref = ws.dimensions
            colour_cells(ws, 3)
            ws.freeze_panes = "C2"
            for column in list(ws.columns)[2:]:
                ws.column_dimensions[column[0].column_letter].width = 9


def main():
    sys.stdout.reconfigure(errors="replace")
    tables_reg = load_tables("regular")
    if not tables_reg:
        print("I can't find any downloaded team tables yet. Run this first:")
        print("    python scripts/pull_team_season_reports.py")
        return 1
    tables_po = load_tables("playoffs")

    seasons = sorted({int(s) for frame in tables_reg.values() for s in frame["seasonId"].unique()})
    labels = [season_label(s) for s in seasons]
    print(f"Found {len(tables_reg)} regular-season reports and {len(tables_po)} playoff reports, "
          f"{labels[0]} to {labels[-1]}.")

    grid_reg = build_status_grid(tables_reg, seasons)
    grid_po = build_status_grid(tables_po, seasons) if tables_po else pd.DataFrame(columns=["report", "column"] + labels)
    key = build_key_metrics(grid_reg, grid_po, labels)
    edge = edge_summary()
    others = other_sources(edge)
    counts_reg = row_counts(tables_reg, seasons)
    counts_po = row_counts(tables_po, seasons) if tables_po else pd.DataFrame(columns=["report"] + labels)

    averages = league_averages(tables_reg, seasons)
    try:
        write_excel(key, averages, edge, others, counts_reg, counts_po, grid_reg, grid_po)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    print()
    print("Where does the data really start? (regular season; 'never' = no usable numbers)")
    for _, row in key.iterrows():
        gap = f"   gaps later: {row['Gaps after that']}" if row["Gaps after that"] else ""
        print(f"  {row['Statistic']:<36} {row['First season (regular)']:<14}{gap}")
    print()
    print(f"Saved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

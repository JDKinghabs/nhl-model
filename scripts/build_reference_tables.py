"""
build_reference_tables.py  -  playoff history and season facts, 2005-06 to now
==============================================================================
WHAT IT DOES (in Excel terms):
    Builds the "lookup tables" the research will lean on, like the reference sheets in a workbook:
      * which team won the Cup each season, who reached the conference finals (the "final four"),
      * how far EVERY team went each season (including teams that missed the playoffs),
      * a one-row-per-season fact sheet (teams, games played, flags for unusual seasons).
    Results go into one Excel file you can open, plus CSV copies for later scripts.

HOW TO RUN IT (in the VS Code terminal, from the "Model 2.0" folder):
    python scripts/build_reference_tables.py

WHAT YOU SHOULD SEE:
    One line per season (about 22 lines, with about 40 seconds of waiting), a check of the
    champions against a known list, a quick look at how champions ranked in the regular season,
    and a new file:  outputs/playoff_history.xlsx   (close it in Excel before re-running).

DATA CREDIT:
    Data is (c) NHL, used for personal, non-commercial purposes.

The seasons you want to treat specially (like the 2020 bubble) are listed in
reference/season_flags.csv. You can edit that file in Excel.
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font

import nhl_api          # our shared, polite downloader (scripts/nhl_api.py)

# ----------------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------------
PROJECT_FOLDER = Path(__file__).resolve().parent.parent
FLAGS_FILE = PROJECT_FOLDER / "reference" / "season_flags.csv"
DATA_FOLDER = PROJECT_FOLDER / "data" / "reference"       # ignored by git
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "playoff_history.xlsx"   # ignored by git

FIRST_START_YEAR = 2005    # 2005-06: the first season after the lockout and the start of the salary cap

BRACKET_URL = "https://api-web.nhle.com/v1/playoff-bracket/{year}"      # year = when the playoffs happen
SUMMARY_URL = "https://api.nhle.com/stats/rest/en/team/summary"

# A sanity check, typed in from memory: who won the Cup (keyed by the year the playoffs ended).
# If the NHL's data ever disagrees, the script says so loudly. That is a clue something is off.
KNOWN_CHAMPIONS = {
    2006: "CAR", 2007: "ANA", 2008: "DET", 2009: "PIT", 2010: "CHI", 2011: "BOS", 2012: "LAK",
    2013: "CHI", 2014: "LAK", 2015: "CHI", 2016: "PIT", 2017: "PIT", 2018: "WSH", 2019: "STL",
    2020: "TBL", 2021: "TBL", 2022: "COL", 2023: "VGK", 2024: "FLA", 2025: "FLA",
}

OUTCOME_ORDER = ["Won Cup", "Lost Cup final", "Lost conference final", "Lost round 2",
                 "Lost round 1", "Lost qualifying round", "Missed playoffs"]


def season_label(start_year):
    return f"{start_year}-{str(start_year + 1)[2:]}"


def current_start_year():
    """The season that is under way now (seasons start in the autumn)."""
    today = datetime.now()
    return today.year if today.month >= 9 else today.year - 1


# ----------------------------------------------------------------------------
# Downloading
# ----------------------------------------------------------------------------
def get_regular_season_teams(season_id):
    """One row per team for the regular season (wins, points, goals ...)."""
    params = {"isAggregate": "false", "isGame": "false", "start": 0, "limit": -1,
              "cayenneExp": f"gameTypeId=2 and seasonId={season_id}"}
    raw = nhl_api.get_json(SUMMARY_URL, params=params)
    return (raw or {}).get("data", [])


def get_series(season_id, end_year):
    """Every playoff series that season: who played, who won. Returns a list of rows."""
    raw = nhl_api.get_json(BRACKET_URL.format(year=end_year), quiet_404=True)
    rows = []
    for s in (raw or {}).get("series", []):
        top, bottom = s.get("topSeedTeam") or {}, s.get("bottomSeedTeam") or {}
        rows.append({
            "season_id": season_id,
            "round": s.get("playoffRound"),
            "series_letter": s.get("seriesLetter"),
            "series_title": s.get("seriesTitle"),
            "top_id": top.get("id"),
            "top_abbrev": top.get("abbrev"),
            "top_team": (top.get("name") or {}).get("default"),
            "top_wins": s.get("topSeedWins"),
            "bottom_id": bottom.get("id"),
            "bottom_abbrev": bottom.get("abbrev"),
            "bottom_team": (bottom.get("name") or {}).get("default"),
            "bottom_wins": s.get("bottomSeedWins"),
            "winner_id": s.get("winningTeamId"),
            "loser_id": s.get("losingTeamId"),
        })
    return rows


# ----------------------------------------------------------------------------
# Working out how far each team went
# ----------------------------------------------------------------------------
def outcome_for(reached, is_champion, status, played_qualifier=False):
    if status == "Regular season in progress":
        return "Season in progress"
    if status != "Finished":
        return "Unknown (no playoff data)"
    if is_champion:
        return "Won Cup"
    if reached == 0 and played_qualifier:
        return "Lost qualifying round"          # 2019-20 only
    return {0: "Missed playoffs", 1: "Lost round 1", 2: "Lost round 2",
            3: "Lost conference final", 4: "Lost Cup final"}[reached]


def build_season(start_year, current_year, flags):
    """Download one season and return (season fact row, team rows, series rows)."""
    end_year = start_year + 1
    season_id = int(f"{start_year}{end_year}")
    teams = get_regular_season_teams(season_id)
    series = get_series(season_id, end_year) if start_year < current_year else []
    # (The current season has no finished bracket yet, so we don't ask for one.)

    # Who won it all, and how deep did each team go?
    final = next((s for s in series if s["round"] == 4 and s["winner_id"]), None)
    champion_id = final["winner_id"] if final else None
    reached = {}
    qualifiers = set()      # teams that played the 2020 bubble "qualifying round" (called round 0)
    for s in series:
        for team_id in (s["top_id"], s["bottom_id"]):
            if team_id is None:
                continue
            if s["round"] == 0:
                qualifiers.add(team_id)
            else:
                reached[team_id] = max(reached.get(team_id, 0), s["round"])

    if start_year >= current_year:
        status = "Regular season in progress"
    elif final:
        status = "Finished"
    else:
        status = "Bracket not found"

    flag = flags.get(season_id, {})
    include_main = flag.get("include_main", "Y")

    team_rows = []
    for t in teams:
        gp = t.get("gamesPlayed") or 0
        points = t.get("points") or 0
        goals_for, goals_against = t.get("goalsFor"), t.get("goalsAgainst")
        r = reached.get(t["teamId"], 0)
        is_champ = t["teamId"] == champion_id
        team_rows.append({
            "season_id": season_id, "season": season_label(start_year),
            "team_id": t["teamId"], "team": t["teamFullName"],
            "gp": gp, "wins": t.get("wins"), "losses": t.get("losses"), "ot_losses": t.get("otLosses"),
            "points": points, "pts_pct": round(points / (2 * gp), 3) if gp else None,
            "goals_for": goals_for, "goals_against": goals_against,
            "goal_diff": (goals_for - goals_against) if goals_for is not None and goals_against is not None else None,
            "made_playoffs": r > 0 or t["teamId"] in qualifiers, "round_reached": r,
            "outcome": outcome_for(r, is_champ, status, t["teamId"] in qualifiers),
            "conf_finalist": r >= 3, "cup_finalist": r >= 4, "champion": is_champ,
            "include_main": include_main, "season_note": flag.get("note", ""),
        })

    # How good were they in the regular season, compared with the rest of that season's league?
    if team_rows:
        frame = pd.DataFrame(team_rows)
        rank = frame["pts_pct"].rank(ascending=False, method="min")
        for row, rk in zip(team_rows, rank):
            row["pts_pct_rank"] = int(rk)
            n = len(team_rows)
            row["pts_pct_percentile"] = round(100 * (1 - (rk - 1) / (n - 1)), 1) if n > 1 else None

    by_id = {t["teamId"]: t["teamFullName"] for t in teams}
    abbrev = {}
    for s in series:
        abbrev[s["top_id"]], abbrev[s["bottom_id"]] = s["top_abbrev"], s["bottom_abbrev"]
    gps = [t.get("gamesPlayed") or 0 for t in teams]
    finalists = [t for t in team_rows if t["conf_finalist"]]
    runner_up_id = final["loser_id"] if final else None
    season_row = {
        "season_id": season_id, "season": season_label(start_year), "status": status,
        "teams": len(teams), "games_min": min(gps) if gps else None, "games_max": max(gps) if gps else None,
        "series_found": len(series),
        "champion": by_id.get(champion_id, ""), "champion_abbrev": abbrev.get(champion_id, ""),
        "cup_runner_up": by_id.get(runner_up_id, ""),
        "final_series_score": (f"{max(final['top_wins'], final['bottom_wins'])}-{min(final['top_wins'], final['bottom_wins'])}"
                               if final else ""),
        "conference_finalists": "; ".join(t["team"] for t in finalists),
        "include_main": include_main, "issue": flag.get("issue", ""), "note": flag.get("note", ""),
    }

    # Say what we found.
    summary_line = f"{season_label(start_year)}: {len(teams)} teams, {len(series)} series"
    if final:
        summary_line += f", champion {abbrev.get(champion_id, '?')}"
    else:
        summary_line += f"  [{status}]"
    print("  " + summary_line)
    return season_row, team_rows, series


# ----------------------------------------------------------------------------
# Excel output
# ----------------------------------------------------------------------------
README_LINES = [
    "PLAYOFF HISTORY - read me",
    "",
    "What this is: one lookup workbook for the contender research. Built by scripts/build_reference_tables.py from NHL data (c) NHL.",
    "",
    "Sheets:",
    "  Champions        One row per finished season: who won, who they beat, the four conference finalists.",
    "  Final four       The four teams that reached the conference finals each season, with their regular-season rank.",
    "  By outcome       How teams that went different distances ranked in the regular season (main seasons only).",
    "  Team seasons     EVERY team in EVERY season and how far it went (including teams that missed the playoffs).",
    "  Seasons          One row per season: team count, games played, status, and the flags you can edit.",
    "  All series       Every playoff series with both teams and the winner.",
    "",
    "Columns worth knowing:",
    "  pts_pct          Points won out of points available (points / (2 x games played)). Fair across short seasons.",
    "  pts_pct_rank     1 = best in the league that season.",
    "  pts_pct_percentile  100 = best, 0 = worst in that season (like PERCENTRANK in Excel). Lets us compare eras.",
    "  round_reached    0 = missed the playoffs, 1-4 = the furthest round played (4 = the Cup final).",
    "  include_main     Y/N. N means the season is left out of the main results (see below).",
    "",
    "Unusual seasons (edit reference/season_flags.csv to change these):",
    "  2005-06  First season after the lockout. Included, but data is thin.",
    "  2012-13  48-game season. Included, flagged.",
    "  2019-20  Cut short, 24-team bubble playoff. Excluded from main results. The 8 teams that lost the",
    "           qualifying round are labelled 'Lost qualifying round' (they count as making the playoffs).",
    "  2020-21  56 games, division-only schedule and playoffs. Excluded from main results.",
]


def write_excel(sheets):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        readme = pd.DataFrame({"Read me": README_LINES})
        readme.to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 120
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
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 48)


# ----------------------------------------------------------------------------
# main() is where the script starts, like the first line of a macro.
# ----------------------------------------------------------------------------
def main():
    sys.stdout.reconfigure(errors="replace")
    flags = {}
    if FLAGS_FILE.exists():
        flags = {int(r["season_id"]): r for r in pd.read_csv(FLAGS_FILE).fillna("").to_dict("records")}

    current_year = current_start_year()
    seasons, teams, series = [], [], []
    print(f"Downloading {season_label(FIRST_START_YEAR)} to {season_label(current_year)} "
          f"(about {2 * (current_year - FIRST_START_YEAR + 1)} requests, one per second)...")
    for start_year in range(FIRST_START_YEAR, current_year + 1):
        season_row, team_rows, series_rows = build_season(start_year, current_year, flags)
        seasons.append(season_row)
        teams.extend(team_rows)
        series.extend(series_rows)

    seasons_df, teams_df, series_df = pd.DataFrame(seasons), pd.DataFrame(teams), pd.DataFrame(series)

    # --- Check the champions against the list we typed in from memory ---
    print("\nChecking champions against the known list:")
    problems = []
    for s in seasons:
        end_year = int(str(s["season_id"])[4:])
        known = KNOWN_CHAMPIONS.get(end_year)
        if known and s["champion_abbrev"] != known:
            problems.append(f"  MISMATCH {s['season']}: NHL data says '{s['champion_abbrev']}', I expected '{known}'")
    checked = sum(1 for s in seasons if int(str(s["season_id"])[4:]) in KNOWN_CHAMPIONS)
    if problems:
        print("\n".join(problems))
    else:
        print(f"  All {checked} known champions match.")
    unchecked = [f"{s['season']} ({s['champion'] or s['status']})" for s in seasons
                 if int(str(s["season_id"])[4:]) not in KNOWN_CHAMPIONS and s["status"] != "Regular season in progress"]
    if unchecked:
        print("  Not on my known list, so taken from the NHL data only: " + ", ".join(unchecked))

    # --- Sheets for Excel ---
    finished = seasons_df[seasons_df["status"] == "Finished"]
    champions = finished[["season", "champion", "cup_runner_up", "final_series_score", "conference_finalists",
                          "include_main", "note"]].copy()
    final_four = teams_df[teams_df["conf_finalist"] & teams_df["season_id"].isin(finished["season_id"])][
        ["season", "team", "outcome", "points", "pts_pct", "pts_pct_rank", "pts_pct_percentile", "goal_diff",
         "include_main", "season_note"]].copy()
    final_four["outcome"] = pd.Categorical(final_four["outcome"], OUTCOME_ORDER, ordered=True)
    final_four = final_four.sort_values(["season", "outcome"]).reset_index(drop=True)
    final_four["outcome"] = final_four["outcome"].astype(str)

    main_teams = teams_df[(teams_df["include_main"] == "Y") & teams_df["season_id"].isin(finished["season_id"])]
    by_outcome = (main_teams.groupby("outcome")
                  .agg(team_seasons=("team", "count"),
                       avg_rank=("pts_pct_rank", "mean"),
                       avg_percentile=("pts_pct_percentile", "mean"),
                       share_top_8=("pts_pct_rank", lambda x: (x <= 8).mean()),
                       share_top_16=("pts_pct_rank", lambda x: (x <= 16).mean()))
                  .reindex(OUTCOME_ORDER).dropna(subset=["team_seasons"]).round(2).reset_index())
    by_outcome["team_seasons"] = by_outcome["team_seasons"].astype(int)

    # --- Save CSV copies (for later scripts) and the Excel workbook ---
    DATA_FOLDER.mkdir(parents=True, exist_ok=True)
    seasons_df.to_csv(DATA_FOLDER / "seasons.csv", index=False, encoding="utf-8-sig")
    teams_df.to_csv(DATA_FOLDER / "team_seasons.csv", index=False, encoding="utf-8-sig")
    series_df.to_csv(DATA_FOLDER / "playoff_series.csv", index=False, encoding="utf-8-sig")
    try:
        write_excel({"Champions": champions, "Final four": final_four, "By outcome": by_outcome,
                     "Team seasons": teams_df, "Seasons": seasons_df, "All series": series_df})
    except PermissionError:
        print(f"\nCould not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    # --- A quick look ---
    print("\nQuick look: where did each group rank in the REGULAR season? (main seasons only; 1 = best)")
    print(by_outcome.to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print(f"Saved CSV copies in: {DATA_FOLDER}")
    print("Data: (c) NHL")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())

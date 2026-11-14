"""
pull_nhl_team_stats.py  -  team statistics from the NHL's public stats site
===========================================================================
WHAT IT DOES (in Excel terms):
    Downloads team statistics for a season from the NHL and saves each table as a CSV you can
    open in Excel. The NHL calls each table a "report" (summary, powerplay, penaltykill, ...).

HOW TO RUN IT (in the VS Code terminal, from the "Model 2.0" folder):
    python scripts/pull_nhl_team_stats.py

WHAT YOU SHOULD SEE:
    A few printed lines, a "top 5 by points" list, and a new file in  data/nhl/

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
    This is an unofficial interface and can change without notice, so the script
    explains problems in plain English instead of crashing.
"""

import sys
import time
from pathlib import Path

import pandas as pd      # pandas = "Excel for Python": a DataFrame is a worksheet
import requests          # requests = fetches things from the internet

# ----------------------------------------------------------------------------
# SETTINGS  (change these, then run again)
# ----------------------------------------------------------------------------
# The NHL writes a season as both years joined together: 20252026 means 2025-26.
SEASONS = [20252026]

# Which tables ("reports") to download. We'll add more later (powerplay, penaltykill, ...).
REPORTS = ["summary"]

GAME_TYPE = 2            # 2 = regular season, 3 = playoffs

URL = "https://api.nhle.com/stats/rest/en/team/{report}"

PROJECT_FOLDER = Path(__file__).resolve().parent.parent   # the "Model 2.0" folder
SAVE_FOLDER = PROJECT_FOLDER / "data" / "nhl"             # "data" is ignored by git

HEADERS = {"User-Agent": "nhl-model-hobby-project (personal, non-commercial)"}


def season_label(season):
    """20252026 -> '2025-26' (nicer to read)."""
    text = str(season)
    return f"{text[:4]}-{text[6:]}"


# ----------------------------------------------------------------------------
# A "function" is a named recipe, like a custom Excel formula.
# ----------------------------------------------------------------------------
def download_report(report, season):
    # These are the "filters" we send, like choosing options in an Excel Data > From Web query.
    params = {
        "isAggregate": "false",
        "isGame": "false",       # false = one row per team for the whole season (true = one row per game)
        "start": 0,
        "limit": -1,             # -1 means "give me every row"
        "cayenneExp": f"gameTypeId={GAME_TYPE} and seasonId={season}",
    }
    print(f"Downloading NHL team '{report}' report for {season_label(season)}...")

    try:
        response = requests.get(URL.format(report=report), params=params, headers=HEADERS, timeout=60)
        response.raise_for_status()
        rows = response.json().get("data", [])
    except requests.RequestException as problem:
        print(f"  Could not download it: {problem}")
        print("  Check your internet connection. The NHL may also have changed this address.")
        return None
    except ValueError:
        print("  The NHL answered, but not with the data format we expected. It may have changed.")
        return None

    if not rows:
        print("  The NHL returned zero rows. Is the season right? Has it started yet?")
        return None

    sheet = pd.DataFrame(rows)    # turn the list of rows into a worksheet

    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)
    save_path = SAVE_FOLDER / f"team_{report}_{season}.csv"
    try:
        # utf-8-sig makes Excel show accented letters (like the e in Montreal) correctly.
        sheet.to_csv(save_path, index=False, encoding="utf-8-sig")
    except PermissionError:
        print(f"  Could not save {save_path.name}. Is it open in Excel? Close it and run again.")
        return None

    print(f"  Saved: {save_path}")
    print(f"  Rows: {len(sheet)}   Columns: {len(sheet.columns)}")
    return sheet


def show_top_five(sheet):
    """Quick eyeball check you can compare with nhl.com."""
    if {"teamFullName", "points", "gamesPlayed"} <= set(sheet.columns):
        top = sheet.sort_values("points", ascending=False).head(5)
        print("  Top 5 by points:")
        for _, team in top.iterrows():
            print(f"    {team['teamFullName']:<25} {int(team['points'])} pts in {int(team['gamesPlayed'])} games")


# ----------------------------------------------------------------------------
# main() is where the script starts, like the first line of a macro.
# ----------------------------------------------------------------------------
def main():
    sys.stdout.reconfigure(errors="replace")      # never crash on an unusual character
    saved_any = False
    for season in SEASONS:
        for report in REPORTS:
            sheet = download_report(report, season)
            if sheet is not None:
                saved_any = True
                if report == "summary":
                    show_top_five(sheet)
            time.sleep(1)                         # be polite: a one-second pause between requests
    print("Data: (c) NHL")
    return 0 if saved_any else 1


if __name__ == "__main__":
    sys.exit(main())

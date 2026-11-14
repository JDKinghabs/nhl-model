"""
pull_moneypuck_teams.py  -  your first data-pull script
=======================================================
WHAT IT DOES (in Excel terms):
    Downloads one season's team stats from MoneyPuck.com and saves them as a CSV
    you can open in Excel. Think of it as a macro that does "Data > From Web" for you.

HOW TO RUN IT (in the VS Code terminal, from the "Model 2.0" folder):
    python scripts/pull_moneypuck_teams.py

WHAT YOU SHOULD SEE:
    A few printed lines, then a new file in  data/moneypuck/teams_2025.csv

NOTE (found on first run): MoneyPuck blocks automatic downloads unless you have a data license,
    so this script currently stops with an explanation. Files downloaded by hand work fine.

DATA CREDIT:
    Data provided by MoneyPuck.com (free for non-commercial use; credit them wherever shown).
    We only download files listed on https://moneypuck.com/data.htm.
"""

import sys
import time
from pathlib import Path

import pandas as pd      # pandas = "Excel for Python": its main object (a DataFrame) is a worksheet
import requests          # requests = the tool that fetches things from the internet

# ----------------------------------------------------------------------------
# SETTINGS  (change these, then run again)
# ----------------------------------------------------------------------------
# MoneyPuck names a season by the year it STARTS: 2025 means the 2025-26 season.
# This is a list, like a column of cells. Later we'll add more seasons to it.
SEASONS = [2025]

URL = "https://moneypuck.com/moneypuck/playerData/seasonSummary/{season}/regular/teams.csv"

# Where the files get saved. "data" is ignored by git, so it never goes to GitHub.
PROJECT_FOLDER = Path(__file__).resolve().parent.parent   # the "Model 2.0" folder
SAVE_FOLDER = PROJECT_FOLDER / "data" / "moneypuck"

# Tell the website who we are (a plain description, nothing personal).
HEADERS = {"User-Agent": "nhl-model-hobby-project (personal, non-commercial)"}


# ----------------------------------------------------------------------------
# A "function" is a named recipe, like a custom Excel formula. This one downloads one season.
# ----------------------------------------------------------------------------
def download_season(season):
    url = URL.format(season=season)          # puts the season number into the web address
    print(f"Downloading {season}-{season + 1} team stats from MoneyPuck...")

    try:
        response = requests.get(url, headers=HEADERS, timeout=60)
        response.raise_for_status()           # stops with an error if the site said "not found" etc.
    except requests.RequestException as problem:
        print(f"  Could not download it: {problem}")
        print("  Check your internet connection, or the file may have moved on MoneyPuck's data page.")
        return None

    # MoneyPuck sends automatic downloads to a "data license" page instead of the CSV.
    # (A web page starts with "<", a CSV does not.) We respect that and do not work around it.
    if "data_license" in response.url or response.text.lstrip().startswith("<"):
        SAVE_FOLDER.mkdir(parents=True, exist_ok=True)
        print("  MoneyPuck did not send the data file. It treats automatic downloads as scraping and")
        print("  asks for a data license first (see moneypuck.com/data_license.htm).")
        print("  Your options:")
        print("    1) Download the file by hand from https://moneypuck.com/data.htm in your browser and")
        print(f"       save it as  {SAVE_FOLDER / f'teams_{season}.csv'}")
        print("    2) Email MoneyPuck (address on that page) to ask for a data license.")
        return None

    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)       # make the folder if it isn't there yet
    save_path = SAVE_FOLDER / f"teams_{season}.csv"
    try:
        save_path.write_bytes(response.content)          # save the file exactly as received
    except PermissionError:
        print(f"  Could not save {save_path.name}. Is it open in Excel? Close it and run again.")
        return None

    print(f"  Saved: {save_path}")
    return save_path


def summarize(csv_path):
    """Open the saved file as a DataFrame (a worksheet) and print a quick summary."""
    sheet = pd.read_csv(csv_path)
    print(f"  Rows: {len(sheet)}   Columns: {len(sheet.columns)}")
    print(f"  First columns: {', '.join(sheet.columns[:12])} ...")
    if "situation" in sheet.columns:
        print("  Rows per game situation (one row per team for each):")
        print(sheet["situation"].value_counts().to_string())
    if "team" in sheet.columns:
        print(f"  Different teams: {sheet['team'].nunique()}")


# ----------------------------------------------------------------------------
# main() is where the script starts, like the first line of a macro.
# ----------------------------------------------------------------------------
def main():
    saved_any = False
    for season in SEASONS:
        path = download_season(season)
        if path:
            summarize(path)
            saved_any = True
        time.sleep(1)                        # be polite: a one-second pause between downloads
    print("Data: MoneyPuck.com")
    return 0 if saved_any else 1


if __name__ == "__main__":
    sys.exit(main())

"""
import_moneypuck_downloads.py  -  find the MoneyPuck files you downloaded and file them in the project
========================================================================================================
WHY THIS EXISTS:
    MoneyPuck does not allow scripts to download its files, so YOU click each download link in your
    browser (their rules: free for non-commercial use, credit MoneyPuck.com wherever the data is shown).
    Your browser saves the files in your Downloads folder, often with confusing names: every season's team
    file is called teams.csv, so the browser renames them "teams (1).csv" and so on.

WHAT IT DOES (in Excel terms):
    Looks through your Downloads folder for MoneyPuck files, works out what each one is by reading the
    column headings (not by trusting the file name), COPIES it into  data/moneypuck/  with a clear name
    (for example teams_2025.csv), unzips the zip files into their own folders, and prints a checklist of
    what you have and what is still missing. Your Downloads folder is only read, never changed.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/import_moneypuck_downloads.py
    Run it again after each batch of downloads: it only adds what is new.
    Optional:  --downloads "D:\\somewhere\\else"   (if your browser saves files elsewhere)

WHAT YOU SHOULD SEE:
    A checklist with a tick or a "missing" for each file, plus a plain-English line saying what to
    download next. Files end up in  data/moneypuck/  (ignored by git: never uploaded to GitHub).

DATA CREDIT:
    Data is from MoneyPuck.com, used here for non-commercial purposes. Please credit MoneyPuck.com wherever
    anything built from it is shown.
"""

import argparse
import csv
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
TARGET_FOLDER = PROJECT_FOLDER / "data" / "moneypuck"        # "data" is ignored by git
DAYS_BACK = 21                                                # only look at files saved recently

# What we want. kind: "name" = recognized by its file name, "season" = recognized by its contents.
WANTED = [
    {"key": "teams_hist", "kind": "name", "name": "teams_2008_to_2024.zip", "target": "teams_2008_to_2024.zip",
     "what": "Team seasons 2008-09 to 2024-25", "where": "row 'All Previous Seasons Combined', icon under Team Level"},
    {"key": "goalies_hist", "kind": "name", "name": "goalies_2008_to_2024.zip", "target": "goalies_2008_to_2024.zip",
     "what": "Goalie seasons 2008-09 to 2024-25", "where": "row 'All Previous Seasons Combined', icon under Goalies"},
    {"key": "teams_2025", "kind": "season", "table": "teams", "season": 2025, "target": "teams_2025.csv",
     "what": "Team season 2025-26", "where": "row '2025-2026', icon under Team Level"},
    {"key": "goalies_2025", "kind": "season", "table": "goalies", "season": 2025, "target": "goalies_2025.csv",
     "what": "Goalie season 2025-26", "where": "row '2025-2026', icon under Goalies"},
    {"key": "teams_2026", "kind": "season", "table": "teams", "season": 2026, "target": "teams_2026.csv",
     "what": "Team season 2026-27 (so far)", "where": "row '2026-2027', icon under Team Level"},
    {"key": "goalies_2026", "kind": "season", "table": "goalies", "season": 2026, "target": "goalies_2026.csv",
     "what": "Goalie season 2026-27 (so far)", "where": "row '2026-2027', icon under Goalies"},
    {"key": "all_teams", "kind": "name", "name": "all_teams.csv", "target": "all_teams.csv",
     "what": "Every team's every game (the big one)", "where": "text link 'here' under Game By Game Level Data"},
    {"key": "dictionary", "kind": "prefix", "name": "MoneyPuckDataDictionary", "target": "MoneyPuckDataDictionaryForPlayers.csv",
     "what": "Data dictionary (what every column means)", "where": "text link 'here' under Data Dictionary"},
]


def clean_name(filename):
    """'teams (1).csv' -> 'teams.csv': the browser adds ' (1)' when a name is already taken."""
    return re.sub(r" \(\d+\)(?=\.[A-Za-z0-9]+$)", "", filename)


def read_header(path):
    """The column names of a CSV, and the first value of its 'season' column (if it has one)."""
    try:
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as handle:
            reader = csv.reader(handle)
            header = next(reader, [])
            first = next(reader, [])
    except OSError:
        return [], None
    season = None
    if "season" in header and first and len(first) == len(header):
        try:
            season = int(float(first[header.index("season")]))
        except ValueError:
            season = None
    return header, season


def classify_columns(header):
    """Is this a MoneyPuck 'teams' file or a 'goalies' file? Looked at from its column headings."""
    cols = set(header)
    if "playerId" in cols and "xGoals" in cols:
        return "goalies"
    if "team" in cols and "xGoalsFor" in cols and "playerId" not in cols:
        return "teams"
    return None


def find_matches(downloads, days_back):
    """Look through the Downloads folder. Returns ({key: path}, [paths we could not identify])."""
    cutoff = time.time() - days_back * 86400
    found, unknown = {}, []
    if not downloads.exists():
        return found, unknown
    for path in sorted(downloads.iterdir(), key=lambda p: p.stat().st_mtime):    # oldest first: newest wins
        if not path.is_file() or path.suffix.lower() not in (".csv", ".zip") or path.stat().st_mtime < cutoff:
            continue
        name = clean_name(path.name)
        matched = False
        for item in WANTED:
            if item["kind"] == "name" and name.lower() == item["name"].lower():
                found[item["key"]] = path
                matched = True
            elif item["kind"] == "prefix" and name.lower().startswith(item["name"].lower()):
                found[item["key"]] = path
                matched = True
        if matched or path.suffix.lower() != ".csv":
            continue
        header, season = read_header(path)
        table = classify_columns(header)
        if table and season:
            key = f"{table}_{season}"
            if any(item["key"] == key for item in WANTED):
                found[key] = path
                continue
        if "xGoalsFor" in header or "xGoals" in header or "I_F_xGoals" in header or "highDangerShotsFor" in header:
            unknown.append(path)                      # looks like MoneyPuck, but not one we asked for
    return found, unknown


def copy_if_newer(source, destination):
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists() and destination.stat().st_mtime >= source.stat().st_mtime \
            and destination.stat().st_size == source.stat().st_size:
        return False
    shutil.copy2(source, destination)
    return True


def unzip_csvs(zip_path, folder):
    """Unzip only the .csv files inside, into their own folder (python removes any '..' tricks from names)."""
    folder.mkdir(parents=True, exist_ok=True)
    extracted = []
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            if member.filename.lower().endswith(".csv") and not member.is_dir():
                archive.extract(member, folder)
                extracted.append(member.filename)
    return extracted


def describe_csv(path):
    """A one-line description: size, and the seasons inside if there is a 'season' column."""
    size = f"{path.stat().st_size / 1e6:.1f} MB"
    header, _ = read_header(path)
    if "season" not in header:
        return size
    try:
        import pandas as pd
        seasons = pd.read_csv(path, usecols=["season"], encoding="utf-8-sig")["season"].dropna().astype(int)
        return f"{size}, seasons {seasons.min()}-{seasons.max()}, {seasons.nunique()} season(s)"
    except Exception:
        return size


def main():
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="File the MoneyPuck downloads in data/moneypuck.")
    parser.add_argument("--downloads", help="folder your browser saves downloads in (default: your Downloads folder)")
    args = parser.parse_args()
    downloads = Path(args.downloads) if args.downloads else Path.home() / "Downloads"
    TARGET_FOLDER.mkdir(parents=True, exist_ok=True)

    print(f"Looking in {downloads} for MoneyPuck files saved in the last {DAYS_BACK} days...")
    found, unknown = find_matches(downloads, DAYS_BACK)

    copied = []
    for item in WANTED:
        source = found.get(item["key"])
        if source is not None and copy_if_newer(source, TARGET_FOLDER / item["target"]):
            copied.append(item["target"])
    for target in copied:
        print(f"  Filed: {target}")

    # Unzip the history zips into their own folders (the zip itself stays too)
    for item in WANTED:
        zip_path = TARGET_FOLDER / item["target"]
        if item["target"].endswith(".zip") and zip_path.exists():
            folder = TARGET_FOLDER / zip_path.stem
            if item["target"] in copied or not any(folder.glob("*.csv")):
                names = unzip_csvs(zip_path, folder)
                print(f"  Unzipped {zip_path.name}: {', '.join(names) if names else 'no csv files found inside!'}")

    print()
    print("CHECKLIST")
    missing = []
    for item in WANTED:
        have = TARGET_FOLDER / item["target"]
        if have.exists():
            extra = describe_csv(have) if have.suffix.lower() == ".csv" else f"{have.stat().st_size / 1e6:.1f} MB"
            print(f"  [x] {item['what']:<44} {extra}")
        else:
            print(f"  [ ] {item['what']:<44} MISSING")
            missing.append(item)

    if unknown:
        print()
        print("These look like MoneyPuck files but are not on today's list (left alone):")
        for path in unknown:
            print(f"  {path.name}")

    print()
    if not missing:
        print("Everything on the list is here. Tell Claude: the next step is the loader and the goals check.")
    else:
        print("STILL TO DOWNLOAD (from https://moneypuck.com/data.htm):")
        for item in missing:
            print(f"  - {item['what']}: {item['where']}")
        print("Then run this script again.")
    print(f"\nFiles are in {TARGET_FOLDER}")
    print("Data: MoneyPuck.com (non-commercial use, credit required)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

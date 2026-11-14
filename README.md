# NHL Model 2.0

A hockey analytics project that rates NHL **teams** and **players** in different areas (offense, defense, special teams, goaltending, chance creation, two-way impact) and compares today's teams and players with history. The goal is to understand the sport better. Results are published as a GitHub Pages website.

> **Status:** early build. Step 1 is exploring the data and choosing which metrics matter.

## How it works (plain English)

1. **Scripts** in `scripts/` download public hockey data into a local `data/` folder.
2. Other scripts calculate ratings and compare today's teams and players with past seasons.
3. A site-builder script turns the results into web pages in `docs/`.

## Updating the website

The site is the `docs/` folder, published with GitHub Pages. To refresh the home page (tonight's games and
standings):

```
python scripts/pull_schedule_standings.py
python scripts/build_site.py
```

To also refresh the **Contender fit** page (the current-season ranking comes from the latest games), update the
game data and the database first. This takes a few minutes because it waits politely between requests:

```
python scripts/pull_game_results.py
python scripts/build_database.py
python scripts/build_contender_fit.py
python scripts/pull_schedule_standings.py
python scripts/build_site.py
```

Then commit and push the changed files in `docs/`. The pages show summary results only, never the raw data files.

## Rebuilding the data

The `data/` folder is **not stored on GitHub** (it is listed in `.gitignore`). To recreate it:

1. Install Python 3 and run `pip install -r requirements.txt`.
2. Run the scripts in `scripts/` (each one explains itself at the top of the file). For example,
   `python scripts/build_reference_tables.py` builds `outputs/playoff_history.xlsx`: every playoff
   series, champion and conference finalist since 2005-06, and how far every team went each season.
3. Download every team's season statistics (about 11 minutes, it waits politely between requests),
   the NHL's tracking-based zone time (2021-22 on), and build a data coverage report that shows
   which statistic is really available from which season:
   ```
   python scripts/pull_team_season_reports.py
   python scripts/pull_edge_zone_time.py
   python scripts/build_coverage_report.py
   ```
   The report is `outputs/data_coverage.xlsx`: start with its "Key metrics" sheet.
4. Download every team's result in every game (about 6 minutes), work out the checkpoint dates
   (Nov 1, Thanksgiving, Jan 1, trade deadline, end of season) and the standings on each, load it all
   into one database file (`data/nhl.db`), and build a workbook showing who was in a playoff spot at
   each checkpoint and whether they made it:
   ```
   python scripts/pull_game_results.py
   python scripts/build_checkpoints.py
   python scripts/build_database.py
   python scripts/build_checkpoint_report.py
   ```
   The workbook is `outputs/checkpoint_standings.xlsx`: start with its "First look" sheet.
   `reference/trade_deadlines.csv` and `reference/franchises.csv` are small hand-checked lists you can
   edit in Excel. Some deadline dates are marked "from memory - please check".
5. Test the popular stats (plus/minus, PDO, Corsi, hits, shooting %, save %, power play, penalty kill):
   does each one explain results, repeat from one part of the season to the next, and help predict
   the rest of the season beyond the record?
   ```
   python scripts/build_popular_stats_check.py
   ```
   The workbook is `outputs/popular_stats_check.xlsx`: start with its "Verdicts" sheet.
6. Build the contender-fit score (version 1, NHL stats only) and test it honestly: does it pick the
   real conference finalists better than simply ranking by points %? It also shows how champions
   looked on each stat, a "top 5 in K factors" rule test, and a current-season ranking:
   ```
   python scripts/build_contender_fit.py
   ```
   The workbook is `outputs/contender_fit.xlsx`: start with its "Verdict" sheet.

## Data sources and credits

- **NHL**: schedule, standings and statistics come from the NHL's public, unofficial web API. NHL data is © NHL; it is used here for personal, non-commercial purposes. The API is undocumented and can change without notice.
- **[MoneyPuck.com](https://moneypuck.com/data.htm)**: expected-goals and related data are provided by MoneyPuck.com and used here for non-commercial purposes. Only the files posted on MoneyPuck's data page are used. Please credit MoneyPuck.com wherever their data is shown.

This is an independent hobby project. It is not affiliated with or endorsed by the NHL, any NHL club, or MoneyPuck.

## Disclaimer

For information and entertainment only. Nothing here is financial or betting advice.

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

The research pages (Contender fit, Red flags, Slow starters, Goaltending in context and the Studies list) are built
from the workbooks the studies write in `outputs/`. To refresh them for the current season, rerun the study scripts
listed in the sections below (`build_contender_fit.py`, `build_red_flags.py`, `build_slow_starters.py`,
`build_goaltending_study.py`) and then `python scripts/build_site.py`. A page whose workbook is missing is skipped and
its existing page in `docs/` is left as it is.

The **Breakouts** page has a "This season" table. To bring it up to date, download the play-by-play of the new games
(about a second a game) and rebuild the table. Nothing is searched or re-fitted: the two frozen recipes are applied as
they are.

```
python scripts/pull_game_results.py
python scripts/build_database.py
python scripts/load_moneypuck.py
python scripts/pull_play_by_play.py
python scripts/build_breakout_now.py
python scripts/build_site.py
```

`build_breakout_now.py` writes `outputs/breakout_now.xlsx`. A game that the MoneyPuck game file does not have yet uses
5-on-5 time counted from the events instead, so the table does not wait for a new MoneyPuck download.

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
   looked on each stat, a "top 5 in K factors" rule test (on the website each rule can be clicked to see
   which of today's teams fit it, with their rank on each of the four factors), and a current-season ranking:
   ```
   python scripts/build_contender_fit.py
   ```
   The workbook is `outputs/contender_fit.xlsx`: start with its "Verdict" sheet.
7. Test six red flags (running hot, weak shot share, record flatters goals, weak special teams, penalty
   trouble, cold lately) on teams that were in a playoff spot: do flagged teams miss the playoffs more often
   than equally safe teams, and does stacking flags raise the risk? Also lists today's in-spot teams with
   their flags (this asks the NHL for one standings page):
   ```
   python scripts/build_red_flags.py
   ```
   The workbook is `outputs/red_flags.xlsx`: start with its "Verdict" sheet.
8. The slow-starter study: for teams that were out of a playoff spot at Nov 1, Thanksgiving, Jan 1 or the
   deadline, how often do they recover (by how far behind), what points pace does it take, do six "signs of a
   coming recovery" separate the ones that do, and who are today's slow starters? (asks the NHL for one
   standings page):
   ```
   python scripts/build_slow_starters.py
   ```
   The workbook is `outputs/slow_starters.xlsx`: start with its "Verdict" sheet.
9. Goaltending in context (needs the MoneyPuck tables from `load_moneypuck.py`): splits each team's goals
   against into defense (shot quality allowed) and goaltending (goals saved above expected), labels teams
   "goalie bailing out the defense" / "defense let down by the goalie", tests which part repeats and whether the
   labels predict the rest of the season, and asks whether anything foresees a hot playoff goalie:
   ```
   python scripts/build_goaltending_study.py
   ```
   The workbook is `outputs/goaltending_in_context.xlsx`: start with its "Verdict" sheet.
10. Contender fit version 2 (needs the MoneyPuck tables): adds expected-goals share, high-danger shot share
    and goaltending in context to version 1 and tests, with the same hold-out method, whether any of them picks
    the real conference finalists better. It also checks why high-danger share is such a noisy statistic:
    ```
    python scripts/build_contender_fit_v2.py
    ```
    The workbook is `outputs/contender_fit_v2.xlsx`: start with its "Verdict" sheet.
11. Breakout Index (needs the play-by-play download and the MoneyPuck tables): builds three "getting out of your own
    end" statistics from the play-by-play at 5-on-5 (zone exit time, costly giveaways, transition shots), blends them
    into one index, and tests whether it repeats, whether it lines up with the NHL's real zone time and with winning,
    which mix of weights for the three works best (chosen on all seasons but one, scored on the one left out),
    and whether it improves the contender score on seasons the score has never seen. It also scores five more
    candidate stats (icings, pinned-in spells, own-zone turnovers, defence-to-offence rate, transition shot quality)
    against a pass rule written down in advance: repeats, points the right way, and adds something beyond
    shot-attempt share in the other half of the season. It uses only seasons that are
    fully downloaded, so it can be run while the download is still going and rerun when it has finished:
    ```
    python scripts/build_breakout_index.py
    ```
    The workbook is `outputs/breakout_index.xlsx`: start with its "Verdict" sheet. The definitions (the 20-second exit
    cap, the 10-second costly-giveaway window, the 5-second transition window, the three weights) are settings at the
    top of the script; `--weights 0.5,0.2,0.3` tries other weights for one run.

12. Do breakouts make teams good? (needs the same data as step 11): takes 11 breakout stats, makes every
    combination of them (2,047), and asks whether the best one goes with winning, predicts games it has not seen,
    carries into next season, and is more than shot share. To avoid fooling ourselves, the best combination is
    always chosen on all seasons but one and scored on the one left out, and the whole search is repeated on
    1,000 random shuffles of the results to measure luck. Tougher checks repeat it with puck-movement stats only
    (no shot counts) and after taking out shot quality:
    ```
    python scripts/build_breakout_search.py
    ```
    The workbook is `outputs/breakout_search.xlsx`: start with its "Verdict" sheet.
13. The search for the right breakout metric, round 2, and the blind test. Every stat is also built from close-score
    moments only and as a "battle" (ours minus the opponent's in the same games), with a few tuned time windows:
    plus BER (100 x (controlled exits + 0.5 x safe clears) / breakout attempts, read from events as well as they
    allow): 68 building blocks, every combination of up to four (about 867,000), plus a data-weighted version. The best
    recipe for "our own breakouts" and the best "breakout battle" recipe are frozen in
    `reference/breakout_metric_frozen.json`:
    ```
    python scripts/build_breakout_round2.py
    ```
    The workbook is `outputs/breakout_round2.xlsx`. Five seasons (2009-10 to 2013-14) are kept out of every breakout
    script (the `LOCKBOX` list in `build_breakout_index.py`) so that the frozen recipes can be scored on them once,
    when they have all downloaded:
    ```
    python scripts/build_breakout_round2.py --blind-test
    ```
    That writes `outputs/breakout_blind_test.xlsx` and records the result in the frozen file. Finish line: the recipe
    predicts games it has not seen at 0.46 or better (Corsi's level); to count as a separate edge it must also add
    something after Corsi and expected goals are taken out.

## The play-by-play download (overnight)

Every NHL game has a play-by-play: one row per event (faceoff, shot, hit, giveaway, takeaway, penalty, goal) with the
time, the players and, for most events, where on the ice it happened and in which zone. This feeds the zone-time and
breakout estimates and a shot-location definition of our own. It is about 22,000 games at one request a second
(roughly 6 hours), so it is built to run unattended:

```
python scripts/pull_play_by_play.py            start (or resume) the download
python scripts/pull_play_by_play.py --status   how far along is it, and when will it finish?
```

It saves into one database file (`data/pbp.db`, about 1-2 GB, never uploaded), skips games it already has, goes
newest seasons first, backs off if the NHL starts refusing, and asks Windows not to sleep while it runs (closing a
laptop lid still sleeps the PC). To stop cleanly, create an empty file named `STOP` in `data/pbp/`; to start again,
just run the command again.

## Downloading the MoneyPuck files (by hand)

MoneyPuck does not allow scripts to download its files, so you click the links yourself in your browser
(free for non-commercial use; credit MoneyPuck.com wherever the data is shown). On
<https://moneypuck.com/data.htm>, in the **Season Level Data** table, download the file under:

| Row | Click the icon under | What it is |
|---|---|---|
| All Previous Seasons Combined | Team Level, Goalies | team and goalie seasons 2008-09 to 2024-25 (two zip files) |
| 2025-2026 | Team Level, Goalies | last season (two csv files) |
| 2026-2027 | Team Level, Goalies | this season so far (two csv files; re-download to refresh) |

Also download the text link **here** under *Game By Game Level Data* (`all_teams.csv`, every team's every game)
and the **data dictionary** link at the top of the page. Your browser saves them in your Downloads folder (with
names like `teams (1).csv`: that is fine). Then run:

```
python scripts/import_moneypuck_downloads.py
```

It recognizes each file by its contents, copies it into `data/moneypuck/` with a clear name, unzips the zips,
and prints a checklist of anything still missing. Run it again after each batch of downloads. Then add the
MoneyPuck tables to the database and check them against the NHL's numbers (goals, 5v5 goals, shot attempts):

```
python scripts/load_moneypuck.py
```

(`build_database.py` rebuilds the database from scratch, so run `load_moneypuck.py` again after it.)

## Data sources and credits

- **NHL**: schedule, standings and statistics come from the NHL's public, unofficial web API. NHL data is © NHL; it is used here for personal, non-commercial purposes. The API is undocumented and can change without notice.
- **[MoneyPuck.com](https://moneypuck.com/data.htm)**: expected-goals and related data are provided by MoneyPuck.com and used here for non-commercial purposes. Only the files posted on MoneyPuck's data page are used. Please credit MoneyPuck.com wherever their data is shown.

This is an independent hobby project. It is not affiliated with or endorsed by the NHL, any NHL club, or MoneyPuck.

## Disclaimer

For information and entertainment only. Nothing here is financial or betting advice.

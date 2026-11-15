"""
build_breakout_search.py  -  do breakouts make teams good? Every combination of breakout stats, tested honestly
==============================================================================================================
WHAT IT DOES (in Excel terms):
    We believe that getting out of your own end cleanly is what makes a team good. This script takes the 11
    breakout statistics we can build from the play-by-play, makes EVERY combination of them (2,047 different
    "breakout scores", each the plain average of the stats in it), and asks four questions of each one:

      1. DO GOOD-BREAKOUT TEAMS WIN?        the score next to points % and goal difference, same season
      2. DOES IT PREDICT UNSEEN GAMES?      the score from odd-numbered games next to the RESULTS of
                                            even-numbered games (and the reverse)
      3. IS IT MORE THAN SHOT SHARE?        question 2 again, after taking out what Corsi already explains
      4. DOES IT CARRY INTO NEXT SEASON?    the score this season next to next season's results

    THREE TOUGHER CHECKS (added after the first results came out strong, to make sure we are not fooling ourselves)
      * PUCK-MOVEMENT STATS ONLY   four of the 11 stats are counts of shot attempts, so they are close relatives of
                                   Corsi. Every question is answered again with the other seven alone.
      * SHOT QUALITY               question 5 repeats question 3 after also taking out expected-goals share.
      * THE ECHO                   half a season of Corsi is itself a noisy reading, so ANY second reading of the same
                                   thing seems to "add" a little. The workbook shows how big that echo would be.

    THE TRAP, AND HOW THIS AVOIDS IT
    Test 2,047 combinations and one of them will always look great by luck. Two protections, both fixed in advance:
      * HOLD-OUT    the best combination is CHOSEN on all seasons but one and SCORED on the season left out.
                    Only that left-out score is reported as the result.
      * SHUFFLE     the teams' results are shuffled at random (so there is truly nothing to find) and the whole
                    search is run again, 1,000 times. That shows how good "the best of 2,047" looks by pure luck,
                    and gives the odds that our real result is luck.

    THE RULE FOR "CONFIRMED" (written before any result was looked at)
      A question is answered YES only if the left-out score is above zero by more than its noise margin AND the
      shuffle test puts the odds of luck under 5%.
      The belief counts as SUPPORTED if questions 1 and 2 are both YES. It counts as MORE THAN SHOT SHARE only if
      question 3 is YES too. Corsi alone is shown next to every answer as the yardstick.

    THE 11 STATS (all 5-on-5; built by build_breakout_index.py, where each definition is written out)
      zone exit time, share of own-zone starts that cleared the zone, share of exits within 8 seconds, costly
      giveaways per 60, share of own-zone giveaways that were costly, transition shot attempts per 60, icings per 60,
      pinned-in spells per 60, own-zone turnovers per 60, defence-to-offence rate, transition shot quality.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_breakout_search.py
    Safe to run while the download is still going; it uses the seasons that are fully downloaded so far.
    Options:  --seasons 20232024,20242025     only those seasons
              --shuffles 200                  fewer shuffles (faster, rougher odds)

WHAT YOU SHOULD SEE:
    A summary in the terminal and a new file:  outputs/breakout_search.xlsx   (close it in Excel first).

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. 5-on-5 ice time and shot attempts for the
    Corsi yardstick come from MoneyPuck.com (non-commercial use, credit required).
"""

import argparse
import sqlite3
import sys

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_breakout_index as bi         # builds the breakout stats from the play-by-play
import build_popular_stats_check as ps    # shared stat helpers

OUTPUT_FILE = bi.PROJECT_FOLDER / "outputs" / "breakout_search.xlsx"

POOL = [  # (column, short name, higher is better?)
    ("exit_time", "Zone exit time", False),
    ("exit_success", "Zone-clear rate", True),
    ("quick_exit", "Quick exits", True),
    ("costly_pg", "Costly giveaways", False),
    ("costly_share", "Costly share of giveaways", False),
    ("trans_pg", "Transition shots", True),
    ("icing_pg", "Icings", False),
    ("pinned_pg", "Pinned-in spells", False),
    ("turnover_pg", "Own-zone turnovers", False),
    ("d2o_rate", "Defence-to-offence", True),
    ("trans_slot_share", "Transition shot quality", True),
]
Z = [f"z_{c}" for c, _, _ in POOL]
NAMES = [n for _, n, _ in POOL]
CURRENT_INDEX = ["exit_time", "costly_pg", "trans_pg"]   # the three stats in the Breakout Index today
# Four of the stats are COUNTS OF SHOT ATTEMPTS (so they are close relatives of Corsi). The other seven count where the
# puck went or who lost it. The search is also run on those seven alone ("puck-movement stats only"): a tougher check.
SHOT_BASED = ["trans_pg", "pinned_pg", "d2o_rate", "trans_slot_share"]
N_SHUFFLES = 1000
ODDS_BAR = 0.05          # the shuffle test must put the odds of luck under this
SEED = 20261114          # so the shuffles (and the answers) are the same every run
TARGETS = ["pts_pct", "gd_pg"]

QUESTIONS = [  # (key, the question, what the score is compared with)
    ("same", "1. Do good-breakout teams win?", "points % and goal difference, same season"),
    ("other", "2. Does it predict unseen games?", "results of the OTHER half of the season's games"),
    ("beyond", "3. Is it more than shot share?", "the same, after taking out what Corsi explains"),
    ("next", "4. Does it carry into next season?", "points % and goal difference NEXT season"),
    ("beyond2", "5. Is it more than shot share AND shot quality?", "as question 3, also taking out expected-goals share"),
]

READ_ME = [
    "Do breakouts make teams good? Every combination of 11 breakout stats, tested honestly",
    "",
    "WHAT WAS DONE",
    "  Every combination of the 11 breakout stats (2,047 of them) was turned into a score: each stat is put on the same scale within its season (steps above or below the",
    "  league average, turned so that higher is always better) and the stats in the combination are averaged. Each score was then asked four questions (sheet 'Questions').",
    "",
    "HOW TO READ THE NUMBERS",
    "  Every number is a rank correlation, like CORREL on the teams' rankings, worked out inside each season and averaged: 0 = no relationship, 1 = perfect.",
    "  'Honest result'  the best combination was chosen WITHOUT one season and scored on that season; this is the average of those left-out scores.",
    "  'Noise margin'   two standard errors across the seasons. A result smaller than its margin could be zero.",
    "  'Odds it is luck'  the teams' results were shuffled at random and the whole search run again 1,000 times. This is the share of those runs that did as well as the real one.",
    "  'Luck line'      what the search reaches one time in twenty when there is truly nothing to find.",
    "  'Corsi alone'    shot-attempt share at 5-on-5, the yardstick: the plainest measure of who controls play.",
    "  'Puck-movement stats only'  four of the 11 stats count shot attempts (transition shots, pinned-in spells, defence-to-offence, transition shot quality), so they are",
    "                   close relatives of Corsi. These columns repeat the whole search with the other seven only: exits, giveaways, turnovers, icings.",
    "  'Echo'           half a season of Corsi is a noisy reading, so any second reading of the same thing seems to add a little to it. This is how much (an upper estimate).",
    "  Question 5       question 3 again, after also taking out expected-goals share (shot quality). If the answer turns to 'no', what the combination added was shot quality.",
    "",
    "THE RULE (fixed before the results were looked at)",
    f"  A question is answered YES only if the honest result is above zero by more than its noise margin AND the odds of luck are under {ODDS_BAR:.0%}.",
    "  The belief is SUPPORTED if questions 1 and 2 are YES; it is MORE THAN SHOT SHARE only if question 3 is YES too.",
    "",
    "WHAT THIS CANNOT SAY",
    "  These are event counts, not puck tracking: nobody recorded passes, carries or who had the puck. A real breakout (a controlled exit under pressure) is not in the data,",
    "  so a weak result here can mean 'breakouts matter less than we think' OR 'we cannot see breakouts well enough'. And a link is not proof of cause: good teams may",
    "  break out well because they are good, rather than the other way round.",
    "",
    "Data: (c) NHL. Corsi yardstick and 5-on-5 ice time: MoneyPuck.com.",
]


# ----------------------------------------------------------------------------
# Tables: the 11 stats per team, on the same scale
# ----------------------------------------------------------------------------
def add_pool_z(t):
    """Each stat as steps above or below the league average that season, turned so higher = better."""
    t = t.copy()
    for col, _, higher in POOL:
        z = ps.zscore_by_season(t, col)
        t[f"z_{col}"] = (z if higher else -z).fillna(0.0)
    return t


def half_tables(game_rows, nhl, seasons):
    """Two tables: the stats from odd-numbered games next to the RESULTS of even-numbered games, and the reverse."""
    ids = ",".join(str(int(s)) for s in seasons)
    results = pd.read_sql_query(
        f"SELECT season_id, game_id, team_id, points, goalsFor - goalsAgainst AS gd FROM team_games WHERE game_type = 2 AND season_id IN ({ids})", nhl)
    bench = bi.benchmark_halves(nhl, seasons)                            # corsi_a = odd games, corsi_b = even games
    frames = []
    for parity, suffix in ((1, "_a"), (0, "_b")):
        z = add_pool_z(bi.summarize(game_rows[game_rows["game_id"] % 2 == parity]))[["season_id", "team_id"] + Z]
        r = results[results["game_id"] % 2 != parity].groupby(["season_id", "team_id"]).agg(
            pts=("points", "sum"), n=("points", "size"), gd=("gd", "sum")).reset_index()
        r["pts_pct"], r["gd_pg"] = r["pts"] / (2 * r["n"]), r["gd"] / r["n"]
        c = bench[["season_id", "team_id", f"corsi{suffix}", f"xg_share{suffix}"]].rename(columns={f"corsi{suffix}": "corsi", f"xg_share{suffix}": "xg_share"})
        frames.append(z.merge(r[["season_id", "team_id", "pts_pct", "gd_pg"]], on=["season_id", "team_id"]).merge(c, on=["season_id", "team_id"]))
    return frames


def next_season_table(team):
    """Each team's stats this season next to its results NEXT season (only where the next season is also in the data)."""
    following = team[["season_id", "team_id"] + TARGETS].copy()
    following["season_id"] = following["season_id"] - 10001              # 20232024 becomes 20222023: lines up with the season before
    return team[["season_id", "team_id", "corsi"] + Z].merge(following, on=["season_id", "team_id"])


# ----------------------------------------------------------------------------
# The engine: the link of all 2,047 combinations at once
# ----------------------------------------------------------------------------
def all_combinations():
    """One row per combination, one column per stat: 1 if the stat is in it."""
    m = len(POOL)
    return np.array([[(k >> j) & 1 for j in range(m)] for k in range(1, 2 ** m)], dtype=float)


def unit(a):
    """Centre each column and scale it to length 1, so that  column . column  is a correlation."""
    a = a - a.mean(axis=0)
    return a / np.sqrt((a ** 2).sum(axis=0))


def take_out(a, controls):
    """What is left of each column of `a` after the best straight-line fit on the control columns is removed."""
    return unit(a - controls @ (np.linalg.pinv(controls) @ a))


def make_part(frame, combos, controls=None):
    """
    Prepares one season (or one half of one) so that  X.T @ Y  gives the rank correlation of every combination with
    each target. With `controls` (e.g. ["corsi"]), both sides first have those taken out (a partial correlation).
    """
    d = frame.dropna(subset=TARGETS + ["corsi"] + (controls or []))
    scores = d[Z].to_numpy() @ combos.T
    x = unit(pd.DataFrame(scores).rank().to_numpy())
    y = unit(d[TARGETS].rank().to_numpy())
    c = unit(d[["corsi"]].rank().to_numpy())
    part = {"corsi": float((c.T @ y).mean()), "overlap": (c.T @ x).ravel(), "xg_beyond": np.nan}   # overlap = each combination's link with Corsi itself
    if controls:
        k = unit(d[controls].rank().to_numpy())
        if controls == ["corsi"] and "xg_share" in d:                   # yardstick: what does expected-goals share add beyond Corsi on the same test?
            part["xg_beyond"] = float((take_out(unit(d[["xg_share"]].rank().to_numpy()), k).T @ take_out(y, k)).mean())
        x, y = take_out(x, k), take_out(y, k)
    part["X"], part["Y"] = x, y
    return part


def fold_links(parts, rng=None):
    """The link of every combination for one season: averaged over the targets and over the parts (e.g. odd->even and even->odd)."""
    out = []
    for part in parts:
        y = part["Y"] if rng is None else part["Y"][rng.permutation(len(part["Y"]))]     # shuffled: teams get someone else's results
        out.append((part["X"].T @ y).mean(axis=1))
    return np.mean(out, axis=0)


def honest(m):
    """Choose the best combination on all seasons but one, score it on the season left out. Returns (left-out scores, combination chosen)."""
    total = m.sum(axis=0)
    chosen = np.array([int((total - m[i]).argmax()) for i in range(len(m))])
    return m[np.arange(len(m)), chosen], chosen


def combo_name(row):
    return " + ".join(n for n, used in zip(NAMES, row) if used)


def run_question(folds, labels, combos, rng, n_shuffles):
    """
    Everything about one question: the honest result, its noise margin, the yardsticks and the shuffle test.
    Done twice: with all 11 stats allowed, and with only the seven puck-movement stats (no shot counts) allowed.
    """
    idx = {c: i for i, c in enumerate(c for c, _, _ in POOL)}
    movement = np.where(combos[:, [idx[c] for c in SHOT_BASED]].sum(axis=1) == 0)[0]       # combinations that use no shot-count stat
    m = np.vstack([fold_links(parts) for parts in folds])               # seasons x combinations
    n = len(m)
    held, chosen = honest(m)
    held_mv, pick = honest(m[:, movement])
    chosen_mv = movement[pick]
    luck = np.empty((n_shuffles, 2))
    for b in range(n_shuffles):
        shuffled = np.vstack([fold_links(parts, rng) for parts in folds])
        luck[b] = honest(shuffled)[0].mean(), honest(shuffled[:, movement])[0].mean()
    every = m.mean(axis=0)
    all_stats = int(np.where(combos.sum(axis=1) == len(POOL))[0][0])
    current = int(np.where((combos == np.isin(np.arange(len(POOL)), [idx[c] for c in CURRENT_INDEX]).astype(float)).all(axis=1))[0][0])
    singles = [int(np.where((combos.sum(axis=1) == 1) & (combos[:, j] == 1))[0][0]) for j in range(len(POOL))]

    def margin(v):
        return float(2 * v.std(ddof=1) / np.sqrt(n)) if n > 1 else float("nan")

    def odds(v, col):
        return float(((luck[:, col] >= v.mean()).sum() + 1) / (n_shuffles + 1))

    def overlap(picks):                                                  # how much the chosen combination moves with Corsi itself
        return float(np.mean([np.mean([p["overlap"][picks[i]] for p in parts]) for i, parts in enumerate(folds)]))

    return {"m": m, "held": held, "chosen": chosen, "labels": labels, "every": every, "n": n,
            "result": float(held.mean()), "margin": margin(held), "odds": odds(held, 0), "luck_line": float(np.quantile(luck[:, 0], 0.95)),
            "seasons_positive": int((held > 0).sum()), "overlap": overlap(chosen),
            "mv_held": held_mv, "mv_chosen": chosen_mv, "mv_result": float(held_mv.mean()), "mv_margin": margin(held_mv), "mv_odds": odds(held_mv, 1),
            "mv_luck_line": float(np.quantile(luck[:, 1], 0.95)), "mv_seasons_positive": int((held_mv > 0).sum()), "mv_overlap": overlap(chosen_mv),
            "mv_best": int(movement[every[movement].argmax()]),
            "corsi": float(np.mean([np.mean([p["corsi"] for p in parts]) for parts in folds])),
            "xg_beyond": float(np.nanmean([np.nanmean([p["xg_beyond"] for p in parts]) for parts in folds])) if not np.isnan(folds[0][0]["xg_beyond"]) else float("nan"),
            "all_stats": float(every[all_stats]), "current": float(every[current]), "singles": every[singles], "best": int(every.argmax())}


def echo(overlap, corsi_link, reliability):
    """
    If a score were ONLY a second, independent reading of the same thing Corsi measures, it would still seem to 'add' a
    little to Corsi, because half a season of Corsi is itself a noisy reading. This is how big that echo would be
    (an upper estimate: the two readings come from the same games, so they are not fully independent).
    """
    expected = overlap * corsi_link / reliability                        # the link with results the score would have through Corsi alone
    return (expected - overlap * corsi_link) / np.sqrt((1 - overlap ** 2) * (1 - corsi_link ** 2))


# ----------------------------------------------------------------------------
# Tables for the workbook
# ----------------------------------------------------------------------------
def is_yes(result, margin, odds):
    return result > margin and odds < ODDS_BAR


def questions_table(res, reliability):
    rows = []
    for key, question, against in QUESTIONS:
        if key not in res:
            continue
        r = res[key]
        controlled = key in ("beyond", "beyond2")
        rows.append({"Question": question, "The score is compared with": against, "Seasons": r["n"],
                     "Honest result (best combination, chosen on the other seasons)": round(r["result"], 3), "Noise margin (+/-)": round(r["margin"], 3),
                     "Seasons above zero": r["seasons_positive"], "Luck line (1 in 20 by chance)": round(r["luck_line"], 3),
                     "Odds it is luck": round(r["odds"], 3), "Answer": "YES" if is_yes(r["result"], r["margin"], r["odds"]) else "no",
                     "Corsi alone": None if controlled else round(r["corsi"], 3),
                     "How much the chosen combination moves with Corsi (0 = not at all, 1 = identical)": round(r["overlap"], 2),
                     "Yardstick: expected-goals share on the same test": round(r["xg_beyond"], 3) if key == "beyond" and pd.notna(r["xg_beyond"]) else None,
                     "Echo: what a mere second reading of shot share would show (upper estimate)":
                         round(float(echo(r["overlap"], r["corsi"], reliability)), 3) if key == "beyond" else None,
                     "PUCK-MOVEMENT STATS ONLY: honest result": round(r["mv_result"], 3), "Puck-movement only: noise margin (+/-)": round(r["mv_margin"], 3),
                     "Puck-movement only: odds it is luck": round(r["mv_odds"], 3),
                     "Puck-movement only: answer": "YES" if is_yes(r["mv_result"], r["mv_margin"], r["mv_odds"]) else "no",
                     "All 11 stats averaged": round(r["all_stats"], 3),
                     "Today's 3-stat index": round(r["current"], 3), "Best single stat": f"{NAMES[int(r['singles'].argmax())]} ({r['singles'].max():.2f})",
                     "Best combination using ALL seasons (flattering)": f"{combo_name(COMBOS[r['best']])} ({r['every'][r['best']]:.2f})",
                     "Best puck-movement combination using ALL seasons (flattering)": f"{combo_name(COMBOS[r['mv_best']])} ({r['every'][r['mv_best']]:.2f})"})
    return pd.DataFrame(rows)


def stats_table(res):
    """For each stat: how it does alone, and whether combinations that include it do better or worse than those that leave it out."""
    rows = []
    for j, name in enumerate(NAMES):
        row = {"Statistic": name}
        for key, question, _ in QUESTIONS:
            if key not in res:
                continue
            r, tag = res[key], question.split(".")[0]
            inside = COMBOS[:, j] == 1
            top = np.argsort(-r["every"])[:100]
            row[f"Q{tag}: on its own"] = round(float(r["singles"][j]), 2)
            row[f"Q{tag}: combinations WITH it minus WITHOUT it"] = round(float(r["every"][inside].mean() - r["every"][~inside].mean()), 3)
            row[f"Q{tag}: share of the 100 best combinations that use it"] = round(float(inside[top].mean()), 2)
            row[f"Q{tag}: times in the chosen combination ({r['n']} seasons)"] = int(COMBOS[r["chosen"], j].sum())
        rows.append(row)
    return pd.DataFrame(rows)


def chosen_table(res):
    rows = []
    for key, question, _ in QUESTIONS:
        if key not in res:
            continue
        r = res[key]
        for i, label in enumerate(r["labels"]):
            rows.append({"Question": question, "Season left out": label, "Combination chosen on the other seasons": combo_name(COMBOS[r["chosen"][i]]),
                         "Its score on the season left out": round(float(r["held"][i]), 3),
                         "Puck-movement stats only: combination chosen": combo_name(COMBOS[r["mv_chosen"][i]]),
                         "Puck-movement only: score on the season left out": round(float(r["mv_held"][i]), 3)})
    return pd.DataFrame(rows)


def all_table(res):
    table = pd.DataFrame({"Combination": [combo_name(row) for row in COMBOS], "Stats in it": COMBOS.sum(axis=1).astype(int)})
    for key, question, _ in QUESTIONS:
        if key in res:
            table[question + " (all seasons)"] = res[key]["every"].round(3)
    return table.sort_values(table.columns[2], ascending=False).reset_index(drop=True)


def top_table(res, n=15):
    pieces = []
    full = all_table(res)
    for key, question, _ in QUESTIONS:
        if key in res:
            col = question + " (all seasons)"
            part = full.sort_values(col, ascending=False).head(n).copy()
            part.insert(0, "Ranked by", question)
            part.insert(1, "Rank", np.arange(1, len(part) + 1))
            pieces.append(part)
    return pd.concat(pieces, ignore_index=True)


def verdict_lines(res, n_seasons, season_list, n_shuffles, reliability):
    def odds_text(odds):
        return "less than 1 in 1,000" if odds <= 1 / (n_shuffles + 1) + 1e-12 else f"about 1 in {max(round(1 / odds), 1):,}"

    def answer(key):
        r = res[key]
        yes, mv_yes = is_yes(r["result"], r["margin"], r["odds"]), is_yes(r["mv_result"], r["mv_margin"], r["mv_odds"])
        text = (f"{'YES' if yes else 'NO'}. The best combination, chosen without the season it is scored on, gets {r['result']:+.2f} "
                f"(noise margin +/-{r['margin']:.2f}; above zero in {r['seasons_positive']} of {r['n']} seasons). "
                f"Shuffled results reach {r['luck_line']:+.2f} one time in twenty, and the odds ours is luck are {odds_text(r['odds'])}. ")
        if key not in ("beyond", "beyond2"):
            text += f"Yardstick: Corsi alone gets {r['corsi']:+.2f}. "
        if key == "beyond":
            text += (f"Two yardsticks on the same test: expected-goals share gets {r['xg_beyond']:+.2f}; and a score that was only a second reading of shot share "
                     f"would show up to about {echo(r['overlap'], r['corsi'], reliability):+.2f} (the chosen combination moves with Corsi at {r['overlap']:.2f}). ")
        text += (f"TOUGHER CHECK, puck-movement stats only (no shot counts): {r['mv_result']:+.2f} (margin +/-{r['mv_margin']:.2f}, odds of luck {odds_text(r['mv_odds'])}): "
                 f"{'YES' if mv_yes else 'NO'}. ")
        text += f"For reference: all 11 stats averaged {r['all_stats']:+.2f}; today's 3-stat index {r['current']:+.2f}; best single stat {NAMES[int(r['singles'].argmax())]} ({r['singles'].max():+.2f})."
        return yes, mv_yes, text

    lines = [("What was done", f"All {len(COMBOS):,} combinations of {len(POOL)} breakout stats, on {n_seasons} seasons ({season_list}). The best combination is always chosen on the "
                               f"other seasons and scored on the one left out, and the whole search is repeated on {n_shuffles:,} random shuffles to measure luck. "
                               "Four of the 11 stats are counts of shot attempts (close relatives of Corsi), so every question is also answered with the other seven alone.")]
    yes, mv = {}, {}
    for key, question, _ in QUESTIONS:
        if key in res:
            yes[key], mv[key], text = answer(key)
            lines.append((question, text))

    supported = yes.get("same") and yes.get("other")
    if supported and yes.get("beyond"):
        bottom = "SUPPORTED by the rule set in advance: teams that break out well win more, it shows in games the stats never saw, and something is still there after shot share is taken out. "
    elif supported:
        bottom = "SUPPORTED, BUT NOT SEPARABLE FROM SHOT SHARE: teams that break out well win more and it shows in games the stats never saw, but nothing clear is left once shot share is taken out. "
    elif yes.get("same"):
        bottom = "WEAK: good-breakout teams win more in the same season, but the score does not reliably predict games it has not seen. "
    else:
        bottom = "NOT CONFIRMED with this data: no combination reliably goes with winning once it is tested on seasons it was not chosen on. "
    if supported:
        o, b = res["other"], res["beyond"]
        bottom += (f"How strong: the best recipe predicts unseen games at {o['result']:+.2f}; Corsi alone does {o['corsi']:+.2f}. "
                   f"Three cautions. (1) The recipe leans on stats that are counts of shot attempts; with the seven puck-movement stats alone the answers are: "
                   f"wins {'YES' if mv.get('same') else 'NO'} ({res['same']['mv_result']:+.2f}), predicts unseen games {'YES' if mv.get('other') else 'NO'} ({o['mv_result']:+.2f}), "
                   f"more than shot share {'YES' if mv.get('beyond') else 'NO'} ({b['mv_result']:+.2f}). "
                   f"(2) Part of 'more than shot share' is an echo: any second reading of the same thing adds a little (up to about {echo(b['overlap'], b['corsi'], reliability):+.2f} here, against our {b['result']:+.2f}). ")
        if "beyond2" in res:
            bottom += (f"(3) After also taking out shot quality (expected goals), the result is {res['beyond2']['result']:+.2f} "
                       f"(margin +/-{res['beyond2']['margin']:.2f}): {'still there' if yes.get('beyond2') else 'no longer clear'}. ")
        bottom += "A link is not proof of cause: this shows good teams break out well, not that breaking out well is what makes them good."
    lines.append(("Bottom line", bottom))
    return pd.DataFrame(lines, columns=["", "Answer"])


def write_excel(sheets, notes):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 150
        ws["A1"].font = Font(bold=True)
        for name, frame in sheets.items():
            start = 1 if name in notes else 0
            frame.to_excel(writer, sheet_name=name, index=False, startrow=start)
            ws = writer.sheets[name]
            header_row = start + 1
            if name in notes:
                ws.cell(row=1, column=1, value=notes[name]).font = Font(bold=True)
            ws.freeze_panes = ws.cell(row=header_row + 1, column=2)
            for cell in ws[header_row]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column in ws.columns:
                values = [c.value for c in column[header_row - 1:300] if c.value is not None and (not isinstance(c.value, str) or len(c.value) < 70)]
                longest = max((len(str(v)) for v in values), default=8)
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 46)
        ws = writer.sheets["Verdict"]
        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["B"].width = 135
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(bold=True)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")


COMBOS = all_combinations()


def main():
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Test every combination of breakout stats, honestly.")
    parser.add_argument("--seasons", help="comma-separated season ids to use, e.g. 20232024,20242025")
    parser.add_argument("--shuffles", type=int, default=N_SHUFFLES, help=f"how many random shuffles for the luck check (default {N_SHUFFLES})")
    args = parser.parse_args()
    if not bi.PBP_DB.exists() or not bi.NHL_DB.exists():
        print("I need data/pbp.db (python scripts/pull_play_by_play.py) and data/nhl.db (python scripts/build_database.py).")
        return 1
    pbp = sqlite3.connect(f"file:{bi.PBP_DB}?mode=ro", uri=True)        # read-only: safe while the download is running
    nhl = sqlite3.connect(bi.NHL_DB)
    loaded = bi.load_game_rows(pbp, nhl, args.seasons)
    if loaded is None:
        return 1
    game_rows, usable, odd_seasons = loaded
    seasons = [s for s in usable if s not in odd_seasons]
    if len(seasons) < 4:
        print(f"Only {len(seasons)} normal seasons are ready; the hold-out needs at least 4. Try again when more have downloaded.")
        return 1
    game_rows = game_rows[game_rows["season_id"].isin(seasons)]

    team = add_pool_z(bi.summarize(game_rows)).merge(bi.comparison_numbers(nhl, seasons), on=["season_id", "team_id"], how="left")
    halves = half_tables(game_rows, nhl, seasons)
    following = next_season_table(team)
    label = bi.season_label
    folds = {
        "same": ([[make_part(team[team["season_id"] == s], COMBOS)] for s in seasons], [label(s) for s in seasons]),
        "other": ([[make_part(h[h["season_id"] == s], COMBOS) for h in halves] for s in seasons], [label(s) for s in seasons]),
        "beyond": ([[make_part(h[h["season_id"] == s], COMBOS, ["corsi"]) for h in halves] for s in seasons], [label(s) for s in seasons]),
        "beyond2": ([[make_part(h[h["season_id"] == s], COMBOS, ["corsi", "xg_share"]) for h in halves] for s in seasons], [label(s) for s in seasons]),
    }
    bench = bi.benchmark_halves(nhl, seasons)
    reliability = float(ps.season_spearman(bench, "corsi_a", "corsi_b").mean())      # how well half a season of Corsi repeats (for the echo estimate)
    pairs = sorted(int(s) for s in following["season_id"].unique())
    if len(pairs) >= 4:
        folds["next"] = ([[make_part(following[following["season_id"] == s], COMBOS)] for s in pairs], [f"{label(s)} -> {label(s + 10001)}" for s in pairs])

    rng = np.random.default_rng(SEED)
    res = {}
    print(f"\nTesting {len(COMBOS):,} combinations of {len(POOL)} stats on {len(seasons)} seasons, with {args.shuffles:,} shuffles per question...")
    for key, question, _ in QUESTIONS:
        if key in folds:
            res[key] = run_question(folds[key][0], folds[key][1], COMBOS, rng, args.shuffles)
            print(f"  done: {question}")

    # a check against the Breakout Index workbook: the same three-stat index and pinned-in spells should give the same numbers there
    pinned = NAMES.index("Pinned-in spells")
    print(f"\nCheck against build_breakout_index.py: today's 3-stat index, same season {res['same']['current']:+.2f}; "
          f"pinned-in spells alone, beyond Corsi in the other half {res['beyond']['singles'][pinned]:+.3f}")

    verdict = verdict_lines(res, len(seasons), ", ".join(label(s) for s in seasons), args.shuffles, reliability)
    questions, stats, chosen, top, everything = questions_table(res, reliability), stats_table(res), chosen_table(res), top_table(res), all_table(res)
    sheets = {"Verdict": verdict, "Questions": questions, "Which stats matter": stats, "Chosen by season": chosen,
              "Top combinations": top, "All combinations": everything}
    notes = {
        "Questions": "One row per question. 'Honest result' is the number to trust: the best combination was chosen without the season it was scored on. Answer = YES only if it beats its noise margin and the odds of luck are under 5%.",
        "Which stats matter": "For each stat and each question (Q1 to Q4): its link on its own; whether combinations that include it score higher or lower than those that leave it out; how often it is in the best combinations.",
        "Chosen by season": "Which combination won on the OTHER seasons, and how it then did on the season left out. If the chosen combination keeps changing, no single recipe is clearly best.",
        "Top combinations": "The 15 best combinations for each question using ALL seasons at once. This flatters them: with 2,047 tries the top of the list is partly luck. The 'Questions' sheet is the fair test.",
        "All combinations": "Every combination with its score on each question, all seasons at once (sorted by question 1). Flattering for the same reason: use it to browse, not to pick a winner.",
    }
    try:
        write_excel(sheets, notes)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 250)
    pd.set_option("display.max_colwidth", 110)
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nQuestions:"); print(questions.set_index("Question").T.to_string())
    print("\nWhich stats matter:"); print(stats.set_index("Statistic").T.to_string())
    print("\nChosen by season:"); print(chosen.to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL; Corsi yardstick and ice time: MoneyPuck.com")
    return 0


if __name__ == "__main__":
    sys.exit(main())

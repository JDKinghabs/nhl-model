"""
build_breakout_round2.py  -  the search for the right breakout metric (round 2), the freeze, and the blind test
==============================================================================================================
WHAT IT DOES (in Excel terms):
    Round 1 (build_breakout_search.py) found that the best mix of our breakout stats goes with winning almost as
    well as Corsi, but does not beat it. This round widens the search in four directions that each have a hockey
    reason, then FREEZES the winner so it can be tried once on five seasons that no test has ever touched.

    THE NEW DIRECTIONS
      * CLOSE GAMES ONLY     a team sitting on a lead lets itself be pinned in, which muddies the stats. Every stat is
                             also built from close-score moments only (within one goal in periods 1-2, tied after).
      * THE BREAKOUT BATTLE  ours minus the opponent's in the same games: the pinned-in spells we suffer against the
                             ones we inflict, our zone-clear rate against theirs. (This also measures our forecheck.)
      * DANGER AND WINDOWS   pinned-in spells that include a shot from the slot; pinned-in spells with a 15 and a 45
                             second gap; defence-to-offence within 10 and 30 seconds.
      * DATA-SET WEIGHTS     instead of a plain average, let the data weight the stats (a "ridge regression": like
                             LINEST, but it shrinks the weights so a few seasons cannot push them to extremes).

          * BER                  your breakout efficiency rating, 100 x (controlled exits + 0.5 x safe clears) / breakout attempts,
                             read from events as well as they allow (the definitions are at the top of build_breakout_index.py).

    That gives 68 building blocks (17 stats x 4 forms: ours / ours in close games / battle / battle in close games)
    and every combination of up to four of them: about 867,000 breakout scores.

    TWO WINNERS ARE FROZEN, BY A RULE FIXED IN ADVANCE
      "Our own breakouts"   the best score built only from our own stats (no battle blocks): the pure test of the belief.
      "Breakout battle"     the best score when battle blocks are allowed.
      For each, the winner is the combination that best PREDICTS UNSEEN GAMES (the stats from odd-numbered games against
      the results of even-numbered games, and the reverse). The data-weighted version replaces it only if it does
      better by more than 0.02 on seasons it was not fitted on. The two recipes are written to
      reference/breakout_metric_frozen.json, which goes into git BEFORE the blind test is run.

    THE FINISH LINE (agreed before this round)
      A real team-strength metric:  predicts unseen games at 0.46 or better (Corsi's level) on seasons it has never seen.
      A separate edge:              it also still adds something after Corsi and expected goals are taken out.

    THE BLIND TEST
      Seasons 2009-10 to 2013-14 are kept out of every script. When they have all downloaded:
          python scripts/build_breakout_round2.py --blind-test
      scores the two frozen recipes on them, once. Whatever comes out is the answer.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_breakout_round2.py                 the search and the freeze (about 5 minutes)
    python scripts/build_breakout_round2.py --blind-test    the blind test (only when all five seasons are downloaded)
    python scripts/build_breakout_round2.py --rehearse      runs the blind-test steps on the SEARCH seasons, to check the
                                                            code works; it proves nothing and saves nothing to the freeze

WHAT YOU SHOULD SEE:
    outputs/breakout_round2.xlsx   (and after the blind test, outputs/breakout_blind_test.xlsx)

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. 5-on-5 ice time, Corsi and expected goals for
    the yardsticks come from MoneyPuck.com (non-commercial use, credit required).
"""

import argparse
import datetime
import itertools
import json
import sqlite3
import sys

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_breakout_index as bi         # builds the per-game counts from the play-by-play
import build_breakout_search as bs        # round 1: shared helpers

OUTPUT_FILE = bi.PROJECT_FOLDER / "outputs" / "breakout_round2.xlsx"
BLIND_FILE = bi.PROJECT_FOLDER / "outputs" / "breakout_blind_test.xlsx"
FROZEN_FILE = bi.PROJECT_FOLDER / "reference" / "breakout_metric_frozen.json"

STATS = [  # (key, short name, higher is better?)
    ("exit_time", "Zone exit time", False),
    ("exit_success", "Zone-clear rate", True),
    ("quick_exit", "Quick exits", True),
    ("costly_pg", "Costly giveaways", False),
    ("costly_share", "Costly share of giveaways", False),
    ("trans_pg", "Transition shots", True),
    ("icing_pg", "Icings", False),
    ("pinned_pg", "Pinned-in spells", False),
    ("pinned15_pg", "Pinned-in spells (15 s gap)", False),
    ("pinned45_pg", "Pinned-in spells (45 s gap)", False),
    ("pinned_danger_pg", "Dangerous pinned-in spells", False),
    ("turnover_pg", "Own-zone turnovers", False),
    ("d2o_rate", "Defence-to-offence", True),
    ("d2o10_rate", "Defence-to-offence (10 s)", True),
    ("d2o30_rate", "Defence-to-offence (30 s)", True),
    ("trans_slot_share", "Transition shot quality", True),
    ("ber", "BER (breakout efficiency)", True),
]
FORMS = [  # (key, close-score only?, battle?, label)
    ("own", False, False, "ours"),
    ("own_c", True, False, "ours, close games"),
    ("battle", False, True, "battle"),
    ("battle_c", True, True, "battle, close games"),
]
BLOCKS = [f"{s}|{f}" for s, _, _ in STATS for f, _, _, _ in FORMS]
BLOCK_NAMES = [f"{name} [{label}]" for _, name, _ in STATS for _, _, _, label in FORMS]
IS_BATTLE = np.array([battle for _ in STATS for _, _, battle, _ in FORMS])
IS_CLOSE = np.array([close for _ in STATS for _, close, _, _ in FORMS])
TIERS = [("own", "Our own breakouts"), ("battle", "Breakout battle")]

MAX_SIZE = 4                 # combinations of up to this many blocks
CHUNK = 60000                # combinations worked on at a time (keeps memory modest)
ALPHAS = [3, 10, 30, 100, 300, 1000, 3000]   # how hard the data-set weights are shrunk; the best is chosen on other seasons
RIDGE_EDGE = 0.02            # the data-weighted version must beat the plain average by more than this to replace it
FINISH_LINE = 0.46           # "predicts unseen games" at Corsi's level
ODDS_BAR = 0.05
BLIND_SHUFFLES = 2000
SEED = 20261115
TARGETS = ["pts_pct", "gd_pg"]
QUESTIONS = [("same", "Goes with winning (same season)"), ("other", "PREDICTS UNSEEN GAMES (the finish line)"),
             ("beyond", "More than shot share (beyond Corsi)"), ("beyond2", "More than shot share and shot quality (beyond Corsi and expected goals)")]

READ_ME = [
    "The search for the right breakout metric, round 2",
    "",
    "WHAT WAS DONE",
    "  17 breakout stats, each in four forms (ours / ours in close games / battle = ours minus the opponent's / battle in close games) = 68 building blocks.",
    "  Every combination of up to four blocks (about 867,000) was turned into a score (the plain average of its blocks) and judged on how well it PREDICTS UNSEEN GAMES:",
    "  the stats from odd-numbered games against the results of even-numbered games, and the reverse. A data-weighted version (ridge regression) was tried as well.",
    "",
    "HOW TO READ THE NUMBERS",
    "  Every number is a rank correlation (CORREL on the teams' rankings, inside each season, averaged): 0 = no relationship, 1 = perfect.",
    "  'Honest' numbers come from choosing the recipe WITHOUT one season and scoring it on that season. They are still from the seasons used to search:",
    "  with this many tries the only clean judge is the BLIND TEST on the five seasons no test has touched.",
    "  Yardsticks: Corsi predicts unseen games at about 0.46, expected-goals share at about 0.51, and a team's own goal difference at about 0.51.",
    "",
    "THE TWO FROZEN RECIPES",
    "  'Our own breakouts'  only our own stats (no battle blocks): the pure test of 'breakouts make teams good'.",
    "  'Breakout battle'    battle blocks allowed: breakout and forecheck together.",
    "  They are saved in reference/breakout_metric_frozen.json before the blind test is run.",
    "",
    "THE FINISH LINE",
    f"  A real team-strength metric: predicts unseen games at {FINISH_LINE:.2f} or better on the blind-test seasons.",
    "  A separate edge: on those seasons it also adds something after Corsi and expected goals are taken out (more than the noise margin, odds of luck under 5%).",
    "",
    "Data: (c) NHL. Corsi, expected goals and 5-on-5 ice time: MoneyPuck.com.",
]


# ----------------------------------------------------------------------------
# The 64 building blocks
# ----------------------------------------------------------------------------
def with_opponent(rows):
    """Puts the opponent's counts from the same game next to each team's own (columns ending _opp)."""
    counts = [c for c in rows.columns if c not in ("game_id", "team_id", "season_id", "toi5", "toi5_c", "ev5", "ev5_c")]
    opp = rows[["game_id", "team_id"] + counts].rename(columns={"team_id": "opp_id", **{c: f"{c}_opp" for c in counts}})
    m = rows.merge(opp, on="game_id")
    return m[m["team_id"] != m["opp_id"]].drop(columns="opp_id")


def raw_stats(g, close, side):
    """The 16 stats from summed counts. close = '_c' for close-score moments; side = '_opp' for the opponent's numbers."""
    def c(name):
        return g[f"{name}{close}{side}"]
    hours = g[f"toi5{close}"] / 3600                                     # the same ice time for both sides
    return {
        "exit_time": c("exit_sum") / c("exit_starts"), "exit_success": c("exit_ok") / c("exit_starts"), "quick_exit": c("exit_fast") / c("exit_starts"),
        "costly_pg": c("costly") / hours, "costly_share": c("costly") / c("giveaways"), "trans_pg": c("trans_for") / hours, "icing_pg": c("icings") / hours,
        "pinned_pg": c("pinned") / hours, "pinned15_pg": c("pinned15") / hours, "pinned45_pg": c("pinned45") / hours, "pinned_danger_pg": c("pinned_danger") / hours,
        "turnover_pg": (c("giveaways") + c("stripped")) / hours,
        "d2o_rate": c("d2o_ok") / c("d2o_starts"), "d2o10_rate": c("d2o_ok10") / c("d2o_starts"), "d2o30_rate": c("d2o_ok30") / c("d2o_starts"),
        "trans_slot_share": c("trans_slot") / c("trans_unblocked"),
        # BER = 100 x (controlled exits + 0.5 x safe clears) / breakout attempts that could be judged
        "ber": 100 * (c("ber_ctrl") + 0.5 * c("ber_safe")) / (c("ber_ctrl") + c("ber_safe") + c("ber_fail")),
    }


def block_table(rows):
    """One row per team-season: the 64 blocks, each as steps above or below that season's average, turned so higher = better."""
    g = rows.groupby(["season_id", "team_id"]).sum(numeric_only=True)
    out = {}
    for close in ("", "_c"):
        own, opp = raw_stats(g, close, ""), raw_stats(g, close, "_opp")
        for key, _, higher in STATS:
            sign = 1 if higher else -1
            out[f"{key}|own{close}"] = sign * own[key]
            out[f"{key}|battle{close}"] = sign * (own[key] - opp[key])
    t = pd.DataFrame(out)[BLOCKS].replace([np.inf, -np.inf], np.nan)
    season = t.index.get_level_values("season_id")
    z = (t - t.groupby(season).transform("mean")) / t.groupby(season).transform("std")
    return z.fillna(0.0).reset_index()


def half_frames(rows, nhl, seasons):
    """Two tables: the blocks from odd-numbered games next to the RESULTS of even-numbered games, and the reverse."""
    ids = ",".join(str(int(s)) for s in seasons)
    results = pd.read_sql_query(
        f"SELECT season_id, game_id, team_id, points, goalsFor - goalsAgainst AS gd FROM team_games WHERE game_type = 2 AND season_id IN ({ids})", nhl)
    bench = bi.benchmark_halves(nhl, seasons)                            # corsi_a / xg_share_a = odd games, _b = even games
    frames = []
    for parity, suffix in ((1, "_a"), (0, "_b")):
        z = block_table(rows[rows["game_id"] % 2 == parity])
        r = results[results["game_id"] % 2 != parity].groupby(["season_id", "team_id"]).agg(
            pts=("points", "sum"), n=("points", "size"), gd=("gd", "sum")).reset_index()
        r["pts_pct"], r["gd_pg"] = r["pts"] / (2 * r["n"]), r["gd"] / r["n"]
        c = bench[["season_id", "team_id", f"corsi{suffix}", f"xg_share{suffix}"]].rename(columns={f"corsi{suffix}": "corsi", f"xg_share{suffix}": "xg_share"})
        frames.append(z.merge(r[["season_id", "team_id", "pts_pct", "gd_pg"]], on=["season_id", "team_id"]).merge(c, on=["season_id", "team_id"]))
    return frames


# ----------------------------------------------------------------------------
# The engine
# ----------------------------------------------------------------------------
def rank_unit(scores):
    """Each column turned into ranks, centred and scaled to length 1, so that column . column is a rank correlation."""
    n = scores.shape[0]
    r = scores.argsort(axis=0).argsort(axis=0).astype(np.float32)
    return (r - (n - 1) / 2) / np.float32(np.sqrt(n * (n * n - 1) / 12))


def prepare(frame):
    """Everything about one season (or one half of one) that does not depend on the recipe being tested."""
    d = frame.dropna(subset=TARGETS + ["corsi", "xg_share"])
    y = bs.unit(d[TARGETS].rank().to_numpy())
    k1 = bs.unit(d[["corsi"]].rank().to_numpy())
    k2 = bs.unit(d[["corsi", "xg_share"]].rank().to_numpy())
    xg = bs.unit(d[["xg_share"]].rank().to_numpy())
    gd = d["gd_pg"].to_numpy()
    return {"Z": d[BLOCKS].to_numpy(dtype=np.float32), "Y": y, "K1": k1, "P1": np.linalg.pinv(k1), "Y1": bs.take_out(y, k1),
            "K2": k2, "P2": np.linalg.pinv(k2), "Y2": bs.take_out(y, k2), "gd_z": (gd - gd.mean()) / gd.std(),
            "corsi": float((k1.T @ y).mean()), "xg": float((xg.T @ y).mean()), "xg_beyond": float((bs.take_out(xg, k1).T @ bs.take_out(y, k1)).mean())}


def links(part, x, shuffle=None):
    """
    The link of each column of x (already rank_unit) with the results: plain, beyond Corsi, and beyond Corsi and expected goals.
    `shuffle` is an optional reordering of the teams' results (for the luck check).
    """
    def take(a):
        return a if shuffle is None else a[shuffle]
    out = {"other": (x.T @ take(part["Y"])).mean(axis=1)}
    for key, k, p, y in (("beyond", "K1", "P1", "Y1"), ("beyond2", "K2", "P2", "Y2")):
        left = x - part[k] @ (part[p] @ x)
        out[key] = (left.T @ take(part[y])).mean(axis=1) / np.sqrt((left ** 2).sum(axis=0))
    return out


def all_combinations(max_size):
    """A list of index tables, one per size: every way of picking 1, 2, ... blocks."""
    return [np.array(list(itertools.combinations(range(len(BLOCKS)), k)), dtype=np.int16) for k in range(1, max_size + 1)]


def search(full_parts, half_parts, combos):
    """The link of every combination, for every season: m[question] has one row per season and one column per combination."""
    total = sum(len(c) for c in combos)
    m = {q: np.zeros((len(full_parts), total), dtype=np.float32) for q, _ in QUESTIONS}
    for f in range(len(full_parts)):
        start = 0
        for table in combos:
            for a in range(0, len(table), CHUNK):
                idx = table[a:a + CHUNK]
                cols = slice(start + a, start + a + len(idx))
                m["same"][f, cols] = (rank_unit(full_parts[f]["Z"][:, idx].sum(axis=2)).T @ full_parts[f]["Y"]).mean(axis=1)
                for part in half_parts[f]:
                    got = links(part, rank_unit(part["Z"][:, idx].sum(axis=2)))
                    for q in ("other", "beyond", "beyond2"):
                        m[q][f, cols] += got[q] / len(half_parts[f])
            start += len(table)
        print(f"  season {f + 1} of {len(full_parts)} done")
    return m


def combo_blocks(combos, i):
    """The blocks (as positions in BLOCKS) of combination number i."""
    for table in combos:
        if i < len(table):
            return [int(b) for b in table[i]]
        i -= len(table)
    raise IndexError(i)


def recipe_name(weights):
    return " + ".join(BLOCK_NAMES[BLOCKS.index(b)] for b in weights)


def honest_pick(m, allowed):
    """For each season: the best allowed combination on the OTHER seasons (by 'predicts unseen games'), as a column number."""
    target = m["other"][:, allowed]
    total = target.sum(axis=0)
    return np.array([int(allowed[(total - target[i]).argmax()]) for i in range(len(target))])


# ----------------------------------------------------------------------------
# Data-set weights (ridge regression)
# ----------------------------------------------------------------------------
def ridge_fit(parts, cols, alpha):
    """Weights that best predict the other half's goal difference from the blocks, shrunk towards zero by `alpha`."""
    x = np.vstack([p["Z"][:, cols] for p in parts]).astype(float)
    y = np.concatenate([p["gd_z"] for p in parts])
    return np.linalg.solve(x.T @ x + alpha * np.eye(len(cols)), x.T @ y)


def score_metric(part, weights_vector):
    """One recipe (a weight per block) scored on one part: plain, beyond Corsi, beyond Corsi and expected goals."""
    got = links(part, rank_unit((part["Z"].astype(float) @ weights_vector)[:, None]))
    return {q: float(v[0]) for q, v in got.items()}


def ridge_honest(full_parts, half_parts, cols):
    """Leave one season out: fit on the others (shrinkage chosen on them too), score on the one left out."""
    n = len(half_parts)
    rows, chosen_alpha = [], []
    for i in range(n):
        train = [j for j in range(n) if j != i]
        best_alpha, best_score = ALPHAS[0], -np.inf
        for alpha in ALPHAS:                                             # choose the shrinkage without season i: leave each training season out in turn
            inner = []
            for j in train:
                beta = ridge_fit([p for k in train if k != j for p in half_parts[k]], cols, alpha)
                inner.append(np.mean([score_metric(p, full_vector(beta, cols))["other"] for p in half_parts[j]]))
            if np.mean(inner) > best_score:
                best_alpha, best_score = alpha, float(np.mean(inner))
        w = full_vector(ridge_fit([p for k in train for p in half_parts[k]], cols, best_alpha), cols)
        got = {q: float(np.mean([score_metric(p, w)[q] for p in half_parts[i]])) for q in ("other", "beyond", "beyond2")}
        got["same"] = float((rank_unit((full_parts[i]["Z"].astype(float) @ w)[:, None]).T @ full_parts[i]["Y"]).mean())
        rows.append(got)
        chosen_alpha.append(best_alpha)
    final_alpha = max(set(chosen_alpha), key=chosen_alpha.count)
    final = full_vector(ridge_fit([p for k in range(n) for p in half_parts[k]], cols, final_alpha), cols)
    return pd.DataFrame(rows), final, final_alpha


def full_vector(beta, cols):
    w = np.zeros(len(BLOCKS))
    w[cols] = beta
    return w


# ----------------------------------------------------------------------------
# Scoring a frozen recipe on a set of seasons (used by the blind test and the rehearsal)
# ----------------------------------------------------------------------------
def judge(weights, full_parts, half_parts, next_parts, rng, shuffles):
    """One frozen recipe on a set of seasons: each question's average, noise margin, odds of luck and the yardsticks."""
    w = np.array([weights.get(b, 0.0) for b in BLOCKS])

    def one(part, key, shuffle=None):
        x = rank_unit((part["Z"].astype(float) @ w)[:, None])
        if key in ("same", "next"):
            y = part["Y"] if shuffle is None else part["Y"][shuffle]
            return float((x.T @ y).mean())
        return float(links(part, x, shuffle)[key][0])

    groups = {"same": [[p] for p in full_parts], "other": half_parts, "beyond": half_parts, "beyond2": half_parts, "next": [[p] for p in next_parts]}
    out = {}
    for key, folds in groups.items():
        if not folds:
            continue
        per_season = np.array([np.mean([one(p, key) for p in parts]) for parts in folds])
        luck = np.array([np.mean([np.mean([one(p, key, rng.permutation(len(p["Y"]))) for p in parts]) for parts in folds]) for _ in range(shuffles)])
        n = len(per_season)
        out[key] = {"result": float(per_season.mean()), "margin": float(2 * per_season.std(ddof=1) / np.sqrt(n)) if n > 1 else float("nan"),
                    "odds": float(((luck >= per_season.mean()).sum() + 1) / (shuffles + 1)), "seasons": n, "seasons_above_zero": int((per_season > 0).sum()),
                    "corsi": float(np.mean([np.mean([p["corsi"] for p in parts]) for parts in folds])),
                    "xg": float(np.mean([np.mean([p["xg"] for p in parts]) for parts in folds])),
                    "xg_beyond_corsi": float(np.mean([np.mean([p["xg_beyond"] for p in parts]) for parts in folds])), "per_season": per_season.round(3).tolist()}
    return out


JUDGE_LABELS = [("same", "Goes with winning (same season)"), ("other", "PREDICTS UNSEEN GAMES (the finish line)"), ("next", "Carries into next season"),
                ("beyond", "More than shot share (beyond Corsi)"), ("beyond2", "More than shot share and shot quality")]


def judge_table(results):
    rows = []
    for tier, label in TIERS:
        for key, question in JUDGE_LABELS:
            if tier in results and key in results[tier]:
                r = results[tier][key]
                controlled = key in ("beyond", "beyond2")
                rows.append({"Recipe": label, "Question": question, "Seasons": r["seasons"], "Result": round(r["result"], 3), "Noise margin (+/-)": round(r["margin"], 3),
                             "Seasons above zero": r["seasons_above_zero"], "Odds it is luck": round(r["odds"], 4),
                             "Corsi on the same seasons": None if controlled else round(r["corsi"], 3),
                             "Expected-goals share on the same seasons": round(r["xg_beyond_corsi"], 3) if key == "beyond" else (None if controlled else round(r["xg"], 3)),
                             "Season by season": ", ".join(f"{v:+.2f}" for v in r["per_season"])})
    return pd.DataFrame(rows)


def finish_line_text(r):
    """The two finish-line answers for one recipe, from judge() results."""
    o, b = r["other"], r["beyond2"]
    strength = o["result"] >= FINISH_LINE
    edge = b["result"] > b["margin"] and b["odds"] < ODDS_BAR
    return strength, edge, (
        f"Predicts unseen games at {o['result']:+.2f} (noise margin +/-{o['margin']:.2f}; Corsi on the same seasons {o['corsi']:+.2f}, expected goals {o['xg']:+.2f}): "
        f"{'REACHES' if strength else 'MISSES'} the {FINISH_LINE:.2f} line"
        + (" and beats Corsi on these seasons. " if o["result"] > o["corsi"] else " and does not beat Corsi on these seasons. ")
        + f"After Corsi and expected goals are taken out: {b['result']:+.2f} (margin +/-{b['margin']:.2f}, odds of luck {b['odds']:.3f}): "
        f"{'A SEPARATE EDGE' if edge else 'no separate edge'}.")


# ----------------------------------------------------------------------------
def write_excel(path, sheets, notes, read_me):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": read_me}).to_excel(writer, sheet_name="Read me", index=False)
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


def build_parts(rows, nhl, seasons, with_next=False):
    """The prepared season tables: full season, the two halves, and (optionally) this season against next season's results."""
    rows = with_opponent(rows)
    compare = bi.comparison_numbers(nhl, seasons)[["season_id", "team_id"] + TARGETS + ["corsi", "xg_share"]]
    table = block_table(rows).merge(compare, on=["season_id", "team_id"], how="left")
    halves = half_frames(rows, nhl, seasons)
    full_parts = [prepare(table[table["season_id"] == s]) for s in seasons]
    half_parts = [[prepare(h[h["season_id"] == s]) for h in halves] for s in seasons]
    next_parts = []
    if with_next:                                                        # next season's results can come from any season in the database
        after = [s + 10001 for s in seasons]
        later = bi.comparison_numbers(nhl, sorted(set(seasons) | set(after)))[["season_id", "team_id"] + TARGETS]
        later = later.assign(season_id=later["season_id"] - 10001)
        nxt = table[["season_id", "team_id", "corsi", "xg_share"] + BLOCKS].merge(later, on=["season_id", "team_id"])
        next_parts = [prepare(nxt[nxt["season_id"] == s]) for s in seasons if (nxt["season_id"] == s).sum() >= 10]
    return full_parts, half_parts, next_parts


def block_repeats(rows):
    """How well each block repeats: its ranking of the teams from odd-numbered games against the ranking from even-numbered games."""
    rows = with_opponent(rows)
    m = block_table(rows[rows["game_id"] % 2 == 1]).merge(block_table(rows[rows["game_id"] % 2 == 0]), on=["season_id", "team_id"], suffixes=("_a", "_b"))
    return [float(m.groupby("season_id").apply(lambda d, b=b: d[f"{b}_a"].rank().corr(d[f"{b}_b"].rank())).mean()) for b in BLOCKS]


def ber_breakdown(rows):
    """What became of the breakout attempts: the shares that were controlled exits, safe clears, failed, or could not be judged."""
    s = rows[["ber_attempts", "ber_ctrl", "ber_safe", "ber_fail"]].sum()
    judged = s["ber_ctrl"] + s["ber_safe"] + s["ber_fail"]
    return {"attempts_per_team_game": float(s["ber_attempts"] / len(rows)), "judged": float(judged / s["ber_attempts"]),
            "controlled": float(s["ber_ctrl"] / judged), "safe": float(s["ber_safe"] / judged), "failed": float(s["ber_fail"] / judged),
            "league_ber": float(100 * (s["ber_ctrl"] + 0.5 * s["ber_safe"]) / judged)}


def settings_now():
    return {"EXIT_CAP": bi.EXIT_CAP, "FAST_EXIT": bi.FAST_EXIT, "GIVEAWAY_WINDOW": bi.GIVEAWAY_WINDOW, "TRANSITION_WINDOW": bi.TRANSITION_WINDOW,
            "PINNED_GAP": bi.PINNED_GAP, "D2O_WINDOW": bi.D2O_WINDOW, "SLOT_HALF_WIDTH": bi.SLOT_HALF_WIDTH, "SLOT_DEPTH": bi.SLOT_DEPTH,
            "BER_WINDOW": bi.BER_WINDOW, "SAFE_CLEAR_WINDOW": bi.SAFE_CLEAR_WINDOW, "FEATURES_VERSION": bi.FEATURES_VERSION}


# ----------------------------------------------------------------------------
# The blind test (and its rehearsal)
# ----------------------------------------------------------------------------
def blind_test(pbp, nhl, rehearse):
    if not FROZEN_FILE.exists():
        print("There is no frozen recipe yet. Run the search first:  python scripts/build_breakout_round2.py")
        return 1
    frozen = json.loads(FROZEN_FILE.read_text(encoding="utf-8"))
    if frozen["settings"] != settings_now():
        print("The stat settings at the top of build_breakout_index.py have changed since the recipes were frozen. Put them back, or run the search again to re-freeze.")
        return 1
    if not rehearse:                                                     # all five or nothing: scoring some of them early would be a peek
        coverage = bi.season_coverage(pbp, nhl).set_index("season_id")["coverage"]
        missing = [bi.season_label(s) for s in bi.LOCKBOX if coverage.get(s, 0) < bi.MIN_COVERAGE]
        if missing:
            print(f"The blind test needs ALL five kept-back seasons, and these have not finished downloading: {', '.join(missing)}.")
            print("Nothing was read or scored. Run this again when  python scripts/pull_play_by_play.py --status  shows them complete.")
            return 1
    loaded = bi.load_game_rows(pbp, nhl, lockbox=not rehearse)
    if loaded is None:
        return 1
    rows, seasons, odd_seasons = loaded
    seasons = [s for s in seasons if s not in odd_seasons]
    rows = rows[rows["season_id"].isin(seasons)]
    full_parts, half_parts, next_parts = build_parts(rows, nhl, seasons, with_next=True)
    rng = np.random.default_rng(SEED)
    results = {tier: judge(frozen["metrics"][tier]["weights"], full_parts, half_parts, next_parts, rng, BLIND_SHUFFLES) for tier, _ in TIERS}

    title = "REHEARSAL on the search seasons (proves nothing: these seasons chose the recipes)" if rehearse else "BLIND TEST on the five seasons no test had touched"
    lines = [("What this is", f"{title}: {', '.join(bi.season_label(s) for s in seasons)}. The recipes were frozen on {frozen['frozen_on']}.")]
    verdicts = {}
    for tier, label in TIERS:
        strength, edge, text = finish_line_text(results[tier])
        verdicts[tier] = {"reaches_finish_line": bool(strength), "separate_edge": bool(edge)}
        lines.append((f"{label}: {recipe_name(frozen['metrics'][tier]['weights'])}", text))
    verdict = pd.DataFrame(lines, columns=["", "Answer"])
    table = judge_table(results)
    pd.set_option("display.width", 250)
    pd.set_option("display.max_colwidth", 120)
    print(f"\n{title}")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print(); print(table.drop(columns="Season by season").to_string(index=False))
    if rehearse:
        print("\nRehearsal only: nothing was saved.")
        return 0
    try:
        write_excel(BLIND_FILE, {"Verdict": verdict, "Blind test": table},
                    {"Blind test": "The two frozen recipes on the five kept-back seasons. 'Odds it is luck' = the share of 2,000 random shuffles of the results that did as well."},
                    ["The blind test of the two frozen breakout recipes", "", "The recipes were chosen on the ten search seasons and frozen in reference/breakout_metric_frozen.json.",
                     "These five seasons were kept out of every script until this test, so nothing about them could influence the recipes.",
                     f"Finish line: predicts unseen games at {FINISH_LINE:.2f} or better; a separate edge needs a clear result after Corsi and expected goals are taken out.",
                     "", "Data: (c) NHL. Corsi, expected goals and 5-on-5 ice time: MoneyPuck.com."])
    except PermissionError:
        print(f"Could not save {BLIND_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1
    if frozen.get("blind_test") is None:
        frozen["blind_test"] = {"run_on": datetime.date.today().isoformat(), "seasons": seasons, "verdicts": verdicts,
                                "results": {tier: {k: {a: b for a, b in v.items()} for k, v in results[tier].items()} for tier, _ in TIERS}}
        FROZEN_FILE.write_text(json.dumps(frozen, indent=2), encoding="utf-8")
        print(f"\nThe result is recorded in {FROZEN_FILE.name}. The kept-back seasons are now used: they cannot give a clean test again.")
    print(f"Saved: {BLIND_FILE}")
    return 0


# ----------------------------------------------------------------------------
# The search and the freeze
# ----------------------------------------------------------------------------
def main():
    sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Round 2 of the search for a breakout metric, the freeze, and the blind test.")
    parser.add_argument("--blind-test", action="store_true", help="score the frozen recipes on the five kept-back seasons (once they have all downloaded)")
    parser.add_argument("--rehearse", action="store_true", help="run the blind-test steps on the search seasons, to check the code (proves nothing)")
    parser.add_argument("--max-size", type=int, default=MAX_SIZE, help=f"largest combination to try (default {MAX_SIZE})")
    args = parser.parse_args()
    if not bi.PBP_DB.exists() or not bi.NHL_DB.exists():
        print("I need data/pbp.db (python scripts/pull_play_by_play.py) and data/nhl.db (python scripts/build_database.py).")
        return 1
    pbp = sqlite3.connect(f"file:{bi.PBP_DB}?mode=ro", uri=True)        # read-only: safe while the download is running
    nhl = sqlite3.connect(bi.NHL_DB)
    if args.blind_test or args.rehearse:
        return blind_test(pbp, nhl, rehearse=args.rehearse)

    if FROZEN_FILE.exists() and json.loads(FROZEN_FILE.read_text(encoding="utf-8")).get("blind_test"):
        print(f"The blind test has already been run on the recipes in {FROZEN_FILE.name}. A new search cannot be blind-tested on those seasons again,")
        print("so this script will not overwrite the frozen recipes. (Delete that file by hand only if you accept that the next result is not a blind test.)")
        return 1
    loaded = bi.load_game_rows(pbp, nhl)
    if loaded is None:
        return 1
    rows, usable, odd_seasons = loaded
    seasons = [s for s in usable if s not in odd_seasons]
    rows = rows[rows["season_id"].isin(seasons)]
    full_parts, half_parts, _ = build_parts(rows, nhl, seasons)
    combos = all_combinations(args.max_size)
    total = sum(len(c) for c in combos)
    battle_used = np.concatenate([IS_BATTLE[c].any(axis=1) for c in combos])
    sizes = np.concatenate([np.full(len(c), c.shape[1]) for c in combos])
    print(f"\nRound 2: {len(BLOCKS)} blocks, {total:,} combinations of up to {args.max_size}, on {len(seasons)} seasons...")
    m = search(full_parts, half_parts, combos)
    every = {q: m[q].mean(axis=0) for q, _ in QUESTIONS}
    labels = [bi.season_label(s) for s in seasons]
    corsi = float(np.mean([np.mean([p["corsi"] for p in parts]) for parts in half_parts]))
    xg = float(np.mean([np.mean([p["xg"] for p in parts]) for parts in half_parts]))

    tiers, chosen_rows, frozen_metrics = {}, [], {}
    for tier, label in TIERS:
        allowed = np.where(~battle_used)[0] if tier == "own" else np.arange(total)
        cols = np.where(~IS_BATTLE)[0] if tier == "own" else np.arange(len(BLOCKS))
        picks = honest_pick(m, allowed)
        held = {q: m[q][np.arange(len(seasons)), picks] for q, _ in QUESTIONS}
        best = int(allowed[every["other"][allowed].argmax()])
        ridge_rows, ridge_w, ridge_alpha = ridge_honest(full_parts, half_parts, cols)
        plain, weighted = float(held["other"].mean()), float(ridge_rows["other"].mean())
        use_ridge = weighted > plain + RIDGE_EDGE
        if use_ridge:
            kept = np.argsort(-np.abs(ridge_w))
            weights = {BLOCKS[i]: round(float(ridge_w[i]), 4) for i in kept if abs(ridge_w[i]) > 1e-9}
        else:
            blocks = combo_blocks(combos, best)
            weights = {BLOCKS[b]: round(1 / len(blocks), 4) for b in blocks}
        tiers[tier] = {"label": label, "held": held, "picks": picks, "best": best, "ridge": ridge_rows, "ridge_w": ridge_w, "ridge_alpha": ridge_alpha,
                       "use_ridge": use_ridge, "weights": weights, "allowed": allowed, "n": len(allowed)}
        for i, s in enumerate(labels):
            chosen_rows.append({"Recipe": label, "Season left out": s, "Combination chosen on the other seasons": " + ".join(BLOCK_NAMES[b] for b in combo_blocks(combos, picks[i])),
                                **{f"{name}: its score on the season left out": round(float(held[q][i]), 3) for q, name in QUESTIONS},
                                "Data-weighted version, predicts unseen games": round(float(ridge_rows["other"].iloc[i]), 3)})
        frozen_metrics[tier] = {"label": label, "kind": "data-set weights (ridge)" if use_ridge else "plain average", "weights": weights,
                                "on_the_search_seasons": {"predicts_unseen_games_honest": round(plain, 3), "data_weighted_honest": round(weighted, 3),
                                                          "frozen_recipe_all_seasons": round(float(every["other"][best]), 3), "combinations_tried": int(len(allowed))}}

    # ---- the freeze: written BEFORE any blind test ----
    frozen = {"frozen_on": datetime.date.today().isoformat(),
              "note": "Chosen on the search seasons only. To be scored ONCE on the lockbox seasons with: python scripts/build_breakout_round2.py --blind-test",
              "search_seasons": seasons, "lockbox": bi.LOCKBOX, "settings": settings_now(),
              "finish_line": {"predicts_unseen_games_at_least": FINISH_LINE,
                              "separate_edge": "after Corsi and expected goals are taken out: above its noise margin, odds of luck under 5%"},
              "metrics": frozen_metrics, "blind_test": None}
    FROZEN_FILE.write_text(json.dumps(frozen, indent=2), encoding="utf-8")

    # ---- tables ----
    summary = []
    for tier, label in TIERS:
        t = tiers[tier]
        for q, name in QUESTIONS:
            h, r = t["held"][q], t["ridge"][q].to_numpy()
            summary.append({"Recipe": label, "Question": name, "Honest result (plain average of the chosen blocks)": round(float(h.mean()), 3),
                            "Noise margin (+/-)": round(float(2 * h.std(ddof=1) / np.sqrt(len(h))), 3), "Seasons above zero": int((h > 0).sum()),
                            "Honest result (data-set weights)": round(float(r.mean()), 3),
                            "Frozen recipe, all search seasons (flattering)": round(float(every[q][t["best"]]), 3),
                            "Yardstick: Corsi": round(corsi, 3) if q == "other" else (round(float(np.mean([p["corsi"] for p in full_parts])), 3) if q == "same" else None),
                            "Yardstick: expected-goals share": round(xg, 3) if q == "other" else (round(float(np.mean([p["xg"] for p in full_parts])), 3) if q == "same" else
                                                                                                   (round(float(np.mean([np.mean([p["xg_beyond"] for p in parts]) for parts in half_parts])), 3) if q == "beyond" else None))})
    summary = pd.DataFrame(summary)

    single = np.arange(len(BLOCKS))                                       # the first combinations are the single blocks, in order
    top200 = np.argsort(-every["other"])[:200]
    blocks = pd.DataFrame({"Block": BLOCK_NAMES, "Form": [label for _ in STATS for _, _, _, label in FORMS],
                           "Repeats (odd vs even games)": np.round(block_repeats(rows), 2),
                           "On its own: predicts unseen games": every["other"][single].round(3), "On its own: goes with winning": every["same"][single].round(3),
                           "On its own: beyond Corsi": every["beyond"][single].round(3), "On its own: beyond Corsi and expected goals": every["beyond2"][single].round(3)})
    member = np.zeros((len(BLOCKS), total), dtype=bool)
    start = 0
    for table in combos:
        for k in range(table.shape[1]):
            member[table[:, k], start + np.arange(len(table))] = True
        start += len(table)
    blocks["Combinations WITH it minus WITHOUT it (predicts unseen games)"] = [round(float(every["other"][member[b]].mean() - every["other"][~member[b]].mean()), 3) for b in range(len(BLOCKS))]
    blocks["Share of the 200 best combinations that use it"] = [round(float(member[b][top200].mean()), 2) for b in range(len(BLOCKS))]
    for tier, label in TIERS:
        blocks[f"{label}: times in the chosen combination ({len(seasons)} seasons)"] = [int(sum(b in combo_blocks(combos, p) for p in tiers[tier]["picks"])) for b in range(len(BLOCKS))]
        blocks[f"{label}: data-set weight"] = tiers[tier]["ridge_w"].round(3)
    forms = blocks.groupby("Form", sort=False)[["On its own: predicts unseen games", "On its own: goes with winning", "On its own: beyond Corsi and expected goals"]].mean().round(3).reset_index()
    forms.columns = ["Form", "Average block: predicts unseen games", "Average block: goes with winning", "Average block: beyond Corsi and expected goals"]

    tops = []
    for tier, label in TIERS:
        order = tiers[tier]["allowed"][np.argsort(-every["other"][tiers[tier]["allowed"]])[:25]]
        for rank, i in enumerate(order, 1):
            tops.append({"Recipe": label, "Rank": rank, "Combination": " + ".join(BLOCK_NAMES[b] for b in combo_blocks(combos, int(i))), "Blocks": int(sizes[i]),
                         **{f"{name} (all search seasons)": round(float(every[q][i]), 3) for q, name in QUESTIONS}})
    tops = pd.DataFrame(tops)

    lines = [("What was done", f"{len(BLOCKS)} building blocks ({len(STATS)} stats, each as ours / ours in close games / battle / battle in close games), every combination of up to {args.max_size}: "
                               f"{total:,} breakout scores on {len(seasons)} seasons ({', '.join(labels)}). Yardsticks for predicting unseen games: Corsi {corsi:+.2f}, expected-goals share {xg:+.2f}.")]
    for tier, label in TIERS:
        t = tiers[tier]
        h = t["held"]
        margin = 2 * h["other"].std(ddof=1) / np.sqrt(len(seasons))
        lines.append((label, f"{t['n']:,} combinations. Chosen without the season it is scored on, the best one PREDICTS UNSEEN GAMES at {h['other'].mean():+.2f} (noise margin +/-{margin:.2f}); "
                             f"goes with winning in the same season {h['same'].mean():+.2f}; beyond Corsi {h['beyond'].mean():+.2f}; beyond Corsi and expected goals {h['beyond2'].mean():+.2f}. "
                             f"Data-set weights instead of a plain average: {t['ridge']['other'].mean():+.2f}, "
                             + ("better by more than 0.02, so the weighted version is frozen." if t["use_ridge"] else "not better by more than 0.02, so the plain average is kept.")
                             + f" FROZEN RECIPE: {recipe_name(t['weights'])}."))
    ber, ber_rows = ber_breakdown(rows), blocks[blocks["Block"].str.startswith("BER")]
    in_frozen = [label for tier, label in TIERS if any(b.startswith("ber|") for b in tiers[tier]["weights"]) and not tiers[tier]["use_ridge"]]
    lines.append(("BER (your formula, read from events)",
                  f"{ber['attempts_per_team_game']:.0f} breakout attempts per team per game; {ber['judged']:.0%} of them could be judged. Of those: controlled exit {ber['controlled']:.0%}, "
                  f"safe clear {ber['safe']:.0%}, failed {ber['failed']:.0%}; league-average BER {ber['league_ber']:.0f}. "
                  + " | ".join(f"{r['Form']}: repeats {r['Repeats (odd vs even games)']:.2f}, predicts unseen games {r['On its own: predicts unseen games']:+.2f}, goes with winning "
                               f"{r['On its own: goes with winning']:+.2f}, beyond Corsi and expected goals {r['On its own: beyond Corsi and expected goals']:+.2f}" for _, r in ber_rows.iterrows())
                  + (f". It is part of the frozen plain-average recipe for: {', '.join(in_frozen)}." if in_frozen else ". It is not in either plain-average recipe the search chose.")))
    best_form = forms.sort_values("Average block: predicts unseen games", ascending=False).iloc[0]
    lines.append(("What helped", "Average block, predicts unseen games, by form: " + "; ".join(f"{r.iloc[0]} {r.iloc[1]:+.2f}" for _, r in forms.iterrows())
                  + f". The strongest form is '{best_form.iloc[0]}'."))
    own_ok, battle_ok = tiers["own"]["held"]["other"].mean() >= FINISH_LINE, tiers["battle"]["held"]["other"].mean() >= FINISH_LINE
    lines.append(("Against the finish line (provisional)", f"On the search seasons, 'our own breakouts' is {'at or above' if own_ok else 'below'} the {FINISH_LINE:.2f} line "
                  f"and 'breakout battle' is {'at or above' if battle_ok else 'below'} it. These seasons chose the recipes, so this is not the verdict: "
                  "the verdict is the blind test on 2009-10 to 2013-14, run once when they have all downloaded."))
    lines.append(("The freeze", f"Both recipes are saved in reference/{FROZEN_FILE.name}. Next step:  python scripts/build_breakout_round2.py --blind-test"))
    verdict = pd.DataFrame(lines, columns=["", "Answer"])

    sheets = {"Verdict": verdict, "Summary": summary, "By form": forms, "Chosen by season": pd.DataFrame(chosen_rows), "Blocks": blocks, "Top combinations": tops}
    notes = {"Summary": "'Honest' = the recipe was chosen without the season it was scored on. Still from the search seasons: the blind test is the judge.",
             "By form": "The average single block in each form. Shows whether close-game and battle versions help, before any combining.",
             "Chosen by season": "Which combination won on the OTHER seasons (by 'predicts unseen games') and how it did on the season left out, on every question.",
             "Blocks": "Each of the 64 blocks: how it does on its own, whether combinations that include it do better, and how often it is chosen.",
             "Top combinations": "The 25 best combinations for each recipe using ALL search seasons at once. Flattering: with this many tries the top is partly luck."}
    try:
        write_excel(OUTPUT_FILE, sheets, notes, READ_ME)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1
    pd.set_option("display.width", 250)
    pd.set_option("display.max_colwidth", 130)
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nSummary:"); print(summary.to_string(index=False))
    print("\nBy form:"); print(forms.to_string(index=False))
    print("\nChosen by season:"); print(pd.DataFrame(chosen_rows).iloc[:, :3].to_string(index=False))
    print(f"\nFrozen recipes saved: {FROZEN_FILE}")
    print(f"Saved: {OUTPUT_FILE}")
    print("Data: (c) NHL; Corsi, expected goals and ice time: MoneyPuck.com")
    return 0


if __name__ == "__main__":
    sys.exit(main())

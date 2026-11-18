"""
build_zone_chances.py  -  time in the offensive zone, and the scoring chances a team gets from it
=================================================================================================
WHAT IT DOES (in Excel terms):
    The point of the game is to create scoring chances. This study splits a team's chances into two parts:

        chances per 60 minutes  =  share of time in the offensive zone  x  chances created per minute in that zone
                                   ("how much")                            ("how well")

    and the same for defence: time spent in its own zone, and chances allowed per minute there. Then it tests,
    honestly, which of these goes with winning and predicts games the stats have not seen.

    THE MEASURES (all at 5-on-5, regular season)
      SCORING CHANCE   a shot on goal, a missed shot or a goal from the "home plate" area: from the goal posts out
                       to the faceoff dots, then straight out to the top of the circles. Read from the NHL's shot
                       coordinates. (Blocked shots are left out: their location is where the block happened.)
      ZONE TIME        the share of 5-on-5 time spent in each zone. Two sources:
                         * REAL: the NHL's EDGE puck tracking, 2021-22 on (used for the team tables on the website);
                         * ESTIMATE: from the play-by-play, for every season and every game (used for the tests,
                           which need game-by-game numbers). The time between two events is credited to the zone the
                           play was in. Four ways of doing that are compared with the real numbers on 2021-22 to
                           2023-24; the best is chosen there and confirmed on 2024-25 and 2025-26 before being used.
      CHANCES PER 10 MINUTES IN THE ZONE   chances for / minutes in the offensive zone x 10 (and the same against,
                       per 10 minutes in the team's own zone).
      CHANCE SHARE     chances for / (chances for + chances against).

    THE TESTS (finished seasons 2009-10 to 2025-26; 2019-20 and 2020-21 left out as unusual seasons)
      * REPEATS            odd-numbered games against even-numbered games: a real team trait or luck?
      * GOES WITH WINNING  against points % and goal difference in the same season (flattered: goals are chances too)
      * PREDICTS UNSEEN    the stat from odd-numbered games against the RESULTS of even-numbered games, and the reverse.
        GAMES              This is the fair test, because the games being predicted are not the ones the stat came from.
      * MORE THAN SHOT     the same, after taking out what shot share (Corsi), and then also expected goals, explain.
        SHARE / QUALITY
      * CONTENDERS         added to the contender score (version 1) and tried on seasons the score has never seen.
      * CHAMPIONS          where Cup winners and finalists ranked.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_zone_chances.py
    The first run reads every season's play-by-play (several minutes); later runs reuse saved copies and only
    re-read the season in progress.

WHAT YOU SHOULD SEE:
    A summary in the terminal and a new file:  outputs/zone_chances.xlsx   (close it in Excel first).

DATA CREDIT:
    Data is (c) NHL and NHL EDGE, used here for personal, non-commercial purposes. 5-on-5 ice time, Corsi and expected
    goals from MoneyPuck.com (non-commercial use, credit required).
"""

import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_breakout_index as bi          # season list, events, MoneyPuck halves, comparison numbers
import build_contender_fit as cf           # the contender score and its hold-out test
import build_popular_stats_check as ps     # season_spearman, zscore_by_season, prepare_games

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "zone_chances.xlsx"
EDGE_FILE = PROJECT_FOLDER / "data" / "nhl" / "edge" / "zone_time_regular.csv"
CACHE = PROJECT_FOLDER / "data" / "zone"
VERSION = 1                       # raise when the per-game counting changes, so saved copies are rebuilt

FIVE_ON_FIVE = "1551"
SHOT_TYPES = ["shot-on-goal", "missed-shot", "goal"]           # unblocked shots: their x,y is where the shot was taken
FLIP = {"O": "D", "D": "O", "N": "N"}
FIRST_SEASON, LAST_FINISHED = 20092010, 20252026
CHOOSE_ON, CONFIRM_ON = [20212022, 20222023, 20232024], [20242025, 20252026]
VARIANTS = {  # name: (how the time between two events is split, longest gap counted in seconds)
    "fwd30": ("all to the zone of the first event", 30), "fwd60": ("all to the zone of the first event", 60),
    "mid30": ("half to each event's zone", 30), "mid60": ("half to each event's zone", 60),
}

# The measures shown and tested: (column, plain name, higher is better?)
MEASURES = [
    ("chance_share", "Chance share (ours / both teams')", True),
    ("cf60", "Chances for per 60", True),
    ("ca60", "Chances against per 60", False),
    ("oz_share", "Offensive-zone time (share of 5-on-5 time)", True),
    ("cf_per10oz", "Chances per 10 minutes in the offensive zone", True),
    ("dz_share", "Own-zone time (share of 5-on-5 time)", False),
    ("ca_per10dz", "Chances against per 10 minutes in the own zone", False),
]


def season_label(season_id):
    text = str(int(season_id))
    return f"{text[:4]}-{text[6:]}"


def in_home_plate(x, y):
    """The scoring-chance area, in feet from centre ice (the goal line is 89 feet out, the faceoff dots 69 feet out and 22 feet wide)."""
    ax, ay = x.abs(), y.abs()
    slot = (ax >= 54) & (ax < 69) & (ay <= 22)                                  # between the dots, up to the top of the circles
    low = (ax >= 69) & (ax <= 89) & (ay <= 4 + (89 - ax) * 18 / 20)             # narrowing from the dots to the goal posts
    return slot | low


# ----------------------------------------------------------------------------
# Per game: zone seconds (four ways) and chances, for each team
# ----------------------------------------------------------------------------
def game_rows(ev):
    ev = ev.sort_values(["game_id", "sort_order"]).reset_index(drop=True)
    ev["abs_t"] = (ev["period"] - 1) * 1200 + ev["time_s"]
    has_zone = ev["zone"].isin(["O", "N", "D"]) & ev["team_id"].notna()
    owner_home = ev["team_id"].eq(ev["home_team_id"])
    tag_is_home = np.where(ev["type"].eq("blocked-shot"), ~owner_home, owner_home)     # blocked shots are tagged from the blocker's side
    ev["hz"] = pd.Series(np.where(tag_is_home, ev["zone"], ev["zone"].map(FLIP)), index=ev.index).where(has_zone)   # zone from the HOME team's side
    by_period = ev.groupby(["game_id", "period"])
    gap = (by_period["abs_t"].shift(-1) - ev["abs_t"]).clip(lower=0)
    next_zone = by_period["hz"].shift(-1)
    counted = ev["situation_code"].eq(FIVE_ON_FIVE) & ev["hz"].notna() & gap.notna()

    pieces = []
    for name, (_, cap) in VARIANTS.items():
        g = gap.clip(upper=cap).where(counted, 0.0)
        if name.startswith("fwd"):
            parts = [(ev["hz"], g)]
        else:
            split = next_zone.notna()
            parts = [(ev["hz"], g * np.where(split, 0.5, 1.0)), (next_zone, g.where(split, 0.0) * 0.5)]
        frame = pd.concat([pd.DataFrame({"game_id": ev["game_id"], "zone": z, "s": s}) for z, s in parts])
        sec = frame[frame["s"] > 0].pivot_table(index="game_id", columns="zone", values="s", aggfunc="sum").reindex(columns=["O", "N", "D"]).fillna(0)
        sec.columns = [f"home_{z}_{name}" for z in sec.columns]
        pieces.append(sec)
    home = pd.concat(pieces, axis=1).reset_index()

    teams = ev.groupby("game_id")[["home_team_id", "away_team_id"]].first().reset_index()
    home = teams.merge(home, on="game_id", how="left").fillna(0)
    rows = []
    for side in ("home", "away"):
        r = pd.DataFrame({"game_id": home["game_id"], "team_id": home[f"{side}_team_id"].astype(int)})
        for name in VARIANTS:
            o, n, d = home[f"home_O_{name}"], home[f"home_N_{name}"], home[f"home_D_{name}"]
            r[f"oz_{name}"], r[f"nz_{name}"], r[f"dz_{name}"] = (o, n, d) if side == "home" else (d, n, o)
        rows.append(r)
    rows = pd.concat(rows, ignore_index=True)

    shots = ev[ev["situation_code"].eq(FIVE_ON_FIVE) & ev["type"].isin(SHOT_TYPES) & ev["zone"].eq("O") & ev["x"].notna()].copy()
    shots["chance"] = in_home_plate(shots["x"], shots["y"])
    shots["opp"] = shots["home_team_id"] + shots["away_team_id"] - shots["team_id"]
    cf_ = shots.groupby(["game_id", "team_id"]).agg(cf=("chance", "sum"), ub_f=("chance", "size")).reset_index()
    ca_ = shots.groupby(["game_id", "opp"]).agg(ca=("chance", "sum"), ub_a=("chance", "size")).reset_index().rename(columns={"opp": "team_id"})
    for frame in (cf_, ca_):
        frame["team_id"] = frame["team_id"].astype(int)
    return rows.merge(cf_, on=["game_id", "team_id"], how="left").merge(ca_, on=["game_id", "team_id"], how="left").fillna(0)


def load_rows(pbp, nhl):
    """Every season with play-by-play from 2009-10 on: per-team-per-game rows (saved in data/zone/ after the first run)."""
    coverage = bi.season_coverage(pbp, nhl)
    current = int(coverage["season_id"].max())
    ready = [int(r.season_id) for r in coverage.itertuples() if r.season_id >= FIRST_SEASON and (r.coverage >= bi.MIN_COVERAGE or r.season_id == current) and r.saved > 0]
    CACHE.mkdir(parents=True, exist_ok=True)
    pieces = []
    for season in ready:
        saved = int(coverage.loc[coverage["season_id"] == season, "saved"].iloc[0])
        path = CACHE / f"zone_rows_{season}_v{VERSION}_{saved}.pkl"
        if path.exists():
            rows = pd.read_pickle(path)
        else:
            for old in CACHE.glob(f"zone_rows_{season}_*.pkl"):
                old.unlink()
            rows = game_rows(bi.load_events(pbp, season)).assign(season_id=season)
            rows.to_pickle(path)
            print(f"  {season_label(season)}: {rows['game_id'].nunique():,} games read")
        pieces.append(rows)
    rows = pd.concat(pieces, ignore_index=True)
    toi = pd.read_sql_query("SELECT game_id, team_id, iceTime AS toi5 FROM mp_team_games WHERE situation = '5on5' AND playoffGame = 0", nhl)
    rows = rows.merge(toi, on=["game_id", "team_id"], how="left")
    clock = rows[["oz_fwd60", "nz_fwd60", "dz_fwd60"]].sum(axis=1)
    rows["toi5"] = rows["toi5"].fillna(clock.where(clock > 0))                    # newest games: 5-on-5 time counted from the events
    return rows.dropna(subset=["toi5"]), current


# ----------------------------------------------------------------------------
# Team numbers
# ----------------------------------------------------------------------------
def team_numbers(rows, variant, edge=None):
    """One row per team-season (or per whatever is in rows). Zone time from the estimate, or the real EDGE share if given."""
    g = rows.groupby(["season_id", "team_id"]).agg(
        games=("game_id", "nunique"), oz=(f"oz_{variant}", "sum"), nz=(f"nz_{variant}", "sum"), dz=(f"dz_{variant}", "sum"),
        cf=("cf", "sum"), ca=("ca", "sum"), ub_f=("ub_f", "sum"), ub_a=("ub_a", "sum"), toi5=("toi5", "sum")).reset_index()
    total = g[["oz", "nz", "dz"]].sum(axis=1)
    g["oz_share"], g["nz_share"], g["dz_share"] = g["oz"] / total, g["nz"] / total, g["dz"] / total
    if edge is not None:
        g = g.drop(columns=["oz_share", "nz_share", "dz_share"]).merge(edge, on=["season_id", "team_id"], how="inner")
    minutes = g["toi5"] / 60
    g["cf60"], g["ca60"] = g["cf"] / minutes * 60, g["ca"] / minutes * 60
    g["cf_per10oz"] = g["cf"] / (minutes * g["oz_share"]) * 10
    g["ca_per10dz"] = g["ca"] / (minutes * g["dz_share"]) * 10
    g["chance_share"] = g["cf"] / (g["cf"] + g["ca"])
    g["oz_min60"], g["dz_min60"] = g["oz_share"] * 60, g["dz_share"] * 60
    return g


def load_edge():
    e = pd.read_csv(EDGE_FILE, encoding="utf-8-sig").rename(columns={"teamId": "team_id", "es_oz_pct": "oz_share", "es_nz_pct": "nz_share", "es_dz_pct": "dz_share"})
    e["team_id"] = e["team_id"].astype(int)
    return e[["season_id", "team_id", "oz_share", "nz_share", "dz_share"]].dropna()


def choose_estimate(rows, edge):
    """Compare the four ways of estimating zone time with the real EDGE numbers; choose on 2021-22 to 2023-24, confirm on the next two."""
    out = []
    for name, (how, cap) in VARIANTS.items():
        t = team_numbers(rows, name).merge(edge, on=["season_id", "team_id"], suffixes=("", "_edge"))
        row = {"Estimate": name, "How the time between two events is split": how, "Longest gap counted (seconds)": cap}
        for label, seasons in (("Choosing seasons", CHOOSE_ON), ("Confirming seasons", CONFIRM_ON)):
            part = t[t["season_id"].isin(seasons)]
            oz = ps.season_spearman(part, "oz_share", "oz_share_edge").mean()
            dz = ps.season_spearman(part, "dz_share", "dz_share_edge").mean()
            row[f"{label}: offensive zone"] = round(oz, 3)
            row[f"{label}: own zone"] = round(dz, 3)
            row[f"{label}: average"] = round((oz + dz) / 2, 3)
            row[f"{label}: average gap to EDGE (percentage points, offensive zone)"] = round(float((part["oz_share"] - part["oz_share_edge"]).abs().mean() * 100), 2)
        out.append(row)
    table = pd.DataFrame(out)
    best = table.loc[table["Choosing seasons: average"].idxmax(), "Estimate"]
    return table, best


# ----------------------------------------------------------------------------
# The tests
# ----------------------------------------------------------------------------
def unit(a):
    a = a - a.mean(axis=0)
    return a / np.sqrt((a ** 2).sum(axis=0))


def partial(d, x, y, controls):
    """Rank correlation of x with y after taking out the control columns (straight-line fit on ranks), within one season."""
    d = d.dropna(subset=[x, y] + controls)
    if len(d) < 10:
        return np.nan
    rx, ry = unit(d[[x]].rank().to_numpy()), unit(d[y].rank().to_numpy()[:, None])
    if controls:
        k = unit(d[controls].rank().to_numpy())
        rx = unit(rx - k @ np.linalg.lstsq(k, rx, rcond=None)[0])
        ry = unit(ry - k @ np.linalg.lstsq(k, ry, rcond=None)[0])
    return float((rx.T @ ry)[0, 0])


def halves(rows, variant, nhl, seasons):
    """The stats from odd-numbered games next to the results (and Corsi, xG) of even-numbered games, and the reverse."""
    ids = ",".join(str(int(s)) for s in seasons)
    results = pd.read_sql_query(f"SELECT season_id, game_id, team_id, points, goalsFor - goalsAgainst AS gd FROM team_games WHERE game_type = 2 AND season_id IN ({ids})", nhl)
    bench = bi.benchmark_halves(nhl, seasons)
    frames, same = [], []
    for parity, suffix, other_suffix in ((1, "_a", "_b"), (0, "_b", "_a")):
        stats = team_numbers(rows[rows["game_id"] % 2 == parity], variant)
        r = results[results["game_id"] % 2 != parity].groupby(["season_id", "team_id"]).agg(pts=("points", "sum"), n=("points", "size"), gd=("gd", "sum")).reset_index()
        r["pts_next"], r["gd_next"] = r["pts"] / (2 * r["n"]), r["gd"] / r["n"]
        b = bench[["season_id", "team_id", f"corsi{suffix}", f"xg_share{suffix}"]].rename(columns={f"corsi{suffix}": "corsi", f"xg_share{suffix}": "xg_share"})
        frames.append(stats.merge(r[["season_id", "team_id", "pts_next", "gd_next"]], on=["season_id", "team_id"]).merge(b, on=["season_id", "team_id"]))
        same.append(stats)
    return frames, same


def test_tables(rows, variant, nhl, seasons, season_table):
    frames, same = halves(rows, variant, nhl, seasons)
    rep = same[0].merge(same[1], on=["season_id", "team_id"], suffixes=("_a", "_b"))
    columns = MEASURES + [("corsi", "Yardstick: Corsi (shot-attempt share)", True), ("xg_share", "Yardstick: expected-goals share", True)]
    bench = bi.benchmark_halves(nhl, seasons)
    repeat_rows, unseen_rows = [], []
    for col, name, higher in columns:
        sign = 1 if higher else -1
        if col in ("corsi", "xg_share"):
            r = ps.season_spearman(bench, f"{col}_a", f"{col}_b")
        else:
            r = ps.season_spearman(rep, f"{col}_a", f"{col}_b")
        half = float(r.mean())
        repeat_rows.append({"Statistic": name, "Repeats (odd vs even games)": round(half, 2), "Estimate for a full season": round(2 * half / (1 + half), 2)})

        per_season = {"plain": [], "beyond_corsi": [], "beyond_both": []}
        for s in seasons:
            vals = {"plain": [], "beyond_corsi": [], "beyond_both": []}
            for f in frames:
                d = f[f["season_id"] == s]
                for key, controls in (("plain", []), ("beyond_corsi", ["corsi"]), ("beyond_both", ["corsi", "xg_share"])):
                    if col in controls:
                        continue
                    vals[key].append(np.nanmean([partial(d, col, "pts_next", controls), partial(d, col, "gd_next", controls)]) * sign)
            for key in per_season:
                per_season[key].append(np.mean(vals[key]) if vals[key] else np.nan)
        row = {"Statistic": name}
        for key, label in (("plain", "Predicts unseen games"), ("beyond_corsi", "Beyond shot share (Corsi)"), ("beyond_both", "Beyond shot share and expected goals")):
            v = np.array(per_season[key], dtype=float)
            if np.isnan(v).all():
                row[label], row[f"{label}: noise margin"] = None, None
                continue
            v = v[~np.isnan(v)]
            row[label] = round(float(v.mean()), 3)
            row[f"{label}: noise margin"] = round(float(2 * v.std(ddof=1) / np.sqrt(len(v))), 3)
            if key == "plain":
                row["Seasons above zero"] = int((v > 0).sum())
        # same season (flattered: goals are chances too)
        if col not in ("corsi", "xg_share"):
            row["Goes with winning, same season (points %)"] = round(float(ps.season_spearman(season_table, col, "pts_pct").mean()) * sign, 2)
        else:
            row["Goes with winning, same season (points %)"] = round(float(ps.season_spearman(season_table, col, "pts_pct").mean()), 2)
        unseen_rows.append(row)

    # time in the zone or what you do with it: each one with the other taken out
    both = []
    for col, other, name in (("oz_share", "cf_per10oz", "Offensive-zone time"), ("cf_per10oz", "oz_share", "Chances per 10 minutes in the offensive zone"),
                             ("dz_share", "ca_per10dz", "Own-zone time"), ("ca_per10dz", "dz_share", "Chances against per 10 minutes in the own zone")):
        sign = -1 if col in ("dz_share", "ca_per10dz") else 1
        v = []
        for s in seasons:
            v.append(np.mean([np.nanmean([partial(f[f["season_id"] == s], col, "pts_next", [other]), partial(f[f["season_id"] == s], col, "gd_next", [other])]) for f in frames]) * sign)
        v = np.array(v)
        overlap = np.mean([ps.season_spearman(f, col, other).mean() for f in frames])
        both.append({"Statistic": name, "Predicts unseen games, with the other part taken out": round(float(v.mean()), 3),
                     "Noise margin": round(float(2 * v.std(ddof=1) / np.sqrt(len(v))), 3), "How much the two parts move together": round(float(overlap), 2)})
    return pd.DataFrame(repeat_rows), pd.DataFrame(unseen_rows), pd.DataFrame(both)


def contender_test(rows, variant, nhl, seasons):
    """Each measure added on its own to the contender score (version 1), tried on seasons the score has never seen."""
    games_raw, seasons_tbl, checkpoints = cf.load(nhl)
    games, _ = ps.prepare_games(games_raw, seasons_tbl)
    tables = cf.build_checkpoint_tables(games, seasons_tbl, checkpoints)
    dates = pd.read_sql_query("SELECT game_id, game_date FROM games", nhl)
    rows = rows.merge(dates, on="game_id")
    v1 = cf.MODELS["Model 3: + power kill (PP% + PK%)"]
    keys = ["chance_share", "oz_share", "cf_per10oz", "ca_per10dz"]
    gains = {k: [] for k in keys}
    out = []
    for cp in cf.CHECKPOINTS:
        d = checkpoints[checkpoints["checkpoint"] == cp][["season_id", "date"]]
        r = rows.merge(d, on="season_id")
        feats = team_numbers(r[r["game_date"] <= r["date"]], variant)
        for k in keys:
            z = ps.zscore_by_season(feats, k)
            feats[f"z_{k}"] = (z if k != "ca_per10dz" else -z).fillna(0.0)
        t = tables[cp].merge(feats[["season_id", "team_id"] + [f"z_{k}" for k in keys]], on=["season_id", "team_id"], how="inner")
        t = t[t["season_id"].isin(seasons)].copy()
        if t["season_id"].nunique() < 5:
            continue
        t["v1"] = cf.held_out_scores(t, v1)
        base = cf.per_season_metrics(t, "v1")
        row = {"Checkpoint": cf.PLAIN_CP[cp], "Seasons": int(t["season_id"].nunique()), "AUC version 1": round(float(base["auc"].mean()), 3)}
        for k in keys:
            t["plus"] = cf.held_out_scores(t, v1 + [k])
            diff = cf.per_season_metrics(t, "plus")["auc"] - base["auc"]
            gains[k].append(diff)
            row[f"Gain: {k}"] = round(float(diff.mean()), 3)
            row[f"Margin: {k}"] = round(float(2 * diff.std(ddof=1) / np.sqrt(len(diff))), 3)
        out.append(row)
    summary = {}
    for k, parts in gains.items():
        by_season = pd.concat(parts, axis=1).mean(axis=1)
        summary[k] = (float(by_season.mean()), float(2 * by_season.std(ddof=1) / np.sqrt(len(by_season))))
    return pd.DataFrame(out), summary


def champions(season_table):
    d = season_table[season_table["outcome"].isin(cf.OUTCOME_ORDER)].copy()
    rows = []
    for outcome in cf.OUTCOME_ORDER + ["ALL teams"]:
        part = d if outcome == "ALL teams" else d[d["outcome"] == outcome]
        row = {"How the season ended": outcome, "Team-seasons": len(part)}
        for col, name, higher in MEASURES:
            pct = d.groupby("season_id")[col].rank(pct=True, ascending=higher) * 100
            row[f"{name} (percentile)"] = round(float(pct[part.index].mean()), 0)
        rows.append(row)
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
def tables_for(t, season):
    d = t[t["season_id"] == season].sort_values("chance_share", ascending=False)
    return pd.DataFrame({"Team": d["team"], "Games": d["games"].astype(int),
                         "Offensive-zone minutes per 60": d["oz_min60"].round(1), "Chances per 10 offensive-zone minutes": d["cf_per10oz"].round(2),
                         "Chances for per 60": d["cf60"].round(2), "Own-zone minutes per 60": d["dz_min60"].round(1),
                         "Chances against per 10 own-zone minutes": d["ca_per10dz"].round(2), "Chances against per 60": d["ca60"].round(2),
                         "Chance share": d["chance_share"].round(3), "Points %": d["pts_pct"].round(3), "How the season ended": d["outcome"]}).reset_index(drop=True)


def verdict_lines(est_table, best, unseen, both, cont, champs, seasons, league):
    u = unseen.set_index("Statistic")
    b = both.set_index("Statistic")
    best_row = est_table.set_index("Estimate").loc[best]

    def g(name, col="Predicts unseen games"):
        return float(u.loc[name, col])

    share, corsi, xg = g("Chance share (ours / both teams')"), g("Yardstick: Corsi (shot-attempt share)"), g("Yardstick: expected-goals share")
    beyond = u.loc["Chance share (ours / both teams')"]
    oz_alone, rate_alone = b.loc["Offensive-zone time"], b.loc["Chances per 10 minutes in the offensive zone"]
    lines = [
        ("What was done", f"Scoring chances (shots from the home-plate area) and time in each zone at 5-on-5, for {len(seasons)} seasons "
                          f"({season_label(min(seasons))} to {season_label(max(seasons))}, without 2019-20 and 2020-21). League average: {league['cf60']:.1f} chances per 60 minutes, "
                          f"{league['share_of_unblocked']:.0%} of shots that reach the net or miss it, {league['per10']:.1f} chances per 10 minutes in the offensive zone."),
        ("Short answer", f"Creating chances is what goes with winning. Chance share predicts the results of games it has not seen at {share:+.2f}, "
                         f"against {corsi:+.2f} for shot-attempt share (Corsi) and {xg:+.2f} for expected goals. "
                         + ("It still adds something after shot share is taken out" if beyond["Beyond shot share (Corsi)"] > beyond["Beyond shot share (Corsi): noise margin"] else "Once shot share is taken out, little is left")
                         + f" ({beyond['Beyond shot share (Corsi)']:+.2f}, noise margin {beyond['Beyond shot share (Corsi): noise margin']:.2f}), and after expected goals as well "
                         f"{beyond['Beyond shot share and expected goals']:+.2f} (margin {beyond['Beyond shot share and expected goals: noise margin']:.2f})."),
        ("Time in the zone, or what you do with it?", f"Both count, and they are different skills (they move together at only {oz_alone['How much the two parts move together']:+.2f}). "
                         f"With the other taken out, offensive-zone time predicts unseen games at {oz_alone['Predicts unseen games, with the other part taken out']:+.2f} and chances per "
                         f"minute in the zone at {rate_alone['Predicts unseen games, with the other part taken out']:+.2f} (noise margins about {oz_alone['Noise margin']:.2f})."),
        ("The zone-time estimate", f"Four ways of estimating zone time from the play-by-play were compared with the NHL's tracking. The best on 2021-22 to 2023-24 "
                                   f"('{best_row['How the time between two events is split']}', gaps up to {int(best_row['Longest gap counted (seconds)'])} seconds) agreed at "
                                   f"{best_row['Choosing seasons: average']:+.2f}, and at {best_row['Confirming seasons: average']:+.2f} on 2024-25 and 2025-26, which it had not been chosen on. "
                                   "The website's team tables use the NHL's real numbers; the tests use the estimate because they need game-by-game numbers."),
    ]
    if len(cont[1]):
        s = cont[1]["chance_share"]
        lines.append(("Does it help pick contenders?", f"Added to the contender score and tried on seasons the score had not seen, chance share changes the AUC by {s[0]:+.3f} "
                                                       f"(noise margin {s[1]:.3f}): " + ("more than the noise." if s[0] > s[1] else "within the noise.")))
    c = champs.set_index("How the season ended")
    key = MEASURES[0][1] + " (percentile)"
    if "Won Cup" in c.index:
        lines.append(("Champions", f"Cup winners averaged the {c.loc['Won Cup', key]:.0f}th percentile on chance share (100 = best), finalists the "
                                   f"{c.loc['Lost Cup final', key]:.0f}th, teams that missed the playoffs the {c.loc['Missed playoffs', key]:.0f}th."))
    return pd.DataFrame(lines, columns=["", "Answer"])


def write_excel(sheets, notes, read_me):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": read_me}).to_excel(writer, sheet_name="Read me", index=False)
        writer.sheets["Read me"].column_dimensions["A"].width = 140
        for name, frame in sheets.items():
            start = 1 if name in notes else 0
            frame.to_excel(writer, sheet_name=name, index=False, startrow=start)
            ws = writer.sheets[name]
            if name in notes:
                ws.cell(row=1, column=1, value=notes[name]).font = Font(bold=True)
            for cell in ws[start + 1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            ws.freeze_panes = ws.cell(row=start + 2, column=2)
            ws.column_dimensions["A"].width = 34
        ws = writer.sheets["Verdict"]
        ws.column_dimensions["B"].width = 140
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(bold=True)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")


def main():
    sys.stdout.reconfigure(errors="replace")
    if not bi.PBP_DB.exists() or not bi.NHL_DB.exists() or not EDGE_FILE.exists():
        print("I need data/pbp.db, data/nhl.db and data/nhl/edge/zone_time_regular.csv.")
        return 1
    pbp = sqlite3.connect(f"file:{bi.PBP_DB}?mode=ro", uri=True)
    nhl = sqlite3.connect(bi.NHL_DB)
    print("Reading the play-by-play (saved copies are reused)...")
    rows, current = load_rows(pbp, nhl)
    edge = load_edge()

    flags = pd.read_csv(PROJECT_FOLDER / "reference" / "season_flags.csv")
    odd = set(int(s) for s in flags.loc[flags["include_main"] == "N", "season_id"])
    seasons = [s for s in sorted(rows["season_id"].unique()) if FIRST_SEASON <= s <= LAST_FINISHED and s not in odd]
    finished = rows[rows["season_id"].isin(seasons)]

    est_table, best = choose_estimate(finished[finished["season_id"].isin(CHOOSE_ON + CONFIRM_ON)], edge)
    print(f"Zone-time estimate chosen: {best}")
    compare = bi.comparison_numbers(nhl, sorted(rows["season_id"].unique()))[["season_id", "team_id", "team", "pts_pct", "gd_pg", "corsi", "xg_share", "outcome"]]
    season_table = team_numbers(finished, best).merge(compare, on=["season_id", "team_id"], how="left")
    repeats, unseen, both = test_tables(finished, best, nhl, seasons, season_table)
    cont = contender_test(finished, best, nhl, seasons)
    champs = champions(season_table)
    league = {"cf60": float(season_table["cf60"].mean()), "share_of_unblocked": float(season_table["cf"].sum() / season_table["ub_f"].sum()),
              "per10": float(season_table["cf_per10oz"].mean())}
    verdict = verdict_lines(est_table, best, unseen, both, cont, champs, seasons, league)

    # the website's team tables: real NHL EDGE zone time, this season and last season
    real = team_numbers(rows[rows["season_id"] >= min(edge["season_id"])], best, edge=edge).merge(compare, on=["season_id", "team_id"], how="left")
    dates = pd.read_sql_query("SELECT game_id, game_date FROM games WHERE season_id = ?", nhl, params=[current])
    through = dates[dates["game_id"].isin(rows.loc[rows["season_id"] == current, "game_id"])]["game_date"].max()
    this_season, last_season = tables_for(real, current), tables_for(real, LAST_FINISHED)

    read_me = ["Time in the offensive zone, and the scoring chances a team gets from it", "",
               "Scoring chance: a shot on goal, missed shot or goal at 5-on-5 from the home-plate area (from the goal posts out to the faceoff dots, then out to the top of the circles).",
               "Zone time: the share of 5-on-5 time in each zone. The team tables use the NHL's EDGE tracking (2021-22 on); the tests use an estimate from the play-by-play,",
               "chosen against EDGE on 2021-22 to 2023-24 and confirmed on 2024-25 and 2025-26 ('Estimate check').",
               "Chances per 10 minutes in the zone = chances / minutes spent in that zone x 10. Chance share = chances for / (for + against).",
               "Tests are rank correlations within each season, averaged: 0 = no relationship, 1 = perfect. 'Predicts unseen games' compares the stat from odd-numbered games",
               "with the results of even-numbered games (and the reverse), so it is not flattered by the fact that goals are chances too.", "",
               "Data: (c) NHL and NHL EDGE. 5-on-5 ice time, Corsi and expected goals: MoneyPuck.com."]
    sheets = {"Verdict": verdict, "This season": this_season, "Last season": last_season, "Unseen games": unseen, "Time or what you do": both,
              "Repeats": repeats, "Contender test": cont[0], "Champions": champs, "Estimate check": est_table,
              "Team-seasons": season_table.sort_values(["season_id", "chance_share"], ascending=[False, False]).round(4)}
    notes = {"This season": f"Season {season_label(current)}: games through {through}. Zone time from NHL EDGE. Early in a season these numbers are noisy.",
             "Last season": f"Season {season_label(LAST_FINISHED)}. Zone time from NHL EDGE.",
             "Unseen games": "Each stat built from odd-numbered games against the results of even-numbered games (and the reverse); turned so positive = good. Noise margin = 2 standard errors across seasons.",
             "Time or what you do": "Offensive-zone time and chances per minute in the zone, each with the other taken out, predicting unseen games.",
             "Contender test": "Each measure added on its own to the contender score (version 1) and tried on seasons the score had not seen: change in AUC.",
             "Estimate check": "Rank agreement between the play-by-play estimate of zone time and the NHL's EDGE tracking, per season, averaged."}
    try:
        write_excel(sheets, notes, read_me)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 250)
    pd.set_option("display.max_colwidth", 140)
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nEstimate check:"); print(est_table.to_string(index=False))
    print("\nUnseen games:"); print(unseen.to_string(index=False))
    print("\nTime or what you do with it:"); print(both.to_string(index=False))
    print("\nRepeats:"); print(repeats.to_string(index=False))
    print("\nContender test:"); print(cont[0].to_string(index=False))
    print("\nChampions:"); print(champs.to_string(index=False))
    print(f"\nThis season (through {through}), top 5:"); print(this_season.head(5).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL and NHL EDGE; ice time, Corsi and expected goals: MoneyPuck.com")
    return 0


if __name__ == "__main__":
    sys.exit(main())

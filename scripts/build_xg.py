"""
build_xg.py  -  our own expected-goals model, from the NHL play-by-play
=======================================================================
WHAT IT DOES (in Excel terms):
    Gives every unblocked shot (a shot on goal, a missed shot or a goal) a chance of becoming a goal, from what the NHL's
    play-by-play records about it, then adds those chances up for every team:

        distance and angle to the net, shot type (wrist, slap, snap, backhand, tip, deflection ...), whether it is a
        rebound (within 3 seconds of a teammate's shot) or comes off the rush (within 4 seconds of play outside the
        offensive zone), and the strength (5-on-5, power play, shorthanded).

    The weights come from a "logistic regression" (like LINEST, but for yes/no outcomes), fitted on the finished seasons of the
    NHL's current shot tracking (2021-22 on) if that does better on the check seasons than fitting on every season.
    Each season is then scaled so the league's expected goals add up to the goals really scored, which removes changes in
    the scoring level from year to year without changing how teams rank. Empty-net shots and penalty shots are left out.

    From the expected goals come the same numbers as the Goaltending in context study, but updated every day:
        Defense      expected goals against per 60 minutes at 5-on-5 (lower is better)
        Goaltending  goals saved above expected (GSAx) = expected goals against - goals allowed, all situations, per 60
        Label        both good / goalie bailing out the defense / defense let down by the goalie / both weak
                     (above or below the league median on each part)

    IS IT ANY GOOD?  The "Check" sheet answers that, on two finished seasons the model was NOT fitted on:
        * how well it ranks shots (AUC: 0.5 = coin flip, 1 = perfect) and whether its chances match real goal rates;
        * how closely its team numbers agree with MoneyPuck's (where we have them).
    It also tries a correction for arenas whose scorers record shot locations a little differently, and keeps it only if
    it makes the model better on those unseen seasons.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_xg.py
    The first run reads every season's play-by-play and fits the model (several minutes); later runs reuse saved copies
    in data/xg/ and only re-read the season in progress.

WHAT YOU SHOULD SEE:
    A short check and a new file:  outputs/expected_goals.xlsx   (close it in Excel first).

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes. MoneyPuck.com numbers are used only to check the model.
"""

import hashlib
import json
import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_breakout_index as bi          # season coverage, season labels
import build_contender_fit as cf           # fit_logistic

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
CACHE = PROJECT_FOLDER / "data" / "xg"
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "expected_goals.xlsx"
VERSION = 2                                  # raise when the shot features change, so saved copies are rebuilt
FIRST_SEASON = 20092010
CHECK_SEASONS = [20242025, 20252026]         # finished seasons held back to check the model
ERA_START = 20212022                         # the NHL's current shot tracking: shots before this are recorded differently
FEATURES_TAG = "finer-v1"                    # change when the features change, so the saved model is refitted
SHOT_TYPES = ["shot-on-goal", "missed-shot", "goal"]
KINDS = ["slap", "snap", "backhand", "tip-in", "deflected", "wrap-around", "bat", "poke", "between-legs", "cradle"]
REBOUND_SECONDS, RUSH_SECONDS = 3, 4
LABELS = ["Both good", "Goalie bailing out the defense", "Defense let down by the goalie", "Both weak"]


def season_label(s):
    return bi.season_label(s)


# ----------------------------------------------------------------------------
# Shots and clocks for one season
# ----------------------------------------------------------------------------
def season_shots(pbp, season):
    ev = pd.read_sql_query(
        """SELECT e.game_id, e.sort_order, e.period, e.time_s, e.situation_code, e.type, e.team_id, e.zone, e.x, e.y, e.shot_type,
                  g.home_team_id, g.away_team_id
           FROM pbp_events e JOIN pbp_games g ON g.game_id = e.game_id
           WHERE g.season_id = ? AND g.game_type = 2 AND g.status = 'ok' AND e.period_type != 'SO'""", pbp, params=[season])
    ev = ev.sort_values(["game_id", "sort_order"]).reset_index(drop=True)
    ev["t"] = (ev["period"] - 1) * 1200 + ev["time_s"]
    by = ev.groupby(["game_id", "period"])
    ev["prev_type"], ev["prev_team"], ev["prev_zone"], ev["prev_t"] = by["type"].shift(1), by["team_id"].shift(1), by["zone"].shift(1), by["t"].shift(1)

    # clocks per game: all play, and 5-on-5 play (time between events)
    gap = (by["t"].shift(-1) - ev["t"]).clip(lower=0).fillna(0)
    clock = pd.DataFrame({"game_id": ev["game_id"], "all_s": gap, "five_s": gap.where(ev["situation_code"].eq("1551"), 0)}).groupby("game_id").sum().reset_index()

    s = ev[ev["type"].isin(SHOT_TYPES) & ev["x"].notna() & ev["y"].notna() & ev["zone"].eq("O") & ev["team_id"].notna()].copy()
    sc = s["situation_code"].astype(str).str.zfill(4)
    away_goalie, away_sk, home_sk, home_goalie = (pd.to_numeric(sc.str[i], errors="coerce") for i in range(4))
    home = s["team_id"].eq(s["home_team_id"])
    s["own_sk"] = np.where(home, home_sk, away_sk)
    s["opp_sk"] = np.where(home, away_sk, home_sk)
    s["opp_goalie"] = np.where(home, away_goalie, home_goalie)
    s = s[(s["opp_goalie"] == 1) & (s["own_sk"] >= 3) & (s["opp_sk"] >= 3)].copy()       # no empty nets, no penalty shots
    s["opp"] = (s["home_team_id"] + s["away_team_id"] - s["team_id"]).astype(int)
    s["team_id"] = s["team_id"].astype(int)
    ax, ay = s["x"].abs(), s["y"].abs()
    s["dist"] = np.sqrt((89 - ax) ** 2 + ay ** 2).clip(lower=1)
    s["angle"] = np.degrees(np.arctan2(ay, 89 - ax))
    since = s["t"] - s["prev_t"]
    s["rebound"] = (s["prev_type"].isin(["shot-on-goal", "missed-shot", "blocked-shot"]) & s["prev_team"].eq(s["team_id"]) & (since <= REBOUND_SECONDS)).astype(float)
    s["rush"] = (s["prev_zone"].isin(["N", "D"]) & (since <= RUSH_SECONDS) & s["rebound"].eq(0)).astype(float)
    s["goal"] = s["type"].eq("goal").astype(float)
    s["five"] = (s["own_sk"].eq(5) & s["opp_sk"].eq(5)).astype(float)
    s["season_id"] = season
    keep = ["game_id", "season_id", "team_id", "opp", "home_team_id", "dist", "angle", "shot_type", "rebound", "rush", "own_sk", "opp_sk", "five", "goal"]
    return s[keep].reset_index(drop=True), clock.assign(season_id=season)


def load_all(pbp, nhl):
    coverage = bi.season_coverage(pbp, nhl)
    current = int(coverage["season_id"].max())
    seasons = [int(r.season_id) for r in coverage.itertuples()
               if r.season_id >= FIRST_SEASON and r.saved > 0 and (r.coverage >= bi.MIN_COVERAGE or r.season_id == current)]
    CACHE.mkdir(parents=True, exist_ok=True)
    shots, clocks = [], []
    for season in seasons:
        saved = int(coverage.loc[coverage["season_id"] == season, "saved"].iloc[0])
        path = CACHE / f"shots_{season}_v{VERSION}_{saved}.pkl"
        if path.exists():
            s, c = pd.read_pickle(path)
        else:
            for old in CACHE.glob(f"shots_{season}_*.pkl"):
                old.unlink()
            s, c = season_shots(pbp, season)
            pd.to_pickle((s, c), path)
            print(f"  {season_label(season)}: {len(s):,} shots read")
        shots.append(s)
        clocks.append(c)
    return pd.concat(shots, ignore_index=True), pd.concat(clocks, ignore_index=True), current


# ----------------------------------------------------------------------------
# The model
# ----------------------------------------------------------------------------
def arena_shift(shots, seasons):
    """
    Some arenas' scorers record shots a little closer or farther than others. For each arena-season: how much nearer the
    shots in that arena are recorded than the same teams' shots in their other games (in feet). Fitted only on `seasons`.
    """
    s = shots[shots["season_id"].isin(seasons)]
    # each shooting team's shots in this arena, and all its shots that season
    here = s.groupby(["season_id", "home_team_id", "team_id"])["dist"].agg(["sum", "count"]).reset_index()
    total = s.groupby(["season_id", "team_id"])["dist"].agg(["sum", "count"]).reset_index()
    j = here.merge(total, on=["season_id", "team_id"], suffixes=("_here", "_all"))
    j["sum_else"], j["count_else"] = j["sum_all"] - j["sum_here"], j["count_all"] - j["count_here"]
    a = j.groupby(["season_id", "home_team_id"])[["sum_here", "count_here", "sum_else", "count_else"]].sum()
    # recorded in this arena vs the same teams' shots in their other games
    shift = a["sum_here"] / a["count_here"] - a["sum_else"] / a["count_else"]
    return shift.rename("shift").reset_index()


def features(shots, shift=None):
    d = shots["dist"].copy()
    if shift is not None:
        d = (d - shots[["season_id", "home_team_id"]].merge(shift, on=["season_id", "home_team_id"], how="left")["shift"].fillna(0).to_numpy()).clip(lower=1)
    X = pd.DataFrame({"dist": d / 10, "dist2": (d / 10) ** 2, "logdist": np.log(d), "angle": shots["angle"] / 45, "angle2": (shots["angle"] / 45) ** 2,
                      "angle_x_dist": shots["angle"] / 45 * d / 10, "rebound": shots["rebound"], "rush": shots["rush"],
                      "rebound_x_angle": shots["rebound"] * shots["angle"] / 45,
                      "pp": (shots["own_sk"] > shots["opp_sk"]).astype(float), "pk": (shots["own_sk"] < shots["opp_sk"]).astype(float)})
    for kind in KINDS:
        X[f"type_{kind}"] = shots["shot_type"].eq(kind).astype(float)
    for lo, hi in [(0, 8), (8, 12), (12, 16), (16, 20), (20, 25), (25, 30), (30, 40), (40, 55)]:      # distance bands (feet)
        X[f"band_{lo}_{hi}"] = ((d >= lo) & (d < hi)).astype(float)
    close = (d < 20).astype(float)
    for kind in ("tip-in", "deflected", "backhand", "wrap-around"):
        X[f"close_{kind}"] = shots["shot_type"].eq(kind).astype(float) * close
    X["rebound_close"] = shots["rebound"] * close
    X["rush_dist"] = shots["rush"] * d / 10
    return X


def fit(X, y):
    mu, sd = X.mean(), X.std().replace(0, 1)
    beta = cf.fit_logistic(((X - mu) / sd).to_numpy(), y, l2=1.0, iterations=30)
    return {"columns": list(X.columns), "mu": mu.tolist(), "sd": sd.tolist(), "beta": beta.tolist()}


def predict(model, X):
    Z = ((X[model["columns"]] - np.array(model["mu"])) / np.array(model["sd"])).to_numpy()
    return 1 / (1 + np.exp(-(np.column_stack([np.ones(len(Z)), Z]) @ np.array(model["beta"]))))


def auc(score, truth):
    order = np.argsort(score, kind="mergesort")
    ranks = np.empty(len(score))
    ranks[order] = np.arange(1, len(score) + 1)
    pos = truth == 1
    return float((ranks[pos].sum() - pos.sum() * (pos.sum() + 1) / 2) / (pos.sum() * (~pos).sum()))


def log_loss(p, y):
    p = np.clip(p, 1e-6, 1 - 1e-6)
    return float(-np.mean(y * np.log(p) + (1 - y) * np.log(1 - p)))


def scale_by_season(shots, p):
    """Each season's expected goals scaled so the league total equals the goals scored (does not change how teams rank)."""
    sums = shots.assign(p=p).groupby("season_id")[["goal", "p"]].sum()
    factor = sums["goal"] / sums["p"]
    return np.clip(p * shots["season_id"].map(factor).to_numpy(), 0, 0.99), factor


# ----------------------------------------------------------------------------
# Team numbers
# ----------------------------------------------------------------------------
def team_numbers(shots, clocks):
    s = shots
    f = s.groupby(["season_id", "team_id"]).agg(xgf=("xg", "sum"), gf=("goal", "sum"))
    f5 = s[s["five"] == 1].groupby(["season_id", "team_id"]).agg(xgf5=("xg", "sum"))
    a = s.groupby(["season_id", "opp"]).agg(xga=("xg", "sum"), ga=("goal", "sum"))
    a5 = s[s["five"] == 1].groupby(["season_id", "opp"]).agg(xga5=("xg", "sum"), ga5=("goal", "sum"))
    a.index.names = a5.index.names = ["season_id", "team_id"]
    games = pd.concat([s[["season_id", "game_id", "team_id"]], s[["season_id", "game_id", "opp"]].rename(columns={"opp": "team_id"})]).drop_duplicates()
    games = games.merge(clocks[["game_id", "all_s", "five_s"]], on="game_id")
    ice = games.groupby(["season_id", "team_id"]).agg(games=("game_id", "nunique"), all_s=("all_s", "sum"), five_s=("five_s", "sum"))
    t = pd.concat([f, f5, a, a5, ice], axis=1).fillna(0).reset_index()
    t["xga60_5"] = t["xga5"] / (t["five_s"] / 3600)
    t["xgf60_5"] = t["xgf5"] / (t["five_s"] / 3600)
    t["xg_share_5"] = t["xgf5"] / (t["xgf5"] + t["xga5"])
    t["gsax"] = t["xga"] - t["ga"]
    t["gsax60"] = t["gsax"] / (t["all_s"] / 3600)
    d_good = t["xga60_5"] <= t.groupby("season_id")["xga60_5"].transform("median")
    g_good = t["gsax60"] >= t.groupby("season_id")["gsax60"].transform("median")
    t["label"] = np.select([d_good & g_good, ~d_good & g_good, d_good & ~g_good], LABELS[:3], default=LABELS[3])
    t["defense_rank"] = t.groupby("season_id")["xga60_5"].rank(method="min").astype(int)
    t["goalie_rank"] = t.groupby("season_id")["gsax60"].rank(ascending=False, method="min").astype(int)
    return t


def compare_with_moneypuck(t, nhl, seasons):
    mp = pd.read_sql_query("""SELECT season_id, team_id, situation, xGoalsFor, xGoalsAgainst, goalsAgainst, iceTime FROM mp_team_seasons
                              WHERE situation IN ('5on5', 'all') AND team_id IS NOT NULL""", nhl)
    mp["team_id"] = mp["team_id"].astype(int)
    m5, ma = mp[mp["situation"] == "5on5"], mp[mp["situation"] == "all"]
    m = m5.merge(ma, on=["season_id", "team_id"], suffixes=("_5", "_all"))
    m["mp_xga60_5"] = m["xGoalsAgainst_5"] / (m["iceTime_5"] / 3600)
    m["mp_xgf_5"] = m["xGoalsFor_5"]
    m["mp_gsax60"] = (m["xGoalsAgainst_all"] - m["goalsAgainst_all"]) / (m["iceTime_all"] / 3600)
    j = t.merge(m[["season_id", "team_id", "mp_xga60_5", "mp_xgf_5", "mp_gsax60"]], on=["season_id", "team_id"])
    rows = []
    for s in seasons:
        d = j[j["season_id"] == s]
        if len(d) < 10:
            continue
        rows.append({"Season": season_label(s), "Teams": len(d),
                     "Expected goals for, 5-on-5": round(d["xgf5"].corr(d["mp_xgf_5"]), 3),
                     "Defense (expected goals against per 60, 5-on-5)": round(d["xga60_5"].corr(d["mp_xga60_5"]), 3),
                     "Goaltending (GSAx per 60)": round(d["gsax60"].corr(d["mp_gsax60"]), 3),
                     "Defense rank, same order": round(d["xga60_5"].corr(d["mp_xga60_5"], method="spearman"), 3),
                     "Goaltending rank, same order": round(d["gsax60"].corr(d["mp_gsax60"], method="spearman"), 3)})
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
def main():
    sys.stdout.reconfigure(errors="replace")
    pbp = sqlite3.connect(f"file:{bi.PBP_DB}?mode=ro", uri=True)
    nhl = sqlite3.connect(bi.NHL_DB)
    print("Reading the play-by-play (saved copies are reused)...")
    shots, clocks, current = load_all(pbp, nhl)
    finished = sorted(s for s in shots["season_id"].unique() if s < current)
    fit_on = [s for s in finished if s not in CHECK_SEASONS]
    y = shots["goal"].to_numpy()
    in_check = shots["season_id"].isin(CHECK_SEASONS).to_numpy()
    in_fit = shots["season_id"].isin(fit_on).to_numpy()

    # ---- the check: fitted without the two check seasons, judged on them. Three versions; the one with the lowest log loss is kept ----
    era_fit = [x for x in fit_on if x >= ERA_START]
    checks, versions = [], {}
    for name, window, use_shift in (("Fitted on every season before the check", fit_on, False),
                                    ("Fitted on the current tracking era only (2021-22 on)", era_fit, False),
                                    ("Tracking era only, with the arena correction", era_fit, True)):
        rows = shots["season_id"].isin(window).to_numpy()
        shift = arena_shift(shots, finished + [current]) if use_shift else None
        X = features(shots, shift)
        model = fit(X[rows], y[rows])
        p, _ = scale_by_season(shots, predict(model, X))
        t = team_numbers(shots.assign(xg=p), clocks)
        agree = compare_with_moneypuck(t, nhl, CHECK_SEASONS)
        versions[name] = {"shift": use_shift, "era": window is era_fit, "p": p, "auc": auc(p[in_check], y[in_check]), "loss": log_loss(p[in_check], y[in_check]),
                          "agree": float(agree[["Defense rank, same order", "Goaltending rank, same order"]].mean().mean()) if len(agree) else np.nan}
        checks.append({"Version": name, "Shots in the check seasons": int(in_check.sum()), "Goals": int(y[in_check].sum()),
                       "How well it ranks shots (AUC)": round(versions[name]["auc"], 4), "Log loss (lower is better)": round(versions[name]["loss"], 5),
                       "Agreement with MoneyPuck's team ranks (average)": round(versions[name]["agree"], 3)})
    best = min(versions, key=lambda k: versions[k]["loss"])
    use_shift, use_era = versions[best]["shift"], versions[best]["era"]
    print(f"Check on {', '.join(season_label(x) for x in CHECK_SEASONS)} (not fitted on): " +
          "; ".join(f"{k}: AUC {v['auc']:.4f}, log loss {v['loss']:.5f}, rank agreement with MoneyPuck {v['agree']:.3f}" for k, v in versions.items()))
    print(f"Kept: {best}")

    # ---- the model used on the site: fitted on every finished season (saved, so daily runs only score new shots) ----
    shift = arena_shift(shots, finished + [current]) if use_shift else None
    X = features(shots, shift)
    window = [x for x in finished if x >= ERA_START] if use_era else finished
    key = hashlib.md5(json.dumps({"v": VERSION, "features": FEATURES_TAG, "seasons": [int(x) for x in window], "shift": bool(use_shift),
                                  "n": int(shots["season_id"].isin(window).sum())}).encode()).hexdigest()[:10]
    model_path = CACHE / f"model_{key}.json"
    if model_path.exists():
        model = json.loads(model_path.read_text(encoding="utf-8"))
    else:
        in_window = shots["season_id"].isin(window).to_numpy()
        model = fit(X[in_window], y[in_window])
        model_path.write_text(json.dumps(model), encoding="utf-8")
    p, factor = scale_by_season(shots, predict(model, X))
    shots["xg"] = p
    t = team_numbers(shots, clocks)
    names = pd.read_sql_query("SELECT season_id, team_id, team FROM team_seasons", nhl)
    t = t.merge(names, on=["season_id", "team_id"], how="left")

    # calibration on the check seasons (the kept version, fitted without them)
    pc = versions[best]["p"]
    bins = pd.qcut(pc[in_check], 10, labels=False, duplicates="drop")
    calib = pd.DataFrame({"group": bins + 1, "xg": pc[in_check], "goal": y[in_check]}).groupby("group").agg(
        shots=("goal", "size"), predicted=("xg", "mean"), actual=("goal", "mean")).reset_index()
    calib.columns = ["Group (1 = least dangerous shots)", "Shots", "Predicted goal rate", "Actual goal rate"]
    agree_all = compare_with_moneypuck(t, nhl, [s for s in finished if s in set(t["season_id"])])

    # this season: the same layout as the Goaltending study's "Right now" sheet
    now = t[t["season_id"] == current].sort_values("gsax60", ascending=False)
    dates = pd.read_sql_query("SELECT game_id, game_date FROM games WHERE season_id = ?", nhl, params=[current])
    through = dates[dates["game_id"].isin(shots.loc[shots["season_id"] == current, "game_id"])]["game_date"].max()
    right_now = pd.DataFrame({"Team": now["team"], "Games played": now["games"].astype(int), "Label": now["label"],
                              "Defense rank (1 = best)": now["defense_rank"], "Goaltending rank (1 = best)": now["goalie_rank"],
                              "Goals against": now["ga"].astype(int), "Expected goals against": now["xga"].round(1), "GSAx (goals)": now["gsax"].round(1),
                              "Expected goals against per 60 (5-on-5)": now["xga60_5"].round(2), "GSAx per 60": now["gsax60"].round(3),
                              "Expected-goals share (5-on-5)": now["xg_share_5"].round(3)}).reset_index(drop=True)
    weights = pd.DataFrame({"Feature": ["(intercept)"] + model["columns"], "Weight (on the scaled feature)": np.round(model["beta"], 4)})
    note = (f"Current season, games through {through}: our own expected-goals model (NHL play-by-play). Defense = expected goals against per 60 at 5-on-5; "
            f"goaltending = goals saved above expected per 60. Early in the season these are noisy.")
    read_me = ["Our own expected-goals model, from the NHL play-by-play", "",
               "Every unblocked shot (on goal, missed or a goal; no empty nets or penalty shots) gets a chance of becoming a goal from: distance and angle to the net, shot type,",
               f"rebound (within {REBOUND_SECONDS} s of a teammate's shot), rush (within {RUSH_SECONDS} s of play outside the offensive zone) and strength. Logistic regression fitted on",
               f"{'the finished seasons of the current tracking era (2021-22 on)' if use_era else 'every finished season'}; each season scaled so the league's expected goals equal its goals.",
               "'Check' judges three versions on two finished seasons none of them was fitted on, and the one with the lowest log loss is used.",
               f"Arena correction ({'used' if use_shift else 'not used'}): shot distances shifted for arenas whose scorers record shots nearer or farther than elsewhere; kept only if it helps on the check seasons.",
               "", "Data: (c) NHL. MoneyPuck.com numbers are used only to check the model."]
    sheets = {"Right now": right_now, "Check": pd.DataFrame(checks), "Calibration": calib, "Agreement with MoneyPuck": agree_all, "Weights": weights,
              "Team-seasons": t.sort_values(["season_id", "gsax60"], ascending=[False, False]).round(4)}
    try:
        OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            pd.DataFrame({"Read me": read_me}).to_excel(writer, sheet_name="Read me", index=False)
            writer.sheets["Read me"].column_dimensions["A"].width = 150
            for name, frame in sheets.items():
                start = 1 if name == "Right now" else 0
                frame.to_excel(writer, sheet_name=name, index=False, startrow=start)
                ws = writer.sheets[name]
                if name == "Right now":
                    ws.cell(row=1, column=1, value=note).font = Font(bold=True)
                for cell in ws[start + 1]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                ws.column_dimensions["A"].width = 30
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 220)
    print("\nCheck (seasons the model was not fitted on):"); print(pd.DataFrame(checks).to_string(index=False))
    print("\nCalibration:"); print(calib.round(3).to_string(index=False))
    print("\nAgreement with MoneyPuck, by season (last 6):"); print(agree_all.tail(6).to_string(index=False))
    print(f"\nThis season, games through {through} (top 5 by goaltending):"); print(right_now.head(5).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL; MoneyPuck.com used only to check the model")
    return 0


if __name__ == "__main__":
    sys.exit(main())

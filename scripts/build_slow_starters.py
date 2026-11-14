"""
build_slow_starters.py  -  teams that start out of a playoff spot: who recovers, and why?
========================================================================================
WHAT IT DOES (in Excel terms):
    Takes every team that was OUT of a playoff spot at Nov 1, Thanksgiving, Jan 1 or the trade deadline
    (using only games played up to that date) and asks:

      1. HOW OFTEN DO THEY RECOVER?  The plain recovery rate, split by how far behind the team was,
         and "what it takes": the points pace the team needed for the rest of the season.
         (This tests the popular claim that a team out of a spot at Thanksgiving rarely recovers.)
      2. WHAT SEPARATES THE ONES THAT DO?  Six "signs of a coming recovery", the mirror image of the red flags,
         fixed in advance as top / bottom quarter of the league:
            running cold (PDO)       shooting % + save % in the bottom quarter: luck should turn
            strong shot share        Corsi in the top quarter: the team is playing better than its record
            goals say better         ranks higher on goal difference than on points
            strong special teams     power-play % + penalty-kill % in the top quarter
            stays out of the box     among the league's least short-handed teams
            hot lately               last-10-games points % in the top quarter
         Do teams with these signs recover more than equally far behind teams without them?
         And does the number of signs improve a forecast on seasons it has never seen?
      3. WHO ARE THE SLOW STARTERS RIGHT NOW? Today's teams out of a spot, with their signs and
         how often teams that far behind at this point recovered.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py          (first, so the database is up to date)
    python scripts/build_slow_starters.py

WHAT YOU SHOULD SEE:
    A summary in the terminal and a new file:  outputs/slow_starters.xlsx   (close it in Excel first).
    Start with its "Verdict" sheet. The 'Right now' list asks the NHL for one standings page.

DATA CREDIT:
    Data is (c) NHL, used here for personal, non-commercial purposes.
"""

import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

import build_checkpoint_report as cr
import build_checkpoints as bc
import build_contender_fit as cf
import build_popular_stats_check as ps
import build_red_flags as rf

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "slow_starters.xlsx"

CHECKPOINTS = ["nov1", "thanksgiving", "jan1", "deadline"]
PLAIN_CP = {"nov1": "Nov 1", "thanksgiving": "Thanksgiving", "jan1": "Jan 1", "deadline": "Trade deadline"}
HEADLINE = "thanksgiving"
OUT_BANDS = ["Out: just out (within 0.03 below)", "Out: behind (0.03 to 0.08 below)", "Out: far behind (more than 0.08 below)"]
FULL_SEASON = 82

# (key, short name, plain-English meaning)
GREEN = [
    ("g_pdo", "Running cold (PDO)", "Shooting % + save % is in the bottom quarter of the league: bad luck that tends to turn."),
    ("g_corsi", "Strong shot share (Corsi)", "Takes a bigger share of the shot attempts than most teams (top quarter)."),
    ("g_goals", "Goals say better than record", "Ranks at least 25 percentile points higher on goal difference than on points."),
    ("g_st", "Strong special teams", "Power-play % + penalty-kill % in the top quarter of the league."),
    ("g_pen", "Stays out of the box", "Among the league's least short-handed teams (bottom quarter in times short-handed per game)."),
    ("g_form", "Hot lately", "Points % over the last 10 games in the top quarter of the league."),
]
GREEN_KEYS = [g[0] for g in GREEN]
GREEN_NAME = {g[0]: g[1] for g in GREEN}
GREEN_TEXT = {g[0]: g[2] for g in GREEN}

COMPARE_STATS = [  # (column, plain name, higher is better?)
    ("pts_pct", "Points %", True), ("gd", "Goal difference per game", True), ("corsi", "Corsi (shot-attempt share)", True),
    ("sog_share", "Shots-on-goal share", True), ("sh_pct", "Shooting %", True), ("sv_pct", "Save % (raw)", True),
    ("pdo", "PDO", True), ("pp_pct", "Power-play %", True), ("pk_pct", "Penalty-kill %", True),
    ("tsh_pg", "Times short-handed per game (fewer = higher)", False), ("l10_pct", "Last-10-games points %", True),
]

READ_ME = [
    "Slow starters: teams out of a playoff spot, who recovers, and what separates them?",
    "",
    "THE POPULATION",
    "  Teams that were OUT of a playoff spot at the checkpoint (Nov 1, Thanksgiving, Jan 1 or the trade deadline), 15 main",
    "  seasons from 2009-10 on (2019-20 and 2020-21 left out). About 14-15 teams per season. 'Recovered' = made the playoffs.",
    "",
    "HOW FAR BEHIND (the 'cushion')",
    "  Distance below the weakest team that was in a spot, in points % (0.03 is about 5 points over a full season).",
    "  Just out = within 0.03; behind = 0.03 to 0.08; far behind = more than 0.08. A few 'just out' teams are slightly ABOVE",
    "  that line in points % (a division leader with fewer points % can hold a spot ahead of them): shown as a small positive number.",
    "",
    "WHAT IT TAKES",
    "  'Required pace' = the points % the team must earn in its REMAINING games to reach the typical final cut line",
    "  (the average points % of the weakest playoff team, across these seasons). It uses the typical cut line, not that season's",
    "  actual one, so it is what a fan could work out at the time.",
    "",
    "THE SIX SIGNS OF A COMING RECOVERY (cut-offs fixed in advance: top or bottom quarter of the league that season)",
] + [f"  {name}: {text}" for _, name, text in GREEN] + [
    "",
    "THE FAIR COMPARISON",
    "  A sign could just be a stand-in for 'only slightly behind'. So each team with a sign is compared with what teams the same",
    "  distance behind did. 'Expected recoveries' is how many would recover if the sign meant nothing; 'Excess' = actual minus",
    "  expected; 'z' is the excess divided by its normal wobble (about 2 or more is unlikely to be chance, but with 6 signs x 4",
    "  checkpoints a couple of results will look interesting by luck). A sign is called 'most consistent' if it beats the",
    "  expectation at all four checkpoints (z of 2 or more at one), and a 'hint' if it does at three of four (z of 1 or more).",
    "",
    "THE HOLD-OUT TEST",
    "  Each season is scored by a forecast built on the OTHER seasons: distance behind alone, against distance behind plus the",
    "  number of signs, and plus each sign separately. AUC: 1.0 = every team that recovered was ranked above every team that",
    "  did not; 0.5 = coin flip. The +/- is about two standard errors across 15 seasons.",
    "",
    "WHAT THIS CANNOT SAY",
    "  Small samples (about 220 teams and 50-70 recoveries per checkpoint), team stats only (no injuries, trades, schedule, goalie",
    "  changes). A sign is a reason to look closer, not a prediction.",
    "",
    "Data: (c) NHL.",
]


# ----------------------------------------------------------------------------
# Building the tables
# ----------------------------------------------------------------------------
def add_green_flags(table):
    t = table.copy()
    t["l10_pct"] = t["l10_pts"] / (2 * t["l10_gp"])
    p = lambda col: rf.percentile(t, col)                                # noqa: E731
    t["g_pdo"] = p("pdo") <= 0.25
    t["g_corsi"] = p("corsi") >= 0.75
    t["g_goals"] = (p("gd") - p("pts_pct")) >= 0.25
    t["g_st"] = p("z_power_kill") >= 0.75
    t["g_pen"] = p("tsh_pg") <= 0.25
    t["g_form"] = p("l10_pct") >= 0.75
    t["n_green"] = t[GREEN_KEYS].sum(axis=1)
    return t


def typical_cut_line(snaps, seasons_used):
    """The average points % of the weakest team in a playoff spot on the last day of the season."""
    end = snaps[(snaps["checkpoint"] == "season_end") & (snaps["in_playoff_spot"] == 1) &
                snaps["season_id"].isin(seasons_used)]
    return float(end.groupby(["season_id", "conference"])["pts_pct"].min().mean())


def add_required_pace(t, cut, season_games):
    """Points % needed in the remaining games to reach the typical cut line."""
    games = season_games if np.isscalar(season_games) else t["season_id"].map(season_games)
    points = t["pts_pct"] * 2 * t["gp"]
    remaining = games - t["gp"]
    t["required_pace"] = (cut * 2 * games - points) / (2 * remaining)
    return t


def build_tables(tables, snaps, season_games, cut):
    out = {}
    for cp in CHECKPOINTS:
        s = snaps[snaps["checkpoint"] == cp][["season_id", "team_id", "in_playoff_spot", "conference", "cushion",
                                              "band", "l10_gp", "l10_pts"]]
        t = add_green_flags(tables[cp].merge(s, on=["season_id", "team_id"]))
        t["recovered"] = (t["made_playoffs"] == 1).astype(int)
        out[cp] = add_required_pace(t, cut, season_games)
    return out


def out_of_spot(table):
    d = table[table["in_playoff_spot"] == 0].copy()
    d["p_band"] = d["band"].map(d.groupby("band")["recovered"].mean())
    return d


def excess(group):
    obs, exp = int(group["recovered"].sum()), float(group["p_band"].sum())
    var = float((group["p_band"] * (1 - group["p_band"])).sum())
    return obs, exp, (obs - exp) / np.sqrt(var) if var > 0 else np.nan


# ----------------------------------------------------------------------------
# The analyses
# ----------------------------------------------------------------------------
def recovery_rates(flag_tables):
    rows = []
    for cp, t in flag_tables.items():
        d = out_of_spot(t)
        groups = [("ALL teams out of a spot", d)] + [(b, d[d["band"] == b]) for b in OUT_BANDS]
        for name, g in groups:
            rows.append({"Checkpoint": PLAIN_CP[cp], "Group": name, "Teams": len(g), "Recovered": int(g["recovered"].sum()),
                         "% recovered": round(100 * g["recovered"].mean(), 1) if len(g) else None})
    return pd.DataFrame(rows)


def what_it_takes(flag_tables):
    edges = [-np.inf, 0.60, 0.65, 0.70, np.inf]
    labels = ["Under .600", ".600 to .650", ".650 to .700", ".700 or more"]
    rows = []
    for cp, t in flag_tables.items():
        d = out_of_spot(t)
        d["need"] = pd.cut(d["required_pace"], bins=edges, labels=labels, right=False)
        for label in labels:
            g = d[d["need"] == label]
            rows.append({"Checkpoint": PLAIN_CP[cp], "Points % needed the rest of the way": label, "Teams": len(g),
                         "Recovered": int(g["recovered"].sum()),
                         "% recovered": round(100 * g["recovered"].mean(), 1) if len(g) else None,
                         "Their points % so far": round(g["pts_pct"].mean(), 3) if len(g) else None})
    return pd.DataFrame(rows)


def sign_rows(flag_tables):
    rows = []
    for cp, t in flag_tables.items():
        d = out_of_spot(t)
        for key in GREEN_KEYS:
            flagged, clean = d[d[key]], d[~d[key]]
            obs, exp, z = excess(flagged)
            rows.append({"key": key, "Sign": GREEN_NAME[key], "Checkpoint": PLAIN_CP[cp], "Out-of-spot teams": len(d),
                         "Teams with the sign": len(flagged), "Recovered (with sign)": obs,
                         "Recovery rate, with sign": obs / len(flagged) if len(flagged) else np.nan,
                         "Recovery rate, without": clean["recovered"].mean(), "Recovery rate, all": d["recovered"].mean(),
                         "Expected recoveries if sign meant nothing": exp, "Excess recoveries": obs - exp, "z": z})
    return pd.DataFrame(rows)


def stacking_rows(flag_tables):
    rows = []
    for cp, t in flag_tables.items():
        d = out_of_spot(t)
        d["group"] = d["n_green"].clip(upper=3).map({0: "0 signs", 1: "1 sign", 2: "2 signs", 3: "3 or more signs"})
        for name in ["0 signs", "1 sign", "2 signs", "3 or more signs"]:
            g = d[d["group"] == name]
            if g.empty:
                continue
            obs, exp, z = excess(g)
            rows.append({"Checkpoint": PLAIN_CP[cp], "Signs": name, "Teams": len(g), "Recovered": obs,
                         "Recovery rate": obs / len(g), "Expected from distance alone": exp,
                         "Excess recoveries": obs - exp, "z": z})
    return pd.DataFrame(rows)


def holdout_rows(flag_tables):
    models = {"Distance behind alone": [], "Distance behind + number of signs": ["n_green"],
              "Distance behind + each sign separately": GREEN_KEYS}
    rows = []
    for cp, t in flag_tables.items():
        d = out_of_spot(t)
        d["x_cushion"] = d["cushion"] / 0.05
        per_season = {}
        for name, extra in models.items():
            x = d[["x_cushion"] + extra].astype(float)
            scores = pd.Series(index=d.index, dtype=float)
            for season in sorted(d["season_id"].unique()):
                train, test = d["season_id"] != season, d["season_id"] == season
                beta = cf.fit_logistic(x[train].to_numpy(), d.loc[train, "recovered"].to_numpy(dtype=float))
                scores[test] = cf.predict_logistic(beta, x[test].to_numpy())
            aucs = {}
            for season, g in d.groupby("season_id"):
                if 0 < g["recovered"].sum() < len(g):
                    aucs[season] = cf.auc(scores[g.index], g["recovered"] == 1)
            per_season[name] = pd.Series(aucs)
        base = per_season["Distance behind alone"]
        for name, aucs in per_season.items():
            diff = aucs - base
            rows.append({"Checkpoint": PLAIN_CP[cp], "Forecast": name, "AUC": aucs.mean(), "Gain vs distance alone": diff.mean(),
                         "Margin (+/-, 2 std. errors)": 2 * diff.std(ddof=1) / np.sqrt(len(diff)) if name != "Distance behind alone" else 0.0,
                         "Seasons better": int((diff > 0).sum()), "Seasons worse": int((diff < 0).sum())})
    return pd.DataFrame(rows)


def recovered_vs_not(flag_tables):
    """How did teams that recovered rank, against those that did not, on each stat? (avg percentile, 100 = best)"""
    rows = []
    for key, name, higher in COMPARE_STATS:
        row = {"Statistic": name}
        for cp in ("thanksgiving", "jan1"):
            d = out_of_spot(flag_tables[cp])
            pct = rf.percentile(flag_tables[cp], key)
            pct = pct if higher else 1 - pct + 1 / flag_tables[cp].groupby("season_id")[key].transform("count")
            pct = pct.loc[d.index] * 100
            rec, non = pct[d["recovered"] == 1].mean(), pct[d["recovered"] == 0].mean()
            row[f"{PLAIN_CP[cp]}: recovered"] = round(rec, 0)
            row[f"{PLAIN_CP[cp]}: did not"] = round(non, 0)
            row[f"{PLAIN_CP[cp]}: difference"] = round(rec - non, 0)
        rows.append(row)
    return pd.DataFrame(rows)


def who_recovered(flag_tables):
    d = out_of_spot(flag_tables[HEADLINE])
    d = d[d["recovered"] == 1]
    signs = d.apply(lambda r: "; ".join(GREEN_NAME[k] for k in GREEN_KEYS if r[k]) or "None", axis=1)
    out = pd.DataFrame({
        "Season": d["season"], "Team": d["team"], "Games played": d["gp"], "Points % at Thanksgiving": d["pts_pct"].round(3),
        "Points % vs the cut line (negative = below)": d["cushion"].round(3), "Points % needed the rest of the way": d["required_pace"].round(3),
        "Number of signs": d["n_green"].astype(int), "Signs": signs, "How the season ended": d["outcome"]})
    return out.sort_values(["Season", "Points % vs the cut line (negative = below)"], ascending=[True, False]).reset_index(drop=True)


# ----------------------------------------------------------------------------
# Verdict
# ----------------------------------------------------------------------------
def classify(rows):
    """Most consistent: beats the expectation at ALL checkpoints (z of 2+ at one). Hint: at all but one (z of 1+)."""
    positive = int((rows["Excess recoveries"] > 0).sum())
    best_z = rows["z"].max()
    if positive == len(rows) and best_z >= 2:
        return "Most consistent"
    if positive >= len(rows) - 1 and best_z >= 1:
        return "Hint"
    return "No clear effect"


def verdict_lines(rates, takes, sign_df, stack_df, hold_df, who, cut):
    words = {k: classify(sign_df[sign_df["key"] == k]) for k in GREEN_KEYS}
    best = [GREEN_NAME[k] for k, w in words.items() if w == "Most consistent"]
    hints = [GREEN_NAME[k] for k, w in words.items() if w == "Hint"]
    nothing = [GREEN_NAME[k] for k, w in words.items() if w == "No clear effect"]

    tg = rates[rates["Checkpoint"] == "Thanksgiving"].set_index("Group")
    allrow = tg.loc["ALL teams out of a spot"]
    band_text = ", ".join(f"{b.split(' (')[0].replace('Out: ', '')} {tg.loc[b, '% recovered']:.0f}% ({int(tg.loc[b, 'Recovered'])} of {int(tg.loc[b, 'Teams'])})" for b in OUT_BANDS)
    tk = takes[takes["Checkpoint"] == "Thanksgiving"].set_index("Points % needed the rest of the way")
    take_text = "; ".join(f"needing {k.lower()}: {r['% recovered']:.0f}% ({int(r['Recovered'])} of {int(r['Teams'])})" for k, r in tk.iterrows() if r["Teams"] > 0)

    hold = hold_df[hold_df["Forecast"] == "Distance behind + number of signs"]
    gain, margin = hold["Gain vs distance alone"].mean(), hold["Margin (+/-, 2 std. errors)"].mean()
    jan = stack_df[stack_df["Checkpoint"] == "Thanksgiving"].set_index("Signs")
    stack_text = "; ".join(f"{g}: {int(r['Recovered'])} of {int(r['Teams'])} ({r['Recovery rate']:.0%})" for g, r in jan.iterrows())
    finalists = int(who["How the season ended"].isin(["Won Cup", "Lost Cup final"]).sum())
    champs = who[who["How the season ended"] == "Won Cup"]
    champ_text = f", including {', '.join(champs['Team'] + ' (' + champs['Season'] + ')')} who won the Cup" if len(champs) else ""
    lines = [
        ("Question", "Teams out of a playoff spot at Thanksgiving: how often do they recover, what does it take, and is there anything that separates the ones that do?"),
        ("The Thanksgiving claim", f"Of {int(allrow['Teams'])} teams out of a spot at Thanksgiving, {int(allrow['Recovered'])} ({allrow['% recovered']:.0f}%) made the playoffs. "
                                   f"It depends on how far behind: {band_text}. So 'rarely recover' is true for teams well behind and not for teams just out."),
        ("What it takes", f"Pace needed over the rest of the season to reach the typical final cut line (about {cut:.3f} points %): {take_text}."),
        ("Signs of recovery", ("Most consistent (teams with the sign recovered more than equally far behind teams at all four checkpoints, clearly so at one or more): " + ", ".join(best) + ". " if best else "No sign is consistent at all four checkpoints. ") +
                              ("Hints (small samples): " + ", ".join(hints) + ". " if hints else "") +
                              ("No clear effect: " + ", ".join(nothing) + ". " if nothing else "") +
                              "With 24 sign-and-checkpoint tests, a few results this strong can happen by chance: read them as 'worth a look'."),
        ("Stacking (Thanksgiving)", stack_text + "."),
        ("Honest test", f"Adding the number of signs to the distance behind changes the hold-out AUC by {gain:+.3f} on average (margin about +/-{margin:.3f}): "
                        + ("larger than the noise, so the signs carry some information." if gain > margin else "within the noise, so the signs do not clearly improve on simply asking how far behind the team was.")),
        ("Who recovered", f"{len(who)} teams out of a spot at Thanksgiving made it; {finalists} of them reached the Cup final{champ_text}."),
        ("How to use it", "Treat a slow start as a reason to check HOW far behind the team is and what pace it needs, then look at the signs for whether its record undersells it. Early in a season the signs are noisy."),
    ]
    return pd.DataFrame(lines, columns=["", "Answer"]), words


# ----------------------------------------------------------------------------
# Today's slow starters
# ----------------------------------------------------------------------------
def right_now(games_raw, seasons, flag_tables, cut):
    current = int(games_raw["season_id"].max())
    cur = games_raw[games_raw["season_id"] == current]
    fake_main = seasons[seasons["season_id"] == current].assign(include_main="Y", outcome="in progress")
    cur, _ = ps.prepare_games(cur, fake_main)
    now = cf.add_features(cf.summarize_all(cur))
    through = str(cur["game_date"].max())
    median_gp = int(now["gp"].median())

    games_csv = bc.load_games()
    if games_csv is None:
        return None, current, through, None
    rows = bc.fetch_snapshot(current, "now", through, bc.team_abbreviations(games_csv))
    if not rows:
        return None, current, through, None
    snap = bc.mark_playoff_spots(pd.DataFrame(rows))
    snap["pts_pct"] = snap["pts_pct"].astype(float)
    snap = cr.add_cushion_bands(snap)
    t = now.merge(snap[["team_id", "conference", "in_playoff_spot", "cushion", "band", "l10_gp", "l10_pts"]], on="team_id")
    t = t.merge(seasons[seasons["season_id"] == current][["team_id", "team"]], on="team_id")
    t = add_required_pace(add_green_flags(t), cut, FULL_SEASON)
    t = t[t["in_playoff_spot"] == 0]

    cp = min(CHECKPOINTS, key=lambda c: abs(float(flag_tables[c]["gp"].median()) - median_gp))
    hist = out_of_spot(flag_tables[cp]).groupby("band")["recovered"].agg(["sum", "count"])
    t["history"] = t["band"].map(lambda b: f"{int(hist.loc[b, 'sum'])} of {int(hist.loc[b, 'count'])} recovered" if b in hist.index else "")
    t["Signs"] = t.apply(lambda r: "; ".join(GREEN_NAME[k] for k in GREEN_KEYS if r[k]) or "None", axis=1)
    out = pd.DataFrame({
        "Team": t["team"], "Conference": t["conference"], "Games played": t["gp"], "Points %": t["pts_pct"].round(3),
        "Points % vs the cut line (negative = below)": t["cushion"].round(3),
        "How far behind": t["band"].str.replace("Out: ", "", regex=False).str.split(" \\(", regex=True).str[0],
        f"Teams that far behind at {PLAIN_CP[cp]} (history)": t["history"],
        "Points % needed the rest of the way": t["required_pace"].round(3),
        "Number of signs": t["n_green"].astype(int), "Signs": t["Signs"],
        "Corsi %": (t["corsi"] * 100).round(1), "PDO": t["pdo"].round(3),
        "Last 10 (points)": t["l10_pts"].astype(int).astype(str) + " of " + (2 * t["l10_gp"]).astype(int).astype(str)})
    out = out.sort_values("Points % vs the cut line (negative = below)", ascending=False).reset_index(drop=True)
    return out, current, through, cp


# ----------------------------------------------------------------------------
def write_excel(sheets, notes):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 124
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
                values = [c.value for c in column[header_row - 1:300] if c.value is not None
                          and (not isinstance(c.value, str) or len(c.value) < 70)]
                longest = max((len(str(v)) for v in values), default=8)
                ws.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 9), 44)
        for name, col, width in (("Verdict", "B", 130), ("The signs", "B", 100)):
            ws = writer.sheets[name]
            ws.column_dimensions["A"].width = 26 if name == "Verdict" else 30
            ws.column_dimensions[col].width = width
            for row in ws.iter_rows(min_row=2):
                row[0].font = Font(bold=True)
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        if "Right now" in sheets:
            ws = writer.sheets["Right now"]
            ws.column_dimensions["J"].width = 60
            for row in ws.iter_rows(min_row=3):
                row[9].alignment = Alignment(wrap_text=True, vertical="top")


def main():
    sys.stdout.reconfigure(errors="replace")
    if not cf.ps.DB_FILE.exists():
        print("I can't find the database yet. Run this first:")
        print("    python scripts/build_database.py")
        return 1
    connection = sqlite3.connect(cf.ps.DB_FILE)
    games_raw, seasons, checkpoints = cf.load(connection)
    games, _ = ps.prepare_games(games_raw, seasons)
    stat_tables = cf.build_checkpoint_tables(games, seasons, checkpoints)
    snaps = rf.load_snapshots(connection)
    connection.close()

    season_games = games.groupby(["season_id", "team_id"]).size().groupby("season_id").median()
    cut = typical_cut_line(snaps, set(games["season_id"].unique()))
    flag_tables = build_tables(stat_tables, snaps, season_games, cut)
    n_seasons = games["season_id"].nunique()
    n_out = len(out_of_spot(flag_tables[HEADLINE]))
    print(f"Using {n_seasons} main seasons from 2009-10 on; {n_out} out-of-spot team-seasons at Thanksgiving; typical final cut line {cut:.3f} points %.")

    rates = recovery_rates(flag_tables)
    takes = what_it_takes(flag_tables)
    sign_df = sign_rows(flag_tables)
    stack_df = stacking_rows(flag_tables)
    hold_df = holdout_rows(flag_tables)
    who = who_recovered(flag_tables)
    verdict, words = verdict_lines(rates, takes, sign_df, stack_df, hold_df, who, cut)
    now_table, current, through, used_cp = right_now(games_raw, seasons, flag_tables, cut)

    signs_out = sign_df.drop(columns="key").copy()
    for col in ["Recovery rate, with sign", "Recovery rate, without", "Recovery rate, all"]:
        signs_out[col] = signs_out[col].round(3)
    signs_out[["Expected recoveries if sign meant nothing", "Excess recoveries", "z"]] = \
        signs_out[["Expected recoveries if sign meant nothing", "Excess recoveries", "z"]].round(2)
    signs_out.insert(0, "Verdict across the 4 checkpoints", signs_out["Sign"].map({GREEN_NAME[k]: w for k, w in words.items()}))

    team_cols = ["season", "team", "outcome", "n_green"] + GREEN_KEYS + ["cushion", "required_pace", "pts_pct", "gd", "corsi", "pdo"]
    teams = pd.concat([out_of_spot(t)[team_cols].assign(Checkpoint=PLAIN_CP[cp]) for cp, t in flag_tables.items()], ignore_index=True)
    teams = teams.rename(columns={**{k: GREEN_NAME[k] for k in GREEN_KEYS}, "n_green": "Number of signs", "season": "Season", "team": "Team",
                                  "outcome": "How the season ended", "cushion": "Points % vs the cut line (negative = below)",
                                  "required_pace": "Points % needed the rest of the way", "pts_pct": "Points %",
                                  "gd": "Goal diff. per game", "corsi": "Corsi", "pdo": "PDO"})
    teams = teams[["Season", "Checkpoint", "Team", "How the season ended", "Number of signs"] + [GREEN_NAME[k] for k in GREEN_KEYS] +
                  ["Points % vs the cut line (negative = below)", "Points % needed the rest of the way", "Points %", "Goal diff. per game", "Corsi", "PDO"]].round(3)

    sheets = {"Verdict": verdict, "Recovery rates": rates, "What it takes": takes,
              "The signs": pd.DataFrame([(GREEN_NAME[k], GREEN_TEXT[k]) for k in GREEN_KEYS], columns=["Sign", "What it means"]),
              "Sign by sign": signs_out, "Stacking": stack_df.round(3), "Hold-out test": hold_df.round(3),
              "Recovered vs not": recovered_vs_not(flag_tables), "Who recovered": who}
    notes = {}
    if now_table is not None:
        sheets["Right now"] = now_table
        notes["Right now"] = (f"Season {str(current)[:4]}-{str(current)[6:]}: teams out of a playoff spot, games through {through}. "
                              f"History column uses {PLAIN_CP[used_cp]} (closest to now). Early in a season the signs are noisy: a reason to look closer, not a prediction.")
    sheets["Teams"] = teams
    try:
        write_excel(sheets, notes)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 230)
    pd.set_option("display.max_colwidth", 130)
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nRecovery rates:")
    print(rates.to_string(index=False))
    print("\nWhat it takes (Thanksgiving + Jan 1):")
    print(takes[takes["Checkpoint"].isin(["Thanksgiving", "Jan 1"])].to_string(index=False))
    print("\nSign by sign:")
    print(sign_df[["Sign", "Checkpoint", "Teams with the sign", "Recovery rate, with sign", "Recovery rate, without", "Excess recoveries", "z"]].round(2).to_string(index=False))
    print("\nStacking:")
    print(stack_df.round(2).to_string(index=False))
    print("\nHold-out test:")
    print(hold_df.round(3).to_string(index=False))
    if now_table is not None:
        print(f"\nRight now (through {through}):")
        print(now_table.iloc[:, [0, 2, 3, 4, 5, 7, 8, 9]].to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

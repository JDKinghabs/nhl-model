"""
build_red_flags.py  -  warning signs for teams that look safe: do they really miss more often?
=============================================================================================
WHAT IT DOES (in Excel terms):
    Takes every team that was IN a playoff spot at Thanksgiving, Jan 1 or the trade deadline and
    checks six warning signs ("red flags"), using only games played up to that date:

        1. Running hot            shooting % + save % (PDO) in the top quarter of the league
        2. Weak shot share        shot-attempt share (Corsi) in the bottom quarter
        3. Record flatters goals  the team ranks much higher on points than on goal difference
        4. Weak special teams     power-play % + penalty-kill % in the bottom quarter
        5. Penalty trouble        short-handed more often than almost anyone (top quarter)
        6. Cold lately            points % over the last 10 games in the bottom quarter

    The flags and their cut-offs were fixed BEFORE looking at the results (simple quarters, not tuned).
    Then it asks, for each flag and for the number of flags a team has:
      * Did flagged teams miss the playoffs more than other teams that were just as safe on the
        standings (same cushion above the cut line)? That is the fair comparison, because a flag could
        just be a stand-in for "was barely in a spot".
      * Does knowing the flag count improve a forecast of who misses, on seasons the forecast has
        never seen? (the same hold-out test as before)

    It also lists today's in-spot teams with their flags.

HOW TO RUN IT (in the terminal, from the "Model 2.0" folder):
    python scripts/build_database.py          (first, so the database is up to date)
    python scripts/build_red_flags.py

WHAT YOU SHOULD SEE:
    A summary in the terminal and a new file:  outputs/red_flags.xlsx   (close it in Excel first).
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

import build_checkpoint_report as cr      # cushion above the playoff cut line
import build_checkpoints as bc            # standings snapshots and the 'in a playoff spot' rule
import build_contender_fit as cf          # checkpoint stat tables and the small logistic-regression helper
import build_popular_stats_check as ps    # shared stat calculations

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_FILE = PROJECT_FOLDER / "outputs" / "red_flags.xlsx"

CHECKPOINTS = ["thanksgiving", "jan1", "deadline"]
PLAIN_CP = {"thanksgiving": "Thanksgiving", "jan1": "Jan 1", "deadline": "Trade deadline"}
IN_BANDS = ["In: far clear (more than 0.08 above the line)", "In: comfortable (0.03 to 0.08 above)",
            "In: bubble (within 0.03 of the line)"]

# (key, short name, plain-English meaning)
FLAGS = [
    ("f_pdo", "Running hot (PDO)", "Shooting % + save % is in the top quarter of the league: luck that tends to even out."),
    ("f_corsi", "Weak shot share (Corsi)", "Takes a smaller share of the shot attempts than most teams (bottom quarter)."),
    ("f_record", "Record flatters goals", "Ranks at least 25 percentile points higher on points than on goal difference: wins came in close games."),
    ("f_st", "Weak special teams", "Power-play % + penalty-kill % in the bottom quarter of the league."),
    ("f_pen", "Penalty trouble", "Short-handed more often than almost anyone: top quarter in times short-handed per game."),
    ("f_form", "Cold lately", "Points % over the last 10 games in the bottom quarter of the league."),
]
FLAG_KEYS = [f[0] for f in FLAGS]
FLAG_NAME = {f[0]: f[1] for f in FLAGS}
FLAG_TEXT = {f[0]: f[2] for f in FLAGS}

READ_ME = [
    "Red flags: do warning signs make a team that is 'in' more likely to miss the playoffs?",
    "",
    "THE POPULATION",
    "  Teams that were IN a playoff spot at the checkpoint (Thanksgiving, Jan 1 or the trade deadline), 15 main seasons",
    "  from 2009-10 on (2019-20 and 2020-21 left out). About 16 teams per season, so roughly 240 team-seasons per checkpoint.",
    "  'Missed' means the team did not make the playoffs after all.",
    "",
    "THE SIX FLAGS (cut-offs fixed in advance: top or bottom quarter of the league that season)",
] + [f"  {name}: {text}" for _, name, text in FLAGS] + [
    "",
    "THE FAIR COMPARISON",
    "  A flag might only be a stand-in for 'barely in a spot'. So each flagged team is compared with what teams with the same",
    "  cushion above the cut line did (three cushion bands: bubble, comfortable, far clear). 'Expected misses' is how many",
    "  flagged teams would have missed if flags meant nothing. 'Excess' = actual minus expected. 'z' is the excess divided by its",
    "  normal wobble: about 2 or more is unlikely to be chance (but with 6 flags x 3 checkpoints, one or two will look",
    "  interesting by luck).",
    "",
    "THE HOLD-OUT TEST",
    "  Each season is scored by a forecast built on the OTHER seasons. Cushion alone is compared with cushion + number of flags,",
    "  and with cushion + each flag separately. AUC: 1.0 = every team that missed was ranked as riskier than every team that",
    "  made it; 0.5 = coin flip. The +/- is about two standard errors across 15 seasons.",
    "",
    "WHAT THIS CANNOT SAY",
    "  Small samples (about 240 teams and 55-60 misses per checkpoint). The flags use team stats only: no shot quality,",
    "  no injuries, no schedule. A flag is a reason to look closer, not a prediction.",
    "",
    "Data: (c) NHL.",
]


# ----------------------------------------------------------------------------
# Building the flag tables
# ----------------------------------------------------------------------------
def percentile(table, column):
    """Where each team ranks in the league that season: 0 = worst, 1 = best (by this column's size)."""
    return table.groupby("season_id")[column].rank(pct=True)


def add_flags(table):
    """Add the six True/False flags and the flag count. `table` needs the cf stat columns + l10 columns."""
    t = table.copy()
    t["l10_pct"] = t["l10_pts"] / (2 * t["l10_gp"])
    t["f_pdo"] = percentile(t, "pdo") >= 0.75
    t["f_corsi"] = percentile(t, "corsi") <= 0.25
    t["f_record"] = (percentile(t, "pts_pct") - percentile(t, "gd")) >= 0.25
    t["f_st"] = percentile(t, "z_power_kill") <= 0.25
    t["f_pen"] = percentile(t, "tsh_pg") >= 0.75
    t["f_form"] = percentile(t, "l10_pct") <= 0.25
    t["n_flags"] = t[FLAG_KEYS].sum(axis=1)
    return t


def load_snapshots(connection):
    snaps = pd.read_sql_query(
        """SELECT season_id, checkpoint, team_id, conference, pts_pct AS snap_pts_pct, in_playoff_spot, l10_gp, l10_pts
           FROM standings_snapshots""", connection)
    snaps["pts_pct"] = snaps["snap_pts_pct"]
    return cr.add_cushion_bands(snaps)


def build_flag_tables(tables, snaps):
    """For each checkpoint: every team with its flags, cushion and what happened. Returns {checkpoint: table}."""
    out = {}
    for cp in CHECKPOINTS:
        s = snaps[snaps["checkpoint"] == cp][["season_id", "team_id", "in_playoff_spot", "conference", "cushion",
                                              "band", "l10_gp", "l10_pts"]]
        t = tables[cp].merge(s, on=["season_id", "team_id"])
        t = add_flags(t)
        t["missed"] = (t["made_playoffs"] == 0).astype(int)
        out[cp] = t
    return out


# ----------------------------------------------------------------------------
# The tests
# ----------------------------------------------------------------------------
def in_spot(table):
    """Only the teams that were in a playoff spot, with the miss rate of teams as safe as they were."""
    d = table[table["in_playoff_spot"] == 1].copy()
    rate = d.groupby("band")["missed"].mean()
    d["p_band"] = d["band"].map(rate)
    return d


def excess(group):
    """Actual misses vs the misses expected from teams with the same cushion, and how unusual the gap is."""
    obs, exp = int(group["missed"].sum()), float(group["p_band"].sum())
    var = float((group["p_band"] * (1 - group["p_band"])).sum())
    return obs, exp, (obs - exp) / np.sqrt(var) if var > 0 else np.nan


def flag_rows(flag_tables):
    rows = []
    for cp, table in flag_tables.items():
        d = in_spot(table)
        base = d["missed"].mean()
        for key in FLAG_KEYS:
            flagged, clean = d[d[key]], d[~d[key]]
            obs, exp, z = excess(flagged)
            rows.append({
                "key": key, "Flag": FLAG_NAME[key], "Checkpoint": PLAIN_CP[cp], "In-spot teams": len(d),
                "Flagged teams": len(flagged), "Missed (flagged)": obs,
                "Miss rate, flagged": obs / len(flagged), "Miss rate, not flagged": clean["missed"].mean(),
                "Miss rate, all in-spot": base, "Expected misses if flag meant nothing": exp,
                "Excess misses": obs - exp, "z": z})
    return pd.DataFrame(rows)


def stacking_rows(flag_tables):
    rows = []
    for cp, table in flag_tables.items():
        d = in_spot(table)
        d["group"] = d["n_flags"].clip(upper=3).map({0: "0 flags", 1: "1 flag", 2: "2 flags", 3: "3 or more flags"})
        for name in ["0 flags", "1 flag", "2 flags", "3 or more flags"]:
            g = d[d["group"] == name]
            if g.empty:
                continue
            obs, exp, z = excess(g)
            rows.append({"Checkpoint": PLAIN_CP[cp], "Flags": name, "Teams": len(g), "Missed": obs,
                         "Miss rate": obs / len(g), "Expected from cushion alone": exp, "Excess misses": obs - exp, "z": z})
    return pd.DataFrame(rows)


def holdout_rows(flag_tables):
    """Does the flag count (or each flag) improve a forecast of who misses, beyond the cushion? Leave-one-season-out."""
    models = {
        "Cushion alone": [],
        "Cushion + number of flags": ["n_flags"],
        "Cushion + each flag separately": FLAG_KEYS,
    }
    rows = []
    for cp, table in flag_tables.items():
        d = in_spot(table)
        d["x_cushion"] = d["cushion"] / 0.05
        per_season = {}
        for name, extra in models.items():
            cols = ["x_cushion"] + extra
            x = d[cols].astype(float)
            scores = pd.Series(index=d.index, dtype=float)
            for season in sorted(d["season_id"].unique()):
                train, test = d["season_id"] != season, d["season_id"] == season
                beta = cf.fit_logistic(x[train].to_numpy(), d.loc[train, "missed"].to_numpy(dtype=float))
                scores[test] = cf.predict_logistic(beta, x[test].to_numpy())
            aucs = {}
            for season, g in d.groupby("season_id"):
                if 0 < g["missed"].sum() < len(g):
                    aucs[season] = cf.auc(scores[g.index], g["missed"] == 1)
            per_season[name] = pd.Series(aucs)
        base = per_season["Cushion alone"]
        for name, aucs in per_season.items():
            diff = aucs - base
            rows.append({"Checkpoint": PLAIN_CP[cp], "Forecast": name, "AUC": aucs.mean(), "Gain vs cushion alone": diff.mean(),
                         "Margin (+/-, 2 std. errors)": 2 * diff.std(ddof=1) / np.sqrt(len(diff)) if name != "Cushion alone" else 0.0,
                         "Seasons better": int((diff > 0).sum()), "Seasons worse": int((diff < 0).sum())})
    return pd.DataFrame(rows)


def signal_word(rows):
    """Rule of thumb for one flag across the three checkpoints (documented in the Read me)."""
    positive = int((rows["Excess misses"] > 0).sum())
    best_z = rows["z"].max()
    if positive == len(rows) and best_z >= 2:
        return "Most consistent"
    if positive >= 2 and best_z >= 1:
        return "Hint"
    return "No clear effect"


def stacking_text(stack_df):
    """'Teams with 3 or more flags vs teams with 0-2 flags', at each checkpoint, in plain words."""
    parts = []
    for cp in [PLAIN_CP[c] for c in CHECKPOINTS]:
        sub = stack_df[stack_df["Checkpoint"] == cp]
        hi = sub[sub["Flags"] == "3 or more flags"]
        lo = sub[sub["Flags"] != "3 or more flags"]
        if hi.empty:
            continue
        hi_n, hi_m = int(hi["Teams"].iloc[0]), int(hi["Missed"].iloc[0])
        lo_n, lo_m = int(lo["Teams"].sum()), int(lo["Missed"].sum())
        parts.append(f"{cp}: {hi_m} of {hi_n} teams with 3+ flags missed ({hi_m / hi_n:.0%}) against {lo_m} of {lo_n} with 0-2 flags ({lo_m / lo_n:.0%})")
    return ("; ".join(parts) + ". The 3-or-more group is small (8-16 teams) and is one of four groups, so treat the gap as a hint. "
            "The Stacking sheet shows every group.")


def verdict_lines(flag_df, stack_df, hold_df):
    words = {k: signal_word(flag_df[flag_df["key"] == k]) for k in FLAG_KEYS}
    best = [FLAG_NAME[k] for k, w in words.items() if w == "Most consistent"]
    hints = [FLAG_NAME[k] for k, w in words.items() if w == "Hint"]
    nothing = [FLAG_NAME[k] for k, w in words.items() if w == "No clear effect"]
    hold = hold_df[hold_df["Forecast"] == "Cushion + number of flags"]
    gain, margin = hold["Gain vs cushion alone"].mean(), hold["Margin (+/-, 2 std. errors)"].mean()
    base_rates = flag_df.groupby("Checkpoint")["Miss rate, all in-spot"].first()
    lines = [
        ("Question", "Teams in a playoff spot sometimes show warning signs. Do those teams really miss the playoffs more often than equally safe teams without them?"),
        ("Base rate", "Of the teams in a spot at Thanksgiving, " + f"{base_rates['Thanksgiving']:.0%} missed; at Jan 1, {base_rates['Jan 1']:.0%}; at the trade deadline, {base_rates['Trade deadline']:.0%}."),
        ("Flags that matter", ("Most consistent (flagged teams missed more than equally safe teams at all three checkpoints, clearly so at one or more): " + ", ".join(best) + ". " if best else "No flag is consistent at all three checkpoints. ") +
                              ("Hints (small samples): " + ", ".join(hints) + ". " if hints else "") +
                              ("No clear effect: " + ", ".join(nothing) + ". " if nothing else "") +
                              "With 18 flag-and-checkpoint tests, a couple of results this strong can happen by chance, so read them as 'worth a look', not proof."),
        ("Stacking", stacking_text(stack_df)),
        ("Honest test", f"Adding the number of flags to the cushion changes the hold-out AUC by {gain:+.3f} on average (margin about +/-{margin:.3f}): "
                        + ("larger than the noise, so the flags carry some information." if gain > margin else "within the noise, so the flags do not clearly improve on simply asking how close to the line the team was.")),
        ("How to use it", "Treat flags as reasons to look closer, not predictions. They are most useful for explaining WHY a team's spot may be less safe than the standings suggest. "
                          "Early in a season (about 20 games) the flags are noisy."),
    ]
    return pd.DataFrame(lines, columns=["", "Answer"]), words


# ----------------------------------------------------------------------------
# Today's list
# ----------------------------------------------------------------------------
def right_now(games_raw, seasons):
    """In-spot teams right now, with their flags. Asks the NHL for one standings page."""
    current = int(games_raw["season_id"].max())
    cur = games_raw[games_raw["season_id"] == current]
    fake_main = seasons[seasons["season_id"] == current].assign(include_main="Y", outcome="in progress")
    cur, _ = ps.prepare_games(cur, fake_main)
    now = cf.add_features(cf.summarize_all(cur))
    through = str(cur["game_date"].max())

    games_csv = bc.load_games()
    if games_csv is None:
        return None, current, through
    rows = bc.fetch_snapshot(current, "now", through, bc.team_abbreviations(games_csv))
    if not rows:
        return None, current, through
    snap = bc.mark_playoff_spots(pd.DataFrame(rows))
    snap["pts_pct"] = snap["pts_pct"].astype(float)
    snap = cr.add_cushion_bands(snap)
    t = now.merge(snap[["team_id", "conference", "in_playoff_spot", "cushion", "l10_gp", "l10_pts"]], on="team_id")
    t = t.merge(seasons[seasons["season_id"] == current][["team_id", "team"]], on="team_id")
    t = add_flags(t)
    t = t[t["in_playoff_spot"] == 1]
    t["Flags"] = t.apply(lambda r: "; ".join(FLAG_NAME[k] for k in FLAG_KEYS if r[k]) or "None", axis=1)
    out = pd.DataFrame({
        "Team": t["team"], "Conference": t["conference"], "Games played": t["gp"], "Points %": t["pts_pct"].round(3),
        "Cushion above the cut line": t["cushion"].round(3), "Number of flags": t["n_flags"].astype(int), "Flags": t["Flags"],
        "Corsi %": (t["corsi"] * 100).round(1), "PDO": t["pdo"].round(3), "Goal diff. per game": t["gd"].round(2),
        "PP% + PK%": ((t["pp_pct"] + t["pk_pct"]) * 100).round(1),
        "Last 10 (points)": t["l10_pts"].astype(int).astype(str) + " of " + (2 * t["l10_gp"]).astype(int).astype(str)})
    out = out.sort_values(["Number of flags", "Cushion above the cut line"], ascending=[False, True]).reset_index(drop=True)
    return out, current, through


# ----------------------------------------------------------------------------
def write_excel(sheets, notes):
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame({"Read me": READ_ME}).to_excel(writer, sheet_name="Read me", index=False)
        ws = writer.sheets["Read me"]
        ws.column_dimensions["A"].width = 122
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
        ws = writer.sheets["Verdict"]
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 130
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(bold=True)
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        if "Right now" in sheets:
            ws = writer.sheets["Right now"]
            ws.column_dimensions["G"].width = 60
            for row in ws.iter_rows(min_row=3):
                row[6].alignment = Alignment(wrap_text=True, vertical="top")
        ws = writer.sheets["The flags"]
        ws.column_dimensions["B"].width = 100
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")


def main():
    sys.stdout.reconfigure(errors="replace")
    if not cf.ps.DB_FILE.exists():
        print("I can't find the database yet. Run this first:")
        print("    python scripts/build_database.py")
        return 1
    connection = sqlite3.connect(cf.ps.DB_FILE)
    games_raw, seasons, checkpoints = cf.load(connection)
    games, _ = ps.prepare_games(games_raw, seasons)
    tables = cf.build_checkpoint_tables(games, seasons, checkpoints)
    snaps = load_snapshots(connection)
    connection.close()

    flag_tables = build_flag_tables(tables, snaps)
    n_seasons = games["season_id"].nunique()
    n_in = len(in_spot(flag_tables["jan1"]))
    print(f"Using {n_seasons} main seasons from 2009-10 on; {n_in} in-spot team-seasons at Jan 1.")

    flag_df = flag_rows(flag_tables)
    stack_df = stacking_rows(flag_tables)
    hold_df = holdout_rows(flag_tables)
    verdict, words = verdict_lines(flag_df, stack_df, hold_df)

    now_table, current, through = right_now(games_raw, seasons)

    flags_out = flag_df.drop(columns="key").copy()
    pct_cols = ["Miss rate, flagged", "Miss rate, not flagged", "Miss rate, all in-spot"]
    flags_out[pct_cols] = flags_out[pct_cols].round(3)
    flags_out[["Expected misses if flag meant nothing", "Excess misses", "z"]] = \
        flags_out[["Expected misses if flag meant nothing", "Excess misses", "z"]].round(2)
    flags_out.insert(0, "Verdict across the 3 checkpoints", flags_out["Flag"].map({FLAG_NAME[k]: w for k, w in words.items()}))

    team_cols = ["season", "team", "outcome", "n_flags"] + FLAG_KEYS + ["cushion", "pts_pct", "gd", "corsi", "pdo"]
    teams = pd.concat([in_spot(t)[team_cols].assign(Checkpoint=PLAIN_CP[cp]) for cp, t in flag_tables.items()], ignore_index=True)
    teams = teams.rename(columns={**{k: FLAG_NAME[k] for k in FLAG_KEYS}, "n_flags": "Number of flags", "season": "Season",
                                  "team": "Team", "outcome": "How the season ended", "cushion": "Cushion above the cut line",
                                  "pts_pct": "Points %", "gd": "Goal diff. per game", "corsi": "Corsi", "pdo": "PDO"})
    teams = teams[["Season", "Checkpoint", "Team", "How the season ended", "Number of flags"] + [FLAG_NAME[k] for k in FLAG_KEYS] +
                  ["Cushion above the cut line", "Points %", "Goal diff. per game", "Corsi", "PDO"]].round(3)

    sheets = {"Verdict": verdict, "The flags": pd.DataFrame([(FLAG_NAME[k], FLAG_TEXT[k]) for k in FLAG_KEYS], columns=["Flag", "What it means"]),
              "Flag by flag": flags_out, "Stacking": stack_df.round(3), "Hold-out test": hold_df.round(3)}
    notes = {}
    if now_table is not None:
        sheets["Right now"] = now_table
        notes["Right now"] = (f"Season {str(current)[:4]}-{str(current)[6:]}: teams in a playoff spot, games through {through}. "
                              f"Flags use only games to date; early in the season they are noisy. A reason to look closer, not a prediction.")
    sheets["Teams"] = teams
    try:
        write_excel(sheets, notes)
    except PermissionError:
        print(f"Could not save {OUTPUT_FILE.name}. Is it open in Excel? Close it and run again.")
        return 1

    pd.set_option("display.width", 220)
    pd.set_option("display.max_colwidth", 120)
    print("\nVERDICT")
    for _, r in verdict.iterrows():
        print(f"  [{r.iloc[0]}] {r.iloc[1]}")
    print("\nFlag by flag (in-spot teams): miss rate flagged vs not, excess misses over cushion-only expectation, z")
    show = flag_df[["Flag", "Checkpoint", "Flagged teams", "Miss rate, flagged", "Miss rate, not flagged", "Excess misses", "z"]].round(2)
    print(show.to_string(index=False))
    print("\nStacking:")
    print(stack_df.round(2).to_string(index=False))
    print("\nHold-out test:")
    print(hold_df.round(3).to_string(index=False))
    if now_table is not None:
        print(f"\nRight now (through {through}), most flags first:")
        print(now_table[["Team", "Games played", "Points %", "Cushion above the cut line", "Number of flags", "Flags"]].head(8).to_string(index=False))
    print(f"\nSaved: {OUTPUT_FILE}")
    print("Data: (c) NHL")
    return 0


if __name__ == "__main__":
    sys.exit(main())

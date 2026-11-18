"""
build_site.py  -  turns your CSV files into the website pages
==============================================================
WHAT IT DOES (in Excel terms):
    Think of a mail merge. Each web page is a template with blanks. This script reads
    data/standings.csv and data/schedule_today.csv and fills the blanks, then writes the
    finished pages into the  docs/  folder. GitHub Pages publishes whatever is in  docs/.

HOW TO RUN IT (after pull_schedule_standings.py):
    python scripts/build_site.py

    The Contender fit page also reads outputs/contender_fit.xlsx (made by build_contender_fit.py).

WHAT YOU SHOULD SEE:
    "Wrote docs/index.html", "Wrote docs/about.html" and "Wrote docs/contender-fit.html".
    Then double-click docs/index.html to preview the site in your browser.

The look of the site lives in docs/style.css and docs/site.js (those are edited by hand).
"""

import hashlib
import html
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

try:
    from zoneinfo import ZoneInfo
    EASTERN = ZoneInfo("America/New_York")
except Exception:
    EASTERN = timezone(timedelta(hours=-5))

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
DATA_FOLDER = PROJECT_FOLDER / "data"
SITE_FOLDER = PROJECT_FOLDER / "docs"

SITE_NAME = "NHL Model"
TAGLINE = ("Independent hockey analytics: who is really good, who is likely to miss the playoffs, "
           "and how teams change over a season.")


def esc(value):
    """Make text safe to put in a web page (turns & and < into harmless codes)."""
    return html.escape(str(value), quote=True)


# ----------------------------------------------------------------------------
# The page "frame": the header, menu and footer that every page shares.
# ----------------------------------------------------------------------------
def page(title, active, body, built_text, moneypuck=False):
    def link(href, label, key):
        current = ' aria-current="page"' if key == active else ""
        return f'<a href="{href}"{current}>{label}</a>'

    credit = ('\n    <p>Shot-quality data: <a href="https://moneypuck.com/data.htm">MoneyPuck.com</a>, used for non-commercial purposes with credit.</p>'
              if moneypuck else "")
    # A short fingerprint of the style and script files, added to their links. When either file changes the link changes,
    # so a browser can never mix a new page with an old cached copy of its script or style.
    version = hashlib.md5(b"".join((SITE_FOLDER / name).read_bytes() for name in ("style.css", "site.js"))).hexdigest()[:8]

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#0b1f3a">
<meta name="color-scheme" content="light">
<title>{esc(title)}</title>
<meta name="description" content="{esc(TAGLINE)}">
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>🏒</text></svg>">
<link rel="stylesheet" href="style.css?v={version}">
</head>
<body>
<header class="site-header">
  <div class="wrap">
    <a class="brand" href="index.html">{esc(SITE_NAME)}</a>
    <nav>{link("index.html", "Home", "home")}{link("contender-fit.html", "Contender fit", "contender")}{link("studies.html", "Studies", "studies")}{link("about.html", "About", "about")}</nav>
  </div>
</header>
<main class="wrap">
{body}
</main>
<footer class="site-footer">
  <div class="wrap">
    <p>Data: &copy; NHL, used for personal, non-commercial purposes. This is an independent hobby project, not affiliated with or endorsed by the NHL or any club.</p>{credit}
    <p>For information and entertainment only. Nothing here is betting or financial advice.</p>
    <p>{esc(built_text)}</p>
  </div>
</footer>
<script src="site.js?v={version}"></script>
</body>
</html>
"""


# ----------------------------------------------------------------------------
# Small formatting helpers
# ----------------------------------------------------------------------------
def fmt_pct(value):
    """0.8 -> '.800' (hockey style)."""
    text = f"{float(value):.3f}"
    return text[1:] if text.startswith("0") else text


def fmt_gd(value):
    """Goal difference with a sign: 36 -> '+36', 0 -> '0'."""
    value = int(value)
    return "0" if value == 0 else f"{value:+d}"


def nice_date(date_text):
    d = datetime.strptime(date_text, "%Y-%m-%d")
    return f"{d:%A, %B} {d.day}"


# ----------------------------------------------------------------------------
# Tonight's games
# ----------------------------------------------------------------------------
def games_section(schedule, standings):
    if schedule.empty:
        return "<h2>Tonight</h2><p class='note'>No games found on this week's schedule.</p>"

    records = {row["team"]: f"{row['w']}-{row['l']}-{row['otl']}" for _, row in standings.iterrows()}
    first = schedule.iloc[0]
    heading = "Tonight" if bool(first["is_today"]) else "Next games"
    sub = nice_date(first["game_date"])

    cards = []
    for _, g in schedule.iterrows():
        live = g["status"] == "In progress"
        if pd.isna(g["away_score"]):
            status = g["status"]
        else:
            status = f"{g['status']} &middot; {esc(g['away'])} {int(g['away_score'])}, {esc(g['home'])} {int(g['home_score'])}"
        cards.append(f"""
    <li class="game">
      <div class="matchup">
        <span>{esc(g['away'])}</span> <span class="rec">{esc(records.get(g['away'], ''))}</span>
        <span class="at">at</span>
        <span>{esc(g['home'])}</span> <span class="rec">{esc(records.get(g['home'], ''))}</span>
      </div>
      <div class="meta"><time datetime="{esc(g['start_utc'])}" data-local>{esc(g['start_et'])}</time> &middot; {esc(g['venue'])}</div>
      <div class="status{' live' if live else ''}">{status}</div>
    </li>""")

    return f"""
<h2>{heading} <span class="note">&middot; {esc(sub)}</span></h2>
<ul class="games">{''.join(cards)}
</ul>"""


# ----------------------------------------------------------------------------
# The standings table
# ----------------------------------------------------------------------------
# (heading, column in the CSV, "num" or "text", tooltip, which way to sort on first click)
COLUMNS = [
    ("#", "league_rank", "num", "Rank in the league (by points)", "ascending"),
    ("Team", "team", "text", "Team", "ascending"),
    ("GP", "gp", "num", "Games played", None),
    ("W", "w", "num", "Wins", None),
    ("L", "l", "num", "Regulation losses", None),
    ("OTL", "otl", "num", "Overtime / shootout losses", None),
    ("PTS", "pts", "num", "Points", None),
    ("P%", "pts_pct", "num", "Points percentage: points won out of points available", None),
    ("ROW", "row", "num", "Regulation + overtime wins (first tiebreaker after points)", None),
    ("GF", "gf", "num", "Goals for", None),
    ("GA", "ga", "num", "Goals against", "ascending"),
    ("GD", "gd", "num", "Goal difference", None),
    ("L10", "l10", "num", "Record in the last 10 games (W-L-OTL)", None),
    ("STRK", "streak", "text", "Current streak (W = winning, L = losing, OTL = overtime losses)", None),
]


def standings_section(standings):
    head = []
    for label, key, kind, tip, first in COLUMNS:
        left = " left" if key in ("team", "streak") else ""
        first_attr = f' data-first="{first}"' if first else ""
        head.append(f'<th class="{left.strip()}" data-sort="{kind}"{first_attr} title="{esc(tip)}" scope="col">'
                    f'<button type="button">{esc(label)}</button></th>')

    rows = []
    for _, t in standings.iterrows():
        l10_wins = int(str(t["l10"]).split("-")[0])
        cells = [
            f'<td data-v="{int(t["league_rank"])}">{int(t["league_rank"])}</td>',
            f'<td class="team left" data-v="{esc(t["team_name"])}">{esc(t["team"])}<span class="full">{esc(t["team_name"])}</span></td>',
            f'<td data-v="{int(t["gp"])}">{int(t["gp"])}</td>',
            f'<td data-v="{int(t["w"])}">{int(t["w"])}</td>',
            f'<td data-v="{int(t["l"])}">{int(t["l"])}</td>',
            f'<td data-v="{int(t["otl"])}">{int(t["otl"])}</td>',
            f'<td class="pts" data-v="{int(t["pts"])}">{int(t["pts"])}</td>',
            f'<td data-v="{float(t["pts_pct"])}">{fmt_pct(t["pts_pct"])}</td>',
            f'<td data-v="{int(t["row"])}">{int(t["row"])}</td>',
            f'<td data-v="{int(t["gf"])}">{int(t["gf"])}</td>',
            f'<td data-v="{int(t["ga"])}">{int(t["ga"])}</td>',
            f'<td class="{"pos" if t["gd"] > 0 else "neg" if t["gd"] < 0 else ""}" data-v="{int(t["gd"])}">{fmt_gd(t["gd"])}</td>',
            f'<td data-v="{l10_wins}">{esc(t["l10"])}</td>',
            f'<td class="left" data-v="">{esc(t["streak"])}</td>',
        ]
        rows.append(f'<tr data-conf="{esc(t["conference"])}">{"".join(cells)}</tr>')

    as_of = nice_date(standings["as_of"].iloc[0])
    return f"""
<h2>Standings <span class="note">&middot; through {esc(as_of)}</span></h2>
<div class="chips" data-filter-group="standings" role="group" aria-label="Filter standings">
  <button type="button" data-filter="all" aria-pressed="true">League</button>
  <button type="button" data-filter="Eastern" aria-pressed="false">Eastern</button>
  <button type="button" data-filter="Western" aria-pressed="false">Western</button>
</div>
<div class="table-wrap">
<table class="sortable ranked" id="standings">
<caption class="note" style="text-align:left;padding:0 0 8px;caption-side:bottom">Click a column heading to sort. Hover a heading for what it means.</caption>
<thead><tr>{''.join(head)}</tr></thead>
<tbody>
{chr(10).join(rows)}
</tbody>
</table>
</div>"""


# ----------------------------------------------------------------------------
# The two pages
# ----------------------------------------------------------------------------
def home_page(schedule, standings, built_text):
    body = f"""
<h1>{esc(SITE_NAME)}</h1>
{games_section(schedule, standings)}
{standings_section(standings)}
"""
    return page(SITE_NAME, "home", body, built_text)


def about_page(built_text):
    body = """
<h1>About this project</h1>

<div class="prose">
<h2>What this is</h2>
<p>Most hockey tables tell you what has happened. This project tries to tell you what is really going on underneath: which teams are genuinely good, which are getting lucky, and which are headed for trouble. It looks at things like the quality of shots a team creates and allows, how its goalie performs against the shots he faces, and how special teams are doing, then compares today's teams with teams from past seasons.</p>
<p>Every view is a hypothesis to test against history, not a fact. When the data disagrees with a popular hockey belief, this site will say so.</p>
</div>

<h2>What is planned</h2>
<ul class="roadmap">
  <li><strong><a href="contender-fit.html">Contender fit</a></strong><span class="tag live">Live</span><br><span class="note">How closely each team's profile matches teams that reached the conference finals in past seasons, and why. Version 1 uses NHL team stats only; shot-quality data comes next.</span></li>
  <li><strong><a href="red-flags.html">Red flags</a></strong><span class="tag live">Live</span><br><span class="note">A simple checklist of warning signs, and how often teams with those signs have missed the playoffs.</span></li>
  <li><strong><a href="slow-starters.html">Slow starters</a></strong><span class="tag live">Live</span><br><span class="note">Which teams that began badly tend to recover, and what separated them from the ones that did not.</span></li>
  <li><strong><a href="goaltending.html">Goaltending in context</a></strong><span class="tag live">Live</span><br><span class="note">Whether a goalie is bailing out a weak defense or a strong defense is being let down, and whether a hot playoff goalie can be foreseen.</span></li>
  <li><strong><a href="breakouts.html">Breakouts</a></strong><span class="tag live">Live</span><br><span class="note">Whether teams that get out of their own end cleanly win more, measured with a breakout efficiency rating (BER) and other stats read from the play-by-play.</span></li>
  <li><strong>Zone time</strong><span class="tag">Planned</span><br><span class="note">How much of the game a team spends in the offensive zone.</span></li>
  <li><strong>Team profiles over time</strong><span class="tag">Planned</span><br><span class="note">How each team's strengths and weaknesses change as the season goes on.</span></li>
</ul>

<h2>Reading the standings</h2>
<dl class="glossary">
  <dt>GP</dt><dd>Games played.</dd>
  <dt>W, L, OTL</dt><dd>Wins, regulation losses, and overtime or shootout losses.</dd>
  <dt>PTS</dt><dd>Points: 2 for a win, 1 for an overtime or shootout loss.</dd>
  <dt>P%</dt><dd>Points percentage: points won out of points available. Fair when teams have played different numbers of games.</dd>
  <dt>ROW</dt><dd>Regulation plus overtime wins (shootout wins do not count). The first tiebreaker after points.</dd>
  <dt>GF, GA, GD</dt><dd>Goals for, goals against, and the difference.</dd>
  <dt>L10</dt><dd>Record over the last 10 games.</dd>
  <dt>STRK</dt><dd>Current streak: W for wins, L for losses, OT for overtime losses.</dd>
</dl>

<h2>Terms used on the research pages</h2>
<dl class="glossary">
  <dt>Points %</dt><dd>Points won out of points available.</dd>
  <dt>Corsi</dt><dd>Share of all shot attempts (shots on goal, misses and blocks) a team takes at 5-on-5. A measure of who controls play.</dd>
  <dt>PDO</dt><dd>Shooting % plus save %. The league average is about 1.000, so a high PDO is usually good luck that tends to even out.</dd>
  <dt>Expected goals</dt><dd>Each shot gets a probability of being a goal from where it was taken, the shot type and what happened before it. Adding them up measures shot quality.</dd>
  <dt>GSAx</dt><dd>Goals saved above expected: the expected goals a goalie faced minus the goals he allowed.</dd>
  <dt>High-danger shot</dt><dd>A shot from the most dangerous area of the ice (the slot), as defined by MoneyPuck.</dd>
  <dt>Cut line, cushion</dt><dd>The cut line is the weakest team currently in a playoff spot. The cushion is how far above or below it a team is, in points %.</dd>
  <dt>Percentile</dt><dd>Where a team ranks that season: 100 is the best in the league, 50 is average.</dd>
  <dt>AUC</dt><dd>How well a ranking puts the teams that did something (such as reaching the conference finals) above the ones that did not. 1.0 is perfect, 0.5 is a coin flip.</dd>
  <dt>Correlation</dt><dd>A number from -1 to +1 for how closely two things move together. About 0.3 is a modest link and 0.5 a strong one; it is usually worked out on ranks, so each season is compared fairly.</dd>
  <dt>R-squared</dt><dd>The share of the differences between teams that a forecast explains. 0.25 means a quarter.</dd>
  <dt>Lift</dt><dd>How much more often something happens to teams that meet a rule than to a random team.</dd>
  <dt>Hold-out test</dt><dd>A rule is built on all seasons but one and tried on the one it has never seen, then repeated for each season.</dd>
  <dt>Blind test</dt><dd>A few seasons are kept out of all the testing. The final recipe is frozen, then scored on them once.</dd>
  <dt>Breakout attempt</dt><dd>A team wins the puck in its own zone at 5-on-5: a faceoff win, a takeaway, a blocked shot, or the opponent missing the net.</dd>
  <dt>Controlled exit</dt><dd>A breakout where the team is next seen beyond its blue line still holding the puck.</dd>
  <dt>Safe clear</dt><dd>A breakout where the puck gets out but the opponent has it, and it does not come straight back.</dd>
  <dt>BER</dt><dd>Breakout efficiency rating: 100 &times; (controlled exits + half the safe clears) &divide; breakout attempts. Read from play-by-play events, so it is an approximation.</dd>
  <dt>BEI</dt><dd>Breakout efficiency index: own-zone giveaways per hour of own-zone time (45%), share of time in the own zone (30%) and transition shot attempts per 60 (25%), each as a league percentile. 0 to 100.</dd>
  <dt>Pinned-in spell</dt><dd>The opponent gets two or more shot attempts in a row, each within 30 seconds of the last, without play leaving the zone or stopping.</dd>
  <dt>Battle</dt><dd>A team's number minus its opponents' number in the same games: for example the pinned-in spells it suffers against the ones it inflicts.</dd>
</dl>

<div class="prose">
<h2>Data and credits</h2>
<p>Schedule and standings come from the NHL's public web data, &copy; NHL, used here for personal, non-commercial purposes. That data is unofficial and can change without notice. Shot-quality data (expected goals, shots by danger, goals saved above expected) comes from <a href="https://moneypuck.com/data.htm">MoneyPuck.com</a>, used for non-commercial purposes with credit, and is credited on every page that uses it.</p>
<p>No NHL logos are used. This site has no ads and does not sell anything.</p>

<h2>Disclaimer</h2>
<p>This is an independent hobby project. It is not affiliated with or endorsed by the NHL, any NHL club, or any data provider. It is for information and entertainment only and is not betting or financial advice.</p>
</div>
"""
    return page("About - " + SITE_NAME, "about", body, built_text)


# ----------------------------------------------------------------------------
# The Contender fit page (reads the workbook written by build_contender_fit.py)
# ----------------------------------------------------------------------------
CONTENDER_FILE = PROJECT_FOLDER / "outputs" / "contender_fit.xlsx"
REPO_URL = "https://github.com/JDKinghabs/nhl-model"

VERDICT_LABELS = {
    "Short answer": "Short answer", "Season by season": "Season by season", "When it helps": "When it helps",
    "What each idea adds": "What each idea adds", "How predictable is it at all?": "How predictable is it?",
    "Single stats": "Single stats on their own", "Not tested yet": "Not tested yet",
}
PROFILE_COLUMNS = [  # (column in the workbook, heading on the page)
    ("Points %", "Points %"), ("Goal difference per game", "Goal diff."), ("Corsi (shot-attempt share)", "Corsi"),
    ("Shots-on-goal share", "Shots share"), ("Shooting %", "Shooting %"), ("Save % (raw)", "Save %"),
    ("PDO", "PDO"), ("Hits per game", "Hits"), ("Power-play %", "PP %"), ("Penalty-kill %", "PK %"),
]


def read_contender(path):
    """Read the sheets this page needs from the workbook."""
    note = str(pd.read_excel(path, sheet_name="Right now", header=None, nrows=1).iloc[0, 0])
    score = pd.read_excel(path, sheet_name="Scorecard")
    data = {
        "note": note,
        "right": pd.read_excel(path, sheet_name="Right now", header=1),
        "verdict": pd.read_excel(path, sheet_name="Verdict"),
        "score": score[score["Checkpoint"] == "ALL FIVE (average)"],
        "profile": pd.read_excel(path, sheet_name="Profile"),
        "rules": pd.read_excel(path, sheet_name="Rule test"),
        "factors": None,
    }
    try:                                    # today's rank on each factor (an older workbook may not have this sheet)
        data["factors"] = pd.read_excel(path, sheet_name="Factors now", header=1)
    except ValueError:
        print("  (the contender workbook has no 'Factors now' sheet: the rule table will not be clickable. "
              "Rerun python scripts/build_contender_fit.py)")
    return data


def right_now_section(right, note, standings):
    conference = {row["team_name"]: row["conference"] for _, row in standings.iterrows()}
    match = re.search(r"games through (\d{4}-\d{2}-\d{2})", note)
    through = nice_date(match.group(1)) if match else "the latest games"

    head = [
        ("#", "num", "ascending", "Fit rank: where the team ranks on contender fit"),
        ("Team", "text", "ascending", "Team, with the main reasons for its rating"),
        ("Fit", "num", None, "The model's estimate that the team ends up one of the four conference finalists"),
        ("P%", "num", None, "Points percentage"),
        ("P% rank", "num", "ascending", "Rank in the league by points percentage, for comparison"),
        ("GP", "num", None, "Games played"),
    ]
    th = "".join(
        f'<th class="{"left" if label == "Team" else ""}" data-sort="{kind}"'
        f'{f" data-first={chr(34)}{first}{chr(34)}" if first else ""} title="{esc(tip)}" scope="col">'
        f'<button type="button">{esc(label)}</button></th>' for label, kind, first, tip in head)

    rows = []
    for _, t in right.iterrows():
        rows.append(
            f'<tr data-conf="{esc(conference.get(t["Team"], ""))}">'
            f'<td data-v="{int(t["Rank"])}">{int(t["Rank"])}</td>'
            f'<td class="who left" data-v="{esc(t["Team"])}"><strong>{esc(t["Team"])}</strong>'
            f'<span class="why">{esc(t["Why"])}</span></td>'
            f'<td class="pts" data-v="{float(t["Final-four chance (model)"])}">{float(t["Final-four chance (model)"]):.0f}%</td>'
            f'<td data-v="{float(t["Points %"])}">{fmt_pct(t["Points %"])}</td>'
            f'<td data-v="{int(t["Rank by points %"])}">{int(t["Rank by points %"])}</td>'
            f'<td data-v="{int(t["Games played"])}">{int(t["Games played"])}</td>'
            f'</tr>')
    return f"""
<h2>Right now <span class="note">&middot; games through {esc(through)}</span></h2>
<p class="note">Ranked by the model's estimate that each team ends up one of the four conference finalists. Because there are four
spots, the 32 percentages add up to 400%. Click a heading to sort; compare the Fit rank with the points rank to see where
the model disagrees with the standings.</p>
<div class="chips" data-filter-group="fit-table" role="group" aria-label="Filter teams">
  <button type="button" data-filter="all" aria-pressed="true">League</button>
  <button type="button" data-filter="Eastern" aria-pressed="false">Eastern</button>
  <button type="button" data-filter="Western" aria-pressed="false">Western</button>
</div>
<div class="table-wrap">
<table class="sortable ranked" id="fit-table">
<thead><tr>{th}</tr></thead>
<tbody>
{chr(10).join(rows)}
</tbody>
</table>
</div>"""


def verdict_section(verdict, score):
    items = []
    for _, row in verdict.iterrows():
        key = str(row.iloc[0])
        if key == "Question":
            continue
        items.append(f"<dt>{esc(VERDICT_LABELS.get(key, key))}</dt><dd>{esc(row.iloc[1])}</dd>")

    body = []
    for _, s in score.iterrows():
        body.append(
            f"<tr><td class='who left'>{esc(s['Method'])}</td><td>{s['AUC']:.3f}</td>"
            f"<td>{s['Avg rank of final-four teams']:.1f}</td><td>{s['Final-four teams in top 8']:.0%}</td>"
            f"<td class='{'pos' if s['Gain vs points % (AUC)'] > 0.005 else 'neg' if s['Gain vs points % (AUC)'] < -0.005 else ''}'>"
            f"{s['Gain vs points % (AUC)']:+.3f}</td></tr>")
    return f"""
<h2>Does it work?</h2>
<p>The honest test: for each past season the score is built from the <em>other</em> seasons, then used to rank that
season's teams, and we look at where the four real conference finalists landed. The benchmark is simply ranking by points %.</p>
<dl class="verdict">{''.join(items)}</dl>
<div class="table-wrap">
<table>
<thead><tr><th class="left" scope="col">Ranking method</th><th scope="col" title="1.0 = every finalist ranked above every non-finalist; 0.5 = a coin flip">AUC</th>
<th scope="col" title="Lower is better; a random list averages about 16">Avg rank of the 4 finalists</th>
<th scope="col">Finalists in top 8</th><th scope="col" title="Change in AUC against ranking by points %">Gain vs points %</th></tr></thead>
<tbody>{''.join(body)}</tbody>
</table>
</div>
<p class="note">Averages over five checkpoints (Nov 1, Thanksgiving, Jan 1, trade deadline, end of season) and 15 seasons from 2009-10,
leaving out the 2019-20 and 2020-21 seasons. AUC: 1.0 is perfect, 0.5 is a coin flip. Only games up to each date are used.</p>"""


def profile_section(profile):
    head = "".join(f'<th scope="col">{esc(label)}</th>' for _, label in PROFILE_COLUMNS)
    rows = []
    for _, r in profile.iterrows():
        cells = "".join(
            f'<td class="heat" style="--p:{float(r[col]):.0f}">{float(r[col]):.0f}</td>' for col, _ in PROFILE_COLUMNS)
        bold = " style='font-weight:700'" if r["How the season ended"] == "ALL teams" else ""
        rows.append(f"<tr><td class='left'{bold}>{esc(r['How the season ended'])}</td><td>{int(r['Team-seasons'])}</td>{cells}</tr>")
    return f"""
<h2>What champions looked like</h2>
<p>How each group of teams ranked in the league at the end of the regular season, on average (100 = best in the league that
season, 50 = average). Darker means higher. 15 seasons, 2009-10 on.</p>
<div class="table-wrap">
<table>
<thead><tr><th class="left" scope="col">How the season ended</th><th scope="col">Teams</th>{head}</tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
</div>
<p class="note">Champions and finalists were strong on results and on shot-attempt share (Corsi) and ordinary on hits.
PDO is shooting % plus save %, a common luck meter: the champions ran a little hot.</p>"""


def rules_section_plain(rules):
    """The rule table without the click-to-open team lists (used when today's factor ranks are not in the workbook)."""
    rows = []
    for _, r in rules.iterrows():
        rows.append(
            f"<tr><td class='who left'>{esc(r['Rule'])}</td><td>{int(r['Teams meeting the rule'])}</td>"
            f"<td>{int(r['Champions among them'])}</td><td>{r['Share of champions that met it']:.0%}</td>"
            f"<td>{r['Champion rate for such teams']:.0%}</td><td>{r['Lift over a random team']:.1f}x</td></tr>")
    return f"""
<h2>Is "top 5 in most of the factors" a rule for finding the champion?</h2>
<p>A popular idea: champions are usually near the top of the league in several key areas, so teams that match should be
favorites. The four areas here are record, goal difference, Corsi and the power kill (power-play % plus penalty-kill %).
A random team wins the Cup about 3% of the time (1 in 31).</p>
<div class="table-wrap">
<table>
<thead><tr><th class="left" scope="col">Rule (end of regular season)</th><th scope="col">Teams that met it</th>
<th scope="col">Champions among them</th><th scope="col" title="How many of the 15 champions met the rule">Champions that met it</th>
<th scope="col" title="Champions divided by teams that met the rule">Champion rate</th><th scope="col" title="Champion rate divided by the 3% base rate">Lift</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
</div>
<p class="note">Such teams do win more often than a random team, but most champions do not meet the bar, and most teams that
do meet it do not win. It narrows the field; it does not find the winner.</p>"""


FACTOR_COLUMNS = [  # (heading, how to show the value, rank column in the workbook)
    ("Record", lambda r: fmt_pct(r["Points %"]), "Record rank"),
    ("Goal difference", lambda r: f"{r['Goal difference / game']:+.2f} / game", "Goal difference rank"),
    ("Corsi (5v5)", lambda r: f"{r['Corsi %']:.1f}%", "Corsi rank"),
    ("Power kill", lambda r: f"PP {r['PP%']:.1f}% / PK {r['PK%']:.1f}%", "Power kill rank"),
]


def rule_detail(factors, top_n, k, through):
    """The list of teams that fit one 'top N in at least K of 4 factors' rule today, with the stats behind it."""
    f = factors.copy()
    f["n_hit"] = sum((f[rank] <= top_n).astype(int) for _, _, rank in FACTOR_COLUMNS)
    fit = f[f["n_hit"] >= k].sort_values(["n_hit", "Record rank"], ascending=[False, True])
    near = f[f["n_hit"] == k - 1].sort_values(["Record rank"]) if k > 1 else f.iloc[0:0]

    explain = (f"<p class='note'>Ranks are among all 32 teams, games through {esc(through)}. A highlighted number means the team is in the "
               f"top {top_n} of the league on that factor. Early in a season ranks move a lot.</p>")
    if fit.empty:
        body = f"<p><strong>No team fits this rule right now.</strong></p>"
    else:
        head = "".join(f'<th scope="col">{esc(label)}</th>' for label, _, _ in FACTOR_COLUMNS)
        rows = []
        for _, r in fit.iterrows():
            cells = "".join(
                f'<td class="{"hit" if r[rank] <= top_n else ""}">{esc(show(r))} <span class="rk">#{int(r[rank])}</span></td>'
                for _, show, rank in FACTOR_COLUMNS)
            rows.append(f'<tr><td class="who left"><strong>{esc(r["Team"])}</strong></td>{cells}<td>{int(r["n_hit"])} of 4</td></tr>')
        body = (f"<p><strong>{len(fit)} team{'s' if len(fit) != 1 else ''} fit{'s' if len(fit) == 1 else ''} this rule right now:</strong></p>"
                f'<div class="table-wrap"><table><thead><tr><th class="left" scope="col">Team</th>{head}<th scope="col">Factors in the top {top_n}</th></tr></thead>'
                f"<tbody>{''.join(rows)}</tbody></table></div>")
    if len(near):
        short = {"Record": "record", "Goal difference": "goal difference", "Corsi (5v5)": "Corsi", "Power kill": "power kill"}
        shown = [f"{r['Team']} (top {top_n} in: " + ", ".join(short[label] for label, _, rank in FACTOR_COLUMNS if r[rank] <= top_n) + ")"
                 for _, r in near.head(8).iterrows()]
        more = f" and {len(near) - 8} more" if len(near) > 8 else ""
        body += f"<p class='note'>One factor short: {esc('; '.join(shown))}{more}.</p>"
    return explain + body


def rules_section(rules, factors, through):
    if factors is None:
        return rules_section_plain(rules)
    rows, panels = [], []
    for i, (_, r) in enumerate(rules.iterrows()):
        match = re.match(r"Top (\d+) in at least (\d+) of (\d+) factors", str(r["Rule"]))
        top_n, k = int(match.group(1)), int(match.group(2))
        n_fit = int(sum((factors[rank] <= top_n).astype(int) for _, _, rank in FACTOR_COLUMNS).ge(k).sum())
        rows.append(
            f"<tr class='rule-row'><td class='who left'><button type='button' class='rule-toggle' aria-expanded='false' "
            f"aria-controls='rule-panel-{i}' data-panel='rule-panel-{i}'>{esc(r['Rule'])}</button></td>"
            f"<td>{int(r['Teams meeting the rule'])}</td><td>{int(r['Champions among them'])}</td><td>{r['Share of champions that met it']:.0%}</td>"
            f"<td>{r['Champion rate for such teams']:.0%}</td><td>{r['Lift over a random team']:.1f}x</td>"
            f"<td class='pts' data-v='{n_fit}'>{n_fit}</td></tr>")
        panels.append(
            f"<div class='rule-panel' id='rule-panel-{i}'><h3>Teams that fit &ldquo;{esc(r['Rule'])}&rdquo; right now</h3>"
            f"{rule_detail(factors, top_n, k, through)}</div>")
    return f"""
<h2>Is "top 5 in most of the factors" a rule for finding the champion?</h2>
<p>A popular idea: champions are usually near the top of the league in several key areas, so teams that match should be
favorites. The four areas here are record (points %), goal difference, Corsi (shot-attempt share at 5-on-5) and the power kill
(power-play % plus penalty-kill %). A random team wins the Cup about 3% of the time (1 in 31).</p>
<p class="note"><strong>Click a rule</strong> to see which of today's teams fit it, and the numbers behind it.</p>
<div class="table-wrap">
<table class="rules">
<thead><tr><th class="left" scope="col">Rule (end of regular season)</th><th scope="col">Teams that met it</th>
<th scope="col">Champions among them</th><th scope="col" title="How many of the 15 champions met the rule">Champions that met it</th>
<th scope="col" title="Champions divided by teams that met the rule">Champion rate</th><th scope="col" title="Champion rate divided by the 3% base rate">Lift</th>
<th scope="col" title="How many of today's teams fit the rule (games through {esc(through)})">Fit right now</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
</div>
<div class="rule-panels" aria-live="polite">{''.join(panels)}</div>
<p class="note">Such teams do win more often than a random team, but most champions do not meet the bar, and most teams that
do meet it do not win. It narrows the field; it does not find the winner.</p>"""


def contender_page(data, standings, built_text):
    about_html = about_this_page("How closely does each team's season so far look like the seasons of teams that went on to reach the conference finals?", f"""Version 1, built from NHL team statistics for 15 past seasons. It is an estimate from past patterns,
not a prediction of who will win the Cup, and not odds or a betting line. The next versions will add shot quality,
zone time and goaltending in context.""")
    body = f"""
<h1>Contender fit</h1>
{right_now_section(data["right"], data["note"], standings)}
{verdict_section(data["verdict"], data["score"])}
{profile_section(data["profile"])}
{rules_section(data["rules"], data.get("factors"), through_text(data["note"]))}
<h2>What this does not know yet</h2>
<ul>
  <li>How dangerous a team's shots are (slot and high-danger shots), and how much of the goalie's results come from the defense in front of him.</li>
  <li>How much of the game a team spends in the offensive zone, and how cleanly it breaks out.</li>
  <li>Depth: a top-line center, an elite defense pair, or a goalie who gets hot in the playoffs.</li>
</ul>
{about_html}
<p class="note">Method and code are public: <a href="{REPO_URL}">{REPO_URL.replace('https://', '')}</a>. Data: &copy; NHL.
Shot-quality data from <a href="https://moneypuck.com/data.htm">MoneyPuck.com</a> will be credited here when it is added.</p>
"""
    return page("Contender fit - " + SITE_NAME, "contender", body, built_text)


# ----------------------------------------------------------------------------
# The research pages: Red flags, Slow starters, Goaltending in context, and the Studies hub.
# Each one reads the workbook its study wrote (outputs/*.xlsx) and shows the summary only.
# ----------------------------------------------------------------------------
RED_FILE = PROJECT_FOLDER / "outputs" / "red_flags.xlsx"
SLOW_FILE = PROJECT_FOLDER / "outputs" / "slow_starters.xlsx"
GOAL_FILE = PROJECT_FOLDER / "outputs" / "goaltending_in_context.xlsx"

MONEYPUCK_NOTE = ('Shot-quality data on this page is from <a href="https://moneypuck.com/data.htm">MoneyPuck.com</a>, '
                  'used here for non-commercial purposes with credit.')


def read_sheet(path, name, header=0):
    return pd.read_excel(path, sheet_name=name, header=header)


def read_note(path, sheet="Right now"):
    return str(pd.read_excel(path, sheet_name=sheet, header=None, nrows=1).iloc[0, 0])


def through_text(note):
    match = re.search(r"games through (\d{4}-\d{2}-\d{2})", note)
    return nice_date(match.group(1)) if match else "the latest games"


def first_sentence(text, limit=330):
    text = str(text).strip()
    cut = text.find(". ")
    sentence = text if cut == -1 else text[:cut + 1]
    if len(sentence) > limit:
        sentence = sentence[:limit].rsplit(" ", 1)[0] + "..."
    return sentence


def about_this_page(question, text):
    """The page's question and notes, shown near the bottom instead of under the title. text may contain links (HTML)."""
    return f"""
<h2>About this page</h2>
<p><strong>The question:</strong> {esc(question)}</p>
<p class="note">{text}</p>
"""


def web_text(text):
    """The workbook sentences sometimes point at 'the X sheet'. On a web page that means nothing, so drop those sentences."""
    sentences = re.split(r"(?<=[.!?])\s+", str(text).strip())
    kept = [s for s in sentences if not re.search(r"\bsheets?\b|\bworkbook\b|\bExcel\b", s)]
    return " ".join(kept)


def verdict_list(verdict, skip=("Question",)):
    items = []
    for _, row in verdict.iterrows():
        key = str(row.iloc[0])
        if key in skip:
            continue
        items.append(f"<dt>{esc(key)}</dt><dd>{esc(web_text(row.iloc[1]))}</dd>")
    return f"<dl class='verdict'>{''.join(items)}</dl>"


def static_table(headers, rows, left=(0,)):
    """A plain table. headers: list of text. rows: list of lists of already-formatted text. left: columns shown left-aligned."""
    th = "".join(f'<th class="{"left" if i in left else ""}" scope="col">{esc(h)}</th>' for i, h in enumerate(headers))
    body = "".join("<tr>" + "".join(f'<td class="{"who left" if i in left else ""}">{esc(v)}</td>' for i, v in enumerate(row)) + "</tr>"
                   for row in rows)
    return f'<div class="table-wrap"><table><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table></div>'


def team_table(table_id, df, cols, conference, sort_note=True, rank_tip="Rank: position in the list as first shown"):
    """
    A sortable team table with the League / Eastern / Western buttons, and a "#" rank column on the left
    (each team's position in the list as it is first shown; rank_tip says what that order is).
    cols: list of dicts with  label, get (a function of a row), kind ('num' or 'text'), fmt (function), tip,
          first ('ascending' or None), sub (optional function giving a small second line, for the team column).
    """
    head = [f'<th data-sort="num" data-first="ascending" title="{esc(rank_tip)}" scope="col"><button type="button">#</button></th>']
    for c in cols:
        left = " left" if c.get("sub") else ""
        first = f' data-first="{c["first"]}"' if c.get("first") else ""
        head.append(f'<th class="{left.strip()}" data-sort="{c["kind"]}"{first} title="{esc(c.get("tip", c["label"]))}" scope="col">'
                    f'<button type="button">{esc(c["label"])}</button></th>')
    rows = []
    for position, (_, r) in enumerate(df.iterrows(), 1):
        cells = [f'<td class="rank" data-v="{position}">{position}</td>']
        for c in cols:
            raw = c["get"](r)
            shown = c["fmt"](raw) if c.get("fmt") else str(raw)
            if c["kind"] == "num":
                try:
                    data_v = "" if pd.isna(raw) else str(float(raw))
                except (TypeError, ValueError):
                    data_v = ""
            else:
                data_v = esc(raw)
            if c.get("sub"):
                cells.append(f'<td class="who left" data-v="{data_v}"><strong>{esc(shown)}</strong><span class="why">{esc(c["sub"](r))}</span></td>')
            else:
                cells.append(f'<td data-v="{data_v}">{esc(shown)}</td>')
        rows.append(f'<tr data-conf="{esc(conference.get(r.get("Team", ""), ""))}">{"".join(cells)}</tr>')
    return f"""
<div class="chips" data-filter-group="{table_id}" role="group" aria-label="Filter teams">
  <button type="button" data-filter="all" aria-pressed="true">League</button>
  <button type="button" data-filter="Eastern" aria-pressed="false">Eastern</button>
  <button type="button" data-filter="Western" aria-pressed="false">Western</button>
</div>
<div class="table-wrap">
<table class="sortable ranked" id="{table_id}">
<thead><tr>{''.join(head)}</tr></thead>
<tbody>
{chr(10).join(rows)}
</tbody>
</table>
</div>"""


def conference_map(standings):
    return {row["team_name"]: row["conference"] for _, row in standings.iterrows()}


def rate_text(count, total):
    return f"{int(count)} of {int(total)} ({count / total:.0%})" if total else "-"


# ---- Red flags -------------------------------------------------------------
def red_flags_page(standings, built_text):
    v = read_sheet(RED_FILE, "Verdict")
    flags = read_sheet(RED_FILE, "The flags")
    by_flag = read_sheet(RED_FILE, "Flag by flag")
    stacking = read_sheet(RED_FILE, "Stacking")
    right = read_sheet(RED_FILE, "Right now", header=1)
    note = read_note(RED_FILE)

    flag_rows = []
    for name in flags["Flag"]:
        part = by_flag[by_flag["Flag"] == name]
        row = [name, str(part["Verdict across the 3 checkpoints"].iloc[0])]
        for cp in ("Thanksgiving", "Jan 1"):
            p = part[part["Checkpoint"] == cp].iloc[0]
            row += [f"{int(p['Flagged teams'])} teams", f"{p['Miss rate, flagged']:.0%} vs {p['Miss rate, not flagged']:.0%}"]
        flag_rows.append(row)
    flag_table = static_table(["Flag", "Verdict across 3 checkpoints", "Thanksgiving: flagged", "Missed: flagged vs not",
                               "Jan 1: flagged", "Missed: flagged vs not"], flag_rows, left=(0, 1))

    stack_rows = []
    for group in ["0 flags", "1 flag", "2 flags", "3 or more flags"]:
        row = [group]
        for cp in ("Thanksgiving", "Jan 1", "Trade deadline"):
            p = stacking[(stacking["Checkpoint"] == cp) & (stacking["Flags"] == group)]
            row.append(rate_text(p["Missed"].iloc[0], p["Teams"].iloc[0]) if len(p) else "-")
        stack_rows.append(row)
    stack_table = static_table(["Flags a team had", "Missed at Thanksgiving", "Missed at Jan 1", "Missed at the deadline"], stack_rows)

    definitions = "".join(f"<dt>{esc(r['Flag'])}</dt><dd>{esc(r['What it means'])}</dd>" for _, r in flags.iterrows())
    right["Flags"] = right["Flags"].fillna("None")
    cols = [
        {"label": "Team", "get": lambda r: r["Team"], "kind": "text", "first": "ascending", "sub": lambda r: r["Flags"]},
        {"label": "Flags", "get": lambda r: r["Number of flags"], "kind": "num", "fmt": lambda x: str(int(x)), "tip": "Number of warning signs"},
        {"label": "Cushion", "get": lambda r: r["Cushion above the cut line"], "kind": "num", "fmt": lambda x: f"{x:+.3f}",
         "tip": "Points % above the weakest team that is in a playoff spot"},
        {"label": "P%", "get": lambda r: r["Points %"], "kind": "num", "fmt": fmt_pct, "tip": "Points percentage"},
        {"label": "GP", "get": lambda r: r["Games played"], "kind": "num", "fmt": lambda x: str(int(x)), "tip": "Games played"},
        {"label": "Last 10", "get": lambda r: r["Last 10 (points)"], "kind": "text", "tip": "Points won in the last 10 games"},
    ]
    about_html = about_this_page("Being in a playoff spot is not the same as being safe. Do warning signs make a team that is \"in\" more likely to miss?", f"""Six warning signs, checked on teams that were in a playoff spot at Thanksgiving, Jan 1 or the trade deadline in 15 past seasons
(2009-10 on). A flag is a reason to look closer, not a prediction. Terms such as cushion, PDO and AUC are explained on the <a href="about.html">About</a> page.""")
    body = f"""
<h1>Red flags</h1>
<h2>Right now <span class="note">&middot; teams in a playoff spot, games through {esc(through_text(note))}</span></h2>
<p class="note">Sorted by number of flags, then by how thin the cushion is. Early in a season (about 20 games) flags are noisy.</p>
{team_table("flags-table", right, cols, conference_map(standings), rank_tip="Rank on this list: most flags first, then the thinnest cushion")}

<h2>What we found</h2>
{verdict_list(v)}

<h2>The six flags, one by one</h2>
<p>"Missed" is how many flagged teams missed the playoffs after all, against equally safe teams without the flag. "Most consistent"
means flagged teams missed more than equally safe teams at all three checkpoints.</p>
{flag_table}

<h2>Do flags stack?</h2>
<p>How often teams in a playoff spot missed, by how many flags they had.</p>
{stack_table}

<h2>What the flags mean</h2>
<dl class="glossary">{definitions}</dl>

{about_html}
<h2>How it was tested, and the limits</h2>
<ul>
  <li>The six flags and their cut-offs (top or bottom quarter of the league that season) were fixed before looking at results.</li>
  <li>Each flagged team is compared with teams the same distance above the playoff line, so a flag cannot win just by meaning "barely in".</li>
  <li>With 6 flags at 3 checkpoints, a couple of strong-looking results can happen by chance, and the groups with 3 or more flags are small.</li>
  <li>Only team statistics are used: no injuries, trades, schedule or shot quality.</li>
</ul>
<p class="note">Method and code are public: <a href="{REPO_URL}">{REPO_URL.replace('https://', '')}</a>. Data: &copy; NHL.</p>
"""
    return page("Red flags - " + SITE_NAME, "studies", body, built_text)


# ---- Slow starters ---------------------------------------------------------
BAND_LABEL = {
    "ALL teams out of a spot": "All teams out of a spot",
    "Out: just out (within 0.03 below)": "Just out (within about 5 points)",
    "Out: behind (0.03 to 0.08 below)": "Behind (about 5 to 13 points)",
    "Out: far behind (more than 0.08 below)": "Far behind (more than about 13 points)",
}


def slow_starters_page(standings, built_text):
    v = read_sheet(SLOW_FILE, "Verdict")
    rates = read_sheet(SLOW_FILE, "Recovery rates")
    takes = read_sheet(SLOW_FILE, "What it takes")
    signs = read_sheet(SLOW_FILE, "The signs")
    by_sign = read_sheet(SLOW_FILE, "Sign by sign")
    who = read_sheet(SLOW_FILE, "Who recovered")
    right = read_sheet(SLOW_FILE, "Right now", header=1)
    note = read_note(SLOW_FILE)

    checkpoints = ["Nov 1", "Thanksgiving", "Jan 1", "Trade deadline"]
    rate_rows = []
    for group, label in BAND_LABEL.items():
        row = [label]
        for cp in checkpoints:
            p = rates[(rates["Checkpoint"] == cp) & (rates["Group"] == group)]
            row.append(rate_text(p["Recovered"].iloc[0], p["Teams"].iloc[0]) if len(p) else "-")
        rate_rows.append(row)
    rate_table = static_table(["How far behind"] + [f"Out at {c}" if c != "Trade deadline" else "Out at the deadline" for c in checkpoints], rate_rows)

    bands = [b for b in takes["Points % needed the rest of the way"].unique()]
    take_rows = []
    for band in bands:
        row = [f"Needing {band.lower()}"]
        for cp in ("Thanksgiving", "Jan 1"):
            p = takes[(takes["Checkpoint"] == cp) & (takes["Points % needed the rest of the way"] == band)]
            row.append(rate_text(p["Recovered"].iloc[0], p["Teams"].iloc[0]) if len(p) and p["Teams"].iloc[0] > 0 else "-")
        take_rows.append(row)
    take_table = static_table(["Points % needed over the rest of the season", "Recovered, out at Thanksgiving", "Recovered, out at Jan 1"], take_rows)

    sign_rows = []
    for _, s in signs.iterrows():
        part = by_sign[by_sign["Sign"] == s["Sign"]]
        sign_rows.append([s["Sign"], str(part["Verdict across the 4 checkpoints"].iloc[0]), s["What it means"]])
    sign_table = static_table(["Sign of a coming recovery", "Verdict across 4 checkpoints", "What it means"], sign_rows, left=(0, 1, 2))

    finalists = who[who["How the season ended"].isin(["Won Cup", "Lost Cup final"])]
    final_rows = [[r["Season"], r["Team"], fmt_pct(r["Points % at Thanksgiving"]), fmt_pct(r["Points % needed the rest of the way"]), r["How the season ended"]]
                  for _, r in finalists.iterrows()]
    final_table = static_table(["Season", "Team", "Points % at Thanksgiving", "Points % needed the rest of the way", "How the season ended"],
                               final_rows, left=(0, 1, 4))

    right["Signs"] = right["Signs"].fillna("No signs")
    history_col = [c for c in right.columns if c.startswith("Teams that far behind")][0]
    history_at = re.search(r"behind at (.+?) \(history\)", history_col)
    history_at = history_at.group(1) if history_at else "this point"
    cols = [
        {"label": "Team", "get": lambda r: r["Team"], "kind": "text", "first": "ascending",
         "sub": lambda r: f"{r['How far behind'].capitalize()} (of teams this far behind at {history_at}, {r[history_col]}). {r['Signs']}"},
        {"label": "Needs", "get": lambda r: r["Points % needed the rest of the way"], "kind": "num", "fmt": fmt_pct,
         "tip": "Points % needed in the remaining games to reach the typical final cut line"},
        {"label": "vs line", "get": lambda r: r["Points % vs the cut line (negative = below)"], "kind": "num", "fmt": lambda x: f"{x:+.3f}",
         "tip": "Points % compared with the weakest team in a playoff spot (negative = below it)"},
        {"label": "P%", "get": lambda r: r["Points %"], "kind": "num", "fmt": fmt_pct, "tip": "Points percentage"},
        {"label": "GP", "get": lambda r: r["Games played"], "kind": "num", "fmt": lambda x: str(int(x)), "tip": "Games played"},
        {"label": "Signs", "get": lambda r: r["Number of signs"], "kind": "num", "fmt": lambda x: str(int(x)), "tip": "Number of signs of a coming recovery"},
    ]
    about_html = about_this_page("A team out of a playoff spot at Thanksgiving: how often does it recover, what does it take, and is there any sign of which ones will?", f"""Teams that were out of a playoff spot at Nov 1, Thanksgiving, Jan 1 or the trade deadline in 15 past seasons (2009-10 on).
A slow start is a reason to check how far behind a team is and what pace it needs, not a prediction. Terms such as cut line and AUC are explained on the <a href="about.html">About</a> page.""")
    body = f"""
<h1>Slow starters</h1>
<h2>Right now <span class="note">&middot; teams out of a playoff spot, games through {esc(through_text(note))}</span></h2>
<p class="note">"Needs" is the points % the team must earn in its remaining games to reach the typical final cut line. Early in a season the signs are noisy.</p>
{team_table("slow-table", right, cols, conference_map(standings), rank_tip="Rank on this list: closest to a playoff spot first")}

<h2>What we found</h2>
{verdict_list(v)}

<h2>How often teams recover</h2>
<p>Teams out of a playoff spot at each checkpoint that made the playoffs after all, by how far behind they were.</p>
{rate_table}

<h2>What it takes</h2>
<p>The points % a team had to earn over the rest of the season to reach the typical final cut line, and how often teams that needed it recovered.</p>
{take_table}

<h2>Signs of a coming recovery</h2>
<p>Six signs, fixed in advance as the top or bottom quarter of the league: the mirror image of the <a href="red-flags.html">red flags</a>.
"Most consistent" would mean teams with the sign recovered more than equally far behind teams at every checkpoint.</p>
{sign_table}

<h2>Slow starters that reached the Cup final</h2>
<p>Of the {len(who)} teams out of a spot at Thanksgiving that recovered, these went all the way to the final.</p>
{final_table}

{about_html}
<h2>How it was tested, and the limits</h2>
<ul>
  <li>Each team with a sign is compared with teams the same distance behind, so a sign cannot win just by meaning "only slightly behind".</li>
  <li>With 6 signs at 4 checkpoints, a few results will look interesting by chance. The groups are small (about 220 teams per checkpoint, 50 to 70 recoveries).</li>
  <li>Team statistics only: no injuries, trades, schedule or goalie changes.</li>
</ul>
<p class="note">Method and code are public: <a href="{REPO_URL}">{REPO_URL.replace('https://', '')}</a>. Data: &copy; NHL.</p>
"""
    return page("Slow starters - " + SITE_NAME, "studies", body, built_text)


# ---- Goaltending in context ------------------------------------------------
def goaltending_page(standings, built_text):
    v = read_sheet(GOAL_FILE, "Verdict")
    repeats = read_sheet(GOAL_FILE, "What repeats")
    follow = read_sheet(GOAL_FILE, "Labels after Jan 1")
    playoff = read_sheet(GOAL_FILE, "Playoff goaltending")
    hot = read_sheet(GOAL_FILE, "Hot goalies")
    right = read_sheet(GOAL_FILE, "Right now", header=1)
    note = read_note(GOAL_FILE)

    repeat_rows = [[r["Statistic"], r["Better when"], f"{r['Repeats (before vs after Jan 1)']:.2f}", str(r["Carry-over (share of the gap still there)"])]
                   for _, r in repeats.iterrows()]
    repeat_table = static_table(["Statistic", "Better when", "Repeats (before vs after Jan 1)", "Share of the gap still there"], repeat_rows)

    points_col = "Standings points after Jan 1, vs what the record predicted"
    follow_rows = [[r["Label (as of Jan 1)"], str(int(r["Team-seasons"])), fmt_pct(r["Points % before Jan 1"]), fmt_pct(r["Points % after Jan 1"]),
                    f"{r[points_col]:+.1f} (+/-{r['Margin (+/-, 2 std. errors, in points)']:.1f})"] for _, r in follow.iterrows()]
    follow_table = static_table(["Label as of Jan 1", "Team-seasons", "Points % before", "Points % after", "Points after Jan 1 vs what the record predicted"],
                                follow_rows)

    playoff_rows = [[r["How the playoffs ended"], str(int(r["Teams"])), f"{r['Playoff goaltending (avg percentile among playoff teams)']:.0f}",
                     f"{r['Regular-season goaltending (avg percentile among playoff teams)']:.0f}", str(r["Share with positive playoff GSAx"]),
                     f"{r['Average playoff GSAx (goals)']:+.1f}"] for _, r in playoff.iterrows()]
    playoff_table = static_table(["How the playoffs ended", "Teams", "Playoff goaltending (percentile)", "Regular-season goaltending (percentile)",
                                  "Had positive playoff GSAx", "Average playoff GSAx (goals)"], playoff_rows)

    hot_rows = [[r["Season"], r["Team"], r["Starter"], r["How the playoffs ended"], f"{r['Playoff GSAx (goals)']:+.1f}",
                 str(r["Regular-season rank among playoff teams"])] for _, r in hot.head(10).iterrows()]
    hot_table = static_table(["Season", "Team", "Starter", "How the playoffs ended", "Playoff GSAx (goals)", "Regular-season goaltending rank"],
                             hot_rows, left=(0, 1, 2, 3))

    cols = [
        {"label": "Team", "get": lambda r: r["Team"], "kind": "text", "first": "ascending", "sub": lambda r: r["Label"]},
        {"label": "Defense", "get": lambda r: r["Defense rank (1 = best)"], "kind": "num", "fmt": lambda x: str(int(x)), "first": "ascending",
         "tip": "Rank in the league for defense (expected goals against per 60, 5v5): 1 = best"},
        {"label": "Goalie", "get": lambda r: r["Goaltending rank (1 = best)"], "kind": "num", "fmt": lambda x: str(int(x)), "first": "ascending",
         "tip": "Rank in the league for goaltending (goals saved above expected per 60): 1 = best"},
        {"label": "GSAx", "get": lambda r: r["GSAx (goals)"], "kind": "num", "fmt": lambda x: f"{x:+.1f}",
         "tip": "Goals saved above expected so far this season"},
        {"label": "GA", "get": lambda r: r["Goals against"], "kind": "num", "fmt": lambda x: str(int(x)), "tip": "Goals against"},
        {"label": "xGA", "get": lambda r: r["Expected goals against"], "kind": "num", "fmt": lambda x: f"{x:.1f}", "tip": "Expected goals against"},
        {"label": "GP", "get": lambda r: r["Games played"], "kind": "num", "fmt": lambda x: str(int(x)), "tip": "Games played"},
    ]
    about_html = about_this_page("Is a goalie bailing out a weak defense, or is a strong defense being let down by its goalie? And can anything foresee a hot playoff goalie?", f"""Goals against split into two parts with MoneyPuck's shot-quality data: <strong>defense</strong> (how many goals the shots allowed should produce:
expected goals against per 60 minutes at 5-on-5) and <strong>goaltending</strong> (goals saved above expected, "GSAx"). 15 past seasons, 2009-10 on.
Terms are explained on the <a href="about.html">About</a> page. {MONEYPUCK_NOTE}""")
    body = f"""
<h1>Goaltending in context</h1>
<h2>Right now <span class="note">&middot; games through {esc(through_text(note))}</span></h2>
<p class="note">Teams are labeled "both good", "goalie bailing out the defense", "defense let down by the goalie" or "both weak" by whether each part is above or below the league median.
Early in a season these labels are noisy.</p>
{team_table("goalie-table", right, cols, conference_map(standings), rank_tip="Rank on this list: by goaltending (goals saved above expected per 60), best first")}

<h2>What we found</h2>
{verdict_list(v)}

<h2>What repeats</h2>
<p>Does a team's number before Jan 1 look like its number after Jan 1? Near 0 means mostly luck; higher means a stable feature of the team.</p>
{repeat_table}

<h2>What happens after Jan 1, by label</h2>
<p>How teams did over the rest of the season compared with what their record before Jan 1 predicted. Margins are about two standard errors.</p>
{follow_table}

<h2>Playoff goaltending</h2>
<p>How each group of playoff teams ranked on goaltending, in the regular season and in the playoffs (percentile among playoff teams, 100 = best).</p>
{playoff_table}

<h2>The hottest playoff runs</h2>
<p>The ten playoff runs with the most goals saved above expected, and where that team's goaltending ranked among the 16 playoff teams in the regular season.</p>
{hot_table}

{about_html}
<h2>How it was tested, and the limits</h2>
<ul>
  <li>The team numbers blend every goalie a team used, including the backup. Playoff data for individual goalies is not in the free files.</li>
  <li>"Expected goals" is one reasonable shot-quality model, not the truth: shot quality it cannot see (screens, passes before the shot) shows up as goaltending.</li>
  <li>"Defense" here means shot quality allowed, not individual defenders.</li>
  <li>Forecast tests are built on the other seasons and used on the one they have never seen, with the same noise margins as the other pages.</li>
</ul>
<p class="note">Method and code are public: <a href="{REPO_URL}">{REPO_URL.replace('https://', '')}</a>. Data: &copy; NHL and MoneyPuck.com.</p>
"""
    return page("Goaltending in context - " + SITE_NAME, "studies", body, built_text, moneypuck=True)


# ---- Breakouts -------------------------------------------------------------
BREAKOUT_FILE = PROJECT_FOLDER / "outputs" / "breakout_round2.xlsx"
BREAKOUT_INDEX_FILE = PROJECT_FOLDER / "outputs" / "breakout_index.xlsx"
BREAKOUT_NOW_FILE = PROJECT_FOLDER / "outputs" / "breakout_now.xlsx"
BREAKOUT_BEI_FILE = PROJECT_FOLDER / "outputs" / "breakout_bei.xlsx"
BREAKOUT_FROZEN = PROJECT_FOLDER / "reference" / "breakout_metric_frozen.json"
BREAKOUT_STATS = [  # (name in the workbook, what it counts)
    ("BER (breakout efficiency)", "Controlled exits, plus half credit for safe clears, as a share of the breakout attempts that could be judged."),
    ("Pinned-in spells", "Times the opponent gets two or more shot attempts in a row without play leaving the zone. Fewer is better."),
    ("Dangerous pinned-in spells", "Pinned-in spells that include a shot from the slot. Fewer is better."),
    ("Defence-to-offence", "Own-zone puck wins that become the team's own shot attempt within 20 seconds."),
    ("Zone-clear rate", "Own-zone puck wins after which play is next seen outside the zone within 20 seconds."),
    ("Zone exit time", "Average seconds from an own-zone puck win until play is seen outside the zone. Shorter is better."),
    ("Transition shot quality", "Share of quick-strike shots (within 5 seconds of play coming up the ice) taken from the slot."),
    ("Own-zone turnovers", "Giveaways in the team's own zone plus pucks taken off it there. Fewer is better."),
    ("Quick exits", "Share of own-zone puck wins that are out of the zone within 8 seconds."),
    ("Costly giveaways", "Own-zone giveaways followed by a shot attempt against within 10 seconds. Fewer is better."),
    ("Transition shots", "Shot attempts taken within 5 seconds of play coming up the ice."),
    ("Icings", "How often a team ices the puck. Fewer is better."),
]


def signed(value):
    return f"{float(value):+.2f}"


def breakouts_page(standings, built_text):
    import json
    v = read_sheet(BREAKOUT_FILE, "Verdict")
    summary = read_sheet(BREAKOUT_FILE, "Summary", header=1)
    blocks = read_sheet(BREAKOUT_FILE, "Blocks", header=1).set_index("Block")
    forms = read_sheet(BREAKOUT_FILE, "By form", header=1).set_index("Form")
    yard = read_sheet(BREAKOUT_FILE, "Yardsticks")
    teams = read_sheet(BREAKOUT_FILE, "Teams")
    frozen = json.loads(BREAKOUT_FROZEN.read_text(encoding="utf-8"))
    predicts, wins, repeats = "On its own: predicts unseen games", "On its own: goes with winning", "Repeats (odd vs even games)"
    corsi = yard[yard["Stat"].str.startswith("Corsi")].iloc[0]

    # --- the stats, one by one (each team's own number), best predictor first, with the two familiar yardsticks underneath ---
    stat_rows = sorted(([name, what, blocks.loc[f"{name} [ours]"]] for name, what in BREAKOUT_STATS), key=lambda r: -r[2][predicts])
    stat_table = static_table(
        ["Stat", "What it counts", "Goes with winning", "Predicts unseen games", "Repeats"],
        [[name.replace(" (breakout efficiency)", ""), what, signed(b[wins]), signed(b[predicts]), f"{b[repeats]:.2f}"] for name, what, b in stat_rows]
        + [[f"Yardstick: {r['Stat']}", "", signed(r["Goes with winning"]), signed(r["Predicts unseen games"]), f"{r[repeats]:.2f}"] for _, r in yard.iterrows()],
        left=(0, 1))

    # --- the two frozen recipes ---
    def honest(tier_label, question, kind):
        row = summary[(summary["Recipe"] == tier_label) & summary["Question"].str.startswith(question)].iloc[0]
        col = "Honest result (data-set weights)" if kind.startswith("data") else "Honest result (plain average of the chosen blocks)"
        return float(row[col])

    def in_words(metric):
        names = [n.replace("[ours", "[the team's own").replace("[", "(").replace("]", ")") for n in metric["names"].values()]
        if metric["kind"].startswith("data"):
            return f"A weighted mix of all {len(names)} versions of the team's own stats. Heaviest weights: {'; '.join(names[:4])}."
        return "The plain average of: " + "; ".join(names) + "."

    recipe_rows, numbers = [], {}
    for tier, label, plain in (("own", "Our own breakouts", "Own breakouts"), ("battle", "Breakout battle", "Breakout battle")):
        m = frozen["metrics"][tier]
        numbers[tier] = {q: honest(label, start, m["kind"]) for q, start in
                         (("wins", "Goes with winning"), ("predicts", "PREDICTS"), ("corsi", "More than shot share (beyond Corsi)"), ("both", "More than shot share and shot quality"))}
        n = numbers[tier]
        recipe_rows.append([plain, in_words(m), signed(n["wins"]), signed(n["predicts"]), signed(n["corsi"]), signed(n["both"])])
    recipe_table = static_table(["Recipe", "What is in it", "Goes with winning", "Predicts unseen games", "Left after shot share", "Left after shot share and shot quality"],
                                recipe_rows, left=(0, 1))

    # --- the blind test: pending, or its result ---
    blind = frozen.get("blind_test")
    if blind:
        rows = []
        for tier, plain in (("own", "Own breakouts"), ("battle", "Breakout battle")):
            o, b = blind["results"][tier]["other"], blind["results"][tier]["beyond2"]
            rows.append([plain, f"{signed(o['result'])} (+/-{o['margin']:.2f})", signed(o["corsi"]), signed(o["xg"]), f"{signed(b['result'])} (+/-{b['margin']:.2f})",
                         "Yes" if blind["verdicts"][tier]["reaches_finish_line"] else "No", "Yes" if blind["verdicts"][tier]["separate_edge"] else "No"])
        blind_html = (f"<p>Run on {esc(blind['run_on'])}, once, on seasons 2009-10 to 2013-14, which no earlier test had touched. The two recipes were frozen and "
                      "published in the project's code before it was run. "
                      "The finish line was set beforehand: predict unseen games at 0.46 or better; and, to count as a separate edge, still show something clear "
                      "after shot share and shot quality are taken out. The numbers in brackets are noise margins.</p>"
                      + "<p class='note'>\"Corsi\" and \"xG share\" are how those two yardsticks predicted unseen games on the same five seasons. "
                        "\"Left over\" is what remained after shot share and shot quality were taken out.</p>"
                      + static_table(["Recipe", "Predicts unseen games", "Corsi", "xG share", "Left over", "Reached 0.46?", "Separate edge?"], rows))
    else:
        blind_html = (f"<p>Not run yet. The two recipes were frozen on {esc(frozen['frozen_on'])} and are public in the project's code. Five seasons (2009-10 to 2013-14) "
                      "have been kept out of every test. When their play-by-play has finished downloading, the recipes will be scored on them once, and that result "
                      "will be reported here whatever it is. The finish line was set beforehand: predict unseen games at 0.46 or better (Corsi's level); and, to count "
                      "as a separate edge, still show something clear after shot share and shot quality are taken out.</p>")

    # --- what we found, in plain words ---
    best = max(stat_rows, key=lambda r: r[2][predicts])
    ber = blocks.loc["BER (breakout efficiency) [ours]"]
    battle_left = summary[(summary["Recipe"] == "Breakout battle") & summary["Question"].str.startswith("More than shot share and shot quality")].iloc[0]
    short = web_text(v[v.iloc[:, 0] == "Short answer"].iloc[0, 1]).replace("(ours minus the opponent's)", "(the team's numbers minus the opponent's)")
    short = short.replace("a blind test on five kept-back seasons is the final word", "the blind test on five kept-back seasons is the final word")
    close_call = "That is too close to call before the blind test."
    if blind:
        own_b, battle_b = blind["results"]["own"], blind["results"]["battle"]
        reached = [name for tier, name in (("own", "own-breakouts"), ("battle", "breakout-battle")) if blind["verdicts"][tier]["reaches_finish_line"]]
        edge = [name for tier, name in (("own", "own-breakouts"), ("battle", "breakout-battle")) if blind["verdicts"][tier]["separate_edge"]]
        all_positive = all(r["other"]["seasons_above_zero"] == r["other"]["seasons"] for r in (own_b, battle_b))
        blind_text = (f"On five seasons that no test had touched (2009-10 to 2013-14), the own-breakouts recipe predicted unseen games at {signed(own_b['other']['result'])} "
                      f"and the breakout-battle recipe at {signed(battle_b['other']['result'])}, against {signed(own_b['other']['corsi'])} for Corsi on the same seasons. "
                      + ("Both reached" if len(reached) == 2 else f"Only the {reached[0]} recipe reached" if reached else "Neither reached") + " the 0.46 line set beforehand. "
                      f"After shot share and shot quality were taken out, what was left was {signed(own_b['beyond2']['result'])} and {signed(battle_b['beyond2']['result'])}: "
                      + ("a separate edge for the " + " and ".join(edge) + " recipe. " if edge else "no separate edge. ")
                      + ("Both links were positive in all five seasons, so the plain finding stands: teams that break out well do win more, "
                         "but the breakout scores are not a better measure of a team than its shot share." if all_positive and not reached else ""))
        close_call = "The blind test, on fresh seasons, " + ("found an edge." if edge else "found nothing clear either.")
    found = [("Short answer", short),
             ("The best single stats", f"The best single stat is {best[0].lower()}: it predicts games it has not seen at {signed(best[2][predicts])}. "
                                       f"BER, the best stat that is not built from shot counts, is at {signed(ber[predicts])}. "
                                       f"Shot-attempt share (Corsi) is at {signed(corsi['Predicts unseen games'])}."),
             ("What helped", f"Counting each stat as a battle, the team's number minus the opponent's in the same games, helped: the average stat predicts unseen games at "
                             f"{signed(forms.loc['battle'].iloc[0])} as a battle against {signed(forms.loc['ours'].iloc[0])} on its own. Counting only moments when the score is close "
                             f"did not help ({signed(forms.loc['ours, close games'].iloc[0])})."),
             ("Is it more than out-shooting the opponent?", f"Not clearly. After shot share and shot quality are taken out, the breakout-battle recipe leaves {signed(numbers['battle']['both'])}, "
                                                            f"about the size of its noise margin ({battle_left['Noise margin (+/-)']:.2f}), and the own-breakouts recipe leaves "
                                                            f"{signed(numbers['own']['both'])}. {close_call}")]
    if blind:
        found.insert(1, ("The blind test", blind_text))
    if BREAKOUT_INDEX_FILE.exists():
        card = read_sheet(BREAKOUT_INDEX_FILE, "New stats", header=1).dropna(subset=["Contender score gain (AUC)"])
        if len(card):
            helped = int((card["Contender score gain (AUC)"] > card["Contender margin (+/-)"]).sum())
            found.append(("Does it help pick contenders?", f"Not so far. Each of the first {len(card)} breakout stats was added to the contender score and tried on seasons the score "
                                                           f"had not seen: {helped} of {len(card)} improved it by more than the noise. BER and the two recipes have not been put through that test yet."))
    found_html = "<dl class='verdict'>" + "".join(f"<dt>{esc(k)}</dt><dd>{esc(t)}</dd>" for k, t in found) + "</dl>"

    # --- team by team, the latest season in the workbook ---
    season = teams["Season"].max()
    latest = teams[teams["Season"] == season]
    share = latest[["Controlled exits (share of judged attempts)", "Safe clears (share)", "Failed (share)"]].mean()
    own_col, battle_col = "Our own breakouts score (0 = league average)", "Breakout battle score (0 = league average)"
    cols = [
        {"label": "Team", "get": lambda r: r["Team"], "kind": "text", "first": "ascending", "sub": lambda r: r["How the season ended"]},
        {"label": "BER", "get": lambda r: r["BER"], "kind": "num", "fmt": lambda x: f"{x:.1f}",
         "tip": "Breakout efficiency rating: 100 x (controlled exits + half the safe clears) / breakout attempts that could be judged"},
        {"label": "Ctrl", "get": lambda r: r["Controlled exits (share of judged attempts)"], "kind": "num", "fmt": lambda x: f"{x:.0%}",
         "tip": "Controlled exits, as a share of the breakout attempts that could be judged"},
        {"label": "Clear", "get": lambda r: r["Safe clears (share)"], "kind": "num", "fmt": lambda x: f"{x:.0%}", "tip": "Safe clears, as a share of the breakout attempts that could be judged"},
        {"label": "Fail", "get": lambda r: r["Failed (share)"], "kind": "num", "fmt": lambda x: f"{x:.0%}", "first": "ascending",
         "tip": "Failed breakouts, as a share of the attempts that could be judged (lower is better)"},
        {"label": "Pinned", "get": lambda r: r["Pinned-in spells per 60"], "kind": "num", "fmt": lambda x: f"{x:.1f}", "first": "ascending",
         "tip": "Pinned-in spells per 60 minutes of 5-on-5 (lower is better)"},
        {"label": "Own", "get": lambda r: r[own_col], "kind": "num", "fmt": lambda x: f"{x:+.2f}",
         "tip": "Own-breakouts recipe: steps above or below the league average (0 = average, +1 = a clearly good team)"},
        {"label": "Battle", "get": lambda r: r[battle_col], "kind": "num", "fmt": lambda x: f"{x:+.2f}",
         "tip": "Breakout-battle recipe: steps above or below the league average (0 = average, +1 = a clearly good team)"},
        {"label": "P%", "get": lambda r: r["Points %"], "kind": "num", "fmt": fmt_pct, "tip": "Points percentage that season"},
    ]

    # --- this season so far (written by build_breakout_now.py), if it has been built ---
    now_html = ""
    if BREAKOUT_NOW_FILE.exists():
        now = read_sheet(BREAKOUT_NOW_FILE, "Right now", header=1)
        now_cols = [{**cols[0], "sub": lambda r: f"{int(r['Games'])} games"}] + cols[1:]
        now_html = f"""
<h2>This season: BER <span class="note">&middot; games through {esc(through_text(read_note(BREAKOUT_NOW_FILE)))}</span></h2>
<p class="note">Sorted by BER. Teams have played {int(now['Games'].min())} to {int(now['Games'].max())} games, so these numbers are still noisy: a first impression, not a verdict.
"Own" and "Battle" are the two recipes described below, each compared with the other teams this season (0 = league average).
Click a heading to sort by another column, and hover over it for what it means.</p>
{team_table("breakout-now", now, now_cols, conference_map(standings), rank_tip="Rank by BER this season")}
"""

    # --- BEI, the breakout efficiency index (written by build_breakout_bei.py), this season and last ---
    bei_now_html = bei_last_html = bei_explain_html = ""
    if BREAKOUT_BEI_FILE.exists():
        bei_now = read_sheet(BREAKOUT_BEI_FILE, "This season", header=1)
        bei_last = read_sheet(BREAKOUT_BEI_FILE, "Last season", header=1)
        bei_test = read_sheet(BREAKOUT_BEI_FILE, "Test", header=1).set_index("Statistic")
        bei_last_season = str(read_note(BREAKOUT_BEI_FILE, "Last season")).replace("Season ", "")
        giveaway_col, zone_col, trans_col = ("Own-zone giveaways per hour of own-zone time", "Share of even-strength time in own zone",
                                             "Transition shot attempts per 60")

        def bei_cols(sub):
            return [
                {"label": "Team", "get": lambda r: r["Team"], "kind": "text", "first": "ascending", "sub": sub},
                {"label": "BEI", "get": lambda r: r["BEI"], "kind": "num", "fmt": lambda x: f"{x:.1f}",
                 "tip": "Breakout efficiency index, 0 to 100: the three parts as league percentiles, weighted 45 / 30 / 25"},
                {"label": "DZ GA/hr", "get": lambda r: r[giveaway_col], "kind": "num", "fmt": lambda x: f"{x:.1f}", "first": "ascending",
                 "tip": "Own-zone giveaways at 5-on-5 per hour of 5-on-5 time spent in the own zone (lower is better; 45% of BEI)"},
                {"label": "Own zone", "get": lambda r: r[zone_col], "kind": "num", "fmt": lambda x: f"{x:.1%}", "first": "ascending",
                 "tip": "Share of even-strength time spent in the own zone, from NHL EDGE puck tracking (lower is better; 30% of BEI)"},
                {"label": "Trans/60", "get": lambda r: r[trans_col], "kind": "num", "fmt": lambda x: f"{x:.2f}",
                 "tip": "Transition shot attempts per 60 minutes at 5-on-5 (higher is better; 25% of BEI)"},
                {"label": "BER", "get": lambda r: r["BER"], "kind": "num", "fmt": lambda x: f"{x:.1f}", "tip": "Breakout efficiency rating, for comparison"},
                {"label": "P%", "get": lambda r: r["Points %"], "kind": "num", "fmt": fmt_pct, "tip": "Points percentage"},
            ]

        bei_now_html = f"""
<h2>This season: BEI <span class="note">&middot; games through {esc(through_text(read_note(BREAKOUT_BEI_FILE, "This season")))}</span></h2>
<p class="note">Sorted by BEI, the breakout efficiency index (version 1): own-zone giveaways, time spent in the own zone and transition shots, each as a league
percentile, weighted 45 / 30 / 25. 100 would be best in the league at all three. Early in a season these numbers are noisy.</p>
{team_table("bei-now", bei_now, bei_cols(lambda r: f"{int(r['Games'])} games"), conference_map(standings), rank_tip="Rank by BEI this season")}
"""
        bei_last_html = f"""
<h2>Last season: BEI <span class="note">&middot; {esc(bei_last_season)}</span></h2>
<p class="note">Sorted by BEI. Click a heading to sort by another column, and hover over it for what it means.</p>
{team_table("bei-last", bei_last, bei_cols(lambda r: r["How the season ended"]), conference_map(standings), rank_tip="Rank by BEI last season")}
"""
        t = bei_test
        part = {name: t.loc[name] for name in t.index}
        bei_row, ber_row = part["BEI (breakout efficiency index)"], part["BER (breakout efficiency rating)"]
        g_row, z_row, tr_row = (part["BEI part: own-zone giveaways per hour of own-zone time"], part["BEI part: share of time in own zone"],
                                part["BEI part: transition shot attempts per 60"])
        corsi_row = part["Yardstick: Corsi (shot-attempt share)"]
        test_rows = [[name.replace("BEI part: ", "Part: ").replace("Yardstick: ", "Yardstick: "), signed(r["Link with points %"]), signed(r["Link with goal difference"]),
                      f"{r['Lowest season (points %)']:+.2f} to {r['Highest season (points %)']:+.2f}"] for name, r in t.iterrows()]
        bei_explain_html = f"""
<h2>How BEI is built</h2>
<p>BEI (breakout efficiency index, version 1) blends three numbers. Each is turned into a league percentile within the season (100 for the best team, 0 for the worst)
and then weighted:</p>
<ul>
  <li><strong>45%: own-zone giveaways per hour of own-zone time</strong> (fewer is better). Giveaways in the team's own zone at 5-on-5, from the play-by-play,
  divided by the hours of 5-on-5 play the team spent in its own zone.</li>
  <li><strong>30%: share of time spent in the own zone</strong> (less is better), at even strength, from the NHL's EDGE puck and player tracking.</li>
  <li><strong>25%: transition shot attempts per 60 minutes</strong> (more is better) at 5-on-5: a shot attempt within 5 seconds of play coming up the ice.</li>
</ul>
<p>NHL EDGE zone time starts in 2021-22, so BEI does too. The weights were set by hand; they were not fitted to the data or tested on seasons they were not set on.</p>
<h3>How BEI lines up with winning</h3>
<p>Each number below compares a team's rating with its points % and goal difference in the same season, on the {int(bei_row['Seasons'])} finished seasons with zone time.
0 means no relationship, 1 a perfect one.</p>
{static_table(["Statistic", "Points %", "Goal difference", "Range across seasons (points %)"], test_rows)}
<p>Own-zone time does most of the work: on its own it lines up with winning at {signed(z_row['Link with points %'])}, more than BEI as a whole
({signed(bei_row['Link with points %'])}). The giveaway part, which carries 45% of the weight, is barely linked with winning ({signed(g_row['Link with points %'])}),
and transition shots only a little ({signed(tr_row['Link with points %'])}). On the same seasons BER is at {signed(ber_row['Link with points %'])} and Corsi at
{signed(corsi_row['Link with points %'])}. Giving more weight to own-zone time would make BEI line up better with winning, but most of that is the same
information as shot share.</p>
"""

    about_html = about_this_page("Do teams that get out of their own end cleanly win more? And is that anything more than out-shooting the opponent?", f"""Built from the NHL's play-by-play at 5-on-5: every faceoff, shot, hit, giveaway and takeaway, with where on the ice it happened.
Ten regular seasons, 2014-15 to 2025-26 (2019-20 and 2020-21 are left out as unusual seasons). Every number in the tables is a correlation worked out on the teams' rankings
within each season: 0 means no relationship, 1 a perfect one. Terms are explained on the <a href="about.html">About</a> page.
The Corsi and expected-goals yardsticks and 5-on-5 ice time are from <a href="https://moneypuck.com/data.htm">MoneyPuck.com</a>, used here for non-commercial purposes with credit.""")
    body = f"""
<h1>Breakouts</h1>
{now_html}
{bei_now_html}
<h2>{'Last season: BER' if now_html else 'Team by team'} <span class="note">&middot; {esc(season)}{'' if now_html else ', the last full season'}</span></h2>
<p class="note">Sorted by BER. Click a heading to sort by another column, and hover over it for what it means. "Own" and "Battle" are the two recipes described below.{'' if now_html else " This season's teams are not in the table yet."}</p>
{team_table("breakout-table", latest, cols, conference_map(standings), rank_tip="Rank by BER last season")}
{bei_last_html}

<h2>What we found (BER)</h2>
{found_html}

<h2>The stats, one by one</h2>
<p>Each row is a team's own number. "Goes with winning" compares it with points % and goal difference in the same season. "Predicts unseen games" builds the stat from
half of a season's games and compares it with the results of the other half. "Repeats" compares the stat from one half of the games with the same stat from the other half:
the higher it is, the more it is a real team trait and the less it is luck.</p>
{stat_table}

<h2>Two recipes</h2>
<p>Every combination of up to four of these stats was tried (about 867,000 of them), each also counted in close-score moments only and as a battle against the opponent.
To avoid keeping a lucky winner, the best combination was always chosen without the season it was scored on. "Own breakouts" uses only the team's own numbers;
"breakout battle" also uses what the team does to its opponents, so it measures the forecheck as well as the breakout.</p>
{recipe_table}

<h2>The blind test</h2>
{blind_html}

<h2>How BER is read from the play-by-play</h2>
<p>BER (breakout efficiency rating) is 100 &times; (controlled exits + half the safe clears) &divide; breakout attempts. The play-by-play does not record who has the puck,
so each breakout attempt is judged by the first telling event in the 20 seconds after it:</p>
<ul>
  <li><strong>Breakout attempt:</strong> the team wins the puck in its own zone: a faceoff win, a takeaway, a blocked shot, or the opponent missing the net.</li>
  <li><strong>Controlled exit:</strong> the next telling event is beyond the blue line and shows the team still had the puck: its own shot attempt, its own giveaway, a hit on one of its players, or the puck being taken off it there.</li>
  <li><strong>Safe clear:</strong> the next telling event is beyond the blue line but the opponent has the puck, and it does not come straight back.</li>
  <li><strong>Failed:</strong> a shot attempt against, a turnover in the zone, or an icing comes first.</li>
  <li><strong>Cannot be judged:</strong> nothing telling happens in time. These attempts are left out. About 4 in 10 attempts end up here.</li>
</ul>
<p>In {esc(season)}, of the attempts that could be judged, {share.iloc[0]:.0%} were controlled exits, {share.iloc[1]:.0%} safe clears and {share.iloc[2]:.0%} failed.</p>
{bei_explain_html}
{about_html}
<h2>How it was tested, and the limits</h2>
<ul>
  <li>These are event counts, not puck tracking. A clean exit often leaves no event behind, while a shot against is always recorded, so failures are over-counted.
  The level of BER here cannot be compared with hand-tracked numbers; only the ranking of teams means something.</li>
  <li>A weak result can mean breakouts matter less than believed, or that events cannot see breakouts well enough. This page cannot tell those apart.</li>
  <li>A link is not proof of cause: good teams may break out well because they are good, not the other way round.</li>
  <li>These are season-level results about how strong a team is. They say nothing about who wins a single game.</li>
  <li>With this many combinations tried, the numbers from the ten search seasons are flattered a little by the search. The blind test above is the fair judge.</li>
  <li>BEI is a first version: its weights were set by hand and it has not been through the hold-out or blind tests that BER went through. Its transition count is this
  site's own play-by-play version, so it will not match tracking-based transition numbers published elsewhere.</li>
</ul>
<p class="note">Method and code are public: <a href="{REPO_URL}">{REPO_URL.replace('https://', '')}</a>. Data: &copy; NHL and MoneyPuck.com.</p>
"""
    return page("Breakouts - " + SITE_NAME, "studies", body, built_text, moneypuck=True)


# ---- The Studies hub -------------------------------------------------------
def studies_page(available, built_text):
    """available: list of (title, href, question, headline_text, tag) for the pages that exist."""
    cards = "".join(
        f'<li><strong><a href="{esc(href)}">{esc(title)}</a></strong><span class="tag live">Live</span><br>'
        f'<span class="note">{esc(question)}</span><br><span>{esc(headline)}</span></li>'
        for title, href, question, headline in available)
    body = f"""
<h1>Studies</h1>
<ul class="roadmap">{cards}</ul>
<h2>Coming next</h2>
<ul class="roadmap">
  <li><strong>Zone time</strong><span class="tag">Planned</span><br><span class="note">How much of the game a team spends in the offensive zone.</span></li>
  <li><strong>Team profiles over time</strong><span class="tag">Planned</span><br><span class="note">How each team's strengths and weaknesses change as the season goes on.</span></li>
</ul>
<h2>About these studies</h2>
<p class="note">Each study asks one hockey question of up to 15 seasons of NHL data (2009-10 on), tests the answer on seasons it has not seen,
and says plainly when the evidence is thin. Every page has a team-by-team list and a plain-English summary of what the history says.
The pages are estimates from past patterns, not predictions or betting advice.</p>
"""
    return page("Studies - " + SITE_NAME, "studies", body, built_text)


# ----------------------------------------------------------------------------
# main() is where the script starts.
# ----------------------------------------------------------------------------
def main():
    sys.stdout.reconfigure(errors="replace")
    standings_path = DATA_FOLDER / "standings.csv"
    schedule_path = DATA_FOLDER / "schedule_today.csv"
    if not standings_path.exists():
        print("I can't find data/standings.csv yet.")
        print("Run this first:  python scripts/pull_schedule_standings.py")
        return 1

    standings = pd.read_csv(standings_path)
    try:
        schedule = pd.read_csv(schedule_path)
    except (FileNotFoundError, pd.errors.EmptyDataError):
        schedule = pd.DataFrame()               # no games file, or an empty one: show a friendly message

    now = datetime.now(EASTERN)
    built_text = f"Page built {now:%b} {now.day}, {now:%Y} at {now.strftime('%I:%M %p').lstrip('0')} ET."

    SITE_FOLDER.mkdir(exist_ok=True)
    pages = {
        "index.html": home_page(schedule, standings, built_text),
        "about.html": about_page(built_text),
    }
    available = []                                   # pages for the Studies hub
    if CONTENDER_FILE.exists():
        pages["contender-fit.html"] = contender_page(read_contender(CONTENDER_FILE), standings, built_text)
        verdict = read_sheet(CONTENDER_FILE, "Verdict")
        available.append(("Contender fit", "contender-fit.html",
                          "Does a score built from NHL team stats pick the real conference finalists better than just ranking by points %?",
                          "Headline: " + first_sentence(verdict[verdict.iloc[:, 0] == "Short answer"].iloc[0, 1])))
    else:
        print("Skipping the Contender fit page: outputs/contender_fit.xlsx not found.")
        print("  (run python scripts/build_contender_fit.py first; the existing page in docs/ is left as it is)")

    study_pages = [
        (RED_FILE, "red-flags.html", red_flags_page, "Red flags",
         "Being in a playoff spot is not the same as being safe: do six warning signs make a team that is in more likely to miss?", "Base rate",
         "python scripts/build_red_flags.py"),
        (SLOW_FILE, "slow-starters.html", slow_starters_page, "Slow starters",
         "A team out of a playoff spot at Thanksgiving: how often does it recover, what does it take, and is there any sign of which ones will?", "The Thanksgiving claim",
         "python scripts/build_slow_starters.py"),
        (GOAL_FILE, "goaltending.html", goaltending_page, "Goaltending in context",
         "Is a goalie bailing out a weak defense, or a strong defense being let down by its goalie? And can anything foresee a hot playoff goalie?", "What repeats",
         "python scripts/build_goaltending_study.py"),
        (BREAKOUT_FILE, "breakouts.html", breakouts_page, "Breakouts",
         "Do teams that get out of their own end cleanly win more, and is that anything more than out-shooting the opponent? Two measures: BER and BEI.", "Short answer",
         "python scripts/build_breakout_round2.py"),
    ]
    for workbook, filename, builder, title, question, key, command in study_pages:
        if workbook.exists():
            pages[filename] = builder(standings, built_text)
            verdict = read_sheet(workbook, "Verdict")
            available.append((title, filename, question, "Headline: " + first_sentence(verdict[verdict.iloc[:, 0] == key].iloc[0, 1])))
        else:
            print(f"Skipping the {title} page: {workbook.name} not found (run {command} first; the existing page in docs/ is left as it is)")
    pages["studies.html"] = studies_page(available, built_text)
    for name, content in pages.items():
        (SITE_FOLDER / name).write_text(content, encoding="utf-8")
        print(f"Wrote docs/{name}  ({len(content):,} characters)")
    print("Open docs/index.html in your browser to preview the site.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

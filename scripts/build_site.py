"""
build_site.py  -  turns your CSV files into the website pages
==============================================================
WHAT IT DOES (in Excel terms):
    Think of a mail merge. Each web page is a template with blanks. This script reads
    data/standings.csv and data/schedule_today.csv and fills the blanks, then writes the
    finished pages into the  docs/  folder. GitHub Pages publishes whatever is in  docs/.

HOW TO RUN IT (after pull_schedule_standings.py):
    python scripts/build_site.py

WHAT YOU SHOULD SEE:
    "Wrote docs/index.html" and "Wrote docs/about.html".
    Then double-click docs/index.html to preview the site in your browser.

The look of the site lives in docs/style.css and docs/site.js (those are edited by hand).
"""

import html
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
def page(title, active, body, built_text):
    def link(href, label, key):
        current = ' aria-current="page"' if key == active else ""
        return f'<a href="{href}"{current}>{label}</a>'

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)}</title>
<meta name="description" content="{esc(TAGLINE)}">
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>🏒</text></svg>">
<link rel="stylesheet" href="style.css">
</head>
<body>
<header class="site-header">
  <div class="wrap">
    <a class="brand" href="index.html">{esc(SITE_NAME)}</a>
    <nav>{link("index.html", "Home", "home")}{link("about.html", "About", "about")}</nav>
  </div>
</header>
<main class="wrap">
{body}
</main>
<footer class="site-footer">
  <div class="wrap">
    <p>Data: &copy; NHL, used for personal, non-commercial purposes. This is an independent hobby project, not affiliated with or endorsed by the NHL or any club.</p>
    <p>For information and entertainment only. Nothing here is betting or financial advice.</p>
    <p>{esc(built_text)}</p>
  </div>
</footer>
<script src="site.js"></script>
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
<table class="sortable" id="standings">
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
<p class="lead">{esc(TAGLINE)}</p>
<p class="note">Early version: tonight's games and the standings. Contender ratings and more are on the way. See <a href="about.html">About</a>.</p>
{games_section(schedule, standings)}
{standings_section(standings)}
"""
    return page(SITE_NAME, "home", body, built_text)


def about_page(built_text):
    body = """
<h1>About this project</h1>
<p class="lead">A hockey analytics project built to understand the sport better, one question at a time.</p>

<div class="prose">
<h2>What this is</h2>
<p>Most hockey tables tell you what has happened. This project tries to tell you what is really going on underneath: which teams are genuinely good, which are getting lucky, and which are headed for trouble. It looks at things like the quality of shots a team creates and allows, how its goalie performs against the shots he faces, and how special teams are doing, then compares today's teams with teams from past seasons.</p>
<p>Every view is a hypothesis to test against history, not a fact. When the data disagrees with a popular hockey belief, this site will say so.</p>
</div>

<h2>What is planned</h2>
<ul class="roadmap">
  <li><strong>Contender fit</strong><span class="tag">Planned</span><br><span class="note">How closely each team's profile matches teams that reached the conference finals in past seasons, and why.</span></li>
  <li><strong>Red flags</strong><span class="tag">Planned</span><br><span class="note">A simple checklist of warning signs, and how often teams with those signs have missed the playoffs.</span></li>
  <li><strong>Slow starters</strong><span class="tag">Planned</span><br><span class="note">Which teams that began badly tend to recover, and what separated them from the ones that did not.</span></li>
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

<div class="prose">
<h2>Data and credits</h2>
<p>Schedule and standings come from the NHL's public web data, &copy; NHL, used here for personal, non-commercial purposes. That data is unofficial and can change without notice. Other data sources, including <a href="https://moneypuck.com/data.htm">MoneyPuck.com</a>, will be added to later pages and credited wherever their data is shown.</p>
<p>No NHL logos are used. This site has no ads and does not sell anything.</p>

<h2>Disclaimer</h2>
<p>This is an independent hobby project. It is not affiliated with or endorsed by the NHL, any NHL club, or any data provider. It is for information and entertainment only and is not betting or financial advice.</p>
</div>
"""
    return page("About - " + SITE_NAME, "about", body, built_text)


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
    for name, content in pages.items():
        (SITE_FOLDER / name).write_text(content, encoding="utf-8")
        print(f"Wrote docs/{name}  ({len(content):,} characters)")
    print("Open docs/index.html in your browser to preview the site.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

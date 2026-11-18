"""
daily_update.py  -  refresh the website's data every morning and publish it
===========================================================================
WHAT IT DOES:
    Runs, in order, the same steps you would run by hand to bring the site up to date:

      1. tonight's games and the standings              pull_schedule_standings.py
      2. this season's game-by-game results             pull_game_results.py
      3. playoff and season tables                      build_reference_tables.py
      4. standings on any checkpoint date just passed   build_checkpoints.py
      5. rebuild the database, add MoneyPuck back       build_database.py, load_moneypuck.py
      6. the contender fit                              build_contender_fit.py
      7. play-by-play of the new games, and the
         Breakouts "This season" table                  pull_play_by_play.py, build_breakout_now.py
      8. NHL EDGE zone time and the breakout
         efficiency index (BEI)                         pull_edge_zone_time.py, build_breakout_bei.py
      9. rebuild the web pages                          build_site.py

    Then it commits the changed pages (the docs/ folder only) and pushes them to GitHub, which publishes the site.

    SAFETY RULES (so it never publishes something half-finished):
      * If any step fails, it stops and publishes nothing.
      * It does not run at all if the project has unsaved changes to its code (scripts/, reference/, README.md ...)
        or commits that have not been pushed yet: those are waiting for your OK, and a daily run must not publish them.
      * It only ever commits files in docs/.
      * Only one run at a time.

    Windows Task Scheduler starts it every morning (see "Daily update" in README.md). While it runs it asks Windows not to
    go back to sleep, and lets it sleep again when it is done.

HOW TO RUN IT BY HAND (in the terminal, from the "Model 2.0" folder):
    python scripts/daily_update.py              the full update, including publishing
    python scripts/daily_update.py --no-push    refresh everything on this computer but do not publish

WHERE TO LOOK AFTERWARDS:
    data/daily/daily_update.log      one line per run: when, how long, and whether it published
    data/daily/logs/<date>.log       everything each step printed, for that day's run (kept for 30 days)

DATA CREDIT:
    Data is (c) NHL, used for personal, non-commercial purposes. MoneyPuck.com data is credited on the pages that use it.
"""

import argparse
import datetime
import os
import subprocess
import sys
import time
from pathlib import Path

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
DAILY_FOLDER = PROJECT_FOLDER / "data" / "daily"
LOG_FOLDER = DAILY_FOLDER / "logs"
SUMMARY_LOG = DAILY_FOLDER / "daily_update.log"
LOCK_FILE = DAILY_FOLDER / "RUNNING"
KEEP_LOGS_DAYS = 30
STEP_TIME_LIMIT = 45 * 60          # seconds: a step that takes longer than this is treated as stuck

STEPS = [  # (script, plain name)
    ("pull_schedule_standings.py", "tonight's games and the standings"),
    ("pull_game_results.py", "this season's game results"),
    ("build_reference_tables.py", "playoff and season tables"),
    ("build_checkpoints.py", "checkpoint standings"),
    ("build_database.py", "rebuild the database"),
    ("load_moneypuck.py", "add the MoneyPuck tables back"),
    ("build_contender_fit.py", "the contender fit"),
    ("pull_play_by_play.py", "play-by-play of the new games"),
    ("build_breakout_now.py", "the Breakouts This season table"),
    ("pull_edge_zone_time.py", "this season's NHL EDGE zone time"),
    ("build_breakout_bei.py", "the breakout efficiency index (BEI)"),
    ("build_site.py", "rebuild the web pages"),
]
# Changes in these places mean someone is in the middle of work: a daily run must not publish on top of it.
PROTECTED = ["scripts", "reference", "README.md", "requirements.txt", ".gitignore"]

NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)       # run each step without popping up a window


def python_exe():
    """The console version of Python (python.exe), even when this script was started with the window-less pythonw.exe."""
    here = Path(sys.executable)
    console = here.with_name("python.exe")
    return str(console if here.name.lower() == "pythonw.exe" and console.exists() else here)


def stay_awake(on):
    """Ask Windows not to sleep while this runs (on=True), and allow it again afterwards (on=False). No settings change."""
    try:
        import ctypes
        ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | (ES_SYSTEM_REQUIRED if on else 0))
    except (AttributeError, OSError):
        pass


def git(*args, check=True):
    result = subprocess.run(["git", *args], cwd=PROJECT_FOLDER, capture_output=True, text=True, encoding="utf-8", errors="replace",
                            creationflags=NO_WINDOW, env={**os.environ, "GIT_TERMINAL_PROMPT": "0"}, timeout=300)
    if check and result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed: {(result.stderr or result.stdout).strip()[:500]}")
    return result.stdout.strip()


def ready_to_publish():
    """Returns None if it is safe to publish, or the reason it is not."""
    if git("rev-parse", "--abbrev-ref", "HEAD") != "main":
        return "the project is not on the main branch"
    dirty = git("status", "--porcelain", "--", *PROTECTED)
    if dirty:
        return "there are unsaved changes to the project's code:\n    " + dirty.replace("\n", "\n    ")
    git("fetch", "--quiet", "origin", "main")
    ahead = git("rev-list", "--count", "origin/main..main")
    behind = git("rev-list", "--count", "main..origin/main")
    if ahead != "0":
        return f"there are {ahead} commit(s) on this computer that have not been pushed yet (they are waiting for your OK)"
    if behind != "0":
        return f"GitHub has {behind} commit(s) this computer does not have"
    return None


def write_summary(line):
    DAILY_FOLDER.mkdir(parents=True, exist_ok=True)
    with SUMMARY_LOG.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M}  {line}\n")


def tidy_old_logs():
    cutoff = time.time() - KEEP_LOGS_DAYS * 86400
    for path in LOG_FOLDER.glob("*.log"):
        if path.stat().st_mtime < cutoff:
            path.unlink(missing_ok=True)


def run(push):
    started = time.time()
    LOG_FOLDER.mkdir(parents=True, exist_ok=True)
    day_log = LOG_FOLDER / f"{datetime.date.today():%Y-%m-%d}.log"

    def say(text):
        print(text, flush=True)
        with day_log.open("a", encoding="utf-8") as f:
            f.write(text + "\n")

    say(f"\n===== Daily update started {datetime.datetime.now():%Y-%m-%d %H:%M:%S}" + ("" if push else "  (no publishing)") + " =====")
    if push:
        reason = ready_to_publish()
        if reason:
            say(f"Not running today: {reason}")
            write_summary(f"SKIPPED: {reason.splitlines()[0]}")
            return 0

    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
    for number, (script, name) in enumerate(STEPS, 1):
        say(f"\n--- Step {number} of {len(STEPS)}: {name} ({script}) ---")
        t0 = time.time()
        try:
            result = subprocess.run([python_exe(), "-u", str(PROJECT_FOLDER / "scripts" / script)], cwd=PROJECT_FOLDER, capture_output=True,
                                    text=True, encoding="utf-8", errors="replace", env=env, creationflags=NO_WINDOW, timeout=STEP_TIME_LIMIT)
            output, code = (result.stdout or "") + (result.stderr or ""), result.returncode
        except subprocess.TimeoutExpired:
            output, code = f"(stopped: took longer than {STEP_TIME_LIMIT // 60} minutes)", -1
        with day_log.open("a", encoding="utf-8") as f:
            f.write(output.rstrip() + "\n")
        if code != 0:
            say(f"FAILED (exit code {code}) after {time.time() - t0:.0f} s. Nothing was published. The step's output is above.")
            write_summary(f"FAILED at step {number} ({script}); nothing published. See {day_log.relative_to(PROJECT_FOLDER)}")
            return 1
        say(f"done in {time.time() - t0:.0f} s")

    minutes = (time.time() - started) / 60
    if not push:
        say(f"\nAll steps done in {minutes:.1f} minutes. Not published (--no-push).")
        write_summary(f"OK in {minutes:.1f} min, not published (--no-push)")
        return 0

    git("add", "--", "docs")
    if not git("status", "--porcelain", "--", "docs"):
        say("\nNo page changed, so there is nothing to publish.")
        write_summary(f"OK in {minutes:.1f} min, no page changed")
        return 0
    stamp = datetime.date.today().isoformat()
    git("commit", "--quiet", "-m", f"Daily update {stamp}: standings, games, contender fit and breakouts", "--", "docs")
    git("push", "--quiet", "origin", "main")
    say(f"\nPublished: {git('log', '-1', '--format=%h %s')}. All done in {(time.time() - started) / 60:.1f} minutes.")
    write_summary(f"PUBLISHED in {(time.time() - started) / 60:.1f} min ({git('log', '-1', '--format=%h')})")
    return 0


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass                                                  # started with pythonw.exe: there is no window to print to
    parser = argparse.ArgumentParser(description="Refresh the website's data and publish it.")
    parser.add_argument("--no-push", action="store_true", help="refresh everything on this computer but do not publish")
    args = parser.parse_args()

    DAILY_FOLDER.mkdir(parents=True, exist_ok=True)
    if LOCK_FILE.exists() and time.time() - LOCK_FILE.stat().st_mtime < 3 * 3600:
        write_summary("SKIPPED: another daily update is still running")
        return 0
    LOCK_FILE.write_text(str(os.getpid()), encoding="utf-8")
    stay_awake(True)
    try:
        return run(push=not args.no_push)
    except Exception as problem:                                    # noqa: BLE001  (any surprise is logged, never silent)
        write_summary(f"FAILED: {problem}")
        print(f"FAILED: {problem}")
        return 1
    finally:
        stay_awake(False)
        LOCK_FILE.unlink(missing_ok=True)
        tidy_old_logs()


if __name__ == "__main__":
    sys.exit(main())

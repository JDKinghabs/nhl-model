"""
nhl_api.py  -  shared helper for downloading from the NHL's public web data
===========================================================================
Not meant to be run on its own. Other scripts "import" it, the way one Excel macro can
call another. Keeping this in one place means every script is equally polite to the NHL:

  * it waits at least one second between requests,
  * it retries a couple of times if the connection hiccups,
  * it explains problems in plain English instead of crashing.

Data is (c) NHL, used for personal, non-commercial purposes. The interface is unofficial
and can change without notice.
"""

import time

import requests

HEADERS = {"User-Agent": "nhl-model-hobby-project (personal, non-commercial)"}
PAUSE_SECONDS = 1.0          # minimum gap between requests

_last_request_time = 0.0     # remembers when we last asked, so we can wait the right amount


def get_json_with_status(url, params=None, tries=3):
    """
    Like get_json, but also says WHY there was no data, so a long download can tell a game that
    does not exist (not_found) from a connection problem (error: ...).
    Returns (data, status) where status is "ok", "not_found" or "error: <reason>".
    """
    global _last_request_time
    for attempt in range(1, tries + 1):
        wait = PAUSE_SECONDS - (time.time() - _last_request_time)
        if wait > 0:
            time.sleep(wait)
        _last_request_time = time.time()
        try:
            response = requests.get(url, params=params, headers=HEADERS, timeout=60)
            if response.status_code == 404:
                return None, "not_found"
            response.raise_for_status()
            return response.json(), "ok"
        except (requests.RequestException, ValueError) as problem:
            if attempt == tries:
                return None, f"error: {problem}"
            time.sleep(2 * attempt)


PAGE_SIZE = 100              # the NHL stats site sends at most 100 rows per request (since November 2026), whatever "limit" asks for


def get_all_rows(url, params, sort_by, page_size=PAGE_SIZE):
    """
    For the NHL stats site (api.nhle.com/stats/rest): ask for every row of a table, one page of up to 100 rows at a time,
    until the number of rows the site says exist ("total") has arrived. sort_by is a list of column names that make the
    order fixed from page to page (for example ["gameId", "teamId"]), so no row is skipped or sent twice.
    Returns the list of rows, or None if any page could not be fetched (so a half-finished table is never saved).
    """
    import json
    rows, start, total = [], 0, None
    sort = json.dumps([{"property": name, "direction": "ASC"} for name in sort_by])
    while total is None or start < total:
        raw = get_json(url, params={**params, "start": start, "limit": page_size, "sort": sort})
        if raw is None:
            return None
        page = raw.get("data", [])
        total = int(raw.get("total", len(page)))
        if not page:
            break
        rows.extend(page)
        start += len(page)
    if total is not None and len(rows) < total:
        print(f"  Warning: the site said {total} rows but sent {len(rows)}.")
        return None
    seen, unique = set(), []
    for row in rows:                                   # a safety net: never keep the same row twice
        key = tuple(row.get(name) for name in sort_by)
        if key not in seen:
            seen.add(key)
            unique.append(row)
    return unique


def get_json(url, params=None, tries=3, quiet_404=False):
    """
    Download a web address that returns data and give it back as Python data.
    Returns None if it could not be fetched (and prints why).
    """
    global _last_request_time
    for attempt in range(1, tries + 1):
        wait = PAUSE_SECONDS - (time.time() - _last_request_time)
        if wait > 0:
            time.sleep(wait)
        _last_request_time = time.time()
        try:
            response = requests.get(url, params=params, headers=HEADERS, timeout=60)
            if response.status_code == 404:
                if not quiet_404:
                    print(f"  Not found: {url}")
                return None
            response.raise_for_status()
            return response.json()
        except (requests.RequestException, ValueError) as problem:
            if attempt == tries:
                print(f"  Could not get {url}")
                print(f"  Reason: {problem}")
                return None
            time.sleep(2 * attempt)

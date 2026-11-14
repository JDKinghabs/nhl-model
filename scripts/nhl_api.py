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

# -*- coding: utf-8 -*-
"""Resolve optional real-world fixtures for e2e tests (repo copy or env/Desktop)."""
import os

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO_EXCEL1 = os.path.join(_TESTS, "fixtures", "excel1")
_DESKTOP_EXCEL1 = os.path.join(os.path.expanduser("~"), "Desktop", "excel1")


def excel1_dir():
    env = os.environ.get("EVRAK_EXCEL1_DIR")
    if env and os.path.isdir(env):
        return env
    if os.path.isdir(_REPO_EXCEL1):
        return _REPO_EXCEL1
    if os.path.isdir(_DESKTOP_EXCEL1):
        return _DESKTOP_EXCEL1
    return None


def mart_dir():
    env = os.environ.get("EVRAK_MART_DIR")
    if env and os.path.isdir(env):
        return env
    base = excel1_dir()
    if base:
        m = os.path.join(base, "mart")
        if os.path.isdir(m):
            return m
    return None


def has_mart():
    c, x = mart_pair()
    return bool(c and x and os.path.isfile(c) and os.path.isfile(x))


def has_excel1():
    d = excel1_dir()
    return bool(d and os.path.isdir(d))


def mart_pair():
    """Return (datev_csv, journal_xlsx) from mart folder, or (None, None)."""
    d = mart_dir()
    if not d:
        return None, None
    csvs = sorted(f for f in os.listdir(d) if f.lower().endswith(".csv"))
    xlsxs = sorted(f for f in os.listdir(d) if f.lower().endswith((".xlsx", ".xlsm")))
    if not csvs or not xlsxs:
        return None, None
    return os.path.join(d, csvs[0]), os.path.join(d, xlsxs[0])


def datev_pair():
    """Return two DATEV CSV paths from excel1, or (None, None)."""
    d = excel1_dir()
    if not d:
        return None, None
    csvs = sorted(
        f for f in os.listdir(d)
        if f.lower().endswith(".csv") and f.upper().startswith("EXTF")
    )
    if len(csvs) < 2:
        return None, None
    return os.path.join(d, csvs[0]), os.path.join(d, csvs[1])

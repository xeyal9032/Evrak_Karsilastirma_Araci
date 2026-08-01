# -*- coding: utf-8 -*-
"""Shared file logging for GUI and CLI runs."""
import logging
import os
from datetime import datetime

from paths import app_root

ROOT = app_root()
LOG_DIR = os.path.join(ROOT, "logs")
_LOGGER = None


def get_logger():
    global _LOGGER
    if _LOGGER is not None:
        return _LOGGER
    os.makedirs(LOG_DIR, exist_ok=True)
    day = datetime.now().strftime("%Y%m%d")
    path = os.path.join(LOG_DIR, f"compare_{day}.log")
    logger = logging.getLogger("evrak_karsilastirma")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        fh = logging.FileHandler(path, encoding="utf-8")
        fh.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))
        logger.addHandler(fh)
    _LOGGER = logger
    return logger


def log_run_start(f1, f2, fmt1=None, fmt2=None, source="gui"):
    log = get_logger()
    log.info("--- compare start (%s) ---", source)
    log.info("file1=%s format1=%s", f1, fmt1)
    log.info("file2=%s format2=%s", f2, fmt2)


def log_run_ok(result, elapsed_s, out_path):
    log = get_logger()
    log.info(
        "ok out=%s elapsed=%.2fs match=%s mismatch=%s only1=%s only2=%s",
        out_path,
        elapsed_s,
        result.get("match_count"),
        result.get("mismatch_count"),
        result.get("only1_count"),
        result.get("only2_count"),
    )


def log_run_err(exc, tb=None):
    log = get_logger()
    log.error("failed: %s", exc)
    if tb:
        log.error(tb)

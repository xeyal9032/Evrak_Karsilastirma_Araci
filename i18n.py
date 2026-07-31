# -*- coding: utf-8 -*-
"""Minimal JSON-based i18n helper."""
import json
import locale
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SUPPORTED = ("tr", "ru", "de", "en")
_CACHE = {}
_DEFAULT_LANG = "tr"


def _locales_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, "locales")
    return os.path.join(ROOT, "locales")


LOCALES_DIR = _locales_dir()


def detect_system_lang():
    loc = ""
    try:
        loc = locale.getlocale()[0] or ""
    except Exception:
        loc = ""
    if not loc:
        try:
            # Fallback for older Python / unset locale
            loc = locale.getdefaultlocale()[0] or ""  # noqa: DOC deprecated until 3.15
        except Exception:
            loc = ""
    loc = loc.lower().replace("-", "_")
    if loc.startswith("ru"):
        return "ru"
    if loc.startswith("de"):
        return "de"
    if loc.startswith("en"):
        return "en"
    if loc.startswith("tr"):
        return "tr"
    return _DEFAULT_LANG


def _load(lang):
    lang = (lang or _DEFAULT_LANG).lower()
    if lang not in SUPPORTED:
        lang = _DEFAULT_LANG
    if lang in _CACHE:
        return _CACHE[lang]
    path = os.path.join(_locales_dir(), f"{lang}.json")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    _CACHE[lang] = data
    return data


def t(key, lang=None, **kwargs):
    """Translate key; falls back to Turkish then to the key itself."""
    lang = lang or _DEFAULT_LANG
    data = _load(lang)
    text = data.get(key)
    if text is None:
        text = _load(_DEFAULT_LANG).get(key, key)
    if kwargs:
        try:
            return text.format(**kwargs)
        except (KeyError, ValueError):
            return text
    return text


def locale_keys(lang=None):
    return set(_load(lang or _DEFAULT_LANG).keys())

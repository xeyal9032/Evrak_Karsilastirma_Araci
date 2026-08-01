# -*- coding: utf-8 -*-
"""Writable app paths — safe for source runs and frozen PyInstaller exe."""
import os
import sys


def app_root():
    """Directory for writable data (logs/, data/).

    - Source: package directory next to this module
    - Frozen (PyInstaller): directory containing the .exe (never _MEIPASS)
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def resource_root():
    """Directory for read-only bundled resources (locales/)."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

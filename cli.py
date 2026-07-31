# -*- coding: utf-8 -*-
"""CLI entry shortcut: python cli.py file1 file2 ..."""
import sys

from karsilastir import run_cli

if __name__ == "__main__":
    raise SystemExit(run_cli(sys.argv[1:]))

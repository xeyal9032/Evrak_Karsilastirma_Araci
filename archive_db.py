# -*- coding: utf-8 -*-
"""SQLite archive for comparison runs."""
import json
import os
import sqlite3
from datetime import datetime

from paths import app_root

DEFAULT_DB = os.path.join(app_root(), "data", "compare_archive.db")


def _connect(db_path):
    os.makedirs(os.path.dirname(os.path.abspath(db_path)) or ".", exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS comparisons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            file1 TEXT NOT NULL,
            file2 TEXT NOT NULL,
            f1_label TEXT,
            f2_label TEXT,
            match_count INTEGER,
            mismatch_count INTEGER,
            only1_count INTEGER,
            only2_count INTEGER,
            f1_total INTEGER,
            f2_total INTEGER,
            out_path TEXT,
            elapsed_s REAL,
            source TEXT,
            meta_json TEXT
        )
        """
    )
    conn.commit()
    return conn


def save_comparison(result, *, file1, file2, f1_label="", f2_label="",
                    elapsed_s=None, source="cli", db_path=None, extra=None):
    db_path = db_path or DEFAULT_DB
    conn = _connect(db_path)
    try:
        cur = conn.execute(
            """
            INSERT INTO comparisons (
                created_at, file1, file2, f1_label, f2_label,
                match_count, mismatch_count, only1_count, only2_count,
                f1_total, f2_total, out_path, elapsed_s, source, meta_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(timespec="seconds"),
                file1, file2, f1_label, f2_label,
                result.get("match_count"), result.get("mismatch_count"),
                result.get("only1_count"), result.get("only2_count"),
                result.get("f1_total"), result.get("f2_total"),
                result.get("out_path"), elapsed_s, source,
                json.dumps(extra or {}, ensure_ascii=False),
            ),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def list_comparisons(db_path=None, limit=50):
    db_path = db_path or DEFAULT_DB
    if not os.path.isfile(db_path):
        return []
    conn = _connect(db_path)
    try:
        cur = conn.execute(
            """
            SELECT id, created_at, file1, file2, match_count, mismatch_count,
                   only1_count, only2_count, out_path, source
            FROM comparisons
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    finally:
        conn.close()

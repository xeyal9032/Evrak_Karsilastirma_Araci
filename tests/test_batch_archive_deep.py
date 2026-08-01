# -*- coding: utf-8 -*-
"""Batch esleme ve SQLite arsiv derin testleri."""
import json
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import archive_db
import batch_compare
from tests.helpers_datev import drow, write_datev_csv


class TestListAndPair(unittest.TestCase):
    def test_list_document_files_filters_ext_and_dirs(self):
        with tempfile.TemporaryDirectory() as tmp:
            write_datev_csv(os.path.join(tmp, "a.csv"), [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            with open(os.path.join(tmp, "ignore.txt"), "w", encoding="utf-8") as fh:
                fh.write("x")
            os.makedirs(os.path.join(tmp, "sub"))
            write_datev_csv(os.path.join(tmp, "sub", "nested.csv"), [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            files = batch_compare.list_document_files(tmp)
            names = [os.path.basename(p) for p in files]
            self.assertEqual(names, ["a.csv"])

    def test_pair_by_stem_case_insensitive(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "A")
            b = os.path.join(tmp, "B")
            os.makedirs(a)
            os.makedirs(b)
            write_datev_csv(os.path.join(a, "Mart.CSV"), [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(os.path.join(b, "mart.csv"), [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            pairs, ua, ub = batch_compare.pair_by_stem(a, b)
            self.assertEqual(len(pairs), 1)
            self.assertEqual(pairs[0][2], "mart")
            self.assertEqual(ua, [])
            self.assertEqual(ub, [])

    def test_pair_by_stem_duplicate_stem_takes_first(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "A")
            b = os.path.join(tmp, "B")
            os.makedirs(a)
            os.makedirs(b)
            # same stem different ext -> list sorts; first of each side used
            write_datev_csv(os.path.join(a, "dup.csv"), [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(os.path.join(a, "dup.xlsx"), [drow("2,00", "S", "1", "2", "0101", "Y", "t")])
            write_datev_csv(os.path.join(b, "dup.csv"), [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            pairs, *_ = batch_compare.pair_by_stem(a, b)
            self.assertEqual(len(pairs), 1)
            self.assertTrue(pairs[0][0].endswith("dup.csv") or pairs[0][0].lower().endswith("dup.csv"))

    def test_pair_zip_sorted_unequal_lengths(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "A")
            b = os.path.join(tmp, "B")
            os.makedirs(a)
            os.makedirs(b)
            write_datev_csv(os.path.join(a, "1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            write_datev_csv(os.path.join(a, "2.csv"), [drow("2,00", "S", "1", "2", "0101", "B", "t")])
            write_datev_csv(os.path.join(a, "3.csv"), [drow("3,00", "S", "1", "2", "0101", "C", "t")])
            write_datev_csv(os.path.join(b, "x.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            pairs, ua, ub = batch_compare.pair_zip_sorted(a, b)
            self.assertEqual(len(pairs), 1)
            self.assertEqual(len(ua), 2)
            self.assertEqual(len(ub), 0)


class TestRunBatch(unittest.TestCase):
    def test_run_batch_zip_writes_reports(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "A")
            b = os.path.join(tmp, "B")
            out = os.path.join(tmp, "out")
            os.makedirs(a)
            os.makedirs(b)
            write_datev_csv(os.path.join(a, "a1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            write_datev_csv(os.path.join(a, "a2.csv"), [drow("2,00", "S", "1", "2", "0101", "B", "t")])
            write_datev_csv(os.path.join(b, "b1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            write_datev_csv(os.path.join(b, "b2.csv"), [drow("2,00", "S", "1", "2", "0101", "B", "t")])
            write_datev_csv(os.path.join(b, "b3.csv"), [drow("3,00", "S", "1", "2", "0101", "C", "t")])
            summary = batch_compare.run_batch(a, b, out, mode="zip")
            self.assertEqual(summary["pair_count"], 2)
            self.assertEqual(len(summary["unmatched_b"]), 1)
            self.assertEqual(len([n for n in os.listdir(out) if n.endswith(".xlsx")]), 2)

    def test_run_batch_archive_extra_stem(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "A")
            b = os.path.join(tmp, "B")
            out = os.path.join(tmp, "out")
            db = os.path.join(tmp, "arch.db")
            os.makedirs(a)
            os.makedirs(b)
            write_datev_csv(os.path.join(a, "stem1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            write_datev_csv(os.path.join(b, "stem1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            batch_compare.run_batch(a, b, out, mode="stem", archive_db=db)
            conn = archive_db._connect(db)
            try:
                row = conn.execute("SELECT meta_json FROM comparisons").fetchone()
            finally:
                conn.close()
            meta = json.loads(row[0])
            self.assertEqual(meta.get("stem"), "stem1")

    def test_run_batch_archive_failure_is_soft_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "A")
            b = os.path.join(tmp, "B")
            out = os.path.join(tmp, "out")
            os.makedirs(a)
            os.makedirs(b)
            write_datev_csv(os.path.join(a, "stem1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            write_datev_csv(os.path.join(b, "stem1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            # Directory path is not a valid sqlite file → archive fails soft.
            bad_db = os.path.join(tmp, "not_a_db_dir")
            os.makedirs(bad_db)
            summary = batch_compare.run_batch(a, b, out, mode="stem", archive_db=bad_db)
            self.assertEqual(summary["pair_count"], 1)
            self.assertTrue(os.path.isfile(os.path.join(out, "Karsilastirma_stem1.xlsx")))
            self.assertTrue(any("archive:" in w for w in summary.get("warnings") or []))


class TestArchiveDeep(unittest.TestCase):
    def test_list_comparisons_missing_db_empty(self):
        self.assertEqual(archive_db.list_comparisons(db_path=os.path.join(tempfile.gettempdir(), "no_such_archive_xyz.db")), [])

    def test_list_comparisons_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            db = os.path.join(tmp, "a.db")
            result = {
                "match_count": 0, "mismatch_count": 0,
                "only1_count": 0, "only2_count": 0,
                "f1_total": 0, "f2_total": 0, "out_path": "x.xlsx",
            }
            for i in range(5):
                archive_db.save_comparison(
                    result, file1=f"a{i}.csv", file2=f"b{i}.csv",
                    source="test", db_path=db,
                )
            rows = archive_db.list_comparisons(db_path=db, limit=2)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["file1"], "a4.csv")

    def test_save_comparison_creates_parent_dirs(self):
        with tempfile.TemporaryDirectory() as tmp:
            db = os.path.join(tmp, "nested", "deep", "arch.db")
            rid = archive_db.save_comparison(
                {"match_count": 1, "mismatch_count": 0, "only1_count": 0,
                 "only2_count": 0, "f1_total": 1, "f2_total": 1, "out_path": "o.xlsx"},
                file1="a", file2="b", db_path=db,
            )
            self.assertIsInstance(rid, int)
            self.assertTrue(os.path.isfile(db))


if __name__ == "__main__":
    unittest.main()

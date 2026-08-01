# -*- coding: utf-8 -*-
"""CLI bayrak matrisi: batch zip, pdf, archive, exit kodlari."""
import os
import sys
import tempfile
import unittest
from io import StringIO
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import i18n
from karsilastir import _wants_cli, run_cli
from tests.helpers_datev import drow, write_datev_csv


class TestWantsCli(unittest.TestCase):
    def test_wants_cli_help(self):
        self.assertTrue(_wants_cli(["-h"]))
        self.assertTrue(_wants_cli(["--help"]))
        self.assertFalse(_wants_cli([]))
        self.assertTrue(_wants_cli(["a.csv", "b.csv"]))


class TestCliFlags(unittest.TestCase):
    def test_cli_batch_mode_zip(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "A")
            b = os.path.join(tmp, "B")
            out = os.path.join(tmp, "out")
            os.makedirs(a)
            os.makedirs(b)
            write_datev_csv(os.path.join(a, "z1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            write_datev_csv(os.path.join(a, "z2.csv"), [drow("2,00", "S", "1", "2", "0101", "B", "t")])
            write_datev_csv(os.path.join(b, "y1.csv"), [drow("1,00", "S", "1", "2", "0101", "A", "t")])
            write_datev_csv(os.path.join(b, "y2.csv"), [drow("2,00", "S", "1", "2", "0101", "B", "t")])
            buf = StringIO()
            with mock.patch("sys.stdout", buf):
                code = run_cli([a, b, "--batch", "--batch-mode", "zip", "-o", out, "--quiet"])
            self.assertEqual(code, 0)
            self.assertIn("pairs=2", buf.getvalue())
            self.assertEqual(len([n for n in os.listdir(out) if n.endswith(".xlsx")]), 2)

    def test_cli_pdf_flag(self):
        try:
            import reportlab  # noqa: F401
        except ImportError:
            self.skipTest("reportlab not installed")
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            out = os.path.join(tmp, "r.xlsx")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            buf = StringIO()
            with mock.patch("sys.stdout", buf):
                code = run_cli([f1, f2, "-o", out, "--pdf", "--quiet"])
            self.assertEqual(code, 0)
            self.assertTrue(os.path.isfile(out.replace(".xlsx", ".pdf")))

    def test_cli_default_output_timestamp_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            cwd = os.getcwd()
            try:
                os.chdir(tmp)
                buf = StringIO()
                with mock.patch("sys.stdout", buf):
                    code = run_cli([f1, f2, "--quiet"])
                self.assertEqual(code, 0)
                xlsx = [n for n in os.listdir(tmp) if n.startswith("Karsilastirma_") and n.endswith(".xlsx")]
                self.assertEqual(len(xlsx), 1)
            finally:
                os.chdir(cwd)

    def test_cli_archive_const_default_path(self):
        import archive_db
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            out = os.path.join(tmp, "r.xlsx")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            before = {r["id"] for r in archive_db.list_comparisons(limit=500)}
            code = run_cli([f1, f2, "-o", out, "--archive", "--quiet"])
            self.assertEqual(code, 0)
            self.assertTrue(os.path.isfile(archive_db.DEFAULT_DB))
            rows = archive_db.list_comparisons(limit=500)
            new_rows = [r for r in rows if r["id"] not in before]
            self.assertGreaterEqual(len(new_rows), 1)
            self.assertTrue(any(r["file1"] == f1 for r in new_rows))

    def test_cli_archive_failure_prints_warning_still_ok(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            out = os.path.join(tmp, "r.xlsx")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            err = StringIO()
            with mock.patch("sys.stderr", err):
                with mock.patch("archive_db.save_comparison", side_effect=RuntimeError("boom")):
                    code = run_cli([f1, f2, "-o", out, "--archive", os.path.join(tmp, "x.db"), "--quiet"])
            self.assertEqual(code, 0)
            self.assertIn("archive warning", err.getvalue())

    def test_cli_batch_missing_folder_exit_1(self):
        with tempfile.TemporaryDirectory() as tmp:
            code = run_cli([os.path.join(tmp, "nope"), tmp, "--batch", "--quiet"])
            self.assertEqual(code, 1)

    def test_cli_unexpected_error_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            out = os.path.join(tmp, "r.xlsx")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            with mock.patch("karsilastir_motor.build_report", side_effect=RuntimeError("explode")):
                code = run_cli([f1, f2, "-o", out, "--quiet"])
            self.assertEqual(code, 2)

    def test_cli_py_entrypoint(self):
        import subprocess
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            out = os.path.join(tmp, "r.xlsx")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            cli = os.path.join(ROOT, "cli.py")
            proc = subprocess.run(
                [sys.executable, cli, f1, f2, "-o", out, "--quiet"],
                capture_output=True, text=True, cwd=ROOT,
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr)
            self.assertTrue(os.path.isfile(out))

    def test_cli_all_supported_langs_ok_message(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            for lang in i18n.SUPPORTED:
                out = os.path.join(tmp, f"r_{lang}.xlsx")
                buf = StringIO()
                with mock.patch("sys.stdout", buf):
                    code = run_cli([f1, f2, "-o", out, "--lang", lang, "--quiet"])
                self.assertEqual(code, 0, msg=lang)
                self.assertTrue(buf.getvalue().strip())


if __name__ == "__main__":
    unittest.main()

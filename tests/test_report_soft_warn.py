# -*- coding: utf-8 -*-
"""Soft-warn path: Excel OK even if HTML/PDF fail."""
import os
import sys
import tempfile
import unittest
from io import StringIO
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(__file__))

import karsilastir_motor as motor
from helpers_datev import drow, write_datev_csv
from karsilastir import run_cli


class TestReportSoftWarn(unittest.TestCase):
    def test_pdf_failure_keeps_excel_and_cli_exit_0(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            out = os.path.join(tmp, "r.xlsx")
            write_datev_csv(f1, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "t")])
            with mock.patch("report_extra.write_pdf_report", side_effect=RuntimeError("pdf boom")):
                result = motor.build_report(
                    f1, f2, out, write_pdf=True, write_html=False, lang="en",
                )
            self.assertTrue(os.path.isfile(out))
            self.assertIsNone(result.get("pdf_path"))
            self.assertTrue(any("pdf:" in w for w in result.get("warnings") or []))

            err = StringIO()
            with mock.patch("sys.stderr", err):
                with mock.patch("report_extra.write_pdf_report", side_effect=RuntimeError("pdf boom")):
                    code = run_cli([f1, f2, "-o", out, "--pdf", "--quiet"])
            self.assertEqual(code, 0)
            self.assertIn("warning:", err.getvalue())
            self.assertIn("pdf", err.getvalue().lower())


if __name__ == "__main__":
    unittest.main()

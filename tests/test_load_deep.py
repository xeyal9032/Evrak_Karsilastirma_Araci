# -*- coding: utf-8 -*-
"""Derin yukleme / format algilama testleri."""
import os
import sys
import tempfile
import unittest
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(__file__))

import karsilastir_motor as motor
from helpers_datev import (
    DATEV_HEADER_LONG, drow, write_datev_csv, write_datev_xlsx, write_journal_xlsx,
)


class TestLoadHeaderVariants(unittest.TestCase):
    def test_long_header_names(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "long.csv")
            write_datev_csv(
                path,
                [drow("21,48", "S", "5900", "70055", "0201", "INV-1", "Barely Digital")],
                header=DATEV_HEADER_LONG,
            )
            e = motor.load(path)
            self.assertEqual(len(e), 1)
            self.assertEqual(e[0]["belegfeld1"], "INV-1")
            self.assertEqual(e[0]["amt"], 21.48)
            self.assertEqual(e[0]["text"], "Barely Digital")

    def test_header_on_first_row_without_extf_meta(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bare.csv")
            # detect needs DATEV header - write without EXTF meta: header at row 0
            import csv
            with open(path, "w", encoding="latin-1", newline="") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow([
                    "Umsatz", "Soll/Haben-Kennzeichen", "WKZ Umsatz", "Kurs", "Basisumsatz",
                    "WKZ Basisumsatz", "Konto", "Gegenkonto", "BU-Schluessel", "Belegdatum",
                    "Belegfeld 1", "Belegfeld 2", "Skonto", "Buchungstext",
                ])
                w.writerow(drow("10,00", "S", "1", "2", "0101", "B1", "T"))
            self.assertEqual(motor.detect_format(path), motor.FORMAT_DATEV_CSV)
            e = motor.load(path)
            self.assertEqual(len(e), 1)
            self.assertEqual(e[0]["amt"], 10.0)


class TestLoadSkips(unittest.TestCase):
    def test_skips_empty_amount_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "a.csv")
            write_datev_csv(path, [
                drow("", "S", "1", "2", "0101", "X", "empty"),
                drow("10,00", "S", "1", "2", "0101", "Y", "ok"),
            ])
            e = motor.load(path)
            self.assertEqual(len(e), 1)
            self.assertEqual(e[0]["belegfeld1"], "Y")

    def test_journal_keeps_zero_skips_missing_accounts(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "j.xlsx")
            write_journal_xlsx(path, [
                (datetime(2026, 1, 1), "Z0", "zero", 0, 1000, 2000),
                (datetime(2026, 1, 2), "Z1", "nosoll", 50, None, 2000),
                (datetime(2026, 1, 3), "OK", "good", 50.5, 1000, 2000),
            ])
            e = motor.load(path)
            self.assertEqual(len(e), 2)
            belegs = {x["belegfeld1"] for x in e}
            self.assertEqual(belegs, {"Z0", "OK"})
            zero = next(x for x in e if x["belegfeld1"] == "Z0")
            self.assertEqual(zero["amt"], 0.0)
            self.assertEqual(zero["umsatz"], "0,00")


class TestLoadEncoding(unittest.TestCase):
    def test_cp1252_umlaut_preserved(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "uml.csv")
            text = "Miete März Talweg"
            write_datev_csv(
                path,
                [drow("985,00", "S", "6310", "1801", "0203", "M1", text)],
                encoding="cp1252",
            )
            e = motor.load(path)
            self.assertEqual(e[0]["text"], text)

    def test_utf8_sig_preserved(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "u8.csv")
            text = "Straße / Größe"
            write_datev_csv(
                path,
                [drow("1,00", "S", "1", "2", "0101", "U8", text)],
                encoding="utf-8-sig",
            )
            e = motor.load(path)
            self.assertEqual(e[0]["text"], text)


class TestDetectErrors(unittest.TestCase):
    def test_missing_file(self):
        with self.assertRaises(FileNotFoundError):
            motor.detect_format(r"C:\no\such\file.csv")

    def test_empty_csv(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "empty.csv")
            open(path, "wb").close()
            with self.assertRaises(motor.FormatError):
                motor.load(path)

    def test_journal_vs_datev_xlsx_detection(self):
        with tempfile.TemporaryDirectory() as tmp:
            j = os.path.join(tmp, "j.xlsx")
            d = os.path.join(tmp, "d.xlsx")
            write_journal_xlsx(j, [(datetime(2026, 1, 1), "1", "t", 10, 1, 2)])
            write_datev_xlsx(d, [drow("10,00", "S", "1", "2", "0101", "1", "t")])
            self.assertEqual(motor.detect_format(j), motor.FORMAT_JOURNAL_XLSX)
            self.assertEqual(motor.detect_format(d), motor.FORMAT_DATEV_XLSX)


class TestFindColumns(unittest.TestCase):
    def test_missing_required_column(self):
        with self.assertRaises(motor.FormatError):
            motor.find_columns(["Umsatz", "Konto"], fpath="x")

    def test_gegenkonto_prefix_match(self):
        header = [
            "Umsatz", "Soll/Haben-Kennzeichen", "Konto",
            "Gegenkonto (ohne BU-Schluessel XYZ)", "Belegdatum", "Belegfeld 1", "Buchungstext",
        ]
        idx = motor.find_columns(header, fpath="x")
        self.assertIn("gegenkonto", idx)


if __name__ == "__main__":
    unittest.main()

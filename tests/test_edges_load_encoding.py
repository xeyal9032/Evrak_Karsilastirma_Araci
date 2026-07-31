# -*- coding: utf-8 -*-
"""Yukleme / encoding / bilimsel beleg / kopya satir kenar durumlari."""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import karsilastir_motor as motor
from tests.helpers_datev import DATEV_HEADER, drow, write_datev_csv


class TestEmptyAndShortFiles(unittest.TestCase):
    def test_truly_empty_file_raises(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "empty.csv")
            open(path, "wb").close()
            with self.assertRaises(motor.FormatError):
                motor.load(path)

    def test_csv_too_short_one_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "one.csv")
            with open(path, "w", encoding="latin-1") as f:
                f.write("EXTF;700;21\n")
            with self.assertRaises(motor.FormatError):
                motor.load(path)

    def test_header_only_datev_csv_raises_format_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "hdr.csv")
            write_datev_csv(path, [])
            with self.assertRaises(motor.FormatError) as cm:
                motor.load(path)
            msg = str(cm.exception).lower()
            self.assertTrue("satir" in msg or "строк" in msg or "row" in msg or "veri" in msg or "empty" in msg or "yok" in msg or "нет" in msg or "kayit" in msg)


class TestEncoding(unittest.TestCase):
    def test_invalid_utf8_falls_to_cp1252_or_latin1(self):
        """utf-8 olmayan baytlar (umlaut) yine de okunmali."""
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "cp.csv")
            header = ";".join(DATEV_HEADER)
            # 14 DATEV alan: ... Belegfeld1; Belegfeld2; Skonto; Buchungstext
            data = "10,00;S;;;;;4400;1800;;0101;B1;;;Buchung mit \xf6\r\n"
            lines = [
                b"EXTF;700;21;Buchungsstapel;test\r\n",
                (header + "\r\n").encode("latin-1"),
                data.encode("latin-1"),
            ]
            with open(path, "wb") as f:
                f.writelines(lines)
            entries = motor.load(path)
            self.assertEqual(len(entries), 1)
            self.assertIn("Buchung", entries[0]["text"])


class TestCorruptXlsx(unittest.TestCase):
    def test_corrupt_xlsx_raises_format_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bad.xlsx")
            with open(path, "wb") as f:
                f.write(b"PK\x03\x04not-a-real-workbook")
            with self.assertRaises(motor.FormatError):
                motor.detect_format(path)


class TestScientificBelegEdges(unittest.TestCase):
    def test_beleg_scientific_lowercase_e_minus(self):
        self.assertEqual(motor.norm_belegfeld("1.2e-3"), "")
        self.assertEqual(motor.norm_belegfeld("+1,5E+10"), "")
        self.assertEqual(motor.norm_belegfeld("1e+11"), "")

    def test_beleg_re2026_not_treated_as_scientific(self):
        self.assertEqual(motor.norm_belegfeld("RE2026-0363"), "RE2026-0363")
        self.assertTrue(motor.norm_belegfeld("E2026ABC"))
        self.assertNotEqual(motor.norm_belegfeld("K273-933848/1"), "")

    def test_beleg_float_scientific_string_empties_key1(self):
        e = motor.make_entry(1, "10,00", "S", "1", "2", "0101", "1.22E+11", "t", 10.0)
        self.assertEqual(e["beleg_n"], "")
        self.assertIsNone(motor.key1(e))

    def test_norm_amount_scientific_string(self):
        # Motor Alman/CSV format bekler; bilimsel string tutarli sekilde None veya sayi
        val = motor.norm_amount("1,5E+2")
        self.assertTrue(val is None or abs(val - 150.0) < 0.01 or abs(val - 1.5) < 0.01)


class TestDuplicateRows(unittest.TestCase):
    def test_duplicate_rows_identical_files_all_match(self):
        row = drow("1,00", "S", "1", "2", "0101", "D1", "dup")
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [row] * 4)
            write_datev_csv(f2, [row] * 4)
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(all(e["matched"] for e in e1 + e2))
            self.assertEqual(sum(1 for e in e1 if e.get("pair_kind") == "MATCH"), 4)

    def test_duplicate_rows_one_side_extra(self):
        row = drow("2,00", "S", "1", "2", "0101", "D2", "dup2")
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [row] * 3)
            write_datev_csv(f2, [row] * 2)
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertEqual(sum(1 for e in e1 if e["matched"] and e.get("pair")), 2)
            self.assertEqual(sum(1 for e in e1 if not e["matched"]), 1)


if __name__ == "__main__":
    unittest.main()

# -*- coding: utf-8 -*-
"""Eslestirme katmanlari: yanlis pozitif/negatif olmamali."""
import csv
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import karsilastir_motor as motor


DATEV_HEADER = [
    "Umsatz", "Soll/Haben-Kennzeichen", "WKZ Umsatz", "Kurs", "Basisumsatz", "WKZ Basisumsatz",
    "Konto", "Gegenkonto", "BU-Schluessel", "Belegdatum", "Belegfeld 1", "Belegfeld 2",
    "Skonto", "Buchungstext",
]


def write_datev(path, data_rows):
    with open(path, "w", encoding="latin-1", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["EXTF", "700", "21", "Buchungsstapel", "test"])
        w.writerow(DATEV_HEADER)
        for row in data_rows:
            w.writerow(row)


def row(umsatz, sh, konto, gegen, datum, beleg, text):
    return [umsatz, sh, "", "", "", "", konto, gegen, "", datum, beleg, "", "", text]


class TestExactMatch(unittest.TestCase):
    def test_identical_files_all_match(self):
        rows = [
            row("100,00", "S", "1200", "8400", "1501", "RE-A", "Miete"),
            row("50,50", "H", "8400", "1200", "1601", "RE-B", "Waren"),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, rows)
            write_datev(f2, rows)
            e1, e2, n1, n2 = motor.compare(f1, f2)
            self.assertEqual(n1, 0)
            self.assertEqual(n2, 0)
            self.assertTrue(all(e["matched"] and e["pair_kind"] == "MATCH" for e in e1))
            self.assertTrue(all(e["matched"] and e["pair_kind"] == "MATCH" for e in e2))
            # Harf/rakam birebir
            for a in e1:
                b = a["pair"]
                self.assertEqual(a["amt"], b["amt"])
                self.assertEqual(a["belegfeld1"], b["belegfeld1"])
                self.assertEqual(a["text"], b["text"])
                self.assertEqual(a["konto"], b["konto"])
                self.assertEqual(a["datum"], b["datum"])


class TestValueMismatch(unittest.TestCase):
    def test_same_id_different_amount_is_orange(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, [row("100,00", "S", "1200", "8400", "1501", "RE-X", "Test")])
            write_datev(f2, [row("200,00", "S", "1200", "8400", "1501", "RE-X", "Test")])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertEqual(e1[0]["pair_kind"], "MISMATCH")
            self.assertEqual(e1[0]["amt"], 100.0)
            self.assertEqual(e1[0]["pair"]["amt"], 200.0)
            self.assertEqual(e1[0]["belegfeld1"], "RE-X")
            self.assertEqual(e1[0]["pair"]["belegfeld1"], "RE-X")


class TestOnlyOneSide(unittest.TestCase):
    def test_only1_and_only2(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, [row("10,00", "S", "1", "2", "0101", "ONLY1", "A")])
            write_datev(f2, [row("20,00", "S", "3", "4", "0202", "ONLY2", "B")])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertFalse(e1[0]["matched"])
            self.assertFalse(e2[0]["matched"])
            self.assertEqual(e1[0]["belegfeld1"], "ONLY1")
            self.assertEqual(e2[0]["belegfeld1"], "ONLY2")


class TestStornoNetting(unittest.TestCase):
    def test_storno_pair_netted_not_real_diff(self):
        rows = [
            row("100,00", "S", "2000", "2800", "0503", "RE-S", "Original Buchung"),
            row("100,00", "S", "2800", "2000", "0503", "RE-S", "<Storno> Original Buchung"),
            row("50,00", "S", "1000", "2000", "0603", "RE-OK", "Normal"),
        ]
        twin = [
            row("50,00", "S", "1000", "2000", "0603", "RE-OK", "Normal"),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, rows)
            write_datev(f2, twin)
            e1, e2, n1, n2 = motor.compare(f1, f2)
            self.assertEqual(n1, 1)
            self.assertEqual(n2, 0)
            netted = [e for e in e1 if e["netted"]]
            self.assertEqual(len(netted), 2)
            ok = [e for e in e1 if e["belegfeld1"] == "RE-OK"][0]
            self.assertTrue(ok["matched"])
            self.assertEqual(ok["pair_kind"], "MATCH")


class TestShOrientation(unittest.TestCase):
    def test_datev_h_matches_equivalent_s(self):
        # Ayni ekonomik islem: H/8400/1200 == S/1200/8400
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, [row("250,50", "H", "8400", "1200", "2001", "RE-2002", "Wareneinkauf")])
            write_datev(f2, [row("250,50", "S", "1200", "8400", "2001", "RE-2002", "Wareneinkauf")])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(e1[0]["matched"])
            self.assertEqual(e1[0]["pair_kind"], "MATCH")
            self.assertEqual(e1[0]["amt"], e1[0]["pair"]["amt"])
            self.assertEqual(e1[0]["belegfeld1"], e1[0]["pair"]["belegfeld1"])


class TestNoFalsePositiveStrict(unittest.TestCase):
    def test_generic_text_different_beleg_not_merged_on_value_tier(self):
        # Ayni genel metin + ayni konto/tarih ama farkli beleg ve tutar -> eslesmemeli
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, [
                row("100,00", "S", "1000", "2000", "0101", "A1", "Zahlungseing.: diverse Rechnungen"),
                row("200,00", "S", "1000", "2000", "0101", "A2", "Zahlungseing.: diverse Rechnungen"),
            ])
            write_datev(f2, [
                row("150,00", "S", "1000", "2000", "0101", "B1", "Zahlungseing.: diverse Rechnungen"),
                row("250,00", "S", "1000", "2000", "0101", "B2", "Zahlungseing.: diverse Rechnungen"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            # Bos beleg degil - farkli belegler; tutarlar farkli -> MATCH olmamali
            # key_value2 strict: ayni text+konto+tarih birden fazla aday -> eslesmez
            unmatched1 = sum(1 for e in e1 if not e["matched"])
            unmatched2 = sum(1 for e in e2 if not e["matched"])
            self.assertEqual(unmatched1, 2)
            self.assertEqual(unmatched2, 2)

    def test_empty_belegfeld_no_value_mismatch(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, [row("100,00", "S", "1000", "2000", "0101", "", "Alpha")])
            write_datev(f2, [row("200,00", "S", "1000", "2000", "0101", "", "Beta")])
            e1, e2, *_ = motor.compare(f1, f2)
            # key_value1 beleg bos -> None; farkli metin -> MATCH/MISMATCH beleg ile olmamali
            self.assertNotEqual(e1[0].get("pair_kind"), "MISMATCH")


class TestKontoRecode(unittest.TestCase):
    def test_same_text_amount_date_different_konto_matches(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, [row("500,00", "S", "4336", "1000", "1003", "FS001", "BIV Group Trans FS/26/01/2026")])
            write_datev(f2, [row("500,00", "S", "4339", "1000", "1003", "FS001", "BIV Group Trans FS/26/01/2026")])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(e1[0]["matched"])
            self.assertEqual(e1[0]["pair_kind"], "MATCH")
            self.assertEqual(e1[0]["amt"], 500.0)
            self.assertEqual(e1[0]["pair"]["amt"], 500.0)


class TestSubstringBelegInText(unittest.TestCase):
    def test_beleg_moved_into_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev(f1, [row("300,00", "S", "5000", "1000", "1203", "FS/26/01/2026", "BIV Group")])
            write_datev(f2, [row("300,00", "S", "5100", "1000", "1203", "", "BIV Group Trans FS/26/01/2026 extra")])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(e1[0]["matched"])
            self.assertEqual(e1[0]["pair_kind"], "MATCH")


class TestDiffNote(unittest.TestCase):
    def test_diff_note_lists_changed_fields_exactly(self):
        a = motor.make_entry(1, "100,00", "S", "1", "2", "0101", "X", "Hello", 100.0)
        b = motor.make_entry(2, "200,00", "S", "1", "2", "0101", "X", "Hello", 200.0)
        note = motor.diff_note(a, b)
        self.assertIn("100,00", note)
        self.assertIn("200,00", note)
        self.assertIn("Umsatz", note)


if __name__ == "__main__":
    unittest.main()

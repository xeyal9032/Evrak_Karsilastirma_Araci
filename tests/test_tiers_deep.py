# -*- coding: utf-8 -*-
"""Derin eslestirme katmanlari: klon, capraz alan, supheli, EB-Wert."""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(__file__))

import karsilastir_motor as motor
from helpers_datev import drow, write_datev_csv


class TestCloneRound(unittest.TestCase):
    def test_identical_duplicates_match_in_clone_pass(self):
        # Ayni anahtar, ayirt edilemez klonlar - ikinci turda eslesmeli
        rows = [
            drow("100,00", "S", "1200", "8400", "1501", "DUP", "Same Text"),
            drow("100,00", "S", "1200", "8400", "1501", "DUP", "Same Text"),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, rows)
            write_datev_csv(f2, rows)
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(all(e["matched"] for e in e1))
            self.assertTrue(all(e["pair_kind"] == "MATCH" for e in e1))
            self.assertTrue(any("klon" in (e["tier"] or "").lower() for e in e1))


class TestCrossFieldMismatch(unittest.TestCase):
    def test_belegfeld_swapped_with_text(self):
        # Dosya1: Belegfeld1=RE2026-0363, Dosya2: ayni kimlik Buchungstext'te, farkli tutar
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("1900,00", "S", "4400", "1000", "1003", "RE2026-0363", "STELLANTIS"),
            ])
            write_datev_csv(f2, [
                drow("1800,00", "S", "4400", "1000", "1003", "", "RE2026-0363"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(e1[0]["matched"])
            self.assertEqual(e1[0]["pair_kind"], "MISMATCH")
            self.assertEqual(e1[0]["amt"], 1900.0)
            self.assertEqual(e1[0]["pair"]["amt"], 1800.0)


class TestSuspiciousMatch(unittest.TestCase):
    def test_amount_konto_date_only_is_mismatch(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("77,00", "S", "5000", "1000", "0505", "AAA", "Alpha Unique"),
            ])
            write_datev_csv(f2, [
                drow("77,00", "S", "5000", "1000", "0505", "BBB", "Beta Unique"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(e1[0]["matched"])
            self.assertEqual(e1[0]["pair_kind"], "MISMATCH")
            self.assertIn("SUPHELI", e1[0]["tier"].upper())


class TestEbWertSkippedInBalance(unittest.TestCase):
    def test_eb_wert_excluded_from_account_diff(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            # Sadece bir tarafta EB-Wert - bakiye farkina girmemeli
            write_datev_csv(f1, [
                drow("29468,00", "S", "135", "9000", "0101", "1", "EB-Wert"),
                drow("100,00", "S", "1200", "8400", "0201", "R1", "Normal"),
            ])
            write_datev_csv(f2, [
                drow("100,00", "S", "1200", "8400", "0201", "R1", "Normal"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            diffs = motor.account_balance_diff(e1, e2)
            kontos = {d[0] for d in diffs}
            self.assertNotIn("135", kontos)
            self.assertNotIn("9000", kontos)

    def test_gewinnvortrag_and_zahllast_skipped(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("500,00", "S", "9000", "8000", "0101", "G", "Gewinnvortrag 2025"),
                drow("200,00", "H", "1800", "4000", "0101", "Z", "Zahllast USt"),
            ])
            write_datev_csv(f2, [])
            # empty f2 invalid - need at least one readable row
            write_datev_csv(f2, [
                drow("1,00", "S", "1111", "2222", "0303", "X", "KeepAlive"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            diffs = motor.account_balance_diff(e1, e2)
            # 9000/8000/1800/4000 EB-benzeri atlanir; 1111/2222 farkli kalabilir
            kontos = {d[0] for d in diffs}
            self.assertNotIn("9000", kontos)
            self.assertNotIn("8000", kontos)


class TestAccountBalanceSign(unittest.TestCase):
    def test_soll_haben_signs(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [drow("100,00", "S", "1000", "2000", "0101", "A", "T")])
            write_datev_csv(f2, [drow("100,00", "S", "1000", "2000", "0101", "A", "T")])
            e1, e2, *_ = motor.compare(f1, f2)
            diffs = motor.account_balance_diff(e1, e2)
            self.assertEqual(diffs, [])


class TestCombinedGroups(unittest.TestCase):
    def test_group_kinds_and_sort(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("10,00", "S", "1", "2", "1502", "M", "match"),
                drow("20,00", "S", "1", "2", "1002", "D", "diffamt"),
                drow("30,00", "S", "3", "4", "0101", "O1", "only1"),
            ])
            write_datev_csv(f2, [
                drow("10,00", "S", "1", "2", "1502", "M", "match"),
                drow("25,00", "S", "1", "2", "1002", "D", "diffamt"),
                drow("40,00", "S", "5", "6", "2003", "O2", "only2"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            groups = motor.build_combined_rows(e1, e2)
            kinds = [g["kind"] for g in groups]
            self.assertEqual(kinds.count("MATCH"), 1)
            self.assertEqual(kinds.count("MISMATCH"), 1)
            self.assertEqual(kinds.count("ONLY1"), 1)
            self.assertEqual(kinds.count("ONLY2"), 1)
            # Kronolojik: 0101, 1002, 1502, 2003
            dates = []
            for g in groups:
                ref = g["e1"] or g["e2"]
                dates.append(ref["datum"])
            self.assertEqual(dates, ["0101", "1002", "1502", "2003"])


class TestFillColors(unittest.TestCase):
    def test_fill_for_kinds(self):
        e = motor.make_entry(1, "1,00", "S", "1", "2", "0101", "X", "t", 1.0)
        e["matched"] = False
        self.assertEqual(motor.fill_for(e).start_color.rgb, motor.RED.start_color.rgb)

        e["matched"] = True
        e["tier"] = "IC-NETLESME (dosya icinde)"
        e["pair_kind"] = "MATCH"
        self.assertEqual(motor.fill_for(e).start_color.rgb, motor.GREEN.start_color.rgb)

        e["tier"] = "TAM ESLESME"
        e["pair_kind"] = "MATCH"
        self.assertEqual(motor.fill_for(e).start_color.rgb, motor.YELLOW.start_color.rgb)

        e["pair_kind"] = "MISMATCH"
        self.assertEqual(motor.fill_for(e).start_color.rgb, motor.ORANGE.start_color.rgb)


if __name__ == "__main__":
    unittest.main()

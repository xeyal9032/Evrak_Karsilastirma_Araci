# -*- coding: utf-8 -*-
"""Ic-netlesme, diff_note ve birlesik satir tutarliligi."""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import karsilastir_motor as motor
import report_extra
from tests.helpers_datev import drow, write_datev_csv


class TestNetting(unittest.TestCase):
    def test_netting_marks_matched_with_ic_netlesme_tier(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            # Twin + storno (gegen/konto swapped), same amt/sh/datum
            write_datev_csv(f1, [
                drow("100,00", "S", "4400", "1800", "0101", "T1", "Normal booking"),
                drow("100,00", "S", "1800", "4400", "0101", "T1", "<Storno> Normal booking"),
                drow("5,00", "S", "1", "2", "0101", "KEEP", "keep me"),
            ])
            write_datev_csv(f2, [
                drow("5,00", "S", "1", "2", "0101", "KEEP", "keep me"),
            ])
            e1, e2, net1, net2 = motor.compare(f1, f2)[:4]
            self.assertGreaterEqual(net1, 1)
            netted = [e for e in e1 if e.get("netted")]
            self.assertEqual(len(netted), 2)
            for e in netted:
                self.assertTrue(e["matched"])
                self.assertIn("IC-NETLESME", e["tier"] or "")

    def test_netted_excluded_from_build_combined_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("100,00", "S", "4400", "1800", "0101", "T1", "Normal booking"),
                drow("100,00", "S", "1800", "4400", "0101", "T1", "<Storno> Normal booking"),
                drow("5,00", "S", "1", "2", "0101", "KEEP", "keep me"),
            ])
            write_datev_csv(f2, [
                drow("5,00", "S", "1", "2", "0101", "KEEP", "keep me"),
            ])
            e1, e2, net1, *_ = motor.compare(f1, f2)
            groups = motor.build_combined_rows(e1, e2)
            texts = []
            for g in groups:
                if g.get("e1"):
                    texts.append(g["e1"]["text"])
                if g.get("e2"):
                    texts.append(g["e2"]["text"])
            self.assertTrue(any("keep me" in t for t in texts))
            self.assertFalse(any("<Storno>" in t for t in texts))
            self.assertFalse(any(t == "Normal booking" for t in texts))
            self.assertGreaterEqual(net1, 1)

    def test_storno_picks_first_non_netted_twin(self):
        entries = [
            motor.make_entry(1, "10,00", "S", "1", "2", "0101", "A", "twin", 10.0),
            motor.make_entry(2, "10,00", "S", "1", "2", "0101", "B", "twin2", 10.0),
            motor.make_entry(3, "10,00", "S", "2", "1", "0101", "C", "<Storno> cancel", 10.0),
        ]
        n = motor.net_self_cancelling(entries)
        self.assertEqual(n, 1)
        self.assertTrue(entries[0]["netted"])
        self.assertTrue(entries[2]["netted"])
        self.assertFalse(entries[1]["netted"])


class TestDiffNote(unittest.TestCase):
    def test_diff_note_ignores_sh_and_datum(self):
        a = motor.make_entry(1, "10,00", "S", "1", "2", "0101", "B", "t", 10.0)
        b = motor.make_entry(2, "10,00", "H", "1", "2", "0202", "B", "t", 10.0)
        # Same umsatz/konto/gegen/beleg/text — SH and datum differ but not in diff_note fields
        self.assertEqual(motor.diff_note(a, b), "")

    def test_diff_note_gegenkonto_only(self):
        a = motor.make_entry(1, "10,00", "S", "1", "2", "0101", "B", "t", 10.0)
        b = motor.make_entry(2, "10,00", "S", "1", "9", "0101", "B", "t", 10.0)
        note = motor.diff_note(a, b)
        self.assertIn("Gegenkonto", note)
        self.assertIn("'2'", note)
        self.assertIn("'9'", note)

    def test_mismatch_note_in_html_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [drow("10,00", "S", "1", "2", "0101", "D", "same")])
            write_datev_csv(f2, [drow("99,00", "S", "1", "2", "0101", "D", "same")])
            e1, e2, *_ = motor.compare(f1, f2)
            groups = motor.build_combined_rows(e1, e2)
            rows = report_extra.groups_to_rows(groups, "A", "B")
            mism = [r for r in rows if r["status"] == "MISMATCH"]
            self.assertEqual(len(mism), 2)
            expected = motor.diff_note(e1[0], e2[0])
            self.assertEqual(mism[1]["note"], expected)
            self.assertIn("Umsatz", expected)


if __name__ == "__main__":
    unittest.main()

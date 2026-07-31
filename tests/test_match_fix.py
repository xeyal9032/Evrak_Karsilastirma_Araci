# -*- coding: utf-8 -*-
"""Regression tests for scientific Belegfeld and text ref-id matching."""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import karsilastir_motor as motor
from tests.helpers_datev import drow, write_datev_csv


class TestNormBelegScientific(unittest.TestCase):
    def test_scientific_german_and_english_rejected(self):
        self.assertEqual(motor.norm_belegfeld("1,22001E+11"), "")
        self.assertEqual(motor.norm_belegfeld("1.22001E+11"), "")
        self.assertEqual(motor.norm_belegfeld("122001E+11"), "")
        self.assertEqual(motor.norm_belegfeld("6,74002E+11"), "")
        self.assertEqual(motor.norm_belegfeld("122E+11"), "")

    def test_normal_beleg_kept(self):
        self.assertEqual(motor.norm_belegfeld("RE2026-0363"), "RE2026-0363")
        self.assertEqual(motor.norm_belegfeld("K273-933848/1"), "K273-933848/1")
        self.assertEqual(motor.norm_belegfeld("2196"), "2196")

    def test_key1_none_when_scientific(self):
        e = motor.make_entry(1, "10,00", "S", "1", "2", "0101", "1,22001E+11", "t", 10.0)
        self.assertEqual(e["beleg_n"], "")
        self.assertIsNone(motor.key1(e))


class TestExtractRefIds(unittest.TestCase):
    def test_strips_leading_zeros(self):
        ids = motor.extract_ref_ids("Zahlungsausg.: 0000119564441 Vodafone")
        self.assertIn("119564441", ids)

    def test_does_not_glue_adjacent_numbers(self):
        ids = motor.extract_ref_ids("0000119564441 0020744324955 Vodafone")
        self.assertIn("119564441", ids)
        self.assertIn("20744324955", ids)

    def test_space_in_beleg_still_matches_compact(self):
        a = motor.extract_ref_ids("2026 130000011")
        b = motor.extract_ref_ids("2026130000011")
        self.assertTrue(a & b)

    def test_from_text_and_beleg(self):
        ids = motor.extract_ref_ids("RE-1001", "Vodafone 119564441 Telekommunikation")
        self.assertIn("119564441", ids)


class TestScientificFalsePositiveFixed(unittest.TestCase):
    def test_different_vodafone_invoices_not_merged_by_sci_beleg(self):
        """Ayni bozuk Belegfeld + ayni tutar ama FARKLI fatura no -> eslesmemeli."""
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("415,39", "S", "6805", "70004", "1301", "122001E+11", "Vodafone 119316577 Telekommunikation"),
                drow("415,39", "S", "70004", "1801", "2301", "122001E+11", "Zahlungsausg.: 0000119564441 Vodafone Rechnung"),
            ])
            write_datev_csv(f2, [
                drow("415,39", "S", "6805", "70026", "1301", "1,22001E+11", "Vodafone 119564441 Telekommunikation"),
                drow("415,39", "S", "70026", "1801", "2301", "1,22001E+11", "Zahlungsausg.: Vodafone 119564441 Telekommunikation"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            # Expense rows have different invoice numbers -> must NOT match each other
            exp1 = [e for e in e1 if e["konto"] == "6805"][0]
            exp2 = [e for e in e2 if e["konto"] == "6805"][0]
            self.assertFalse(exp1["matched"] and exp1.get("pair") is exp2)
            self.assertFalse(exp2["matched"] and exp2.get("pair") is exp1)
            # Payment rows share 119564441 -> SHOULD match despite konto 70004 vs 70026
            pay1 = [e for e in e1 if e["konto"] == "70004"][0]
            pay2 = [e for e in e2 if e["konto"] == "70026"][0]
            self.assertTrue(pay1["matched"], msg=pay1)
            self.assertTrue(pay2["matched"], msg=pay2)
            self.assertIs(pay1["pair"], pay2)
            self.assertEqual(pay1["pair_kind"], "MATCH")
            self.assertIn("REF-ID", pay1["tier"] or "")


class TestSharedRefIdTier(unittest.TestCase):
    def test_konto_recode_with_shared_invoice_in_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("100,00", "S", "70001", "1801", "1501", "X", "Zahlung 99887766 Lieferant A"),
            ])
            write_datev_csv(f2, [
                drow("100,00", "S", "70099", "1801", "1501", "Y", "Zahlung Lieferant A 99887766"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(e1[0]["matched"])
            self.assertIs(e1[0]["pair"], e2[0])


if __name__ == "__main__":
    unittest.main()

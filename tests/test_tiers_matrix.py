# -*- coding: utf-8 -*-
"""Derin eslesme katmani matrisi: key1-4, klon, substring, ref-id, H/S."""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import karsilastir_motor as motor
from tests.helpers_datev import drow, write_datev_csv


def _cmp(rows1, rows2):
    with tempfile.TemporaryDirectory() as tmp:
        f1 = os.path.join(tmp, "a.csv")
        f2 = os.path.join(tmp, "b.csv")
        write_datev_csv(f1, rows1)
        write_datev_csv(f2, rows2)
        return motor.compare(f1, f2)


class TestTierPrecedence(unittest.TestCase):
    def test_key1_beats_key2_when_beleg_present(self):
        e1, e2, *_ = _cmp(
            [drow("10,00", "S", "1", "2", "0101", "BELEG99", "metin-a")],
            [drow("10,00", "S", "1", "2", "0101", "BELEG99", "metin-b")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertEqual(e1[0]["pair_kind"], "MATCH")
        self.assertIn("TAM ESLESME", e1[0]["tier"])

    def test_key2_before_key3_when_konto_same(self):
        e1, e2, *_ = _cmp(
            [drow("25,50", "S", "4400", "1800", "1501", "", "Ayni metin XYZ")],
            [drow("25,50", "S", "4400", "1800", "1501", "DIF", "Ayni metin XYZ")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertIn("Konto+Tutar+Metin", e1[0]["tier"])
        self.assertNotIn("Yeniden-kodlanmis", e1[0]["tier"])

    def test_key3_konto_recode_tier_name(self):
        e1, e2, *_ = _cmp(
            [drow("100,00", "S", "70001", "1801", "1501", "A", "Ortak aciklama 123")],
            [drow("100,00", "S", "70099", "1801", "1501", "B", "Ortak aciklama 123")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertEqual(e1[0]["pair_kind"], "MATCH")
        self.assertIn("Yeniden-kodlanmis", e1[0]["tier"])
        self.assertIs(e1[0]["pair"], e2[0])


class TestCanonSH(unittest.TestCase):
    def test_h_flipped_matches_s_via_canon(self):
        """DATEV H (ters kontolar) <-> Journal/DATEV S ayni ekonomik islem."""
        e1, e2, *_ = _cmp(
            [drow("50,00", "H", "1801", "6805", "2001", "INV50", "Benzin")],
            [drow("50,00", "S", "6805", "1801", "2001", "INV50", "Benzin")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertIs(e1[0]["pair"], e2[0])
        self.assertEqual(e1[0]["pair_kind"], "MATCH")


class TestClonePass(unittest.TestCase):
    def test_clone_unequal_counts_pairs_min(self):
        row = drow("5,00", "S", "1", "2", "0101", "DUP", "klon metin")
        e1, e2, *_ = _cmp([row, row, row], [row, row])
        matched1 = sum(1 for e in e1 if e["matched"] and e.get("pair_kind") == "MATCH")
        matched2 = sum(1 for e in e2 if e["matched"] and e.get("pair_kind") == "MATCH")
        self.assertEqual(matched1, 2)
        self.assertEqual(matched2, 2)
        self.assertEqual(sum(1 for e in e1 if not e["matched"]), 1)
        self.assertTrue(any("klon" in (e["tier"] or "") for e in e1 if e["matched"]))

    def test_clone_skips_disjoint_ref_ids(self):
        """Ayni beleg/tutar ama farkli fatura no (klon gecisi atlamali)."""
        a = drow("80,00", "S", "6805", "70001", "1301", "SCI", "Fatura 11111111 Alpha")
        b = drow("80,00", "S", "6805", "70001", "1301", "SCI", "Fatura 22222222 Beta")
        # Bilimsel beleg ile key1 kapali; key2/key3 metin farkli -> eslesmemeli
        # Ayni key4 icin strict SUPHELI tek aday ister; burda iki farkli metin
        e1, e2, *_ = _cmp(
            [drow("80,00", "S", "6805", "70001", "1301", "X1", "Fatura 11111111 Alpha"),
             drow("80,00", "S", "6805", "70001", "1301", "X1", "Fatura 11111111 Alpha")],
            [drow("80,00", "S", "6805", "70001", "1301", "X1", "Fatura 22222222 Beta"),
             drow("80,00", "S", "6805", "70001", "1301", "X1", "Fatura 22222222 Beta")],
        )
        # key1 ayni (X1) ama metin farkli -> clone _all_identical her tarafta True,
        # ref_ids disjoint -> eslesme yok
        for e in e1 + e2:
            if e["matched"] and e.get("pair"):
                ia = e.get("ref_ids") or set()
                ib = e["pair"].get("ref_ids") or set()
                if ia and ib:
                    self.assertFalse(ia.isdisjoint(ib), msg=(e["text"], e["pair"]["text"]))

    def test_clone_pass_key2_identical_text_dupes(self):
        row = drow("12,00", "S", "3", "4", "0505", "", "Bos beleg ayni metin")
        e1, e2, *_ = _cmp([row, row], [row, row])
        self.assertTrue(all(e["matched"] for e in e1 + e2))
        self.assertTrue(all("klon" in (e["tier"] or "") for e in e1))


class TestSharedRefAndSubstring(unittest.TestCase):
    def test_shared_ref_ambiguous_two_hits_no_match(self):
        e1, e2, *_ = _cmp(
            [drow("40,00", "S", "1", "2", "1001", "A", "Pay 99887766 only")],
            [
                drow("40,00", "S", "9", "8", "1001", "B", "Pay 99887766 left"),
                drow("40,00", "S", "7", "6", "1001", "C", "Pay 99887766 right"),
            ],
        )
        self.assertFalse(e1[0]["matched"])

    def test_shared_ref_min_len_boundary(self):
        self.assertEqual(motor.extract_ref_ids("id 123456", min_len=7), set())
        self.assertIn("1234567", motor.extract_ref_ids("id 1234567", min_len=7))
        e1, e2, *_ = _cmp(
            [drow("15,00", "S", "10", "20", "0101", "P", "Ref 1234567 Vendor")],
            [drow("15,00", "S", "99", "20", "0101", "Q", "Vendor Ref 1234567")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertIn("REF-ID", e1[0]["tier"] or "")

    def test_substring_min_len_5_boundary(self):
        # beleg len 4 -> no substring match; len 5 -> match when in text
        e1, e2, *_ = _cmp(
            [drow("7,00", "S", "1", "2", "0101", "ABCD", "plain")],
            [drow("7,00", "S", "9", "2", "0101", "Z", "contains ABCD here")],
        )
        self.assertFalse(e1[0]["matched"])

        e1, e2, *_ = _cmp(
            [drow("7,00", "S", "1", "2", "0101", "ABCDE", "plain")],
            [drow("7,00", "S", "9", "2", "0101", "Z", "contains ABCDE here")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertIn("Belegfeld1 metne", e1[0]["tier"] or "")

    def test_substring_ref_id_conflict_blocks(self):
        e1, e2, *_ = _cmp(
            [drow("9,00", "S", "1", "2", "0101", "FS2601", "Fatura 11111111")],
            [drow("9,00", "S", "9", "2", "0101", "ZZ", "FS2601 and 22222222")],
        )
        # Substring would hit FS2601, but disjoint long ids -> block
        self.assertFalse(e1[0]["matched"])


class TestValueAndSupheli(unittest.TestCase):
    def test_value1_amount_diff_is_mismatch(self):
        e1, e2, *_ = _cmp(
            [drow("10,00", "S", "1", "2", "0101", "RE100", "t")],
            [drow("99,00", "S", "1", "2", "0101", "RE100", "t")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertEqual(e1[0]["pair_kind"], "MISMATCH")
        self.assertIn("DEGER-FARKI", e1[0]["tier"])

    def test_supheli_only_when_earlier_tiers_fail(self):
        e1, e2, *_ = _cmp(
            [drow("33,00", "S", "5", "6", "0808", "AAA", "Metin bir")],
            [drow("33,00", "S", "5", "6", "0808", "BBB", "Metin iki")],
        )
        self.assertTrue(e1[0]["matched"])
        self.assertEqual(e1[0]["pair_kind"], "MISMATCH")
        self.assertIn("SUPHELI", e1[0]["tier"])

    def test_match_tier_ref_id_filter_in_non_strict_clone(self):
        """Non-strict klon: ayni key1, taraflar kendi icinde ayni, ref_ids carpisirsa atla."""
        e1, e2, *_ = _cmp(
            [
                drow("60,00", "S", "1", "2", "0101", "K1", "No 11111111 A"),
                drow("60,00", "S", "1", "2", "0101", "K1", "No 11111111 A"),
            ],
            [
                drow("60,00", "S", "1", "2", "0101", "K1", "No 99999999 B"),
                drow("60,00", "S", "1", "2", "0101", "K1", "No 99999999 B"),
            ],
        )
        # Clone skips; may still get SUPHELI if strict allows 1:1 — with 2 vs 2
        # strict key1 fails; clone skips disjoint; SUPHELI strict also fails (len!=1)
        matched_pairs = [
            e for e in e1
            if e["matched"] and e.get("pair") is not None and e.get("pair_kind") == "MATCH"
        ]
        for e in matched_pairs:
            self.assertFalse(
                (e.get("ref_ids") or set()).isdisjoint(e["pair"].get("ref_ids") or set())
            )


if __name__ == "__main__":
    unittest.main()

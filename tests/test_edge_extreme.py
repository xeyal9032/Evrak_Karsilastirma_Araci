# -*- coding: utf-8 -*-
"""Derin kenar durumlar: storno, substring belirsizlik, strict, bos dosya."""
import os
import sys
import tempfile
import unittest
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(__file__))

import karsilastir_motor as motor
from helpers_datev import drow, write_datev_csv, write_journal_xlsx


class TestStornoEdgeCases(unittest.TestCase):
    def test_storno_without_twin_remains(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("100,00", "S", "2800", "2000", "0503", "RE-S", "<Storno> Orphan"),
            ])
            write_datev_csv(f2, [
                drow("50,00", "S", "1000", "2000", "0603", "OK", "Normal"),
            ])
            e1, e2, n1, n2 = motor.compare(f1, f2)
            self.assertEqual(n1, 0)
            self.assertFalse(e1[0]["netted"])
            self.assertFalse(e1[0]["matched"])

    def test_storno_wrong_amount_not_netted(self):
        rows = [
            drow("100,00", "S", "2000", "2800", "0503", "RE-S", "Original"),
            drow("99,00", "S", "2800", "2000", "0503", "RE-S", "<Storno> Original"),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, rows)
            write_datev_csv(f2, [drow("1,00", "S", "1", "2", "0101", "X", "pad")])
            e1, e2, n1, *_ = motor.compare(f1, f2)
            self.assertEqual(n1, 0)
            self.assertEqual(sum(1 for e in e1 if e["netted"]), 0)

    def test_two_storno_pairs_both_netted(self):
        rows = [
            drow("10,00", "S", "A", "B", "0101", "1", "Orig1"),
            drow("10,00", "S", "B", "A", "0101", "1", "<Storno> Orig1"),
            drow("20,00", "S", "C", "D", "0202", "2", "Orig2"),
            drow("20,00", "S", "D", "C", "0202", "2", "<Storno> Orig2"),
            drow("5,00", "S", "1", "2", "0303", "OK", "Keep"),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, rows)
            write_datev_csv(f2, [drow("5,00", "S", "1", "2", "0303", "OK", "Keep")])
            e1, e2, n1, n2 = motor.compare(f1, f2)
            self.assertEqual(n1, 2)
            self.assertEqual(sum(1 for e in e1 if e["netted"]), 4)
            keep = [e for e in e1 if e["belegfeld1"] == "OK"][0]
            self.assertEqual(keep["pair_kind"], "MATCH")


class TestSubstringAmbiguity(unittest.TestCase):
    def test_multiple_substring_hits_match_none(self):
        # Ayni tutar+tarih+SH, beleg iki farkli metinde geciyor -> belirsiz, esleme
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("300,00", "S", "5000", "1000", "1203", "FS12345", "Left"),
            ])
            write_datev_csv(f2, [
                drow("300,00", "S", "5100", "1000", "1203", "", "AAA FS12345 XXX"),
                drow("300,00", "S", "5200", "1000", "1203", "", "BBB FS12345 YYY"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            # substring_id_tier birden fazla hit olunca eslemez; key4 ile
            # tutar+konto+tarih de farkli konto oldugu icin MATCH olmaz.
            # Supheli key4 konto ayni degil -> eslesmemeli
            self.assertFalse(e1[0]["matched"])

    def test_short_beleg_below_min_len_no_substring(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("50,00", "S", "1", "2", "0101", "AB", "ShortBeleg"),
            ])
            write_datev_csv(f2, [
                drow("50,00", "S", "9", "2", "0101", "", "Contains AB inside"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            # min_len=5; "AB" substring ile eslenmemeli (konto farkli, key1-3 yok)
            self.assertFalse(e1[0]["matched"])


class TestStrictValueMismatch(unittest.TestCase):
    def test_ambiguous_value_tier_does_not_pair(self):
        # Ayni beleg_n yok; ayni genel metin+konto+tarih, farkli tutarlar, 2 aday
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("100,00", "S", "1000", "2000", "0101", "A1", "Generic Text"),
                drow("200,00", "S", "1000", "2000", "0101", "A2", "Generic Text"),
            ])
            write_datev_csv(f2, [
                drow("150,00", "S", "1000", "2000", "0101", "B1", "Generic Text"),
                drow("250,00", "S", "1000", "2000", "0101", "B2", "Generic Text"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            self.assertTrue(all(not e["matched"] for e in e1))
            self.assertTrue(all(not e["matched"] for e in e2))


class TestCloneRequiresIdentity(unittest.TestCase):
    def test_same_key_different_text_not_cloned(self):
        # key1 ayni (bos beleg?) - better: same amt/sh/konto/datum/beleg but different text
        # key1 uses beleg not text - so duplicates with different text still same key1
        # clone pass requires identical (belegfeld1, text) on ALL candidates
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [
                drow("100,00", "S", "1200", "8400", "1501", "DUP", "Text One"),
                drow("100,00", "S", "1200", "8400", "1501", "DUP", "Text Two"),
            ])
            write_datev_csv(f2, [
                drow("100,00", "S", "1200", "8400", "1501", "DUP", "Text One"),
                drow("100,00", "S", "1200", "8400", "1501", "DUP", "Text Two"),
            ])
            e1, e2, *_ = motor.compare(f1, f2)
            # strict fails (2 candidates); clone fails because not all identical within side
            # Actually each side has TWO different texts so _all_identical is False
            unmatched = sum(1 for e in e1 if not e["matched"])
            # They might match via key2 (text) one-to-one in strict mode!
            # key2: (amt, sh, konto, datum, text_n) - each text unique -> 1:1 MATCH
            self.assertEqual(unmatched, 0)
            self.assertTrue(all(e["pair_kind"] == "MATCH" for e in e1))


class TestJournalDateFormats(unittest.TestCase):
    def test_journal_datetime_and_stringish(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "j.xlsx")
            write_journal_xlsx(path, [
                (datetime(2026, 12, 31), "YEND", "Silvester", 10.0, 100, 200),
                (datetime(2026, 1, 1), "NY", "Neujahr", 20.5, 100, 200),
            ])
            e = motor.load(path)
            by_b = {x["belegfeld1"]: x for x in e}
            self.assertEqual(by_b["YEND"]["datum"], "3112")
            self.assertEqual(by_b["NY"]["datum"], "0101")
            self.assertEqual(by_b["NY"]["amt"], 20.5)


class TestTextTruncationKey(unittest.TestCase):
    def test_text_n_cuts_at_60_but_raw_preserved(self):
        long = "X" * 100
        e = motor.make_entry(1, "1,00", "S", "1", "2", "0101", "B", long, 1.0)
        self.assertEqual(len(e["text_n"]), 60)
        self.assertEqual(e["text"], long)

    def test_two_long_texts_same_prefix_collide_on_key2(self):
        t1 = "A" * 60 + "UNIQUE1"
        t2 = "A" * 60 + "UNIQUE2"
        e1 = motor.make_entry(1, "5,00", "S", "1", "2", "0101", "B1", t1, 5.0)
        e2 = motor.make_entry(2, "5,00", "S", "1", "2", "0101", "B2", t2, 5.0)
        # key2 ignores beleg; text_n identical -> same key
        self.assertEqual(motor.key2(e1), motor.key2(e2))
        # key1 differs by beleg
        self.assertNotEqual(motor.key1(e1), motor.key1(e2))


class TestSameFileCompare(unittest.TestCase):
    def test_file_against_itself_all_match(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "a.csv")
            write_datev_csv(path, [
                drow("1,00", "S", "1", "2", "0101", "A", "t1"),
                drow("2,00", "H", "3", "4", "0202", "B", "t2"),
                drow("3,50", "S", "5", "6", "0303", "C", "t3"),
            ])
            r = motor.build_report(path, path, os.path.join(tmp, "r.xlsx"), f1_label="A", f2_label="A")
            self.assertEqual(r["match_count"], 3)
            self.assertEqual(r["mismatch_count"], 0)
            self.assertEqual(r["only1_count"], 0)
            self.assertEqual(r["only2_count"], 0)
            self.assertEqual(r["diffs"], [])


class TestFormatErrorMessages(unittest.TestCase):
    def test_bilingual_missing_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bad.csv")
            with open(path, "w", encoding="latin-1", newline="") as f:
                f.write("EXTF;1\n")
                # detect_format gecer (Umsatz+SH+Konto var), find_columns Belegdatum'da duser
                f.write("Umsatz;Soll/Haben-Kennzeichen;Konto;Gegenkonto;Belegfeld 1;Buchungstext\n")
                f.write("1,00;S;1000;2000;B1;Text\n")
            with self.assertRaises(motor.FormatError) as cm:
                motor.load(path)
            msg = str(cm.exception).lower()
            self.assertTrue("belegdatum" in msg or "sutun" in msg or "столбец" in msg)
            self.assertTrue("столбец" in msg or "дата" in msg or "belegdatum" in msg)

    def test_unsupported_extension_bilingual(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "x.pdf")
            with open(path, "wb") as f:
                f.write(b"%PDF")
            with self.assertRaises(motor.FormatError) as cm:
                motor.detect_format(path)
            msg = str(cm.exception)
            self.assertIn("xlsx", msg.lower())
            self.assertIn("pdf", msg.lower())


class TestAccountBalanceGegenkonto(unittest.TestCase):
    def test_gegenkonto_receives_opposite_sign(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1 = os.path.join(tmp, "a.csv")
            f2 = os.path.join(tmp, "b.csv")
            write_datev_csv(f1, [drow("100,00", "S", "1000", "2000", "0101", "A", "T")])
            write_datev_csv(f2, [drow("100,00", "S", "1000", "2000", "0101", "A", "T")])
            e1, e2, *_ = motor.compare(f1, f2)
            # identical -> no diffs
            self.assertEqual(motor.account_balance_diff(e1, e2), [])
            # single file balance: 1000 += 100, 2000 -= 100
            diffs = motor.account_balance_diff(e1, [])
            by_k = {k: (b1, b2, d) for k, b1, b2, d, *_ in diffs}
            self.assertEqual(by_k["1000"][0], 100.0)
            self.assertEqual(by_k["2000"][0], -100.0)


class TestMatchTierDirect(unittest.TestCase):
    def _entry(self, **kwargs):
        base = dict(
            line=1, umsatz="1,00", sh="S", konto="1", gegenkonto="2", datum="0101",
            belegfeld1="B", text="T", amt=1.0, beleg_n="B", text_n="T",
            matched=False, tier=None, netted=False, pair=None, pair_kind=None,
        )
        base.update(kwargs)
        return base

    def test_strict_blocks_multi_candidate(self):
        a1 = self._entry(line=1, belegfeld1="X", beleg_n="X", amt=10.0)
        a2 = self._entry(line=2, belegfeld1="X", beleg_n="X", amt=10.0)
        b1 = self._entry(line=3, belegfeld1="X", beleg_n="X", amt=10.0)
        motor.match_tier([a1, a2], [b1], motor.key1, "T", kind="MATCH", strict=True)
        self.assertFalse(a1["matched"])
        self.assertFalse(a2["matched"])
        self.assertFalse(b1["matched"])

    def test_non_strict_pairs_min_count(self):
        a1 = self._entry(line=1, belegfeld1="X", beleg_n="X", amt=10.0)
        a2 = self._entry(line=2, belegfeld1="X", beleg_n="X", amt=10.0)
        b1 = self._entry(line=3, belegfeld1="X", beleg_n="X", amt=10.0)
        motor.match_tier([a1, a2], [b1], motor.key1, "T", kind="MATCH", strict=False)
        self.assertEqual(sum(1 for e in (a1, a2) if e["matched"]), 1)
        self.assertTrue(b1["matched"])


if __name__ == "__main__":
    unittest.main()

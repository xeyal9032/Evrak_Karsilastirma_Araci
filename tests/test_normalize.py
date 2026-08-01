# -*- coding: utf-8 -*-
"""Normalize / parse: harf ve rakam dogrulugu (kayip/bozulma olmamali)."""
import os
import sys
import tempfile
import unittest
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import karsilastir_motor as motor


class TestNormAmount(unittest.TestCase):
    def test_german_thousands(self):
        self.assertEqual(motor.norm_amount("1.234.567,89"), 1234567.89)

    def test_simple_comma(self):
        self.assertEqual(motor.norm_amount("1,99"), 1.99)
        self.assertEqual(motor.norm_amount("0,01"), 0.01)
        self.assertEqual(motor.norm_amount("1000,00"), 1000.0)

    def test_float_passthrough(self):
        self.assertEqual(motor.norm_amount(110.7), 110.7)
        self.assertEqual(motor.norm_amount(1), 1.0)

    def test_empty_and_invalid(self):
        self.assertIsNone(motor.norm_amount(""))
        self.assertIsNone(motor.norm_amount(None))
        self.assertIsNone(motor.norm_amount("abc"))

    def test_no_silent_rounding_error_on_cents(self):
        # Tipik EUR kuruslari 2 ondalikta kalmali
        for s, expected in [
            ("0,10", 0.10),
            ("0,20", 0.20),
            ("99,99", 99.99),
            ("123,45", 123.45),
        ]:
            self.assertEqual(motor.norm_amount(s), expected)

    def test_english_decimal_dot(self):
        self.assertEqual(motor.norm_amount("12.34"), 12.34)
        self.assertEqual(motor.norm_amount("1234.56"), 1234.56)
        self.assertEqual(motor.norm_amount("0.01"), 0.01)

    def test_english_thousands_comma(self):
        self.assertEqual(motor.norm_amount("1,234.56"), 1234.56)
        self.assertEqual(motor.norm_amount("1,234,567.89"), 1234567.89)

    def test_german_thousands_dot_only(self):
        self.assertEqual(motor.norm_amount("1.234"), 1234.0)
        self.assertEqual(motor.norm_amount("1.234.567"), 1234567.0)

    def test_signed_amounts(self):
        self.assertEqual(motor.norm_amount("-12,50"), -12.5)
        self.assertEqual(motor.norm_amount("-12.50"), -12.5)


class TestNormBelegfeld(unittest.TestCase):
    def test_preserves_letters_digits_after_strip(self):
        self.assertEqual(motor.norm_belegfeld("RE2026-0494"), "RE2026-0494")
        self.assertEqual(motor.norm_belegfeld("INV-HU-8556562"), "INV-HU-8556562")

    def test_removes_space_comma_dot_only(self):
        self.assertEqual(motor.norm_belegfeld("RE 2026.01"), "RE202601")
        self.assertEqual(motor.norm_belegfeld("12,34"), "1234")

    def test_casefold_upper(self):
        self.assertEqual(motor.norm_belegfeld("re2026"), "RE2026")

    def test_does_not_alter_slash_or_hyphen_content_wrongly(self):
        # Slash korunur (sadece space/comma/dot silinir)
        self.assertEqual(motor.norm_belegfeld("FS/26/01/2026"), "FS/26/01/2026")


class TestNormText(unittest.TestCase):
    def test_storno_prefix_stripped_for_key(self):
        self.assertEqual(
            motor.norm_text("<Storno> Original Buchung"),
            motor.norm_text("Original Buchung"),
        )

    def test_truncates_at_60(self):
        long = "A" * 80
        self.assertEqual(len(motor.norm_text(long)), 60)

    def test_preserves_umlaut_letters_upper(self):
        # latin-1/unicode harfler bozulmamali
        self.assertIn("MIE", motor.norm_text("Miete März"))


class TestBelegdatum(unittest.TestCase):
    def test_datetime(self):
        self.assertEqual(motor.format_belegdatum(datetime(2026, 3, 1)), "0103")
        self.assertEqual(motor.format_belegdatum(datetime(2026, 12, 31)), "3112")

    def test_date(self):
        self.assertEqual(motor.format_belegdatum(date(2026, 4, 15)), "1504")

    def test_string_formats(self):
        self.assertEqual(motor.format_belegdatum("0103"), "0103")
        self.assertEqual(motor.format_belegdatum("01.03.2026"), "0103")
        self.assertEqual(motor.format_belegdatum("1/3/2026"), "0103")

    def test_sort_key_chronological(self):
        self.assertLess(motor.date_sort_key("0103"), motor.date_sort_key("0203"))
        self.assertLess(motor.date_sort_key("3103"), motor.date_sort_key("0104"))


class TestCellStrAndUmsatz(unittest.TestCase):
    def test_float_account_no_decimal(self):
        self.assertEqual(motor._cell_str(1200.0), "1200")
        self.assertEqual(motor._cell_str(70064.0), "70064")

    def test_format_umsatz_german(self):
        self.assertEqual(motor.format_umsatz(1.99), "1,99")
        self.assertEqual(motor.format_umsatz(1000.0), "1000,00")

    def test_make_entry_preserves_exact_strings(self):
        e = motor.make_entry(
            line=10,
            umsatz_raw="1,99",
            sh="S",
            konto="6640",
            gegenkonto="70064",
            datum="0103",
            belegfeld1="8796",
            text="Aral Tankstelle Leonart Bewirtungskosten (nur MA)",
            amt=1.99,
        )
        self.assertEqual(e["umsatz"], "1,99")
        self.assertEqual(e["sh"], "S")
        self.assertEqual(e["konto"], "6640")
        self.assertEqual(e["gegenkonto"], "70064")
        self.assertEqual(e["datum"], "0103")
        self.assertEqual(e["belegfeld1"], "8796")
        self.assertEqual(e["text"], "Aral Tankstelle Leonart Bewirtungskosten (nur MA)")
        self.assertEqual(e["amt"], 1.99)


class TestCanonSh(unittest.TestCase):
    def test_s_unchanged(self):
        e = {"sh": "S", "konto": "1200", "gegenkonto": "8400"}
        self.assertEqual(motor.canon_sh_konto(e), ("S", "1200"))

    def test_h_flips_to_soll_viewpoint(self):
        e = {"sh": "H", "konto": "8400", "gegenkonto": "1200"}
        self.assertEqual(motor.canon_sh_konto(e), ("S", "1200"))


class TestSanitizeSheetName(unittest.TestCase):
    def test_invalid_chars_removed(self):
        taken = set()
        name = motor.sanitize_sheet_name("A:B*C?D", "Dosya1", taken)
        self.assertNotRegex(name, r'[:*?/\\\[\]]')

    def test_unique(self):
        taken = set()
        a = motor.sanitize_sheet_name("Same", "F1", taken)
        b = motor.sanitize_sheet_name("Same", "F1", taken)
        self.assertNotEqual(a.lower(), b.lower())


if __name__ == "__main__":
    unittest.main()

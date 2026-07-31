# -*- coding: utf-8 -*-
"""Property / fuzz benzeri derin normalize ve anahtar testleri."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import karsilastir_motor as motor


class TestPropertyNormAmount(unittest.TestCase):
    def test_many_cent_values(self):
        for cents in range(0, 10000, 7):
            euros = cents / 100.0
            s = f"{euros:.2f}".replace(".", ",")
            self.assertEqual(motor.norm_amount(s), round(euros, 2))

    def test_thousands_patterns(self):
        samples = [
            ("1.000,00", 1000.0),
            ("12.345,67", 12345.67),
            ("1.234.567,89", 1234567.89),
            ("0,01", 0.01),
            ("999.999,99", 999999.99),
        ]
        for s, expected in samples:
            self.assertEqual(motor.norm_amount(s), expected)

    def test_format_umsatz_roundtrip(self):
        for amt in (0.01, 1.0, 1.99, 110.7, 1000.0, 95286.74):
            s = motor.format_umsatz(amt)
            self.assertEqual(motor.norm_amount(s), round(amt, 2))


class TestPropertyBelegKeys(unittest.TestCase):
    def test_norm_beleg_idempotent(self):
        samples = ["RE-1001", "re 2026.01", "CBT/10520", "FS/26/01/2026", "2026 130000011"]
        for s in samples:
            n1 = motor.norm_belegfeld(s)
            n2 = motor.norm_belegfeld(n1)
            self.assertEqual(n1, n2)

    def test_same_economic_keys_for_sh_flip(self):
        a = motor.make_entry(1, "100,00", "H", "8400", "1200", "2001", "RE-X", "T", 100.0)
        b = motor.make_entry(2, "100,00", "S", "1200", "8400", "2001", "RE-X", "T", 100.0)
        self.assertEqual(motor.key1(a), motor.key1(b))
        self.assertEqual(motor.key_value1(a), motor.key_value1(b))


class TestPropertyDateKeys(unittest.TestCase):
    def test_format_belegdatum_idempotent_on_ddmm(self):
        for d in ("0101", "1503", "3112", "2902"):
            self.assertEqual(motor.format_belegdatum(d), d)

    def test_sort_order_months(self):
        seq = ["1501", "0102", "2802", "0103", "3112"]
        sorted_seq = sorted(seq, key=motor.date_sort_key)
        self.assertEqual(sorted_seq, seq)


if __name__ == "__main__":
    unittest.main()

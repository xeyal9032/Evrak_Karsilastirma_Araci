# -*- coding: utf-8 -*-
"""
Derin E2E: gercek excel1/mart dosyalari uzerinde alan butunlugu,
eslesen ciftlerde sessiz degisiklik yok, rapor tutarliligi.
"""
import os
import sys
import tempfile
import unittest
from collections import Counter

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(__file__))

import karsilastir_motor as motor
from fixture_paths import mart_dir, excel1_dir, has_mart

MART = mart_dir() or ""
EXCEL1 = excel1_dir() or ""
FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def _mart_paths():
    files = os.listdir(MART)
    csv_path = os.path.join(MART, [f for f in files if f.lower().endswith(".csv")][0])
    xlsx_path = os.path.join(MART, [f for f in files if f.lower().endswith(".xlsx")][0])
    return csv_path, xlsx_path


@unittest.skipUnless(has_mart(), "mart yok")
class TestE2EMartDeepIntegrity(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.csv_path, cls.xlsx_path = _mart_paths()
        cls.e1, cls.e2, cls.net1, cls.net2 = motor.compare(cls.csv_path, cls.xlsx_path)

    def test_no_duplicate_pair_objects(self):
        """Her eslesen kayit tek partnere bagli; capraz cift bozulmasi yok."""
        pairs = []
        for e in self.e1:
            if e["pair"] is not None and not e["netted"]:
                pairs.append((id(e), id(e["pair"])))
                self.assertIs(e["pair"]["pair"], e)
        # Partner id'leri tekil
        partner_ids = [p for _, p in pairs]
        self.assertEqual(len(partner_ids), len(set(partner_ids)))

    def test_all_match_pairs_amounts_equal(self):
        bad = []
        for e in self.e1:
            if e["pair_kind"] == "MATCH" and not e["netted"]:
                if e["amt"] != e["pair"]["amt"]:
                    bad.append((e["belegfeld1"], e["amt"], e["pair"]["amt"], e["tier"]))
        self.assertEqual(bad, [], msg=f"MATCH ama tutar farkli: {bad[:5]}")

    def test_all_mismatch_pairs_differ_in_some_field(self):
        for e in self.e1:
            if e["pair_kind"] != "MISMATCH" or e["netted"]:
                continue
            p = e["pair"]
            changed = (
                e["amt"] != p["amt"]
                or e["konto"] != p["konto"]
                or e["gegenkonto"] != p["gegenkonto"]
                or e["belegfeld1"] != p["belegfeld1"]
                or e["text"] != p["text"]
            )
            self.assertTrue(changed, msg=f"MISMATCH ama hic alan farkli degil: {e['belegfeld1']}")

    def test_belegfeld_characters_never_corrupted_on_load(self):
        raw_entries = motor.load(self.csv_path)
        for e in raw_entries:
            # Belegfeld sadece yazdirilabilir / alfanumerik / yaygin semboller
            b = e["belegfeld1"]
            self.assertEqual(b, b.strip().strip('"'))
            self.assertNotIn("\x00", b)
            self.assertNotIn("\ufffd", b)  # replacement char = bozulma
            self.assertNotIn("\ufffd", e["text"])

    def test_umsatz_string_parses_back_to_amt(self):
        for e in motor.load(self.csv_path):
            parsed = motor.norm_amount(e["umsatz"])
            self.assertEqual(parsed, e["amt"], msg=f"umsatz={e['umsatz']!r} amt={e['amt']}")

    def test_journal_umsatz_string_parses_back(self):
        for e in motor.load(self.xlsx_path):
            parsed = motor.norm_amount(e["umsatz"])
            self.assertEqual(parsed, e["amt"])

    def test_tier_histogram_sane(self):
        tiers = Counter()
        for e in self.e1:
            if e["matched"] and not e["netted"]:
                # Kisa etiket
                t = (e["tier"] or "?")[:40]
                tiers[t] += 1
        # En az bir tam/guclu eslesme olmali
        self.assertTrue(any("ESLESME" in k.upper() or "СОВПАД" in k.upper() or "Konto" in k for k in tiers))
        total_matched = sum(1 for e in self.e1 if e["matched"] and not e["netted"])
        self.assertGreaterEqual(total_matched, 600)

    def test_report_group_counts_match_result_dict(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "deep.xlsx")
            result = motor.build_report(
                self.csv_path, self.xlsx_path, out,
                f1_label="DATEV", f2_label="Journal",
            )
            groups = motor.build_combined_rows(self.e1, self.e2)
            self.assertEqual(result["match_count"], sum(1 for g in groups if g["kind"] == "MATCH"))
            self.assertEqual(result["mismatch_count"], sum(1 for g in groups if g["kind"] == "MISMATCH"))
            self.assertEqual(result["only1_count"], sum(1 for g in groups if g["kind"] == "ONLY1"))
            self.assertEqual(result["only2_count"], sum(1 for g in groups if g["kind"] == "ONLY2"))

            wb = openpyxl.load_workbook(out)
            ws = wb["KARSILASTIRMA-SRAVNENIE"]
            # Her grup 2 satir -> data satiri = 2 * group_count
            data_rows = ws.max_row - 1
            self.assertEqual(data_rows, 2 * (
                result["match_count"] + result["mismatch_count"]
                + result["only1_count"] + result["only2_count"]
            ))

    def test_kraftstoff_and_bewirtung_both_present_and_match(self):
        e1 = self.e1
        bew = [e for e in e1 if e["text"] == "Aral Tankstelle Leonart Bewirtungskosten (nur MA)"]
        kraft = [e for e in e1 if e["text"] == "Aral Tankstelle Leonart Kraftstoff/Ladestrom"]
        self.assertEqual(len(bew), 1)
        self.assertEqual(len(kraft), 1)
        self.assertEqual(bew[0]["amt"], 1.99)
        self.assertEqual(kraft[0]["amt"], 110.7)
        self.assertTrue(bew[0]["matched"] and bew[0]["pair_kind"] == "MATCH")
        self.assertTrue(kraft[0]["matched"] and kraft[0]["pair_kind"] == "MATCH")
        self.assertEqual(bew[0]["pair"]["text"], bew[0]["text"])
        self.assertEqual(kraft[0]["pair"]["text"], kraft[0]["text"])


@unittest.skipUnless(
    os.path.isfile(os.path.join(EXCEL1, "EXTF_Buchungsstapel_01012026_bis_30062026.csv"))
    and os.path.isfile(os.path.join(EXCEL1, "EXTF_MB_01-06.csv")),
    "excel1 DATEV cifti yok",
)
class TestE2ETwoDatevDeep(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.f1 = os.path.join(EXCEL1, "EXTF_Buchungsstapel_01012026_bis_30062026.csv")
        cls.f2 = os.path.join(EXCEL1, "EXTF_MB_01-06.csv")
        cls.e1, cls.e2, cls.net1, cls.net2 = motor.compare(cls.f1, cls.f2)

    def test_load_roundtrip_amounts(self):
        for path in (self.f1, self.f2):
            for e in motor.load(path):
                self.assertEqual(motor.norm_amount(e["umsatz"]), e["amt"])

    def test_match_pairs_preserve_beleg_when_tam_eslesme(self):
        checked = 0
        for e in self.e1:
            if not e["matched"] or e["netted"]:
                continue
            if e["tier"] and e["tier"].startswith("TAM ESLESME"):
                # Eslesme norm_belegfeld ile yapilir (bosluk/nokta silinir);
                # ham metin farkli olabilir ama anlam ayni kalmali.
                self.assertEqual(e["beleg_n"], e["pair"]["beleg_n"])
                self.assertEqual(e["amt"], e["pair"]["amt"])
                self.assertEqual(e["datum"], e["pair"]["datum"])
                checked += 1
        self.assertGreater(checked, 50)

    def test_build_report_smoke_large(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "large.xlsx")
            result = motor.build_report(self.f1, self.f2, out, f1_label="BS", f2_label="MB")
            self.assertTrue(os.path.isfile(out))
            self.assertGreater(result["match_count"], 100)
            size = os.path.getsize(out)
            self.assertGreater(size, 50_000)


class TestE2EFixtureMatrix(unittest.TestCase):
    """Tum fixture kombinasyonlari hatasiz ve tutarli."""

    def test_all_fixture_pairs(self):
        sample_a = os.path.join(FIXTURES, "sample_a.csv")
        sample_b = os.path.join(FIXTURES, "sample_b.csv")
        journal = os.path.join(FIXTURES, "sample_journal.xlsx")
        datev_x = os.path.join(FIXTURES, "sample_datev.xlsx")
        pairs = [
            (sample_a, sample_b),
            (sample_a, journal),
            (sample_a, datev_x),
            (sample_b, journal),
            (datev_x, journal),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            for i, (a, b) in enumerate(pairs):
                out = os.path.join(tmp, f"p{i}.xlsx")
                result = motor.build_report(a, b, out, f1_label="F1", f2_label="F2")
                self.assertTrue(os.path.isfile(out))
                self.assertGreaterEqual(result["f1_total"], 1)
                self.assertGreaterEqual(result["f2_total"], 1)
                # MATCH ciftlerinde tutar esit
                e1, e2, *_ = motor.compare(a, b)
                for e in e1:
                    if e["pair_kind"] == "MATCH" and not e["netted"]:
                        self.assertEqual(e["amt"], e["pair"]["amt"])


class TestE2EGuiImportSmoke(unittest.TestCase):
    def test_gui_module_imports(self):
        import karsilastir as gui
        self.assertTrue(hasattr(gui, "App"))
        self.assertTrue(hasattr(gui, "open_file"))
        self.assertTrue(hasattr(gui, "desktop_dir"))
        self.assertTrue(callable(gui.desktop_dir))
        d = gui.desktop_dir()
        self.assertTrue(os.path.isdir(d))


if __name__ == "__main__":
    unittest.main()

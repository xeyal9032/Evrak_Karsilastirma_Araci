# -*- coding: utf-8 -*-
"""
Rapor i18n / PDF Unicode SOZLESME testleri.

Bu dosya gecmise ait gercek regresyonlari kilitler:
1) HTML arayuzu secilen dile gore uretilmeli (hardcoded EN yasak)
2) PDF Turkce/Kiril/Almanca glifleri ■■■ olmamali (Helvetica yasak, Unicode font)
3) build_report(lang=...) ve CLI --lang HTML+PDF'e aktarilmali
4) Locale'deki html_* / pdf_* anahtarlari ciktiya yansiyabilmeli
"""
import os
import re
import sys
import tempfile
import unittest
from io import StringIO
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import i18n
import karsilastir_motor as motor
import report_extra
from karsilastir import run_cli
from tests.helpers_datev import drow, write_datev_csv

# Hardcoded EN chrome - non-en HTML'de asla gorunmemeli (dil-ortak kelimeler haric)
FORBIDDEN_EN_IN_NON_EN_HTML = (
    "Match (yellow)",
    "Mismatch (orange)",
    "Only file1 (red)",
    "Only file2 (red)",
    "Account differences",
    "No significant diffs",
    "Generated ",
    'lang="en"',
    ">All</button>",
    'placeholder="Filter text',
    ">Source</th>",
    ">Amount</th>",
    "Critical rows (red / orange)",
    "Document Comparison",
)

# Her dil icin PDF/HTML'de mutlaka bulunmasi gereken ayirici metinler
LANG_MARKERS = {
    "tr": ("Eşleşme (sarı)", "Hesap farkları", "Oluşturulma:", "Kritik satırlar"),
    "en": ("Match (yellow)", "Account differences", "Generated ", "Critical rows"),
    "de": ("Treffer (gelb)", "Kontodifferenzen", "Erstellt:", "Kritische Zeilen"),
    "ru": ("Совпадение (жёлтый)", "Разницы по счетам", "Создано:", "Критические строки"),
}

# Unicode stress: bu karakterler Helvetica ile PDF'de kaybolur / ■ olur
UNICODE_STRESS = {
    "tr": "şğıüöçıŞĞİÜÖÇ",
    "de": "äöüßÄÖÜ",
    "ru": "ЖёлтыйСчётРазница",
    "en": "Comparison",
}


def _have_reportlab():
    try:
        import reportlab  # noqa: F401
        return True
    except ImportError:
        return False


def _have_pypdf():
    try:
        from pypdf import PdfReader  # noqa: F401
        return True
    except ImportError:
        return False


def _pdf_text(path):
    from pypdf import PdfReader
    return "".join((p.extract_text() or "") for p in PdfReader(path).pages)


def _make_pair(tmp):
    f1 = os.path.join(tmp, "a.csv")
    f2 = os.path.join(tmp, "b.csv")
    write_datev_csv(f1, [
        drow("10,00", "S", "1", "2", "0101", "M", "ok match"),
        drow("20,00", "S", "1", "2", "0202", "O1", "only file1"),
    ], encoding="utf-8")
    write_datev_csv(f2, [
        drow("10,00", "S", "1", "2", "0101", "M", "ok match"),
        drow("40,00", "S", "1", "2", "0404", "O2", "only file2"),
    ], encoding="utf-8")
    return f1, f2


def _read(path):
    with open(path, encoding="utf-8") as fh:
        return fh.read()


class TestHtmlI18nContract(unittest.TestCase):
    def test_every_supported_lang_html_chrome(self):
        for lang in i18n.SUPPORTED:
            with self.subTest(lang=lang):
                with tempfile.TemporaryDirectory() as tmp:
                    path = os.path.join(tmp, f"{lang}.html")
                    report_extra.write_html_report(
                        path, f1_path="a.csv", f2_path="b.csv",
                        f1_label="A", f2_label="B",
                        match_count=1, mismatch_count=2, only1_count=3, only2_count=4,
                        groups=[], lang=lang,
                    )
                    text = _read(path)
                    self.assertIn(f'lang="{lang}"', text)
                    for marker in LANG_MARKERS[lang][:3]:
                        self.assertIn(marker, text, msg=f"{lang} missing {marker!r}")
                    # locale keys must appear as rendered values
                    for key in (
                        "html_card_match", "html_account_diffs", "html_filter_all",
                        "html_col_source", "html_col_status", "html_rows",
                    ):
                        val = i18n.t(key, lang=lang)
                        self.assertIn(val, text, msg=f"{lang}:{key}")

    def test_non_en_html_forbids_english_chrome(self):
        for lang in ("tr", "de", "ru"):
            with self.subTest(lang=lang):
                with tempfile.TemporaryDirectory() as tmp:
                    path = os.path.join(tmp, f"{lang}.html")
                    report_extra.write_html_report(
                        path, f1_path="a", f2_path="b", f1_label="A", f2_label="B",
                        match_count=0, mismatch_count=0, only1_count=0, only2_count=0,
                        groups=[], lang=lang,
                    )
                    text = _read(path)
                    for bad in FORBIDDEN_EN_IN_NON_EN_HTML:
                        self.assertNotIn(bad, text, msg=f"{lang} still has English {bad!r}")

    def test_html_missing_note_localized(self):
        with tempfile.TemporaryDirectory() as tmp:
            f1, f2 = _make_pair(tmp)
            e1, e2, *_ = motor.compare(f1, f2)
            groups = motor.build_combined_rows(e1, e2)
            rows_tr = report_extra.groups_to_rows(groups, "A", "B", lang="tr")
            rows_en = report_extra.groups_to_rows(groups, "A", "B", lang="en")
            notes_tr = {r["note"] for r in rows_tr if r["note"]}
            notes_en = {r["note"] for r in rows_en if r["note"]}
            self.assertIn("EKSİK", notes_tr)
            self.assertIn("MISSING", notes_en)
            self.assertNotIn("MISSING", notes_tr)

    def test_report_extra_source_has_no_hardcoded_en_html_lang(self):
        """Eski bug: <html lang=\"en\"> sabitti. Tekrar girmemeli."""
        src = _read(os.path.join(ROOT, "report_extra.py"))
        self.assertNotIn('html lang="en"', src)
        self.assertIn('html lang="{_esc(lang)}"', src)
        # UI metinleri _t / i18n uzerinden gelmeli
        self.assertIn('_t("html_card_match"', src)
        self.assertIn('_t("html_account_diffs"', src)
        self.assertNotIn("Match (yellow)", src)


class TestPdfUnicodeContract(unittest.TestCase):
    @unittest.skipUnless(_have_reportlab() and _have_pypdf(), "reportlab/pypdf yok")
    def test_pdf_all_langs_no_tofu_boxes(self):
        for lang in i18n.SUPPORTED:
            with self.subTest(lang=lang):
                with tempfile.TemporaryDirectory() as tmp:
                    path = os.path.join(tmp, f"{lang}.pdf")
                    report_extra.write_pdf_report(
                        path, f1_path="a.csv", f2_path="b.csv",
                        f1_label="A", f2_label="B",
                        match_count=1, mismatch_count=0, only1_count=0, only2_count=0,
                        groups=[], lang=lang,
                    )
                    text = _pdf_text(path)
                    self.assertNotIn("■■", text, msg=f"{lang} PDF has tofu boxes")
                    for marker in LANG_MARKERS[lang]:
                        self.assertIn(marker, text, msg=f"{lang} PDF missing {marker!r}")

    @unittest.skipUnless(_have_reportlab() and _have_pypdf(), "reportlab/pypdf yok")
    def test_pdf_unicode_stress_in_file_paths_and_title(self):
        """Dosya yolu / baslik icindeki ozel karakterler bozulmamali."""
        with tempfile.TemporaryDirectory() as tmp:
            for lang, stress in UNICODE_STRESS.items():
                path = os.path.join(tmp, f"stress_{lang}.pdf")
                fake = os.path.join(tmp, f"dosya_{stress}.csv")
                report_extra.write_pdf_report(
                    path, f1_path=fake, f2_path=fake,
                    f1_label=stress, f2_label=stress,
                    match_count=0, mismatch_count=0, only1_count=0, only2_count=0,
                    groups=[], lang=lang, title=f"Başlık {stress}",
                )
                text = _pdf_text(path)
                self.assertNotIn("■■", text, msg=lang)
                # En azindan dil isaretleri saglam kalsin
                self.assertIn(LANG_MARKERS[lang][0], text)

    @unittest.skipUnless(_have_reportlab(), "reportlab yok")
    def test_pdf_registers_unicode_font_when_available(self):
        font, bold = report_extra._register_pdf_fonts()
        # Windows CI / lokal: Arial/Segoe bulunmali. Linux: DejaVu.
        # Helvetica ancak hic TTF yoksa kabul; o zaman Unicode testi skip edilir.
        windir = os.environ.get("WINDIR", r"C:\Windows")
        has_ttf = any(os.path.isfile(p) for p in (
            os.path.join(windir, "Fonts", "segoeui.ttf"),
            os.path.join(windir, "Fonts", "arial.ttf"),
            os.path.join(windir, "Fonts", "arialuni.ttf"),
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ))
        if has_ttf:
            self.assertNotEqual(font, "Helvetica")
            self.assertTrue(font.startswith("EvrakSans"))
        else:
            self.skipTest("Unicode TTF yok - Helvetica fallback")


class TestBuildReportLangWiring(unittest.TestCase):
    def test_build_report_lang_flows_to_html_and_pdf(self):
        if not _have_reportlab():
            self.skipTest("reportlab yok")
        with tempfile.TemporaryDirectory() as tmp:
            f1, f2 = _make_pair(tmp)
            out = os.path.join(tmp, "r.xlsx")
            for lang in ("tr", "ru"):
                with self.subTest(lang=lang):
                    r = motor.build_report(
                        f1, f2, out.replace(".xlsx", f"_{lang}.xlsx"),
                        write_html=True, write_pdf=True, lang=lang,
                        detail_sheets=False,
                    )
                    html = _read(r["html_path"])
                    self.assertIn(f'lang="{lang}"', html)
                    self.assertIn(LANG_MARKERS[lang][0], html)
                    if lang != "en":
                        self.assertNotIn("Match (yellow)", html)
                    if _have_pypdf():
                        pdf = _pdf_text(r["pdf_path"])
                        self.assertNotIn("■■", pdf)
                        self.assertIn(LANG_MARKERS[lang][0], pdf)

    def test_cli_lang_tr_html_pdf(self):
        if not _have_reportlab():
            self.skipTest("reportlab yok")
        with tempfile.TemporaryDirectory() as tmp:
            f1, f2 = _make_pair(tmp)
            out = os.path.join(tmp, "cli.xlsx")
            buf = StringIO()
            with mock.patch("sys.stdout", buf):
                code = run_cli([f1, f2, "-o", out, "--html", "--pdf", "--lang", "tr", "--quiet"])
            self.assertEqual(code, 0)
            html = _read(out.replace(".xlsx", ".html"))
            self.assertIn('lang="tr"', html)
            self.assertIn("Eşleşme (sarı)", html)
            self.assertNotIn("Match (yellow)", html)
            if _have_pypdf():
                pdf = _pdf_text(out.replace(".xlsx", ".pdf"))
                self.assertIn("Eşleşme", pdf)
                self.assertNotIn("■■", pdf)


class TestLocaleKeysConsumedByReports(unittest.TestCase):
    """report_extra'nin kullandigi her i18n anahtari tum dillerde dolu olmali."""

    REPORT_KEYS = {
        "html_title", "html_generated", "html_file1", "html_file2",
        "html_card_match", "html_card_mismatch", "html_card_only1", "html_card_only2",
        "html_account_diffs", "html_no_diffs",
        "html_th_konto", "html_th_bal1", "html_th_bal2", "html_th_diff",
        "html_th_n1", "html_th_n2", "html_rows",
        "html_filter_all", "html_filter_match", "html_filter_mismatch",
        "html_filter_only1", "html_filter_only2", "html_filter_red",
        "html_filter_placeholder",
        "html_col_source", "html_col_line", "html_col_amount", "html_col_sh",
        "html_col_konto", "html_col_gegen", "html_col_date", "html_col_beleg",
        "html_col_text", "html_col_note", "html_col_status",
        "html_missing", "html_count", "pdf_critical_rows",
    }

    def test_report_keys_exist_and_differ_across_langs(self):
        for key in self.REPORT_KEYS:
            vals = {}
            for lang in i18n.SUPPORTED:
                val = i18n.t(key, lang=lang)
                self.assertNotEqual(val, key, msg=f"missing translation {lang}.{key}")
                self.assertTrue(str(val).strip(), msg=f"empty {lang}.{key}")
                vals[lang] = val
            # Baslik / kart gibi kullaniciya gorunen anahtarlar en az 2 dilde farkli olmali
            if key in ("html_card_match", "html_title", "html_account_diffs", "html_missing"):
                self.assertGreaterEqual(len(set(vals.values())), 3, msg=key)

    def test_report_extra_references_only_existing_keys(self):
        src = _read(os.path.join(ROOT, "report_extra.py"))
        used = set(re.findall(r'_t\(\s*"([a-z0-9_]+)"', src))
        used |= set(re.findall(r'i18n\.t\(\s*"([a-z0-9_]+)"', src))
        self.assertTrue(used)
        for key in used:
            for lang in i18n.SUPPORTED:
                self.assertNotEqual(i18n.t(key, lang=lang), key, msg=f"{lang}.{key}")


class TestGuiAndCliLangWiringSource(unittest.TestCase):
    """GUI/CLI kaynak kodu lang parametresini unutmamali."""

    def test_gui_passes_lang_to_build_report(self):
        src = _read(os.path.join(ROOT, "karsilastir.py"))
        self.assertIn("lang=self.lang.get()", src)
        self.assertIn("lang=lang", src)

    def test_batch_passes_lang(self):
        src = _read(os.path.join(ROOT, "batch_compare.py"))
        self.assertIn("lang=lang", src)
        self.assertIn("lang=None", src)


class TestCiFontAvailabilityHint(unittest.TestCase):
    """Linux CI'de DejaVu yoksa PDF Unicode testi anlamsiz kalir - en az bir font yolu dene."""

    def test_at_least_one_unicode_font_candidate_or_windows(self):
        windir = os.environ.get("WINDIR", r"C:\Windows")
        candidates = [
            os.path.join(windir, "Fonts", "segoeui.ttf"),
            os.path.join(windir, "Fonts", "arial.ttf"),
            os.path.join(windir, "Fonts", "arialuni.ttf"),
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        ]
        found = [p for p in candidates if os.path.isfile(p)]
        if sys.platform.startswith("linux") and not found:
            self.fail(
                "Linux'ta Unicode TTF yok. CI'ye 'fonts-dejavu-core' eklenmeli. "
                f"Aranan: {candidates}"
            )
        # Windows'ta genelde segoe/arial vardir
        if sys.platform == "win32":
            self.assertTrue(found, msg="Windows Fonts altinda Arial/Segoe beklenir")


if __name__ == "__main__":
    unittest.main()

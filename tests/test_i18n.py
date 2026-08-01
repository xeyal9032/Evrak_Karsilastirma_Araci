# -*- coding: utf-8 -*-
"""
Kapsamli i18n testleri: locale dosyalari, anahtar parity, placeholder'lar,
format, fallback, cache, sistem dili algilama, GUI/CLI kullanim anahtarlari.
"""
import json
import os
import re
import unittest
from unittest import mock

import i18n

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCALES_DIR = os.path.join(ROOT, "locales")

# karsilastir.py icinde gercekten kullanilan anahtarlar (regresyon)
GUI_KEYS = {
    "app_title", "app_subtitle", "drop_hint", "lang_label",
    "file1", "file2", "compare", "pick_title",
    "format_error_title", "missing_title", "missing_files",
    "comparing", "done", "warn_open_title", "warn_open_body",
    "summary_sheet", "summary_yellow", "summary_orange",
    "summary_red1", "summary_red2", "summary_diffs", "summary_diff_line",
    "summary_no_diffs", "done_title", "error_status",
    "err_permission", "err_not_found", "err_unexpected",
    "opt_html", "opt_pdf", "opt_fast", "opt_archive",
    "btn_batch", "btn_archive", "batch_title", "batch_pick_a", "batch_pick_b",
    "batch_done", "archive_title", "archive_empty", "archive_warn", "report_warn",
}
EXCEL_KEYS = {
    "excel_h_line", "excel_h_amount", "excel_h_sh", "excel_h_konto", "excel_h_gegen",
    "excel_h_date", "excel_h_beleg", "excel_h_text", "excel_h_status",
    "excel_h_source", "excel_h_note", "excel_missing", "excel_only_here",
    "excel_summary_title", "excel_file1", "excel_file2",
    "excel_col_total", "excel_col_matched", "excel_col_only", "excel_net_dup",
    "excel_combined_summary", "excel_same_yellow", "excel_diff_orange",
    "excel_only_f1", "excel_only_f2", "excel_legend_title",
    "excel_legend_yellow", "excel_legend_yellow_desc",
    "excel_legend_orange", "excel_legend_orange_desc",
    "excel_legend_green", "excel_legend_green_desc",
    "excel_legend_red", "excel_legend_red_desc",
    "excel_source_file", "excel_bal", "excel_diff", "excel_n_entries",
}
CLI_KEYS = {"cli_ok", "cli_counts"}
HTML_KEYS = {
    "html_title", "html_generated", "html_file1", "html_file2",
    "html_card_match", "html_card_mismatch", "html_card_only1", "html_card_only2",
    "html_account_diffs", "html_no_diffs",
    "html_th_konto", "html_th_bal1", "html_th_bal2", "html_th_diff", "html_th_n1", "html_th_n2",
    "html_rows",
    "html_filter_all", "html_filter_match", "html_filter_mismatch",
    "html_filter_only1", "html_filter_only2", "html_filter_red", "html_filter_placeholder",
    "html_col_source", "html_col_line", "html_col_amount", "html_col_sh", "html_col_konto",
    "html_col_gegen", "html_col_date", "html_col_beleg", "html_col_text", "html_col_note",
    "html_col_status", "html_missing", "html_count", "pdf_critical_rows",
}

# format() placeholder'lari (key -> beklenen alan adlari)
PLACEHOLDERS = {
    "pick_title": {"which"},
    "comparing": {"fmt", "pct"},
    "done": {"name"},
    "warn_open_body": {"path"},
    "err_not_found": {"err"},
    "err_unexpected": {"err"},
    "summary_yellow": {"n"},
    "summary_orange": {"n"},
    "summary_red1": {"n"},
    "summary_red2": {"n"},
    "summary_diff_line": {"k", "d"},
    "cli_ok": {"path"},
    "cli_counts": {"m", "mm", "o1", "o2"},
    "html_generated": {"ts"},
    "html_file1": {"path"},
    "html_file2": {"path"},
    "excel_only_f1": {"name"},
    "excel_only_f2": {"name"},
    "excel_bal": {"name"},
    "excel_n_entries": {"name"},
    "batch_done": {"n", "path"},
    "archive_warn": {"err"},
    "report_warn": {"err"},
}

_PLACEHOLDER_RE = re.compile(r"\{([a-zA-Z_][a-zA-Z0-9_]*)(?::[^}]*)?\}")


def _placeholders_in(text):
    return set(_PLACEHOLDER_RE.findall(text or ""))


class TestLocaleFilesOnDisk(unittest.TestCase):
    def test_supported_tuple(self):
        self.assertEqual(i18n.SUPPORTED, ("tr", "ru", "de", "en"))
        self.assertEqual(i18n._DEFAULT_LANG, "tr")

    def test_each_locale_file_exists_and_is_object(self):
        for lang in i18n.SUPPORTED:
            path = os.path.join(LOCALES_DIR, f"{lang}.json")
            self.assertTrue(os.path.isfile(path), msg=path)
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            self.assertIsInstance(data, dict)
            self.assertTrue(data)

    def test_locales_dir_helper_points_to_repo_locales(self):
        self.assertEqual(
            os.path.normcase(os.path.abspath(i18n._locales_dir())),
            os.path.normcase(os.path.abspath(LOCALES_DIR)),
        )

    def test_no_extra_unexpected_json_required(self):
        """SUPPORTED disindaki json dosyalari varsa da sorun degil; SUPPORTED hepsi mevcut."""
        for lang in i18n.SUPPORTED:
            self.assertTrue(os.path.isfile(os.path.join(LOCALES_DIR, f"{lang}.json")))


class TestKeyParity(unittest.TestCase):
    def test_all_langs_same_key_set(self):
        base = i18n.locale_keys("tr")
        self.assertGreaterEqual(len(base), 20)
        for lang in i18n.SUPPORTED:
            self.assertEqual(base, i18n.locale_keys(lang), msg=lang)

    def test_gui_keys_present_in_every_locale(self):
        for lang in i18n.SUPPORTED:
            keys = i18n.locale_keys(lang)
            missing = GUI_KEYS - keys
            self.assertFalse(missing, msg=f"{lang} missing GUI keys: {missing}")

    def test_cli_keys_present_in_every_locale(self):
        for lang in i18n.SUPPORTED:
            keys = i18n.locale_keys(lang)
            missing = CLI_KEYS - keys
            self.assertFalse(missing, msg=f"{lang} missing CLI keys: {missing}")

    def test_html_keys_present_all_langs(self):
        for lang in i18n.SUPPORTED:
            keys = i18n.locale_keys(lang)
            missing = HTML_KEYS - keys
            self.assertFalse(missing, msg=f"{lang} missing HTML keys: {missing}")

    def test_excel_keys_present_all_langs(self):
        for lang in i18n.SUPPORTED:
            keys = i18n.locale_keys(lang)
            missing = EXCEL_KEYS - keys
            self.assertFalse(missing, msg=f"{lang} missing Excel keys: {missing}")

    def test_locale_keys_default_is_turkish(self):
        self.assertEqual(i18n.locale_keys(), i18n.locale_keys("tr"))

    def test_all_values_are_nonempty_strings(self):
        for lang in i18n.SUPPORTED:
            data = i18n._load(lang)
            for key, val in data.items():
                self.assertIsInstance(val, str, msg=f"{lang}.{key}")
                self.assertTrue(val.strip(), msg=f"{lang}.{key} empty")


class TestPlaceholders(unittest.TestCase):
    def test_placeholder_sets_match_across_langs(self):
        for key, expected in PLACEHOLDERS.items():
            ref = None
            for lang in i18n.SUPPORTED:
                text = i18n._load(lang)[key]
                found = _placeholders_in(text)
                self.assertEqual(
                    found, expected,
                    msg=f"{lang}.{key}: got {found}, expected {expected}",
                )
                if ref is None:
                    ref = found
                else:
                    self.assertEqual(found, ref, msg=f"{lang}.{key} drift")

    def test_keys_without_placeholders_have_none(self):
        for lang in i18n.SUPPORTED:
            data = i18n._load(lang)
            for key, val in data.items():
                if key in PLACEHOLDERS:
                    continue
                self.assertFalse(
                    _placeholders_in(val),
                    msg=f"{lang}.{key} unexpectedly has placeholders: {val!r}",
                )


class TestTranslate(unittest.TestCase):
    def test_known_translations(self):
        self.assertEqual(i18n.t("compare", lang="tr"), "Karşılaştır")
        self.assertEqual(i18n.t("compare", lang="en"), "Compare")
        self.assertEqual(i18n.t("compare", lang="de"), "Vergleichen")
        self.assertEqual(i18n.t("compare", lang="ru"), "Сравнить")

    def test_default_lang_when_lang_none(self):
        self.assertEqual(i18n.t("compare"), i18n.t("compare", lang="tr"))

    def test_case_insensitive_lang_code(self):
        self.assertEqual(i18n.t("compare", lang="EN"), "Compare")
        self.assertEqual(i18n.t("compare", lang="De"), "Vergleichen")
        self.assertEqual(i18n.t("compare", lang="TR"), "Karşılaştır")

    def test_unsupported_lang_falls_back_to_tr(self):
        self.assertEqual(i18n.t("compare", lang="fr"), i18n.t("compare", lang="tr"))
        self.assertEqual(i18n.t("compare", lang="xx"), i18n.t("compare", lang="tr"))
        self.assertEqual(i18n.t("compare", lang=""), i18n.t("compare", lang="tr"))

    def test_missing_key_returns_key_name(self):
        self.assertEqual(i18n.t("___missing___", lang="en"), "___missing___")
        self.assertEqual(i18n.t("___missing___", lang="tr"), "___missing___")

    def test_languages_not_all_identical_for_title(self):
        titles = {i18n.t("app_title", lang=lang) for lang in i18n.SUPPORTED}
        self.assertGreaterEqual(len(titles), 3)


class TestFormatKwargs(unittest.TestCase):
    def test_done_inserts_name(self):
        for lang in i18n.SUPPORTED:
            text = i18n.t("done", lang=lang, name="Rapor_1.xlsx")
            self.assertIn("Rapor_1.xlsx", text)

    def test_pick_title_which(self):
        for lang in i18n.SUPPORTED:
            text = i18n.t("pick_title", lang=lang, which=2)
            self.assertIn("2", text)

    def test_comparing_fmt_and_pct(self):
        for lang in i18n.SUPPORTED:
            text = i18n.t("comparing", lang=lang, fmt=" (DATEV)", pct=67)
            self.assertIn("67", text)
            self.assertIn("DATEV", text)

    def test_summary_counts(self):
        for lang in i18n.SUPPORTED:
            self.assertIn("12", i18n.t("summary_yellow", lang=lang, n=12))
            self.assertIn("3", i18n.t("summary_orange", lang=lang, n=3))
            self.assertIn("7", i18n.t("summary_red1", lang=lang, n=7))
            self.assertIn("9", i18n.t("summary_red2", lang=lang, n=9))

    def test_summary_diff_line_number_format(self):
        text = i18n.t("summary_diff_line", lang="en", k="4400", d=1234.5)
        self.assertIn("4400", text)
        self.assertIn("1,234.50", text)
        self.assertIn("EUR", text)

    def test_cli_ok_and_counts(self):
        ok = i18n.t("cli_ok", lang="en", path=r"C:\out.xlsx")
        self.assertIn("out.xlsx", ok)
        counts = i18n.t("cli_counts", lang="de", m=1, mm=2, o1=3, o2=4)
        self.assertIn("match=1", counts)
        self.assertIn("mismatch=2", counts)
        self.assertIn("only1=3", counts)
        self.assertIn("only2=4", counts)

    def test_warn_open_body_path_and_newline(self):
        text = i18n.t("warn_open_body", lang="en", path="/tmp/a.xlsx")
        self.assertIn("/tmp/a.xlsx", text)
        self.assertIn("\n", text)

    def test_missing_format_kwarg_returns_unformatted_template(self):
        # {name} eksik -> KeyError yakalanir, ham sablon doner
        raw = i18n._load("en")["done"]
        self.assertEqual(i18n.t("done", lang="en"), raw)

    def test_wrong_format_kwarg_returns_template(self):
        raw = i18n._load("en")["done"]
        self.assertEqual(i18n.t("done", lang="en", wrong="x"), raw)

    def test_extra_unused_kwargs_still_format(self):
        text = i18n.t("done", lang="en", name="ok.xlsx", unused=123)
        self.assertIn("ok.xlsx", text)


class TestUnicode(unittest.TestCase):
    def test_turkish_special_chars(self):
        title = i18n.t("app_title", lang="tr")
        self.assertTrue(any(c in title for c in "şığüöçŞİĞÜÖÇ"))

    def test_russian_cyrillic(self):
        text = i18n.t("compare", lang="ru")
        self.assertTrue(any("\u0400" <= c <= "\u04FF" for c in text))

    def test_german_umlaut_or_plain(self):
        # "Datei" / "wählen" vb. - en azindan Almanca kelime farkli olmali
        self.assertNotEqual(
            i18n.t("file1", lang="de"),
            i18n.t("file1", lang="en"),
        )

    def test_json_roundtrip_preserves_unicode(self):
        for lang in i18n.SUPPORTED:
            path = os.path.join(LOCALES_DIR, f"{lang}.json")
            with open(path, encoding="utf-8") as f:
                raw = f.read()
            data = json.loads(raw)
            dumped = json.dumps(data, ensure_ascii=False)
            again = json.loads(dumped)
            self.assertEqual(data, again)


class TestCache(unittest.TestCase):
    def test_load_caches_same_dict(self):
        i18n._CACHE.clear()
        a = i18n._load("en")
        b = i18n._load("en")
        self.assertIs(a, b)

    def test_cache_clear_reloads(self):
        i18n._CACHE.clear()
        first = i18n._load("de")
        i18n._CACHE.clear()
        second = i18n._load("de")
        self.assertEqual(first, second)
        self.assertIsNot(first, second)

    def test_unsupported_does_not_pollute_cache_with_bad_code(self):
        i18n._CACHE.clear()
        i18n._load("nope")
        self.assertIn("tr", i18n._CACHE)
        self.assertNotIn("nope", i18n._CACHE)


class TestDetectSystemLang(unittest.TestCase):
    def test_maps_common_locales(self):
        cases = [
            (("ru_RU", "UTF-8"), "ru"),
            (("de_DE", "UTF-8"), "de"),
            (("en_US", "UTF-8"), "en"),
            (("en_GB", "UTF-8"), "en"),
            (("tr_TR", "UTF-8"), "tr"),
            (("ru", None), "ru"),
            (("DE", None), "de"),
            (("en-US", None), "en"),  # hyphen normalized via replace on getlocale? 
        ]
        for loc_tuple, expected in cases:
            with self.subTest(loc=loc_tuple, expected=expected):
                # getlocale may return en-US style; detect lowercases and replaces -
                normalized = (loc_tuple[0].replace("-", "_"), loc_tuple[1])
                with mock.patch("locale.getlocale", return_value=normalized):
                    with mock.patch("locale.getdefaultlocale", return_value=(None, None)):
                        # For en-US input: we pass already normalized in most cases
                        if loc_tuple[0] == "en-US":
                            with mock.patch("locale.getlocale", return_value=("en-US", None)):
                                self.assertEqual(i18n.detect_system_lang(), "en")
                        else:
                            self.assertEqual(i18n.detect_system_lang(), expected)

    def test_unknown_locale_defaults_to_tr(self):
        with mock.patch("locale.getlocale", return_value=("ja_JP", "UTF-8")):
            with mock.patch("locale.getdefaultlocale", return_value=(None, None)):
                self.assertEqual(i18n.detect_system_lang(), "tr")

    def test_empty_locale_falls_back_getdefaultlocale(self):
        with mock.patch("locale.getlocale", return_value=(None, None)):
            with mock.patch("locale.getdefaultlocale", return_value=("de_AT", "UTF-8")):
                self.assertEqual(i18n.detect_system_lang(), "de")

    def test_both_locale_apis_fail_defaults_tr(self):
        with mock.patch("locale.getlocale", side_effect=RuntimeError("x")):
            with mock.patch("locale.getdefaultlocale", side_effect=RuntimeError("y")):
                self.assertEqual(i18n.detect_system_lang(), "tr")


class TestFrozenLocalesDir(unittest.TestCase):
    def test_meipass_used_when_frozen(self):
        # resource_root() uses paths.sys; MEIPASS = ROOT → locales = ROOT/locales
        import paths

        with mock.patch.object(paths.sys, "frozen", True, create=True):
            with mock.patch.object(paths.sys, "_MEIPASS", ROOT, create=True):
                path = i18n._locales_dir()
                self.assertEqual(
                    os.path.normcase(path),
                    os.path.normcase(LOCALES_DIR),
                )
                self.assertTrue(os.path.isdir(path))


class TestLoadEdgeCases(unittest.TestCase):
    def test_load_none_uses_default(self):
        i18n._CACHE.clear()
        data = i18n._load(None)
        self.assertEqual(data, i18n._load("tr"))

    def test_every_key_translatable_without_exception(self):
        sample_kwargs = {
            "which": 1, "fmt": "", "pct": 0, "name": "n.xlsx", "path": "p",
            "err": "e", "n": 0, "k": "1", "d": 0.0, "m": 0, "mm": 0, "o1": 0, "o2": 0,
        }
        for lang in i18n.SUPPORTED:
            for key in i18n.locale_keys(lang):
                # kwargs fazla gonderilse bile format gerekmeyenlerde sorun yok;
                # format gerekenlerde gerekli alanlar var
                text = i18n.t(key, lang=lang, **sample_kwargs)
                self.assertIsInstance(text, str)
                self.assertTrue(text)


if __name__ == "__main__":
    unittest.main()

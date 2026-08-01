# -*- coding: utf-8 -*-
"""Frozen-safe writable vs resource path helpers."""
import os
import sys
import unittest
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import paths


class TestAppRoot(unittest.TestCase):
    def test_source_uses_package_dir(self):
        self.assertEqual(
            os.path.normcase(paths.app_root()),
            os.path.normcase(ROOT),
        )

    def test_frozen_uses_exe_dir_not_meipass(self):
        fake_exe = r"C:\Program Files\App\tool.exe"
        fake_meipass = r"C:\Users\tmp\_MEI12345"
        with mock.patch.object(paths.sys, "frozen", True, create=True):
            with mock.patch.object(paths.sys, "executable", fake_exe):
                with mock.patch.object(paths.sys, "_MEIPASS", fake_meipass, create=True):
                    root = paths.app_root()
                    self.assertEqual(
                        os.path.normcase(root),
                        os.path.normcase(r"C:\Program Files\App"),
                    )
                    self.assertNotEqual(
                        os.path.normcase(root),
                        os.path.normcase(fake_meipass),
                    )


class TestResourceRoot(unittest.TestCase):
    def test_source_uses_package_dir(self):
        self.assertEqual(
            os.path.normcase(paths.resource_root()),
            os.path.normcase(ROOT),
        )

    def test_frozen_uses_meipass(self):
        fake_meipass = os.path.join(ROOT, "fake_meipass")
        with mock.patch.object(paths.sys, "frozen", True, create=True):
            with mock.patch.object(paths.sys, "_MEIPASS", fake_meipass, create=True):
                self.assertEqual(
                    os.path.normcase(paths.resource_root()),
                    os.path.normcase(fake_meipass),
                )


class TestLogAndArchiveDefaults(unittest.TestCase):
    def test_log_dir_under_app_root(self):
        import app_log

        self.assertEqual(
            os.path.normcase(app_log.LOG_DIR),
            os.path.normcase(os.path.join(paths.app_root(), "logs")),
        )

    def test_archive_db_under_app_root(self):
        import archive_db

        self.assertEqual(
            os.path.normcase(archive_db.DEFAULT_DB),
            os.path.normcase(
                os.path.join(paths.app_root(), "data", "compare_archive.db")
            ),
        )


if __name__ == "__main__":
    unittest.main()

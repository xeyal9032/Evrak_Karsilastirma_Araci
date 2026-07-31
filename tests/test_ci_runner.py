# -*- coding: utf-8 -*-
"""ci_runner smoke/full rapor uretimi."""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from tools import ci_runner  # noqa: E402


class TestCiRunner(unittest.TestCase):
    def test_smoke_writes_summary_and_junit(self):
        with tempfile.TemporaryDirectory() as tmp:
            code = ci_runner.main(["--suite", "smoke", "--out-dir", tmp, "--quiet"])
            self.assertEqual(code, 0)
            summary = os.path.join(tmp, "summary.md")
            junit = os.path.join(tmp, "junit.xml")
            self.assertTrue(os.path.isfile(summary))
            self.assertTrue(os.path.isfile(junit))
            with open(summary, encoding="utf-8") as fh:
                text = fh.read()
            self.assertIn("Verdict: **PASS**", text)
            self.assertIn("CI Report", text)
            with open(junit, encoding="utf-8") as fh:
                xml = fh.read()
            self.assertIn("<testsuite", xml)
            self.assertIn("testcase", xml)


if __name__ == "__main__":
    unittest.main()

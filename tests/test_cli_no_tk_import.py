# -*- coding: utf-8 -*-
"""CLI entry must not require Tkinter at import time."""
import importlib
import sys
import unittest


class TestCliImportNoTk(unittest.TestCase):
    def test_importing_karsilastir_does_not_load_tkinter(self):
        # Drop cached modules so this test is meaningful even if GUI ran earlier
        # in the same process (full suite). Prefer a subprocess-like isolation via
        # checking karsilastir module dict after fresh import when possible.
        for name in list(sys.modules):
            if name == "karsilastir" or name.startswith("karsilastir."):
                del sys.modules[name]
            if name == "karsilastir_gui":
                del sys.modules[name]
        # If tkinter already loaded by other tests, we still assert karsilastir
        # itself does not import it as a dependency of the CLI module body.
        import karsilastir as k
        self.assertNotIn("tkinter", k.__dict__)
        self.assertTrue(hasattr(k, "run_cli"))
        self.assertNotIn("karsilastir_gui", sys.modules)


if __name__ == "__main__":
    unittest.main()

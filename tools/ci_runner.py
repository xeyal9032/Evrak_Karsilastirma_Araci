#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CI test runner: unittest + Markdown/JUnit rapor.

Kullanim:
  python tools/ci_runner.py
  python tools/ci_runner.py --suite smoke
  python tools/ci_runner.py --suite full --out-dir ci-artifacts

Cikti:
  <out-dir>/summary.md   -> GitHub Step Summary / Cursor icin okunabilir rapor
  <out-dir>/junit.xml    -> GitHub test annotation / artifact
  <out-dir>/failures.txt -> sadece basarisiz testler (kisa)
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import traceback
import unittest
import xml.etree.ElementTree as ET
from datetime import datetime, timezone


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SMOKE_MODULES = [
    "tests.test_report_i18n_contract",
    "tests.test_match_fix",
    "tests.test_normalize",
    "tests.test_matching",
    "tests.test_faz23",
    "tests.test_i18n",
]


class _Result(unittest.TextTestResult):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.successes = []

    def addSuccess(self, test):
        super().addSuccess(test)
        self.successes.append(test)


def _test_id(test):
    return test.id()


def _exc_text(err):
    return "".join(traceback.format_exception(*err))


def _load_suite(kind: str) -> unittest.TestSuite:
    loader = unittest.TestLoader()
    if kind == "smoke":
        suite = unittest.TestSuite()
        for mod in SMOKE_MODULES:
            try:
                suite.addTests(loader.loadTestsFromName(mod))
            except Exception as ex:  # noqa: BLE001 - smoke must report load errors
                print(f"WARN: could not load {mod}: {ex}", file=sys.stderr)
        return suite
    return loader.discover(os.path.join(ROOT, "tests"), pattern="test_*.py")


def _write_junit(path: str, result: _Result, elapsed: float, suite_name: str):
    tests = (
        result.testsRun
    )
    failures = len(result.failures)
    errors = len(result.errors)
    skipped = len(result.skipped)

    root = ET.Element(
        "testsuite",
        name=suite_name,
        tests=str(tests),
        failures=str(failures),
        errors=str(errors),
        skipped=str(skipped),
        time=f"{elapsed:.3f}",
        timestamp=datetime.now(timezone.utc).isoformat(timespec="seconds"),
    )

    seen = set()

    def add_case(test, status=None, message="", detail=""):
        tid = _test_id(test)
        if tid in seen:
            return
        seen.add(tid)
        parts = tid.rsplit(".", 2)
        classname = ".".join(parts[:-1]) if len(parts) > 1 else tid
        name = parts[-1] if parts else tid
        case = ET.SubElement(
            root, "testcase", classname=classname, name=name, time="0"
        )
        if status == "failure":
            el = ET.SubElement(case, "failure", message=message[:300])
            el.text = detail
        elif status == "error":
            el = ET.SubElement(case, "error", message=message[:300])
            el.text = detail
        elif status == "skipped":
            ET.SubElement(case, "skipped", message=str(message)[:300])

    for test in result.successes:
        add_case(test)
    for test, reason in result.skipped:
        add_case(test, "skipped", str(reason))
    for test, err in result.failures:
        add_case(test, "failure", str(err).splitlines()[-1] if err else "fail", err)
    for test, err in result.errors:
        add_case(test, "error", str(err).splitlines()[-1] if err else "error", err)

    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    ET.ElementTree(root).write(path, encoding="utf-8", xml_declaration=True)


def _write_summary(path: str, result: _Result, elapsed: float, suite_name: str, os_name: str, py_ver: str):
    failed = result.failures + result.errors
    lines = [
        f"# CI Report — `{suite_name}`",
        "",
        f"- Time (UTC): {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')}",
        f"- OS: `{os_name}`",
        f"- Python: `{py_ver}`",
        f"- Elapsed: **{elapsed:.1f}s**",
        f"- Ran: **{result.testsRun}**",
        f"- OK: **{len(result.successes)}**",
        f"- Failed: **{len(result.failures)}**",
        f"- Errors: **{len(result.errors)}**",
        f"- Skipped: **{len(result.skipped)}**",
        f"- Verdict: **{'PASS' if result.wasSuccessful() else 'FAIL'}**",
        "",
    ]
    if failed:
        lines.append("## Failures / Errors")
        lines.append("")
        for test, err in failed:
            kind = "FAIL" if (test, err) in result.failures else "ERROR"
            # failures stored as list of tuples - membership by identity may fail; use ids
            kind = "FAIL"
            for t, e in result.errors:
                if t is test and e is err:
                    kind = "ERROR"
                    break
            lines.append(f"### `{_test_id(test)}` ({kind})")
            lines.append("")
            lines.append("```")
            # keep last ~40 lines of traceback for readable summary
            tb_lines = (err or "").strip().splitlines()
            lines.extend(tb_lines[-40:])
            lines.append("```")
            lines.append("")
    else:
        lines.append("## All tests passed")
        lines.append("")
        lines.append("No failures. Contract / suite green.")
        lines.append("")

    if result.skipped:
        lines.append("## Skipped")
        lines.append("")
        for test, reason in result.skipped[:30]:
            lines.append(f"- `{_test_id(test)}`: {reason}")
        lines.append("")

    lines.append("---")
    lines.append("_Generated by `tools/ci_runner.py` — share this file with Cursor for triage._")
    lines.append("")

    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def _write_failures_txt(path: str, result: _Result):
    rows = []
    for test, err in result.failures + result.errors:
        rows.append(_test_id(test))
        rows.append((err or "").strip().splitlines()[-1] if err else "")
        rows.append("")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(rows) if rows else "none\n")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="CI unittest runner with markdown/junit report")
    ap.add_argument("--suite", choices=("smoke", "full"), default="full")
    ap.add_argument("--out-dir", default=os.path.join(ROOT, "ci-artifacts"))
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    os.chdir(ROOT)
    if ROOT not in sys.path:
        sys.path.insert(0, ROOT)

    suite = _load_suite(args.suite)
    stream = open(os.devnull, "w", encoding="utf-8") if args.quiet else sys.stderr
    runner = unittest.TextTestRunner(stream=stream, verbosity=0 if args.quiet else 2, resultclass=_Result)

    t0 = time.time()
    result = runner.run(suite)
    elapsed = time.time() - t0
    if args.quiet:
        stream.close()

    os_name = f"{os.name}/{sys.platform}"
    py_ver = sys.version.split()[0]
    out = args.out_dir
    os.makedirs(out, exist_ok=True)

    summary_path = os.path.join(out, "summary.md")
    junit_path = os.path.join(out, "junit.xml")
    fail_path = os.path.join(out, "failures.txt")

    _write_summary(summary_path, result, elapsed, args.suite, os_name, py_ver)
    _write_junit(junit_path, result, elapsed, args.suite)
    _write_failures_txt(fail_path, result)

    # GitHub Actions job summary
    gh_sum = os.environ.get("GITHUB_STEP_SUMMARY")
    if gh_sum:
        with open(summary_path, encoding="utf-8") as src, open(gh_sum, "a", encoding="utf-8") as dst:
            dst.write(src.read())
            dst.write("\n")

    print(f"CI_REPORT={summary_path}")
    print(f"JUNIT={junit_path}")
    print(
        f"RESULT={'PASS' if result.wasSuccessful() else 'FAIL'} "
        f"ran={result.testsRun} fail={len(result.failures)} err={len(result.errors)} "
        f"skip={len(result.skipped)} time={elapsed:.1f}s"
    )
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())

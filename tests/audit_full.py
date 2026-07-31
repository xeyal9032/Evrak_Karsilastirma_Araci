# -*- coding: utf-8 -*-
"""Tam proje denetimi."""
import os
import sys
import tempfile
import py_compile
from datetime import datetime

ROOT = r"C:\Users\xeyal\Desktop\Evrak_Karsilastirma_Araci"
sys.path.insert(0, ROOT)

import karsilastir_motor as motor
import karsilastir as gui

issues = []
ok = []


def add_ok(msg):
    ok.append(msg)


def add_issue(msg):
    issues.append(msg)


def main():
    bat = r"C:\Users\xeyal\Desktop\Evrak_Karsilastir.bat"
    exe = os.path.join(ROOT, "dist", "Evrak_Karsilastirma_Araci.exe")
    pyw = r"C:\Python314\pythonw.exe"
    for label, p in [
        ("bat", bat),
        ("exe", exe),
        ("pythonw", pyw),
        ("gui", os.path.join(ROOT, "karsilastir.py")),
        ("motor", os.path.join(ROOT, "karsilastir_motor.py")),
        ("requirements", os.path.join(ROOT, "requirements.txt")),
        ("readme", os.path.join(ROOT, "README.md")),
        ("spec", os.path.join(ROOT, "Evrak_Karsilastirma_Araci.spec")),
    ]:
        if os.path.isfile(p):
            add_ok(f"exists:{label}")
        else:
            add_issue(f"MISSING:{label}:{p}")

    src_m = max(
        os.path.getmtime(os.path.join(ROOT, "karsilastir.py")),
        os.path.getmtime(os.path.join(ROOT, "karsilastir_motor.py")),
    )
    exe_m = os.path.getmtime(exe)
    if exe_m < src_m - 1:
        add_issue(
            f"EXE_STALE: exe={datetime.fromtimestamp(exe_m)} "
            f"src={datetime.fromtimestamp(src_m)}"
        )
    else:
        add_ok("exe_fresh_vs_src")

    with open(bat, encoding="utf-8", errors="replace") as f:
        bat_txt = f.read()
    if "karsilastir.py" in bat_txt and "pythonw.exe" in bat_txt:
        add_ok("bat_points_to_source")
    else:
        add_issue("bat_unexpected_content")

    import openpyxl
    import tkinter  # noqa: F401

    add_ok(f"openpyxl={openpyxl.__version__}")
    add_ok("tkinter_ok")

    d = gui.desktop_dir()
    if not os.path.isdir(d):
        add_issue(f"desktop_dir_invalid:{d}")
    else:
        add_ok(f"desktop={d}")

    mart = r"C:\Users\xeyal\Desktop\excel1\mart"
    if os.path.isdir(mart):
        files = os.listdir(mart)
        csvs = [f for f in files if f.lower().endswith(".csv")]
        xlsxs = [f for f in files if f.lower().endswith(".xlsx")]
        if csvs and xlsxs:
            cpath = os.path.join(mart, csvs[0])
            xpath = os.path.join(mart, xlsxs[0])
            f1 = motor.detect_format(cpath)
            f2 = motor.detect_format(xpath)
            add_ok(f"mart_fmt={f1}+{f2}")
            with tempfile.TemporaryDirectory() as tmp:
                out = os.path.join(tmp, "audit.xlsx")
                r = motor.build_report(cpath, xpath, out, f1_label="DATEV", f2_label="Journal")
                add_ok(
                    f"mart totals {r['f1_total']}/{r['f2_total']} "
                    f"match={r['match_count']} mism={r['mismatch_count']} "
                    f"o1={r['only1_count']} o2={r['only2_count']}"
                )
                e1, e2, *_ = motor.compare(cpath, xpath)
                bad_amt = sum(
                    1
                    for e in e1
                    if e.get("pair_kind") == "MATCH"
                    and not e["netted"]
                    and e["amt"] != e["pair"]["amt"]
                )
                if bad_amt:
                    add_issue(f"mart_MATCH_amount_mismatch_count={bad_amt}")
                else:
                    add_ok("mart_all_MATCH_amounts_equal")

                bew = [
                    e
                    for e in e1
                    if e["text"] == "Aral Tankstelle Leonart Bewirtungskosten (nur MA)"
                ]
                kraft = [
                    e
                    for e in e1
                    if e["text"] == "Aral Tankstelle Leonart Kraftstoff/Ladestrom"
                ]
                if len(bew) == 1 and bew[0]["amt"] == 1.99 and bew[0]["pair_kind"] == "MATCH":
                    add_ok("mart_bewirtung_ok")
                else:
                    add_issue(f"mart_bewirtung_bad:{[(b['amt'], b.get('pair_kind')) for b in bew]}")
                if (
                    len(kraft) == 1
                    and abs(kraft[0]["amt"] - 110.7) < 0.001
                    and kraft[0]["pair_kind"] == "MATCH"
                ):
                    add_ok("mart_kraftstoff_ok")
                else:
                    add_issue(f"mart_kraftstoff_bad:{kraft}")

                if r["match_count"] + r["mismatch_count"] + r["only1_count"] != r["f1_total"] - 2 * r["net1"]:
                    add_issue("mart_group_count_inconsistent")
                else:
                    add_ok("mart_group_counts_consistent")
        else:
            add_issue("mart_files_incomplete")
    else:
        add_issue("mart_missing")

    ex = r"C:\Users\xeyal\Desktop\excel1"
    f_bs = os.path.join(ex, "EXTF_Buchungsstapel_01012026_bis_30062026.csv")
    f_mb = os.path.join(ex, "EXTF_MB_01-06.csv")
    if os.path.isfile(f_bs) and os.path.isfile(f_mb):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "two.xlsx")
            r = motor.build_report(f_bs, f_mb, out, f1_label="BS", f2_label="MB")
            add_ok(f"two_datev match={r['match_count']} f1={r['f1_total']} f2={r['f2_total']}")
            if r["match_count"] < 100:
                add_issue("two_datev_too_few_matches")
    else:
        add_issue("two_datev_missing")

    kars = os.path.join(ex, "KARSILASTIRMA_ISARETLI.xlsx")
    if os.path.isfile(kars):
        try:
            motor.detect_format(kars)
            add_issue("should_reject_KARSILASTIRMA_output")
        except motor.FormatError:
            add_ok("rejects_report_xlsx")

    for name in ("karsilastir.py", "karsilastir_motor.py"):
        py_compile.compile(os.path.join(ROOT, name), doraise=True)
    add_ok("py_compile_ok")

    # required symbols
    for sym in (
        "detect_format", "load", "compare", "build_report",
        "FORMAT_DATEV_CSV", "FORMAT_JOURNAL_XLSX", "FormatError",
    ):
        if not hasattr(motor, sym):
            add_issue(f"missing_symbol:{sym}")
        else:
            add_ok(f"symbol:{sym}")

    print("=== OK ===")
    for x in ok:
        print("+", x)
    print("=== ISSUES ===")
    if not issues:
        print("(none)")
    else:
        for x in issues:
            print("!", x)
    print("ISSUE_COUNT", len(issues))
    return 1 if issues else 0


if __name__ == "__main__":
    raise SystemExit(main())

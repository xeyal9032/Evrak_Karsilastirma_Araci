# -*- coding: utf-8 -*-
"""
Benchmark: generate N synthetic DATEV rows and time compare/build_report.

Usage:
  python benchmarks/run_benchmark.py
  python benchmarks/run_benchmark.py --rows 100000 --fast
"""
import argparse
import os
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from tests.helpers_datev import drow, write_datev_csv
import karsilastir_motor as motor


def make_pair(tmp, n):
    f1 = os.path.join(tmp, "a.csv")
    f2 = os.path.join(tmp, "b.csv")
    rows1 = []
    rows2 = []
    for i in range(n):
        beleg = f"B{i:06d}"
        amt = f"{(i % 500) + 1},00"
        text = f"Buchung {i}"
        rows1.append(drow(amt, "S", str(1000 + (i % 50)), "2000", "0101", beleg, text))
        # 90% exact match, 5% amount drift, 5% only-one-side
        mod = i % 20
        if mod == 0:
            rows2.append(drow(f"{(i % 500) + 2},00", "S", str(1000 + (i % 50)), "2000", "0101", beleg, text))
        elif mod == 1:
            rows1_only = True
            # skip adding to file2 -> only1
            continue
        else:
            rows2.append(drow(amt, "S", str(1000 + (i % 50)), "2000", "0101", beleg, text))
    # add some only2
    for j in range(max(1, n // 50)):
        rows2.append(drow("9,00", "S", "9999", "1", "0202", f"ONLY2-{j}", "extra"))
    write_datev_csv(f1, rows1)
    write_datev_csv(f2, rows2)
    return f1, f2


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=10000)
    ap.add_argument("--fast", action="store_true")
    ap.add_argument("--html", action="store_true")
    args = ap.parse_args()
    with tempfile.TemporaryDirectory() as tmp:
        print(f"Generating ~{args.rows} rows...")
        t0 = time.perf_counter()
        f1, f2 = make_pair(tmp, args.rows)
        gen_s = time.perf_counter() - t0
        out = os.path.join(tmp, "bench.xlsx")
        print(f"Generated in {gen_s:.2f}s. Comparing...")
        t1 = time.perf_counter()
        result = motor.build_report(
            f1, f2, out,
            detail_sheets=not args.fast,
            write_html=args.html,
        )
        cmp_s = time.perf_counter() - t1
        print("--- benchmark ---")
        print(f"rows_requested={args.rows}")
        print(f"f1_total={result['f1_total']} f2_total={result['f2_total']}")
        print(f"match={result['match_count']} mismatch={result['mismatch_count']} "
              f"only1={result['only1_count']} only2={result['only2_count']}")
        print(f"compare_build_seconds={cmp_s:.3f}")
        print(f"fast={args.fast} html={args.html}")
        print(f"xlsx_bytes={os.path.getsize(out)}")
        if cmp_s > 0:
            print(f"rows_per_second~={(result['f1_total'] + result['f2_total']) / cmp_s:.0f}")


if __name__ == "__main__":
    main()

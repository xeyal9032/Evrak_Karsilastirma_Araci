# -*- coding: utf-8 -*-
"""Batch / folder comparison helpers."""
import os
from collections import defaultdict

SUPPORTED_EXT = {".csv", ".xlsx", ".xlsm"}


def list_document_files(folder):
    files = []
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        if os.path.splitext(name)[1].lower() in SUPPORTED_EXT:
            files.append(path)
    return files


def pair_by_stem(folder_a, folder_b):
    """Pair files with the same stem (basename without extension) across two folders."""
    a_map = defaultdict(list)
    b_map = defaultdict(list)
    for p in list_document_files(folder_a):
        stem = os.path.splitext(os.path.basename(p))[0].lower()
        a_map[stem].append(p)
    for p in list_document_files(folder_b):
        stem = os.path.splitext(os.path.basename(p))[0].lower()
        b_map[stem].append(p)
    pairs = []
    unmatched_a = []
    unmatched_b = []
    for stem, paths in sorted(a_map.items()):
        if stem in b_map and b_map[stem]:
            pairs.append((paths[0], b_map[stem][0], stem))
        else:
            unmatched_a.extend(paths)
    for stem, paths in sorted(b_map.items()):
        if stem not in a_map:
            unmatched_b.extend(paths)
    return pairs, unmatched_a, unmatched_b


def pair_zip_sorted(folder_a, folder_b):
    """Pair by sorted order (min length)."""
    a = list_document_files(folder_a)
    b = list_document_files(folder_b)
    n = min(len(a), len(b))
    pairs = [(a[i], b[i], f"pair_{i+1}") for i in range(n)]
    return pairs, a[n:], b[n:]


def run_batch(folder_a, folder_b, out_dir, *, mode="stem", progress_cb=None,
              write_html=False, write_pdf=False, archive_db=None, detail_sheets=True):
    """
    Compare paired files from two folders. Returns list of result dicts.
    mode: 'stem' (same basename) or 'zip' (sorted order).
    """
    import karsilastir_motor as motor
    import archive_db as arch
    from report_extra import write_html_report, write_pdf_report

    os.makedirs(out_dir, exist_ok=True)
    if mode == "zip":
        pairs, ua, ub = pair_zip_sorted(folder_a, folder_b)
    else:
        pairs, ua, ub = pair_by_stem(folder_a, folder_b)

    results = []
    total = max(len(pairs), 1)
    for i, (f1, f2, stem) in enumerate(pairs):
        if progress_cb:
            progress_cb(int(i * 100 / total), f"batch:{stem}")
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in stem)[:40]
        xlsx = os.path.join(out_dir, f"Karsilastirma_{safe}.xlsx")
        name1 = os.path.splitext(os.path.basename(f1))[0][:31]
        name2 = os.path.splitext(os.path.basename(f2))[0][:31]
        result = motor.build_report(
            f1, f2, xlsx,
            f1_label=name1, f2_label=name2,
            write_html=write_html,
            write_pdf=write_pdf,
            detail_sheets=detail_sheets,
        )
        result["batch_stem"] = stem
        result["unmatched_a"] = ua
        result["unmatched_b"] = ub
        if archive_db:
            arch.save_comparison(
                result, file1=f1, file2=f2,
                f1_label=name1, f2_label=name2,
                source="batch", db_path=archive_db,
                extra={"stem": stem},
            )
        results.append(result)
    if progress_cb:
        progress_cb(100, "batch_done")
    return {
        "pairs": results,
        "unmatched_a": ua,
        "unmatched_b": ub,
        "pair_count": len(pairs),
    }

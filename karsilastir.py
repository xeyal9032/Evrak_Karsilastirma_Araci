# -*- coding: utf-8 -*-
"""
Evrak Karsilastirma Araci - GUI + CLI (TR / RU / DE / EN).

GUI:  python karsilastir.py
CLI:  python karsilastir.py file1.csv file2.xlsx [-o out.xlsx] [--lang en]
"""
import argparse
import os
import sys
import time
import traceback
import datetime
import platform
import subprocess

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import karsilastir_motor as motor
import i18n
import app_log

FILETYPES = [
    ("Documents", "*.csv *.xlsx *.xlsm"),
    ("CSV", "*.csv"),
    ("Excel", "*.xlsx *.xlsm"),
    ("All files", "*.*"),
]


def desktop_dir():
    home = os.path.expanduser("~")
    for name in ("Desktop", "Masaüstü", "Рабочий стол"):
        path = os.path.join(home, name)
        if os.path.isdir(path):
            return path
    return home


def open_file(path):
    system = platform.system()
    if system == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
    elif system == "Darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


DESKTOP = desktop_dir()



def _cli_err(msg, code=1):
    print(msg, file=sys.stderr)
    return code


def _ensure_cli_console():
    """Windowed frozen exe has no console — attach/allocate one for CLI I/O."""
    if not getattr(sys, "frozen", False):
        return
    if sys.platform != "win32":
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        # Attach to parent console (cmd/PowerShell) if present; else create one.
        if kernel32.GetConsoleWindow() == 0:
            if kernel32.AttachConsole(-1) == 0:
                kernel32.AllocConsole()
        sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
        sys.stderr = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
        try:
            sys.stdin = open("CONIN$", "r", encoding="utf-8", errors="replace")
        except OSError:
            pass
    except Exception:
        pass


def _require_path(path, *, as_dir=False):
    if as_dir:
        if not os.path.isdir(path):
            raise FileNotFoundError(
                f"Folder not found: {path}\n"
                f"Klasor bulunamadi / Папка не найдена.\n"
                f"Example: python karsilastir.py tests\\fixtures_batch_a tests\\fixtures_batch_b --batch -o batch_out"
            )
    else:
        if not os.path.isfile(path):
            raise FileNotFoundError(
                f"File not found: {path}\n"
                f"Dosya bulunamadi / Файл не найден.\n"
                f"Example: python karsilastir.py tests\\fixtures\\sample_a.csv tests\\fixtures\\sample_b.csv -o out.xlsx --html"
            )


def run_cli(argv=None):
    _ensure_cli_console()
    import archive_db

    parser = argparse.ArgumentParser(
        prog="Evrak_Karsilastirma_Araci",
        description="Compare DATEV / Journal exports and write Excel/HTML/PDF reports.",
        epilog=(
            "Examples:\n"
            "  python karsilastir.py tests/fixtures/sample_a.csv tests/fixtures/sample_b.csv -o out.xlsx --html --pdf --archive\n"
            "  python karsilastir.py folder_a folder_b --batch -o batch_out --fast --html\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("file1", help="First document or folder (with --batch)")
    parser.add_argument("file2", help="Second document or folder (with --batch)")
    parser.add_argument("-o", "--output", help="Output Excel path (or directory with --batch)")
    parser.add_argument("--lang", default="en", choices=list(i18n.SUPPORTED))
    parser.add_argument("--quiet", action="store_true", help="Suppress progress on stderr")
    parser.add_argument("--html", action="store_true", help="Also write HTML summary")
    parser.add_argument("--pdf", action="store_true", help="Also write PDF summary (needs reportlab)")
    parser.add_argument(
        "--archive", metavar="DB", nargs="?", const=archive_db.DEFAULT_DB,
        help=f"Save run into SQLite archive (default: {archive_db.DEFAULT_DB})",
    )
    parser.add_argument("--batch", action="store_true",
                        help="Treat file1/file2 as folders and compare paired files")
    parser.add_argument("--batch-mode", choices=("stem", "zip"), default="stem",
                        help="Pairing strategy for --batch (default: stem)")
    parser.add_argument("--fast", action="store_true",
                        help="Skip per-file detail sheets (faster for large files / batch)")
    args = parser.parse_args(argv)

    lang = args.lang

    def progress_cb(pct, stage):
        if not args.quiet:
            print(f"\r{pct:3d}% {stage}", end="", file=sys.stderr, flush=True)

    if args.batch:
        out_dir = args.output or os.path.join(os.getcwd(), "batch_out")
        try:
            _require_path(args.file1, as_dir=True)
            _require_path(args.file2, as_dir=True)
            import batch_compare
            t0 = time.time()
            summary = batch_compare.run_batch(
                args.file1, args.file2, out_dir,
                mode=args.batch_mode,
                progress_cb=progress_cb,
                write_html=args.html,
                write_pdf=args.pdf,
                archive_db=args.archive,
                detail_sheets=not args.fast,
                lang=lang,
            )
            if not args.quiet:
                print(file=sys.stderr)
            for w in summary.get("warnings") or []:
                print(f"warning: {w}", file=sys.stderr)
            print(f"OK batch pairs={summary['pair_count']} out={out_dir}")
            print(f"unmatched_a={len(summary['unmatched_a'])} unmatched_b={len(summary['unmatched_b'])}")
            print(f"elapsed={time.time() - t0:.2f}s")
            return 0
        except (motor.FormatError, FileNotFoundError, OSError, PermissionError) as ex:
            if not args.quiet:
                print(file=sys.stderr)
            app_log.log_run_err(ex, traceback.format_exc())
            return _cli_err(str(ex), 1)
        except Exception as ex:
            if not args.quiet:
                print(file=sys.stderr)
            app_log.log_run_err(ex, traceback.format_exc())
            return _cli_err(str(ex), 2)

    f1, f2 = args.file1, args.file2
    try:
        _require_path(f1)
        _require_path(f2)
    except FileNotFoundError as ex:
        app_log.log_run_err(ex)
        return _cli_err(str(ex), 1)

    if not args.output:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(os.getcwd(), f"Karsilastirma_{ts}.xlsx")
    else:
        out_path = args.output

    name1 = os.path.splitext(os.path.basename(f1))[0][:31]
    name2 = os.path.splitext(os.path.basename(f2))[0][:31]

    fmt1 = fmt2 = None
    try:
        fmt1 = motor.detect_format(f1)
        fmt2 = motor.detect_format(f2)
    except (motor.FormatError, FileNotFoundError, OSError) as ex:
        app_log.log_run_err(ex)
        return _cli_err(str(ex), 1)

    app_log.log_run_start(f1, f2, fmt1, fmt2, source="cli")
    t0 = time.time()

    try:
        result = motor.build_report(
            f1, f2, out_path,
            f1_label=name1, f2_label=name2,
            progress_cb=progress_cb,
            write_html=args.html,
            write_pdf=args.pdf,
            detail_sheets=not args.fast,
            lang=lang,
        )
    except (motor.FormatError, FileNotFoundError, OSError, PermissionError) as ex:
        if not args.quiet:
            print(file=sys.stderr)
        app_log.log_run_err(ex, traceback.format_exc())
        return _cli_err(str(ex), 1)
    except Exception as ex:
        if not args.quiet:
            print(file=sys.stderr)
        app_log.log_run_err(ex, traceback.format_exc())
        return _cli_err(str(ex), 2)

    if not args.quiet:
        print(file=sys.stderr)
    elapsed = time.time() - t0
    app_log.log_run_ok(result, elapsed, out_path)
    for w in result.get("warnings") or []:
        print(f"warning: {w}", file=sys.stderr)
    if args.archive:
        try:
            import archive_db
            archive_db.save_comparison(
                result, file1=f1, file2=f2,
                f1_label=name1, f2_label=name2,
                elapsed_s=elapsed, source="cli", db_path=args.archive,
            )
        except Exception as ex:
            print(f"archive warning: {ex}", file=sys.stderr)
    print(i18n.t("cli_ok", lang=lang, path=out_path))
    if result.get("html_path"):
        print(f"HTML -> {result['html_path']}")
    if result.get("pdf_path"):
        print(f"PDF -> {result['pdf_path']}")
    print(i18n.t(
        "cli_counts", lang=lang,
        m=result["match_count"], mm=result["mismatch_count"],
        o1=result["only1_count"], o2=result["only2_count"],
    ))
    return 0


def _wants_cli(argv):
    if len(argv) == 0:
        return False
    if argv[0] in ("-h", "--help"):
        return True
    if argv[0].startswith("-"):
        return True
    return True


def __getattr__(name):
    if name == "App":
        from karsilastir_gui import App as _App
        return _App
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def main():
    argv = sys.argv[1:]
    if argv and _wants_cli(argv):
        raise SystemExit(run_cli(argv))
    from karsilastir_gui import App
    App().mainloop()


if __name__ == "__main__":
    main()

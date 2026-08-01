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
import threading
import platform
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.lang = tk.StringVar(value=i18n.detect_system_lang())
        self._busy = False
        self._fmt1 = None
        self._fmt2 = None
        self._progress_pct = 0

        self.file1 = tk.StringVar()
        self.file2 = tk.StringVar()

        self.geometry("720x560")
        self.resizable(False, False)

        top = tk.Frame(self)
        top.pack(fill="x", padx=20, pady=(12, 0))
        self.lbl_lang = tk.Label(top, text="")
        self.lbl_lang.pack(side="right", padx=(8, 0))
        self.lang_box = ttk.Combobox(
            top, textvariable=self.lang, values=list(i18n.SUPPORTED),
            width=5, state="readonly",
        )
        self.lang_box.pack(side="right")
        self.lang_box.bind("<<ComboboxSelected>>", lambda _e: self._apply_lang())

        self.title_lbl = tk.Label(self, font=("Segoe UI", 13, "bold"), wraplength=660, justify="center")
        self.title_lbl.pack(pady=(8, 4))
        self.sub_lbl = tk.Label(self, fg="#555", wraplength=660, justify="center")
        self.sub_lbl.pack(pady=(0, 8))
        self.drop_lbl = tk.Label(self, fg="#888", wraplength=660, justify="center")
        self.drop_lbl.pack(pady=(0, 8))

        f1_frame = tk.Frame(self)
        f1_frame.pack(fill="x", padx=20, pady=6)
        self.btn1 = tk.Button(f1_frame, width=22, command=self.pick_file1)
        self.btn1.pack(side="left")
        tk.Label(f1_frame, textvariable=self.file1, anchor="w", fg="#333").pack(
            side="left", padx=10, fill="x", expand=True
        )

        f2_frame = tk.Frame(self)
        f2_frame.pack(fill="x", padx=20, pady=6)
        self.btn2 = tk.Button(f2_frame, width=22, command=self.pick_file2)
        self.btn2.pack(side="left")
        tk.Label(f2_frame, textvariable=self.file2, anchor="w", fg="#333").pack(
            side="left", padx=10, fill="x", expand=True
        )

        self.status = tk.Label(self, text="", fg="#0a6")
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100, length=640)
        self.progress.pack(pady=(0, 8))

        opts = tk.Frame(self)
        opts.pack(pady=(0, 4))
        self.var_html = tk.BooleanVar(value=True)
        self.var_pdf = tk.BooleanVar(value=False)
        self.var_fast = tk.BooleanVar(value=False)
        self.var_archive = tk.BooleanVar(value=True)
        self.chk_html = tk.Checkbutton(opts, text="HTML", variable=self.var_html)
        self.chk_html.pack(side="left", padx=6)
        self.chk_pdf = tk.Checkbutton(opts, text="PDF", variable=self.var_pdf)
        self.chk_pdf.pack(side="left", padx=6)
        self.chk_fast = tk.Checkbutton(opts, text="Fast", variable=self.var_fast)
        self.chk_fast.pack(side="left", padx=6)
        self.chk_archive = tk.Checkbutton(opts, text="Archive", variable=self.var_archive)
        self.chk_archive.pack(side="left", padx=6)

        extra = tk.Frame(self)
        extra.pack(pady=(0, 4))
        self.btn_batch = tk.Button(extra, command=self.run_batch)
        self.btn_batch.pack(side="left", padx=6)
        self.btn_archive = tk.Button(extra, command=self.show_archive)
        self.btn_archive.pack(side="left", padx=6)

        self.compare_btn = tk.Button(
            self, font=("Segoe UI", 11, "bold"),
            bg="#4472C4", fg="white", width=26, height=2, command=self.run_compare,
        )
        self.compare_btn.pack(pady=10)

        self._apply_lang()
        self._setup_dnd()

    def _t(self, key, **kwargs):
        return i18n.t(key, lang=self.lang.get(), **kwargs)

    def _apply_lang(self):
        self.title(self._t("app_title"))
        self.title_lbl.config(text=self._t("app_title"))
        self.sub_lbl.config(text=self._t("app_subtitle"))
        self.drop_lbl.config(text=self._t("drop_hint"))
        self.lbl_lang.config(text=self._t("lang_label") + ":")
        self.btn1.config(text=self._t("file1"))
        self.btn2.config(text=self._t("file2"))
        self.compare_btn.config(text=self._t("compare"))
        self.chk_html.config(text=self._t("opt_html"))
        self.chk_pdf.config(text=self._t("opt_pdf"))
        self.chk_fast.config(text=self._t("opt_fast"))
        self.chk_archive.config(text=self._t("opt_archive"))
        self.btn_batch.config(text=self._t("btn_batch"))
        self.btn_archive.config(text=self._t("btn_archive"))
        self._update_format_status()

    def _setup_dnd(self):
        try:
            import windnd  # type: ignore
        except ImportError:
            return

        def on_drop(files):
            paths = []
            for f in files:
                if isinstance(f, bytes):
                    f = f.decode(sys.getfilesystemencoding(), errors="replace")
                paths.append(f)
            for p in paths:
                if not self.file1.get():
                    self._assign_path(1, p)
                elif not self.file2.get():
                    self._assign_path(2, p)
                else:
                    self._assign_path(1, p)
                    break

        windnd.hook_dropfiles(self, func=on_drop)

    def _assign_path(self, which, path):
        if not path or not os.path.isfile(path):
            return
        try:
            fmt = motor.detect_format(path)
        except (motor.FormatError, FileNotFoundError, OSError) as ex:
            messagebox.showerror(self._t("format_error_title"), str(ex))
            return
        if which == 1:
            self.file1.set(path)
            self._fmt1 = fmt
        else:
            self.file2.set(path)
            self._fmt2 = fmt
        self._update_format_status()

    def _update_format_status(self):
        if self._busy:
            return
        parts = []
        if self._fmt1:
            parts.append(f"1: {motor.format_label(self._fmt1)}")
        if self._fmt2:
            parts.append(f"2: {motor.format_label(self._fmt2)}")
        if parts:
            self.status.config(text=" / ".join(parts), fg="#06a")
        else:
            self.status.config(text="", fg="#0a6")

    def _pick(self, which):
        p = filedialog.askopenfilename(
            title=self._t("pick_title", which=which),
            filetypes=FILETYPES,
        )
        if not p:
            return
        self._assign_path(which, p)

    def pick_file1(self):
        self._pick(1)

    def pick_file2(self):
        self._pick(2)

    def _on_progress(self, pct, stage, *, lang, fmt1, fmt2):
        self._progress_pct = pct
        fmt_txt = ""
        if fmt1 and fmt2:
            fmt_txt = f" ({motor.format_label(fmt1)} ↔ {motor.format_label(fmt2)})"
        text = i18n.t("comparing", lang=lang, fmt=fmt_txt, pct=pct)
        self.after(0, lambda: self._set_progress_ui(pct, text))

    def _set_progress_ui(self, pct, text):
        self.progress["value"] = pct
        self.status.config(text=text, fg="#c60")

    def run_compare(self):
        if self._busy:
            return
        f1 = self.file1.get()
        f2 = self.file2.get()
        if not f1 or not f2:
            messagebox.showwarning(self._t("missing_title"), self._t("missing_files"))
            return

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"Karsilastirma_{ts}.xlsx"
        out_path = os.path.join(DESKTOP, out_name)
        name1 = os.path.splitext(os.path.basename(f1))[0][:31]
        name2 = os.path.splitext(os.path.basename(f2))[0][:31]
        # Snapshot Tk vars on the main thread before starting the worker.
        lang = self.lang.get()
        write_html = bool(self.var_html.get())
        write_pdf = bool(self.var_pdf.get())
        detail_sheets = not bool(self.var_fast.get())
        do_archive = bool(self.var_archive.get())
        fmt1, fmt2 = self._fmt1, self._fmt2

        self._busy = True
        self._progress_pct = 0
        self.progress["value"] = 0
        self.compare_btn.config(state="disabled")
        app_log.log_run_start(f1, f2, fmt1, fmt2, source="gui")
        t0 = time.time()

        def progress_cb(pct, stage):
            self._on_progress(pct, stage, lang=lang, fmt1=fmt1, fmt2=fmt2)

        def worker():
            try:
                result = motor.build_report(
                    f1, f2, out_path,
                    f1_label=name1, f2_label=name2,
                    progress_cb=progress_cb,
                    write_html=write_html,
                    write_pdf=write_pdf,
                    detail_sheets=detail_sheets,
                    lang=lang,
                )
                elapsed = time.time() - t0
                app_log.log_run_ok(result, elapsed, out_path)
                archive_err = None
                if do_archive:
                    try:
                        import archive_db
                        archive_db.save_comparison(
                            result, file1=f1, file2=f2,
                            f1_label=name1, f2_label=name2,
                            elapsed_s=elapsed, source="gui",
                        )
                    except Exception as ex:
                        archive_err = str(ex)
                        app_log.log_run_err(ex, traceback.format_exc())
                self.after(
                    0,
                    lambda r=result, n=out_name, p=out_path, ae=archive_err, lg=lang:
                        self._on_compare_ok(r, n, p, archive_err=ae, lang=lg),
                )
            except Exception as ex:
                tb = traceback.format_exc()
                app_log.log_run_err(ex, tb)
                self.after(0, lambda e=ex, t=tb: self._on_compare_err(e, t))

        threading.Thread(target=worker, daemon=True).start()

    def run_batch(self):
        if self._busy:
            return
        a = filedialog.askdirectory(title=self._t("batch_pick_a"))
        if not a:
            return
        b = filedialog.askdirectory(title=self._t("batch_pick_b"))
        if not b:
            return
        out_dir = os.path.join(DESKTOP, f"batch_out_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
        lang = self.lang.get()
        write_html = bool(self.var_html.get())
        write_pdf = bool(self.var_pdf.get())
        detail_sheets = not bool(self.var_fast.get())
        import archive_db as arch_mod
        archive_db_path = arch_mod.DEFAULT_DB if self.var_archive.get() else None

        self._busy = True
        self.compare_btn.config(state="disabled")
        self.progress["value"] = 0

        def progress_cb(pct, stage):
            self.after(0, lambda: self._set_progress_ui(pct, f"{pct}% {stage}"))

        def worker():
            try:
                import batch_compare
                summary = batch_compare.run_batch(
                    a, b, out_dir,
                    progress_cb=progress_cb,
                    write_html=write_html,
                    write_pdf=write_pdf,
                    archive_db=archive_db_path,
                    detail_sheets=detail_sheets,
                    lang=lang,
                )
                self.after(
                    0,
                    lambda s=summary, p=out_dir, lg=lang: self._on_batch_ok(s, p, lg),
                )
            except Exception as ex:
                tb = traceback.format_exc()
                app_log.log_run_err(ex, tb)
                self.after(0, lambda e=ex, t=tb: self._on_compare_err(e, t))

        threading.Thread(target=worker, daemon=True).start()

    def show_archive(self):
        try:
            import archive_db
            rows = archive_db.list_comparisons(limit=30)
        except Exception as ex:
            messagebox.showerror(self._t("archive_title"), str(ex))
            return
        if not rows:
            messagebox.showinfo(self._t("archive_title"), self._t("archive_empty"))
            return
        lines = []
        for r in rows:
            lines.append(
                f"#{r['id']} {r['created_at']}  "
                f"m={r['match_count']} mm={r['mismatch_count']} "
                f"o1={r['only1_count']} o2={r['only2_count']}\n"
                f"  {os.path.basename(r['file1'] or '')} ↔ {os.path.basename(r['file2'] or '')}"
            )
        messagebox.showinfo(self._t("archive_title"), "\n".join(lines[:40]))

    def _on_batch_ok(self, summary, out_dir, lang):
        self._busy = False
        self.compare_btn.config(state="normal")
        self.progress["value"] = 100
        msg = i18n.t("batch_done", lang=lang, n=summary.get("pair_count", 0), path=out_dir)
        self.status.config(text=msg, fg="#0a6")
        messagebox.showinfo(self._t("batch_title"), msg)

    def _on_compare_ok(self, result, out_name, out_path, archive_err=None, lang=None):
        lang = lang or self.lang.get()
        self._busy = False
        self.compare_btn.config(state="normal")
        self.progress["value"] = 100
        self.status.config(text=i18n.t("done", lang=lang, name=out_name), fg="#0a6")
        warns = list(result.get("warnings") or [])
        if archive_err:
            warns.append(archive_err)
            messagebox.showwarning(
                i18n.t("warn_open_title", lang=lang),
                i18n.t("archive_warn", lang=lang, err=archive_err),
            )
        if result.get("warnings"):
            messagebox.showwarning(
                i18n.t("warn_open_title", lang=lang),
                i18n.t("report_warn", lang=lang, err="\n".join(result["warnings"])),
            )
        try:
            open_file(out_path)
        except OSError:
            messagebox.showwarning(
                i18n.t("warn_open_title", lang=lang),
                i18n.t("warn_open_body", lang=lang, path=out_path),
            )

        top_diffs = result["diffs"][:5]
        msg_lines = [
            i18n.t("summary_sheet", lang=lang),
            i18n.t("summary_yellow", lang=lang, n=result["match_count"]),
            i18n.t("summary_orange", lang=lang, n=result["mismatch_count"]),
            i18n.t("summary_red1", lang=lang, n=result["only1_count"]),
            i18n.t("summary_red2", lang=lang, n=result["only2_count"]),
            "",
        ]
        if top_diffs:
            msg_lines.append(i18n.t("summary_diffs", lang=lang))
            for k, b1, b2, d, c1, c2 in top_diffs:
                msg_lines.append(i18n.t("summary_diff_line", lang=lang, k=k, d=d))
        else:
            msg_lines.append(i18n.t("summary_no_diffs", lang=lang))
        messagebox.showinfo(i18n.t("done_title", lang=lang), "\n".join(msg_lines))

    def _on_compare_err(self, ex, tb):
        self._busy = False
        self.compare_btn.config(state="normal")
        self.progress["value"] = 0
        self.status.config(text=self._t("error_status"), fg="#c00")
        if isinstance(ex, PermissionError):
            messagebox.showerror(self._t("format_error_title"), self._t("err_permission"))
        elif isinstance(ex, FileNotFoundError):
            messagebox.showerror(self._t("format_error_title"), self._t("err_not_found", err=ex))
        elif isinstance(ex, motor.FormatError):
            messagebox.showerror(self._t("format_error_title"), str(ex))
        else:
            messagebox.showerror(
                self._t("format_error_title"),
                self._t("err_unexpected", err=f"{ex}\n\n{tb}"),
            )


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


def main():
    argv = sys.argv[1:]
    if argv and _wants_cli(argv):
        raise SystemExit(run_cli(argv))
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

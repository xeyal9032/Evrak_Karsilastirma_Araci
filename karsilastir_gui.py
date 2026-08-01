# -*- coding: utf-8 -*-
"""Tkinter GUI for Evrak Karsilastirma Araci (imported only for GUI mode)."""
import datetime
import os
import sys
import threading
import time
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import karsilastir_motor as motor
import i18n
import app_log
from karsilastir import DESKTOP, FILETYPES, open_file


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
        warns = list(summary.get("warnings") or [])
        if warns:
            messagebox.showwarning(
                i18n.t("warn_open_title", lang=lang),
                i18n.t("report_warn", lang=lang, err="\n".join(warns)),
            )
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


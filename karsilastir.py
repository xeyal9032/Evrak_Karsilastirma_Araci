# -*- coding: utf-8 -*-
"""
Evrak Karsilastirma Araci - basit pencere (TR / RU).
Инструмент сравнения документов - простое окно (TR / RU).

Iki muhasebe evraki (DATEV CSV/Excel veya Journal Excel) secip
"Karsilastir / Сравнить" butonuna basiyorsunuz; format otomatik taninir,
renkli bir Excel raporu masaustune kaydedilip otomatik aciliyor.
"""
import os
import sys
import traceback
import datetime
import threading
import platform
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import karsilastir_motor as motor


FILETYPES = [
    ("Evrak dosyalari / Документы", "*.csv *.xlsx *.xlsm"),
    ("CSV", "*.csv"),
    ("Excel", "*.xlsx *.xlsm"),
    ("Tum dosyalar / Все файлы", "*.*"),
]


def desktop_dir():
    home = os.path.expanduser("~")
    for name in ("Desktop", "Masaüstü", "Рабочий стол"):
        path = os.path.join(home, name)
        if os.path.isdir(path):
            return path
    return home


def open_file(path):
    """Raporu varsayilan uygulama ile acar (Windows / macOS / Linux)."""
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
        self.title("Evrak Karsilastirma Araci / Инструмент сравнения документов")
        self.geometry("700x320")
        self.resizable(False, False)
        self._busy = False
        self._fmt1 = None
        self._fmt2 = None

        self.file1 = tk.StringVar()
        self.file2 = tk.StringVar()

        tk.Label(self, text="Evrak Karsilastirma Araci / Инструмент сравнения документов",
                 font=("Segoe UI", 13, "bold"), wraplength=660, justify="center").pack(pady=(16, 4))
        tk.Label(self,
                 text="DATEV CSV/Excel veya Journal Excel secin - format otomatik taninir.\n"
                      "Выберите DATEV CSV/Excel или Journal Excel - формат определяется автоматически.",
                 fg="#555", wraplength=660, justify="center").pack(pady=(0, 16))

        f1_frame = tk.Frame(self)
        f1_frame.pack(fill="x", padx=20, pady=6)
        tk.Button(f1_frame, text="Dosya 1 Sec / Файл 1...", width=22, command=self.pick_file1).pack(side="left")
        tk.Label(f1_frame, textvariable=self.file1, anchor="w", fg="#333").pack(side="left", padx=10, fill="x", expand=True)

        f2_frame = tk.Frame(self)
        f2_frame.pack(fill="x", padx=20, pady=6)
        tk.Button(f2_frame, text="Dosya 2 Sec / Файл 2...", width=22, command=self.pick_file2).pack(side="left")
        tk.Label(f2_frame, textvariable=self.file2, anchor="w", fg="#333").pack(side="left", padx=10, fill="x", expand=True)

        self.status = tk.Label(self, text="", fg="#0a6")
        self.status.pack(pady=8)

        self.compare_btn = tk.Button(self, text="Karsilastir / Сравнить", font=("Segoe UI", 11, "bold"),
                                      bg="#4472C4", fg="white", width=26, height=2, command=self.run_compare)
        self.compare_btn.pack(pady=10)

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
            title=f"Dosya {which} sec / Выбрать файл {which}",
            filetypes=FILETYPES,
        )
        if not p:
            return
        try:
            fmt = motor.detect_format(p)
        except (motor.FormatError, FileNotFoundError, OSError) as ex:
            messagebox.showerror("Format hatasi / Ошибка формата", str(ex))
            return
        if which == 1:
            self.file1.set(p)
            self._fmt1 = fmt
        else:
            self.file2.set(p)
            self._fmt2 = fmt
        self._update_format_status()

    def pick_file1(self):
        self._pick(1)

    def pick_file2(self):
        self._pick(2)

    def run_compare(self):
        if self._busy:
            return
        f1 = self.file1.get()
        f2 = self.file2.get()
        if not f1 or not f2:
            messagebox.showwarning(
                "Eksik dosya / Не хватает файла",
                "Lutfen iki dosyayi da secin.\nПожалуйста, выберите оба файла.",
            )
            return

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"Karsilastirma_{ts}.xlsx"
        out_path = os.path.join(DESKTOP, out_name)
        name1 = os.path.splitext(os.path.basename(f1))[0][:31]
        name2 = os.path.splitext(os.path.basename(f2))[0][:31]

        self._busy = True
        fmt_txt = ""
        if self._fmt1 and self._fmt2:
            fmt_txt = f" ({motor.format_label(self._fmt1)} ↔ {motor.format_label(self._fmt2)})"
        self.status.config(
            text=f"Karsilastiriliyor{fmt_txt}... / Сравнение, подождите...",
            fg="#c60",
        )
        self.compare_btn.config(state="disabled")

        def worker():
            try:
                result = motor.build_report(f1, f2, out_path, f1_label=name1, f2_label=name2)
                self.after(0, lambda: self._on_compare_ok(result, out_name, out_path))
            except Exception as ex:
                tb = traceback.format_exc()
                self.after(0, lambda: self._on_compare_err(ex, tb))

        threading.Thread(target=worker, daemon=True).start()

    def _on_compare_ok(self, result, out_name, out_path):
        self._busy = False
        self.compare_btn.config(state="normal")
        self.status.config(text=f"Tamamlandi / Готово. Rapor / Отчёт: {out_name}", fg="#0a6")
        try:
            open_file(out_path)
        except OSError:
            messagebox.showwarning(
                "Uyari / Предупреждение",
                f"Rapor kaydedildi ama otomatik acilamadi. Elle acabilirsiniz:\n{out_path}\n\n"
                f"Отчёт сохранён, но не открылся автоматически. Откройте вручную:\n{out_path}",
            )

        top_diffs = result["diffs"][:5]
        msg_lines = [
            "'KARSILASTIRMA-SRAVNENIE' sayfasi / Лист 'KARSILASTIRMA-SRAVNENIE':",
            f"  Sari / Жёлтый (ayni/совпадает): {result['match_count']}",
            f"  Turuncu / Оранжевый (deger farkli/значение отличается): {result['mismatch_count']}",
            f"  Kirmizi / Красный (sadece Dosya1/только в файле 1): {result['only1_count']}",
            f"  Kirmizi / Красный (sadece Dosya2/только в файле 2): {result['only2_count']}",
            "",
        ]
        if top_diffs:
            msg_lines.append("En buyuk hesap bazli farklar / Крупнейшие разницы по счетам:")
            for k, b1, b2, d, c1, c2 in top_diffs:
                msg_lines.append(f"  Konto / Счёт {k}: fark / разница {d:,.2f} EUR")
        else:
            msg_lines.append("Hesap bazli onemli bir fark bulunamadi. / Существенных разниц по счетам не найдено.")
        messagebox.showinfo("Karsilastirma tamamlandi / Сравнение завершено", "\n".join(msg_lines))

    def _on_compare_err(self, ex, tb):
        self._busy = False
        self.compare_btn.config(state="normal")
        self.status.config(text="Hata olustu. / Произошла ошибка.", fg="#c00")
        if isinstance(ex, PermissionError):
            messagebox.showerror(
                "Hata / Ошибка",
                "Rapor dosyasi kaydedilemedi - hedef dosya baska bir programda (orn. Excel'de) acik olabilir.\n"
                "Не удалось сохранить файл отчёта - возможно, он открыт в другой программе (напр. в Excel).",
            )
        elif isinstance(ex, FileNotFoundError):
            messagebox.showerror(
                "Dosya bulunamadi / Файл не найден",
                f"Secilen dosya bulunamadi ya da erisilemedi.\nВыбранный файл не найден или недоступен.\n\n{ex}",
            )
        elif isinstance(ex, motor.FormatError):
            messagebox.showerror("Format hatasi / Ошибка формата", str(ex))
        else:
            messagebox.showerror(
                "Beklenmeyen hata / Непредвиденная ошибка",
                f"Bir hata olustu / Произошла ошибка:\n\n{ex}\n\n{tb}",
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()

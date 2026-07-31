# -*- coding: utf-8 -*-
"""Journal ve DATEV Excel fixture uretici (bir kez calistirilir / test setup)."""
import os
from datetime import datetime

import openpyxl

FIXTURES = os.path.dirname(os.path.abspath(__file__))


def build_journal_fixture():
    path = os.path.join(FIXTURES, "sample_journal.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tabelle1"
    ws.append(["Journal Test"])
    ws.append(["Belegdat.", "Zusatzang.", "Periode", "Belegnr.", "Buchungstext",
               "Betrag", "Whrg", "Sollkto", "Habenkto", "USt-EUR", "USt Kto"])
    # Matches sample_a RE-1001 (DATEV S/1200/8400)
    ws.append([datetime(2026, 1, 15), None, 1, "RE-1001", "Miete Januar",
               1000.00, "EUR", 1200, 8400, None, None])
    # DATEV sample_a RE-2002 is H/8400/1200 (= Journal Soll 1200 / Haben 8400), amount differs
    ws.append([datetime(2026, 1, 20), None, 1, "RE-2002", "Wareneinkauf",
               999.99, "EUR", 1200, 8400, None, None])
    # Only in journal
    ws.append([datetime(2026, 2, 11), None, 2, "RE-ONLY2", "Nur in Journal",
               750.00, "EUR", 1700, 1800, None, None])
    wb.save(path)
    wb.close()
    return path


def build_datev_xlsx_fixture():
    path = os.path.join(FIXTURES, "sample_datev.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export"
    ws.append(["EXTF", "700", "21", "Buchungsstapel"])
    ws.append([
        "Umsatz (ohne Soll/Haben-Kz)", "Soll/Haben-Kennzeichen", "WKZ Umsatz", "Kurs",
        "Basis-Umsatz", "WKZ Basis-Umsatz", "Konto", "Gegenkonto (ohne BU-Schluessel)",
        "BU-Schlüssel", "Belegdatum", "Belegfeld 1", "Belegfeld 2", "Skonto", "Buchungstext",
    ])
    ws.append(["1000,00", "S", "EUR", "", "", "", "1200", "8400", "", "1501", "RE-1001", "", "", "Miete Januar"])
    ws.append(["250,50", "H", "EUR", "", "", "", "8400", "1200", "", "2001", "RE-2002", "", "", "Wareneinkauf"])
    wb.save(path)
    wb.close()
    return path


if __name__ == "__main__":
    print(build_journal_fixture())
    print(build_datev_xlsx_fixture())

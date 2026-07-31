# -*- coding: utf-8 -*-
"""Test yardimcilari: DATEV/Journal dosya uretimi."""
import csv
import os
from datetime import datetime

import openpyxl

DATEV_HEADER = [
    "Umsatz", "Soll/Haben-Kennzeichen", "WKZ Umsatz", "Kurs", "Basisumsatz", "WKZ Basisumsatz",
    "Konto", "Gegenkonto", "BU-Schluessel", "Belegdatum", "Belegfeld 1", "Belegfeld 2",
    "Skonto", "Buchungstext",
]

DATEV_HEADER_LONG = [
    "Umsatz (ohne Soll/Haben-Kz)", "Soll/Haben-Kennzeichen", "WKZ Umsatz", "Kurs",
    "Basis-Umsatz", "WKZ Basis-Umsatz", "Konto", "Gegenkonto (ohne BU-Schluessel)",
    "BU-Schlüssel", "Belegdatum", "Belegfeld 1", "Belegfeld 2", "Skonto", "Buchungstext",
]


def drow(umsatz, sh, konto, gegen, datum, beleg, text):
    return [umsatz, sh, "", "", "", "", konto, gegen, "", datum, beleg, "", "", text]


def write_datev_csv(path, data_rows, header=None, encoding="latin-1"):
    header = header or DATEV_HEADER
    with open(path, "w", encoding=encoding, newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["EXTF", "700", "21", "Buchungsstapel", "test"])
        w.writerow(header)
        for row in data_rows:
            w.writerow(row)


def write_datev_xlsx(path, data_rows, header=None):
    header = header or DATEV_HEADER_LONG
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["EXTF", "700", "21", "Buchungsstapel"])
    ws.append(header)
    for row in data_rows:
        ws.append(row)
    wb.save(path)
    wb.close()


def write_journal_xlsx(path, data_rows):
    """data_rows: list of (datetime|None, belegnr, text, betrag, soll, haben)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Journal Test"])
    ws.append([
        "Belegdat.", "Zusatzang.", "Periode", "Belegnr.", "Buchungstext",
        "Betrag", "Whrg", "Sollkto", "Habenkto", "USt-EUR", "USt Kto",
    ])
    for belegdat, belegnr, text, betrag, soll, haben in data_rows:
        ws.append([belegdat, None, 1, belegnr, text, betrag, "EUR", soll, haben, None, None])
    wb.save(path)
    wb.close()

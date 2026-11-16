"""Buat file contoh RAB 1:1 (data dummy)."""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "sample"

# Item: (Nama, Satuan, Harga Satuan, Qty Total)
ITEMS = [
    ("Service", "UNIT", 240_000, 20),
    ("Oli Mesin", "LITER", 100_000, 20),
    ("Kampas Rem Depan", "SET", 100_000, 8),
    ("Kampas Rem Belakang", "SET", 100_000, 8),
    ("Kampas Kopling", "SET", 330_000, 4),
    ("Bohlam Depan", "PCS", 65_000, 6),
    ("Bohlam Belakang", "PCS", 20_000, 6),
    ("Busi", "PCS", 30_000, 10),
    ("Filter Udara", "PCS", 70_000, 6),
    ("Lager Roda Belakang", "PCS", 113_000, 4),
    ("Rantai & Gir", "SET", 250_000, 3),
    ("Air Radiator", "LITER", 35_000, 5),
]

THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="C0392B")
HEAD_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")


def build_rab() -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RAB"
    ws["A1"] = "RAB PENGADAAN ALAT TULIS - CONTOH"
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = "Proyek: Pengadaan Alat Tulis Kantor 2026"
    ws["A3"] = "Format: No | Nama Barang | Satuan | Harga Satuan | Qty Total | Termin 1..4"

    headers = [
        "No",
        "Nama Barang",
        "Satuan",
        "Harga Satuan",
        "Qty Total",
        "Termin 1",
        "Termin 2",
        "Termin 3",
        "Termin 4",
    ]
    header_row = 5
    for index, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=index, value=header)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for index, (name, unit, price, qty) in enumerate(ITEMS, start=1):
        row = header_row + index
        values = [index, name, unit, price, qty, 0, 0, 0, 0]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.border = BORDER
            if col == 4:
                cell.number_format = '#,##0'
            if col >= 5:
                cell.alignment = Alignment(horizontal="right")

    total_row = header_row + len(ITEMS) + 1
    ws.cell(row=total_row, column=2, value="TOTAL").font = Font(bold=True)
    ws.cell(row=total_row, column=5, value=sum(q for _n, _u, _p, q in ITEMS))
    widths = {"A": 5, "B": 28, "C": 10, "D": 15, "E": 10, "F": 10, "G": 10, "H": 10, "I": 10}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A6"

    OUT.mkdir(exist_ok=True)
    path = OUT / "rab_sample.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    print(build_rab())
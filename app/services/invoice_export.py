"""Export invoice 1:1 berdasarkan template invoice.

Hanya mode 'files' (satu file Excel per invoice). Rekap/pivot/workbook dihapus.
"""
from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Sequence

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.properties import PageSetupProperties

from ..utils import sanitize_filename, unique_path
from .template_profile import DEFAULT_PROFILE, TemplateProfile, detect_header_row
from .template_renderer import DesignConfig, render_invoice_html


class InvoiceExportError(Exception):
    pass


@dataclass
class InvoiceLine:
    name: str
    qty: int
    unit_price: int

    @property
    def amount(self) -> int:
        return int(self.qty) * int(self.unit_price)


@dataclass
class InvoiceData:
    label: str
    client_name: str
    client_address: str
    company_name: str
    company_address: str
    po_number: str
    invoice_date: date
    invoice_date_text: str = ""
    lines: list[InvoiceLine] = field(default_factory=list)
    project_name: str = ""

    @property
    def subtotal(self) -> int:
        return sum(line.amount for line in self.lines)

    @property
    def total(self) -> int:
        return self.subtotal


def _copy_style(src, dst) -> None:
    """Salin style antar-workbook dengan aman."""
    dst.value = src.value
    if not src.has_style:
        return
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)
    dst.number_format = src.number_format


def _clone_block(ws, tpl_ws, rows: range, cols: range) -> None:
    for row in rows:
        for col in cols:
            src = tpl_ws.cell(row=row, column=col)
            dst = ws.cell(row=row, column=col)
            _copy_style(src, dst)


def _clone_row_style(ws, tpl_ws, src_row: int, dst_row: int, cols: range) -> None:
    for col in cols:
        src = tpl_ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        _copy_style(src, dst)
        dst.value = None


def _set_style_only(ws, tpl_ws, coord: str, dst_coord: str) -> None:
    src = tpl_ws[coord]
    dst = ws[dst_coord]
    _copy_style(src, dst)


def _unique_sheet_name(wb: Workbook, desired: str) -> str:
    cleaned = "".join(
        ch for ch in str(desired or "Invoice") if ch not in '[]:*?/\\'
    ).strip()
    cleaned = cleaned[:31] or "Invoice"
    name = cleaned
    index = 2
    while name in wb.sheetnames:
        suffix = f" {index}"
        name = cleaned[: 31 - len(suffix)] + suffix
        index += 1
    return name


def build_invoice_workbook(
    template_path: str | Path,
    invoice: InvoiceData,
    workbook: Workbook | None = None,
    sheet_name: str | None = None,
    profile: TemplateProfile | None = None,
) -> Workbook:
    """Bangun workbook invoice dengan layout & style dari template."""
    template_path = Path(template_path)
    if not template_path.exists():
        raise InvoiceExportError(
            f"Template invoice tidak ditemukan: {template_path}"
        )
    try:
        tpl_wb = openpyxl.load_workbook(str(template_path), data_only=False)
    except Exception as exc:  # noqa: BLE001
        raise InvoiceExportError(f"Template tidak bisa dibuka: {exc}") from exc

    profile = profile or DEFAULT_PROFILE
    index = min(profile.sheet_index, len(tpl_wb.worksheets) - 1)
    tpl_ws = tpl_wb.worksheets[index]
    header_row = detect_header_row(tpl_ws)
    if header_row is None:
        raise InvoiceExportError(
            "Template invoice tidak sesuai: header 'Qty' dan 'Nama Barang' "
            f"tidak ditemukan di {template_path.name}."
        )

    if workbook is None:
        wb = Workbook()
        default = wb.active
        wb.remove(default)
    else:
        wb = workbook

    title = _unique_sheet_name(
        wb,
        sheet_name
        or invoice.po_number
        or invoice.label
        or "Invoice",
    )
    ws = wb.create_sheet(title=title)

    # Lebar kolom & tinggi baris
    for key, dim in tpl_ws.column_dimensions.items():
        if dim.width:
            ws.column_dimensions[key].width = dim.width
    for row in range(1, header_row + 1):
        height = tpl_ws.row_dimensions[row].height
        if height:
            ws.row_dimensions[row].height = height

    # Salin area header (sampai baris header tabel item) termasuk merge
    _clone_block(ws, tpl_ws, range(1, header_row + 1), range(1, 6))
    for merged in tpl_ws.merged_cells.ranges:
        if merged.min_row <= header_row and merged.max_row <= header_row:
            ws.merge_cells(str(merged))

    # Isi header: peta sel -> kunci nilai berasal dari profil template
    invoice_date = invoice.invoice_date
    if isinstance(invoice_date, date) and not isinstance(invoice_date, datetime):
        invoice_date = datetime(invoice_date.year, invoice_date.month, invoice_date.day)

    values = {
        "title": "INVOICE",
        "company_name": invoice.company_name or "",
        "company_address": invoice.company_address or "",
        "po_number": invoice.po_number or "",
        "invoice_date": invoice.invoice_date_text or invoice_date,
        "client_name": invoice.client_name or "",
        "client_address": invoice.client_address or "",
        "label": invoice.label or "",
        "entity_type": invoice.project_name or "",
    }
    for coord, key in profile.fields.items():
        if key in values:
            ws[coord] = values[key]

    # Tabel item
    lines = invoice.lines or []
    first_row = header_row + 1
    columns = profile.columns
    style_item = profile.style_rows.get("item", first_row)
    for index, line in enumerate(lines):
        row = first_row + index
        _clone_row_style(ws, tpl_ws, style_item, row, range(1, 5))
        ws.cell(row=row, column=int(columns.get("qty", 1)), value=int(line.qty))
        ws.cell(row=row, column=int(columns.get("name", 2)), value=line.name)
        ws.cell(row=row, column=int(columns.get("unit_price", 3)), value=int(line.unit_price))
        ws.cell(row=row, column=int(columns.get("amount", 4)), value=line.amount)
        height = tpl_ws.row_dimensions[style_item].height
        ws.row_dimensions[row].height = height or 15

    subtotal_row = first_row + len(lines)
    total_row = subtotal_row + 1
    style_subtotal = profile.style_rows.get("subtotal", 21)
    style_total = profile.style_rows.get("total", 23)

    # Sub total
    _clone_row_style(ws, tpl_ws, style_subtotal, subtotal_row, range(1, 5))
    _set_style_only(ws, tpl_ws, f"A{style_subtotal}", f"A{subtotal_row}")
    ws[f"A{subtotal_row}"] = profile.tax_note
    ws[f"C{subtotal_row}"] = "Sub Total"
    ws[f"D{subtotal_row}"] = invoice.subtotal

    # Total
    _clone_row_style(ws, tpl_ws, style_total, total_row, range(1, 5))
    ws[f"C{total_row}"] = "TOTAL"
    ws[f"D{total_row}"] = invoice.total

    # Footer tanda tangan
    signature_row = total_row + int(profile.footer_offsets.get("signature", 3))
    company_row = total_row + int(profile.footer_offsets.get("company", 6))
    for src_row, dst_row, text in (
        (profile.style_rows.get("signature", 26), signature_row, profile.signature_text),
        (profile.style_rows.get("company", 29), company_row, invoice.company_name or ""),
    ):
        _clone_row_style(ws, tpl_ws, src_row, dst_row, range(3, 5))
        ws.merge_cells(
            start_row=dst_row, start_column=3, end_row=dst_row, end_column=4
        )
        ws.cell(row=dst_row, column=3, value=text)

    # Print setup
    try:
        ws.page_margins = copy(tpl_ws.page_margins)
    except Exception:
        pass
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = f"A1:D{company_row}"

    tpl_wb.close()
    return wb


def invoice_filename(invoice: InvoiceData) -> str:
    label = sanitize_filename(invoice.label or "Invoice")
    return f"Invoice {label}.xlsx"


def export_invoices(
    template_path: str | Path,
    invoices: Sequence[InvoiceData],
    output_dir: str | Path,
    *,
    mode: str = "files",
    overwrite: bool = False,
    profile: TemplateProfile | None = None,
    design_config: DesignConfig | None = None,
    export_html: bool = False,
) -> list[Path]:
    """Export invoice — mode 'files' (file Excel per invoice, dan file HTML A4 presisi)."""
    if not invoices:
        raise InvoiceExportError("Tidak ada invoice yang akan diekspor.")
    if mode != "files":
        raise InvoiceExportError("Mode 'workbook' tidak didukung. Gunakan mode 'files'.")
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)

    paths: list[Path] = []
    for invoice in invoices:
        wb = build_invoice_workbook(template_path, invoice, profile=profile)
        path = output / invoice_filename(invoice)
        if not overwrite:
            path = unique_path(path)
        try:
            wb.save(path)
        except Exception as exc:  # noqa: BLE001
            raise InvoiceExportError(
                f"Gagal menyimpan invoice {invoice.label}: {exc}"
            ) from exc
        paths.append(path)

        # Export dokumen HTML A4 berdesain tinggi (mengadopsi reference-invoicely.html)
        if export_html:
            html_cfg = design_config or DesignConfig()
            html_content = render_invoice_html(invoice, html_cfg)
            html_name = sanitize_filename(invoice.label or "Invoice") + ".html"
            html_path = output / html_name
            if not overwrite:
                html_path = unique_path(html_path)
            try:
                html_path.write_text(html_content, encoding="utf-8")
                paths.append(html_path)
            except Exception:
                pass

    return paths
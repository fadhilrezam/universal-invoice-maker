"""Pembangun data invoice 1:1 (netto)."""
from __future__ import annotations

from datetime import date
from typing import Sequence

from ..models import Item, Project
from ..utils import format_date_id, format_rupiah
from .invoice_export import InvoiceData, InvoiceLine


def build_invoice_netto(
    project: Project,
    items: Sequence[Item],
    *,
    company_name: str = "",
    company_address: str = "",
    label: str = "",
    po_number: str = "",
    invoice_date: date | None = None,
    date_format: str = "",
) -> InvoiceData | None:
    """Invoice 1:1 tanpa termin dan tanpa pengelompokan.

    Satu baris RAB menjadi tepat satu baris invoice: qty diambil apa adanya
    dari ``item.total_qty`` (kolom volume pada blok "sesuai pelaksanaan"),
    harga dari ``item.unit_price``. Tidak ada agregasi, pemecahan qty, atau
    perhitungan PPN — nilai invoice adalah nilai akhir (netto).
    """
    lines = [
        InvoiceLine(
            name=item.name.strip(),
            qty=int(item.total_qty),
            unit_price=int(item.unit_price),
        )
        for item in items
        if (item.name or "").strip()
    ]
    if not lines:
        return None

    label = (label or "").strip() or project.name
    date_value = invoice_date or project.date_obj
    fmt = date_format or project.date_format or "full"
    po = (po_number or "").strip() or f"{project.po_prefix}{label}"
    return InvoiceData(
        label=label,
        client_name=project.client_name,
        client_address=project.client_address,
        company_name=company_name,
        company_address=company_address,
        po_number=po,
        invoice_date=date_value,
        invoice_date_text=format_date_id(date_value, fmt),
        lines=lines,
        project_name=project.name,
    )


def invoice_summary_text(invoice: InvoiceData) -> str:
    return f"{invoice.label}: {len(invoice.lines)} item, total {format_rupiah(invoice.total)}"

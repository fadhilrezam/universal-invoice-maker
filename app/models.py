"""Model data aplikasi Universal Invoice Maker.

Tidak ada kendaraan, termin, atau alokasi: satu proyek = satu RAB, dan RAB
dipetakan satu banding satu menjadi invoice.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .utils import parse_iso_date


@dataclass
class Project:
    id: int
    name: str
    client_name: str = ""
    client_address: str = ""
    po_prefix: str = ""
    invoice_date: str = ""
    date_format: str = "full"
    created_at: str = ""

    @property
    def date_obj(self) -> date:
        return parse_iso_date(self.invoice_date)

    @classmethod
    def from_row(cls, row) -> "Project":
        keys = row.keys()
        return cls(
            id=row["id"],
            name=row["name"],
            client_name=row["client_name"] or "",
            client_address=row["client_address"] or "",
            po_prefix=row["po_prefix"] or "",
            invoice_date=row["invoice_date"] or "",
            date_format=(
                (row["date_format"] or "full")
                if "date_format" in keys
                else "full"
            ),
            created_at=row["created_at"] or "",
        )


@dataclass
class Item:
    id: int
    project_id: int
    row_no: int
    name: str
    unit: str = ""
    unit_price: int = 0
    total_qty: int = 0

    @classmethod
    def from_row(cls, row) -> "Item":
        return cls(
            id=row["id"],
            project_id=row["project_id"],
            row_no=row["row_no"] or 0,
            name=row["name"] or "",
            unit=row["unit"] or "",
            unit_price=int(round(row["unit_price"] or 0)),
            total_qty=int(round(row["total_qty"] or 0)),
        )


@dataclass
class ParsedItem:
    """Item hasil parsing RAB sebelum disimpan."""

    name: str
    unit: str
    unit_price: int
    total_qty: int
    source_row: int = 0


@dataclass
class ValidationIssue:
    """Masalah validasi proyek."""

    item_name: str
    message: str
    row_no: int = 0
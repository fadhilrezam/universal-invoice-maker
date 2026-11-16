"""Validasi data proyek sebelum invoice diekspor."""
from __future__ import annotations

from ..db import Database
from ..models import Item, ValidationIssue


def _check_item(item: Item) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    name = item.name.strip()
    if not name:
        issues.append(ValidationIssue("(tanpa nama)", "Nama barang kosong.", item.row_no))
        return issues
    if item.total_qty <= 0:
        issues.append(
            ValidationIssue(name, f"Qty harus lebih dari 0 (sekarang {item.total_qty}).", item.row_no)
        )
    if item.unit_price < 0:
        issues.append(
            ValidationIssue(
                name,
                f"Harga satuan negatif ({item.unit_price}).",
                item.row_no,
            )
        )
    return issues


def validate_project(db: Database, project_id: int) -> list[ValidationIssue]:
    """Validasi ringan: proyek punya item dan setiap item layak ditagihkan."""
    items = db.list_items(project_id)
    if not items:
        return [ValidationIssue("", "RAB belum diimpor atau tidak berisi item.")]
    issues: list[ValidationIssue] = []
    for item in items:
        issues.extend(_check_item(item))
    return issues

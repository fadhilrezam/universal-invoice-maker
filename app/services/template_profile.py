"""Profil layout template invoice.

Layout template (baris header, sel header, anchor style, kolom tabel) disimpan
sebagai data, bukan ditulis sebagai literal di dalam kode export. Dengan begitu
template lain bisa dipakai tanpa mengubah ``invoice_export.py``.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

# Teks header tabel item yang dipakai untuk menemukan baris header.
HEADER_LABELS = ("qty", "nama barang")

# Peta sel header -> kunci nilai invoice. Kunci yang tidak ada di data invoice
# dibiarkan apa adanya (nilai bawaan template tidak ditimpa).
DEFAULT_FIELDS = {
    "A1": "title",
    "C1": "company_name",
    "C2": "company_address",
    "B4": "po_number",
    "B5": "invoice_date",
    "D4": "client_name",
    "D5": "client_address",
    "D6": "label",
    "D7": "entity_type",
}

# Kolom tabel item: nama peran -> nomor kolom.
DEFAULT_COLUMNS = {"qty": 1, "name": 2, "unit_price": 3, "amount": 4}

# Baris sumber untuk menyalin style (bukan baris tulis).
DEFAULT_STYLE_ROWS = {
    "item": 11,
    "subtotal": 21,
    "discount": 22,
    "total": 23,
    "signature": 26,
    "company": 29,
}


@dataclass
class TemplateProfile:
    """Deskripsi satu layout template invoice."""

    name: str = "Default Invoice"
    sheet_index: int = 0
    header_row: int = 10
    item_first_row: int = 11
    fields: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_FIELDS))
    columns: dict[str, int] = field(default_factory=lambda: dict(DEFAULT_COLUMNS))
    style_rows: dict[str, int] = field(default_factory=lambda: dict(DEFAULT_STYLE_ROWS))
    # Jarak baris footer dari baris total.
    footer_offsets: dict[str, int] = field(
        default_factory=lambda: {"signature": 3, "company": 6}
    )
    signature_text: str = "Hormat Kami,"
    tax_note: str = "* harga sudah termasuk pajak"
    print_area_columns: int = 4
    # Konfigurasi styling & layout A4 (diadopsi dari reference-invoicely.html)
    design_config: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "TemplateProfile":
        known = {f for f in cls.__dataclass_fields__}  # type: ignore[attr-defined]
        clean = {k: v for k, v in (data or {}).items() if k in known}
        return cls(**clean)

    def save(self, path: str | Path) -> None:
        Path(path).write_text(
            json.dumps(self.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8"
        )

    @classmethod
    def load(cls, path: str | Path) -> "TemplateProfile":
        return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))

    def fingerprint(self, header_texts: list[str]) -> str:
        """Sidik jari layout: profil dicocokkan dengan bentuk header file."""
        raw = "|".join(
            [
                self.name,
                str(self.header_row),
                str(len(header_texts)),
                *[t.strip().lower() for t in header_texts],
            ]
        )
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


DEFAULT_PROFILE = TemplateProfile()


def header_texts(ws, row: int) -> list[str]:
    """Ambil teks header pada satu baris (sel kosong jadi string kosong)."""
    return [str(cell.value or "").strip() for cell in ws[row]]


def detect_header_row(ws, max_scan: int = 60) -> int | None:
    """Cari baris header tabel item.

    Baris dianggap header bila memuat semua label di ``HEADER_LABELS``.
    Tidak lagi mengunci ke baris 10 seperti sebelumnya.
    """
    limit = min(ws.max_row or 0, max_scan)
    for row in range(1, limit + 1):
        texts = [t.lower() for t in header_texts(ws, row)]
        if all(any(label in text for text in texts) for label in HEADER_LABELS):
            return row
    return None


def profile_for_template(path: str | Path) -> TemplateProfile:
    """Bangun profil dari file template; jatuh ke default bila header tak ketemu."""
    import openpyxl

    profile = TemplateProfile()
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=False)
    except Exception:  # noqa: BLE001
        return profile
    try:
        if not wb.worksheets:
            return profile
        ws = wb.worksheets[min(profile.sheet_index, len(wb.worksheets) - 1)]
        row = detect_header_row(ws)
        if row:
            profile.header_row = row
            profile.item_first_row = row + 1
            profile.style_rows["item"] = row + 1
    finally:
        wb.close()
    return profile

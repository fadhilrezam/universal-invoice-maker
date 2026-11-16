"""Generic Excel/CSV column mapper adapting Invoicely's Smart Excel Column Mapper.

Supports .xlsx, .xlsm, .xls (via xlrd), and .csv.
Provides sheet selection, header-row detection, auto-mapping, and manual mapping.
Adapter converts mapping results to existing invoice schema (ParsedItem / InvoiceData).
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Literal

import openpyxl
from openpyxl.utils import get_column_letter

# Optional xlrd for .xls
try:
    import xlrd  # type: ignore[import-untyped]
    XLRD_AVAILABLE = True
except Exception:  # pragma: no cover
    xlrd = None  # type: ignore[assignment]
    XLRD_AVAILABLE = False

from .rab_parser import (
    ROLE_IGNORE,
    ROLE_LABELS,
    ROLE_NAME,
    ROLE_PRICE,
    ROLE_TOTAL_QTY,
    ROLE_UNIT,
    MappedColumn,
    RabMapping,
    ParseResult,
    RabParseError,
    to_number,
)
from ..models import ParsedItem, Project, Item
from .invoice_export import InvoiceData, InvoiceLine
from .invoice_builder import build_invoice_netto


# =============================================================================
# Role constants (extending existing with generic invoice fields)
# =============================================================================
ROLE_AMOUNT = "amount"
ROLE_INVOICE_NO = "invoice_no"
ROLE_CLIENT = "client"
ROLE_ADDRESS = "address"
ROLE_DATE = "date"
ROLE_DUE_DATE = "due_date"
ROLE_PO = "po"

ROLE_ALL_LABELS = {
    **ROLE_LABELS,
    ROLE_AMOUNT: "Jumlah/Totals",
    ROLE_INVOICE_NO: "Nomor Invoice",
    ROLE_CLIENT: "Nama Klien",
    ROLE_ADDRESS: "Alamat Klien",
    ROLE_DATE: "Tanggal Invoice",
    ROLE_DUE_DATE: "Jatuh Tempo",
    ROLE_PO: "Nomor PO",
}


# =============================================================================
# Synonyms for auto-mapping (normalized: lowercase, alphanumeric only)
# =============================================================================
COLUMN_SYNONYMS: dict[str, list[str]] = {
    "invoice_no": [
        "noinvoice", "invoice", "nomorinvoice", "invno", "no_invoice",
        "notransaksi", "kodeorder", "orderid", "id", "faktur", "nofaktur",
        "receipt", "struk", "ref", "kodestruk",
    ],
    "client": [
        "namaklien", "klien", "client", "customer", "pelanggan", "namapelanggan",
        "buyer", "pembeli", "perusahaan", "company", "kepada", "billto", "nama",
    ],
    "name": [
        "deskripsiitem", "deskripsi", "item", "produk", "jasa", "namabarang",
        "barang", "layanan", "service", "product", "description", "keterangan", "uraian",
    ],
    "price": [
        "hargasatuan", "harga", "price", "unitprice", "rate", "tarif", "biaya",
        "nominal", "amount", "nilai", "hargajual",
    ],
    "qty": [
        "qty", "jumlah", "kuantitas", "quantity", "vol", "volume", "unit", "pcs", "banyaknya",
    ],
    "unit": [
        "satuan", "sat", "unit", "uom",
    ],
    "date": [
        "tanggal", "date", "tgl", "tanggalinvoice", "invoicedate", "createdat", "waktu", "tgltransaksi",
    ],
    "address": [
        "alamatklien", "alamat", "address", "kota", "lokasi", "alamatpengiriman", "shippingaddress",
    ],
    "due_date": [
        "jatuhtempo", "duedate", "tempo", "batasbayar", "emailklien", "email",
    ],
    "po": [
        "nopo", "nomorpo", "purchaseorder", "ponumber", "no_po", "no_po_", "po",
    ],
    "amount": [
        "jumlahrharga", "totalharga", "jumlahbiaya", "amount", "nilai", "subtotal", "biaya",
        "total", "grandtotal", "grand_total",
    ],
}


def _normalize(text: str) -> str:
    """Normalize header for matching: lowercase, alphanumeric only."""
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _header_display(text: str) -> str:
    """Pretty header text for UI display."""
    return str(text or "").strip() or "(kosong)"


# =============================================================================
# WorkbookSource: unified interface for xlsx / xls / csv
# =============================================================================
class WorkbookSource:
    """Abstraction over openpyxl / xlrd / csv for reading sheets, grids, merges."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.suffix = self.path.suffix.lower()
        self._wb_openpyxl = None
        self._wb_xlrd = None
        self._csv_data: list[list[str]] | None = None

        if self.suffix in (".xlsx", ".xlsm"):
            # Load with formulas for merged cells info; also load data_only for values
            self._wb_openpyxl = openpyxl.load_workbook(str(self.path), data_only=False)
            self._wb_values = openpyxl.load_workbook(str(self.path), data_only=True)
        elif self.suffix == ".xls":
            if not XLRD_AVAILABLE:
                raise RabParseError("xlrd tidak tersedia. Install 'xlrd>=2.0' untuk .xls.")
            self._wb_xlrd = xlrd.open_workbook(str(self.path), formatting_info=True)
        elif self.suffix == ".csv":
            self._load_csv()
        else:
            raise RabParseError(f"Format file tidak didukung: {self.suffix}")

    # -------------------------------------------------------------------------
    # Public API
    # -------------------------------------------------------------------------
    @property
    def sheet_names(self) -> list[str]:
        if self._wb_openpyxl:
            return list(self._wb_openpyxl.sheetnames)
        if self._wb_xlrd:
            return list(self._wb_xlrd.sheet_names())
        if self._csv_data is not None:
            return ["Sheet1"]
        return []

    def grid(self, sheet: str) -> list[list[Any]]:
        """Return 2D grid (0-based rows/cols) of cell values for given sheet.
        For CSV: single sheet 'Sheet1'.
        """
        if self._wb_openpyxl:
            ws = self._wb_values[sheet]
            max_row = ws.max_row or 1
            max_col = ws.max_column or 1
            return [
                [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
                for r in range(1, max_row + 1)
            ]
        if self._wb_xlrd:
            ws = self._wb_xlrd.sheet_by_name(sheet)
            return [
                [ws.cell_value(r, c) for c in range(ws.ncols)]
                for r in range(ws.nrows)
            ]
        if self._csv_data is not None:
            return self._csv_data
        return [[]]

    def merged_ranges(self, sheet: str) -> list[tuple[int, int, int, int]]:
        """Return list of merged ranges as (min_row, min_col, max_row, max_col) 1-based.
        Empty for CSV.
        """
        if self._wb_openpyxl:
            ws = self._wb_openpyxl[sheet]
            return [
                (m.min_row, m.min_col, m.max_row, m.max_col)
                for m in ws.merged_cells.ranges
            ]
        if self._wb_xlrd:
            ws = self._wb_xlrd.sheet_by_name(sheet)
            ranges = []
            for crange in ws.merged_cells:
                rlo, rhi, clo, chi = crange
                # xlrd uses 0-based, half-open [rlo, rhi), [clo, chi)
                ranges.append((rlo + 1, clo + 1, rhi, chi))
            return ranges
        return []

    def cell_value(self, sheet: str, row: int, col: int) -> Any:
        """Get cell value (1-based row/col)."""
        if self._wb_openpyxl:
            ws = self._wb_values[sheet]
            return ws.cell(row=row, column=col).value
        if self._wb_xlrd:
            ws = self._wb_xlrd.sheet_by_name(sheet)
            if row <= ws.nrows and col <= ws.ncols:
                return ws.cell_value(row - 1, col - 1)
            return None
        if self._csv_data is not None:
            if 0 <= row - 1 < len(self._csv_data):
                r = self._csv_data[row - 1]
                if 0 <= col - 1 < len(r):
                    return r[col - 1]
            return None
        return None

    def max_row(self, sheet: str) -> int:
        if self._wb_openpyxl:
            return self._wb_values[sheet].max_row or 1
        if self._wb_xlrd:
            return self._wb_xlrd.sheet_by_name(sheet).nrows
        if self._csv_data is not None:
            return len(self._csv_data)
        return 1

    def max_col(self, sheet: str) -> int:
        if self._wb_openpyxl:
            return self._wb_values[sheet].max_column or 1
        if self._wb_xlrd:
            return self._wb_xlrd.sheet_by_name(sheet).ncols
        if self._csv_data is not None:
            return max((len(r) for r in self._csv_data), default=1)
        return 1

    # -------------------------------------------------------------------------
    # Private helpers
    # -------------------------------------------------------------------------
    def _load_csv(self) -> None:
        # Try common encodings and delimiters
        encodings = ("utf-8-sig", "utf-8", "latin-1", "cp1252")
        sample_size = 8192
        for enc in encodings:
            try:
                with open(self.path, "r", encoding=enc, newline="") as f:
                    sample = f.read(sample_size)
                    f.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                    except Exception:
                        dialect = csv.excel
                    reader = csv.reader(f, dialect)
                    self._csv_data = [list(row) for row in reader]
                    return
            except UnicodeDecodeError:
                continue
        # Fallback: latin-1 never fails
        with open(self.path, "r", encoding="latin-1", newline="") as f:
            try:
                dialect = csv.Sniffer().sniff(f.read(sample_size), delimiters=",;\t|")
                f.seek(0)
            except Exception:
                dialect = csv.excel
            reader = csv.reader(f, dialect)
            self._csv_data = [list(row) for row in reader]

    def close(self) -> None:
        if self._wb_openpyxl:
            self._wb_openpyxl.close()
            self._wb_values.close()
        if self._wb_xlrd:
            self._wb_xlrd.release_resources()


# =============================================================================
# Header detection & auto-mapping
# =============================================================================

# Reuse summary/term regexes from rab_parser
SUMMARY_RE = re.compile(
    r"^(sub\s*total|subtotal|total|grand\s*total|jumlah|ppn|pajak|"
    r"catatan|keterangan|rekap|ringkasan|biaya\s+lain)\b",
    re.IGNORECASE,
)
TERM_RE = re.compile(
    r"^(?:qty\s+)?(?:termin|term|tahap|bap|pembayaran)\s*(?:ke\s*)?[-:#.]?\s*"
    r"(\d+|[ivxlcdm]+)\b",
    re.IGNORECASE,
)
ROMAN_RE = re.compile(r"^[ivxlcdm]+$", re.IGNORECASE)
ROMAN_MAP = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}

NAME_KEYS = ("nama barang", "uraian", "deskripsi", "pekerjaan", "item", "barang")
UNIT_KEYS = ("satuan", "sat.", "unit", "uom")
PRICE_KEYS = ("harga satuan", "harga", "unit price", "harga unit")
QTY_KEYS = ("qty", "quantity", "volume", "jumlah", "banyak")
AMOUNT_KEYS = (
    "jumlah harga", "total harga", "jumlah biaya", "amount", "nilai", "subtotal", "biaya",
)


def _roman_to_int(text: str) -> int:
    text = text.lower()
    total = 0
    prev = 0
    for ch in reversed(text):
        value = ROMAN_MAP.get(ch, 0)
        if value < prev:
            total -= value
        else:
            total += value
            prev = value
    return total


def _term_label(text: str) -> str | None:
    raw = str(text or "").strip()
    m = TERM_RE.match(raw)
    if not m:
        return None
    token = m.group(1)
    number = _roman_to_int(token) if ROMAN_RE.match(token) else int(token)
    return f"Termin {number}"


def _classify_header(text: str) -> str:
    """Return role for a single header text (no parent row)."""
    h = _normalize(text)
    if not h:
        return ROLE_IGNORE

    # Pre-compute normalized keys (space removed, lowercase)
    def _nk(keys: tuple) -> list[str]:
        return [_normalize(k) for k in keys]

    # Priority order: specific multi-word keys first, then shorter ones.
    # Match: normalized key contained in normalized header (k in h), or exact (k == h).
    # NOTE: "h in k" (reverse) is intentionally NOT used — it causes false positives
    # e.g. "harga" matching "jumlah harga".

    # Nomor / index — check early to avoid matching "no" inside other words
    if h in ("no", "nomor", "urut"):
        return ROLE_IGNORE

    # Amount keywords (long, specific — check before price/qty)
    _amount_norm = _nk(AMOUNT_KEYS)
    if any(k == h or k in h for k in _amount_norm):
        return ROLE_AMOUNT

    # Term column
    if _term_label(text):
        return ROLE_TOTAL_QTY

    # Price keywords (harga satuan > harga)
    _price_norm = sorted(_nk(PRICE_KEYS), key=len, reverse=True)
    if any(k == h or k in h for k in _price_norm):
        return ROLE_PRICE

    # Name keywords
    _name_norm = _nk(NAME_KEYS)
    if any(k == h or k in h for k in _name_norm):
        return ROLE_NAME

    # Unit keywords (short — sat., unit, uom)
    _unit_norm = _nk(UNIT_KEYS)
    if any(k == h or k in h for k in _unit_norm):
        return ROLE_UNIT

    # Qty keywords
    _qty_norm = _nk(QTY_KEYS)
    if any(k == h or k in h for k in _qty_norm):
        return ROLE_TOTAL_QTY

    return ROLE_IGNORE


def _effective_header(
    source: WorkbookSource, sheet: str, header_row: int, col: int
) -> str:
    """Combine header row and parent row (like rab_parser.header_text)."""
    # base header
    base = str(source.cell_value(sheet, header_row, col) or "").strip()
    # parent row (one above)
    parent = ""
    if header_row > 1:
        parent = str(source.cell_value(sheet, header_row - 1, col) or "").strip()
    # check if parent is part of a merge covering header_row
    for mr, mc, Mr, Mc in source.merged_ranges(sheet):
        if mr <= header_row <= Mr and mc <= col <= Mc:
            # parent is anchor if merge starts above header_row
            if mr < header_row:
                parent = str(source.cell_value(sheet, mr, mc) or "").strip()
                break
    if parent and base and _normalize(parent) != _normalize(base):
        return f"{parent} — {base}"
    return base or parent


def detect_header_row(source: WorkbookSource, sheet: str, max_scan: int = 40) -> int:
    """Detect best header row (1-based) by scoring role coverage."""
    max_r = min(source.max_row(sheet), max_scan)
    best_row = 1
    best_score = -1
    for row in range(1, max_r + 1):
        cols = min(source.max_col(sheet), 80)
        headers = [_effective_header(source, sheet, row, c) for c in range(1, cols + 1)]
        roles = [_classify_header(h) for h in headers]
        has_name = ROLE_NAME in roles
        has_price = ROLE_PRICE in roles
        has_qty = ROLE_TOTAL_QTY in roles
        has_unit = ROLE_UNIT in roles
        term_count = sum(1 for h in headers if _term_label(h))
        if not (has_name and has_price and (has_qty or term_count)):
            continue
        score = 10 + term_count * 4 + (3 if has_qty else 0) + (1 if has_unit else 0)
        if score > best_score:
            best_score = score
            best_row = row
    if best_score < 0:
        # Fallback: first row with >=2 non-empty cells
        for row in range(1, max_r + 1):
            cols = min(source.max_col(sheet), 30)
            non_empty = sum(
                1 for c in range(1, cols + 1)
                if str(source.cell_value(sheet, row, c) or "").strip()
            )
            if non_empty >= 2:
                return row
    return best_row


def auto_map_columns(source: WorkbookSource, sheet: str, header_row: int) -> dict[str, int]:
    """Return {role: 1-based column index} using synonym matching + classification."""
    cols = min(source.max_col(sheet), 80)
    headers = [_effective_header(source, sheet, header_row, c) for c in range(1, cols + 1)]
    norm_headers = [_normalize(h) for h in headers]

    mapping: dict[str, int] = {}

    # Helper: find first column whose normalized header matches any synonym
    def find_by_synonyms(syns: list[str]) -> int | None:
        for i, nh in enumerate(norm_headers):
            for syn in syns:
                if syn == nh or syn in nh or nh.startswith(syn):
                    return i + 1
        return None

    # Required: name, price
    for role_key, syns in (
        ("name", COLUMN_SYNONYMS["name"]),
        ("price", COLUMN_SYNONYMS["price"]),
        ("qty", COLUMN_SYNONYMS["qty"]),
        ("unit", COLUMN_SYNONYMS["unit"]),
        ("invoice_no", COLUMN_SYNONYMS["invoice_no"]),
        ("client", COLUMN_SYNONYMS["client"]),
        ("address", COLUMN_SYNONYMS["address"]),
        ("date", COLUMN_SYNONYMS["date"]),
        ("due_date", COLUMN_SYNONYMS["due_date"]),
        ("po", COLUMN_SYNONYMS["po"]),
        ("amount", COLUMN_SYNONYMS["amount"]),
    ):
        col = find_by_synonyms(syns)
        if col:
            mapping[role_key] = col

    # Fallback classification for missing required roles
    for c in range(1, cols + 1):
        h = headers[c - 1]
        role = _classify_header(h)
        if role == ROLE_NAME and "name" not in mapping:
            mapping["name"] = c
        elif role == ROLE_PRICE and "price" not in mapping:
            mapping["price"] = c
        elif role == ROLE_TOTAL_QTY and "qty" not in mapping:
            mapping["qty"] = c
        elif role == ROLE_UNIT and "unit" not in mapping:
            mapping["unit"] = c
        elif role == ROLE_AMOUNT and "amount" not in mapping:
            mapping["amount"] = c

    return mapping


# =============================================================================
# MappingResult & Adapter
# =============================================================================

@dataclass
class MappingResult:
    """Result of applying a column mapping to a workbook."""
    path: Path
    sheet: str
    header_row: int
    columns: dict[str, int]          # role -> 1-based column index
    headers: list[str]               # display header texts per column
    rows: list[dict[str, Any]]       # each row: {role: value}
    items: list[ParsedItem] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)
    po_number: str = ""
    invoice_numbers: list[str] = field(default_factory=list)
    # raw grid for advanced use
    raw_rows: list[list[Any]] = field(default_factory=list)

    def validate(self) -> str | None:
        if not self.columns.get("name"):
            return "Kolom 'Nama Barang' wajib dipilih."
        if not self.columns.get("price") and not self.columns.get("amount"):
            return "Kolom 'Harga Satuan' atau 'Jumlah' wajib dipilih."
        return None


def parse_number(value: Any) -> float | None:
    """Parse number handling Indonesian formats, currency symbols, parentheses.
    Mirrors rab_parser.to_number with reference's extra cleaning.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    text = str(value).strip()
    if not text or text in ("-", "–", "—"):
        return None
    # Strip currency symbols, spaces, NBSP
    text = (
        text.replace("\xa0", "")
        .replace("Rp", "")
        .replace("rp", "")
        .replace("IDR", "")
        .replace(" ", "")
        .replace("$", "")
        .replace("€", "")
        .replace("£", "")
    )
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = text.rstrip("%")
    # Handle thousand/decimal separators
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            # 1.234,56
            text = text.replace(".", "").replace(",", ".")
        else:
            # 1,234.56
            text = text.replace(",", "")
    elif "." in text:
        # Could be 1.234 (thousand) or 1.23 (decimal)
        if re.fullmatch(r"[+-]?\d{1,3}(\.\d{3})+", text):
            text = text.replace(".", "")
    elif "," in text:
        # Could be 1,234 (thousand) or 1,23 (decimal)
        if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+", text):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
    # Remove any remaining non-numeric except e/E + - .
    text = re.sub(r"[^0-9eE+\-.]", "", text)
    if text in ("", ".", "-", "+"):
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def _round_qty(value: float) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _is_summary(text: str) -> bool:
    return bool(SUMMARY_RE.match(str(text or "").strip()))


def build_mapping_result(
    path: str | Path,
    *,
    sheet: str,
    header_row: int,
    columns: dict[str, int],
    limit: int | None = None,
) -> MappingResult:
    """Read workbook, apply mapping, and return MappingResult with parsed items."""
    path = Path(path)
    source = WorkbookSource(path)
    try:
        max_row = source.max_row(sheet)
        max_col = source.max_col(sheet)

        # Build header display texts for all columns
        headers = [_effective_header(source, sheet, header_row, c) for c in range(1, max_col + 1)]

        # Iterate data rows
        raw_rows: list[list[Any]] = []
        mapped_rows: list[dict[str, Any]] = []
        items: list[ParsedItem] = []
        issues: list[str] = []
        seen_invoices: set[str] = set()

        # Detect PO from rows above header (reuse rab_parser logic)
        po_number = ""
        if header_row > 1:
            # Use openpyxl sheet if available for detect_po_number
            if source._wb_values:
                po_number = detect_po_number(source._wb_values[sheet], header_row)
        if not po_number:
            # Generic scan first few rows
            for r in range(1, min(header_row, 30)):
                for c in range(1, min(max_col, 20) + 1):
                    v = source.cell_value(sheet, r, c)
                    if v:
                        text = str(v).strip()
                        low = _normalize(text)
                        for pref in ("nopo", "nomorpo", "purchaseorder", "ponumber", "po"):
                            if low.startswith(pref):
                                rest = text[len(pref):].lstrip(" :-=")
                                if rest:
                                    po_number = rest
                                    break
                        if po_number:
                            break
                if po_number:
                    break

        start_row = header_row + 1
        end_row = max_row if limit is None else min(max_row, start_row + limit - 1)

        for row in range(start_row, end_row + 1):
            # Raw row for debugging
            raw = [source.cell_value(sheet, row, c) for c in range(1, max_col + 1)]
            raw_rows.append(raw)

            # Build mapped dict
            row_data = {}
            for role, col in columns.items():
                if 1 <= col <= max_col:
                    row_data[role] = source.cell_value(sheet, row, col)
            mapped_rows.append(row_data)

            # Skip empty name rows
            name_val = row_data.get("name")
            if not name_val or not str(name_val).strip():
                continue
            name = str(name_val).strip()

            # Skip summary rows
            if _is_summary(name):
                continue

            # Parse numeric fields
            price = parse_number(row_data.get("price"))
            qty = parse_number(row_data.get("qty"))
            amount = parse_number(row_data.get("amount"))
            unit = str(row_data.get("unit") or "").strip()

            # Derive price from amount/qty if price missing
            if price is None and amount is not None and qty and qty != 0:
                price = amount / qty
            # Defaults
            if qty is None or qty == 0:
                qty = 1
            if price is None:
                price = 0

            # Warnings
            if isinstance(row_data.get("price"), (int, float, str)) and price is not None:
                # Check for decimal qty
                pass

            unit_price = _round_qty(price)
            total_qty = _round_qty(qty)

            items.append(ParsedItem(
                name=name,
                unit=unit,
                unit_price=unit_price,
                total_qty=total_qty,
                source_row=row,
            ))

            # Collect invoice numbers for metadata
            inv_no = row_data.get("invoice_no")
            if inv_no:
                inv_str = str(inv_no).strip()
                if inv_str:
                    seen_invoices.add(inv_str)

        return MappingResult(
            path=path,
            sheet=sheet,
            header_row=header_row,
            columns=columns,
            headers=headers,
            rows=mapped_rows,
            items=items,
            issues=issues,
            po_number=po_number,
            invoice_numbers=sorted(seen_invoices),
            raw_rows=raw_rows,
        )
    finally:
        source.close()


def to_parsed_items(result: MappingResult) -> list[ParsedItem]:
    """Adapter: return ParsedItem list (compatible with RabImportConfirmDialog)."""
    return result.items


def to_invoice_data(
    result: MappingResult,
    project: Project,
    *,
    company_name: str = "",
    company_address: str = "",
    po_number: str = "",
    invoice_date: date | None = None,
    date_format: str = "",
) -> InvoiceData | None:
    """Adapter: convert MappingResult + Project to InvoiceData for template.
    Reuses build_invoice_netto via temporary Item objects.
    """
    if not result.items:
        return None

    # Convert ParsedItem -> Item (subset fields needed by build_invoice_netto)
    items: list[Item] = []
    for i, pi in enumerate(result.items):
        items.append(Item(
            id=0,
            project_id=project.id,
            row_no=i + 1,
            name=pi.name,
            unit=pi.unit,
            unit_price=pi.unit_price,
            total_qty=pi.total_qty,
        ))

    # Use detected PO if not provided
    final_po = po_number or result.po_number

    return build_invoice_netto(
        project=project,
        items=items,
        company_name=company_name,
        company_address=company_address,
        label=project.name,
        po_number=final_po,
        invoice_date=invoice_date,
        date_format=date_format,
    )


# =============================================================================
# High-level import function used by UI
# =============================================================================

def import_excel(
    path: str | Path,
    *,
    sheet: str | None = None,
    header_row: int | None = None,
    mapping: dict[str, int] | None = None,
) -> MappingResult:
    """Convenience function: auto-detect sheet/header/mapping if not provided."""
    source = WorkbookSource(path)
    try:
        if sheet is None:
            sheet = source.sheet_names[0]
        if header_row is None:
            header_row = detect_header_row(source, sheet)
        if mapping is None:
            mapping = auto_map_columns(source, sheet, header_row)
        return build_mapping_result(path, sheet=sheet, header_row=header_row, columns=mapping)
    finally:
        source.close()


# Re-export detect_po_number from rab_parser for generic use
from .rab_parser import detect_po_number  # noqa: E402
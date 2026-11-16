"""Dialog pemetaan kolom universal (Excel/CSV) berbasis WorkbookSource.

Mendukung .xlsx, .xlsm, .xls (via xlrd), dan .csv.
Alur:
  1. ExcelMapperDialog  — pilih sheet, header row, mapping kolom → konfirmasi
  2. ExcelMapperConfirmDialog — ringkasan ParsedItem sebelum disimpan ke DB
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..models import ParsedItem
from ..services.excel_mapper import (
    ROLE_ALL_LABELS,
    ROLE_AMOUNT,
    ROLE_IGNORE,
    MappingResult,
    WorkbookSource,
    auto_map_columns,
    build_mapping_result,
    detect_header_row,
)
from ..services.rab_parser import RabParseError
from .style import GREEN, MUTED, RED
from .widgets import Banner, button, money_text, muted


# ---------------------------------------------------------------------------
# Peran yang bisa dipilih user (subset yang relevan untuk mapping kolom item)
# ---------------------------------------------------------------------------
_ROLE_CHOICES = [
    ("ignore",      "Abaikan"),
    ("name",        "Nama Barang ✱"),
    ("price",       "Harga Satuan"),
    ("qty",         "Qty / Volume"),
    ("unit",        "Satuan"),
    ("amount",      "Jumlah (harga × qty)"),
    ("invoice_no",  "No. Invoice"),
    ("client",      "Nama Klien"),
    ("address",     "Alamat Klien"),
    ("date",        "Tanggal Invoice"),
    ("due_date",    "Jatuh Tempo"),
    ("po",          "Nomor PO"),
]


class ExcelMapperDialog(QDialog):
    """Modal pemetaan kolom universal: xlsx / xls / csv.

    Setelah accepted, `self.result` berisi `MappingResult` siap pakai.
    """

    def __init__(self, path: str | Path, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Import File — Petakan Kolom")
        self.resize(1120, 780)
        self.path = Path(path)
        self.result: MappingResult | None = None

        self._loading = False
        try:
            self._source = WorkbookSource(self.path)
        except Exception as exc:  # noqa: BLE001
            raise RabParseError(f"File tidak bisa dibuka: {exc}") from exc

        self._build_ui()
        self._load_sheets()

    # ------------------------------------------------------------------ UI

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 14)
        root.setSpacing(10)

        # Title
        title = QLabel("Petakan kolom file ke field invoice")
        title.setStyleSheet("font-size: 14pt; font-weight: 600; color: #1F2937;")
        root.addWidget(title)

        subtitle = muted(
            f"File: {self.path.name}  •  "
            "Pilih sheet dan baris header, lalu tentukan peran setiap kolom. "
            "Kolom ✱ wajib. Auto-mapping dijalankan otomatis — koreksi bila perlu."
        )
        root.addWidget(subtitle)

        # --- Controls row ---
        ctrl = QHBoxLayout()
        ctrl.setSpacing(10)

        ctrl.addWidget(QLabel("Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(180)
        ctrl.addWidget(self.sheet_combo)

        ctrl.addSpacing(12)
        ctrl.addWidget(QLabel("Baris header:"))
        self.header_spin = QSpinBox()
        self.header_spin.setRange(1, 500)
        self.header_spin.setValue(1)
        self.header_spin.setFixedWidth(80)
        ctrl.addWidget(self.header_spin)

        detect_btn = button("🔍 Deteksi Otomatis")
        detect_btn.setToolTip("Deteksi baris header dan auto-mapping kolom")
        detect_btn.clicked.connect(self._detect_auto)
        ctrl.addWidget(detect_btn)

        ctrl.addStretch(1)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet(f"color: {MUTED}; font-size: 9pt;")
        ctrl.addWidget(self.status_label)
        root.addLayout(ctrl)

        # --- Splitter: mapping table (top) + data preview (bottom) ---
        splitter = QSplitter(Qt.Vertical)

        self.map_table = QTableWidget()
        self.map_table.setColumnCount(3)
        self.map_table.setHorizontalHeaderLabels(["Kolom", "Header File", "Peran"])
        self.map_table.verticalHeader().setVisible(False)
        self.map_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.map_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.map_table.setShowGrid(False)
        self.map_table.setAlternatingRowColors(True)
        h = self.map_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.Fixed)
        h.setSectionResizeMode(1, QHeaderView.Stretch)
        h.setSectionResizeMode(2, QHeaderView.Fixed)
        self.map_table.setColumnWidth(0, 70)
        self.map_table.setColumnWidth(2, 220)
        splitter.addWidget(self.map_table)

        preview_wrap = QWidget()
        pv_layout = QVBoxLayout(preview_wrap)
        pv_layout.setContentsMargins(0, 0, 0, 0)
        pv_layout.setSpacing(4)
        pv_layout.addWidget(muted("Pratinjau data (mulai dari baris header):"))

        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.preview_table.setShowGrid(False)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.verticalHeader().setDefaultSectionSize(24)
        pv_layout.addWidget(self.preview_table)
        splitter.addWidget(preview_wrap)

        splitter.setSizes([380, 260])
        root.addWidget(splitter, 1)

        # --- Bottom buttons ---
        bottom = QHBoxLayout()
        bottom.addStretch(1)
        cancel = QPushButton("Batal")
        cancel.clicked.connect(self.reject)
        bottom.addWidget(cancel)

        self.import_btn = button("Lanjut →", primary=True)
        self.import_btn.clicked.connect(self._accept)
        bottom.addWidget(self.import_btn)
        root.addLayout(bottom)

        # Signals
        self.sheet_combo.currentIndexChanged.connect(self._sheet_changed)
        self.header_spin.valueChanged.connect(self._header_changed)

    # ---------------------------------------------------------------- loading

    def _load_sheets(self) -> None:
        self._loading = True
        self.sheet_combo.clear()
        self.sheet_combo.addItems(self._source.sheet_names)
        self._loading = False
        self._sheet_changed()

    def _current_sheet(self) -> str:
        return self.sheet_combo.currentText()

    def _sheet_changed(self) -> None:
        if self._loading:
            return
        sheet = self._current_sheet()
        if not sheet:
            return
        self._loading = True
        max_r = max(1, min(self._source.max_row(sheet), 500))
        self.header_spin.setMaximum(max_r)
        self._loading = False
        self._detect_auto()

    def _header_changed(self) -> None:
        if self._loading:
            return
        sheet = self._current_sheet()
        if not sheet:
            return
        mapping = auto_map_columns(self._source, sheet, self.header_spin.value())
        self._populate_columns(mapping)

    def _detect_auto(self) -> None:
        sheet = self._current_sheet()
        if not sheet:
            return
        row = detect_header_row(self._source, sheet)
        mapping = auto_map_columns(self._source, sheet, row)
        self._loading = True
        self.header_spin.setValue(row)
        self._loading = False
        self._populate_columns(mapping)
        mapped_count = sum(1 for v in mapping.values() if v)
        self.status_label.setText(
            f"Auto-mapping: {mapped_count} kolom terdeteksi dari baris {row}."
        )

    # ---------------------------------------------------------------- columns

    def _populate_columns(self, mapping: dict[str, int]) -> None:
        """Populate mapping table from role->col dict."""
        sheet = self._current_sheet()
        if not sheet:
            return

        header_row = self.header_spin.value()
        max_col = min(self._source.max_col(sheet), 80)

        # Build reverse map: col -> role
        col_to_role: dict[int, str] = {}
        for role, col in mapping.items():
            if col and 1 <= col <= max_col:
                col_to_role[col] = role

        self.map_table.setRowCount(0)
        self.map_table.setRowCount(max_col)

        for idx in range(max_col):
            col = idx + 1
            # Column letter label
            letter = _col_letter(col)
            item_col = QTableWidgetItem(letter)
            item_col.setTextAlignment(Qt.AlignCenter)
            self.map_table.setItem(idx, 0, item_col)

            # Header text from file
            from ..services.excel_mapper import _effective_header
            header_val = _effective_header(self._source, sheet, header_row, col) or "(kosong)"
            item_hdr = QTableWidgetItem(header_val)
            item_hdr.setToolTip(header_val)
            self.map_table.setItem(idx, 1, item_hdr)

            # Role combo
            combo = QComboBox()
            for role_key, role_label in _ROLE_CHOICES:
                combo.addItem(role_label, role_key)
            detected_role = col_to_role.get(col, "ignore")
            idx_in_combo = combo.findData(detected_role)
            if idx_in_combo >= 0:
                combo.setCurrentIndex(idx_in_combo)
            self.map_table.setCellWidget(idx, 2, combo)

        self.map_table.resizeRowsToContents()
        self._populate_preview()

    def _populate_preview(self) -> None:
        sheet = self._current_sheet()
        if not sheet:
            return
        header_row = self.header_spin.value()
        max_col = min(self._source.max_col(sheet), 40)
        max_row = self._source.max_row(sheet)
        row_count = min(max(1, max_row - header_row), 12)

        self.preview_table.clear()
        self.preview_table.setRowCount(row_count)
        self.preview_table.setColumnCount(max_col)
        self.preview_table.setHorizontalHeaderLabels(
            [_col_letter(c) for c in range(1, max_col + 1)]
        )
        self.preview_table.setVerticalHeaderLabels(
            [str(header_row + i) for i in range(row_count)]
        )
        for r in range(row_count):
            for c in range(1, max_col + 1):
                val = self._source.cell_value(sheet, header_row + r, c)
                text = "" if val is None else str(val)
                item = QTableWidgetItem(text)
                item.setToolTip(text)
                self.preview_table.setItem(r, c - 1, item)
        self.preview_table.resizeColumnsToContents()
        for c in range(max_col):
            if self.preview_table.columnWidth(c) > 200:
                self.preview_table.setColumnWidth(c, 200)

    # ---------------------------------------------------------------- accept

    def _build_mapping_dict(self) -> dict[str, int]:
        """Collect {role: col_1based} from combo selections."""
        result: dict[str, int] = {}
        for row in range(self.map_table.rowCount()):
            combo = self.map_table.cellWidget(row, 2)
            if combo is None:
                continue
            role = combo.currentData()
            if role and role != "ignore":
                # Last assignment wins for duplicate roles (except amount)
                result[role] = row + 1
        return result

    def _validate(self, mapping: dict[str, int]) -> str | None:
        if not mapping.get("name"):
            return "Kolom 'Nama Barang ✱' wajib dipilih."
        if not mapping.get("price") and not mapping.get("amount"):
            return "Pilih minimal satu kolom 'Harga Satuan' atau 'Jumlah'."
        return None

    def _accept(self) -> None:
        mapping = self._build_mapping_dict()
        err = self._validate(mapping)
        if err:
            QMessageBox.warning(self, "Pemetaan belum lengkap", err)
            return

        sheet = self._current_sheet()
        header_row = self.header_spin.value()

        try:
            result = build_mapping_result(
                self.path,
                sheet=sheet,
                header_row=header_row,
                columns=mapping,
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Gagal memproses file", str(exc))
            return

        if not result.items:
            QMessageBox.warning(
                self,
                "Tidak ada item",
                "Tidak ada baris data yang terbaca. "
                "Periksa kembali baris header dan pemetaan kolom.",
            )
            return

        self.result = result
        self._source.close()
        super().accept()

    def closeEvent(self, event) -> None:  # type: ignore[override]
        try:
            self._source.close()
        except Exception:  # noqa: BLE001
            pass
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# Confirm dialog — ringkasan sebelum simpan ke DB
# ---------------------------------------------------------------------------

class ExcelMapperConfirmDialog(QDialog):
    """Ringkasan hasil mapping + konfirmasi sebelum simpan ke proyek."""

    def __init__(self, result: MappingResult, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Konfirmasi Import")
        self.setMinimumWidth(760)
        self.result = result
        self.accepted_items: list[ParsedItem] = list(result.items)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 14)
        root.setSpacing(10)

        title = QLabel("Data berhasil dibaca")
        title.setStyleSheet("font-size: 14pt; font-weight: 600; color: #1F2937;")
        root.addWidget(title)

        # Summary banner
        total_qty = sum(item.total_qty for item in result.items)
        total_value = sum(item.total_qty * item.unit_price for item in result.items)
        po_info = f"  •  PO: {result.po_number}" if result.po_number else ""
        inv_info = (
            f"  •  Invoice#: {', '.join(result.invoice_numbers[:3])}"
            if result.invoice_numbers
            else ""
        )
        summary = (
            f"{len(result.items)} baris item  •  "
            f"total qty {total_qty:,}  •  "
            f"nilai Rp {money_text(total_value)}"
            f"{po_info}{inv_info}"
        ).replace(",", ".")
        root.addWidget(Banner(summary, self))

        # Issues
        if result.issues:
            root.addWidget(
                QLabel(
                    "Ada hal yang perlu diperiksa. "
                    "Kamu tetap bisa import dan perbaiki di Excel lalu import ulang."
                )
            )
            issues_box = QPlainTextEdit()
            issues_box.setReadOnly(True)
            issues_box.setMinimumHeight(120)
            issues_box.setPlainText("\n".join(f"• {iss}" for iss in result.issues))
            root.addWidget(issues_box, 1)
        else:
            ok = QLabel("✓  Tidak ada masalah yang terdeteksi.")
            ok.setStyleSheet(f"color: {GREEN};")
            root.addWidget(ok)
            root.addStretch(1)

        # Preview table of items
        preview_label = QLabel(f"Pratinjau item ({min(len(result.items), 20)} baris pertama):")
        preview_label.setStyleSheet(f"color: {MUTED}; font-size: 9pt;")
        root.addWidget(preview_label)

        self.items_table = QTableWidget()
        self.items_table.setColumnCount(5)
        self.items_table.setHorizontalHeaderLabels(
            ["No", "Nama Barang", "Satuan", "Qty", "Harga Satuan"]
        )
        self.items_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.items_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.items_table.setAlternatingRowColors(True)
        self.items_table.setShowGrid(False)
        self.items_table.verticalHeader().setVisible(False)
        hdr = self.items_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        for c in (2, 3, 4):
            hdr.setSectionResizeMode(c, QHeaderView.Fixed)
        self.items_table.setColumnWidth(0, 50)
        self.items_table.setColumnWidth(2, 80)
        self.items_table.setColumnWidth(3, 70)
        self.items_table.setColumnWidth(4, 140)
        self.items_table.setMaximumHeight(240)

        display_items = result.items[:20]
        self.items_table.setRowCount(len(display_items))
        for row, item in enumerate(display_items):
            _set(self.items_table, row, 0, str(row + 1), Qt.AlignCenter)
            _set(self.items_table, row, 1, item.name)
            _set(self.items_table, row, 2, item.unit, Qt.AlignCenter)
            _set(self.items_table, row, 3, str(item.total_qty), Qt.AlignRight)
            _set(self.items_table, row, 4, money_text(item.unit_price), Qt.AlignRight)
        root.addWidget(self.items_table)

        # Acknowledgement for issues
        self.ack_check: QCheckBox | None = None
        if result.issues:
            self.ack_check = QCheckBox(
                "Saya sudah membaca masalah di atas dan tetap ingin import."
            )
            root.addWidget(self.ack_check)

        # Buttons
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel = QPushButton("Batal")
        cancel.clicked.connect(self.reject)
        buttons.addWidget(cancel)

        self.import_btn = button("Import ke Proyek", primary=True)
        self.import_btn.clicked.connect(self._do_accept)
        buttons.addWidget(self.import_btn)
        root.addLayout(buttons)

        self._update_button()
        if self.ack_check:
            self.ack_check.stateChanged.connect(self._update_button)

        self.adjustSize()

    def _update_button(self) -> None:
        ok = True
        if self.ack_check and not self.ack_check.isChecked():
            ok = False
        self.import_btn.setEnabled(ok)

    def _do_accept(self) -> None:
        if self.ack_check and not self.ack_check.isChecked():
            return
        super().accept()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_letter(col: int) -> str:
    """Convert 1-based column index to Excel letter (A, B, ..., Z, AA, ...)."""
    result = ""
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _set(
    table: QTableWidget,
    row: int,
    col: int,
    text: str,
    align: Qt.AlignmentFlag = Qt.AlignLeft | Qt.AlignVCenter,
) -> None:
    item = QTableWidgetItem(text)
    item.setTextAlignment(align)
    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
    table.setItem(row, col, item)

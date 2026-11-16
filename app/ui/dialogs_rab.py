"""Dialog-dialog untuk import RAB dan pengelolaan item."""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
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

from ..services.rab_parser import (
    ROLE_IGNORE,
    ROLE_LABELS,
    ROLE_NAME,
    ROLE_PRICE,
    ROLE_TERM,
    ROLE_TOTAL_QTY,
    ROLE_UNIT,
    MappedColumn,
    ParseResult,
    RabMapping,
    RabParseError,
    detect_best_header,
    detect_columns,
    header_text,
)
from .style import GREEN, MUTED, RED
from .widgets import Banner, button, money_text, muted


class RabMappingDialog(QDialog):
    """Pemetaan kolom Excel RAB ke peran aplikasi."""

    def __init__(self, path: str | Path, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Import RAB - Petakan Kolom")
        self.resize(1080, 760)
        self.path = Path(path)
        self.mapping: RabMapping | None = None
        self._loading = False
        try:
            self._wb = openpyxl.load_workbook(str(self.path), data_only=True)
        except Exception as exc:  # noqa: BLE001
            raise RabParseError(f"File tidak bisa dibuka: {exc}") from exc

        self._build_ui()
        self._load_sheets()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 14)
        root.setSpacing(10)

        title = QLabel("Petakan kolom Excel RAB")
        title.setStyleSheet("font-size: 14pt; font-weight: 600; color: #1F2937;")
        root.addWidget(title)
        subtitle = muted(
            "Pilih sheet dan baris header, lalu tentukan peran setiap kolom. "
            "Kolom termin dikenali tetapi diabaikan: invoice memakai qty total "
            "apa adanya. Baris penanda blok (mis. 'PEKERJAAN PERSIAPAN') "
            "otomatis dilewati."
        )
        root.addWidget(subtitle)

        top = QHBoxLayout()
        top.setSpacing(10)
        top.addWidget(QLabel("Sheet"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(180)
        top.addWidget(self.sheet_combo)
        top.addSpacing(12)
        top.addWidget(QLabel("Baris header"))
        self.header_spin = QSpinBox()
        self.header_spin.setRange(1, 500)
        self.header_spin.setValue(1)
        self.header_spin.setFixedWidth(80)
        top.addWidget(self.header_spin)
        detect_btn = button("Deteksi Otomatis")
        detect_btn.clicked.connect(self._detect_auto)
        top.addWidget(detect_btn)
        top.addStretch(1)
        root.addLayout(top)

        splitter = QSplitter(Qt.Vertical)
        self.map_table = QTableWidget()
        self.map_table.setColumnCount(3)
        self.map_table.setHorizontalHeaderLabels(
            ["Kolom", "Header Excel", "Peran"]
        )
        self.map_table.verticalHeader().setVisible(False)
        self.map_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.map_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.map_table.setShowGrid(False)
        self.map_table.setAlternatingRowColors(True)
        header = self.map_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        self.map_table.setColumnWidth(0, 70)
        self.map_table.setColumnWidth(2, 190)
        splitter.addWidget(self.map_table)

        preview_wrap = QWidget()
        preview_layout = QVBoxLayout(preview_wrap)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(4)
        preview_layout.addWidget(muted("Pratinjau data (baris header ke bawah):"))
        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.preview_table.setShowGrid(False)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.verticalHeader().setDefaultSectionSize(24)
        preview_layout.addWidget(self.preview_table)
        splitter.addWidget(preview_wrap)
        splitter.setSizes([360, 260])
        root.addWidget(splitter, 1)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        cancel = QPushButton("Batal")
        cancel.clicked.connect(self.reject)
        bottom.addWidget(cancel)
        self.import_btn = button("Lanjut", primary=True)
        self.import_btn.clicked.connect(self._accept)
        bottom.addWidget(self.import_btn)
        root.addLayout(bottom)

        self.sheet_combo.currentIndexChanged.connect(self._sheet_changed)
        self.header_spin.valueChanged.connect(self._header_changed)

    # --------------------------------------------------------------- loading
    def _load_sheets(self) -> None:
        self._loading = True
        self.sheet_combo.clear()
        self.sheet_combo.addItems(self._wb.sheetnames)
        self._loading = False
        self._sheet_changed()

    def _current_ws(self):
        name = self.sheet_combo.currentText()
        return self._wb[name] if name in self._wb.sheetnames else None

    def _sheet_changed(self) -> None:
        if self._loading:
            return
        ws = self._current_ws()
        if ws is None:
            return
        self._loading = True
        self.header_spin.setMaximum(max(1, min(ws.max_row or 1, 500)))
        self._loading = False
        self._detect_auto()

    def _header_changed(self) -> None:
        if self._loading:
            return
        ws = self._current_ws()
        if ws is None:
            return
        self._populate_columns(detect_columns(ws, self.header_spin.value()))

    def _detect_auto(self) -> None:
        ws = self._current_ws()
        if ws is None:
            return
        row, columns = detect_best_header(ws)
        self._loading = True
        self.header_spin.setValue(row)
        self._loading = False
        self._populate_columns(columns)

    def _populate_columns(self, columns: list[MappedColumn]) -> None:
        ws = self._current_ws()
        if ws is None:
            return
        self.map_table.setRowCount(0)
        self.map_table.setRowCount(len(columns))
        for index, mapped in enumerate(columns):
            col = mapped.col
            letter = get_column_letter(col)
            header_value = (
                header_text(ws, self.header_spin.value(), col) or "(kosong)"
            )
            item_col = QTableWidgetItem(letter)
            item_col.setTextAlignment(Qt.AlignCenter)
            self.map_table.setItem(index, 0, item_col)
            header_item = QTableWidgetItem(header_value)
            header_item.setToolTip(header_value)
            self.map_table.setItem(index, 1, header_item)

            combo = QComboBox()
            for role in (
                ROLE_NAME,
                ROLE_UNIT,
                ROLE_PRICE,
                ROLE_TOTAL_QTY,
                ROLE_TERM,
                ROLE_IGNORE,
            ):
                combo.addItem(ROLE_LABELS[role], role)
            combo.setCurrentIndex(max(0, combo.findData(mapped.role)))
            self.map_table.setCellWidget(index, 2, combo)

        self.map_table.resizeRowsToContents()
        self._populate_preview()

    def _populate_preview(self) -> None:
        ws = self._current_ws()
        if ws is None:
            return
        header_row = self.header_spin.value()
        max_col = min(ws.max_column or 1, 40)
        row_count = max(1, min((ws.max_row or header_row) - header_row, 10))
        self.preview_table.clear()
        self.preview_table.setRowCount(row_count)
        self.preview_table.setColumnCount(max_col)
        labels = [get_column_letter(c) for c in range(1, max_col + 1)]
        self.preview_table.setHorizontalHeaderLabels(labels)
        vlabels = [str(header_row + i) for i in range(row_count)]
        self.preview_table.setVerticalHeaderLabels(vlabels)
        for r in range(row_count):
            for c in range(1, max_col + 1):
                value = ws.cell(row=header_row + r, column=c).value
                text = "" if value is None else str(value)
                item = QTableWidgetItem(text)
                item.setToolTip(text)
                self.preview_table.setItem(r, c - 1, item)
        self.preview_table.resizeColumnsToContents()
        for c in range(max_col):
            if self.preview_table.columnWidth(c) > 220:
                self.preview_table.setColumnWidth(c, 220)

    # -------------------------------------------------------------- accepting
    def _build_mapping(self) -> RabMapping:
        columns: list[MappedColumn] = []
        term_counter = 0
        for row in range(self.map_table.rowCount()):
            combo = self.map_table.cellWidget(row, 2)
            role = combo.currentData() if combo else ROLE_IGNORE
            term_name = ""
            if role == ROLE_TERM:
                term_counter += 1
                term_name = f"Termin {term_counter}"
            columns.append(MappedColumn(row + 1, role, term_name))
        return RabMapping(
            sheet=self.sheet_combo.currentText(),
            header_row=self.header_spin.value(),
            columns=columns,
        )

    def _accept(self) -> None:
        mapping = self._build_mapping()
        error = mapping.validate()
        if error:
            QMessageBox.warning(self, "Pemetaan belum lengkap", error)
            return
        self.mapping = mapping
        super().accept()


class RabImportConfirmDialog(QDialog):
    """Ringkasan hasil parsing + konfirmasi pembulatan qty desimal."""

    def __init__(
        self,
        result: ParseResult,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Konfirmasi Import RAB")
        self.setMinimumWidth(720)
        self.result = result

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 14)
        root.setSpacing(10)

        title = QLabel("RAB berhasil dibaca")
        title.setStyleSheet("font-size: 14pt; font-weight: 600; color: #1F2937;")
        root.addWidget(title)

        total_qty = sum(item.total_qty for item in result.items)
        total_value = sum(item.total_qty * item.unit_price for item in result.items)
        summary = (
            f"{len(result.items)} baris RAB • total qty {total_qty:,} • "
            f"nilai RAB Rp {money_text(total_value)}"
        ).replace(",", ".")
        root.addWidget(Banner(summary, self))

        if result.term_names:
            root.addWidget(
                muted(
                    "Kolom termin terdeteksi ("
                    + ", ".join(result.term_names)
                    + ") tetapi tidak dipakai: invoice 1:1 memakai qty total."
                )
            )

        if result.issues:
            root.addWidget(
                QLabel(
                    "Ada hal yang perlu diperiksa. Kamu tetap bisa import, "
                    "lalu perbaiki di Excel dan import ulang."
                )
            )
            issues_box = QPlainTextEdit()
            issues_box.setReadOnly(True)
            issues_box.setMinimumHeight(160)
            issues_box.setPlainText("\n".join(f"• {issue}" for issue in result.issues))
            root.addWidget(issues_box, 1)
        else:
            ok = QLabel("Tidak ada masalah qty yang terdeteksi.")
            ok.setStyleSheet(f"color: {GREEN};")
            root.addWidget(ok)
            root.addStretch(1)

        self.decimal_check = None
        if result.decimal_rows:
            rows = ", ".join(str(r) for r in sorted(set(result.decimal_rows))[:12])
            self.decimal_check = QCheckBox(
                f"Saya setuju qty desimal (baris {rows}) dibulatkan ke bilangan bulat."
            )
            root.addWidget(self.decimal_check)

        self.ack_check = None
        if result.issues:
            self.ack_check = QCheckBox(
                "Saya sudah membaca masalah di atas dan tetap ingin import."
            )
            root.addWidget(self.ack_check)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel = QPushButton("Batal")
        cancel.clicked.connect(self.reject)
        buttons.addWidget(cancel)
        self.import_btn = button("Import RAB", primary=True)
        self.import_btn.clicked.connect(self._accept)
        buttons.addWidget(self.import_btn)
        root.addLayout(buttons)

        self._update_button()
        if self.decimal_check:
            self.decimal_check.stateChanged.connect(self._update_button)
        if self.ack_check:
            self.ack_check.stateChanged.connect(self._update_button)
        self.adjustSize()

    def _update_button(self) -> None:
        ok = True
        if self.decimal_check and not self.decimal_check.isChecked():
            ok = False
        if self.ack_check and not self.ack_check.isChecked():
            ok = False
        self.import_btn.setEnabled(ok)

    def _accept(self) -> None:
        if self.decimal_check and not self.decimal_check.isChecked():
            return
        if self.ack_check and not self.ack_check.isChecked():
            return
        super().accept()


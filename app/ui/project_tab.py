"""Tab data proyek: informasi proyek, import RAB, dan pratinjau item."""
from __future__ import annotations

from datetime import date

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..db import Database
from ..models import Project
from ..services.excel_mapper import import_excel
from ..services.rab_parser import (
    RabParseError,
    detect_po_number,
    parse_rab,
)
from ..services.validation import validate_project
from ..utils import format_rupiah
from .dialogs_excel_mapper import ExcelMapperConfirmDialog, ExcelMapperDialog
from .dialogs_rab import RabImportConfirmDialog, RabMappingDialog
from .style import GREEN, RED
from .widgets import Banner, Card, button, money_text, muted, setup_table


class ProjectTab(QWidget):
    data_changed = Signal()

    def __init__(self, db: Database, parent: QWidget | None = None):
        super().__init__(parent)
        self.db = db
        self.project: Project | None = None
        self._build_ui()
        self._set_enabled(False)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        outer.addWidget(scroll)

        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)
        scroll.setWidget(content)

        # --- informasi proyek
        info = Card("Informasi Proyek")
        form = QFormLayout()
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.name_edit = QLineEdit()
        self.client_edit = QLineEdit()
        self.client_edit.setPlaceholderText("mis. POLRESTA BANDUNG")
        self.address_edit = QLineEdit()
        self.address_edit.setPlaceholderText("Alamat klien yang tampil di invoice")
        self.po_edit = QLineEdit()
        self.po_edit.setPlaceholderText("mis. PO- (dipakai bila RAB tidak memuat No PO)")
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd MMM yyyy")
        self.date_format_combo = QComboBox()
        self.date_format_combo.addItem("Tanggal lengkap (1 Juli 2026)", "full")
        self.date_format_combo.addItem("Bulan dan tahun (Juli 2026)", "bulan_tahun")
        form.addRow("Nama proyek", self.name_edit)
        form.addRow("Kepada Yth (klien)", self.client_edit)
        form.addRow("Alamat klien", self.address_edit)
        form.addRow("Prefix No PO", self.po_edit)
        form.addRow("Tanggal invoice", self.date_edit)
        form.addRow("Format tanggal invoice", self.date_format_combo)
        info.add_layout(form)
        save_row = QHBoxLayout()
        save_row.addStretch(1)
        self.save_info_btn = button("Simpan Informasi", primary=True)
        self.save_info_btn.clicked.connect(self._save_info)
        save_row.addWidget(self.save_info_btn)
        info.add_layout(save_row)
        layout.addWidget(info)

        # --- RAB
        rab = Card("RAB Excel")
        self.rab_summary = QLabel("RAB belum diimpor.")
        self.rab_summary.setWordWrap(True)
        rab.add(self.rab_summary)
        self.rab_banner = Banner("", self)
        rab.add(self.rab_banner)
        rab_buttons = QHBoxLayout()
        rab_buttons.setSpacing(8)
        self.import_rab_btn = button("Import RAB", primary=True)
        self.import_rab_btn.clicked.connect(self._import_rab)
        rab_buttons.addWidget(self.import_rab_btn)
        self.import_universal_btn = button("Import Universal (xlsx/xls/csv)")
        self.import_universal_btn.setToolTip(
            "Import file Excel atau CSV dengan format bebas menggunakan mapper kolom universal.\n"
            "Mendukung .xlsx, .xlsm, .xls, .csv dengan berbagai struktur header."
        )
        self.import_universal_btn.clicked.connect(self._import_universal)
        rab_buttons.addWidget(self.import_universal_btn)
        rab_buttons.addStretch(1)
        rab.add_layout(rab_buttons)

        self.issues_table = QTableWidget()
        self.issues_table.setColumnCount(3)
        self.issues_table.setHorizontalHeaderLabels(["Baris", "Item", "Masalah"])
        setup_table(self.issues_table, read_only=True, row_height=27)
        self.issues_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.issues_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.issues_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.issues_table.setColumnWidth(0, 60)
        self.issues_table.setColumnWidth(1, 220)
        self.issues_table.setMinimumHeight(110)
        self.issues_table.setMaximumHeight(180)
        rab.add(self.issues_table)
        layout.addWidget(rab)

        # --- pratinjau item RAB
        items_card = Card("Item RAB (1 baris = 1 baris invoice)")
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(6)
        self.items_table.setHorizontalHeaderLabels(
            ["No", "Nama Barang", "Satuan", "Qty", "Harga Satuan", "Jumlah"]
        )
        setup_table(self.items_table, read_only=True, row_height=27)
        header = self.items_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col in (2, 3, 4, 5):
            header.setSectionResizeMode(col, QHeaderView.Fixed)
        self.items_table.setColumnWidth(0, 50)
        self.items_table.setColumnWidth(2, 80)
        self.items_table.setColumnWidth(3, 70)
        self.items_table.setColumnWidth(4, 130)
        self.items_table.setColumnWidth(5, 140)
        self.items_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.items_table.setMinimumHeight(220)
        items_card.add(self.items_table)
        self.items_note = muted(
            "Qty diambil dari volume RAB apa adanya. Kolom termin diabaikan; "
            "nilai invoice adalah netto tanpa PPN."
        )
        self.items_note.setWordWrap(True)
        items_card.add(self.items_note)
        layout.addWidget(items_card, 1)

    def _set_enabled(self, enabled: bool) -> None:
        for cls in (QLineEdit, QDateEdit, QComboBox):
            for widget in self.findChildren(cls):
                widget.setEnabled(enabled)
        self.save_info_btn.setEnabled(enabled)
        self.import_rab_btn.setEnabled(enabled)
        self.import_universal_btn.setEnabled(enabled)

    # --------------------------------------------------------------- project
    def set_project(self, project: Project | None) -> None:
        self.project = project
        self._set_enabled(project is not None)
        if not project:
            self.name_edit.clear()
            self.client_edit.clear()
            self.address_edit.clear()
            self.po_edit.clear()
            self.rab_summary.setText("RAB belum diimpor.")
            self.items_table.setRowCount(0)
            self.issues_table.setRowCount(0)
            self.rab_banner.set_state("info", "")
            return
        self.name_edit.setText(project.name)
        self.client_edit.setText(project.client_name)
        self.address_edit.setText(project.client_address)
        self.po_edit.setText(project.po_prefix)
        d = project.date_obj
        self.date_edit.setDate(QDate(d.year, d.month, d.day))
        fmt_index = self.date_format_combo.findData(project.date_format or "full")
        self.date_format_combo.setCurrentIndex(fmt_index if fmt_index >= 0 else 0)
        self.refresh()

    def refresh(self) -> None:
        if not self.project:
            return
        self.project = self.db.get_project(self.project.id)
        if not self.project:
            return
        items = self.db.list_items(self.project.id)
        total_qty = sum(item.total_qty for item in items)
        total_value = sum(item.total_qty * item.unit_price for item in items)

        if items:
            self.rab_summary.setText(
                f"<b>{len(items)} baris RAB</b> • total qty "
                f"<b>{f'{total_qty:,}'.replace(',', '.')}</b> • nilai RAB "
                f"<b>{format_rupiah(total_value)}</b>. "
                "Setiap baris akan menjadi satu baris invoice."
            )
        else:
            self.rab_summary.setText(
                "RAB belum diimpor. Klik <b>Import RAB</b> dan pilih file Excel."
            )

        self.items_table.setRowCount(len(items))
        for row, item in enumerate(items):
            values = [
                str(row + 1),
                item.name,
                item.unit,
                str(item.total_qty),
                money_text(item.unit_price),
                money_text(item.total_qty * item.unit_price),
            ]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if col == 0:
                    cell.setTextAlignment(Qt.AlignCenter)
                elif col >= 3:
                    cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                self.items_table.setItem(row, col, cell)

        issues = validate_project(self.db, self.project.id)
        self.issues_table.setRowCount(len(issues))
        for row, issue in enumerate(issues):
            values = [str(issue.row_no or ""), issue.item_name, issue.message]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if col == 0:
                    cell.setTextAlignment(Qt.AlignCenter)
                cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                if col == 2:
                    cell.setForeground(QColor(RED))
                self.issues_table.setItem(row, col, cell)
        self.issues_table.setVisible(bool(issues))

        if not items:
            self.rab_banner.set_state(
                "info", "Belum ada RAB. Import file Excel untuk mulai."
            )
        elif issues:
            self.rab_banner.set_state(
                "warning",
                f"{len(issues)} baris perlu diperiksa sebelum invoice diekspor.",
            )
        else:
            self.rab_banner.set_state(
                "success",
                f"Siap diekspor: {len(items)} baris RAB → {len(items)} baris invoice.",
            )
        self.rab_banner.setVisible(True)
        if not issues:
            self.issues_table.setVisible(False)

    def _save_info(self) -> None:
        if not self.project:
            return
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Belum lengkap", "Nama proyek wajib diisi.")
            return
        qdate = self.date_edit.date()
        self.db.update_project(
            self.project.id,
            name=name,
            client_name=self.client_edit.text().strip(),
            client_address=self.address_edit.text().strip(),
            po_prefix=self.po_edit.text().strip(),
            invoice_date=date(qdate.year(), qdate.month(), qdate.day()),
            date_format=self.date_format_combo.currentData() or "full",
        )
        self.db.log(self.project.id, "edit_proyek", "Informasi proyek diperbarui")
        self.project = self.db.get_project(self.project.id)
        self.refresh()
        self.data_changed.emit()

    # ------------------------------------------------------------------- RAB
    def _import_rab(self) -> None:
        if not self.project:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Pilih file Excel RAB", "", "Excel (*.xlsx *.xlsm)"
        )
        if not path:
            return
        try:
            dialog = RabMappingDialog(path, self)
        except RabParseError as exc:
            QMessageBox.critical(self, "Gagal membaca file", str(exc))
            return
        if dialog.exec() != QDialog.Accepted or not dialog.mapping:
            return
        try:
            result = parse_rab(path, dialog.mapping)
        except RabParseError as exc:
            QMessageBox.critical(self, "Gagal memproses RAB", str(exc))
            return
        confirm = RabImportConfirmDialog(result, self)
        if confirm.exec() != QDialog.Accepted:
            return

        if self.db.list_items(self.project.id):
            answer = QMessageBox.question(
                self,
                "Ganti RAB?",
                "Proyek ini sudah punya RAB. Import ulang akan MENGHAPUS seluruh "
                "item yang ada. Lanjutkan?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
        self.db.replace_rab(self.project.id, result.items)

        # Simpan nomor PO yang terdeteksi di area atas tabel RAB.
        po_number = ""
        try:
            import openpyxl

            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if dialog.mapping.sheet in wb.sheetnames:
                po_number = detect_po_number(
                    wb[dialog.mapping.sheet], dialog.mapping.header_row
                )
            wb.close()
        except Exception:  # noqa: BLE001
            po_number = ""
        self.db.set_setting(f"po_number_{self.project.id}", po_number)
        self.refresh()
        self.data_changed.emit()
        extra = f" Nomor PO terdeteksi: {po_number}." if po_number else ""
        QMessageBox.information(
            self,
            "Import selesai",
            f"{len(result.items)} baris RAB berhasil diimpor.{extra}",
        )

    # ---------------------------------------------------------------- Universal mapper

    def _import_universal(self) -> None:
        """Import file Excel/CSV apapun via mapper kolom universal."""
        if not self.project:
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Pilih file Excel atau CSV",
            "",
            "File data (*.xlsx *.xlsm *.xls *.csv);;Excel (*.xlsx *.xlsm *.xls);;CSV (*.csv)",
        )
        if not path:
            return

        # Step 1: Mapping dialog
        try:
            mapper_dialog = ExcelMapperDialog(path, self)
        except RabParseError as exc:
            QMessageBox.critical(self, "Gagal membaca file", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Gagal membuka file", f"Terjadi kesalahan:\n{exc}")
            return

        if mapper_dialog.exec() != QDialog.Accepted or mapper_dialog.result is None:
            return
        mapping_result = mapper_dialog.result

        # Step 2: Confirm dialog
        confirm = ExcelMapperConfirmDialog(mapping_result, self)
        if confirm.exec() != QDialog.Accepted:
            return

        # Step 3: Konfirmasi jika sudah ada data
        if self.db.list_items(self.project.id):
            answer = QMessageBox.question(
                self,
                "Ganti data?",
                "Proyek ini sudah punya data RAB. Import akan MENGHAPUS seluruh "
                "item yang ada dan menggantinya dengan data baru. Lanjutkan?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        # Step 4: Simpan ke DB
        self.db.replace_rab(self.project.id, mapping_result.items)

        # Simpan nomor PO yang terdeteksi
        po_number = mapping_result.po_number or ""
        self.db.set_setting(f"po_number_{self.project.id}", po_number)

        self.refresh()
        self.data_changed.emit()

        extra_parts = []
        if po_number:
            extra_parts.append(f"Nomor PO: {po_number}")
        if mapping_result.invoice_numbers:
            nums = ", ".join(mapping_result.invoice_numbers[:3])
            extra_parts.append(f"Invoice#: {nums}")
        extra = (" — " + "; ".join(extra_parts)) if extra_parts else ""
        QMessageBox.information(
            self,
            "Import selesai",
            f"{len(mapping_result.items)} baris berhasil diimpor{extra}.",
        )

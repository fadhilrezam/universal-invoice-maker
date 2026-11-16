"""Tab invoice 1:1: satu baris RAB menjadi satu baris invoice.

Tab ini tidak memakai termin maupun alokasi. Qty diambil dari
volume RAB, harga dari harga satuan, dan nilainya netto (PPN diurus user).
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from ..config import default_template_path
from ..db import Database
from ..models import Project
from ..services.invoice_builder import build_invoice_netto
from ..services.invoice_export import InvoiceExportError, export_invoices
from ..utils import format_rupiah
from .widgets import (
    Banner,
    Card,
    button,
    money_text,
    muted,
    set_item,
    setup_table,
)


class InvoiceTab(QWidget):
    """Preview + export invoice yang memetakan RAB apa adanya."""

    def __init__(self, db: Database, parent: QWidget | None = None):
        super().__init__(parent)
        self.db = db
        self.project: Project | None = None
        self.items = []
        self._last_output: Path | None = None
        self._prefilled_project_id: int | None = None
        self._build_ui()
        self._set_enabled(False)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        card = Card("Invoice 1:1 (netto)")

        top = QHBoxLayout()
        top.setSpacing(8)
        top.addWidget(QLabel("Nomor / Label"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("mis. Pengadaan Alat Tulis 2026")
        self.label_edit.setMinimumWidth(240)
        top.addWidget(self.label_edit)

        top.addWidget(QLabel("No PO"))
        self.po_edit = QLineEdit()
        self.po_edit.setPlaceholderText("kosongkan untuk otomatis")
        self.po_edit.setMinimumWidth(200)
        top.addWidget(self.po_edit)
        top.addStretch(1)
        card.add_layout(top)

        card.add(
            muted(
                "Satu baris RAB = satu baris invoice. Qty dari volume "
                "RAB. Termin dan PPN diabaikan — harga dianggap harga akhir."
            )
        )

        self.banner = Banner("", self)
        card.add(self.banner)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["No", "Nama Barang", "Qty", "Satuan", "Harga Satuan", "Jumlah"]
        )
        setup_table(self.table, read_only=True, row_height=28)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col in (2, 3, 4, 5):
            header.setSectionResizeMode(col, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(2, 70)
        self.table.setColumnWidth(3, 80)
        self.table.setColumnWidth(4, 130)
        self.table.setColumnWidth(5, 140)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        card.add(self.table)

        bottom = QHBoxLayout()
        bottom.setSpacing(8)
        bottom.addWidget(QLabel("Folder output"))
        self.output_edit = QLineEdit()
        bottom.addWidget(self.output_edit, 1)
        self.browse_btn = button("Pilih Folder...")
        self.browse_btn.clicked.connect(self._browse)
        bottom.addWidget(self.browse_btn)
        self.design_btn = button("🎨 Desain Dokumen A4...")
        self.design_btn.setToolTip("Kustomisasi tampilan dokumen A4, stempel, warna, dan cetak PDF.")
        self.design_btn.clicked.connect(self._open_designer)
        bottom.addWidget(self.design_btn)
        self.export_btn = button("Export Invoice", primary=True)
        self.export_btn.clicked.connect(self._export)
        bottom.addWidget(self.export_btn)
        card.add_layout(bottom)

        self.result_box = QPlainTextEdit()
        self.result_box.setReadOnly(True)
        self.result_box.setFixedHeight(96)
        self.result_box.setPlaceholderText("Hasil export akan tampil di sini.")
        card.add(self.result_box)

        root.addWidget(card, 1)

    def _set_enabled(self, enabled: bool) -> None:
        for widget in (self.label_edit, self.po_edit, self.browse_btn, self.design_btn, self.export_btn):
            widget.setEnabled(enabled)

    # --------------------------------------------------------------- project
    def set_project(self, project: Project | None) -> None:
        self.project = project
        self._set_enabled(project is not None)
        if not project:
            self.items = []
            self.table.setRowCount(0)
            self.banner.set_state("info", "")
            self.label_edit.clear()
            self.po_edit.clear()
            self._prefilled_project_id = None
            return
        if not self.output_edit.text().strip():
            default_dir = self.db.get_setting(
                "last_output_dir", str(Path.home() / "Documents" / "Invoice")
            )
            self.output_edit.setText(default_dir)
        self.refresh()

    def refresh(self) -> None:
        if not self.project:
            return
        self.project = self.db.get_project(self.project.id) or self.project
        self.items = self.db.list_items(self.project.id)
        if self._prefilled_project_id != self.project.id:
            self._prefilled_project_id = self.project.id
            self._prefill_fields()
        self._fill_table()
        self._update_banner()

    def _prefill_fields(self) -> None:
        """Isi label & No PO sekali per proyek; edit user tidak ditimpa.

        Nomor PO diambil dari RAB bila saat import terdeteksi; kalau tidak,
        dipakai ``prefix + label``.
        """
        assert self.project is not None
        label = self.project.name
        self.label_edit.setText(label)
        detected = self.db.get_setting(
            f"po_number_{self.project.id}", ""
        ).strip()
        self.po_edit.setText(detected or f"{self.project.po_prefix}{label}")

    def _fill_table(self) -> None:
        self.table.setRowCount(len(self.items))
        for row, item in enumerate(self.items):
            amount = int(item.total_qty) * int(item.unit_price)
            set_item(self.table, row, 0, str(row + 1), align=Qt.AlignCenter)
            set_item(self.table, row, 1, item.name)
            set_item(self.table, row, 2, str(int(item.total_qty)), align=Qt.AlignRight)
            set_item(self.table, row, 3, item.unit, align=Qt.AlignCenter)
            set_item(
                self.table,
                row,
                4,
                money_text(item.unit_price),
                align=Qt.AlignRight,
            )
            set_item(self.table, row, 5, money_text(amount), align=Qt.AlignRight)

    def _total(self) -> int:
        return sum(int(item.total_qty) * int(item.unit_price) for item in self.items)

    def _update_banner(self) -> None:
        if not self.items:
            self.banner.set_state(
                "warning", "Belum ada item RAB. Import RAB di tab Data Proyek."
            )
            self.export_btn.setEnabled(False)
            return
        self.banner.set_state(
            "info",
            f"{len(self.items)} baris RAB → {len(self.items)} baris invoice. "
            f"Total netto {format_rupiah(self._total())}.",
        )
        self.export_btn.setEnabled(True)

    # --------------------------------------------------------------- actions
    def _browse(self) -> None:
        start = self.output_edit.text().strip() or str(Path.home() / "Documents")
        folder = QFileDialog.getExistingDirectory(self, "Pilih folder output", start)
        if folder:
            self.output_edit.setText(folder)

    def _open_designer(self) -> None:
        if not self.project:
            return
        if not self.items:
            QMessageBox.warning(
                self, "Belum ada data", "Import RAB terlebih dahulu di tab Data Proyek."
            )
            return

        company_name = self.db.get_setting("company_name", "CV. PANJI PERKASA")
        company_address = self.db.get_setting(
            "company_address", "Kab. Bandung, Jawa Barat, Indonesia"
        )
        invoice = build_invoice_netto(
            self.project,
            self.items,
            company_name=company_name,
            company_address=company_address,
            label=self.label_edit.text(),
            po_number=self.po_edit.text(),
        )
        if invoice is None:
            QMessageBox.warning(
                self, "Tidak ada invoice", "RAB tidak berisi item yang bisa didesain."
            )
            return

        from .dialogs_invoice_designer import InvoiceDesignerDialog

        dialog = InvoiceDesignerDialog(invoice, self.db, self)
        dialog.exec()

    def _export(self) -> None:
        if not self.project:
            return
        if not self.items:
            QMessageBox.warning(
                self, "Belum ada data", "Import RAB terlebih dahulu di tab Data Proyek."
            )
            return
        output_text = self.output_edit.text().strip()
        if not output_text:
            QMessageBox.warning(
                self, "Folder kosong", "Pilih folder output terlebih dahulu."
            )
            return
        output_dir = Path(output_text)

        company_name = self.db.get_setting("company_name", "CV. PANJI PERKASA")
        company_address = self.db.get_setting(
            "company_address", "Kab. Bandung, Jawa Barat, Indonesia"
        )
        template = self.db.get_setting("template_path", str(default_template_path()))

        invoice = build_invoice_netto(
            self.project,
            self.items,
            company_name=company_name,
            company_address=company_address,
            label=self.label_edit.text(),
            po_number=self.po_edit.text(),
        )
        if invoice is None:
            QMessageBox.warning(
                self, "Tidak ada invoice", "RAB tidak berisi item yang bisa diekspor."
            )
            return

        self.export_btn.setEnabled(False)
        try:
            paths = export_invoices(template, [invoice], output_dir, mode="files")
        except InvoiceExportError as exc:
            QMessageBox.critical(self, "Export gagal", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Export gagal", f"Terjadi kesalahan: {exc}")
            return
        finally:
            self.export_btn.setEnabled(True)

        self.db.set_setting("last_output_dir", str(output_dir))
        self.db.log(
            self.project.id,
            "export_invoice_netto",
            f"{len(invoice.lines)} baris, total {invoice.total}",
        )
        self._last_output = output_dir
        lines = [
            f"Invoice dibuat: {len(invoice.lines)} baris RAB → "
            f"{len(invoice.lines)} baris invoice.",
            f"Total netto: {format_rupiah(invoice.total)}",
            "",
        ]
        lines.extend(f"• {path}" for path in paths)
        self.result_box.setPlainText("\n".join(lines))
        QMessageBox.information(
            self,
            "Export selesai",
            f"{len(paths)} file berhasil dibuat di:\n{output_dir}",
        )

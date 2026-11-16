"""Dialog pengaturan, riwayat, dan panduan."""
from __future__ import annotations

import os
import shutil
from pathlib import Path

import openpyxl
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from ..config import bundled_template_path, db_path, default_template_path
from ..db import Database
from ..models import Project
from ..services.template_profile import detect_header_row
from .widgets import button, muted, setup_table


class SettingsDialog(QDialog):
    def __init__(self, db: Database, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Pengaturan")
        self.resize(720, 360)
        self.db = db
        self.saved = False

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 14)
        root.setSpacing(12)
        title = QLabel("Pengaturan Aplikasi")
        title.setStyleSheet("font-size: 14pt; font-weight: 600; color: #1F2937;")
        root.addWidget(title)

        form = QFormLayout()
        form.setSpacing(10)
        self.company_edit = QLineEdit(
            db.get_setting("company_name", "CV. PANJI PERKASA")
        )
        self.address_edit = QLineEdit(
            db.get_setting("company_address", "Kab. Bandung, Jawa Barat, Indonesia")
        )
        self.template_edit = QLineEdit(
            db.get_setting("template_path", str(default_template_path()))
        )
        form.addRow("Nama perusahaan", self.company_edit)
        form.addRow("Alamat perusahaan", self.address_edit)

        template_row = QHBoxLayout()
        template_row.addWidget(self.template_edit, 1)
        browse = QPushButton("Pilih...")
        browse.clicked.connect(self._browse_template)
        template_row.addWidget(browse)
        open_btn = QPushButton("Buka")
        open_btn.clicked.connect(self._open_template)
        template_row.addWidget(open_btn)
        reset_btn = QPushButton("Reset")
        reset_btn.clicked.connect(self._reset_template)
        template_row.addWidget(reset_btn)
        form.addRow("Template invoice", template_row)

        db_label = QLineEdit(str(db_path()))
        db_label.setReadOnly(True)
        form.addRow("Lokasi data", db_label)
        root.addLayout(form)

        hint = muted(
            "Template invoice harus punya header 'Qty' dan 'Nama Barang' di baris "
            "header yang terdeteksi. Kolom A-D dipakai untuk tabel item. Baris total "
            "dan footer akan digeser otomatis sesuai jumlah item."
        )
        hint.setWordWrap(True)
        root.addWidget(hint)
        root.addStretch(1)

        row = QHBoxLayout()
        row.addStretch(1)
        cancel = QPushButton("Batal")
        cancel.clicked.connect(self.reject)
        row.addWidget(cancel)
        save = button("Simpan", primary=True)
        save.clicked.connect(self._save)
        row.addWidget(save)
        root.addLayout(row)

    def _browse_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Pilih template invoice", "", "Excel (*.xlsx)"
        )
        if path:
            self.template_edit.setText(path)

    def _open_template(self) -> None:
        path = Path(self.template_edit.text().strip())
        if not path.exists():
            QMessageBox.warning(self, "Tidak ditemukan", "File template tidak ditemukan.")
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Gagal membuka", str(exc))

    def _reset_template(self) -> None:
        source = bundled_template_path()
        target = default_template_path()
        if not source.exists():
            QMessageBox.warning(self, "Tidak ditemukan", "Template bawaan tidak ditemukan.")
            return
        try:
            if source.resolve() != target.resolve():
                shutil.copyfile(source, target)
            self.template_edit.setText(str(target))
            QMessageBox.information(self, "Template direset", "Template bawaan berhasil dipulihkan.")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Gagal reset", str(exc))

    def _save(self) -> None:
        template = self.template_edit.text().strip()
        path = Path(template)
        if not path.exists():
            QMessageBox.warning(self, "Template tidak ditemukan", "File template tidak ditemukan.")
            return
        try:
            wb = openpyxl.load_workbook(str(path), read_only=True)
            ws = wb.worksheets[0]
            header_row = detect_header_row(ws)
            wb.close()
            if header_row is None:
                raise ValueError("Header tabel tidak ditemukan")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Template tidak valid", f"Gagal membaca template: {exc}")
            return
        self.db.set_setting("company_name", self.company_edit.text().strip())
        self.db.set_setting("company_address", self.address_edit.text().strip())
        self.db.set_setting("template_path", template)
        self.saved = True
        super().accept()


class LogDialog(QDialog):
    def __init__(self, db: Database, project: Project, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Riwayat Aktivitas")
        self.resize(760, 520)
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 14)
        root.setSpacing(10)
        title = QLabel(f"Riwayat Aktivitas - {project.name}")
        title.setStyleSheet("font-size: 14pt; font-weight: 600; color: #1F2937;")
        root.addWidget(title)

        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Waktu", "Aksi", "Detail"])
        setup_table(table, read_only=True, row_height=28)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        table.setColumnWidth(0, 170)
        table.setColumnWidth(1, 160)
        rows = db.list_log(project.id)
        table.setRowCount(len(rows))
        for row, record in enumerate(rows):
            stamp = str(record["ts"]).replace("T", " ")
            table.setItem(row, 0, QTableWidgetItem(stamp))
            table.setItem(row, 1, QTableWidgetItem(record["action"]))
            table.setItem(row, 2, QTableWidgetItem(record["detail"]))
        root.addWidget(table, 1)

        close = button("Tutup", primary=True)
        close.clicked.connect(self.accept)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(close)
        root.addLayout(row)


HELP_HTML = """
<h2>Panduan Singkat</h2>
<p><b>Alur utama:</b> buat proyek → import RAB Excel → export invoice 1:1.</p>
<p><b>1. Buat proyek</b> lalu isi data klien, Prefix No PO, dan tanggal invoice.
Tanggal invoice punya dua format: <b>tanggal lengkap</b> (mis. "1 Juli 2026")
atau <b>bulan dan tahun</b> (mis. "Juli 2026").</p>
<p><b>2. Import RAB</b> dari Excel. Aplikasi mendeteksi header otomatis; koreksi
pemetaan kolom bila perlu. Kolom yang dikenali:</p>
<ul>
  <li><b>Nama Barang</b> — nama item/pekerjaan</li>
  <li><b>Satuan</b> — opsional (unit, pcs, liter, dll.)</li>
  <li><b>Harga Satuan</b> — harga per unit</li>
  <li><b>Qty Total</b> — volume pada blok "sesuai pelaksanaan"</li>
  <li><b>Termin</b> — dikenali tetapi <b>diabaikan</b> saat export 1:1</li>
</ul>
<p>Baris rekap, sub jumlah, PPN, dan grand total tidak dianggap item.</p>
<p><b>3. Tab Invoice 1:1</b>: tabel menampilkan hasil pemetaan. Satu baris RAB menjadi
tepat satu baris invoice — tanpa agregasi dan tanpa pemecahan qty. Isi
<b>Nomor / Label</b> dan <b>No PO</b> (kosongkan untuk otomatis), pilih folder output,
lalu klik <b>Export Invoice</b>.</p>
<p>Nilai invoice adalah <b>netto</b>: PPN tidak dihitung aplikasi. Harga dianggap harga
akhir, sesuai isi RAB.</p>
<h3>Lokasi data</h3>
<p>Data tersimpan lokal di <code>%APPDATA%\\UniversalInvoiceMaker\\data.db</code>.
Template invoice bisa diganti lewat menu Pengaturan.</p>
"""


class HelpDialog(QDialog):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Panduan")
        self.resize(780, 640)
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 14)
        browser = QTextBrowser()
        browser.setHtml(HELP_HTML)
        browser.setOpenExternalLinks(True)
        root.addWidget(browser, 1)
        close = button("Tutup", primary=True)
        close.clicked.connect(self.accept)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(close)
        root.addLayout(row)
"""Dialog Pratinjau & Kustomisasi Desain Dokumen Invoice A4.

Mengadopsi seluruh fitur dari reference-invoicely.html:
1. 4 Gaya Header: Modern Studio, Bold Banner, Korporat Klasik, Minimalis
2. 3 Gaya Tabel: Clean Lines, Zebra Striped, Grid Bordered
3. Stempel Status: LUNAS, BELUM BAYAR, DRAFT, CUSTOM
4. Customization Props: Warna Primer/Aksen, Header Background, Radius, Padding, Font Scale, dsb.
5. Multi-page A4 Pagination & Cetak / Ekspor PDF/HTML presisi.
"""
from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import (
    QCheckBox,
    QColorDialog,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSlider,
    QSpinBox,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from ..db import Database
from ..services.invoice_export import InvoiceData
from ..services.template_renderer import DesignConfig, render_invoice_html
from .style import ACCENT, GREEN, MUTED, RED
from .widgets import button, muted


class InvoiceDesignerDialog(QDialog):
    """Modal interaktif untuk mendesain output dokumen invoice A4."""

    def __init__(
        self,
        invoice: InvoiceData,
        db: Database,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.invoice = invoice
        self.db = db
        self.setWindowTitle(f"Desain & Pratinjau Invoice A4 — {invoice.label}")
        self.resize(1360, 880)

        # Load saved DesignConfig from database settings or default
        self.cfg = self._load_config()

        self._build_ui()
        self._refresh_preview()

    def _load_config(self) -> DesignConfig:
        raw = self.db.get_setting("invoice_design_config", "")
        if raw:
            try:
                data = json.loads(raw)
                return DesignConfig(**data)
            except Exception:
                pass
        return DesignConfig()

    def _save_config(self) -> None:
        raw = json.dumps(self.cfg.__dict__, ensure_ascii=False)
        self.db.set_setting("invoice_design_config", raw)

    def _build_ui(self) -> None:
        root = QHBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        splitter = QSplitter(Qt.Horizontal)

        # ----------------- PANEL KIRI: INSPECTOR PENGATURAN -----------------
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        form_layout = QVBoxLayout(scroll_content)
        form_layout.setContentsMargins(8, 8, 8, 8)
        form_layout.setSpacing(14)

        # 1. Layout & Header
        grp_layout = QGroupBox("1. Gaya Header & Layout A4")
        l1 = QFormLayout(grp_layout)
        l1.setSpacing(6)

        self.layout_combo = QComboBox()
        self.layout_combo.addItem("Modern Studio (Garis Aksen)", "modern")
        self.layout_combo.addItem("Bold Banner (Blok Warna Penuh)", "banner")
        self.layout_combo.addItem("Korporat Klasik (Garis Ganda Formal)", "classic")
        self.layout_combo.addItem("Minimalis (Clean & Simple)", "minimal")
        idx = self.layout_combo.findData(self.cfg.layout)
        if idx >= 0:
            self.layout_combo.setCurrentIndex(idx)
        self.layout_combo.currentIndexChanged.connect(self._on_control_changed)
        l1.addRow("Gaya Header", self.layout_combo)

        self.align_combo = QComboBox()
        self.align_combo.addItem("Kanan (Standar)", "default")
        self.align_combo.addItem("Kiri (Tukar Posisi)", "swapped")
        self.align_combo.addItem("Tengah (Centered)", "centered")
        idx_al = self.align_combo.findData(self.cfg.header_align)
        if idx_al >= 0:
            self.align_combo.setCurrentIndex(idx_al)
        self.align_combo.currentIndexChanged.connect(self._on_control_changed)
        l1.addRow("Posisi Header", self.align_combo)

        self.doc_title_edit = QLineEdit(self.cfg.doc_title)
        self.doc_title_edit.textChanged.connect(self._on_control_changed)
        l1.addRow("Judul Dokumen", self.doc_title_edit)

        form_layout.addWidget(grp_layout)

        # 2. Gaya Tabel Item
        grp_table = QGroupBox("2. Gaya Tabel Item")
        l2 = QFormLayout(grp_table)
        l2.setSpacing(6)

        self.table_style_combo = QComboBox()
        self.table_style_combo.addItem("Clean Lines (Garis Bawah Halus)", "clean")
        self.table_style_combo.addItem("Zebra Striped (Belang Halus)", "zebra")
        self.table_style_combo.addItem("Grid Bordered (Kotak Berbingkai)", "bordered")
        idx_tb = self.table_style_combo.findData(self.cfg.table_style)
        if idx_tb >= 0:
            self.table_style_combo.setCurrentIndex(idx_tb)
        self.table_style_combo.currentIndexChanged.connect(self._on_control_changed)
        l2.addRow("Gaya Tabel", self.table_style_combo)

        self.row_padding_spin = QSpinBox()
        self.row_padding_spin.setRange(4, 24)
        self.row_padding_spin.setValue(self.cfg.row_padding)
        self.row_padding_spin.valueChanged.connect(self._on_control_changed)
        l2.addRow("Kepadatan Baris (px)", self.row_padding_spin)

        self.chk_row_num = QCheckBox("Tampilkan Nomor Urut (#)")
        self.chk_row_num.setChecked(self.cfg.show_row_num)
        self.chk_row_num.toggled.connect(self._on_control_changed)
        l2.addRow(self.chk_row_num)

        form_layout.addWidget(grp_table)

        # 3. Warna & Palet
        grp_color = QGroupBox("3. Warna Dokumen")
        l3 = QFormLayout(grp_color)
        l3.setSpacing(6)

        # Color picker buttons
        self.btn_color_primary = QPushButton(self.cfg.color_primary)
        self.btn_color_primary.clicked.connect(lambda: self._pick_color("color_primary", self.btn_color_primary))
        l3.addRow("Warna Aksen", self.btn_color_primary)

        self.btn_color_thead = QPushButton(self.cfg.color_table_head)
        self.btn_color_thead.clicked.connect(lambda: self._pick_color("color_table_head", self.btn_color_thead))
        l3.addRow("Background Header", self.btn_color_thead)

        self.btn_color_box = QPushButton(self.cfg.color_box_bg)
        self.btn_color_box.clicked.connect(lambda: self._pick_color("color_box_bg", self.btn_color_box))
        l3.addRow("Background Kotak Total", self.btn_color_box)

        form_layout.addWidget(grp_color)

        # 4. Stempel / Status Watermark
        grp_stamp = QGroupBox("4. Stempel Status")
        l4 = QFormLayout(grp_stamp)
        l4.setSpacing(6)

        self.stamp_combo = QComboBox()
        self.stamp_combo.addItem("Tanpa Stempel", "none")
        self.stamp_combo.addItem("LUNAS (Hijau)", "LUNAS")
        self.stamp_combo.addItem("BELUM BAYAR (Merah)", "BELUM BAYAR")
        self.stamp_combo.addItem("DRAFT (Abu-abu)", "DRAFT")
        self.stamp_combo.addItem("Teks Stempel Kustom...", "CUSTOM")
        idx_st = self.stamp_combo.findData(self.cfg.stamp_status)
        if idx_st >= 0:
            self.stamp_combo.setCurrentIndex(idx_st)
        self.stamp_combo.currentIndexChanged.connect(self._on_control_changed)
        l4.addRow("Stempel", self.stamp_combo)

        self.stamp_custom_edit = QLineEdit(self.cfg.stamp_custom_text)
        self.stamp_custom_edit.setPlaceholderText("Ketik teks stempel...")
        self.stamp_custom_edit.textChanged.connect(self._on_control_changed)
        l4.addRow("Teks Kustom", self.stamp_custom_edit)

        form_layout.addWidget(grp_stamp)

        # 5. Visibilitas & Informasi Keuangan
        grp_vis = QGroupBox("5. Visibilitas & Finansial")
        l5 = QFormLayout(grp_vis)
        l5.setSpacing(6)

        self.chk_hero_total = QCheckBox("Kotak Hero Total")
        self.chk_hero_total.setChecked(self.cfg.show_hero_total)
        self.chk_hero_total.toggled.connect(self._on_control_changed)
        l5.addRow(self.chk_hero_total)

        self.chk_payment_box = QCheckBox("Kotak Instruksi Pembayaran")
        self.chk_payment_box.setChecked(self.cfg.show_payment_box)
        self.chk_payment_box.toggled.connect(self._on_control_changed)
        l5.addRow(self.chk_payment_box)

        self.chk_signature = QCheckBox("Blok Tanda Tangan")
        self.chk_signature.setChecked(self.cfg.show_signature)
        self.chk_signature.toggled.connect(self._on_control_changed)
        l5.addRow(self.chk_signature)

        self.tax_spin = QDoubleSpinBox()
        self.tax_spin.setRange(0.0, 100.0)
        self.tax_spin.setSuffix(" %")
        self.tax_spin.setValue(self.cfg.tax_rate)
        self.tax_spin.valueChanged.connect(self._on_control_changed)
        l5.addRow("PPN / Pajak", self.tax_spin)

        self.disc_spin = QDoubleSpinBox()
        self.disc_spin.setRange(0.0, 100.0)
        self.disc_spin.setSuffix(" %")
        self.disc_spin.setValue(self.cfg.discount_rate)
        self.disc_spin.valueChanged.connect(self._on_control_changed)
        l5.addRow("Diskon", self.disc_spin)

        self.bank_info_edit = QLineEdit(self.cfg.bank_info.replace("\n", " | "))
        self.bank_info_edit.textChanged.connect(self._on_control_changed)
        l5.addRow("Info Bank", self.bank_info_edit)

        self.signer_edit = QLineEdit(self.cfg.signer_name)
        self.signer_edit.textChanged.connect(self._on_control_changed)
        l5.addRow("Penandatangan", self.signer_edit)

        form_layout.addWidget(grp_vis)

        # 6. Paginasi A4 & Tipografi
        grp_page = QGroupBox("6. Paginasi & Kertas A4")
        l6 = QFormLayout(grp_page)
        l6.setSpacing(6)

        self.font_scale_spin = QSpinBox()
        self.font_scale_spin.setRange(80, 130)
        self.font_scale_spin.setSuffix(" %")
        self.font_scale_spin.setValue(self.cfg.font_scale)
        self.font_scale_spin.valueChanged.connect(self._on_control_changed)
        l6.addRow("Skala Font", self.font_scale_spin)

        self.items_per_page_spin = QSpinBox()
        self.items_per_page_spin.setRange(0, 100)
        self.items_per_page_spin.setSpecialValueText("Otomatis (Auto)")
        self.items_per_page_spin.setValue(self.cfg.items_per_page)
        self.items_per_page_spin.valueChanged.connect(self._on_control_changed)
        l6.addRow("Item per Lembar", self.items_per_page_spin)

        self.chk_editable = QCheckBox("Mode Klik untuk Edit (Inline)")
        self.chk_editable.setChecked(self.cfg.content_editable)
        self.chk_editable.toggled.connect(self._on_control_changed)
        l6.addRow(self.chk_editable)

        form_layout.addWidget(grp_page)

        scroll.setWidget(scroll_content)
        left_layout.addWidget(scroll, 1)

        # Bottom buttons on left panel
        btn_box = QHBoxLayout()
        self.btn_reset = QPushButton("Reset Desain")
        self.btn_reset.clicked.connect(self._reset_design)
        btn_box.addWidget(self.btn_reset)

        self.btn_save_pref = button("Simpan Preferensi", primary=True)
        self.btn_save_pref.clicked.connect(self._save_preferences)
        btn_box.addWidget(self.btn_save_pref)
        left_layout.addLayout(btn_box)

        splitter.addWidget(left_panel)

        # ----------------- PANEL KANAN: PREVIEW DOKUMEN A4 -----------------
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)

        # Top bar preview
        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("<b>Pratinjau Kertas A4 Multi-Halaman</b>"))
        top_bar.addStretch(1)

        self.btn_print_pdf = button("🖨️ Cetak / Ekspor PDF", primary=True)
        self.btn_print_pdf.clicked.connect(self._print_pdf)
        top_bar.addWidget(self.btn_print_pdf)

        self.btn_export_html = button("Simpan File HTML")
        self.btn_export_html.clicked.connect(self._export_html_file)
        top_bar.addWidget(self.btn_export_html)

        right_layout.addLayout(top_bar)

        # QWebEngineView for Live A4 Rendering
        self.web_view = QWebEngineView()
        right_layout.addWidget(self.web_view, 1)

        splitter.addWidget(right_panel)
        splitter.setSizes([420, 940])

        root.addWidget(splitter)

    def _pick_color(self, attr_name: str, btn: QPushButton) -> None:
        curr_hex = getattr(self.cfg, attr_name, "#2563eb")
        color = QColorDialog.getColor(curr_hex, self, "Pilih Warna")
        if color.isValid():
            hex_val = color.name()
            setattr(self.cfg, attr_name, hex_val)
            btn.setText(hex_val)
            btn.setStyleSheet(f"background-color: {hex_val}; color: #ffffff; font-weight: bold;")
            self._refresh_preview()

    def _sync_cfg_from_inputs(self) -> None:
        self.cfg.layout = self.layout_combo.currentData() or "modern"
        self.cfg.header_align = self.align_combo.currentData() or "default"
        self.cfg.doc_title = self.doc_title_edit.text().strip() or "INVOICE"
        self.cfg.table_style = self.table_style_combo.currentData() or "clean"
        self.cfg.row_padding = self.row_padding_spin.value()
        self.cfg.show_row_num = self.chk_row_num.isChecked()
        self.cfg.stamp_status = self.stamp_combo.currentData() or "none"
        self.cfg.stamp_custom_text = self.stamp_custom_edit.text().strip() or "VERIFIED"
        self.cfg.show_hero_total = self.chk_hero_total.isChecked()
        self.cfg.show_payment_box = self.chk_payment_box.isChecked()
        self.cfg.show_signature = self.chk_signature.isChecked()
        self.cfg.tax_rate = self.tax_spin.value()
        self.cfg.discount_rate = self.disc_spin.value()
        self.cfg.bank_info = self.bank_info_edit.text().replace(" | ", "\n")
        self.cfg.signer_name = self.signer_edit.text().strip()
        self.cfg.font_scale = self.font_scale_spin.value()
        self.cfg.items_per_page = self.items_per_page_spin.value()
        self.cfg.content_editable = self.chk_editable.isChecked()

    def _on_control_changed(self) -> None:
        self._sync_cfg_from_inputs()
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        html_content = render_invoice_html(self.invoice, self.cfg)
        self.web_view.setHtml(html_content, QUrl("qrc:/"))

    def _reset_design(self) -> None:
        self.cfg = DesignConfig()
        self._save_config()
        # Repopulate
        self.layout_combo.setCurrentIndex(0)
        self.align_combo.setCurrentIndex(0)
        self.doc_title_edit.setText(self.cfg.doc_title)
        self.table_style_combo.setCurrentIndex(0)
        self.row_padding_spin.setValue(self.cfg.row_padding)
        self.chk_row_num.setChecked(self.cfg.show_row_num)
        self.stamp_combo.setCurrentIndex(0)
        self.chk_hero_total.setChecked(self.cfg.show_hero_total)
        self.chk_payment_box.setChecked(self.cfg.show_payment_box)
        self.chk_signature.setChecked(self.cfg.show_signature)
        self.tax_spin.setValue(self.cfg.tax_rate)
        self.disc_spin.setValue(self.cfg.discount_rate)
        self._refresh_preview()

    def _save_preferences(self) -> None:
        self._sync_cfg_from_inputs()
        self._save_config()
        QMessageBox.information(self, "Tersimpan", "Preferensi desain berhasil disimpan.")

    def _print_pdf(self) -> None:
        suggested = f"Invoice {self.invoice.label or 'Invoice'}.pdf"
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Simpan sebagai PDF", suggested, "PDF (*.pdf)"
        )
        if not out_path:
            return
        self.web_view.printToPdf(out_path)
        QMessageBox.information(
            self, "Selesai", f"Dokumen berhasil dicetak ke PDF:\n{out_path}"
        )

    def _export_html_file(self) -> None:
        suggested = f"Invoice {self.invoice.label or 'Invoice'}.html"
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Simpan sebagai File HTML", suggested, "HTML (*.html)"
        )
        if not out_path:
            return
        html_content = render_invoice_html(self.invoice, self.cfg)
        Path(out_path).write_text(html_content, encoding="utf-8")
        QMessageBox.information(
            self, "Selesai", f"Dokumen HTML berhasil disimpan:\n{out_path}"
        )

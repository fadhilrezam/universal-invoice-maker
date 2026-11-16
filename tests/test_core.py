from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from app.db import Database
from app.models import ParsedItem, Item
from app.services.invoice_builder import build_invoice_netto
from app.services.invoice_export import InvoiceData, InvoiceLine, export_invoices
from app.services.rab_parser import (
    ROLE_NAME,
    ROLE_PRICE,
    ROLE_TERM,
    ROLE_TOTAL_QTY,
    ROLE_UNIT,
    RabMapping,
    MappedColumn,
    _round_qty,
    detect_best_header,
    detect_po_number,
    parse_rab,
    to_number,
)
from app.services.validation import validate_project


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "app" / "resources" / "invoice_template.xlsx"
SAMPLE_RAB = ROOT / "sample" / "rab_sample.xlsx"


def write_multi_section_rab(path: Path) -> Path:
    """Buat RAB uji dengan penanda blok (2 bagian, 2 termin).

    Dibuat di dalam test supaya tidak bergantung pada file contoh milik user
    (``sample/`` ada di .gitignore dan tidak ikut repo).
    """
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(
        [
            "No",
            "Uraian Pekerjaan",
            "Satuan",
            "Harga Satuan",
            "Jumlah Biaya",
            "Termin 1 Volume",
            "Termin 2 Volume",
        ]
    )
    ws.append([None, "YAMAHA VIXION (3 Unit)", None, None, None, None, None])
    ws.append([1, "Service", "UNIT", 240_000, 1_440_000, 3, 3])
    ws.append([2, "Oli Mesin", "LITER", 100_000, 600_000, 3, 3])
    ws.append([None, "SUZUKI HAYATE (2 Unit)", None, None, None, None, None])
    ws.append([1, "Service", "UNIT", 200_000, 800_000, 2, 2])
    ws.append([2, "Kampas Rem", "SET", 150_000, 300_000, 1, 1])
    wb.save(path)
    wb.close()
    return path


class NumberParsingTests(unittest.TestCase):
    def test_indonesian_formats(self):
        self.assertEqual(to_number("1.000.000"), 1_000_000)
        self.assertEqual(to_number("1,5"), 1.5)
        self.assertEqual(to_number("Rp 240.000"), 240_000)
        self.assertEqual(to_number("1.000,50"), 1000.5)
        self.assertEqual(to_number("1,000.50"), 1000.5)
        self.assertIsNone(to_number("-"))
        self.assertEqual(to_number(4), 4.0)


class RoundingTests(unittest.TestCase):
    def test_qty_round_half_up(self):
        self.assertEqual(_round_qty(2.4), 2)
        self.assertEqual(_round_qty(2.5), 3)
        self.assertEqual(_round_qty(2.6), 3)
        self.assertEqual(_round_qty(3.5), 4)


class RabParserTests(unittest.TestCase):
    def test_detect_and_parse_sample(self):
        import openpyxl

        wb = openpyxl.load_workbook(SAMPLE_RAB, data_only=True)
        row, columns = detect_best_header(wb["RAB"])
        self.assertEqual(row, 5)
        roles = {c.role for c in columns}
        self.assertIn(ROLE_NAME, roles)
        self.assertIn(ROLE_PRICE, roles)
        self.assertEqual(sum(1 for c in columns if c.role == ROLE_TERM), 4)

        mapping = RabMapping(sheet="RAB", header_row=row, columns=columns)
        result = parse_rab(SAMPLE_RAB, mapping)
        self.assertEqual(len(result.items), 12)
        self.assertEqual(result.term_names, ["Termin 1", "Termin 2", "Termin 3", "Termin 4"])
        service = result.items[0]
        self.assertEqual(service.name, "Service")
        self.assertEqual(service.unit_price, 240_000)
        self.assertEqual(service.total_qty, 20)
        self.assertEqual(result.issues, [])

    def test_manual_mapping_no_term(self):
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Nama", "Satuan", "Harga", "Volume"])
        ws.append(["Oli", "liter", 100000, 7])
        path = Path(tempfile.gettempdir()) / "rab_no_term.xlsx"
        wb.save(path)
        columns = [
            MappedColumn(1, ROLE_NAME),
            MappedColumn(2, ROLE_UNIT),
            MappedColumn(3, ROLE_PRICE),
            MappedColumn(4, ROLE_TOTAL_QTY),
        ]
        result = parse_rab(path, RabMapping("Sheet", 1, columns))
        self.assertEqual(result.term_names, [])
        self.assertEqual(result.items[0].total_qty, 7)

    def test_parse_multi_section_rab(self):
        import openpyxl

        with tempfile.TemporaryDirectory() as tmp:
            path = write_multi_section_rab(Path(tmp) / "rab_multi.xlsx")
            wb = openpyxl.load_workbook(path, data_only=True)
            row, columns = detect_best_header(wb["Sheet1"])
            self.assertEqual(row, 1)
            self.assertEqual(sum(1 for c in columns if c.role == ROLE_TERM), 2)
            wb.close()

            result = parse_rab(path, RabMapping("Sheet1", row, columns))
        self.assertEqual(result.term_names, ["Termin 1", "Termin 2"])
        self.assertEqual(len(result.items), 4)
        # Items should be flat, no groups
        self.assertEqual([item.name for item in result.items], ["Service", "Oli Mesin", "Service", "Kampas Rem"])
        self.assertEqual(result.items[0].total_qty, 6)
        self.assertEqual(result.issues, [])

    def test_section_headers_skipped(self):
        import openpyxl

        with tempfile.TemporaryDirectory() as tmp:
            path = write_multi_section_rab(Path(tmp) / "rab_multi.xlsx")
            wb = openpyxl.load_workbook(path, data_only=True)
            row, columns = detect_best_header(wb["Sheet1"])
            wb.close()
            result = parse_rab(path, RabMapping("Sheet1", row, columns))
        # Section header rows ("YAMAHA VIXION (3 Unit)", "SUZUKI HAYATE (2 Unit)") are skipped
        self.assertEqual(len(result.items), 4)
        self.assertEqual([item.source_row for item in result.items], [3, 4, 6, 7])


class DatabaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db = Database(Path(self.tmp.name) / "test.db")
        self.pid = self.db.create_project(
            "Proyek Uji", client_name="POLRESTA", client_address="Bandung"
        )

    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def test_replace_rab_and_validate(self):
        items = [
            ParsedItem("Service", "unit", 240000, 10, 2),
            ParsedItem("Oli", "liter", 100000, 6, 3),
        ]
        self.db.replace_rab(self.pid, items)
        self.assertEqual(validate_project(self.db, self.pid), [])
        stored = self.db.list_items(self.pid)
        self.assertEqual(len(stored), 2)
        self.assertEqual(stored[0].name, "Service")
        self.assertEqual(stored[0].total_qty, 10)

    def test_validate_negative_qty(self):
        items = [ParsedItem("Service", "unit", 240000, -1, 2)]
        self.db.replace_rab(self.pid, items)
        issues = validate_project(self.db, self.pid)
        self.assertEqual(len(issues), 1)
        self.assertIn("Qty harus lebih dari 0", issues[0].message)

    def test_validate_zero_qty(self):
        items = [ParsedItem("Service", "unit", 240000, 0, 2)]
        self.db.replace_rab(self.pid, items)
        issues = validate_project(self.db, self.pid)
        self.assertEqual(len(issues), 1)
        self.assertIn("Qty harus lebih dari 0", issues[0].message)

    def test_validate_negative_price(self):
        items = [ParsedItem("Service", "unit", -50000, 5, 2)]
        self.db.replace_rab(self.pid, items)
        issues = validate_project(self.db, self.pid)
        self.assertEqual(len(issues), 1)
        self.assertIn("Harga satuan negatif", issues[0].message)

    def test_invoice_date_format_options(self):
        import datetime

        from app.utils import format_date_id

        sample = datetime.date(2026, 7, 1)
        self.assertEqual(format_date_id(sample, "full"), "1 Juli 2026")
        self.assertEqual(format_date_id(sample, "bulan_tahun"), "Juli 2026")

        # Format default proyek bisa diubah.
        self.assertEqual(self.db.get_project(self.pid).date_format, "full")
        self.db.update_project(
            self.pid, invoice_date=sample, date_format="bulan_tahun"
        )
        project = self.db.get_project(self.pid)
        self.assertEqual(project.date_format, "bulan_tahun")
        self.assertEqual(project.date_obj, sample)

    def test_import_flat_rab(self):
        import openpyxl

        with tempfile.TemporaryDirectory() as tmp:
            path = write_multi_section_rab(Path(tmp) / "rab_multi.xlsx")
            wb = openpyxl.load_workbook(path, data_only=True)
            row, columns = detect_best_header(wb["Sheet1"])
            wb.close()
            result = parse_rab(path, RabMapping("Sheet1", row, columns))
            self.db.replace_rab(self.pid, result.items)
        stored = self.db.list_items(self.pid)
        self.assertEqual(len(stored), 4)
        self.assertEqual(validate_project(self.db, self.pid), [])
        # Items are flat, no section grouping
        self.assertEqual(stored[0].name, "Service")
        self.assertEqual(stored[1].name, "Oli Mesin")
        self.assertEqual(stored[2].name, "Service")
        self.assertEqual(stored[3].name, "Kampas Rem")


class PoDetectionTests(unittest.TestCase):
    """Nomor PO diambil dari area atas tabel RAB bila ada."""

    def _sheet(self, rows):
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        for row in rows:
            ws.append(row)
        return wb, ws

    def test_value_in_same_cell(self):
        wb, ws = self._sheet(
            [
                ["RENCANA ANGGARAN BIAYA", None, None],
                ["No. PO : VXN-2026-001", None, None],
                ["No", "Uraian Pekerjaan", "Harga Satuan"],
            ]
        )
        self.assertEqual(detect_po_number(ws, 3), "VXN-2026-001")
        wb.close()

    def test_value_in_next_cell(self):
        wb, ws = self._sheet(
            [
                ["Nomor PO", "PO-777", None],
                ["No", "Uraian Pekerjaan", "Harga Satuan"],
            ]
        )
        self.assertEqual(detect_po_number(ws, 2), "PO-777")
        wb.close()

    def test_not_found(self):
        wb, ws = self._sheet(
            [
                ["RAB Perawatan", None],
                ["No", "Uraian Pekerjaan"],
            ]
        )
        self.assertEqual(detect_po_number(ws, 2), "")
        wb.close()

    def test_no_false_positive_on_other_words(self):
        wb, ws = self._sheet(
            [
                ["Potongan harga", None],
                ["Pompa air", None],
                ["No", "Uraian Pekerjaan"],
            ]
        )
        self.assertEqual(detect_po_number(ws, 3), "")
        wb.close()


class TemplateProfileTests(unittest.TestCase):
    """Layout template disimpan sebagai data, bukan literal di kode export."""

    def test_detect_header_row_from_bundled_template(self):
        import openpyxl

        from app.services.template_profile import detect_header_row

        wb = openpyxl.load_workbook(str(TEMPLATE))
        ws = wb.worksheets[0]
        self.assertEqual(detect_header_row(ws), 10)
        wb.close()

    def test_profile_for_template(self):
        from app.services.template_profile import profile_for_template

        profile = profile_for_template(TEMPLATE)
        self.assertEqual(profile.header_row, 10)
        self.assertEqual(profile.item_first_row, 11)
        self.assertEqual(profile.columns["qty"], 1)
        self.assertEqual(profile.columns["name"], 2)
        self.assertEqual(profile.columns["unit_price"], 3)
        self.assertEqual(profile.columns["amount"], 4)

    def test_profile_roundtrip_and_fingerprint(self):
        from app.services.template_profile import TemplateProfile

        profile = TemplateProfile(name="Uji")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "profile.json"
            profile.save(path)
            loaded = TemplateProfile.load(path)
        self.assertEqual(loaded.to_dict(), profile.to_dict())
        header = ["No", "Nama Barang", "Satuan", "Harga Satuan", "Qty Total"]
        self.assertEqual(
            profile.fingerprint(header), loaded.fingerprint(header)
        )
        # Fingerprint berubah bila bentuk header berbeda.
        self.assertNotEqual(
            profile.fingerprint(header),
            profile.fingerprint(["No", "Uraian", "Satuan"]),
        )

    def test_export_uses_profile_layout(self):
        """Baris header hasil export mengikuti template, bukan angka hardcode."""
        import openpyxl

        from app.services.template_profile import profile_for_template

        profile = profile_for_template(TEMPLATE)
        invoice = InvoiceData(
            label="Uji",
            client_name="Klien",
            client_address="Alamat",
            company_name="CV. PANJI PERKASA",
            company_address="Bandung",
            po_number="PO-1",
            invoice_date=__import__("datetime").date(2026, 7, 1),
            lines=[InvoiceLine("Service", 2, 100_000)],
        )
        with tempfile.TemporaryDirectory() as tmp:
            paths = export_invoices(TEMPLATE, [invoice], tmp, mode="files")
            wb = openpyxl.load_workbook(paths[0], data_only=False)
            ws = wb.worksheets[0]
            header = profile.header_row
            self.assertEqual(ws.cell(row=header, column=1).value, "Qty")
            self.assertEqual(ws.cell(row=header, column=2).value, "Nama Barang")
            first = header + 1
            self.assertEqual(ws.cell(row=first, column=1).value, 2)
            self.assertEqual(ws.cell(row=first, column=2).value, "Service")
            self.assertEqual(ws.cell(row=first, column=4).value, 200_000)
            # Sub Total at first+1, TOTAL at first+2 (diskon row removed)
            self.assertEqual(ws.cell(row=first + 1, column=3).value, "Sub Total")
            self.assertEqual(ws.cell(row=first + 2, column=3).value, "TOTAL")
            wb.close()


class InvoiceNettoTests(unittest.TestCase):
    """Invoice 1:1 netto menggunakan build_invoice_netto() tanpa termin."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db = Database(Path(self.tmp.name) / "test.db")
        self.pid = self.db.create_project(
            "Proyek Uji Netto", client_name="POLRESTA", client_address="Bandung"
        )

    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def _item(self, name: str, qty: int, price: int) -> Item:
        return Item(
            id=0,
            project_id=self.pid,
            row_no=0,
            name=name,
            unit="unit",
            unit_price=price,
            total_qty=qty,
        )

    def test_mapping_passthrough(self):
        items = [self._item("Service", 3, 240_000)]
        project = self.db.get_project(self.pid)
        invoice = build_invoice_netto(
            project, items, label="Vixion"
        )
        self.assertIsNotNone(invoice)
        self.assertEqual(len(invoice.lines), 1)
        line = invoice.lines[0]
        self.assertEqual(line.name, "Service")
        self.assertEqual(line.qty, 3)
        self.assertEqual(line.unit_price, 240_000)
        self.assertEqual(line.amount, 720_000)
        self.assertEqual(invoice.total, 720_000)
        self.assertEqual(invoice.po_number, "Vixion")
        self.assertEqual(invoice.label, "Vixion")
        self.assertEqual(invoice.client_name, "POLRESTA")

    def test_multiple_items(self):
        items = [
            self._item("Oli", 6, 100_000),
            self._item("Kampas", 2, 150_000),
            self._item("Busi", 3, 50_000),
        ]
        project = self.db.get_project(self.pid)
        invoice = build_invoice_netto(project, items)
        self.assertIsNotNone(invoice)
        self.assertEqual(len(invoice.lines), 3)
        self.assertEqual(invoice.total, 1_050_000)
        self.assertEqual(invoice.lines[0].amount, 600_000)
        self.assertEqual(invoice.lines[1].amount, 300_000)
        self.assertEqual(invoice.lines[2].amount, 150_000)

    def test_skip_empty_name(self):
        items = [
            self._item("Service", 5, 240_000),
            Item(
                id=0,
                project_id=self.pid,
                row_no=0,
                name="",
                unit="unit",
                unit_price=100_000,
                total_qty=2,
            ),
            self._item("Oli", 3, 100_000),
        ]
        project = self.db.get_project(self.pid)
        invoice = build_invoice_netto(project, items)
        self.assertIsNotNone(invoice)
        self.assertEqual(len(invoice.lines), 2)
        self.assertEqual(invoice.lines[0].name, "Service")
        self.assertEqual(invoice.lines[1].name, "Oli")

    def test_no_items(self):
        items = []
        project = self.db.get_project(self.pid)
        invoice = build_invoice_netto(project, items)
        self.assertIsNone(invoice)


class InvoiceExportTests(unittest.TestCase):
    def test_build_invoice_with_many_lines(self):
        lines = [
            InvoiceLine(f"Barang {i}", i, 10_000) for i in range(1, 16)
        ]
        invoice = InvoiceData(
            label="D 1001 ABC",
            client_name="POLRESTA BANDUNG",
            client_address="JL BHAYANGKARA NO.1",
            company_name="CV. PANJI PERKASA",
            company_address="Kab. Bandung, Jawa Barat",
            po_number="VXN-D 1001 ABC",
            invoice_date=__import__("datetime").date(2026, 7, 1),
            invoice_date_text="1 Juli 2026",
            lines=lines,
        )
        with tempfile.TemporaryDirectory() as tmp:
            paths = export_invoices(TEMPLATE, [invoice], tmp, mode="files")
            self.assertEqual(len(paths), 1)
            import openpyxl

            wb = openpyxl.load_workbook(paths[0], data_only=False)
            ws = wb.active
            self.assertEqual(ws["A1"].value, "INVOICE")
            self.assertEqual(ws["C1"].value, "CV. PANJI PERKASA")
            self.assertEqual(ws["B5"].value, "1 Juli 2026")
            self.assertEqual(ws["D6"].value, "D 1001 ABC")
            self.assertEqual(ws["A11"].value, 1)
            self.assertEqual(ws["B11"].value, "Barang 1")
            self.assertEqual(ws["D11"].value, 10_000)
            self.assertEqual(ws["A25"].value, 15)
            self.assertEqual(ws["D25"].value, 150_000)
            subtotal_row = 26
            self.assertEqual(ws[f"C{subtotal_row}"].value, "Sub Total")
            self.assertEqual(ws[f"D{subtotal_row}"].value, 1_200_000)
            # Total row is subtotal+1 (no discount row)
            self.assertEqual(ws[f"C{subtotal_row + 1}"].value, "TOTAL")
            self.assertEqual(ws[f"D{subtotal_row + 1}"].value, 1_200_000)
            self.assertEqual(ws["C30"].value, "Hormat Kami,")
            self.assertEqual(ws["C33"].value, "CV. PANJI PERKASA")
            wb.close()

    def test_build_invoice_short(self):
        invoice = InvoiceData(
            label="D 2 B",
            client_name="Klien",
            client_address="Alamat",
            company_name="CV. PANJI PERKASA",
            company_address="Bandung",
            po_number="VXN-D 2 B",
            invoice_date=__import__("datetime").date(2026, 7, 1),
            lines=[InvoiceLine("Service", 1, 250_000)],
        )
        with tempfile.TemporaryDirectory() as tmp:
            paths = export_invoices(TEMPLATE, [invoice], tmp, mode="files")
            import openpyxl

            wb = openpyxl.load_workbook(paths[0])
            ws = wb.active
            self.assertEqual(ws["A11"].value, 1)
            self.assertEqual(ws["D11"].value, 250_000)
            self.assertEqual(ws["D12"].value, 250_000)  # sub total
            self.assertEqual(ws["C13"].value, "TOTAL")
            self.assertEqual(ws["C16"].value, "Hormat Kami,")
            wb.close()


class MultiTerminTemplateTests(unittest.TestCase):
    """Tiga layout RAB dengan posisi header berbeda dalam satu workbook.

    Kolom termin harus diabaikan: qty invoice diambil dari blok
    'SESUAI PELAKSANAAN' (volume total), bukan dari kolom termin.
    """

    SAMPLE = ROOT / "sample" / "Template_RAB_Multi_Termin.xlsx"

    def setUp(self):
        if not self.SAMPLE.exists():
            self.skipTest(f"file contoh tidak ada: {self.SAMPLE}")
        import openpyxl

        self.wb = openpyxl.load_workbook(str(self.SAMPLE), data_only=True)

    def tearDown(self):
        self.wb.close()

    def _parse(self, sheet, header_row):
        ws = self.wb[sheet]
        row, columns = detect_best_header(ws)
        self.assertEqual(row, header_row, f"header {sheet} harus di baris {header_row}")
        return parse_rab(str(self.SAMPLE), RabMapping(sheet, row, columns))

    def test_two_terms_header_row_1(self):
        result = self._parse("2_Termin_Header_Baris_1", 2)
        self.assertEqual(len(result.items), 7)
        first = result.items[0]
        self.assertEqual(first.name, "Item Pekerjaan / Jasa Utama 1")
        self.assertEqual(first.unit, "UNIT")
        self.assertEqual(first.unit_price, 250_000)
        self.assertEqual(first.total_qty, 6)
        self.assertEqual(result.term_names, ["Termin 1", "Termin 2"])

    def test_four_terms_header_row_2(self):
        result = self._parse("4_Termin_Header_Baris_2", 3)
        self.assertEqual(len(result.items), 7)
        self.assertEqual(result.items[0].total_qty, 6)
        self.assertEqual(result.items[2].unit_price, 150_000)

    def test_percentage_header_row_3(self):
        result = self._parse("Persentase_Header_Baris_3", 4)
        self.assertEqual(len(result.items), 4)
        self.assertEqual(result.items[0].name, "Pekerjaan Persiapan & Mobilisasi")
        self.assertEqual(result.items[0].unit, "LS")
        self.assertEqual(result.items[0].total_qty, 1)
        self.assertEqual(result.items[1].unit_price, 1_800_000)

    def test_summary_rows_are_not_items(self):
        result = self._parse("2_Termin_Header_Baris_1", 2)
        names = [item.name.lower() for item in result.items]
        for forbidden in ("jumlah", "total", "rekap", "ppn", "grand total"):
            self.assertFalse(
                any(name.startswith(forbidden) for name in names),
                f"baris rekap '{forbidden}' tidak boleh jadi item",
            )


if __name__ == "__main__":
    unittest.main()
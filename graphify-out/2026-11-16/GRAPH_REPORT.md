# Graph Report - universal-invoice-maker  (2026-11-16)

## Corpus Check
- 24 files · ~16,191 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 4 file(s) not represented in the graph (top: .bat 2, (none) 1, .spec 1)

## Summary
- 452 nodes · 1060 edges · 39 communities (12 shown, 27 thin omitted)
- Extraction: 88% EXTRACTED · 12% INFERRED · 0% AMBIGUOUS · INFERRED: 130 edges (avg confidence: 0.89)
- Token cost: 5,500 input · 9,400 output

## Graph Freshness
- Built from commit: `6ceaf07d`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- ExportTab
- main_window.py
- rab_parser.py
- build_invoice_netto
- main.py
- Database
- Invoice Excel Export
- Export Invoice Excel (Satu Workbook Multi-Sheet atau Satu File per Kendaraan)
- TemplateProfile
- Rencana Teknis — Universal Invoice Maker
- RabMappingDialog
- detect_po_number
- Asumsi & Batasan (Satu Pengguna, Offline, Output Hanya Excel)
- Build EXE dengan PyInstaller (build_exe.bat, one-file)
- Lokasi Data Lokal (%APPDATA%\InvoiceAutomation, data.db, error.log)
- Pelacakan Sisa Volume yang Belum Dialokasikan

## God Nodes (most connected - your core abstractions)
1. `Database` - 40 edges
2. `parse_rab()` - 27 edges
3. `MainWindow` - 24 edges
4. `InvoiceTab` - 20 edges
5. `Project` - 20 edges
6. `RabMappingDialog` - 18 edges
7. `ProjectTab` - 17 edges
8. `build_invoice_netto()` - 17 edges
9. `button()` - 16 edges
10. `build_invoice_workbook()` - 16 edges

## Surprising Connections (you probably didn't know these)
- `Catatan blocker (diperbarui)` --references--> `write_multi_section_rab()`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → tests/test_core.py
- `3.3 Model` --references--> `Item`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/models.py
- `9. Risiko` --references--> `to_number()`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/services/rab_parser.py
- `3.2 Mapping` --references--> `RabMappingDialog`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/ui/dialogs_rab.py
- `2.2 Parser RAB — `app/services/rab_parser.py`` --references--> `ParsedItem`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/models.py

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Pipeline Excel: RAB Masuk sampai Invoice Keluar** — readme_format_excel_rab, readme_deteksi_header_otomatis, readme_format_excel_kendaraan, readme_prefix_no_po_per_jenis, readme_template_invoice, readme_export_invoice_excel [INFERRED 0.85]

## Communities (39 total, 27 thin omitted)

### Community 0 - "ExportTab"
Cohesion: 0.09
Nodes (6): HelpDialog, SettingsDialog, InvoiceTab, MainWindow, NewProjectDialog, format_rupiah()

### Community 1 - "main_window.py"
Cohesion: 0.08
Nodes (14): Project, RabParseError, LogDialog, RabImportConfirmDialog, Banner, button(), Card, heading() (+6 more)

### Community 2 - "rab_parser.py"
Cohesion: 0.07
Nodes (32): _base_header_text(), cell_text(), _classify_grouped(), classify_header(), detect_best_header(), detect_columns(), detect_grouped_columns(), _formula_without_cache() (+24 more)

### Community 3 - "build_invoice_netto"
Cohesion: 0.17
Nodes (3): build_invoice_netto(), ProjectTab, InvoiceNettoTests

### Community 4 - "main.py"
Cohesion: 0.09
Nodes (12): app_data_dir(), bundled_template_path(), db_path(), default_template_path(), log_path(), user_template_path(), app_icon(), 4.4 Keputusan yang harus diambil: lokasi data user (+4 more)

### Community 5 - "Database"
Cohesion: 0.09
Nodes (7): Database, Item, ParsedItem, ValidationIssue, _check_item(), validate_project(), DatabaseTests

### Community 6 - "Invoice Excel Export"
Cohesion: 0.07
Nodes (24): invoice_summary_text(), build_invoice_workbook(), _clone_block(), _clone_row_style(), _copy_style(), export_invoices(), invoice_filename(), InvoiceData (+16 more)

### Community 7 - "Export Invoice Excel (Satu Workbook Multi-Sheet atau Satu File per Kendaraan)"
Cohesion: 0.17
Nodes (12): Alur Pemakaian (Proyek, Import RAB, Prefix No PO, Bagi, Kunci, Export), Export Invoice Excel (Satu Workbook Multi-Sheet atau Satu File per Kendaraan), Format Excel Kendaraan (No Kendaraan/Plat/Nopol, Tipe Kendaraan), Format Excel RAB (Header Dua Tingkat, Kolom Termin), Invoice Automation Desktop App, Pengujian: python -m unittest tests.test_core, Prefix No PO per Jenis Kendaraan (No PO = prefix + nomor kendaraan), RAB Multi-Jenis Kendaraan (Blok "NAMA JENIS (N Unit)") (+4 more)

### Community 9 - "TemplateProfile"
Cohesion: 0.12
Nodes (5): detect_header_row(), header_texts(), profile_for_template(), TemplateProfile, TemplateProfileTests

### Community 14 - "Rencana Teknis — Universal Invoice Maker"
Cohesion: 0.06
Nodes (26): 10. Pertanyaan terbuka, 12. Hasil template check — `sample/Template_RAB_Multi_Termin.xlsx`, 1. Keputusan scope yang sudah dikunci, 4.1 Tabel yang dipakai, 4.2 Tabel yang dihapus, 4.3 Strategi migrasi, 4. Perubahan data model (migrasi v2), 7. Rencana kerja & acceptance criteria (+18 more)

### Community 19 - "RabMappingDialog"
Cohesion: 0.20
Nodes (5): RabMappingDialog, 3.2 Mapping, 3.3 Model, 3.4 Output, 3. Arsitektur target

### Community 21 - "detect_po_number"
Cohesion: 0.31
Nodes (4): detect_po_number(), 11. Log status, Catatan blocker (diperbarui), PoDetectionTests

## Knowledge Gaps
- **6 isolated node(s):** `Asumsi & Batasan (Satu Pengguna, Offline, Output Hanya Excel)`, `Build EXE dengan PyInstaller (build_exe.bat, one-file)`, `Lokasi Data Lokal (%APPDATA%\InvoiceAutomation, data.db, error.log)`, `Pelacakan Sisa Volume yang Belum Dialokasikan`, `PySide6>=6.6` (+1 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 152 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **27 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Database` connect `Database` to `ExportTab`, `main_window.py`, `rab_parser.py`, `build_invoice_netto`, `main.py`?**
  _High betweenness centrality (0.130) - this node is a cross-community bridge._
- **Are the 9 inferred relationships involving `Database` (e.g. with `Item` and `ParsedItem`) actually correct?**
  _`Database` has 9 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Asumsi & Batasan (Satu Pengguna, Offline, Output Hanya Excel)`, `Build EXE dengan PyInstaller (build_exe.bat, one-file)`, `Lokasi Data Lokal (%APPDATA%\InvoiceAutomation, data.db, error.log)` to the rest of the system?**
  _6 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `ExportTab` be split into smaller, more focused modules?**
  _Cohesion score 0.0873015873015873 - nodes in this community are weakly interconnected._
- **Why does `Rencana Teknis — Universal Invoice Maker` connect `Rencana Teknis — Universal Invoice Maker` to `rab_parser.py`, `RabMappingDialog`, `detect_po_number`, `Invoice Excel Export`?**
  _High betweenness centrality (0.089) - this node is a cross-community bridge._
- **Are the 3 inferred relationships involving `parse_rab()` (e.g. with `.first()` and `.terms()`) actually correct?**
  _`parse_rab()` has 3 INFERRED edges - model-reasoned connections that need verification._
- **Should `main_window.py` be split into smaller, more focused modules?**
  _Cohesion score 0.07540983606557378 - nodes in this community are weakly interconnected._
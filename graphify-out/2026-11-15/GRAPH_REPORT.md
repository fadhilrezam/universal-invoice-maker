# Graph Report - universal-invoice-maker  (2026-11-15)

## Corpus Check
- 29 files · ~30,118 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 4 file(s) not represented in the graph (top: .bat 2, (none) 1, .spec 1)

## Summary
- 607 nodes · 1715 edges · 18 communities (10 shown, 8 thin omitted)
- Extraction: 82% EXTRACTED · 18% INFERRED · 0% AMBIGUOUS · INFERRED: 306 edges (avg confidence: 0.89)
- Token cost: 5,500 input · 9,400 output

## Graph Freshness
- Built from commit: `6ceaf07d`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- MainWindow
- project_tab.py
- RAB Parser & Test Suite
- RabMappingDialog
- Database
- Invoice Excel Export
- Project Docs (PRD & README)
- RAB Import Dialogs
- Vehicle File Parser
- AllocationTab
- RemainderDialog
- Rencana Teknis — Universal Invoice Maker

## God Nodes (most connected - your core abstractions)
1. `Database` - 92 edges
2. `AllocationTab` - 41 edges
3. `ProjectTab` - 37 edges
4. `Project` - 31 edges
5. `ExportTab` - 29 edges
6. `parse_rab()` - 28 edges
7. `MainWindow` - 26 edges
8. `Banner` - 25 edges
9. `button()` - 25 edges
10. `muted()` - 22 edges

## Surprising Connections (you probably didn't know these)
- `9. Risiko` --references--> `to_number()`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/services/rab_parser.py
- `3.2 Mapping` --references--> `RabMappingDialog`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/ui/dialogs_rab.py
- `3.3 Model` --references--> `Item`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/models.py
- `6. Daftar file: hapus / ubah / baru` --references--> `match_section_header()`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/services/rab_parser.py
- `3.4 Output` --references--> `Item`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/models.py

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Pipeline Excel: RAB Masuk sampai Invoice Keluar** — readme_format_excel_rab, readme_deteksi_header_otomatis, readme_format_excel_kendaraan, readme_prefix_no_po_per_jenis, readme_template_invoice, readme_export_invoice_excel [INFERRED 0.85]
- **Stack Desktop Offline (Python, PySide6, openpyxl, SQLite, PyInstaller)** — prd_aplikasi_pembagian_rab_ke_kendaraan_untuk_invoice_v0_1_rekomendasi_teknis, prd_aplikasi_pembagian_rab_ke_kendaraan_untuk_invoice_v0_1_sqlite_penyimpanan_lokal, requirements_pyside6, requirements_openpyxl, readme_build_exe, readme_lokasi_data [INFERRED 0.85]
- **Alur Pembagian Qty dan Invarian Alokasi** — prd_aplikasi_pembagian_rab_ke_kendaraan_untuk_invoice_v0_1_aturan_pembagian, prd_aplikasi_pembagian_rab_ke_kendaraan_untuk_invoice_v0_1_pembagian_sisa, prd_aplikasi_pembagian_rab_ke_kendaraan_untuk_invoice_v0_1_saran_default_kendaraan, prd_aplikasi_pembagian_rab_ke_kendaraan_untuk_invoice_v0_1_invarian_wajib, readme_aturan_pembagian, readme_invarian_hard [INFERRED 0.90]

## Communities (18 total, 8 thin omitted)

### Community 0 - "MainWindow"
Cohesion: 0.06
Nodes (14): app_data_dir(), bundled_template_path(), db_path(), default_template_path(), log_path(), user_template_path(), SettingsDialog, MainWindow (+6 more)

### Community 1 - "project_tab.py"
Cohesion: 0.06
Nodes (29): Project, allocate_item(), allocations_for_choice(), build_remainder_plan(), rank_vehicles(), RemainderPlan, split_qty(), validate_choice() (+21 more)

### Community 2 - "RAB Parser & Test Suite"
Cohesion: 0.06
Nodes (35): ParsedGroup, ParsedItem, _base_header_text(), cell_text(), _classify_grouped(), classify_header(), detect_best_header(), detect_columns() (+27 more)

### Community 3 - "RabMappingDialog"
Cohesion: 0.13
Nodes (3): ItemsDialog, RabImportConfirmDialog, RabMappingDialog

### Community 5 - "Database"
Cohesion: 0.07
Nodes (5): Database, VehicleType, ProjectTab, date_to_iso(), DatabaseAndValidationTests

### Community 6 - "Invoice Excel Export"
Cohesion: 0.06
Nodes (24): build_invoice_workbook(), build_pivot_sheets(), build_rekap_sheet(), _clone_block(), _clone_row_style(), _copy_style(), export_invoices(), _format_tanggal_pendek() (+16 more)

### Community 7 - "Project Docs (PRD & README)"
Cohesion: 0.06
Nodes (29): Alur Utama (Buat Proyek, Import, Validasi, Bagi, Review, Kunci, Export), Asumsi: Satu RAB, Satu List Kendaraan Tetap, Harga dari RAB, Export Excel Invoice per Termin, Kebutuhan Fungsional F1-F10, Kebutuhan Non-Fungsional (Offline, Lokal, Deterministik, Error Informatif), Kunci Termin (Lock Alokasi Sebelum Export), Pelacakan Alokasi antar Termin / Sisa Volume Belum Ditagihkan, Aplikasi Pembagian RAB ke Kendaraan untuk Invoice (+21 more)

### Community 9 - "Vehicle File Parser"
Cohesion: 0.16
Nodes (8): _cell(), _find_header(), _norm(), parse_vehicle_file(), VehicleParseError, VehicleParseResult, build_rab(), build_vehicles()

### Community 10 - "AllocationTab"
Cohesion: 0.07
Nodes (12): AllocIssue, Item, Term, Vehicle, build_invoice_data(), _item_label(), term_totals(), validate_project() (+4 more)

### Community 14 - "Rencana Teknis — Universal Invoice Maker"
Cohesion: 0.05
Nodes (33): 10. Pertanyaan terbuka, 11. Log status, 1. Keputusan scope yang sudah dikunci, 2.1 Skema database — `app/db.py`, 2.3 Template invoice — `app/services/invoice_export.py`, 2. Kondisi kode saat ini (terverifikasi), 3.2 Mapping, 3.3 Model (+25 more)

## Knowledge Gaps
- **10 isolated node(s):** `Alur Utama (Buat Proyek, Import, Validasi, Bagi, Review, Kunci, Export)`, `Kunci Termin (Lock Alokasi Sebelum Export)`, `Aplikasi Pembagian RAB ke Kendaraan untuk Invoice`, `Pertanyaan Terbuka (Qty Non-Bulat, Perubahan List Kendaraan, Isi Invoice)`, `Ruang Lingkup v1 (In/Out of Scope)` (+5 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 154 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **8 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Database` connect `Database` to `MainWindow`, `project_tab.py`, `RAB Parser & Test Suite`, `RabMappingDialog`, `RAB Import Dialogs`, `AllocationTab`, `RemainderDialog`, `Rencana Teknis — Universal Invoice Maker`?**
  _High betweenness centrality (0.223) - this node is a cross-community bridge._
- **Are the 20 inferred relationships involving `Database` (e.g. with `Item` and `ParsedGroup`) actually correct?**
  _`Database` has 20 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Alur Utama (Buat Proyek, Import, Validasi, Bagi, Review, Kunci, Export)`, `Kunci Termin (Lock Alokasi Sebelum Export)`, `Aplikasi Pembagian RAB ke Kendaraan untuk Invoice` to the rest of the system?**
  _10 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `MainWindow` be split into smaller, more focused modules?**
  _Cohesion score 0.06095791001451379 - nodes in this community are weakly interconnected._
- **Why does `Rencana Teknis — Universal Invoice Maker` connect `Rencana Teknis — Universal Invoice Maker` to `AllocationTab`, `Invoice Excel Export`?**
  _High betweenness centrality (0.070) - this node is a cross-community bridge._
- **Are the 14 inferred relationships involving `AllocationTab` (e.g. with `Database` and `AllocIssue`) actually correct?**
  _`AllocationTab` has 14 INFERRED edges - model-reasoned connections that need verification._
- **Should `project_tab.py` be split into smaller, more focused modules?**
  _Cohesion score 0.06161616161616162 - nodes in this community are weakly interconnected._
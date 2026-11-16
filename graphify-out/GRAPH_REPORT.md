# Graph Report - universal-invoice-maker  (2026-11-16)

## Corpus Check
- 24 files · ~16,191 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 4 file(s) not represented in the graph (top: .bat 2, (none) 1, .spec 1)

## Summary
- 425 nodes · 1075 edges · 14 communities (9 shown, 5 thin omitted)
- Extraction: 86% EXTRACTED · 14% INFERRED · 0% AMBIGUOUS · INFERRED: 150 edges (avg confidence: 0.9)
- Token cost: 5,500 input · 9,400 output

## Graph Freshness
- Built from commit: `6ceaf07d`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- Database
- rab_parser.py
- invoice_export.py
- main_window.py
- MainWindow
- Rencana Teknis — Universal Invoice Maker
- main.py
- RabMappingDialog
- Universal Invoice Maker
- detect_po_number
- make_samples.py

## God Nodes (most connected - your core abstractions)
1. `Database` - 42 edges
2. `parse_rab()` - 27 edges
3. `MainWindow` - 24 edges
4. `InvoiceTab` - 22 edges
5. `RabMappingDialog` - 21 edges
6. `Project` - 20 edges
7. `RabMapping` - 19 edges
8. `ProjectTab` - 19 edges
9. `build_invoice_netto()` - 17 edges
10. `InvoiceData` - 16 edges

## Surprising Connections (you probably didn't know these)
- `3.3 Model` --references--> `Item`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/models.py
- `5. Kontrak template invoice` --references--> `invoice_filename()`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/services/invoice_export.py
- `9. Risiko` --references--> `to_number()`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/services/rab_parser.py
- `3.2 Mapping` --references--> `RabMappingDialog`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → app/ui/dialogs_rab.py
- `Catatan blocker (diperbarui)` --references--> `write_multi_section_rab()`  [INFERRED]
  docs/PLAN-UNIVERSAL-INVOICE-MAKER.md → tests/test_core.py

## Import Cycles
- None detected.

## Communities (14 total, 5 thin omitted)

### Community 0 - "Database"
Cohesion: 0.06
Nodes (14): Database, Item, ParsedItem, ValidationIssue, build_invoice_netto(), _check_item(), validate_project(), date_to_iso() (+6 more)

### Community 1 - "rab_parser.py"
Cohesion: 0.06
Nodes (34): _base_header_text(), cell_text(), _classify_grouped(), classify_header(), detect_best_header(), detect_columns(), detect_grouped_columns(), _formula_without_cache() (+26 more)

### Community 2 - "invoice_export.py"
Cohesion: 0.06
Nodes (24): invoice_summary_text(), build_invoice_workbook(), _clone_block(), _clone_row_style(), _copy_style(), export_invoices(), invoice_filename(), InvoiceData (+16 more)

### Community 3 - "main_window.py"
Cohesion: 0.09
Nodes (13): Project, LogDialog, RabImportConfirmDialog, ProjectTab, Banner, button(), Card, heading() (+5 more)

### Community 4 - "MainWindow"
Cohesion: 0.07
Nodes (6): HelpDialog, SettingsDialog, InvoiceTab, MainWindow, NewProjectDialog, set_item()

### Community 5 - "Rencana Teknis — Universal Invoice Maker"
Cohesion: 0.06
Nodes (30): 10. Pertanyaan terbuka, 12. Hasil template check — `sample/Template_RAB_Multi_Termin.xlsx`, 1. Keputusan scope yang sudah dikunci, 3.2 Mapping, 3.3 Model, 3.4 Output, 3. Arsitektur target, 4.1 Tabel yang dipakai (+22 more)

### Community 6 - "main.py"
Cohesion: 0.14
Nodes (11): app_data_dir(), bundled_template_path(), db_path(), default_template_path(), log_path(), user_template_path(), app_icon(), 4.4 Keputusan yang harus diambil: lokasi data user (+3 more)

### Community 8 - "Universal Invoice Maker"
Cohesion: 0.17
Nodes (11): Alur Pemakaian, Asumsi & Batasan, Build EXE, Fitur, Format Excel RAB, Instalasi & Menjalankan, Lokasi Data, Pengujian (+3 more)

### Community 9 - "detect_po_number"
Cohesion: 0.31
Nodes (4): detect_po_number(), 11. Log status, Catatan blocker (diperbarui), PoDetectionTests

## Knowledge Gaps
- **5 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Database` connect `Database` to `main_window.py`, `MainWindow`, `main.py`?**
  _High betweenness centrality (0.148) - this node is a cross-community bridge._
- **Are the 11 inferred relationships involving `Database` (e.g. with `Item` and `ParsedItem`) actually correct?**
  _`Database` has 11 INFERRED edges - model-reasoned connections that need verification._
- **Should `Database` be split into smaller, more focused modules?**
  _Cohesion score 0.05895061728395062 - nodes in this community are weakly interconnected._
- **Why does `Rencana Teknis — Universal Invoice Maker` connect `Rencana Teknis — Universal Invoice Maker` to `detect_po_number`, `invoice_export.py`, `rab_parser.py`?**
  _High betweenness centrality (0.101) - this node is a cross-community bridge._
- **Are the 3 inferred relationships involving `parse_rab()` (e.g. with `.first()` and `.terms()`) actually correct?**
  _`parse_rab()` has 3 INFERRED edges - model-reasoned connections that need verification._
- **Should `rab_parser.py` be split into smaller, more focused modules?**
  _Cohesion score 0.06142410015649452 - nodes in this community are weakly interconnected._
- **Why does `build_invoice_netto()` connect `Database` to `detect_po_number`, `invoice_export.py`, `main_window.py`?**
  _High betweenness centrality (0.058) - this node is a cross-community bridge._
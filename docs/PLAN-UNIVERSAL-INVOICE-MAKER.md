# Rencana Teknis — Universal Invoice Maker

Status: **draft, menunggu konfirmasi**. Belum ada kode yang diubah dan folder belum dipindahkan.

Dokumen ini mencakup dua hal:
1. Perubahan scope dari *Invoice Automation* (alokasi termin per kendaraan) menjadi *Universal Invoice Maker* (Excel masuk → invoice keluar, 1:1).
2. Prosedur pemindahan project ke `C:\Users\Fadhil\Documents\Data Science\projects\universal-invoice-maker`.

---

## 1. Keputusan scope yang sudah dikunci

| # | Keputusan | Nilai |
|---|---|---|
| 1 | Transformasi data | **1:1** — satu baris RAB = satu baris invoice. Tidak ada agregasi, tidak ada pembagian qty. |
| 2 | Pembagian termin | **Dibuang dari scope.** Termin sepenuhnya urusan user di file Excel-nya. |
| 3 | Alokasi ke unit/kendaraan | **Dibuang dari scope.** |
| 4 | Template invoice | **Disediakan aplikasi** (belum user-supplied). |

Konsekuensi langsung: inti aplikasi berubah dari *calculator* menjadi **Excel mapper + template filler**.

---

## 2. Kondisi kode saat ini (terverifikasi)

### 2.1 Skema database — `app/db.py`

```
projects        id, name, client_name, client_address, po_prefix, vehicle_type,
                invoice_date, date_format, created_at
vehicle_types   id, project_id, name, unit_count, sort_order, po_prefix
vehicles        id, project_id, plate, vehicle_type, sort_order, vehicle_type_id, unit_no
items           id, project_id, row_no, name, unit, unit_price, total_qty, vehicle_type_id
terms           id, project_id, name, seq, invoice_date, date_format, discount_percent, locked
term_qty        id, term_id, item_id, qty
allocations     id, term_id, item_id, vehicle_id, qty
audit_log       id, project_id, ts, action, detail
settings        key, value
```

Pola migrasi sudah ada dan idempoten: `Database._migrate()` di `db.py:121` memakai `PRAGMA table_info` + `ALTER TABLE ADD COLUMN` + backfill. **Pola ini yang akan dipakai ulang**, bukan bikin migrasi baru dari nol.

### 2.2 Parser RAB — `app/services/rab_parser.py`

Sudah generik dan ini aset terbesar:

| Komponen | Lokasi | Catatan |
|---|---|---|
| Role kolom | L18–23 | `ignore`, `name`, `unit`, `price`, `total_qty`, `term` |
| `normalize_header()` | L136 | normalisasi teks header |
| `classify_header()` | L164 | teks header → role |
| `detect_grouped_columns()` | L246 | header dua tingkat (grup di atas kolom) |
| `detect_columns()` | L293 | deteksi kolom satu tingkat |
| `detect_best_header()` | L305 | pilih baris header terbaik |
| `to_number()` | L335 | angka Excel + format Indonesia (`1.000,50`) |
| `_merged_anchor_value()` | L195 | nilai sel di titik jangkar merge |
| `parse_rab()` | L411 | orkestrasi |
| `read_header_preview()` / `list_sheets()` | L621 / L611 | preview untuk dialog mapping |

Yang perlu dibuang hanya: `ROLE_TERM` (L23), `term_label_from_header()` (L154), `match_section_header()` (L396, pola `"NAMA JENIS (N Unit)"`), dan field `term_qty` pada `ParsedItem`.

### 2.3 Template invoice — `app/services/invoice_export.py`

Kontrak layout saat ini **hardcoded**, bukan data:

- Validasi: `_valid_template()` L119 → `A10 == "qty"` dan `B10 == "nama barang"`.
- Tulis header: `A1`, `C1`, `C2`, `B4`, `B5`, `D4`, `D5`, `D6` (plat), `D7` (jenis) — L190–198.
- Baris item mulai `11` — L202.
- Anchor style subtotal/diskon/total: baris `21`, `22`, `23` — L218, L225, L233.
- Fallback nama perusahaan literal `"CV. PANJI PERKASA"` — L191.
- Nama file: `Invoice {term} - {plate}.xlsx` — L598.

Semua ini harus pindah ke **mapping table** (data), bukan `if` di kode.

### 2.4 Ketergantungan pada kendaraan

`InvoiceData` punya field `plate` dan `vehicle_type` (L36–37). `export_invoices()` punya dua mode: `files` (satu file per kendaraan) dan `workbook` (satu sheet per kendaraan + sheet rekap + pivot) — L604–668. `build_rekap_sheet()` dan `build_pivot_sheets()` juga berasumsi "jenis kendaraan".

---

## 3. Arsitektur target

```
[1] INTAKE            [2] MAPPING           [3] MODEL            [4] OUTPUT
upload .xlsx    →   deteksi header &    →   items (1:1)     →   template invoice
(pilih sheet)       klasifikasi kolom       + group opsional      (placeholder)
                          ↓
                    dialog koreksi
                          ↓
                    template profile (disimpan, dipakai ulang)
```

### 3.1 Intake
- Input: satu file `.xlsx`, user pilih sheet (pakai `list_sheets()`).
- Output: grid mentah + `ParseResult` yang sudah ada.

### 3.2 Mapping
- `RabMappingDialog` (`app/ui/dialogs_rab.py:55`) dinaikkan menjadi **Template Profile**:
  - hasil koreksi user → `{fingerprint, mapping: {kolom: role}}`
  - `fingerprint` = hash dari (jumlah kolom, teks header baris terpilih, tipe data per kolom)
  - upload berikutnya dengan fingerprint sama → auto-apply, nol klik
- Tidak ada AI. Fallback selalu manual, bukan gagal.

### 3.3 Model
- Satu baris RAB → satu `Item`. Field minimal: `name`, `unit`, `unit_price`, `qty`.
- Kolom RAB yang tidak dipetakan → masuk `attributes` (JSON) agar tidak hilang.
- Group/section bersifat **opsional**: kalau file punya blok, dipakai untuk memecah invoice; kalau tidak, satu invoice untuk seluruh sheet.

### 3.4 Output
- `InvoiceLine` = `Item` apa adanya. `amount = qty × unit_price` (sudah ada di L29–31).
- Template diisi lewat mapping `cell → field`, bukan literal.

---

## 4. Perubahan data model (migrasi v2)

### 4.1 Tabel yang dipakai

```sql
projects          id, name, client_name, client_address, po_prefix,
                  invoice_date, date_format, created_at
                  -- DROP konsep: vehicle_type
document_groups   id, project_id, name, sort_order, meta_json        -- pengganti vehicle_types
items             id, project_id, group_id, row_no, name, unit,
                  unit_price, qty, attributes_json                   -- qty menggantikan total_qty
template_profiles id, name, fingerprint, mapping_json, created_at    -- BARU
settings          key, value
audit_log         id, project_id, ts, action, detail
```

### 4.2 Tabel yang dihapus

`terms`, `term_qty`, `allocations`, `vehicles`, `vehicle_types`.

### 4.3 Strategi migrasi

Mengikuti pola `_migrate()` yang sudah ada:

1. `ALTER TABLE items ADD COLUMN qty` lalu `UPDATE items SET qty = total_qty` (data lama tidak hilang).
2. Buat `document_groups` dan salin dari `vehicle_types`.
3. Buat `template_profiles`, isi satu default profile = layout template sekarang.
4. **Jangan `DROP` tabel lama di rilis yang sama.** Biarkan `terms`/`allocations` ada tapi tidak dibaca; drop di rilis berikutnya setelah data user terbukti aman.

### 4.4 Keputusan yang harus diambil: lokasi data user

`%APPDATA%\InvoiceAutomation\data.db` adalah milik aplikasi lama. Karena project pindah dan nama berubah:

| Opsi | Efek |
|---|---|
| A. Pertahankan nama folder data | User lama otomatis bawa project-nya; risiko: user baru dapat DB warisan yang sudah termigrasi. |
| B. Ganti ke `%APPDATA%\UniversalInvoiceMaker` | Bersih, tapi project lama user "hilang" sampai di-import manual. |

**Rekomendasi: B**, karena aplikasi ini pada dasarnya produk baru. Perlu satu baris di `app/config.py` (`app_data_dir()`).

---

## 5. Kontrak template invoice

Template tetap milik aplikasi, tapi **layout dinyatakan sebagai data**, bukan kode:

```json
{
  "name": "Default Invoice",
  "sheet": 0,
  "header_row": 10,
  "item_first_row": 11,
  "style_anchors": {"item": 11, "subtotal": 21, "discount": 22, "total": 23},
  "fields": {
    "A1": "title",
    "C1": "company_name",
    "B4": "po_number",
    "B5": "invoice_date",
    "D4": "client_name",
    "D5": "client_address",
    "D6": "label"
  },
  "columns": {"qty": 1, "name": 2, "unit_price": 3, "amount": 4}
}
```

Perubahan yang dibutuhkan di `invoice_export.py`:

- `_valid_template()` (L119) → cari baris header yang mengandung `Qty` + `Nama Barang`, jangan hardcode baris 10.
- Semua penulisan sel L190–198 → loop dari `fields`.
- Anchor style L218–233 → dari `style_anchors`.
- `"CV. PANJI PERKASA"` L191 → dari `settings`.
- `invoice_filename()` L598 → pola nama file dari profile (`{po_number} - {label}.xlsx`).

Catatan: `build_rekap_sheet()` dan `build_pivot_sheets()` **tidak** masuk MVP. Keduanya bergantung pada konsep "jenis kendaraan" dan baru relevan kalau group ada isinya.

---

## 6. Daftar file: hapus / ubah / baru

| Aksi | File | Baris | Alasan |
|---|---|---|---|
| **Hapus** | `app/ui/allocation_tab.py` | 777 | layar alokasi tidak dipakai |
| **Hapus** | `app/ui/dialogs_alloc.py` | 563 | dialog remainder/term manager |
| **Hapus** | `app/services/allocation.py` | 110 | logika split qty |
| **Hapus** | `app/services/validation.py` | 138 | invarian termin |
| **Hapus** | `tests/` bagian alokasi | — | test untuk kode yang dibuang |
| **Ubah** | `app/db.py` | 1.076 | schema v2, migrasi, buang method term/vehicle |
| **Ubah** | `app/models.py` | 195 | hapus `Term`, `Vehicle`, `VehicleType`; `Item.total_qty` → `qty` |
| **Ubah** | `app/services/rab_parser.py` | 642 | buang `ROLE_TERM`, `term_qty`, `match_section_header` |
| **Ubah** | `app/services/invoice_export.py` | 668 | mapping-driven, hapus asumsi plat/jenis |
| **Ubah** | `app/ui/dialogs_rab.py` | 737 | dialog jadi profile mapper; hapus `TermQtyDialog` |
| **Ubah** | `app/ui/project_tab.py` | 759 | buang kolom plat/jenis, tampilkan atribut generik |
| **Ubah** | `app/ui/export_tab.py` | 527 | buang pilihan termin |
| **Ubah** | `app/ui/main_window.py` | 367 | tab alokasi dihapus |
| **Baru** | `app/services/template_profile.py` | ~150 | simpan/ambil mapping + fingerprint |
| **Baru** | `services` template filler | ~120 | `fill_template(profile, invoice)` |

Perkiraan bersih: **−1.590 baris dihapus, +270 baris baru**, sisanya penyuntingan.

---

## 7. Rencana kerja & acceptance criteria

### Fase 0 — Characterization test (sebelum menyentuh apa pun)
- Ambil 2–3 file RAB nyata dengan format berbeda.
- Test: jumlah baris, nilai `qty`/`unit_price`, nama item harus identik sebelum dan sesudah refactor.
- **Gate:** semua test hijau di kode lama.

### Fase 1 — Skema v2 + migrasi
- **AC-1:** DB lama terbuka tanpa error, `items.qty` terisi dari `total_qty`.
- **AC-2:** `python -m unittest tests.test_core` tetap hijau (test alokasi dihapus bersamaan dengan kodenya).

### Fase 2 — Mapping 1:1
- **AC-3:** Upload file → tabel tampil 1:1; `jumlah baris invoice == jumlah baris RAB`.
- **AC-4:** `sum(qty invoice) == sum(qty RAB)` per sheet.
- **AC-5:** Urutan baris dipertahankan.

### Fase 3 — Template profile
- **AC-6:** Upload file yang sama dua kali → percobaan kedua tanpa dialog (auto-apply).
- **AC-7:** Upload format berbeda → dialog muncul dan bisa disimpan sebagai profile baru.

### Fase 4 — Output
- **AC-8:** Invoice keluar dengan style template utuh (border, tinggi baris, merge header).
- **AC-9:** Tidak ada string `"CV. PANJI PERKASA"` di kode; nilainya dari settings.

### Fase 5 — Bersih-bersih
- Hapus tab/dialog/servis mati; pastikan tidak ada import yang menunjuk ke modul terhapus.
- Baru setelah ini: `graphify update .` supaya graph menggambarkan kode baru.

---

## 8. Prosedur pindah folder

Target: `C:\Users\Fadhil\Documents\Data Science\projects\universal-invoice-maker`

### 8.1 Pre-flight
1. `git status` — semua perubahan di-commit atau di-stash (supaya rollback bersih).
2. Tutup aplikasi yang sedang jalan (SQLite file lock).
3. Pastikan folder tujuan belum ada.

### 8.2 Pindah
4. `Move-Item` folder lama → folder baru (satu operasi, bukan copy lalu hapus, agar `.git` dan `graphify-out/` utuh).

### 8.3 Yang ikut otomatis dan valid
| Aset | Status setelah pindah |
|---|---|
| `.git` | ikut, history utuh |
| `graphify-out/graph.json` | **valid** — `source_file` dan node ID relatif terhadap root |
| `graphify-out/GRAPH_REPORT.md`, `graph.html` | valid |
| `graphify-out/manifest.json` | valid — key-nya juga relatif |
| `graphify-out/.graphify_python` | valid — menunjuk interpreter uv global |
| `AGENTS.md` | tidak ikut tapi tetap berlaku (letaknya di `C:\Users\Fadhil\AGENTS.md`) |

### 8.4 Yang wajib di-refresh
5. **`graphify-out/.graphify_root`** — masih berisi path absolut lama. Tulis ulang ke path baru.
6. Verifikasi: `graphify query "what is the main entry point"` → harus jalan tanpa error.

### 8.5 Session
7. Pindahkan sesi obrolan ini ke folder baru (`session_move`), supaya konteks diskusi ikut.

### 8.6 Yang tertinggal (sengaja)
- `%APPDATA%\InvoiceAutomation\` — `data.db`, `error.log`, template user. Data runtime, bukan bagian repo.
- Kalau opsi 4.4-B dipilih, folder ini tidak akan dibaca aplikasi baru sama sekali.

### 8.7 Verifikasi pasca-pindah
8. `python -m unittest tests.test_core`
9. `python main.py` → aplikasi terbuka.
10. `git status` → bersih (kecuali `graphify-out/` yang belum di-gitignore).

### 8.8 Rollback
Pindahkan folder kembali ke path lama, tulis ulang `.graphify_root`, pindahkan sesi kembali. Tidak ada state lain yang berubah.

---

## 9. Risiko

| Risiko | Dampak | Mitigasi |
|---|---|---|
| Variasi header Excel antar file | mapping gagal senyap | AC-3/AC-6 + fallback dialog manual |
| Angka format Indonesia (`1.000,50`) | qty/harga salah | `to_number()` sudah menangani — wajib ada test regresi |
| Merge cell salah baca | baris item hilang/duplikat | test Fase 0 pakai file nyata yang punya merge |
| Migrasi DB merusak data user | kehilangan project lama | `ALTER` + backfill, tidak ada `DROP` di rilis pertama |
| Graph usang setelah refactor | agent salah kira kode lama masih ada | `graphify update .` **setelah** Fase 5, bukan sebelumnya |

## 10. Pertanyaan terbuka

1. **Nama aplikasi & folder data**: `UniversalInvoiceMaker`? (mempengaruhi `app/config.py` dan README)
2. **Group/section**: kalau RAB punya blok, mau dipecah jadi beberapa invoice, atau tetap satu invoice per file?
3. **Rekap & pivot sheet**: masih dibutuhkan, atau cukup invoice per file?
4. **Nomor PO**: masih `prefix + nomor urut`, atau diambil langsung dari file RAB?

---

## 11. Log status

| Tanggal | Item | Status |
|---|---|---|
| 2026-11-15 | Pindah folder ke `universal-invoice-maker` | selesai (99/99 file, 0 selisih) |
| 2026-11-15 | `.graphify_root` di-refresh, graph valid | selesai |
| 2026-11-15 | Sesi obrolan dipindah | selesai |
| 2026-11-15 | `.venv` + `requirements.txt` terpasang | selesai (PySide6 6.12.0, openpyxl 3.1.5) |
| 2026-11-15 | Baseline test | 13 ok, **3 error** (butuh `sample/rab_sample_real.xlsx`) |
| 2026-11-15 | `app/services/template_profile.py` | selesai (baru) |
| 2026-11-15 | `invoice_export.py` layout jadi data | selesai, 13 test tetap hijau |
| 2026-11-15 | `app/config.py` identitas + folder data baru | selesai |
| 2026-11-15 | **Tab Invoice 1:1** (`app/ui/invoice_tab.py`) | selesai — preview + export 1:1 |
| 2026-11-15 | **builder netto** `build_invoice_netto()` | selesai — 1:1 tanpa termin/vehicle |
| 2026-11-15 | **Test self-contained** (3 test bergantung file user) | diganti fixture in-memory, **28 test OK** |
| 2026-11-15 | **Template profile test** | 4 test baru, semua OK |
| 2026-11-15 | **App startup** `python main.py` | selesai — exit 0, window title OK |
| 2026-11-15 | **E2E UI** (import RAB → export 1:1) | selesai — 7/7 baris pas, file terbuka |
| 2026-11-16 | **Deteksi No PO dari RAB** (`detect_po_number`) | selesai — dipakai sebagai default No PO |
| 2026-11-16 | Test suite | **32 test OK**, tanpa file user |
| 2026-11-16 | **Bersih-bersih penuh** — hapus semua jalur kendaraan/termin/alokasi | selesai |
| 2026-11-16 | File dihapus: `allocation.py`, `vehicle_parser.py`, `allocation_tab.py`, `export_tab.py`, `dialogs_alloc.py`, PRD lama, `error.log` | selesai |
| 2026-11-16 | `app/db.py` skema minimal (`projects`, `items`, `audit_log`, `settings`) | selesai |
| 2026-11-16 | `app/models.py`, `validation.py`, `invoice_builder.py`, `project_tab.py`, `dialogs_rab.py` ditulis ulang | selesai |
| 2026-11-16 | `README.md` + `build_exe.bat` + spec diganti nama ke Universal Invoice Maker | selesai |
| 2026-11-16 | Test suite setelah bersih-bersih | **30 test OK** |
| 2026-11-16 | Smoke test offscreen + E2E import RAB → export | selesai — 12 baris, total cocok |
| — | Skema v2 + migrasi | **belum** |

### Catatan blocker (diperbarui)

`sample/rab_sample_real.xlsx` tidak di-commit (gitignored). Tiga test lama
gagal → diganti **fixture self-contained** di `tests/test_core.py`
(`write_multi_section_rab`). Test suite sekarang **30 test OK** tanpa file user.

Pembersihan 2026-11-16 menghapus seluruh sisa proyek lama: tidak ada lagi
tabel/kode kendaraan, termin, maupun alokasi. Database lama di
`%APPDATA%\InvoiceAutomation` tidak dimigrasi; aplikasi memakai folder baru
`%APPDATA%\UniversalInvoiceMaker`.

---

## 12. Hasil template check — `sample/Template_RAB_Multi_Termin.xlsx`

Isi file: **3 layout RAB input**, bukan template invoice. Perbedaannya:

| Sheet | Header row | Termin | Kolom | Blok termin |
|---|---|---|---|---|
| `2_Termin_Header_Baris_1` | 1–2 | 2 | 17 | H–K, L–O |
| `4_Termin_Header_Baris_2` | 2–3 | 4 | 25 | H–K, L–O, P–S, T–W |
| `Persentase_Header_Baris_3` | 3–4 | 3 | 15 | BOBOT % + NILAI TAGIHAN |

Sebelum perbaikan, **ketiganya gagal** dengan
`RabParseError: Kolom Nama Barang dan Harga Satuan wajib dipilih.`

Dua akar masalah:

1. `UNIT_KEYS` tidak mengenali `SAT.` → `detect_grouped_columns()` langsung
   `return None` pada baris `if not any(text == "satuan" ...)`, sehingga deteksi
   header dua tingkat tidak pernah aktif.
2. Precondition grup hanya memeriksa baris bawah. Di template ini kolom satuan
   dan sebagian header ada di baris atas (merge vertikal `B1:B2`, `C1:C2`,
   `C3:C4`), jadi harus diperiksa di kedua baris.

Perbaikan (2 titik di `rab_parser.py`): `UNIT_KEYS` + `"sat."`, dan precondition
memakai `base_texts + parents`.

Hasil setelah perbaikan:

| Sheet | header row | item | grup |
|---|---|---|---|
| `2_Termin_Header_Baris_1` | 2 | 7 | 3 (`Sub-Pekerjaan A.1/A.2/B.1`) |
| `4_Termin_Header_Baris_2` | 3 | 7 | 3 |
| `Persentase_Header_Baris_3` | 4 | 4 | 0 |

Baris `Jumlah …`, `TOTAL PEKERJAAN UTAMA`, `REKAPITULASI`, `PPN 11%`,
`GRAND TOTAL`, dan `BOBOT PROGRESS` benar tidak menjadi item.

**Semantik qty untuk 1:1:** qty diambil dari blok `SESUAI PELAKSANAAN` →
`VOLUME`. Kolom termin diabaikan sepenuhnya — konsisten dengan keputusan
"termin urusan user".

Regresi terkunci di `tests/test_core.py::MultiTerminTemplateTests` (4 test).



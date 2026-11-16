# Universal Invoice Maker

Aplikasi desktop Windows untuk memetakan isi RAB Excel menjadi invoice **satu
banding satu**: satu baris RAB menjadi tepat satu baris invoice. Tidak ada
pengelompokan kendaraan, pembagian termin, atau alokasi qty.

Aplikasi berjalan **offline**, data disimpan lokal di SQLite, dan tidak
memerlukan koneksi internet. Dibangun dengan Python, PySide6, openpyxl, dan
SQLite.

## Fitur

- **Proyek lokal**: nama proyek, klien, alamat klien, prefix No PO, tanggal
  invoice (format lengkap atau bulan–tahun).
- **Import RAB Excel** dengan deteksi header otomatis dan dialog koreksi
  pemetaan kolom.
- **Pemetaan 1:1**: qty diambil dari volume RAB apa adanya, harga dari harga
  satuan. Tidak ada agregasi maupun pemecahan qty.
- **Invoice netto**: PPN tidak dihitung aplikasi. Harga dianggap harga akhir;
  pajak diurus pengguna.
- **No PO otomatis**: diambil dari RAB bila terdeteksi di area atas tabel;
  jika tidak ada, dipakai `prefix + label`.
- **Validasi ringan** sebelum export: item kosong, qty ≤ 0, harga negatif.
- **Export Excel** mengikuti layout `invoice_template.xlsx`.
- **Riwayat aktivitas** (import, edit proyek, export).
- **Template invoice** dapat diganti lewat menu **Pengaturan**.

## Instalasi & Menjalankan

Prasyarat: Python 3.10+ (diuji dengan Python 3.13).

```bat
pip install -r requirements.txt
python main.py
```

Atau klik dua kali `run.bat`. Script itu akan membuat virtual environment
`.venv`, memasang dependensi, lalu menjalankan aplikasi tanpa jendela konsol.

## Alur Pemakaian

1. **Buat proyek** dari sidebar, lalu lengkapi informasi proyek di tab
   **1. Data Proyek**. Tanggal invoice punya dua opsi format: **tanggal
   lengkap** (`1 Juli 2026`) atau **bulan dan tahun** (`Juli 2026`).
2. **Import RAB** (tombol `Import RAB`):
   - pilih file `.xlsx`;
   - periksa sheet, baris header, dan peran kolom di dialog pemetaan;
   - klik `Lanjut`, periksa ringkasan, lalu `Import RAB`.
3. Pratinjau item RAB tampil di tab **1. Data Proyek**. Setiap baris akan
   menjadi satu baris invoice.
4. Buka tab **2. Invoice 1:1**, periksa tabel, isi **Nomor / Label** dan
   **No PO** (kosongkan untuk otomatis), pilih folder output, lalu klik
   **Export Invoice**.
5. Satu file Excel dibuat untuk proyek tersebut. Nilainya netto.

## Format Excel RAB

Format tidak dikunci mati. Aplikasi mendeteksi header, dan jika meleset kamu
bisa mengoreksi peran tiap kolom.

| No | Uraian Pekerjaan | Satuan | Harga Satuan | Jumlah Biaya | Termin 1 | Termin 2 |
|----|------------------|--------|--------------|--------------|----------|----------|
| A  | YAMAHA VIXION (3 Unit) | | | | | |
| 1  | Service | UNIT | 240000 | 1440000 | 3 | 3 |
| 2  | Oli Mesin | LITER | 100000 | 600000 | 3 | 3 |
| B  | SUZUKI HAYATE (2 Unit) | | | | | |
| 1  | Service | UNIT | 200000 | 800000 | 2 | 2 |

Ketentuan:

- Kolom yang dikenali: **Nama Barang**, **Satuan** (opsional), **Harga
  Satuan**, **Qty Total**, dan kolom **Termin**.
- Kolom termin **dikenali tetapi diabaikan**. Qty invoice memakai volume total
  pada blok `SESUAI PELAKSANAAN` / kolom Qty Total.
- Baris penanda blok seperti `NAMA JENIS (N Unit)` otomatis dilewati, bukan
  dijadikan item.
- Baris subtotal, total, PPN, rekap, dan catatan otomatis dilewati.
- Jika tidak ada kolom Qty Total, qty dihitung dari jumlah kolom termin.
- Qty desimal akan memunculkan peringatan dan dibulatkan ke bilangan bulat
  terdekat (0,5 ke atas) setelah kamu menyetujuinya.
- Jika sel berisi formula yang belum punya nilai tersimpan, buka file di Excel
  lalu simpan sekali agar nilainya terbaca.

File contoh: `sample/rab_sample.xlsx` (tidak ikut repo, ada di `.gitignore`).

## Template Invoice

Template default ada di `app/resources/invoice_template.xlsx` dan disalin ke
folder data pengguna saat pertama kali dibutuhkan.

- Baris header tabel item dideteksi otomatis dari label `Qty` dan
  `Nama Barang`; layout tidak dikunci ke baris tertentu.
- Baris 1 sampai header: judul, perusahaan, No PO, tanggal, klien, alamat,
  label proyek.
- Setelah item: `Sub Total` dan `TOTAL`. Nilai `TOTAL` sama dengan `Sub Total`
  karena invoice netto (tanpa diskon/PPN).
- Footer: `Hormat Kami,` dan nama perusahaan.
- Jumlah baris item, total, dan footer digeser otomatis.

Field yang diisi aplikasi (template bawaan):

| Sel | Isi |
|-----|-----|
| `C1` | Nama perusahaan (Pengaturan) |
| `C2` | Alamat perusahaan (Pengaturan) |
| `B4` | No PO (dari RAB, atau prefix + label) |
| `B5` | Tanggal invoice |
| `D4` | Kepada Yth / nama klien |
| `D5` | Alamat klien |
| `D6` | Nomor / Label invoice |
| `D7` | Nama proyek |

Syarat minimal template: memuat header `Qty` dan `Nama Barang` pada baris yang
sama. Untuk mengganti layout, ganti template lewat **Berkas → Pengaturan**.

## Build EXE

```bat
build_exe.bat
```

Hasil ada di `dist\UniversalInvoiceMaker.exe` (one-file). Template ikut
dipaketkan. Build membutuhkan PyInstaller; script akan memasangnya jika belum
ada.

## Lokasi Data

- Database: `%APPDATA%\UniversalInvoiceMaker\data.db`
- Template pengguna: `%APPDATA%\UniversalInvoiceMaker\invoice_template.xlsx`
- Log error: `%APPDATA%\UniversalInvoiceMaker\error.log`

Folder data bisa dioverride dengan environment variable
`UNIVERSAL_INVOICE_APP_DATA`.

## Pengujian

```bat
python -m unittest tests.test_core -v
```

Tes mencakup parsing RAB, deteksi No PO, validasi proyek, pembuatan data
invoice netto, layout template, dan export Excel.

## Asumsi & Batasan

- Satu pengguna, Windows, offline. Tidak ada sinkronisasi cloud/multi-user.
- Satu proyek = satu RAB = satu invoice. Pemecahan per bagian/batch belum ada.
- Satu baris RAB selalu menjadi satu baris invoice; qty tidak dipecah.
- Kolom termin diabaikan sepenuhnya.
- Nilai invoice netto; PPN/diskon tidak dihitung aplikasi.
- Qty desimal dibulatkan setelah konfirmasi.
- Output hanya Excel; tidak ada PDF/Word.
- Template invoice menentukan tampilan akhir.

## Troubleshooting

- **Qty terbaca 0 / item tidak muncul** — periksa baris header dan pemetaan
  kolom di dialog import. Pastikan sel berisi angka, bukan formula tanpa nilai.
- **Tombol export disabled** — belum ada item RAB, atau template tidak
  ditemukan. Lihat banner di tab terkait.
- **File invoice tertimpa** — file baru akan diberi akhiran `(2)`, `(3)`, dst.
- **Template rusak** — buka **Berkas → Pengaturan → Reset** untuk memulihkan
  template bawaan.
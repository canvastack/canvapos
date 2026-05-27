# Dokumentasi CanvaPOS

## Daftar Isi

1. [Arsitektur](#1-arsitektur)
2. [Struktur Data](#2-struktur-data)
3. [Fungsi Global & Konstanta](#3-fungsi-global--konstanta)
4. [Setup System](#4-setup-system)
5. [Sheet Reference](#5-sheet-reference)
6. [Macro Reference](#6-macro-reference)
7. [Trigger System](#7-trigger-system)
8. [Dialog & HTML Service](#8-dialog--html-service)
9. [Engine: BOM & HPP](#9-engine-bom--hpp)
10. [Panduan Pengguna](#10-panduan-pengguna)
11. [Pengembangan](#11-pengembangan)

---

## 1. Arsitektur

### 1.1 Gambaran Umum

CanvaPOS adalah **single-file Google Apps Script** (~2.600 baris) yang berjalan di Google Sheets. Tidak ada dependency eksternal, tidak ada build step, tidak ada package manager.

### 1.2 Alur Data

```
[POS Input] ──simpanTransaksi()──→ [Transaksi Log]
       │                                    │
       │ stockEngineBOM()                   │
       ▼                                    ▼
   [Stock] ←──sync── [Pengeluaran]    [Pendapatan]
       │                                    ▲
       │                                    │
       ▼                              refreshLaporan()
   [Bahan] ──VLOOKUP──→ [Resep (BOM)] ──────┘
```

### 1.3 Pola Desain

- **Sheet-per-module**: Setiap modul bisnis memiliki sheet sendiri
- **Procedural**: Semua fungsi global, tidak ada class
- **Singleton Caching**: `CACHED_VARIAN_LIST`, `CACHED_TOPPING_LIST`
- **Event-Driven**: Triggers `onEdit`, `onSelectionChange`, `onOpen`
- **Dependency Injection via `ss`**: Semua fungsi setup menerima `Spreadsheet` object

### 1.4 Konvensi Kode

- Semua global menggunakan `var` (bukan `const`/`let`) — lihat baris 4
- Warna diakses via getter `getC()` — selalu terdefinisi saat dipanggil
- Nama fungsi: camelCase dengan prefix untuk helper (`_`), setup (`setup_`), build (`build`)
- Bahasa: Indonesia (variabel, komentar, UI)
- Timezone: `Asia/Jakarta`

---

## 2. Struktur Data

### 2.1 Sheet Bahan (Master Bahan Baku)

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Kategori | String | Kategori: Bahan Utama, Topping, Bahan Pendukung, Kemasan |
| B | Nama Bahan | String | Nama unik bahan baku |
| C | Ukuran/Pack | Number | Jumlah satuan per pack |
| D | Satuan | String | Gram, Liter, Kg, Pack, Piece |
| E | Harga Beli | Number (Rp) | Harga per pack |
| F | Harga Per Piece | Formula | `=E/D` — Harga per satuan terkecil |

### 2.2 Sheet Resep (Bill of Materials)

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Nama Menu / Varian | String | Nama produk (key relasi ke POS) |
| B | Nama Bahan Baku | String | Nama bahan (foreign key ke Bahan) |
| C | Takaran (Qty Usage) | Number | Quantity per 1 unit produk |
| D | Satuan Penggunaan | String | Gram, Liter, Piece, dll |
| E | Harga Per Piece | Formula | VLOOKUP dari Bahan!F |
| F | Harga Per Takaran | Formula | `=C*E` |

### 2.3 Sheet Stock

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Kategori | String | Kategori bahan |
| B | Nama Bahan | String | Nama bahan (key) |
| C | Satuan | String | Satuan stok |
| D | Stok Awal | Number | Stok awal periode |
| E | Terjual | Number | Terpotong oleh transaksi (dari stockEngineBOM) |
| F | Sisa Stok | Number | Dikirim manual atau dari Pengeluaran |
| G | Min. Stok | Number | Batas minimal untuk restock alert |
| H | Status | Formula | `⚠ RESTOCK` atau `✓ OK` |

### 2.4 Sheet Transaksi

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | No. Trx | String | Format `TRX-001` |
| B | No. Item | Number | Urutan item dalam satu transaksi |
| C | Tanggal | Date/String | Format DD/MM/YYYY |
| D | Jam | String | Format HH:mm |
| E | Kasir | String | Nama kasir |
| F | Product Variant | String | Nama produk |
| G | Jumlah Cup | Number | QTY |
| H | Topping | String | Dipisah koma: "Keju, Boba" |
| I | Jml Topping | Number | Hitung otomatis dari kolom H |
| J | Harga Base | Number (Rp) | `=G*5000` |
| K | Harga Topping | Number (Rp) | `=I*G*1000` |
| L | Total | Number (Rp) | `=J+K` |

### 2.5 Sheet Pendapatan

**Ringkasan Hari Ini (baris 5-11):**

| Baris | Label | Sumber |
|---|---|---|
| 5 | Tanggal | `TODAY()` |
| 6 | Total Transaksi | `COUNTIF` |
| 7 | Total Cup Terjual | `SUMIF` |
| 8 | Total Pendapatan | `SUMIF` |
| 9 | Total HPP | HPP per cup + HPP per topping |
| 10 | Laba Bersih | `=Pendapatan-HPP` |
| 11 | Margin (%) | `=Laba/Pendapatan*100` |

**Status Kas (baris 12-13):** PC dan UB — LOOKUP dari sheet Kas

**Rekap Harian (baris 17+):** Dinamis dari Transaksi via `refreshLaporan()`

**Rekap Bulanan:** Agregasi dari rekap harian via `refreshLaporan()`

### 2.6 Sheet POS

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | No | Number (formula) | Nomor urut otomatis |
| B | Product Variant | Dropdown | Dari getVarianList() |
| C | Jumlah Cup | Number | Default 1 |
| D | Topping - Jenis | String | Manual atau dari pilihTopping() |
| E | Topping - Jumlah | Formula | Hitung dari koma di D |
| F | Topping - Harga Top | Formula | `=E*C*1000` |
| G | Harga Base | Formula | `=C*5000` |
| H | Total | Formula | `=G+F` |

**Grand Total:** Baris terakhir dengan formula `=SUM(H7:H{n})`

### 2.7 Sheet Pengeluaran

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Tanggal | Date | DD/MM/YYYY |
| B | Kategori | Dropdown | Bahan Utama, Topping, Kemasan, dll |
| C | Nama Item | Dropdown | Dinamis berdasarkan kategori |
| D | Satuan | Auto-fill | Dari Bahan |
| E | Jumlah | Number | QTY |
| F | Harga Satuan | Number (Rp) | Dari Bahan |
| G | Total | Formula | `=E*F` |
| H | Status Stok | Auto-fill | Sisa stok atau `✓ Synced` |

### 2.8 Sheet Kas

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Tanggal | Date | DD/MM/YYYY |
| B | Kategori | String | PC atau UB |
| C | Jenis | String | Saldo Awal, Top Up, Pengeluaran, Setor |
| D | Keterangan | String | Deskripsi transaksi |
| E | Jumlah | Number (Rp) | Nominal |
| F | Saldo | Number (Rp) | Saldo setelah transaksi |

---

## 3. Fungsi Global & Konstanta

### 3.1 Konstanta

| Variabel | Nilai | Deskripsi |
|---|---|---|
| `HARGA_BASE` | 5000 | Harga per cup |
| `HARGA_TOPPING` | 1000 | Harga per jenis topping per cup |
| `HPP_PER_CUP` | 2200 | Estimasi HPP per cup (fallback) |
| `HPP_PER_TOP` | 80 | Estimasi HPP per topping (fallback) |
| `POS_START_ROW` | 7 | Baris pertama data order di POS |
| `POS_INIT_ROWS` | 5 | Jumlah baris order awal |

### 3.2 Caching System

| Fungsi | Deskripsi |
|---|---|
| `getVarianList()` | Ambil daftar menu unik dari Resep, cache di `CACHED_VARIAN_LIST` |
| `getToppingList()` | Ambil daftar topping dari Bahan, cache di `CACHED_TOPPING_LIST` |
| `getDynamicList(kategori)` | Filter bahan by kategori |
| `clearDynamicCache()` | Reset cache (dipanggil saat Bahan diedit) |

### 3.3 Styling Helpers

| Fungsi | Parameter | Deskripsi |
|---|---|---|
| `getC()` | — | Return object warna (9 warna) |
| `setSheetFont(sh)` | Sheet | Set font Nunito ke seluruh sheet |
| `styleHeader(range, bg, fg?)` | Range, warna bg, warna font | Style header row dengan border |
| `styleData(range, bg?)` | Range, warna bg | Style data row dengan border |
| `styleOrderRow(sh, r, no, bg)` | Sheet, row, nomor, warna bg | Style 1 baris order POS + dropdown + formula |
| `setCurrency(range)` | Range | Format `"Rp "#,##0` |
| `deleteSheetIfExists(ss, name)` | Spreadsheet, string | Hapus sheet jika ada |

### 3.4 Color Palette

| Key | Hex | Penggunaan |
|---|---|---|
| `BLUE` | #2E86AB | Bahan, header |
| `GREEN` | #27AE60 | Stock, tombol Simpan, laba positif |
| `ORANGE` | #E67E22 | POS, header, Panduan |
| `RED` | #E74C3C | Pendapatan, tombol Clear, laba negatif |
| `PURPLE` | #8E44AD | Resep, Transaksi |
| `YELLOW` | #F9CA24 | Kas |
| `LIGHT` | #EAF4FB | Background zebra genap |
| `WHITE` | #FFFFFF | Background zebra ganjil |
| `DARK` | #2C3E50 | Teks gelap, header dark |
| `LGREEN` | #D5F5E3 | Background laba positif |
| `LRED` | #FADBD8 | Background laba negatif |
| `INPUT` | #FEF9E7 | Background baris input POS |

---

## 4. Setup System

### 4.1 `setupPOS()`

Fungsi utama yang membangun seluruh workbook:

1. Set timezone `Asia/Jakarta`
2. Hapus semua sheet existing
3. Build sheet berurutan (dengan `SpreadsheetApp.flush()` tiap step):
   `Bahan → Resep → Transaksi → Stock → Pengeluaran → Kas → Pendapatan → POS → Panduan`
4. Reorder sheet ke urutan yang diinginkan
5. Set active sheet ke POS
6. Tampilkan alert sukses

**Dependency order penting:**
- Pengeluaran & Kas harus SEBELUM Pendapatan (formula mereferensi sheet tersebut)
- Bahan harus SEBELUM Resep (VLOOKUP)
- Stock & Transaksi harus SEBELUM Pendapatan (SUMIF/COUNTIF)

### 4.2 Step-by-Step Functions

Jika `setupPOS()` timeout (6 menit batas Apps Script), jalankan:

| Fungsi | Output |
|---|---|
| `setup_1_Bahan()` | Sheet Bahan |
| `setup_2_Resep()` | Sheet Resep |
| `setup_3_Transaksi()` | Sheet Transaksi |
| `setup_4_Stock()` | Sheet Stock |
| `setup_5_Pengeluaran()` | Sheet Pengeluaran |
| `setup_5b_Kas()` | Sheet Kas |
| `setup_5c_Pendapatan()` | Sheet Pendapatan |
| `setup_6_POS()` | Sheet POS |
| `setup_7_Panduan()` | Sheet Panduan |
| `setup_8_Reorder()` | Urutkan sheet |

---

## 5. Sheet Reference

### 5.1 Bahan (Build: `buildBahan`)

- **Tab color:** Blue (#2E86AB)
- **Frozen rows:** 1
- **Data:** 37 baris bahan baku (baris 2-38)
- **Column widths:** 100, 200, 100, 80, 120, 140

### 5.2 Resep (Build: `buildResep`)

- **Tab color:** Purple (#8E44AD)
- **Frozen rows:** 1
- **Struktur:** Many-to-many (1 menu → banyak bahan)
- **Data:** ~170 baris (18 varian Pop Ice × 8 bahan + 2 kopi × 6 + teh × 6 + 6 topping)
- **Formulas:** Kolom E (VLOOKUP Bahan!B:F), Kolom F (C×E)
- **Column widths:** 220, 200, 140, 130, 140, 140

### 5.3 Stock (Build: `buildStock`)

- **Tab color:** Green (#27AE60)
- **Frozen rows:** 1
- **Data:** Dinamis, diisi via Pengeluaran sync
- **Conditional formatting:** 2 rules (RESTOCK = merah, OK = hijau)
- **Note:** Data array kosong — harus diisi via Pengeluaran atau manual

### 5.4 Transaksi (Build: `buildTransaksi`)

- **Tab color:** Purple (#8E44AD)
- **Frozen rows:** 2 (judul + header)
- **Title row:** Merged A1:L1
- **Data:** Mulai baris 3, diisi oleh `simpanTransaksi()`
- **Sample data:** Dikomentari (tidak aktif)

### 5.5 Pendapatan (Build: `buildPendapatan`)

- **Tab color:** Red (#E74C3C)
- **Frozen rows:** 1
- **Sections:**
  - Baris 1: Judul
  - Baris 2: Petunjuk refresh
  - Baris 4-11: Ringkasan Hari Ini (formula)
  - Baris 12-13: Status Kas (LOOkUP)
  - Baris 15-16: Header Rekap Harian
  - Baris 17+: Data rekap harian (dinamis)
  - Setelah rekap harian + 2 baris: Header Rekap Bulanan + data

### 5.6 POS (Build: `buildPOS`)

- **Tab color:** Orange (#E67E22)
- **Frozen rows:** 6 (header bertingkat)
- **Layout:**
  - Baris 1: Judul "🧋 CanvaPOS"
  - Baris 2: Kasir, Tanggal (format panjang Indonesia), Jam
  - Baris 3: Tombol aksi (Simpan, Add Row, Clear)
  - Baris 4: Info harga
  - Baris 5-6: Header kolom bertingkat (Topping: 3 sub-kolom)
  - Baris 7-11: 5 baris order awal
  - Baris 12: Grand Total

### 5.7 Pengeluaran (Build: `buildPengeluaran`)

- **Tab color:** Orange (#E67E22)
- **Frozen rows:** 3
- **Dropdown kategori:** 200 baris
- **Formula total:** 200 baris (E×F)
- **Data validation:** Date picker di kolom A
- **Column widths:** 110, 130, 180, 90, 80, 130, 130, 120

### 5.8 Kas (Build: `buildKas`)

- **Tab color:** Yellow (#F1C40F)
- **Frozen rows:** 4
- **Saldo display:** Baris 2 (PC), Baris 3 (UB) — update real-time
- **Data:** Mulai baris 5
- **Column widths:** 110, 100, 130, 200, 130, 130

### 5.9 Panduan (Build: `buildPanduan`)

- **Tab color:** Dark (#2C3E50)
- **Sections:** Cara Pakai, Harga, HPP Estimasi, Sheet Reference
- **Column widths:** 40, 400, 200, 100

---

## 6. Macro Reference

### 6.1 `simpanTransaksi()`

**Flow:**

1. Baca nama kasir dari POS!B2
2. Generate timestamp (Asia/Jakarta)
3. Scan Transaksi sheet backward untuk cari nomor TRX terakhir
4. Baca semua data POS dari `POS_START_ROW` hingga LastRow
5. Filter baris yang memiliki varian (skip kosong, skip GRAND TOTAL)
6. Batch write ke Transaksi (12 kolom)
7. Tulis summary bar (Total Cup, Total Base, Total Top, Total Bayar)
8. Panggil `stockEngineBOM()` untuk potong stok
9. Panggil `clearPOS()` untuk reset POS
10. Panggil `refreshLaporan()` untuk update pendapatan

**Format TRX:** `TRX-{angka}` (3 digit, padStart)

### 6.2 `stockEngineBOM(shPOS)`

**Flow:**

1. Baca data POS (varian, qty, topping string)
2. Parse topping string menjadi array (split koma)
3. Cocokkan tiap item dengan Resep (BOM)
4. Agregat kebutuhan bahan: `{namaBahan: totalTakaran}`
5. Update Stock:
   - Kolom E (Terjual) = existing + kebutuhan
   - Kolom F (Sisa Stok) = SisaAwal - kebutuhan

### 6.3 `addRowPOS()`

**Flow:**

1. Cari baris Grand Total (deteksi formula `SUM(H` di kolom H)
2. Hitung nomor urut baru
3. `insertRowBefore(totalRow)`
4. `styleOrderRow()` untuk baris baru
5. Update formula SUM Grand Total (geser 1 baris)
6. Update `POS_GRAND_TOTAL_ROW` di PropertiesService
7. Set active range ke kolom B baris baru

### 6.4 `clearPOS()`

**Flow:**

1. Hapus semua baris dari `POS_START_ROW` hingga LastRow
2. Rebuild `POS_INIT_ROWS` (5) baris order
3. Rebuild Grand Total row
4. Update PropertiesService

### 6.5 `pilihTopping()`

**Flow:**

1. Validasi: baris aktif >= `POS_START_ROW`, varian sudah dipilih
2. Ambil daftar topping dari `getToppingList()`
3. Baca topping yang sudah ada di kolom D
4. Tampilkan dialog `ui.prompt()` dengan checklist visual
5. Parse input: split koma, filter index valid, dedup
6. Tulis hasil ke kolom D format: "Keju, Boba"

### 6.6 `simpanPengeluaran()`

**Flow:**

1. Baca data Pengeluaran baris 4+
2. Filter: skip baris kosong, skip yang sudah `✓ Synced`
3. Untuk tiap baris valid:
   - Konversi satuan dari Bahan (`jumlah × packSize`)
   - Cari item di Stock:
     - Jika ada: update kolom F (Sisa Stok)
     - Jika tidak: insert baris baru di Stock
   - Tandai kolom H: `✓ Synced`
4. Panggil `refreshLaporan()`

### 6.7 `deleteRowPOS()`

**Flow:**

1. Validasi: sheet aktif = POS, baris di area order (bukan Grand Total)
2. Konfirmasi dialog
3. Hapus baris dari bawah ke atas
4. Panggil `_renumberAndFixPOS()`

### 6.8 `_renumberAndFixPOS(sh)`

**Flow:**

1. Iterasi dari `POS_START_ROW` hingga Grand Total
2. Update nomor urut di kolom A
3. Set ulang formula di kolom E-H (referensi baris mungkin bergeser)
4. Update formula SUM Grand Total
5. Update PropertiesService

### 6.9 `refreshLaporan()`

**Flow:**

1. Bangun HPP lookup dari BOM via `getHPPLookup()`
2. Baca semua data Transaksi (baris 3+)
3. Agregat per hari: `hariMap[tanggal] = {trx, cup, pendapatan, hppBahan}`
4. Baca Pengeluaran per hari, tambahkan ke `hariMap.pengeluaran`
5. Agregat per bulan: `bulanMap[MM/YYYY]`
6. Update Ringkasan Hari Ini (baris 5-11) dengan BOM HPP
7. Tulis Rekap Harian (baris 17+)
8. Tulis Rekap Bulanan (dinamis, setelah rekap harian + spacer)
9. Tampilkan alert diagnostic: total pendapatan, HPP, laba

### 6.10 Cash Management

| Fungsi | Deskripsi |
|---|---|
| `_getSaldoKas(kategori)` | Kalkulasi saldo dari awal (scan semua baris) |
| `_catatKas(kat, jenis, ket, jumlah)` | Catat transaksi + update display saldo |
| `_updateSaldoDisplay(sh, kat, saldo)` | Update baris 2/3 di sheet Kas |
| `topUpPC()` | Top Up PC ke Rp 100.000 |
| `topUpUB()` | Top Up UB jika < Rp 10.000 |
| `initSaldoKas()` | Set saldo awal harian (cegah duplikat) |

### 6.11 `setupOnEditTrigger()`

Cek apakah trigger `onEdit` sudah terinstall. Jika belum, buat baru.

### 6.12 `syncDropdownPOS()`

Refresh dropdown varian di POS (dipanggil otomatis saat Bahan diedit via `clearDynamicCache()` + `onEdit`).

### 6.13 `syncDropdownPengeluaran()`

Refresh dropdown Nama Item di Pengeluaran berdasarkan data Stock terbaru.

---

## 7. Trigger System

### 7.1 `onOpen()` — Custom Menu

Terpanggil otomatis saat spreadsheet dibuka. Mendaftarkan menu `🧋 POS` dengan item:

```
💾 Simpan Transaksi
➕ Add Row
🍬 Pilih Topping (baris aktif)
──────────
🗑 Clear POS
──────────
💸 Simpan & Sync Stok (Pengeluaran)
💸 Sinkronisasi Dropdown Pengeluaran
🔄 Refresh Laporan Pendapatan
──────────
🗑 Hapus Baris Aktif (POS)
──────────
➕ Tambah Resep / BOM
──────────
💰 Top Up PC ke Rp 100.000
🛒 Top Up UB (jika < Rp 10.000)
📅 Init Saldo Awal PC & UB
📅 Pilih Tanggal (Pengeluaran)
──────────
🔧 Setup Ulang (reset semua)
⚡ Install Auto-Fix Trigger
```

### 7.2 `onEdit(e)`

**Routing:**

| Sheet | Aksi |
|---|---|
| Bahan | `clearDynamicCache()` + `syncDropdownPOS()` |
| Pengeluaran | `onEditPengeluaran(e)` |
| POS | Auto-fix Grand Total jika ada baris dihapus manual |

### 7.3 `onEditPengeluaran(e)`

**Handler khusus untuk sheet Pengeluaran:**

| Kolom | Trigger | Aksi |
|---|---|---|
| B (Kategori) | Change | Update dropdown Nama Item via `_updateNamaItemDropdown()` |
| C (Nama Item) | Change | Auto-fill Satuan (dari Bahan), Harga (dari Bahan), Jumlah=1, Total formula, Status Stok |

### 7.4 `_updateNamaItemDropdown(ss, shPen, row)`

Filter dropdown kolom C berdasarkan kategori di kolom B:

| Kategori | Sumber Dropdown |
|---|---|
| Operasional | List fixed: Listrik, Air, Sewa Tempat, dll |
| Lain-lain | Free text (no validation) |
| Bahan Utama / Topping / Kemasan / Bahan Pendukung | Filter dari Bahan by kategori yang sama |

### 7.5 `onSelectionChange(e)`

Otomatis membuka date picker jika:
- Sheet aktif = Pengeluaran
- Kolom A (Tanggal)
- Baris >= 4
- Cell kosong

---

## 8. Dialog & HTML Service

### 8.1 Tambah Resep (`showTambahResepDialog`)

**HTML dialog untuk menambah resep baru:**

- Dropdown pilih menu existing atau tambah baru
- Input dinamis: tambah baris bahan baku (nama, takaran, satuan)
- Validasi: minimal 1 bahan baku
- Submit via `google.script.run.simpanTambahResep()`

**Spesifikasi:**
- Width: 450px, Height: 500px
- Font: Nunito
- Warna: Hijau (tambah), Merah (hapus), Oranye (submit)

### 8.2 Date Picker (`showDatePicker`, `showDatePickerForRow`)

**HTML dialog untuk memilih tanggal:**

- Input `<input type="date">`
- Tombol Simpan
- Callback: `setDateValue(row, dateStr)`
- Format output: DD/MM/YYYY

**Spesifikasi:**
- Width: 300px, Height: 200px
- Warna: Oranye (#E67E22)

---

## 9. Engine: BOM & HPP

### 9.1 `getHPPLookup()`

**Return:** `{namaProduk: totalHargaPerUnit}`

**Algoritma:**

1. Baca Bahan: build `hargaMap[namaBahan] = hargaPerPiece`
2. Baca Resep: untuk tiap baris, `hppMap[menu] += takaran × hargaMap[bahan]`
3. Return hppMap

### 9.2 HPP di `refreshLaporan()`

Untuk tiap baris transaksi:

```
hppProduk  = hppLookup[varian] × jumlahCup
hppTopping = Σ(hppLookup[tiapTopping]) × jumlahCup
totalHPP   = hppProduk + hppTopping
```

**Fallback:** Jika BOM lookup gagal, gunakan konstanta `HPP_PER_CUP` (2200) dan `HPP_PER_TOP` (80).

### 9.3 BOM Stock Engine

**Input:** Data POS (varian, qty, topping)
**Process:** Cocokkan dengan Resep → agregat kebutuhan bahan
**Output:** Update Stock (kolom Terjual + Sisa Stok)

---

## 10. Panduan Pengguna

### 10.1 Transaksi Harian

1. Buka sheet **POS**
2. Pilih **Product Variant** dari dropdown (kolom B)
3. Isi **Jumlah Cup** (kolom C) — default 1
4. Klik menu **🧋 POS → Pilih Topping** untuk tambah topping
5. Klik **🧋 POS → Simpan Transaksi**
6. Klik **🧋 POS → Clear POS** untuk reset

### 10.2 Menambah Baris Order

Klik **🧋 POS → Add Row**, atau ketik langsung di baris setelah Grand Total lalu klik Add Row.

### 10.3 Menghapus Baris Order

Klik baris yang ingin dihapus → **🧋 POS → Hapus Baris Aktif**. Bisa multi-select.

### 10.4 Manajemen Stok

1. Buka sheet **Pengeluaran**
2. Isi: Tanggal (klik cell kosong → date picker otomatis), Kategori, Nama Item (dropdown), Jumlah
3. Klik **🧋 POS → Simpan & Sync Stok**
4. Cek sheet **Stock** untuk status stok

### 10.5 Menambah Resep Baru

1. Pastikan bahan baku sudah ada di sheet **Bahan**
2. Klik **🧋 POS → Tambah Resep / BOM**
3. Pilih "Tambah Menu Baru" atau pilih menu existing
4. Tambah bahan baku dan takaran
5. Simpan

### 10.6 Laporan Pendapatan

Klik **🧋 POS → Refresh Laporan**. Data akan:
- Update Ringkasan Hari Ini
- Tulis Rekap Harian
- Tulis Rekap Bulanan

### 10.7 Manajemen Kas

- **Init Saldo Awal:** Setiap hari baru, klik **Init Saldo Awal PC & UB**
- **Top Up PC:** Klik **Top Up PC ke Rp 100.000** jika saldo menipis
- **Top Up UB:** Klik **Top Up UB** jika saldo < Rp 10.000

### 10.8 Setup Ulang

Jika ingin reset semua data:
1. Klik **🧋 POS → Setup Ulang**
2. Semua sheet akan dihapus dan dibuat ulang

---

## 11. Pengembangan

### 11.1 Menambah Varian Baru

1. Tambah bahan baku baru di sheet **Bahan** (jika diperlukan)
2. Klik **🧋 POS → Tambah Resep / BOM**
3. Isi nama varian dan komposisi bahan

### 11.2 Mengubah Harga

Edit konstanta di baris 7-10 `CanvaPOS.gs`:

```javascript
var HARGA_BASE    = 5000;   // Harga per cup
var HARGA_TOPPING = 1000;   // Harga per jenis topping
var HPP_PER_CUP   = 2200;   // Fallback HPP per cup
var HPP_PER_TOP   = 80;     // Fallback HPP per topping
```

### 11.3 Keterbatasan

- **Batas waktu eksekusi:** Apps Script memiliki batas 6 menit untuk fungsi biasa, 30 menit untuk trigger. `setupPOS()` mungkin timeout — gunakan step-by-step.
- **Single user:** Tidak ada mekanisme locking untuk multi-user concurrent access.
- **No undo:** Simpan Transaksi dan Clear POS tidak bisa di-undo.
- **Data tidak terenkripsi:** Semua data tersimpan di spreadsheet Google Sheets biasa.
- **No backup otomatis:** Disarankan backup berkala via File → Version History.
- **Font Nunito:** Jika font tidak terinstall di sistem, tampilan akan menggunakan font default.

### 11.4 Testing

Project tidak memiliki test suite. Semua pengujian dilakukan manual di spreadsheet. Disarankan untuk:
- Backup spreadsheet sebelum perubahan besar
- Test di spreadsheet copy terlebih dahulu
- Jalankan step-by-step setup untuk memverifikasi tiap sheet

### 11.5 V8 Runtime Notes

- Semua global menggunakan `var` (bukan `const`/`let`) karena lebih reliable di Apps Script V8
- Hindari destructuring assignment yang kompleks
- Gunakan `for` loop tradisional, bukan `for...of` untuk kompatibilitas maksimal

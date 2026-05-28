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

CanvaPOS adalah **Google Apps Script modular** (~4.330 baris, 13 file) yang berjalan di Google Sheets. Tidak ada dependency eksternal — deploy via clasp ke Apps Script project.

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
                           │
                      generateBOMData()
                      (template-based)
```

### 1.3 Struktur File (13 File)

```
CanvaPOS.gs          — Legacy data arrays (274 baris)
src/
├── appsscript.json       — Manifest (oauthScopes, runtime, timezone)
├── Config.gs             — Konstanta global (HARGA_BASE, COL, SHEET, env)
├── Helpers.gs            — Utility: UnitConverter, timers, styling, LockService
├── Builders.gs           — Semua sheet builder + setupPOS + protect + backup + named ranges
├── POS.gs                — addRowPOS, clearPOS, pilihTopping, deleteRowPOS, safeClear
├── Transaction.gs        — simpanTransaksi (3-phase), stockEngineBOM, TRX counter
├── Pengeluaran.gs        — Expense tracking, sync to Stock, auto-fill, unit validation
├── Kas.gs                — Petty Cash (PC) & Uang Belanja (UB) tracking
├── Laporan.gs            — Pendapatan builder, getHPPLookup (cached), refreshLaporan
├── Triggers.gs           — onOpen (menu), onEdit, onSelectionChange, trigger installer
├── Dialogs.gs            — Multi-select dialog, date picker, tambah resep, onboarding, env picker
└── data/
    ├── BahanData.gs      — getBahanData() — 40 master items
    └── ResepData.gs      — generateBOMData() — template-based BOM generator
```

### 1.4 Pola Desain

- **Modular**: Setiap domain punya file sendiri (Config, Helpers, POS, Transaction, dll)
- **Sheet-per-module**: Setiap modul bisnis memiliki sheet sendiri
- **Procedural**: Semua fungsi global, tidak ada class
- **Caching**: `CACHED_VARIAN_LIST`, `CACHED_TOPPING_LIST`, HPP cache via PropertiesService
- **Event-Driven**: Triggers `onEdit`, `onSelectionChange`, `onOpen`
- **LockService**: `withLock()` wrapper untuk concurrent-safe operations
- **3-Phase Pipeline**: simpanTransaksi → Read → Compute → Write (all-or-nothing)

### 1.5 Konvensi Kode

- Semua global menggunakan `var` (bukan `const`/`let`)
- Warna diakses via getter `getC()` — return object dengan 12 warna
- Nama fungsi: camelCase, helper dengan prefix `_`, setup dengan prefix `setup_`, build `build`
- Bahasa: Indonesia (variabel, komentar, UI)
- Timezone: `Asia/Jakarta`
- JSDoc: Semua fungsi publik punya `@param` + `@return`

---

## 2. Struktur Data

### 2.1 Sheet Bahan (Master Bahan Baku) — 40 Baris

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Kategori | String | Bahan Utama, Topping, Bahan Pendukung, Kemasan |
| B | Nama Bahan | String | Nama unik bahan baku |
| C | Ukuran/Pack | Number | Jumlah satuan per pack |
| D | Satuan | String | Gram, Liter, Kg, Pack, Piece |
| E | Harga Beli | Number (Rp) | Harga per pack |
| F | Harga Per Piece | Formula | `=IFERROR(E/D, 0)` |

### 2.2 Sheet Resep (Bill of Materials) — 168 Baris

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Nama Menu / Varian | String | Nama produk (key relasi ke POS) |
| B | Nama Bahan Baku | String | Nama bahan (foreign key ke Bahan) |
| C | Takaran (Qty Usage) | Number | Quantity per 1 unit produk |
| D | Satuan Penggunaan | String | Gram, Liter, Piece, dll |
| E | Harga Per Piece | Formula | VLOOKUP dari named range BAHAN_Lookup |
| F | Harga Per Takaran | Formula | `=C*E` |

### 2.3 Sheet Stock

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Kategori | String | Kategori bahan |
| B | Nama Bahan | String | Nama bahan (key) |
| C | Satuan | String | Satuan stok |
| D | Stok Awal | Number | Stok awal (0 — diisi via Pengeluaran) |
| E | Terjual | Number | Terpotong oleh transaksi (dari stockEngineBOM) |
| F | Sisa Stok | Number | Dikirim manual atau dari Pengeluaran sync |
| G | Min. Stok | Number | Batas minimal untuk restock alert |
| H | Status | Formula | `⚠ RESTOCK` atau `✓ OK` (conditional formatting) |

### 2.4 Sheet Transaksi

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | No. Trx | String | Format `TRX-001` |
| B | No. Item | Number | Urutan item dalam satu transaksi (1,2,3...) |
| C | Tanggal | Date/String | Format DD/MM/YYYY |
| D | Jam | String | Format HH:mm |
| E | Kasir | String | Nama kasir |
| F | Product Variant | String | Nama produk |
| G | Jumlah Cup | Number | QTY |
| H | Topping | String | Dipisah koma: "Keju, Boba" |
| I | Jml Topping | Number | Hitung dari kolom H |
| J | Harga Base | Number (Rp) | `=Cup*5000` |
| K | Harga Topping | Number (Rp) | `=JmlTop*Cup*1000` |
| L | Total | Number (Rp) | `=J+K` |

### 2.5 Sheet Pendapatan

**Ringkasan Hari Ini (baris 5-11):**

| Baris | Label | Sumber |
|---|---|---|
| 5 | Tanggal | `TODAY()` |
| 6 | Total Transaksi | `COUNTIF(TRX_Tgl, TODAY())` |
| 7 | Total Cup Terjual | `SUMIF(TRX_Tgl, TODAY(), TRX_Cup)` |
| 8 | Total Pendapatan | `SUMIF(TRX_Tgl, TODAY(), TRX_Total)` |
| 9 | Total HPP | `B7*2200 + SUMPRODUCT(...)*80` |
| 10 | Laba Bersih | `=Pendapatan-HPP` |
| 11 | Margin (%) | `=Laba/Pendapatan*100` |

**Status Kas (baris 12-13):** PC dan UB — LOOKUP dari sheet Kas

**Rekap Harian (baris 17+):** Dinamis dari Transaksi via `refreshLaporan()`

**Rekap Bulanan:** Agregasi dari rekap harian via `refreshLaporan()`

### 2.6 Sheet POS

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | No | Number (formula) | Nomor urut otomatis |
| B | Product Variant | Dropdown | Dari getVarianList() (sumber: Resep) |
| C | Jumlah Cup | Number | Default 1 |
| D | Topping - Jenis | String | Manual atau dari pilihTopping() |
| E | Topping - Jumlah | Formula | Hitung dari koma di D via COUNTA/SPLIT |
| F | Topping - Harga Top | Formula | `=E*C*1000` |
| G | Harga Base | Formula | `=C*5000` |
| H | Total | Formula | `=G+F` |

**Grand Total:** Baris terakhir dengan formula `=SUM(H7:H{n})`

### 2.7 Sheet Pengeluaran

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Tanggal | Date | DD/MM/YYYY — auto-fill via onSelectionChange |
| B | Kategori | Dropdown | Bahan Utama, Topping, Kemasan, Bahan Pendukung, Operasional, Lain-lain |
| C | Nama Item | Dropdown | Dinamis berdasarkan kategori |
| D | Satuan | Auto-fill | Dari Bahan |
| E | Jumlah | Number | QTY |
| F | Harga Satuan | Number (Rp) | Dari Bahan |
| G | Total | Formula | `=E*F` |
| H | Status Stok | Auto-fill | Indikator sisa stok atau `✓ Synced` |

### 2.8 Sheet Kas

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Tanggal | Date | DD/MM/YYYY |
| B | Kategori | String | PC atau UB |
| C | Jenis | String | Saldo Awal, Top Up, Pengeluaran, Setor |
| D | Keterangan | String | Deskripsi transaksi |
| E | Jumlah | Number (Rp) | Nominal |
| F | Saldo | Number (Rp) | Running balance |

### 2.9 Sheet Audit (Hidden)

| Kolom | Header | Tipe | Deskripsi |
|---|---|---|---|
| A | Timestamp | String | Tanggal + Jam |
| B | User | String | Email pengguna |
| C | Action | String | Nama aksi |
| D | Detail | String | Detail aksi |

Retensi: 90 hari (auto-clean via `cleanAuditLog()`)

---

## 3. Fungsi Global & Konstanta

### 3.1 Konstanta (`src/Config.gs`)

| Variabel | Nilai | Deskripsi |
|---|---|---|
| `HARGA_BASE` | 5000 | Harga per cup |
| `HARGA_TOPPING` | 1000 | Harga per jenis topping per cup |
| `HPP_PER_CUP` | 2200 | Estimasi HPP per cup (fallback BOM) |
| `HPP_PER_TOP` | 80 | Estimasi HPP per topping (fallback BOM) |
| `POS_START_ROW` | 7 | Baris pertama data order di POS |
| `POS_INIT_ROWS` | 5 | Jumlah baris order awal |

### 3.2 COL Constants (0-based, column indices)

```javascript
var COL = {
  BAHAN:       { KATEGORI:0, NAMA:1, SATUAN:2, PACK:3, HARGA_BELI:4, HARGA_PIECE:5 },
  RESEP:       { MENU:0, BAHAN:1, TAKARAN:2, SATUAN:3 },
  POS:         { NO:0, VARIAN:1, CUP:2, TOPPING:3, JML_TOP:4, HRG_TOP:5, HRG_BASE:6, TOTAL:7 },
  TRANSAKSI:   { NO_TRX:0, NO_ITEM:1, TGL:2, JAM:3, KASIR:4, VARIAN:5, CUP:6, TOPPING:7, JML_TOP:8, HRG_BASE:9, HRG_TOP:10, TOTAL:11 },
  STOCK:       { KATEGORI:0, NAMA:1, SATUAN:2, STOK_AWAL:3, TERJUAL:4, SISA:5, MIN:6, STATUS:7 },
  PENGELUARAN: { TGL:0, KATEGORI:1, NAMA:2, SATUAN:3, JUMLAH:4, HARGA:5, TOTAL:6, STATUS:7 },
  KAS:         { TGL:0, KATEGORI:1, JENIS:2, KET:3, JUMLAH:4, SALDO:5 },
  PENDAPATAN:  { LABEL:1, TRX:2, CUP:3, PENDAPATAN:4, HPP:5, LABA:6 },
};
function COLx(c) { return c + 1; }
```

### 3.3 SHEET Constants

```javascript
var SHEET = {
  PANDUAN: "Panduan", POS: "POS", STOCK: "Stock", TRANSAKSI: "Transaksi",
  PENDAPATAN: "Pendapatan", PENGELUARAN: "Pengeluaran", BAHAN: "Bahan",
  RESEP: "Resep", KAS: "Kas", AUDIT: "Audit"
};
```

### 3.4 Unit Converter

```javascript
UnitConverter.convert(qty, fromUnit, toUnit)  // Kg↔Gram, Liter↔ml
UnitConverter.toBase(qty, unit)                // → Gram / ml
UnitConverter.baseUnit(unit)                   // → "Gram" / "ml"
```

### 3.5 Color Palette — `getC()`

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

### 3.6 Caching System

| Fungsi | Cache | Lifetime |
|---|---|---|
| `getVarianList()` | `CACHED_VARIAN_LIST` | Per session (clear on Bahan edit) |
| `getToppingList()` | `CACHED_TOPPING_LIST` | Per session (clear on Bahan edit) |
| `getHPPLookup()` | PropertiesService (HPP_CACHE) | 1 jam TTL |

---

## 4. Setup System

### 4.1 `setupPOS(env?)`

Fungsi utama yang membangun seluruh workbook:

1. Set timezone `Asia/Jakarta`
2. Hapus semua sheet existing (Panduan, POS, Stock, Transaksi, Pendapatan, Pengeluaran, Kas, Bahan, Resep, Audit)
3. Build sheet berurutan (dengan `SpreadsheetApp.flush()` tiap step):
   `Bahan → Resep → Transaksi → Stock → Pengeluaran → Kas → Audit → Pendapatan → POS → Panduan`
4. Reorder sheet
5. Auto-protect semua formula cells via `protectAll()`
6. Setup named ranges via `setupNamedRanges()`
7. Tampilkan alert sukses

**Dependency order:**
- Pengeluaran & Kas harus SEBELUM Pendapatan (formula mereferensi sheet tersebut)
- Audit dibuat SEBELUM Pendapatan (auditLog dipanggil di protectAll)
- Bahan harus SEBELUM Resep (VLOOKUP)
- Stock & Transaksi harus SEBELUM Pendapatan (SUMIF/COUNTIF via named ranges)

### 4.2 Step-by-Step Functions

Jika `setupPOS()` timeout (6 menit batas Apps Script):

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

### 4.3 Named Ranges (8 ranges)

| Name | Sheet | Kolom | Fungsi |
|---|---|---|---|
| `TRX_Tgl` | Transaksi | C (Tgl) | SUMIF/COUNTIF di Pendapatan |
| `TRX_Cup` | Transaksi | G (Cup) | SUMIF di Pendapatan |
| `TRX_Total` | Transaksi | L (Total) | SUMIF di Pendapatan |
| `TRX_JmlTop` | Transaksi | I (Jml Top) | SUMPRODUCT di Pendapatan |
| `TRX_Varian` | Transaksi | F (Varian) | Lookup |
| `TRX_Topping` | Transaksi | H (Topping) | Lookup |
| `BAHAN_Lookup` | Bahan | B (Nama) | VLOOKUP di Resep |
| `PEN_Tgl` | Pengeluaran | A (Tgl) | Lookup |

---

## 5. Sheet Reference

### 5.1 Bahan (Build: `buildBahan`)

- **Tab color:** Blue (#2E86AB)
- **Frozen rows:** 1
- **Data:** 40 baris bahan baku (baris 2-41)
- **Data source:** `getBahanData()` di `src/data/BahanData.gs`
- **Column widths:** 100, 200, 100, 80, 120, 140

### 5.2 Resep (Build: `buildResep`)

- **Tab color:** Purple (#8E44AD)
- **Frozen rows:** 1
- **Struktur:** Many-to-many (1 menu → banyak bahan)
- **Data source:** `generateBOMData()` di `src/data/ResepData.gs`
- **Output:** 168 baris (18 Pop Ice × 8 + 2 Kopi × 6 + Es Teh × 6 + 6 topping × 1)
- **Formulas:** Kolom E (VLOOKUP named range BAHAN_Lookup), Kolom F (C×E)
- **Column widths:** 220, 200, 140, 130, 140, 140

### 5.3 Stock (Build: `buildStock`)

- **Tab color:** Green (#27AE60)
- **Frozen rows:** 1
- **Data:** Dinamis, diisi via Pengeluaran sync atau manual
- **Conditional formatting:** 2 rules (RESTOCK = merah, OK = hijau)
- **Min. Stok default:** 1 untuk item baru

### 5.4 Transaksi (Build: `buildTransaksi`)

- **Tab color:** Purple (#8E44AD)
- **Frozen rows:** 2
- **Title row:** Merged A1:L1
- **Data:** Mulai baris 3, diisi oleh `simpanTransaksi()`
- **12 kolom:** No.Trx, No.Item, Tanggal, Jam, Kasir, Varian, Cup, Topping, Jml Topping, Harga Base, Harga Top, Total

### 5.5 Pendapatan (Build: `buildPendapatan`)

- **Tab color:** Red (#E74C3C)
- **Frozen rows:** 1
- **Sections:**
  - Baris 1: Judul
  - Baris 2: Petunjuk refresh
  - Baris 4-11: Ringkasan Hari Ini (formula + BOM HPP)
  - Baris 12-13: Status Kas (LOOKUP ke sheet Kas)
  - Baris 15-16: Header Rekap Harian
  - Baris 17+: Data rekap harian (dinamis dari `refreshLaporan()`)
  - Setelah rekap harian + 2 baris: Header Rekap Bulanan + data

### 5.6 POS (Build: `buildPOS`)

- **Tab color:** Orange (#E67E22)
- **Frozen rows:** 6 (header bertingkat)
- **Layout:**
  - Baris 1: Judul "🧋 CanvaPOS"
  - Baris 2: Kasir, Tanggal (format Indonesia panjang), Jam real-time
  - Baris 3: Tombol aksi (Simpan, Add Row, Clear)
  - Baris 4: Info harga
  - Baris 5-6: Header kolom bertingkat (Topping: 3 sub-kolom D, E, F)
  - Baris 7-11: 5 baris order awal
  - Baris 12: Grand Total (posisi disimpan di PropertiesService)

### 5.7 Pengeluaran (Build: `buildPengeluaran`)

- **Tab color:** Orange (#E67E22)
- **Frozen rows:** 3
- **Dropdown kategori:** 200 baris (6 kategori)
- **Formula total:** 200 baris (E×F)
- **Data validation:** Date picker di kolom A (allow invalid)
- **Column widths:** 110, 130, 180, 90, 80, 130, 130, 120

### 5.8 Kas (Build: `buildKas`)

- **Tab color:** Yellow (#F1C40F)
- **Frozen rows:** 4
- **Saldo display:** Baris 2 (PC), Baris 3 (UB) — update real-time
- **Data:** Mulai baris 5
- **Column widths:** 110, 100, 130, 200, 130, 130

### 5.9 Panduan (Build: `buildPanduan`)

- **Tab color:** Dark (#2C3E50)
- **Sections:**
  - Baris 1-8: Dashboard (System Health, Transaksi Terakhir, Pendapatan Hari Ini, Stok Menipis, Saldo PC/UB)
  - Baris 9+: Panduan lengkap: Cara Pakai, Harga, Input Bahan (Gram), HPP Estimasi, Sheet Reference

### 5.10 Audit (Build: `buildAudit`)

- **Tab color:** Grey (#95A5A6)
- **Hidden sheet** — tidak terlihat di UI
- **4 kolom:** Timestamp, User, Action, Detail
- **Retensi:** 90 hari (auto-clean di `onOpen()`)
- **Frozen rows:** 1

---

## 6. Macro Reference

### 6.1 `simpanTransaksi()` — `src/Transaction.gs`

**3-Phase Pipeline (semua di dalam LockService):**

- **Phase 1 — Read & Validate:** Baca data POS, validasi varian, format batch rows
- **Phase 2 — Compute:** Cocokkan dengan Resep (BOM), hitung kebutuhan bahan di memory
- **Phase 3 — Write (all-or-nothing):** Generate TRX number, batch write ke Transaksi, summary bar, stock deduction, clear POS

**Format TRX:** `TRX-{angka}` (3 digit, padStart, O(1) via PropertiesService)

**Key features:**
- Same TRX per batch — semua item dalam 1 transaksi dapat nomor yang sama
- No. Item (1, 2, 3...) per row
- BOM-based stock deduction

### 6.2 `stockEngineBOM(shPOS)` — `src/Transaction.gs`

**Flow:**

1. Baca data POS (varian, qty, topping string)
2. Parse topping string → array
3. Cocokkan tiap item dengan Resep (BOM)
4. Agregat kebutuhan bahan: `{namaBahan: totalTakaran}`
5. Update Stock (batch write):
   - Kolom E (Terjual) = existing + kebutuhan
   - Kolom F (Sisa Stok) = SisaAwal - kebutuhan

### 6.3 `addRowPOS()` — `src/POS.gs`

**Flow:**

1. Cari baris Grand Total (deteksi formula `SUM(H` di kolom H)
2. Hitung nomor urut baru
3. `insertRowBefore(totalRow)`
4. `styleOrderRow()` untuk baris baru
5. Update formula SUM Grand Total
6. Update `POS_GRAND_TOTAL_ROW` di PropertiesService

### 6.4 `clearPOS()` — `src/POS.gs`

**Flow:**

1. Konfirmasi via `confirmAction()`
2. Hapus semua baris dari `POS_START_ROW` hingga LastRow
3. Rebuild `POS_INIT_ROWS` (5) baris order
4. Rebuild Grand Total row
5. Update PropertiesService

### 6.5 `pilihTopping()` — `src/POS.gs`

**Flow:**

1. Validasi: baris aktif >= `POS_START_ROW`, varian sudah dipilih
2. Ambil daftar topping dari `getToppingList()`
3. Baca topping yang sudah ada di kolom D
4. Tampilkan `showMultiSelectDialog()` — HTML checklist
5. Callback `_pilihToppingCallback()` — tulis hasil ke kolom D

### 6.6 `simpanPengeluaran()` — `src/Pengeluaran.gs`

**Flow:**

1. Baca data Pengeluaran baris 4+
2. Filter: skip baris kosong, skip yang sudah `✓ Synced`
3. Validasi unit consistency (P2.5)
4. Untuk tiap baris valid:
   - Konversi satuan dari Bahan (`jumlah × packSize`)
   - Cari item di Stock: update atau insert baris baru
   - Tandai kolom H: `✓ Synced`
5. Panggil `refreshLaporan()` + `refreshNamedRanges()`

### 6.7 `deleteRowPOS()` — `src/POS.gs`

**Flow:**

1. Validasi: sheet aktif = POS, baris di area order (bukan Grand Total)
2. Support multi-row selection
3. Konfirmasi dialog
4. Hapus baris dari bawah ke atas (safe index)
5. Panggil `_renumberAndFixPOS()`

### 6.8 `_renumberAndFixPOS(sh)` — `src/POS.gs`

1. Iterasi dari `POS_START_ROW` hingga Grand Total
2. Update nomor urut di kolom A
3. Set ulang formula di kolom E-H (referensi baris mungkin bergeser)
4. Update formula SUM Grand Total
5. Update PropertiesService

### 6.9 `refreshLaporan()` — `src/Laporan.gs`

**Flow:**

1. HPP lookup dari BOM via `getHPPLookup()` (cached 1 jam)
2. Baca semua data Transaksi (baris 3+)
3. Agregat per hari: `hariMap[tanggal] = {trx, cup, pendapatan, hppBahan}`
4. Baca Pengeluaran per hari, tambahkan ke `hariMap.pengeluaran`
5. Agregat per bulan: `bulanMap[MM/YYYY]`
6. Update Ringkasan Hari Ini (baris 5-11) dengan BOM HPP
7. Tulis Rekap Harian (baris 17+)
8. Tulis Rekap Bulanan (dinamis, setelah rekap harian + spacer)
9. Tampilkan alert diagnostic

### 6.10 Cash Management — `src/Kas.gs`

| Fungsi | Deskripsi |
|---|---|
| `_getSaldoKas(kategori)` | Kalkulasi saldo dari awal (scan semua baris) |
| `getSaldoPC()` | Saldo Petty Cash |
| `getSaldoUB()` | Saldo Uang Belanja |
| `_catatKas(kat, jenis, ket, jumlah)` | Catat transaksi + update display |
| `_updateSaldoDisplay(sh, kat, saldo)` | Update baris 2/3 di sheet Kas |
| `topUpPC()` | Top Up PC ke Rp 100.000 (closing) |
| `topUpUB()` | Top Up UB jika < Rp 10.000 |
| `initSaldoKas()` | Set saldo awal harian (cegah duplikat) |

### 6.11 Safe Clear & Restore — `src/POS.gs`

| Fungsi | Deskripsi |
|---|---|
| `safeClearPOS()` | Backup data POS ke PropertiesService lalu clear |
| `restorePOSFromBackup()` | Pulihkan data dari backup terakhir |

### 6.12 Dashboard — `src/Builders.gs`

| Fungsi | Deskripsi |
|---|---|
| `refreshDashboard()` | Update 6 metric di Panduan baris 2-7 |

### 6.13 Sheet Protection — `src/Builders.gs`

| Fungsi | Deskripsi |
|---|---|
| `protectAll(adminEmail?)` | Proteksi formula cells di semua 10 sheet |
| `unprotectAll(adminEmail?)` | Hapus semua proteksi |

### 6.14 Backup — `src/Builders.gs`

| Fungsi | Deskripsi |
|---|---|
| `backupSpreadsheet()` | Copy spreadsheet dengan timestamp |
| `cleanBackups(maxAgeDays=7)` | Hapus backup > 7 hari |
| `setupBackupTrigger(hour=2)` | Daily backup otomatis |

### 6.15 Named Ranges — `src/Builders.gs`

| Fungsi | Deskripsi |
|---|---|
| `setupNamedRanges()` | Buat 8 named ranges (10000 rows) |
| `refreshNamedRanges()` | Update ke ukuran data aktual |
| `clearNamedRanges()` | Hapus semua named range |

### 6.16 Environment — `src/Config.gs`

| Fungsi | Deskripsi |
|---|---|
| `getEnv()` | Baca environment (default: production) |
| `setEnv(env)` | Set environment |
| `getEnvConfig()` | Dapatkan konfigurasi environment |
| `setupEnv(env)` | Setup + refresh protection/named ranges |
| `showEnvPicker()` | HTML dialog pilih environment |

### 6.17 Audit — `src/Builders.gs`

| Fungsi | Deskripsi |
|---|---|
| `auditLog(action, detail)` | Catat aktivitas ke sheet Audit |
| `cleanAuditLog()` | Hapus log > 90 hari |

### 6.18 Sync Functions — `src/Triggers.gs`

| Fungsi | Deskripsi |
|---|---|
| `syncDropdownPOS()` | Refresh dropdown varian di POS dari Resep |
| `syncDropdownPengeluaran()` | Refresh dropdown item di Pengeluaran dari Stock |
| `setupOnEditTrigger()` | Install trigger onEdit (anti-duplicate) |

---

## 7. Trigger System

### 7.1 `onOpen()` — Custom Menu

Terpanggil otomatis saat spreadsheet dibuka. Mendaftarkan menu `🧋 POS` dengan 26 item:

```
💾 Simpan Transaksi
➕ Add Row
🍬 Pilih Topping (baris aktif)
──────────
🗑 Clear POS
♻ Safe Clear (backup dulu)
↩ Restore POS dari Backup
──────────
💸 Simpan & Sync Stok (Pengeluaran)
💸 Sinkronisasi Dropdown Pengeluaran
🔄 Refresh Laporan Pendapatan
📊 Refresh Dashboard
──────────
🗑 Hapus Baris Aktif (POS)
──────────
➕ Tambah Resep / BOM
──────────
💰 Top Up PC ke Rp 100.000
🛒 Top Up UB (jika < Rp 10.000)
📅 Init Saldo Awal PC & UB
📅 Pilih Tanggal (cell aktif)
──────────
🔧 Setup Ulang (reset semua)
🔒 Proteksi Semua Sheet
🔓 Unprotect Semua Sheet
──────────
📀 Backup Sekarang
⏰ Atur Backup Otomatis
──────────
🚀 Onboarding Wizard
──────────
🏷️ Refresh Named Ranges
──────────
🌐 Set Environment…
──────────
⚡ Install Auto-Fix Trigger
```

Juga menjalankan: `cleanAuditLog()`, `cleanBackups()`, `refreshDashboard()`.

### 7.2 `onEdit(e)`

**Routing:**

| Sheet | Aksi |
|---|---|
| Bahan | `clearDynamicCache()` + `clearHPPLookupCache()` + `syncDropdownPOS()` |
| Resep | `clearHPPLookupCache()` |
| Pengeluaran | `onEditPengeluaran(e)` |
| POS | Auto-fix Grand Total jika ada baris dihapus manual |

### 7.3 `onEditPengeluaran(e)` — `src/Pengeluaran.gs`

| Kolom | Trigger | Aksi |
|---|---|---|
| B (Kategori) | Change | Update dropdown Nama Item via `_updateNamaItemDropdown()` |
| C (Nama Item) | Change | Auto-fill Satuan, Harga, Jumlah=1, Total formula, Status Stok |

### 7.4 `_updateNamaItemDropdown(ss, shPen, row)`

Filter dropdown kolom C berdasarkan kategori:

| Kategori | Sumber Dropdown |
|---|---|
| Operasional | Fixed list: Listrik, Air, Sewa Tempat, dll |
| Lain-lain | Free text (no validation) |
| Bahan Utama / Topping / Kemasan / Bahan Pendukung | Filter dari Bahan by kategori |

### 7.5 `onSelectionChange(e)`

Auto-fill tanggal hari ini jika:
- Sheet aktif = Pengeluaran
- Kolom A (Tanggal)
- Baris >= 4
- Cell kosong

> Date picker tetap tersedia via menu **📅 Pilih Tanggal (cell aktif)**

---

## 8. Dialog & HTML Service

### 8.1 Multi-Select Checklist (`showMultiSelectDialog`)

Dialog generic untuk pilihan multiple:
- Items dengan checkbox
- Tombol "Pilih Semua / Hapus Semua"
- Callback function (parameter string comma-separated)
- Width: 350px, Height: 420px

**Penggunaan:** `pilihTopping()` — callback `_pilihToppingCallback`

### 8.2 Date Picker (`showDatePickerGeneric`, `showDatePickerForRowCol`)

Dialog untuk memilih tanggal:
- Input `<input type="date">`
- Tombol Simpan + Batal
- Callback: `setDateValueGeneric(row, col, dateStr)`
- Format output: DD/MM/YYYY
- Width: 300px, Height: 220px

### 8.3 Tambah Resep (`showTambahResepDialog`)

Dialog untuk menambah BOM:
- Dropdown pilih menu existing atau "+ Tambah Menu Baru"
- Input dinamis: tambah baris bahan baku (nama dropdown, takaran, satuan)
- Validasi: minimal 1 bahan baku
- Submit via `google.script.run.simpanTambahResep()`
- Width: 450px, Height: 500px

### 8.4 Onboarding Wizard (`showOnboardingWizard`)

4-step wizard:
1. Cek Setup — validasi semua sheet ada
2. Input Stok — cek stok sudah terisi
3. Initialize Kas — cek saldo PC/UB
4. Siap Transaksi — panduan memulai

Styling: gradient background, step progress bar, status cards.

### 8.5 Environment Picker (`showEnvPicker`)

Dialog radio button untuk pilih environment:
- Production (konfirmasi ON, debug OFF)
- Staging (konfirmasi ON, debug ON)
- Development (konfirmasi SKIP, debug ON)

---

## 9. Engine: BOM & HPP

### 9.1 BOM Generator — `generateBOMData()` (`src/data/ResepData.gs`)

Template-based generator, bukan hardcoded:

```
POP_ICE_TEMPLATE (7 ingredients) × 18 variants = 144 rows
+ Kopi Tubruk Robusta (6 rows)
+ Kopi Tubruk Arabika (6 rows)
+ Es Teh Original (6 rows)
+ 6 Toppings (6 rows)
= 168 rows total
```

### 9.2 HPP Lookup — `getHPPLookup()` (`src/Laporan.gs`)

**Return:** `{namaProduk: totalHargaPerUnit}`

**Algoritma:**
1. Baca Bahan: build `hargaMap[namaBahan] = hargaPerPiece` (dengan UnitConverter)
2. Baca Resep: `hppMap[menu] += takaran × hargaMap[bahan]`
3. Cache di PropertiesService (1 jam TTL)
4. Auto-clear saat Bahan/Resep diedit

### 9.3 HPP di `refreshLaporan()`

```
hppProduk  = hppLookup[varian] × jumlahCup
hppTopping = Σ(hppLookup[tiapTopping]) × jumlahCup
totalHPP   = hppProduk + hppTopping
```

**Fallback:** Gunakan konstanta `HPP_PER_CUP` (2200) dan `HPP_PER_TOP` (80) jika BOM lookup gagal.

**Catatan:** HPP = Bahan baku ONLY. Biaya operasional TIDAK termasuk HPP.

### 9.4 BOM Stock Engine

**Input:** Data POS (varian, qty, topping)
**Process:** Cocokkan dengan Resep → agregat kebutuhan bahan
**Unit handling:** Normalisasi via UnitConverter (Kg↔Gram, Liter↔ml)
**Output:** Batch update Stock (kolom Terjual + Sisa Stok)

---

## 10. Panduan Pengguna

### 10.1 Transaksi Harian

1. Buka sheet **POS**
2. Pilih **Product Variant** dari dropdown (kolom B)
3. Isi **Jumlah Cup** (kolom C)
4. Klik menu **🧋 POS → 🍬 Pilih Topping** untuk tambah topping
5. Klik **🧋 POS → 💾 Simpan Transaksi**
6. Klik **🧋 POS → 🗑 Clear POS** untuk reset

### 10.2 Menambah/Menghapus Baris Order

- **Tambah:** Klik **🧋 POS → ➕ Add Row**
- **Hapus:** Pilih baris → **🧋 POS → 🗑 Hapus Baris Aktif** (bisa multi-select)

### 10.3 Manajemen Stok via Pengeluaran

1. Buka sheet **Pengeluaran**
2. Isi: Tanggal (auto-fill), Kategori, Nama Item (dropdown), Jumlah
3. Klik **🧋 POS → 💸 Simpan & Sync Stok**
4. Cek sheet **Stock** untuk status

### 10.4 Manajemen Kas

- **Init Saldo Awal:** Setiap hari baru, klik **🧋 POS → 📅 Init Saldo Awal PC & UB**
- **Top Up PC:** Klik **🧋 POS → 💰 Top Up PC ke Rp 100.000**
- **Top Up UB:** Klik **🧋 POS → 🛒 Top Up UB (jika < Rp 10.000)**

### 10.5 Menambah Resep Baru

1. Pastikan bahan baku sudah ada di sheet **Bahan**
2. Klik **🧋 POS → ➕ Tambah Resep / BOM**
3. Pilih "Tambah Menu Baru" atau pilih menu existing
4. Tambah bahan baku dan takaran
5. Simpan

### 10.6 Laporan

Klik **🧋 POS → 🔄 Refresh Laporan Pendapatan** — update Ringkasan Harian, Rekap Harian, dan Bulanan.

### 10.7 Backup

- **Manual:** **🧋 POS → 📀 Backup Sekarang**
- **Otomatis:** **🧋 POS → ⏰ Atur Backup Otomatis** (daily jam 02:00, retensi 7 hari)

### 10.8 Setup Ulang

**🧋 POS → 🔧 Setup Ulang (reset semua)** — HAPUS SEMUA data.

---

## 11. Pengembangan

### 11.1 Menambah Varian Baru

**Cara 1 — Via Dialog (Recommended untuk non-dev):**
1. Tambah bahan baku baru di sheet **Bahan** (jika diperlukan)
2. Klik **🧋 POS → ➕ Tambah Resep / BOM**

**Cara 2 — Via Code (untuk dev, bulk add):**
1. Tambah bahan di `src/data/BahanData.gs` → `getBahanData()`
2. Tambah varian di `src/data/ResepData.gs` → `VARIAN_POP_ICE` array
3. Deploy via `clasp push`

### 11.2 Mengubah Harga

Edit di `src/Config.gs`:

```javascript
var HARGA_BASE    = 5000;
var HARGA_TOPPING = 1000;
```

### 11.3 Deploy

```bash
clasp login                    # OAuth via browser
clasp push                     # Push 13 files ke Apps Script
clasp deployments              # Lihat deployment ID
```

Script ID project: `18Eld0ZbczRsWqIxXYhK3y0DHZm2FKUO9jQ8Dvc-UI9DfHKAmItiElags`

### 11.4 Environment Setup

```javascript
setupEnv("development");  // skip confirmations, debug ON
setupEnv("staging");      // confirmations ON, debug ON
setupEnv("production");   // confirmations ON, debug OFF
```

Atau via menu: **🧋 POS → 🌐 Set Environment…**

### 11.5 Fitur Keamanan

- **LockService:** `withLock()` — aman multi-user concurrent access
- **Sheet Protection:** Formula cells diproteksi (bisa dibuka via `unprotectAll()`)
- **Audit Trail:** Semua operasi kritis tercatat
- **Backup Otomatis:** Daily backup + 7-day retention

### 11.6 Keterbatasan

- **Batas waktu eksekusi:** 6 menit fungsi biasa, 30 menit trigger
- **Font Nunito:** Jika tidak terinstall, pakai font default
- **Hanya untuk F&B UMKM:** Desain spesifik untuk minuman cup

### 11.7 V8 Runtime Notes

- Semua global menggunakan `var` (bukan `const`/`let`)
- Hindari destructuring assignment kompleks
- Gunakan `for` loop tradisional, bukan `for...of`
- `onSelectionChange` sebagai simple trigger — tidak bisa panggil `showModalDialog()`

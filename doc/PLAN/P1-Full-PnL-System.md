# P1 — Multi-Level P&L System (Full Financial Architecture)

## Objective
Upgrade CanvaPOS dari laporan laba-rugi 1-level (`Revenue - HPP = Laba`) menjadi **5-level P&L** proper:
`Revenue → Gross Profit → EBITDA → EBIT → Net Profit`

Dengan breakdown HPP per kategori bahan + integrasi OPEX, Penyusutan, dan Pajak.

---

## Overview P&L Structure (Target)

```
REVENUE (Total Penjualan)
  - HPP Bahan Utama          ← BOM ingredients kategori "Bahan Utama"
  - HPP Topping              ← BOM ingredients kategori "Topping"
  - HPP Bahan Pendukung      ← BOM ingredients kategori "Bahan Pendukung"
  - HPP Kemasan              ← BOM ingredients kategori "Kemasan"
= GROSS PROFIT (Laba Kotor)
  - OPEX (Biaya Operasional) ← Dari Pengeluaran kategori "Operasional"
= EBITDA
  - Penyusutan (Depreciation)← Dari Aset Tetap (module baru)
= EBIT (Laba Operasional)
  - Pajak & Bunga            ← Dari Pengeluaran atau input manual
= NET PROFIT (Laba Bersih Final)
```

---

## Files yang Akan Diubah/Dibuat

| # | File | Action | Deskripsi |
|---|------|--------|-----------|
| 1 | `src/Config.gs` | Edit | Tambah COL.PENDAPATAN baru, konstanta P&L |
| 2 | `src/data/BahanData.gs` | Edit | Tambah mapping kategori HPP per bahan |
| 3 | `src/data/ResepData.gs` | Edit | Tambah kolom kategori ingredient di BOM |
| 4 | `src/Builders.gs` | Edit | Rebuild `buildPendapatan()` dengan struktur baru |
| 5 | `src/Laporan.gs` | Edit | Rewrite `refreshLaporan()` untuk multi-level P&L |
| 6 | `src/Pengeluaran.gs` | Edit | Tambah sub-kategori OPEX untuk mapping akurat |
| 7 | `src/Kas.gs` | No change | Kas tetap sebagai module terpisah, referensi di dashboard |
| 8 | **NEW: `src/Aset.gs`** | **Create** | Module Aset Tetap & Penyusutan |
| **NEW** | **NEW: `src/Dialogs.gs`** | Create (if needed) | Dialog tambah aset |
| 9 | `src/Triggers.gs` | Edit | Menu items baru: Kelola Aset, Hitung Penyusutan |
| 10 | `CHANGELOG.md` | Edit | Dokumentasi perubahan |

---

## Step-by-Step Implementation

### Step 1 — Konfigurasi & Constants (`Config.gs`)

**Perubahan:**
- Update `COL.PENDAPATAN` dari 6 kolom jadi 10+ kolom:
```javascript
PENDAPATAN: {
  LABEL:0,           // A — Tanggal/Bulan
  TRX:1,             // B — Total Transaksi
  CUP:2,             // C — Total Cup
  REVENUE:3,         // D — Pendapatan Kotor
  HPP_BAHAN:4,       // E — HPP Bahan Utama
  HPP_TOPPING:5,     // F — HPP Topping
  HPP_SUPPORT:6,     // G — HPP Bahan Pendukung (gula, SKM, air, es)
  HPP_KEMASAN:7,     // H — HPP Kemasan
  GROSS_PROFIT:8,    // I — Laba Kotor = Revenue - Total HPP
  OPEX:9,            // J — Biaya Operasional (dari Pengeluaran)
  EBITDA:10,         // K — Laba sebelum penyusutan
  DEPRESIASI:11,     // L — Biaya Penyusutan (dari Aset)
  EBIT:12,           // M — Laba Operasional
  PAJAK:13,          // N — Pajak & Bunga
  NET_PROFIT:14      // O — Laba Bersih Final
}
```

- Tambah konstanta:
```javascript
var PAJAK_PERSEN = 0;    // Default 0%, bisa diubah manual (PPh Final UMKM = 0.5%)
```

### Step 2 — Kategori HPP di BOM (`ResepData.gs`)

**Perubahan:** Setiap ingredient di BOM perlu tau kategori HPP-nya.

Strategi: **Tambah `_kategoriBahanMap`** di `generateBOMData()` yang mapping tiap bahan ke kategorinya. Bisa juga dibaca dari `getBahanData()` secara langsung — setiap bahan dari BahanData sudah punya kategori.

```javascript
// Di generateBOMData()
function getKategoriBahan() {
  var data = getBahanData();
  var map = {};
  data.forEach(function(row) {
    map[row[0] /* KATEGORI */] = true;
  });
  return map;
}
```

Tapi pendekatan lebih baik: **Di `getHPPLookup()`, kembalikan breakdown per kategori, bukan total flat.**

**Strategi final:** Buat fungsi baru `getHPPBreakdown(menuName)` yang return:
```javascript
{
  total: 2200,
  byCategory: {
    "Bahan Utama": 1500,    // sachet Pop Ice
    "Bahan Pendukung": 300, // gula + SKM + air + es
    "Kemasan": 400          // cup + tutup + sedotan
  }
}
```

### Step 3 — HPP Breakdown Engine (`Laporan.gs` — fungsi baru)

Buat 2 fungsi baru:

**A. `getCategoryForBahan(bahanName)`**
Mapping tiap nama bahan dari BahanData ke kategori HPP-nya.

```javascript
function getCategoryForBahan(bahanName) {
  var data = getBahanData();
  for (var i = 0; i < data.length; i++) {
    if (data[i][COL.BAHAN.NAMA] === bahanName) {
      var kat = data[i][COL.BAHAN.KATEGORI];
      // Map kategori Bahan ke kategori HPP
      var map = {
        "Bahan Utama": "Bahan Utama",
        "Topping": "Topping",
        "Bahan Pendukung": "Bahan Pendukung",
        "Kemasan": "Kemasan"
      };
      return map[kat] || "Lain-lain";
    }
  }
  return "Lain-lain";
}
```

**B. `getHPPBreakdown(varianMartabak, qty, toppings?)`**
Read Resep BOM + Bahan prices, return breakdown per category:

```javascript
function getHPPBreakdown(varian, qty, toppingArray) {
  var hppCat = { "Bahan Utama":0, "Topping":0, "Bahan Pendukung":0, "Kemasan":0 };
  var resepData = getSheet(SHEET.RESEP).getDataRange().getValues();
  var priceMap = getBahanHargaMap(); // { "Gula Pasir": 20, ... }
  
  // Sum BOM ingredients for this variant
  for (var i = 1; i < resepData.length; i++) {
    if (resepData[i][COL.RESEP.MENU] === varian) {
      var bahan = resepData[i][COL.RESEP.BAHAN];
      var takaran = Number(resepData[i][COL.RESEP.TAKARAN]) || 0;
      var harga = priceMap[bahan] || 0;
      var cat = getCategoryForBahan(bahan);
      hppCat[cat] += takaran * harga * qty;
    }
  }
  
  // Same for toppings
  if (toppingArray) { ... }
  
  return hppCat;
}
```

### Step 4 — Restruktur Pendapatan Sheet (`Builders.gs`)

`buildPendapatan()` diubah menjadi 3 section:

```
Row 1:   💰 Laporan Laba/Rugi — merged
Row 2:   Info bar
Row 4:   ── Ringkasan Hari Ini ── (section header)
Row 5:   Tanggal                         | formula TODAY
Row 6:   Total Transaksi                 | COUNTIF
Row 7:   Total Cup Terjual               | SUMIF
Row 8:   A. TOTAL PENDAPATAN (REVENUE)   | SUMIF → value (BOLD, GREEN)
Row 9:   B.1 HPP — Bahan Utama           | dari BOM breakdown (value)
Row 10:  B.2 HPP — Topping               | dari BOM breakdown (value)
Row 11:  B.3 HPP — Bahan Pendukung       | dari BOM breakdown (value)
Row 12:  B.4 HPP — Kemasan               | dari BOM breakdown (value)
Row 13:  B   TOTAL HPP                   | =SUM(B9:B12) (BOLD)
Row 14:  C   LABA KOTOR (GROSS PROFIT)   | =A-B (BOLD, LGREEN)
Row 16:  ── Biaya Operasional ── (section header)
Row 17:  D.1 Gaji Karyawan               | dari Pengeluaran (filter)
Row 18:  D.2 Sewa Tempat                 | dari Pengeluaran (filter)
Row 19:  D.3 Listrik                     | dari Pengeluaran (filter)
Row 20:  D.4 Transportasi                | dari Pengeluaran (filter)
Row 21:  D.5 Operasional Lainnya          | dari Pengeluaran (filter)
Row 22:  D   TOTAL OPEX                  | =SUM(D1:D5) (BOLD)
Row 23:  E   EBITDA                      | =C-D (BOLD)
Row 25:  ── Penyusutan ── (section header)
Row 26:  F   Total Penyusutan            | dari module Aset (value)
Row 27:  G   EBIT (LABA OPERASIONAL)     | =E-F (BOLD)
Row 29:  ── Pajak & Bunga ── (section header)
Row 30:  H   Pajak & Bunga               | dari Pengeluaran atau manual
Row 31:  I   LABA BERSIH (NET PROFIT)    | =G-H (BOLD, BLUE/WHITE)
Row 33:  💳 PC & UB balance              | formula LOOKUP ke Kas (seperti sekarang)
Row 35:  ── Rekap Harian ── (section header)
Row 36:  Headers: 15 columns
Row 37+: Data rows (dinamis, tiap hari 1 baris)
```

**Kolom layout untuk rekap harian (A-O):**

```
A: Tanggal
B: Total Trx
C: Total Cup
D: Revenue
E: HPP-Bahan Utama
F: HPP-Topping
G: HPP-Bahan Pendukung
H: HPP-Kemasan
I: Gross Profit
J: OPEX
K: EBITDA
L: Depresiasi
M: EBIT
N: Pajak
O: Net Profit
```

### Step 5 — Module Baru: Aset & Penyusutan (`src/Aset.gs`)

**Fungsi:**
- Track aset tetap (blender, sealer, kulkas, dll)
- Hitung depresiasi bulanan (metode garis lurus / straight-line)
- Post depresiasi ke P&L tiap bulan

**Sheet baru: "Aset" (11th sheet)**
```
Row 1: 📦 Daftar Aset Tetap — merged header
Row 2: Headers:
  A: Nama Aset
  B: Kategori Aset
  C: Tanggal Beli
  D: Harga Perolehan
  E: Umur Ekonomis (bulan)
  F: Nilai Residu
  G: Penyusutan per Bulan      (formula: = (D-F) / E )
  H: Akumulasi Penyusutan      (dari Jurnal Bulanan)
  I: Nilai Buku Saat Ini       (formula: = D - H )
Row 3+: Data aset

Row berikutnya:
  Ringkasan: Total Penyusutan Bulan Ini
  Jurnal Bulanan (tiap bulan 1 baris):
    A: Bulan, B: Total Penyusutan, C: Akumulasi
```

**Key functions:**

```javascript
function getAsetData()          // Read all assets from Aset sheet
function getTotalDepresiasiBulan(bulan)  // Get total depreciation for a month
function postingDepresiasi()    // Auto-post monthly depreciation to jurnal
function addAset()              // Dialog tambah aset baru
function buildAset(ss)          // Sheet builder
```

### Step 6 — Update `refreshLaporan()` (`Laporan.gs`)

Rewrite function untuk:

1. **Read Transaksi** → aggregate revenue + cup per hari (sama seperti sekarang)
2. **HPP Breakdown** → untuk tiap transaksi, hitung HPP per kategori pakai `getHPPBreakdown()`
3. **OPEX** → baca Pengeluaran kategori "Operasional", filter per tanggal
4. **Depresiasi** → baca dari sheet Aset, ambil total per bulan
5. **Write multi-level P&L** → tulis ke 15 kolom Rekap Harian + Ringkasan Hari Ini
6. **Rekap Bulanan** → SUM semua kolom per bulan

### Step 7 — Update Pengeluaran UX (minor)

Tambah field **"Sub-Kategori OPEX"** untuk expense kategori Operasional:
- Gaji Karyawan
- Sewa Tempat
- Sewa Tempat Penitipan
- Listrik
- Air
- Transportasi
- Internet
- ATK
- Kebersihan
- Gas LPG
- Lainnya

Ini biar OPEX report bisa breakdown per jenis biaya.

Cara: Dropdown kolom B pilih "Operasional" → kolom C (Nama Item) sudah ada daftar ini dari `ITEM_OPERASIONAL`. Gak perlu perubahan struktur — tinggal di report aja difilter per nama item.

### Step 8 — Update Menu (`Triggers.gs`)

Tambah menu items:
- `📦 Tambah Aset Tetap` → `addAset()`
- `📊 Posting Penyusutan` → `postingDepresiasi()`
- `📋 Lihat Daftar Aset` → navigate to Aset sheet

### Step 9 — Update COL Constants

`COL` di `Config.gs` perlu tambahan:
```javascript
ASET: { NAMA:0, KATEGORI:1, TGL_BELI:2, HARGA:3, UMUR:4, RESIDU:5,
        DEPRESIASI_BLN:6, AKUMULASI:7, NILAI_BUKU:8 }
```

Dan `SHEET`:
```javascript
ASET: "Aset"
```

---

## Tabel Perubahan per File (Ringkasan)

| File | Perubahan |
|------|-----------|
| `Config.gs` | COL.PENDAPATAN 6→15 kolom, tambah COL.ASET, SHEET.ASET, PAJAK_PERSEN |
| `ResepData.gs` | Tambah fungsi `getKategoriBahanMap()` atau inline mapping |
| `BahanData.gs` | No change (kategori sudah ada) |
| `Helpers.gs` | (opsional) Fungsi `getCategoryForBahan()` |
| **NEW `Aset.gs`** | Module Aset: addAset, postingDepresiasi, buildAset, getTotalDepresiasi |
| `Builders.gs` | `buildPendapatan()` → restruktur 5-level P&L. `buildAset()` → sheet baru |
| `Laporan.gs` | `refreshLaporan()` → rewrite. `getHPPBreakdown()` → fungsi baru |
| `Pengeluaran.gs` | Minor: sub-kategori OPEX sudah support via ITEM_OPERASIONAL |
| `Triggers.gs` | Menu items baru: Aset, Depresiasi |
| `Dialogs.gs` | (opsional) Dialog tambah aset |

---

## Data Flow Diagram

```
[Transaksi] ──→ refreshLaporan()
                    │
                    ├──→ [getHPPBreakdown()]
                    │       ├── Baca Resep (BOM)
                    │       ├── Baca Bahan (harga)
                    │       └── Return { "Bahan Utama": ..., "Topping": ..., "Kemasan": ... }
                    │
                    ├──→ [Pengeluaran] ─→ OPEX by category
                    │
                    ├──→ [Aset Sheet] ──→ Total Depresiasi Bulan Ini
                    │
                    └──→ Write ke P&L:
                          Revenue → HPP breakdown → Gross Profit → OPEX
                          → EBITDA → Depresiasi → EBIT → Pajak → Net Profit
```

---

## Risk & Mitigation

| Risk | Mitigasi |
|------|----------|
| `refreshLaporan()` makin lambat | Panggil hanya via menu (manual), bukan otomatis. Tambah progress indicator. |
| HPP breakdown beda dengan HPP total sebelumnya | Test: jumlah semua kategori HPP harus = HPP lama (`HPP_PER_CUP` fallback) |
| Kolom Rekap Harian 15 kolom kepotong lebar sheet | Set column width otomatis (80-150px depending on content) |
| Aset sheet baru bikin total sheet jadi 11 | update `SHEET` constant, `protectAll()`, `setupPOS()` cleanup |
| Depresiasi bulk write banyak baris | `withLock()` + batch operation |

---

## Testing Plan

1. **Unit:** `getHPPBreakdown("Pop Ice - Cokelat", 1)` → cek total = 2200, kategori ratio wajar
2. **Integration:** `refreshLaporan()` → cek Ringkasan Hari Ini 15 kolom terisi
3. **Sheet:** `buildPendapatan()` → cek row alignment, formula, section headers
4. **Aset:** `addAset()` → cek depresiasi formula, posting ke jurnal
5. **Menu:** All new menu items muncul dan jalan
6. **OPEX filter:** `refreshLaporan()` OPEX total = sum Pengeluaran "Operasional" per tanggal

---

## Milestones

| Step | Estimasi | Depends On |
|------|----------|------------|
| Step 1: Config + COL | 15 min | - |
| Step 2: HPP category map | 20 min | BahanData |
| Step 3: HPP Breakdown engine | 45 min | Step 1+2 |
| Step 4: buildPendapatan restructure | 60 min | Step 1 |
| Step 5: Aset.gs (new module) | 90 min | Step 1 |
| Step 6: refreshLaporan rewrite | 90 min | Step 2+3+4 |
| Step 7: Pengeluaran minor | 15 min | - |
| Step 8: Menu items | 10 min | Step 5 |
| Step 9: Docs (CHANGELOG) | 15 min | All steps |
| **Total** | **~6 jam** | |

---

## Catatan Tambahan

- **HPP_PER_CUP (2200)** dan **HPP_PER_TOP (80)** tetap sebagai fallback — jangan dihapus
- **PC/UB** tetap sebagai modules terpisah (tidak digabung ke P&L). Tapi saldonya tetap ditampilkan di dashboard.
- **Pajak**: Default 0%. UMKM dengan omzet < 500jt/tahun = 0% PPh Final (UU HPP). Kalau perlu, user bisa set manual.
- **Format Rp**: All currency columns pakai `"Rp "#,##0`
- **Protection**: Formula columns di Pendapatan harus diproteksi

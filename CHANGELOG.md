# Changelog

## [1.0.0] — 2026-05-28

### Added

- **Setup System** — `setupPOS()` membangun 9 sheet sekaligus dengan dependency ordering
- **Setup Step-by-Step** — 8 fungsi individual (`setup_1_Bahan()` hingga `setup_8_Reorder()`) untuk menghindari timeout
- **Color Palette** — Sistem warna konsisten via `getC()` — biru, hijau, oranye, merah, ungu, kuning

### Sheet: Bahan

- Master data 37 bahan baku: Pop Ice (18 varian), kopi (2 jenis), teh, 6 topping, 5 bahan pendukung, 6 kemasan
- Kolom Harga Per Piece otomatis (`=E/D`)
- Zebra striping + format mata uang Rp

### Sheet: Resep (Bill of Materials)

- 18 varian Pop Ice dengan komposisi 8 bahan masing-masing
- 2 varian Kopi Tubruk (Robusta & Arabika) dengan 6 bahan
- Es Teh Original dengan 6 bahan
- 6 topping dengan takaran gram
- Formula VLOOKUP harga bahan + Harga Per Takaran otomatis

### Sheet: Stock

- Monitoring stok dengan kolom: Stok Awal, Terjual, Sisa Stok, Min. Stok, Status
- Conditional formatting: merah (`⚠ RESTOCK`) / hijau (`✓ OK`)
- Data diisi dinamis dari Pengeluaran

### Sheet: Transaksi

- Log transaksi 12 kolom: No. Trx, No. Item, Tanggal, Jam, Kasir, Varian, Cup, Topping, Jml Topping, Harga Base, Harga Top, Total
- Summary bar per transaksi (Total Cup, Total Bayar)
- Data diisi via macro Simpan Transaksi

### Sheet: Pendapatan

- Ringkasan Hari Ini: Total Transaksi, Cup, Pendapatan, HPP, Laba, Margin
- Status Kas (PC & UB) via formula LOOKUP
- Rekap Harian (dinamis dari Transaksi)
- Rekap Bulanan (agregasi dari harian)
- Perhitungan HPP menggunakan BOM lookup

### Sheet: POS

- Header: Kasir, Tanggal (format Indonesia), Jam real-time
- Tombol aksi: Simpan, Add Row, Clear
- 8 kolom order: No, Varian (dropdown), Cup, Topping (jenis+jumlah+harga), Harga Base, Total
- Grand Total dengan formula SUM
- Dropdown varian dinamis dari sheet Resep

### Sheet: Pengeluaran

- 8 kolom: Tanggal, Kategori (dropdown), Nama Item (dropdown dinamis), Satuan, Jumlah, Harga Satuan, Total, Status Stok
- Auto-fill Satuan & Harga dari Bahan saat Nama Item dipilih
- Indikator status stok (sisa + kritis)
- Date picker otomatis saat klik cell tanggal kosong

### Sheet: Kas

- Dua akun: Petty Cash (PC) & Uang Belanja (UB)
- 6 kolom: Tanggal, Kategori, Jenis, Keterangan, Jumlah, Saldo
- Saldo otomatis kalkulasi ulang dari awal
- Display saldo real-time di baris 2-3

### Sheet: Panduan

- Panduan penggunaan lengkap dalam sheet

### Macro: Simpan Transaksi

- Baca data POS, generate nomor TRX otomatis (scan backward)
- Batch write ke Transaksi + summary bar + separator
- Trigger stock engine BOM, clear POS, refresh laporan
- Zebra striping + format Rp

### Engine: BOM Stock

- Kalkulasi kebutuhan bahan dari POS berdasarkan Resep
- Update kolom Terjual & Sisa Stok di Stock
- Dukung multi-item (produk + topping)

### Macro: Add Row

- Insert baris baru sebelum Grand Total
- Renumber otomatis + update formula SUM

### Macro: Clear POS

- Hapus semua baris order, rebuild 5 baris awal + Grand Total

### Macro: Pilih Topping

- Dialog prompt dengan daftar topping dinamis dari Bahan
- Checklist visual, dukung multi-select via nomor

### Macro: Simpan & Sync Stok (Pengeluaran)

- Baca baris baru, konversi satuan dari Bahan, update/insert ke Stock
- Tandai status `✓ Synced` otomatis
- Skip baris yang sudah di-sync

### Macro: Hapus Baris Aktif POS

- Dukung multi-row selection
- Hapus dari bawah ke atas (safe index)
- Renumber + fix formula otomatis

### Trigger: onEdit

- Routing otomatis: Bahan → sync dropdown POS, Pengeluaran → auto-fill
- Auto-fix Grand Total jika ada baris dihapus manual

### Trigger: onSelectionChange

- Auto date picker di kolom Tanggal Pengeluaran (cell kosong)

### Dialog: Tambah Resep

- Pilih menu existing atau tambah baru
- Tambah bahan baku dinamis (1–∞)
- Simpan ke sheet Resep + styling

### System: Cache

- Dynamic list caching untuk varian & topping
- Clear cache otomatis saat Bahan diedit
- PropertiesService untuk menyimpan posisi Grand Total

### System: Trigger Installer

- `setupOnEditTrigger()` — pasang trigger onEdit dengan pengecekan duplikat

### Date Picker

- HTML dialog kalender untuk kolom Tanggal di Pengeluaran
- Terbuka otomatis via `onSelectionChange`
- Format output: DD/MM/YYYY

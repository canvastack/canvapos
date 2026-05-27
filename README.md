# CanvaPOS

**Point-of-Sale System untuk Google Sheets** — Dirancang untuk usaha minuman skala kecil (Pop Ice blender, kopi tubruk, es teh) dengan manajemen stok berbasis Bill of Materials (BOM).

## Fitur

- **POS Entry** — Input pesanan dengan dropdown varian, jumlah cup, dan topping
- **Hitung Otomatis** — Harga base (Rp 5.000) + topping (Rp 1.000/jenis) terhitung real-time
- **Transaksi Log** — Setiap pesanan tersimpan dengan timestamp, kasir, dan detail lengkap
- **Manajemen Stok** — Stok bahan terpotong otomatis berdasarkan BOM setiap transaksi
- **Laporan Laba/Rugi** — Rekap harian & bulanan dengan perhitungan HPP dari BOM
- **Manajemen Kas** — Petty Cash (PC) & Uang Belanja (UB) dengan saldo otomatis
- **Catatan Pengeluaran** — Sync otomatis ke stok + indikator status stok
- **Dialog Interaktif** — Pilihan topping, date picker, tambah resep via popup
- **Custom Menu** — Menu `🧋 POS` muncul otomatis saat spreadsheet dibuka

## Persyaratan

- Google Sheets
- Google Apps Script (V8 runtime)
- Font **Nunito** (opsional, untuk tampilan optimal)

## Instalasi

1. Buka Google Sheets baru
2. Klik **Extensions → Apps Script**
3. Hapus kode default, salin seluruh isi `CanvaPOS.gs` ke editor
4. Simpan (Ctrl+S) — beri nama project "CanvaPOS"
5. Klik **Run → setupPOS()** — izinkan permissions yang diminta
6. Kembali ke sheet, refresh halaman
7. Menu `🧋 POS` akan muncul di toolbar

> Jika `setupPOS()` timeout, jalankan fungsi satu per satu:
> `setup_1_Bahan()` → `setup_2_Resep()` → `setup_3_Transaksi()` → `setup_4_Stock()` → `setup_5_Pengeluaran()` → `setup_5b_Kas()` → `setup_5c_Pendapatan()` → `setup_6_POS()` → `setup_7_Panduan()` → `setup_8_Reorder()`

## Struktur Sheet

| Sheet | Fungsi |
|---|---|
| **Panduan** | Manual pengguna |
| **POS** | Input pesanan utama |
| **Stock** | Monitoring stok bahan + restock alert |
| **Transaksi** | Log history transaksi |
| **Pendapatan** | Laporan laba/rugi harian & bulanan |
| **Pengeluaran** | Catatan pembelian & biaya operasional |
| **Kas** | Tracking Petty Cash & Uang Belanja |
| **Bahan** | Master data bahan baku (harga, ukuran) |
| **Resep** | Bill of Materials — komposisi setiap produk |

## Menu `🧋 POS`

| Menu | Fungsi |
|---|---|
| 💾 Simpan Transaksi | Simpan pesanan ke log + update stok |
| ➕ Add Row | Tambah baris order baru |
| 🍬 Pilih Topping | Dialog pilihan topping untuk baris aktif |
| 🗑 Clear POS | Reset semua baris order |
| 💸 Simpan & Sync Stok | Simpan pengeluaran → update stok |
| 💸 Sinkronisasi Dropdown | Refresh dropdown di sheet Pengeluaran |
| 🔄 Refresh Laporan | Update laporan pendapatan |
| 🗑 Hapus Baris Aktif | Hapus baris order tertentu |
| ➕ Tambah Resep / BOM | Dialog tambah resep baru |
| 💰 Top Up PC / 🛒 Top Up UB | Top up kas otomatis |
| 📅 Init Saldo Awal | Set saldo awal PC & UB harian |
| 📅 Pilih Tanggal | Date picker untuk Pengeluaran |
| 🔧 Setup Ulang | Reset semua sheet |
| ⚡ Install Auto-Fix Trigger | Pasang trigger onEdit |

## Teknologi

- **Bahasa:** JavaScript (Google Apps Script, V8)
- **UI:** Google Sheets cells + HtmlService (HTML/CSS/JS dialogs)
- **Storage:** Spreadsheet rows + PropertiesService
- **Trigger:** `onEdit`, `onSelectionChange`, `onOpen`
- **Harga:** Rp 5.000/cup, Rp 1.000/jenis topping

## Lisensi

Hak cipta CanvaStack.

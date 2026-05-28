# CanvaPOS

**Point-of-Sale System untuk Google Sheets** — Dirancang untuk usaha minuman skala kecil (Pop Ice blender, kopi tubruk, es teh) dengan manajemen stok berbasis Bill of Materials (BOM). **Modular, 12 file, 90% complete.**

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

- [Node.js](https://nodejs.org) + [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
- Google Apps Script (V8 runtime)
- Font **Nunito** (opsional, untuk tampilan optimal)

## Instalasi (clasp — modular)

```bash
# 1. Clone repo
git clone https://github.com/canvastack/canvapos.git
cd canvapos

# 2. Login clasp
clasp login

# 3. Buat GAS project baru (atau pakai existing)
#    Opsi A: Buat baru dari CLI
clasp create --type sheets --title "CanvaPOS"
#    Opsi B: Pakai project existing
#    Buka script editor → Settings → Script ID → copas ke .clasp.json

# 4. Push semua file ke GAS
clasp push

# 5. Buka di browser
clasp open

# 6. Di Apps Script editor, jalankan setupPOS()
```

### Manual (copy-paste — alternatif)

1. Buka Google Sheets baru → **Extensions → Apps Script**
2. Copy isi setiap file dari `src/` ke editor (buat file baru per modul)
3. Simpan, jalankan `setupPOS()`

> Jika `setupPOS()` timeout, jalankan fungsi satu per satu (lihat menu POS → Setup Partial).

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

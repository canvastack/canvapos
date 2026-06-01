# CanvaPOS

**Point-of-Sale System untuk Google Sheets** — Dirancang untuk usaha minuman skala kecil (Pop Ice blender, kopi tubruk, es teh) dengan manajemen stok berbasis Bill of Materials (BOM) dan laporan laba/rugi multi-level. **14 file modular, deployed ke production.**

## Fitur

- **POS Entry** — Input pesanan dengan dropdown varian yang cerdas: produk dengan stok bahan tidak mencukupi otomatis tidak muncul di dropdown
- **Hitung Otomatis** — Harga base (Rp 5.000) + topping (Rp 1.000/jenis) terhitung real-time
- **Transaksi Log** — Setiap pesanan tersimpan dengan timestamp, kasir, dan detail lengkap
- **Manajemen Stok** — Stok bahan terpotong otomatis berdasarkan BOM setiap transaksi
- **Laporan Laba/Rugi 5-Level** — Revenue → HPP breakdown (4 kategori BOM) → Gross Profit → OPEX → EBITDA → Depresiasi → EBIT → Pajak → Net Profit
- **Manajemen Kas** — Petty Cash (PC) & Uang Belanja (UB) dengan saldo otomatis
- **Catatan Pengeluaran** — Sync otomatis ke stok + indikator status stok
- **Dialog Interaktif** — Pilihan topping, date picker, tambah resep, tambah aset via popup
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

## Struktur Sheet (11 Sheets)

| Sheet | Fungsi | Fase |
|-------|--------|------|
| **Panduan** | Dashboard + Manual pengguna | 🔝 Referensi |
| **Bahan** | Master data bahan baku (harga, ukuran) | ⚙️ Config |
| **Resep** | Bill of Materials — komposisi setiap produk | ⚙️ Config |
| **Kas** | Tracking Petty Cash & Uang Belanja | 🌅 Pre-Op |
| **Stock** | Monitoring stok bahan + restock alert | 🌅 Pre-Op |
| **Pengeluaran** | Catatan pembelian & biaya operasional | 🏪 Operasi |
| **POS** | Input pesanan utama | 🏪 Operasi |
| **Transaksi** | Log history transaksi | 🏪 Operasi |
| **Pendapatan** | Laporan laba/rugi 5-level (harian & bulanan) | 📊 Reports |
| **Aset** | Daftar aset tetap & penyusutan | 📊 Reports |
| **Audit** | Hidden log aktivitas (90-day retention) | 🔧 System |

## Menu `🧋 POS`

| Menu | Fungsi |   
|---|---|   
| 💾 Simpan Transaksi | Simpan pesanan ke log + update stok |   
| ➕ Add Row | Tambah baris order baru |   
| 🍬 Pilih Topping (baris aktif) | Dialog checklist topping untuk baris aktif |   
| 🗑 Clear POS | Reset semua baris order |   
| ♻ Safe Clear (backup dulu) | Backup data lalu clear POS |   
| ↩ Restore POS dari Backup | Pulihkan data dari backup terakhir |   
| 💸 Simpan & Sync Stok (Pengeluaran) | Simpan pengeluaran → update stok |
| ➕ Add Row Pengeluaran | Navigasi ke baris kosong + validation otomatis |
| 🔧 Fix Formula Total (atasi #REF!) | Perbaiki konflik formula Total |
| 💸 Sinkronisasi Dropdown Pengeluaran | Refresh dropdown di sheet Pengeluaran |
| 🔄 Refresh Laporan Pendapatan | Update laporan pendapatan |
| 📊 Refresh Dashboard | Update dashboard di Panduan |
| 🗑 Hapus Baris Aktif (POS) | Hapus baris order tertentu |
| ➕ Tambah Resep / BOM | Dialog tambah resep baru |
| 💰 Top Up PC ke Rp 100.000 | Top up PC (closing harian) |
| 🛒 Top Up UB (jika < Rp 10.000) | Top up UB otomatis |
| 📅 Init Saldo Awal PC & UB | Set saldo awal harian |
| 📅 Pilih Tanggal (cell aktif) | Date picker untuk cell aktif |
| 📦 Tambah Aset Tetap | Dialog tambah aset tetap baru |
| 📉 Posting Penyusutan | Posting depresiasi bulan ini |
| 🔄 Sync Modal Awal → Aset + Kas | Migrasi data Modal Awal ke Aset tetap & Kas |
| 🔧 Setup Ulang (reset semua) | Reset semua sheet |
| 🔒 Proteksi Semua Sheet | Proteksi formula & header |
| 🔓 Unprotect Semua Sheet | Hapus semua proteksi |
| 📀 Backup Sekarang | Backup spreadsheet + timestamp |
| ⏰ Atur Backup Otomatis | Daily backup trigger |
| 🚀 Onboarding Wizard | Panduan setup 4 langkah |
| 📦 Migrasi Stok (Awal → Masuk) | Migrasi ke sistem Stok Masuk kumulatif |
| 🏷️ Refresh Named Ranges | Update ukuran named range |
| 🌐 Set Environment… | Pilih production/staging/development |
| ⚡ Install Auto-Fix Trigger | Pasang trigger onEdit *(auto-installed saat setupPOS)* |   

## Teknologi

- **Bahasa:** JavaScript (Google Apps Script, V8)
- **UI:** Google Sheets cells + HtmlService (HTML/CSS/JS dialogs)
- **Storage:** Spreadsheet rows + PropertiesService + CacheService
- **Trigger:** `onEdit`, `onSelectionChange`, `onOpen`
- **Concurrency:** LockService wrapper (`withLock()`) — aman multi-user
- **Deploy:** clasp push — 14 file modular ke GAS project
- **Harga:** Rp 5.000/cup, Rp 1.000/jenis topping

## Script ID

```
18Eld0ZbczRsWqIxXYhK3y0DHZm2FKUO9jQ8Dvc-UI9DfHKAmItiElags
```

## Lisensi

Hak cipta CanvaStack.ID

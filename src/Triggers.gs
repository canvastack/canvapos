// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Triggers.gs (Event Handlers & Menu)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// CUSTOM MENU — muncul otomatis saat file dibuka
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Buat custom menu POS saat spreadsheet dibuka.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧋 POS")
    .addItem("💾 Simpan Transaksi", "simpanTransaksi")
    .addItem("➕ Add Row", "addRowPOS")
    .addItem("🍬 Pilih Topping (baris aktif)", "pilihTopping")
    .addSeparator()
    .addItem("🗑 Clear POS", "clearPOS")
    .addItem("♻ Safe Clear (backup dulu)", "safeClearPOS")
    .addItem("↩ Restore POS dari Backup", "restorePOSFromBackup")
    .addSeparator()
    .addItem("💸 Simpan & Sync Stok (Pengeluaran)", "simpanPengeluaran")
    .addItem("➕ Add Row Pengeluaran", "addRowPengeluaran")
    .addItem("🔧 Fix Formula Total (atasi #REF!)", "fixPengeluaranFormulas")
    .addItem("📥 Import Data Mei", "importDataMei")
    .addItem('💸 Sinkronisasi Dropdown Pengeluaran', 'syncDropdownPengeluaran')
    .addItem("🔄 Refresh Laporan Pendapatan", "refreshLaporan")
    .addItem("📊 Refresh Dashboard", "refreshDashboard")
    .addSeparator()
    .addItem("🗑 Hapus Baris Aktif (POS)", "deleteRowPOS")
    .addSeparator()
    .addItem("➕ Tambah Resep / BOM", "showTambahResepDialog")
    .addItem("🔄 Sync Data Resep", "syncResepData")
    .addSeparator()
    .addSeparator()
    .addItem("💰 Top Up PC ke Rp 100.000", "topUpPC")
    .addItem("🛒 Top Up UB (jika < Rp 10.000)", "topUpUB")
    .addItem("📅 Init Saldo Awal PC & UB", "initSaldoKas")
    .addItem("📅 Pilih Tanggal (cell aktif)", "showDatePickerGeneric")
    .addSeparator()
    .addItem("📦 Tambah Aset Tetap", "addAset")
    .addItem("📉 Posting Penyusutan", "postingDepresiasi")
    .addSeparator()
    .addItem("🔧 Setup Ulang (reset semua)", "setupPOS")
    .addItem("🔒 Proteksi Semua Sheet", "protectAll")
    .addItem("🔓 Unprotect Semua Sheet", "unprotectAll")
    .addSeparator()
    .addItem("📀 Backup Sekarang", "backupSpreadsheet")
    .addItem("⏰ Atur Backup Otomatis", "setupBackupTrigger")
    .addItem("📂 Lihat Backup", "listBackups")
    .addItem("♻ Restore dari Backup", "restoreBackup")
    .addSeparator()
    .addItem("🚀 Onboarding Wizard", "showOnboardingWizard")
    .addSeparator()
    .addItem("🏷️ Refresh Named Ranges", "refreshNamedRanges")
    .addSeparator()
    .addItem("📦 Migrasi Stok (Stok Awal → Stok Masuk)", "migrasiStok")
    .addSeparator()
    .addItem("🌐 Set Environment…", "showEnvPicker")
    .addSeparator()
    .addItem("⚡ Install Auto-Fix Trigger", "setupOnEditTrigger")
    .addToUi();

  // Bersihkan log audit > 90 hari
  cleanAuditLog();
  // Bersihkan backup lama
  cleanBackups();
  // Refresh dashboard
  refreshDashboard();
  // Refresh named ranges (fix width BAHAN_Lookup etc.)
  refreshNamedRanges();

  // Pasang date validation di kolom A Pengeluaran
  var C = getC();
  var shP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pengeluaran");
  if (shP) {
    shP.getRange("A4:A203").setDataValidation(
      SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
        .setHelpText("Double-click untuk pilih tanggal").build()
    );
    // Kembalikan row 2 ke merge E2:H2 (bersihin checkbox bekas)
    try { shP.getRange("E2:H2").breakApart(); } catch(_) {}
    shP.getRange("E2:H2").merge()
      .setValue("ℹ Isi baris baru di bawah → klik menu POS → Simpan & Sync Stok")
      .setBackground(C.LIGHT).setFontColor(C.DARK).setFontSize(10)
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    shP.getRange("G2").clearAll();
    shP.getRange("H2").clearAll();
    shP.setRowHeight(2, 30);
  }
}

// ── onEdit utama: routing ke handler yang sesuai ─────────────────────────
/**
 * Trigger utama onEdit — routing ke handler sesuai sheet.
 * @param {Object} e - Event object
 */
function onEdit(e) {
  safeExecute("onEdit", function() {
  if (!e || !e.source) return;
  var shName = e.source.getActiveSheet().getName();

  // ── 1. Jika yang diedit adalah tab Bahan, langsung sinkronkan POS ──
  if (shName === "Bahan") { 
    clearDynamicCache();
    clearHPPLookupCache();
    syncDropdownPOS(); 
    return; 
  }

  if (shName === "Resep") {
    clearDynamicCache();
    clearHPPLookupCache();
    syncDropdownPOS();
    return;
  }

  if (shName === "Stock") {
    clearDynamicCache();
    syncDropdownPOS();
    return;
  }

  if (shName === "Pengeluaran") { onEditPengeluaran(e); return; }
  if (shName !== "POS") return;

  // POS: auto-fix GRAND TOTAL kalau ada row yang dihapus manual
  var props = PropertiesService.getDocumentProperties();
  var savedGTRow = parseInt(props.getProperty("POS_GRAND_TOTAL_ROW") || "0");
  if (savedGTRow <= 0) return;

  var sh = e.source.getActiveSheet();
  var lastRow = sh.getLastRow();

  if (savedGTRow > lastRow) {
    for (var r = POS_START_ROW; r <= lastRow; r++) {
      if (sh.getRange(r, 8).getFormula().indexOf("SUM(H") >= 0) {
        props.setProperty("POS_GRAND_TOTAL_ROW", String(r));
        _renumberAndFixPOS(sh);
        break;
      }
    }
  } else if (sh.getRange(savedGTRow, 8).getFormula().indexOf("SUM(H") < 0) {
    for (var r2 = POS_START_ROW; r2 <= lastRow; r2++) {
      if (sh.getRange(r2, 8).getFormula().indexOf("SUM(H") >= 0) {
        props.setProperty("POS_GRAND_TOTAL_ROW", String(r2));
        _renumberAndFixPOS(sh);
        break;
      }
    }
  }
  });
}

// ── Fungsi untuk mereset dropdown di Tab POS secara instan ──
/**
 * Sinkronkan dropdown varian di POS dari data Resep.
 */
function syncDropdownPOS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPOS = getSheet(SHEET.POS);
  if (!shPOS) return;
  
  // Ambil data terbaru dari tab Resep (menu items)
  var varianRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getVarianList(), true)
    .setAllowInvalid(false).build();
    
  var lastRow = shPOS.getLastRow();
  var startRow = POS_START_ROW || 7; 
  
  if (lastRow >= startRow) {
    // Perbarui seluruh baris di kolom Product Variant (Kolom B) secara silent
    shPOS.getRange(startRow, 2, lastRow - startRow + 1).setDataValidation(varianRule);
  }
}

/**
 * Sinkronkan dropdown item di Pengeluaran dari Stock.
 */
function syncDropdownPengeluaran() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPen = getSheet(SHEET.PENGELUARAN);
  var shStock = getSheet(SHEET.STOCK);
  
  if (!shPen || !shStock) return;
  
  var lastStockRow = shStock.getLastRow();
  var rangeStock = shStock.getRange('B2:B' + lastStockRow); 
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeStock, true)
    .setAllowInvalid(false)
    .build();
    
  var lastRow = shPen.getLastRow();
  if (lastRow >= 4) {
    shPen.getRange("C4:C" + lastRow).setDataValidation(rule);
  } else {
    shPen.getRange("C4:C").setDataValidation(rule);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
/**
 * Pasang trigger onEdit jika belum ada.
 */
function setupOnEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onEdit") {
      SpreadsheetApp.getUi().alert("⚠ Trigger onEdit sudah terinstall.");
      return;
    }
  }
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert("✅ Trigger onEdit berhasil dipasang!\n\nFungsi auto-fix POS dan auto-fill Pengeluaran akan berjalan otomatis.");
}

// ═══════════════════════════════════════════════════════════════════════════
// TRIGGER — onSelectionChange: auto show date picker di sheet tertentu
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Trigger onSelectionChange — auto date picker di Pengeluaran.
 * @param {Object} e - Event object
 */
function onSelectionChange(e) {
  safeExecute("onSelectionChange", function() {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var shName = sh ? sh.getName() : "";

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Pengeluaran: kolom Tanggal (A), baris data (4+), cell kosong → auto isi tanggal hari ini
  if (shName === "Pengeluaran" && col === 1 && row >= 4 && e.range.getValue() === "") {
    e.range.setValue(new Date()).setNumberFormat("DD/MM/YYYY");
    return;
  }
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// MIGRASI — Konversi Stok Awal → Stok Masuk (kumulatif)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Migrasi satu kali: back-calculate Stok Masuk = Sisa Stok + Terjual,
 * lalu ubah Sisa Stok jadi formula = D - E.
 * Hanya jalan kalau kolom D masih namanya "Stok Awal" (indikasi belum migrasi).
 */
function migrasiStok() {
  var sh = getSheet(SHEET.STOCK);
  if (!sh) return;
  var ui = SpreadsheetApp.getUi();
  if (!confirmAction(
    "📦 Migrasi Stok Awal → Stok Masuk\n\n" +
    "• Kolom D (Stok Awal) → Stok Masuk (kumulatif)\n" +
    "• Kolom F (Sisa Stok) → formula = D - E\n" +
    "• Data existing akan di-back-calculate: Stok Masuk = Sisa + Terjual\n\n" +
    "Lanjutkan?", "📦 Migrasi Stok?"
  )) return;

  // Cek apakah sudah pernah migrasi (header berubah)
  var headerD = sh.getRange("D1").getValue();
  if (headerD === "Stok Masuk") {
    ui.alert("✅ Stock sudah menggunakan sistem Stok Masuk kumulatif.\nTidak perlu migrasi ulang.");
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert("Stock kosong."); return; }

  var data = sh.getRange("A2:H" + lastRow).getValues();
  var updated = 0;

  for (var i = 0; i < data.length; i++) {
    var r = i + 2;
    var stokAwal = Number(data[i][3]) || 0;  // D = Stok Awal (old)
    var terjual  = Number(data[i][4]) || 0;  // E = Terjual
    var sisa     = Number(data[i][5]) || 0;  // F = Sisa Stok (old value)
    var minStok  = Number(data[i][6]) || 0;  // G = Min. Stok

    if (!data[i][1]) continue; // skip empty rows

    // Back-calculate: Stok Masuk = existing sisa + terjual
    // (Kalau user pernah isi manual Stok Awal, kita pake nilai yg lebih besar)
    var stokMasuk = sisa + terjual;
    if (stokAwal > stokMasuk) stokMasuk = stokAwal; // prefer manual value jk lebih besar

    // Tulis Stok Masuk (D)
    sh.getRange(r, 4).setValue(stokMasuk);

    // Sisa Stok (F) = formula = D - E
    sh.getRange(r, 6).setFormula(F("=IF({d}=0,\"\",{d}-{e})", {d: 'D'+r, e: 'E'+r}))
      .setNumberFormat("#,##0");

    // Status formula sudah pernah di-set buildStock — biarkan

    updated++;
  }

  // Update header
  sh.getRange("D1").setValue("Stok Masuk");
  sh.getRange("A1:H1").setFontWeight("bold"); // keep bold

  SpreadsheetApp.flush();
  ui.alert("✅ Migrasi selesai!\n\n" + updated + " item diperbarui.\n" +
    "Kolom D = Stok Masuk (kumulatif)\n" +
    "Kolom F = Sisa Stok (formula = D - E)");
  auditLog("Migrasi Stok", updated + " item");
}

// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Pengeluaran.gs (Expense Tracking & Stock Sync)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — PENGELUARAN (input langsung di sheet, dropdown dari Bahan)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet Pengeluaran (expense tracking).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function buildPengeluaran(ss) {
  var C = getC();
  var sh = ss.insertSheet("Pengeluaran");
  setSheetFont(sh);
  sh.setTabColor("#E67E22");

  // ── Judul ──────────────────────────────────────────────────────────────
  sh.getRange("A1:H1").merge()
    .setValue("💸 Catatan Pengeluaran")
    .setBackground("#E67E22").setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 40);

  // ── Tombol Simpan & Sync ───────────────────────────────────────────────
  sh.getRange("A2:D2").merge()
    .setValue("💾 Simpan & Sync Stok")
    .setBackground(C.GREEN).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("E2:F2").merge()
    .setValue("ℹ Isi baris → Simpan & Sync Stok")
    .setBackground(C.LIGHT).setFontColor(C.DARK).setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  // Checkbox 📅 — centang buat pilih tanggal
  sh.getRange("G2").setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build()
  ).setValue(false).setBackground("#E67E22").setHorizontalAlignment("center");
  sh.getRange("H2").setValue("📅")
    .setBackground("#E67E22").setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 30);

  // ── Header ────────────────────────────────────────────────────────────
  var headers = ["Tanggal","Kategori","Nama Item","Satuan","Jumlah","Harga Satuan","Total","Status Stok"];
  sh.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground("#E67E22").setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(3, 28);

  // ── Dropdown Kategori untuk 50 baris ke depan ─────────────────────────
  var KATEGORI_LIST = ["Bahan Utama","Topping","Kemasan","Bahan Pendukung","Operasional","Lain-lain"];
  var katRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KATEGORI_LIST, true)
    .setAllowInvalid(false).build();
  sh.getRange(4, 2, 200, 1).setDataValidation(katRule);

  // ── Formula Total otomatis untuk 200 baris ────────────────────────────
  for (var i = 0; i < 200; i++) {
    var r = i + 4;
    sh.getRange(r, 7).setFormula("=IF(E"+r+"*F"+r+"=0,\"\",E"+r+"*F"+r+")")
      .setNumberFormat('"Rp "#,##0');
    sh.getRange(r, 8).setValue("").setFontSize(10);
  }

  // ── Format kolom Tanggal & Harga (sebelum sample data) ──────────────
  sh.getRange("A4:A203").setNumberFormat("DD/MM/YYYY");
  sh.getRange("F4:F203").setNumberFormat('"Rp "#,##0');

  // ── Data kosong — isi manual ────────────────────────────────────────

  // ── Date picker validation ────────────────────────────────────────
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Klik untuk pilih tanggal dari kalender")
    .build();
  sh.getRange("A4:A203").setDataValidation(dateRule);

  [110,130,180,90,80,130,130,120].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(3);
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — SIMPAN & SYNC STOK dari sheet Pengeluaran
// Baca baris yang belum di-sync (kolom H kosong), update stok, tandai selesai
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Sync data Pengeluaran ke Stock (proses baris baru).
 */
function simpanPengeluaran() {
  if (!confirmAction("Yakin mau sync data Pengeluaran ke Stock? Semua baris baru akan diproses.", "💸 Sync ke Stock?")) return;

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var C    = getC();
  var shPen = getSheet(SHEET.PENGELUARAN);
  var shStk = getSheet(SHEET.STOCK);

  if (!shPen || !shStk) {
    ui.alert("Sheet Pengeluaran atau Stock tidak ditemukan.");
    return;
  }

  var lastRow = shPen.getLastRow();
  if (lastRow < 4) { ui.alert("Belum ada data pengeluaran."); return; }

  // ══════════════════════════════════════════════════════════════════════
  // PHASE 1 — Read & Validate (no writes)
  // ══════════════════════════════════════════════════════════════════════
  var KATEGORI_STOK = ["Bahan Utama","Topping","Kemasan","Bahan Pendukung"];

  var PB = COL.BAHAN, PP = COL.PENGELUARAN, PS = COL.STOCK;
  var shBahan = getSheet(SHEET.BAHAN);
  var bahanPackMap = {};
  if (shBahan) {
    var bahanData = shBahan.getDataRange().getValues();
    for (var bi = 1; bi < bahanData.length; bi++) {
      var bNama = String(bahanData[bi][PB.NAMA]).trim().toLowerCase();
      var bPack = Number(bahanData[bi][PB.PACK]) || 1;
      bahanPackMap[bNama] = bPack;
    }
  }

  var penData = shPen.getRange(4, 1, lastRow - 3, 8).getValues();
  var stockLastRow = shStk.getLastRow();
  var stockData = stockLastRow > 1
    ? shStk.getRange(2, 1, stockLastRow - 1, 7).getValues()
    : [];

  // ── Validasi unit consistency (P2.5) ────────────────────────────────
  var warnings = [];
  penData.forEach(function(row) {
    var namaItem = String(row[PP.NAMA]).trim().toLowerCase();
    var unitPen  = String(row[PP.SATUAN]).trim();
    if (!namaItem || !unitPen) return;
    for (var si = 0; si < stockData.length; si++) {
      var namaStok = String(stockData[si][PS.NAMA]).trim().toLowerCase();
      if (namaStok === namaItem) {
        var unitStok = String(stockData[si][PS.SATUAN]).trim();
        if (unitStok && unitPen.toLowerCase() !== unitStok.toLowerCase()) {
          warnings.push("⚠ Unit \"" + row[PP.NAMA] + "\": Pengeluaran = " + unitPen + ", Stock = " + unitStok);
        }
        break;
      }
    }
  });

  // ══════════════════════════════════════════════════════════════════════
  // PHASE 2 — Compute changes in memory & PHASE 3 — Write (inside Lock)
  // ══════════════════════════════════════════════════════════════════════
  var synced = 0, skipped = 0;
  var syncActions = []; // {row, status, stockOps} untuk batch write

  penData.forEach(function(row, i) {
    var r = i + 4;
    var tanggal  = row[PP.TGL];
    var kategori = String(row[PP.KATEGORI]).trim();
    var namaItem = String(row[PP.NAMA]).trim();
    var jumlah   = Number(row[PP.JUMLAH]) || 0;
    var status   = String(row[PP.STATUS]).trim();

    if (!namaItem || !tanggal) return;
    if (status === "✓ Synced") { skipped++; return; }
    synced++;

    var bg = applyZebraRow(shPen, r, i, 8);
    shPen.getRange(r, COLx(PP.HARGA)).setNumberFormat('"Rp "#,##0');

    var syncedAction = {row: r, namaItem: namaItem};
    if (KATEGORI_STOK.indexOf(kategori) >= 0 && jumlah > 0) {
      var packSize = bahanPackMap[namaItem.toLowerCase()] || 1;
      var qtyStok  = jumlah * packSize;

      var found = false;
      for (var j = 0; j < stockData.length; j++) {
        var namaBahan = String(stockData[j][PS.NAMA]).trim().toLowerCase();
        if (namaBahan === namaItem.toLowerCase()) {
          syncedAction.stockRow = j + 2;
          syncedAction.qtyStok = qtyStok;
          found = true;
          break;
        }
      }
      if (!found) {
        syncedAction.newItem = {
          kategori: kategori, namaItem: namaItem,
          satuan: String(row[PP.SATUAN]).trim(), qtyStok: qtyStok
        };
      }
    }
    syncActions.push(syncedAction);
  });

  if (synced === 0 && skipped > 0) {
    ui.alert("Semua baris sudah di-sync (" + skipped + " baris dilewati).");
    return;
  }

  // PHASE 3 — Write batch (inside LockService)
  withLock(15000, function() {
    syncActions.forEach(function(action) {
      var r = action.row;
      if (action.newItem) {
        var newRow = shStk.getLastRow() + 1;
        shStk.getRange(newRow, COLx(PS.KATEGORI)).setValue(action.newItem.kategori);
        shStk.getRange(newRow, COLx(PS.NAMA)).setValue(action.newItem.namaItem);
        shStk.getRange(newRow, COLx(PS.SATUAN)).setValue(action.newItem.satuan);
        shStk.getRange(newRow, COLx(PS.STOK_AWAL)).setValue(action.newItem.qtyStok);
        shStk.getRange(newRow, COLx(PS.TERJUAL)).setValue(0);
        shStk.getRange(newRow, COLx(PS.SISA)).setValue(action.newItem.qtyStok);
        shStk.getRange(newRow, COLx(PS.MIN)).setValue(1);
        shStk.getRange(newRow, COLx(PS.STATUS)).setFormula(F("=IF(F{row}<=G{row},\"⚠ RESTOCK\",\"✓ OK\")", {row: newRow}));
        var rBg = (newRow % 2 === 0) ? C.LIGHT : C.WHITE;
        styleData(shStk.getRange(newRow, 1, 1, 8), rBg);
        shStk.setRowHeight(newRow, 24);
      } else if (action.stockRow) {
        var stokLama = shStk.getRange(action.stockRow, COLx(PS.SISA)).getValue() || 0;
        shStk.getRange(action.stockRow, COLx(PS.SISA)).setValue(stokLama + action.qtyStok);
      }
      shPen.getRange(r, COLx(PP.STATUS)).setValue("✓ Synced")
        .setFontColor(C.GREEN).setFontWeight("bold").setFontSize(10);
    });
    SpreadsheetApp.flush();
  });

  refreshLaporan();
  refreshNamedRanges();
  auditLog("Simpan Pengeluaran", synced + " item sync, " + skipped + " skipped");

  var msg = "✅ Selesai!\n\n" +
    "Baris diproses : " + synced + "\n" +
    "Sudah di-sync  : " + skipped + " (dilewati)";
  if (synced > 0) {
    msg += "\n\nItem baru otomatis ditambahkan ke Stock.\n" +
      "Jumlah dikonversi otomatis berdasarkan pack size dari Bahan.";
  }
  if (warnings.length > 0) {
    msg += "\n\n⚠ Peringatan unit:\n" + warnings.join("\n");
  }
  try { ui.alert(msg); } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════════════
// TRIGGER — onEdit Pengeluaran: dropdown + auto-fill + status stok
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Trigger onEdit untuk auto-fill di sheet Pengeluaran.
 * @param {Object} e - Event object
 */
function onEditPengeluaran(e) {
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  var shPen = e.source.getActiveSheet();
  
  // ─── 1. KONFIGURASI KOLOM TARGET (TAB PENGELUARAN) ───
  var COL_NAMA_ITEM = 3;  // Kolom C
  var COL_SATUAN    = 4;  // Kolom D
  var COL_JUMLAH    = 5;  // Kolom E
  var COL_HARGA     = 6;  // Kolom F
  var COL_TOTAL     = 7;  // Kolom G
  var COL_STATUS    = 8;  // Kolom H
  var START_ROW     = 4;

  var COL_KATEGORI  = 2;  // Kolom B

  // ── Handle checkbox 📅 di baris 2 → date picker ──
  if (row === 2 && col === 7 && range.getValue() === true) {
    if (typeof showDatePickerGeneric === 'function') {
      showDatePickerGeneric();
    }
    range.setValue(false);
    return;
  }

  if (row < START_ROW) return;

  // ── Handle Kategori change → dynamic Nama Item dropdown ──
  if (col === COL_KATEGORI) {
    _updateNamaItemDropdown(e.source, shPen, row);
    return;
  }

  // ── Handle Nama Item change → auto-fill ──
  if (col !== COL_NAMA_ITEM) return;

  var itemName = range.getValue();
  if (!itemName) {
    shPen.getRange(row, COL_HARGA, 1, 5).clearContent();
    return;
  }

  var ss = e.source;
  var shBahan = getSheet(SHEET.BAHAN);
  var shStock = getSheet(SHEET.STOCK);

  if (!shBahan || !shStock) return;

  var bahanData = shBahan.getDataRange().getValues();
  var stockData = shStock.getDataRange().getValues();

  // ─── 2. CONFIG INDEX SOURCE (TAB BAHAN & STOCK) ───
  var IDX_BAHAN_NAMA   = 1; // Kolom B (Nama Bahan)
  var IDX_BAHAN_UKURAN = 2; // Kolom C (Ukuran/Pack)
  var IDX_BAHAN_HARGA  = 4; // Kolom E (Harga Beli per pack)

  var IDX_STOCK_NAMA   = 1; // Kolom B (Nama Item)
  var IDX_STOCK_SISA   = 5; // Kolom F (Sisa Stok)
  var IDX_STOCK_MIN    = 6; // Kolom G (Min. Stok)

  var hargaItem = 0;
  var satuanItem = "—";
  var sisaStock = 0;
  var minStok = 0;
  var itemFoundInBahan = false;
  var itemFoundInStock = false;

  // CARI DATA UTAMA DARI TAB BAHAN
  for (var i = 1; i < bahanData.length; i++) {
    if (bahanData[i][IDX_BAHAN_NAMA] === itemName) {
      hargaItem  = bahanData[i][IDX_BAHAN_HARGA];
      satuanItem = bahanData[i][IDX_BAHAN_UKURAN];
      itemFoundInBahan = true;
      break;
    }
  }

  // CARI DATA MONITORING DARI TAB STOCK
  for (var j = 1; j < stockData.length; j++) {
    if (stockData[j][IDX_STOCK_NAMA] === itemName) {
      sisaStock = stockData[j][IDX_STOCK_SISA];
      minStok   = stockData[j][IDX_STOCK_MIN];
      itemFoundInStock = true;
      break;
    }
  }

  // INJEKSI DATA KE TAB PENGELUARAN
  if (itemFoundInBahan) {
    var hargaBersih = 0;
    if (typeof hargaItem === 'string') {
      hargaBersih = parseInt(hargaItem.replace(/[^\d]/g, '')) || 0;
    } else {
      hargaBersih = hargaItem || 0;
    }

    shPen.getRange(row, COL_SATUAN).setValue(satuanItem);
    shPen.getRange(row, COL_HARGA).setValue(hargaBersih);
    shPen.getRange(row, COL_JUMLAH).setValue(1);

    var cellHarga  = shPen.getRange(row, COL_HARGA).getA1Notation();
    var cellJumlah = shPen.getRange(row, COL_JUMLAH).getA1Notation();
    shPen.getRange(row, COL_TOTAL).setFormula("=IFERROR(" + cellHarga + "*" + cellJumlah + ", 0)");
  }
  
  // INDIKATOR STATUS STOCK
  if (itemFoundInStock) {
    var statusMsg = "Sisa: " + sisaStock;
    var isKritis = sisaStock <= minStok;
    
    shPen.getRange(row, COL_STATUS).setValue(statusMsg)
         .setFontColor(isKritis ? "#E74C3C" : "#27AE60")
         .setFontWeight("bold").setFontSize(10);
  } else {
    shPen.getRange(row, COL_STATUS).setValue("— tidak ada di Stock")
         .setFontColor("#888888").setFontSize(10);
  }
}

// ── Helper: update Nama Item dropdown berdasarkan Kategori ───────────────
function _updateNamaItemDropdown(ss, shPen, row) {
  var kategori = shPen.getRange(row, 2).getValue();
  var COL_NAMA_ITEM = 3;

  // Item operasional untuk kategori yang tidak ada di Bahan
  var ITEM_OPERASIONAL = [
    "Listrik","Air","Sewa Tempat","Sewa Tempat Penitipan",
    "Karyawan","Transportasi","Gas LPG","Internet","ATK","Kebersihan","Lainnya"
  ];

  if (kategori === "Operasional") {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(ITEM_OPERASIONAL, true)
      .setAllowInvalid(true) // allow user to type custom
      .build();
    shPen.getRange(row, COL_NAMA_ITEM).setDataValidation(rule).clearContent();
    return;
  }

  if (kategori === "Lain-lain") {
    // Lain-lain: no validation, free text
    shPen.getRange(row, COL_NAMA_ITEM).setDataValidation(null).clearContent();
    return;
  }

  // Cari item dari Bahan yang sesuai kategori
  var shBahan = getSheet(SHEET.BAHAN);
  if (!shBahan) return;

  var bahanData = shBahan.getDataRange().getValues();
  var filtered = [];
  for (var i = 1; i < bahanData.length; i++) {
    if (String(bahanData[i][0]).trim() === kategori) {
      filtered.push(bahanData[i][1]);
    }
  }

  if (filtered.length > 0) {
    var rule2 = SpreadsheetApp.newDataValidation()
      .requireValueInList(filtered, true)
      .setAllowInvalid(true)
      .build();
    shPen.getRange(row, COL_NAMA_ITEM).setDataValidation(rule2).clearContent();
  } else {
    // Kategori gak dikenal → free text
    shPen.getRange(row, COL_NAMA_ITEM).setDataValidation(null).clearContent();
  }
}

// ── Helper: tampilkan estimasi sisa stok di kolom H ──────────────────────
function _updateStatusStok(shPen, shStk, row, namaItem, jumlahBeli) {
  if (!shStk) return;
  var stockData = shStk.getRange(2, 1, shStk.getLastRow() - 1, 8).getValues();
  for (var j = 0; j < stockData.length; j++) {
    if (String(stockData[j][1]).trim().toLowerCase() === namaItem.toLowerCase()) {
      var stokAwal = Number(stockData[j][3]) || 0; // kolom D = Stok Awal
      var terjual  = Number(shStk.getRange(j + 2, 5).getValue()) || 0; // kolom E = Terjual
      var minStok  = Number(stockData[j][6]) || 0; // kolom G = Min. Stok
      // Estimasi stok setelah pembelian ini
      var sisaEstimasi = stokAwal + jumlahBeli - terjual;
      var C = getC();
      var status = sisaEstimasi <= minStok
        ? "⚠ Estimasi: " + sisaEstimasi + " (hampir habis)"
        : "✓ Estimasi: " + sisaEstimasi;
      shPen.getRange(row, 8).setValue(status)
        .setFontColor(sisaEstimasi <= minStok ? C.RED : C.GREEN)
        .setFontWeight("bold").setFontSize(10);
      return;
    }
  }
  // Tidak ditemukan di stock
  shPen.getRange(row, 8).setValue("— tidak ada di Stock")
    .setFontColor("#888888").setFontSize(10);
}

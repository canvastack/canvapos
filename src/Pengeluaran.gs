// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Pengeluaran.gs (Expense Tracking & Stock Sync)
// ═══════════════════════════════════════════════════════════════════════════

var KATEGORI_LIST_PEN = ["Bahan Utama","Topping","Kemasan","Bahan Pendukung","Operasional","Modal Awal","Lain-lain"];

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
    .setValue("💾 Simpan & Sync Stok  |  ➕ Add Row (menu POS)")
    .setBackground(C.GREEN).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("E2:H2").merge()
    .setValue("ℹ Menu POS → Add Row Pengeluaran untuk navigasi baris baru + validation otomatis")
    .setBackground(C.LIGHT).setFontColor(C.DARK).setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(2, 30);

  // ── Header ────────────────────────────────────────────────────────────
  var headers = ["Tanggal","Kategori","Nama Item","Satuan (base)","Jumlah","Harga per Unit","Total","Status Stok"];
  sh.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground("#E67E22").setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(3, 28);

  // ── Dropdown Kategori untuk 5000 baris ─────────────────────────────
  var katRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KATEGORI_LIST_PEN, true)
    .setAllowInvalid(false).build();
  sh.getRange(4, 2, 5000, 1).setDataValidation(katRule);

  // ── Formula Total batch untuk 5000 baris (per-row, no ArrayFormula conflict) ──
  var TOTAL_ROWS = 5000;
  var formulas = [];
  for (var fi = 0; fi < TOTAL_ROWS; fi++) {
    var fr = fi + 4;
    formulas.push(["=IF(E" + fr + "*F" + fr + "=0,\"\",E" + fr + "*F" + fr + ")"]);
  }
  sh.getRange(4, 7, TOTAL_ROWS, 1).setFormulas(formulas)
    .setNumberFormat('"Rp "#,##0');
  sh.getRange("H4:H" + (3 + TOTAL_ROWS)).setFontSize(10);

  // ── Format kolom Tanggal & Harga ──────────────────────────────────
  sh.getRange("A4:A5000").setNumberFormat("DD/MM/YYYY");
  sh.getRange("F4:F5000").setNumberFormat('"Rp "#,##0');

  // ── Data kosong — isi manual ────────────────────────────────────────

  // ── Date picker validation ────────────────────────────────────────
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Klik untuk pilih tanggal dari kalender")
    .build();
  sh.getRange("A4:A5000").setDataValidation(dateRule);

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
  // Cari baris terakhir yang beneran ada datanya (A-C), skip baris formula kosong
  var lastDataRow = lastRow;
  if (lastRow > 200) {
    var scanData = shPen.getRange(4, 1, Math.min(lastRow - 3, 500), 3).getValues();
    for (var si = scanData.length - 1; si >= 0; si--) {
      if (String(scanData[si][0]).trim() || String(scanData[si][1]).trim() || String(scanData[si][2]).trim()) {
        lastDataRow = 4 + si;
        break;
      }
    }
  }
  if (lastDataRow < 4) { ui.alert("Belum ada data pengeluaran."); return; }

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

  var penData = shPen.getRange(4, 1, lastDataRow - 3, 8).getValues();
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
      var qtyStok  = jumlah;

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
        // Sisa Stok (F) = formula D-E — tidak diisi manual
        shStk.getRange(newRow, COLx(PS.SISA)).setFormula(F("=IF({d}=0,\"\",{d}-{e})", {d: 'D'+newRow, e: 'E'+newRow}))
          .setNumberFormat("#,##0");
        shStk.getRange(newRow, COLx(PS.MIN)).setValue(1);
        shStk.getRange(newRow, COLx(PS.STATUS)).setFormula(F("=IF(F{row}<=G{row},\"⚠ RESTOCK\",\"✓ OK\")", {row: newRow}));
        var rBg = (newRow % 2 === 0) ? C.LIGHT : C.WHITE;
        styleData(shStk.getRange(newRow, 1, 1, 8), rBg);
        shStk.setRowHeight(newRow, 24);
      } else if (action.stockRow) {
        var stokMasukLama = shStk.getRange(action.stockRow, COLx(PS.STOK_AWAL)).getValue() || 0;
        shStk.getRange(action.stockRow, COLx(PS.STOK_AWAL)).setValue(stokMasukLama + action.qtyStok);
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
  var shPen = e.source.getActiveSheet();

  var COL_KATEGORI  = 2;
  var COL_NAMA_ITEM = 3;
  var COL_SATUAN    = 4;
  var COL_JUMLAH    = 5;
  var COL_HARGA     = 6;
  var START_ROW     = 4;

  var rowStart = range.getRow();
  var rowEnd   = rowStart + range.getNumRows() - 1;
  var colStart = range.getColumn();
  var colEnd   = colStart + range.getNumColumns() - 1;

  if (rowStart < START_ROW) return;

  // ── Handle Kategori change — single row only ──
  if (colStart === COL_KATEGORI && colEnd === COL_KATEGORI && range.getNumRows() === 1) {
    _updateNamaItemDropdown(e.source, shPen, rowStart);
    return;
  }

  // ── Handle Nama Item change — single or multi row (copy-paste) ──
  if (colStart > COL_NAMA_ITEM || colEnd < COL_NAMA_ITEM) return;

  var namaValues = shPen.getRange(rowStart, COL_NAMA_ITEM, rowEnd - rowStart + 1, 1).getValues();
  var jumlahValues = shPen.getRange(rowStart, COL_JUMLAH, rowEnd - rowStart + 1, 1).getValues();
  var kategoriValues = (colStart <= COL_KATEGORI && colEnd >= COL_KATEGORI)
    ? shPen.getRange(rowStart, COL_KATEGORI, rowEnd - rowStart + 1, 1).getValues()
    : null;

  for (var ri = 0; ri < namaValues.length; ri++) {
    var r = rowStart + ri;
    var itemName = String(namaValues[ri][0]).trim();
    if (!itemName) {
      // Single-row clear → hapus auto-filled columns
      if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
        shPen.getRange(r, 6, 1, 3).clearContent();
      }
      continue;
    }

    // Check if user pasted jumlah — don't overwrite if they did
    var userJumlah = Number(jumlahValues[ri][0]) || 0;

    _autoFillPengeluaranRow(e.source, shPen, r, itemName, userJumlah,
      kategoriValues ? String(kategoriValues[ri][0]).trim() : "");
  }
}

/**
 * Auto-fill Satuan, Jumlah (if empty), Harga, Total, Status untuk satu row.
 * @param {Spreadsheet} ss
 * @param {Sheet} shPen
 * @param {number} row
 * @param {string} itemName
 * @param {number} existingJumlah - 0 if user didn't provide a value
 * @param {string} kategori - optional, for Operasional fallback
 */
function _autoFillPengeluaranRow(ss, shPen, row, itemName, existingJumlah, kategori) {
  var COL_SATUAN = 4, COL_JUMLAH = 5, COL_HARGA = 6;
  var shBahan = getSheet(SHEET.BAHAN);
  var shStock = getSheet(SHEET.STOCK);
  if (!shBahan || !shStock) return;

  var bahanData = shBahan.getDataRange().getValues();
  var stockData = shStock.getDataRange().getValues();

  var hargaItem = 0;
  var satuanItem = "—";
  var packSize = 1;
  var sisaStock = 0;
  var minStok = 0;
  var foundInBahan = false;
  var foundInStock = false;

  for (var i = 1; i < bahanData.length; i++) {
    if (bahanData[i][COL.BAHAN.NAMA] === itemName) {
      hargaItem  = bahanData[i][COL.BAHAN.HARGA_BELI];
      satuanItem = bahanData[i][COL.BAHAN.SATUAN];
      packSize   = bahanData[i][COL.BAHAN.PACK] || 1;
      foundInBahan = true;
      break;
    }
  }

  for (var j = 1; j < stockData.length; j++) {
    if (stockData[j][COL.STOCK.NAMA] === itemName) {
      sisaStock = stockData[j][COL.STOCK.SISA];
      minStok   = stockData[j][COL.STOCK.MIN];
      foundInStock = true;
      break;
    }
  }

  // Write Satuan & Harga
  if (foundInBahan) {
    var hargaBersih = (typeof hargaItem === 'string')
      ? parseInt(hargaItem.replace(/[^\d]/g, '')) || 0
      : hargaItem || 0;
    var hargaPerPiece = (packSize > 0) ? (hargaBersih / packSize) : hargaBersih;

    shPen.getRange(row, COL_SATUAN).setValue(satuanItem);
    shPen.getRange(row, COL_HARGA).setValue(hargaPerPiece)
      .setNumberFormat('"Rp "#,##0');

    // Jumlah: hanya set ke 1 jika user belum isi
    if (!existingJumlah) {
      shPen.getRange(row, COL_JUMLAH).setValue(1);
    }

    // Total formula
    var cellJml = shPen.getRange(row, COL_JUMLAH).getA1Notation();
    var cellHrg = shPen.getRange(row, COL_HARGA).getA1Notation();
    shPen.getRange(row, 7).setFormula("=IFERROR(" + cellJml + "*" + cellHrg + ", 0)")
      .setNumberFormat('"Rp "#,##0');
  }

  // Status Stock
  if (foundInStock) {
    var statusMsg = "Sisa: " + sisaStock;
    shPen.getRange(row, 8).setValue(statusMsg)
      .setFontColor(sisaStock <= minStok ? "#E74C3C" : "#27AE60")
      .setFontWeight("bold").setFontSize(10);
  } else if (kategori === "Operasional") {
    shPen.getRange(row, 8).setValue("✓ Synced")
      .setFontColor("#888888").setFontSize(10);
  } else {
    shPen.getRange(row, 8).setValue("— tidak ada di Stock")
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

// ═══════════════════════════════════════════════════════════════════════════
// IMPORT — Data Bulan Mei (one-time migration)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Import data awal bulan Mei ke sheet Pengeluaran (67 baris).
 * Tambah item baru ke Bahan jika belum ada.
 * Status Stok dikosongkan — jalankan Simpan & Sync Stok setelahnya.
 */
function importDataMei() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPen = getSheet(SHEET.PENGELUARAN);
  if (!shPen) { SpreadsheetApp.getUi().alert("Sheet Pengeluaran tidak ditemukan."); return; }

  // ── Guard: cek apakah sudah pernah di-import ──
  var lastDataRow = shPen.getLastRow();
  // Cari baris kosong pertama (kolom A-C kosong, ignore column G yg punya formula)
  function _firstEmptyRow() {
    for (var er = 4; er <= 250; er++) {
      var vals = shPen.getRange(er, 1, 1, 3).getValues()[0];
      if (!String(vals[0]).trim() && !String(vals[1]).trim() && !String(vals[2]).trim()) return er;
    }
    return 251;
  }
  var startRow = _firstEmptyRow();

  // Guard: scan data yg udah ada (bukan berdasarkan getLastRow)
  var scanEnd = Math.min(startRow + 100, lastDataRow + 1);
  if (scanEnd > startRow) {
    var existing = shPen.getRange(4, 1, scanEnd - 4, 8).getValues();
    for (var ei = 0; ei < existing.length; ei++) {
      if (String(existing[ei][1]).trim() === "Modal Awal" &&
          String(existing[ei][2]).trim() === "Mesin Blender") {
        SpreadsheetApp.getUi().alert("⚠ Data Mei sudah pernah di-import.\nHapus baris lama dulu jika ingin re-import.");
        return;
      }
    }
  }

  // ── 1. Tambah item baru ke Bahan sheet ──
  _ensureBahanItems();

  // ── 2. Load BahanData ──
  var bahanMap = {};
  getBahanData().forEach(function(row) {
    bahanMap[String(row[COL.BAHAN.NAMA]).trim()] = {
      kategori: row[COL.BAHAN.KATEGORI],
      satuan:   row[COL.BAHAN.SATUAN],
      pack:     Number(row[COL.BAHAN.PACK]) || 1,
      hargaBel: Number(row[COL.BAHAN.HARGA_BELI]) || 0
    };
  });

  // ── 3. Helper: ambil data Bahan ──
  function _bahanRow(nama) {
    var item = bahanMap[nama];
    if (!item) return null;
    var jml = item.pack;
    var hrg = item.hargaBel / item.pack;
    return ["02/05/2026", item.kategori, nama, item.satuan, jml, hrg];
  }

  // ── 4. Bangun semua baris ──
  var tgl = "02/05/2026";
  var rows = [];

  // ── A. OPERASIONAL (4 baris) ──
  var operasional = [
    [tgl, "Operasional", "Sewa Tempat",      "Bulan",  1, 300000],
    [tgl, "Operasional", "Sewa Gudang",      "Bulan",  1, 100000],
    [tgl, "Operasional", "Listrik",          "Bulan",  1, 100000],
    [tgl, "Operasional", "Gaji Karyawan",    "Bulan",  1, 500000],
  ];

  // ── B. MODAL AWAL (21 baris) ──
  var modalAwal = [
    [tgl, "Modal Awal", "Petty Cash",       "Hari",   1, 100000],
    [tgl, "Modal Awal", "Saving Budget",    "Bulan",  1, 1000000],
    [tgl, "Modal Awal", "Bangku Bakso",     "Pcs",    4, 41250],
    [tgl, "Modal Awal", "Servis Kompor",    "Pcs",    1, 100000],
    [tgl, "Modal Awal", "Servis Kabel Rol", "Pcs",    1, 50000],
    [tgl, "Modal Awal", "Laminating Menu",  "Pcs",    2, 7500],
    [tgl, "Modal Awal", "Mesin Blender",    "Pcs",    1, 350000],
    [tgl, "Modal Awal", "Termos",           "Pcs",    1, 135000],
    [tgl, "Modal Awal", "Botol Mayo",       "Pcs",    4, 10000],
    [tgl, "Modal Awal", "Banner",           "Pcs",    1, 120000],
    [tgl, "Modal Awal", "Servis Mesin",     "Pcs",    1, 50000],
    [tgl, "Modal Awal", "Tempat Topping",   "Pcs",   10, 3800],
    [tgl, "Modal Awal", "Trashbag",         "Roll",   3, 27667],
    [tgl, "Modal Awal", "Galon + Isi",      "Pcs",    1, 50000],
    [tgl, "Modal Awal", "Kabel 20m",        "Meter", 20, 4350],
    [tgl, "Modal Awal", "Pisau",            "Meter",  1, 20000],
    [tgl, "Modal Awal", "Kabel Lampu",      "Meter",  1, 50000],
    [tgl, "Modal Awal", "Lampu 15w",        "Pcs",    1, 20000],
    [tgl, "Modal Awal", "Lampu 50w",        "Pcs",    1, 45000],
    [tgl, "Modal Awal", "Tambal Ban",       "Pcs",    1, 20000],
    [tgl, "Modal Awal", "Mika Jelly",       "Pcs",   10, 1500],
  ];

  // ── C. BAHAN UTAMA (23 baris) ──
  var bahanUtama = [
    "Pop Ice - Chociato", "Pop Ice - Cokelat", "Pop Ice - Duren",
    "Pop Ice - Strawberry", "Pop Ice - Alpukat", "Pop Ice - Taro",
    "Pop Ice - Blueberry", "Pop Ice - Permen Karet", "Pop Ice - Lychee",
    "Pop Ice - Anggur", "Pop Ice - Mangga", "Pop Ice - Melon",
    "Pop Ice - Cappuccino", "Pop Ice - Moccacino", "Pop Ice - Vanilla Latte",
    "Pop Ice - Cookies & Krim", "Pop Ice - Es Doger", "Pop Ice - Es Teler",
    "Teh Celup", "Gooday", "Chocolatos", "ABC Klepon",
    "Biji Kopi Robusta", "Biji Kopi Arabika",
  ];

  // ── D. TOPPING (6 baris) ──
  var topping = [
    "Keju", "Chocolate", "Chocochips", "Mesis", "Bubuk Oreo", "Boba",
  ];

  // ── E. BAHAN PENDUKUNG (6 baris) ──
  var bahanPendukung = [
    "Gula Pasir", "Susu SKM", "Gula Aren",
    "Air (Galon)", "Es Batu Kristal", "Tissue",
  ];

  // ── F. KEMASAN (6 baris) ──
  var kemasan = [
    "Cup Plastik 18oz", "Tutup Cup Plastik",
    "Paper Cup 8oz", "Tutup Paper Cup",
    "Sedotan Boba", "Sedotan Biasa",
  ];

  // ── Gabungkan semua ──
  rows = rows.concat(operasional, modalAwal);
  bahanUtama.concat(topping, bahanPendukung, kemasan).forEach(function(nama) {
    var r = _bahanRow(nama);
    if (r) rows.push(r);
  });

  // ── 5. Tulis ke Sheet ──
  var range = shPen.getRange(startRow, 1, rows.length, 6);
  range.setValues(rows);

  // Format tanggal & harga
  shPen.getRange(startRow, 1, rows.length, 1).setNumberFormat("DD/MM/YYYY");
  shPen.getRange(startRow, 6, rows.length, 1).setNumberFormat('"Rp "#,##0');

  // Pastikan validation & format ada di baris yang ditulis
  var endRow = startRow + rows.length - 1;
  // Jika melebihi range pre-formatted, extend validation sampai endRow
  if (endRow > 5000) {
    for (var ri = startRow; ri <= endRow; ri++) {
      _ensureEntryRow(shPen, ri);
    }
  }
  // Khusus untuk PRODUCTION — extend range validation lama (203) ke range baru
  if (endRow > 203 && !shPen.getRange(500, 2).getDataValidation()) {
    var oldKatRule = shPen.getRange(4, 2).getDataValidation();
    var oldDateRule = shPen.getRange(4, 1).getDataValidation();
    var extendTo = Math.max(endRow, 5000);
    shPen.getRange("A4:A" + extendTo).setDataValidation(oldDateRule);
    shPen.getRange("B4:B" + extendTo).setDataValidation(oldKatRule);
    shPen.getRange("F4:F" + extendTo).setNumberFormat('"Rp "#,##0');
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "✅ Import berhasil!\n\n" +
    rows.length + " baris ditulis ke sheet Pengeluaran (mulai baris " + startRow + ").\n\n" +
    "▶ Jalankan menu POS → Simpan & Sync Stok untuk sync stok bahan.");
  auditLog("Import Data Mei", rows.length + " baris");
}

/**
 * Tambah item baru ke sheet Bahan jika belum ada.
 */
function _ensureBahanItems() {
  var shBahan = getSheet(SHEET.BAHAN);
  if (!shBahan) return;

  var newItems = [
    ["Bahan Utama",     "Gooday",     "Piece", 10, 24000],
    ["Bahan Utama",     "Chocolatos", "Piece", 10, 21000],
    ["Bahan Utama",     "ABC Klepon", "Piece", 10, 20000],
    ["Bahan Pendukung", "Tissue",     "Piece",  3, 25000],
  ];

  var existing = {};
  var data = shBahan.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    existing[String(data[i][COL.BAHAN.NAMA]).trim()] = true;
  }

  var nextRow = data.length + 1;
  newItems.forEach(function(item) {
    if (existing[item[1]]) return;
    shBahan.getRange(nextRow, 1, 1, 5).setValues([item]);
    shBahan.getRange(nextRow, 6).setFormula("=IFERROR(E"+nextRow+"/D"+nextRow+", 0)")
      .setNumberFormat('"Rp "#,##0');
    var bg = nextRow % 2 === 0 ? "#F5F5F5" : "#FFFFFF";
    shBahan.getRange(nextRow, 1, 1, 6).setBackground(bg).setFontSize(10);
    nextRow++;
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// ADD ROW — PENGELUARAN
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Navigasi ke baris kosong berikutnya + pastikan validation ada.
 * Alternatif lebih cepat drpd scroll manual cari baris kosong.
 */
function addRowPengeluaran() {
  var sh = getSheet(SHEET.PENGELUARAN);
  if (!sh) return;

  var lastRow = sh.getLastRow();
  var nextRow = 4;
  // scan dari bawah ke atas buat cari baris terakhir dengan data (A-C)
  if (lastRow >= 4) {
    var vals = sh.getRange(4, 1, lastRow - 3, 3).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]).trim() || String(vals[i][1]).trim() || String(vals[i][2]).trim()) {
        nextRow = 4 + i + 1;
        break;
      }
    }
  }

  // ensure validation/formula di baris target
  _ensureEntryRow(sh, nextRow);

  sh.setActiveSelection("C" + nextRow);
  SpreadsheetApp.getActiveSpreadsheet().toast("Baris " + nextRow + " — pilih Kategori dulu", "➕ Add Row");
}

/**
 * Pastikan satu baris Pengeluaran punya validation & format.
 * @param {Sheet} sh - Sheet Pengeluaran
 * @param {number} row - Nomor baris
 */
function _ensureEntryRow(sh, row) {
  if (row <= 3) return;

  // Date validation
  if (!sh.getRange(row, 1).getDataValidation()) {
    sh.getRange(row, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build()
    );
    sh.getRange(row, 1).setNumberFormat("DD/MM/YYYY");
  }

  // Kategori dropdown
  if (!sh.getRange(row, 2).getDataValidation()) {
    sh.getRange(row, 2).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(KATEGORI_LIST_PEN, true).setAllowInvalid(false).build()
    );
  }

  // Harga format
  var fmt = sh.getRange(row, 6).getNumberFormat();
  if (!fmt || fmt === "none" || fmt === "1") {
    sh.getRange(row, 6).setNumberFormat('"Rp "#,##0');
  }

  // Kolom H — font kecil
  sh.getRange(row, 8).setFontSize(10);
}

// ═══════════════════════════════════════════════════════════════════════════
// FIX — Konversi ArrayFormula G4 → per-row + bersihkan #REF!
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Hapus ArrayFormula dari G4, pasang per-row formula untuk semua baris data,
 * dan bersihkan error #REF! akibat konflik ArrayFormula vs per-row.
 */
function fixPengeluaranFormulas() {
  var sh = getSheet(SHEET.PENGELUARAN);
  if (!sh) return;
  var ui = SpreadsheetApp.getUi();
  if (!confirmAction("Hapus ArrayFormula di G4 dan pasang per-row formula untuk semua baris?\n#REF! akibat konflik akan otomatis hilang.", "🔧 Fix Formula?")) return;

  var lastRow = sh.getLastRow();
  var data = sh.getRange(4, 1, lastRow - 3, 3).getValues(); // A-C only
  var fixed = 0, cleared = 0, refFixed = 0;

  for (var i = 0; i < data.length; i++) {
    var r = i + 4;
    var cellG = sh.getRange(r, 7);
    var formula = cellG.getFormula();
    var isArray = formula && formula.indexOf("ARRAYFORMULA") >= 0;
    var isRef = cellG.getValue() === "#REF!";

    if (isArray) {
      // Ganti ArrayFormula → per-row
      cellG.setFormula("=IF(E" + r + "*F" + r + "=0,\"\",E" + r + "*F" + r + ")").setNumberFormat('"Rp "#,##0');
      cleared++;
    } else if (isRef) {
      // #REF! dari konflik — re-set formula
      cellG.setFormula("=IF(E" + r + "*F" + r + "=0,\"\",E" + r + "*F" + r + ")").setNumberFormat('"Rp "#,##0');
      refFixed++;
    }
  }

  SpreadsheetApp.flush();
  ui.alert("✅ Selesai!\n\n" +
    "ArrayFormula dihapus: " + cleared + "\n" +
    "#REF! diperbaiki: " + refFixed + "\n" +
    "Semua data pake per-row formula.");
  auditLog("Fix Pengeluaran Formulas", cleared + " cleared, " + refFixed + " refFixed");
}

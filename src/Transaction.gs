// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Transaction.gs (Transaction Save & BOM Stock Engine)
// ═══════════════════════════════════════════════════════════════════════════

// ── TRX Number Generator (PropertiesService-based, O(1), no backward scan) ─
/**
 * Dapatkan nomor transaksi berikutnya (O(1), via PropertiesService).
 * @return {number} Nomor transaksi berikutnya
 */
function getNextTrxNumber() {
  var prop = PropertiesService.getDocumentProperties();
  var nextVal = parseInt(prop.getProperty("TRX_COUNTER") || "0", 10) + 1;
  prop.setProperty("TRX_COUNTER", String(nextVal));
  return nextVal;
}

/**
 * Reset counter nomor transaksi.
 * @param {number} start - Mulai dari angka ini
 */
function resetTrxCounter(start) {
  PropertiesService.getDocumentProperties().setProperty("TRX_COUNTER", String(start - 1));
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — SIMPAN TRANSAKSI
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Simpan semua order dari POS ke Transaksi & jalankan BOM engine.
 * — Phase 1: Read & Validate (no writes, outside lock)
 * — Phase 2: Compute stock changes in memory
 * — Phase 3: Write batch + stock + clear (all-or-nothing, inside LockService)
 */
function simpanTransaksi() {
  timeStart("simpanTransaksi");
  if (!confirmAction("Yakin mau menyimpan transaksi ini? Data akan diproses dan dipotong dari stok.", "💾 Simpan Transaksi?")) return;

  // ══════════════════════════════════════════════════════════════════════
  // PHASE 1 — Read & Validate
  // ══════════════════════════════════════════════════════════════════════
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPOS = getSheet(SHEET.POS);
  var shTrx = getSheet(SHEET.TRANSAKSI);

  if (!shPOS || !shTrx) {
    SpreadsheetApp.getUi().alert("Sheet POS atau Transaksi tidak ditemukan. Jalankan Setup dulu.");
    return;
  }

  var kasir   = shPOS.getRange("B2").getValue() || "Kasir 1";
  var tanggal = fmtDate(new Date());
  var jam     = fmtTime(new Date());

  var lastPOSRow = shPOS.getLastRow();
  var posRange = shPOS.getRange(POS_START_ROW, 1, lastPOSRow - POS_START_ROW + 1, 8);
  var posData = posRange.getValues();
  var PV = COL.POS;

  // Phase 1a — Build batch data in memory (read only)
  var batchRows = [];
  var batchFormats = [];

  for (var ri = 0; ri < posData.length; ri++) {
    var row = posData[ri];
    var varian = String(row[PV.VARIAN] || "").trim();
    if (!varian || varian === "" || varian.indexOf("GRAND TOTAL") >= 0) continue;

    var jumlahCup  = Number(row[PV.CUP]) || 1;
    var topping    = String(row[PV.TOPPING] || "");
    var hargaBase  = Number(row[PV.HRG_BASE]) || 0;
    var hargaTop   = Number(row[PV.HRG_TOP]) || 0;
    var total      = Number(row[PV.TOTAL]) || 0;

    batchRows.push([
      null, null, // placeholder — diisi di Phase 2
      tanggal, jam, kasir, varian, jumlahCup, topping, Number(row[PV.JML_TOP]) || 0,
      hargaBase, hargaTop, total
    ]);
    batchFormats.push([jumlahCup, hargaBase, hargaTop, total]);
  }

  var saved = batchRows.length;
  if (saved === 0) {
    SpreadsheetApp.getUi().alert("Tidak ada data order di POS. Isi varian dulu ya!");
    return;
  }
  timeEnd("simpanTransaksi:P1_Read");

  // ══════════════════════════════════════════════════════════════════════
  // PHASE 2 — Compute stock changes in memory
  // ══════════════════════════════════════════════════════════════════════
  var shResep = getSheet(SHEET.RESEP);
  var shStock = getSheet(SHEET.STOCK);

  var kebutuhanBahan = {};
  if (shResep && shStock) {
    var resepData = shResep.getDataRange().getValues();
    var PR = COL.RESEP;
    for (var i = 0; i < batchFormats.length; i++) {
      var varianItem = batchRows[i][6]; // col F = varian
      var qty = batchFormats[i][0];
      var toppingStr = batchRows[i][8]; // col H = topping

      var itemsToFind = [varianItem];
      if (toppingStr) {
        var toppingArray = toppingStr.split(",").map(function(t) { return t.trim(); });
        itemsToFind = itemsToFind.concat(toppingArray);
      }

      for (var k = 0; k < itemsToFind.length; k++) {
        for (var r = 1; r < resepData.length; r++) {
          if (resepData[r][PR.MENU] === itemsToFind[k]) {
            var namaBahan = resepData[r][PR.BAHAN];
            var takaran = parseFloat(resepData[r][PR.TAKARAN]) || 0;
            if (!kebutuhanBahan[namaBahan]) kebutuhanBahan[namaBahan] = 0;
            kebutuhanBahan[namaBahan] += (takaran * qty);
          }
        }
      }
    }
  }
  timeEnd("simpanTransaksi:P2_Compute");

  // ══════════════════════════════════════════════════════════════════════
  // PHASE 3 — Write (all-or-nothing inside LockService)
  // ══════════════════════════════════════════════════════════════════════
  withLock(15000, function() {
    var nextNo = getNextTrxNumber();
    var trxLabel = "TRX-" + String(nextNo).padStart(3, "0");

    // Isi nomor transaksi & item number
    for (var b = 0; b < batchRows.length; b++) {
      batchRows[b][0] = trxLabel;
      batchRows[b][1] = b + 1;
    }

    // Write transaction rows
    var insertRow = shTrx.getLastRow() + 1;
    var dataRange = shTrx.getRange(insertRow, 1, saved, 12);
    dataRange.setValues(batchRows);
    dataRange.offset(0, 9, saved, 3).setNumberFormat('"Rp "#,##0');

    for (var bi = 0; bi < saved; bi++) {
      var r = insertRow + bi;
      var bgColor = (r % 2 === 0) ? "#EAF4FB" : "#FFFFFF";
      shTrx.getRange(r, 1, 1, 12)
        .setBackground(bgColor).setVerticalAlignment("middle")
        .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
      shTrx.setRowHeight(r, 22);
    }

    // Grand Total summary
    var totalCup = 0, totalBase = 0, totalTop = 0, totalBayar = 0;
    for (var si = 0; si < saved; si++) {
      var item = batchFormats[si];
      totalCup   += item[0];
      totalBase  += item[1];
      totalTop   += item[2];
      totalBayar += item[3];
    }

    var sumRow = shTrx.getLastRow() + 1;
    var TV = COL.TRANSAKSI;
    shTrx.getRange(sumRow, 1, 1, 6).merge()
      .setValue("TOTAL TRANSAKSI — " + tanggal + " " + jam)
      .setBackground(C.DARK).setFontColor(C.WHITE)
      .setFontWeight("bold").setFontSize(10)
      .setHorizontalAlignment("right").setVerticalAlignment("middle");
    shTrx.getRange(sumRow, COLx(TV.CUP)).setValue(totalCup)
      .setBackground(C.ORANGE).setFontColor(C.WHITE)
      .setFontWeight("bold").setHorizontalAlignment("center");
    shTrx.getRange(sumRow, COLx(TV.TOPPING), 1, 2).setBackground(C.DARK);
    shTrx.getRange(sumRow, COLx(TV.HRG_BASE)).setValue(totalBase)
      .setBackground(C.ORANGE).setFontColor(C.WHITE)
      .setFontWeight("bold").setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shTrx.getRange(sumRow, COLx(TV.HRG_TOP)).setValue(totalTop)
      .setBackground(C.ORANGE).setFontColor(C.WHITE)
      .setFontWeight("bold").setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shTrx.getRange(sumRow, COLx(TV.TOTAL)).setValue(totalBayar)
      .setBackground(C.GREEN).setFontColor(C.WHITE)
      .setFontWeight("bold").setFontSize(11).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shTrx.setRowHeight(sumRow, 28);

    var sepRow = sumRow + 1;
    shTrx.getRange(sepRow, 1, 1, 12).setBackground("#DDDDDD");
    shTrx.setRowHeight(sepRow, 6);

    // Stock deduction (BOM engine inline — batch write)
    if (shStock) {
      var stockData = shStock.getDataRange().getValues();
      var PC = COL.STOCK;
      var stockRows = [];
      for (var s = 1; s < stockData.length; s++) {
        var namaBahanStock = stockData[s][PC.NAMA];
        if (kebutuhanBahan[namaBahanStock]) {
          var terjualAwal = parseFloat(stockData[s][PC.TERJUAL]) || 0;
          var sisaAwal = parseFloat(stockData[s][PC.SISA]) || parseFloat(stockData[s][PC.STOK_AWAL]) || 0;
          stockRows.push({
            row: s + 1,
            terjual: terjualAwal + kebutuhanBahan[namaBahanStock],
            sisa: sisaAwal - kebutuhanBahan[namaBahanStock]
          });
        }
      }
      // Batch write stock updates
      if (stockRows.length > 0) {
        var terjualRange = shStock.getRange(stockRows[0].row, COLx(PC.TERJUAL), stockRows.length, 1);
        var sisaRange = shStock.getRange(stockRows[0].row, COLx(PC.SISA), stockRows.length, 1);
        var terjualData = [];
        var sisaData = [];
        for (var si = 0; si < stockRows.length; si++) {
          terjualData.push([stockRows[si].terjual]);
          sisaData.push([stockRows[si].sisa]);
        }
        terjualRange.setValues(terjualData);
        sisaRange.setValues(sisaData);
      }
    }

    // Clear POS
    _clearPOSRows(shPOS);

    SpreadsheetApp.flush();
    timeEnd("simpanTransaksi:P3_Write");
  });

  refreshLaporan();
  refreshNamedRanges(); // P3.4 — update named range sizes
  var totalBayar = 0;
  for (var ai = 0; ai < batchFormats.length; ai++) totalBayar += batchFormats[ai][3];
  auditLog("Simpan Transaksi", batchRows[0][0] + " — " + saved + " item, Rp " + totalBayar.toLocaleString("id-ID"));
  SpreadsheetApp.getUi().alert("✅ Transaksi " + batchRows[0][0] + " (" + saved + " item) berhasil disimpan!");
}

/**
 * Hapus data order dari POS (tanpa konfirmasi — dipanggil dari Phase 3).
 * @param {Sheet} shPOS - Sheet POS
 */
function _clearPOSRows(shPOS) {
  var lastRow = shPOS.getLastRow();
  var rowCount = lastRow - POS_START_ROW + 1;
  if (rowCount <= POS_INIT_ROWS) {
    shPOS.getRange(POS_START_ROW, 1, POS_INIT_ROWS, 8).clearContent();
    // Reset formulas
    for (var i = 0; i < POS_INIT_ROWS; i++) {
      var r = POS_START_ROW + i;
      shPOS.getRange(r, COLx(COL.POS.VARIAN)).clearDataValidation();
    }
    return;
  }
  // Hapus baris extra (dari PosStart + InitRows sampai LastRow)
  shPOS.deleteRows(POS_START_ROW + POS_INIT_ROWS, rowCount - POS_INIT_ROWS);
  shPOS.getRange(POS_START_ROW, 1, POS_INIT_ROWS, 8).clearContent();
  for (var j = 0; j < POS_INIT_ROWS; j++) {
    shPOS.getRange(POS_START_ROW + j, COLx(COL.POS.VARIAN)).clearDataValidation();
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ENGINE BOM — PEMOTONGAN STOK OTOMATIS (SAFE MODE, legacy, prefer inline)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Engine BOM — potong stok berdasarkan order di POS.
 * @param {Sheet} shPOS - Sheet POS object
 */
function stockEngineBOM(shPOS) {
  timeStart("stockEngineBOM");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shResep = getSheet(SHEET.RESEP);
  var shStock = getSheet(SHEET.STOCK);
  if (!shResep || !shStock) return;

  var lastPosRow = shPOS.getLastRow();
  var startRow = typeof POS_START_ROW !== 'undefined' ? POS_START_ROW : 7;
  if (lastPosRow < startRow) return; 

  var posData = shPOS.getRange(startRow, 2, lastPosRow - startRow + 1, 3).getValues();
  var resepData = shResep.getDataRange().getValues();
  var stockData = shStock.getDataRange().getValues();
  var kebutuhanBahan = {};

  var PV = COL.POS, PC = COL.STOCK, PR = COL.RESEP;
  for (var i = 0; i < posData.length; i++) {
    var varian = posData[i][PV.VARIAN - 1];
    var qty = parseInt(posData[i][PV.CUP - 1]) || 0;
    var toppingStr = posData[i][PV.TOPPING - 1];
    if (!varian || qty === 0) continue;

    var itemsToFind = [varian];
    if (toppingStr) {
      var toppingArray = toppingStr.split(",").map(function(t) { return t.trim(); });
      itemsToFind = itemsToFind.concat(toppingArray);
    }

    for (var k = 0; k < itemsToFind.length; k++) {
      for (var r = 1; r < resepData.length; r++) {
        if (resepData[r][PR.MENU] === itemsToFind[k]) {
          var namaBahan = resepData[r][PR.BAHAN];
          var takaran = parseFloat(resepData[r][PR.TAKARAN]) || 0;
          if (!kebutuhanBahan[namaBahan]) kebutuhanBahan[namaBahan] = 0;
          kebutuhanBahan[namaBahan] += (takaran * qty);
        }
      }
    }
  }

  var stockRows = [];
  for (var s = 1; s < stockData.length; s++) {
    var namaBahanStock = stockData[s][PC.NAMA];
    if (kebutuhanBahan[namaBahanStock]) {
      var unitStok = String(stockData[s][PC.SATUAN]).trim();
      var qtyDeduction = kebutuhanBahan[namaBahanStock];
      if (unitStok) {
        qtyDeduction = UnitConverter.toBase(qtyDeduction, unitStok);
      }
      var terjualAwal = parseFloat(stockData[s][PC.TERJUAL]) || 0;
      var sisaAwal = parseFloat(stockData[s][PC.SISA]) || parseFloat(stockData[s][PC.STOK_AWAL]) || 0;
      stockRows.push({ row: s + 1, terjual: terjualAwal + qtyDeduction, sisa: sisaAwal - qtyDeduction });
      stockData[s][PC.SISA] = sisaAwal - qtyDeduction;
    }
  }
  // Batch write stock updates
  if (stockRows.length > 0) {
    var terjualRange = shStock.getRange(stockRows[0].row, COLx(PC.TERJUAL), stockRows.length, 1);
    var sisaRange = shStock.getRange(stockRows[0].row, COLx(PC.SISA), stockRows.length, 1);
    var terjualData = [];
    var sisaData = [];
    for (var si = 0; si < stockRows.length; si++) {
      terjualData.push([stockRows[si].terjual]);
      sisaData.push([stockRows[si].sisa]);
    }
    terjualRange.setValues(terjualData);
    sisaRange.setValues(sisaData);
  }
  timeEnd("stockEngineBOM");
}

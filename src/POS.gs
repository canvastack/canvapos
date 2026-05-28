// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — POS.gs (Point of Sale Operations)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — ADD ROW
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Tambah baris order baru di POS sebelum Grand Total.
 */
function addRowPOS() {
  withLock(10000, function() {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = getSheet(SHEET.POS);
    var C   = getC();

    var lastRow = sh.getLastRow();
    var totalRow = -1;
    for (var r = POS_START_ROW; r <= lastRow; r++) {
      var formula = sh.getRange(r, 8).getFormula();
      if (formula.indexOf("SUM(H") >= 0) { totalRow = r; break; }
    }
    if (totalRow < 0) totalRow = lastRow + 1;

    var orderCount = totalRow - POS_START_ROW;
    var newNo = orderCount + 1;
    var bg = (newNo % 2 === 0) ? C.WHITE : C.INPUT;

    sh.insertRowBefore(totalRow);
    var newR = totalRow;

    styleOrderRow(sh, newR, newNo, bg);

    var gtRow = totalRow + 1;
    sh.getRange(gtRow, 8).setFormula(F("=SUM(H{start}:H{end})", {start: POS_START_ROW, end: gtRow - 1}));
    PropertiesService.getDocumentProperties().setProperty("POS_GRAND_TOTAL_ROW", String(gtRow));

    ss.setActiveSheet(sh);
    sh.setActiveRange(sh.getRange(newR, 2));
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — CLEAR POS
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Hapus semua baris order & reset POS ke kondisi awal.
 */
function clearPOS() {
  if (!confirmAction("Yakin mau menghapus SEMUA baris order di POS? Data yang belum disimpan akan hilang.", "⚠ Hapus Semua Order?")) return;

  withLock(15000, function() {
    var sh = getSheet(SHEET.POS);
    var C  = getC();

    var lastRow = sh.getLastRow();
    if (lastRow >= POS_START_ROW) {
      sh.deleteRows(POS_START_ROW, lastRow - POS_START_ROW + 1);
    }

    for (var i = 0; i < POS_INIT_ROWS; i++) {
      var r = POS_START_ROW + i;
      var bg = i % 2 === 0 ? C.INPUT : C.WHITE;
      styleOrderRow(sh, r, i + 1, bg);
    }

    var gtRow = POS_START_ROW + POS_INIT_ROWS;
    sh.getRange(gtRow, 1, 1, 7).merge()
      .setValue("GRAND TOTAL")
      .setBackground(C.DARK).setFontColor(C.WHITE)
      .setFontWeight("bold").setFontSize(12)
      .setHorizontalAlignment("right").setVerticalAlignment("middle");
    sh.getRange(gtRow, 8)
      .setFormula(F("=SUM(H{start}:H{end})", {start: POS_START_ROW, end: gtRow - 1}))
      .setNumberFormat('"Rp "#,##0')
      .setBackground(C.ORANGE).setFontColor(C.WHITE)
      .setFontWeight("bold").setFontSize(13)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sh.setRowHeight(gtRow, 36);

    PropertiesService.getDocumentProperties().setProperty("POS_GRAND_TOTAL_ROW", String(gtRow));
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — PILIH TOPPING (popup checklist via UI)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Tampilkan dialog checklist untuk pilih topping.
 */
function pilihTopping() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = getSheet(SHEET.POS);
  var ui  = SpreadsheetApp.getUi();

  var activeRow = ss.getActiveRange().getRow();
  var startRow = typeof POS_START_ROW !== 'undefined' ? POS_START_ROW : 7;

  if (activeRow < startRow) {
    ui.alert("Pilih dulu baris order yang ingin ditambah topping (baris " + startRow + " ke bawah).");
    return;
  }

  var varian = sh.getRange(activeRow, COLx(COL.POS.VARIAN)).getValue();
  if (!varian) {
    ui.alert("Pilih Product Variant dulu di kolom B sebelum menambah topping.");
    return;
  }

  var items = getToppingList();
  var existing = sh.getRange(activeRow, COLx(COL.POS.TOPPING)).getValue() || "";
  var selected = existing ? existing.split(",").map(function(t) { return t.trim(); }) : [];

  showMultiSelectDialog("🍬 Pilih Topping", items, selected, "_pilihToppingCallback");
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — HAPUS BARIS AKTIF DI POS (support multi-row selection)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Hapus baris order yang dipilih (support multi-row).
 */
function deleteRowPOS() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = getSheet(SHEET.POS);
  var ui  = SpreadsheetApp.getUi();

  if (ss.getActiveSheet().getName() !== "POS") {
    ui.alert("Pindah ke sheet POS dulu sebelum hapus baris.");
    return;
  }

  var selection  = ss.getActiveRange();
  var firstRow   = selection.getRow();
  var numRows    = selection.getNumRows();
  var lastSelRow = firstRow + numRows - 1;

  var gtRow = parseInt(PropertiesService.getDocumentProperties()
    .getProperty("POS_GRAND_TOTAL_ROW") || "0");

  if (firstRow < POS_START_ROW || lastSelRow >= gtRow) {
    ui.alert(
      "Pilih baris order yang ingin dihapus.\n" +
      "Baris order dimulai dari baris " + POS_START_ROW +
      " sampai sebelum GRAND TOTAL."
    );
    return;
  }

  var label = numRows === 1
    ? "Hapus baris " + (firstRow - POS_START_ROW + 1) + "?"
    : "Hapus " + numRows + " baris (" +
      (firstRow - POS_START_ROW + 1) + "–" +
      (lastSelRow - POS_START_ROW + 1) + ")?";

  var resp = ui.alert(label,
    "Baris yang dipilih akan dihapus permanen dari POS.",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  withLock(10000, function() {
    for (var r = lastSelRow; r >= firstRow; r--) {
      sh.deleteRow(r);
    }
    _renumberAndFixPOS(sh);
  });
}

// ── Helper: renumber baris order + update GRAND TOTAL property ────────────
function _renumberAndFixPOS(sh) {
  var C = getC();
  var lastRow = sh.getLastRow();
  var no = 1;

  for (var r = POS_START_ROW; r <= lastRow; r++) {
    var formula = sh.getRange(r, 8).getFormula();
    var isGrandTotal = formula.indexOf("SUM(H") >= 0;

    if (isGrandTotal) {
      // Update property dengan posisi terkini
      PropertiesService.getDocumentProperties().setProperty("POS_GRAND_TOTAL_ROW", String(r));
      sh.getRange(r, 8).setFormula(F("=SUM(H{start}:H{end})", {start: POS_START_ROW, end: r - 1}));
      break;
    }

    // Cek apakah ini baris order (ada dropdown di kolom B)
    var colAVal = sh.getRange(r, 1).getValue();
    if (colAVal === "" && sh.getRange(r, 3).getValue() === "") continue; // skip baris kosong aneh

    // Update nomor urut
    sh.getRange(r, 1).setValue(no)
      .setHorizontalAlignment("center").setBackground(C.LIGHT)
      .setFontColor(C.DARK).setFontWeight("bold").setFontSize(10);

    sh.getRange(r, 5).setFormula(
      F("=IF(B{row}=\"\",\"\",IF(D{row}=\"\",0,COUNTA(SPLIT(REGEXREPLACE(TRIM(D{row}),\"\\s*,\\s*$\",\"\"),\",\"))))", {row: r})
    );
    sh.getRange(r, 6).setFormula(F("=IF(B{row}=\"\",\"\",E{row}*C{row}*{HARGA_TOPPING})", {row: r, HARGA_TOPPING: HARGA_TOPPING}));
    sh.getRange(r, 7).setFormula(F("=IF(B{row}=\"\",\"\",C{row}*{HARGA_BASE})", {row: r, HARGA_BASE: HARGA_BASE}));
    sh.getRange(r, 8).setFormula(F("=IF(B{row}=\"\",\"\",G{row}+F{row})", {row: r}));

    no++;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SAFE CLEAR — backup + restore (P2.2.4-5)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Backup data POS ke PropertiesService lalu clear POS.
 */
function safeClearPOS() {
  if (!confirmAction("Yakin mau clear POS? Data akan di-backup dulu.", "⚠ Safe Clear POS?")) return;

  var sh = getSheet(SHEET.POS);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow < POS_START_ROW) { clearPOS(); return; }

  var data = sh.getRange(POS_START_ROW, 1, lastRow - POS_START_ROW + 1, 8).getValues();
  var backup = {timestamp: fmtDate(new Date()) + " " + fmtTime(new Date()), data: data};
  PropertiesService.getDocumentProperties().setProperty("POS_BACKUP", JSON.stringify(backup));

  clearPOS();
  SpreadsheetApp.getUi().alert("✅ POS cleared. Data di-backup. Klik menu → Restore POS untuk mengembalikan.");
}

/**
 * Pulihkan data POS dari backup terakhir.
 */
function restorePOSFromBackup() {
  if (!confirmAction("Pulihkan data POS dari backup terakhir?", "♻ Restore POS?")) return;
  var C = getC();

  var prop = PropertiesService.getDocumentProperties();
  var raw = prop.getProperty("POS_BACKUP");
  if (!raw) {
    SpreadsheetApp.getUi().alert("Tidak ada backup POS yang tersimpan.");
    return;
  }

  var backup;
  try { backup = JSON.parse(raw); } catch(e) {
    SpreadsheetApp.getUi().alert("Backup corrupt. Tidak bisa restore.");
    return;
  }

  var sh = getSheet(SHEET.POS);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow >= POS_START_ROW) {
    sh.deleteRows(POS_START_ROW, lastRow - POS_START_ROW + 1);
  }

  for (var i = 0; i < backup.data.length; i++) {
    var r = POS_START_ROW + i;
    sh.getRange(r, 1, 1, 8).setValues([backup.data[i]]);
    var bg = i % 2 === 0 ? C.INPUT : C.WHITE;
    styleData(sh.getRange(r, 1, 1, 8), bg);
    sh.setRowHeight(r, 24);
  }

  // Restore GRAND TOTAL row
  var gtRow = POS_START_ROW + backup.data.length;
  sh.getRange(gtRow, 1, 1, 7).merge()
    .setValue("GRAND TOTAL").setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange(gtRow, 8)
    .setFormula(F("=SUM(H{start}:H{end})", {start: POS_START_ROW, end: gtRow - 1}))
    .setNumberFormat('"Rp "#,##0').setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(gtRow, 36);
  PropertiesService.getDocumentProperties().setProperty("POS_GRAND_TOTAL_ROW", String(gtRow));

  SpreadsheetApp.getUi().alert("✅ Data POS dipulihkan dari backup (" + backup.timestamp + ").");
}

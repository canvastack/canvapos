// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Helpers.gs (Utility Functions & Styling)
// ═══════════════════════════════════════════════════════════════════════════

// ── Cache ──────────────────────────────────────────────────────────────────
var CACHED_VARIAN_LIST = null;
var CACHED_TOPPING_LIST = null;

/**
 * Dapatkan daftar item berdasarkan kategori dari sheet Bahan.
 * @param {string} kategori - Nama kategori (contoh: "Topping")
 * @return {string[]} Daftar item
 */
function getDynamicList(kategori) {
  var sh = getSheet(SHEET.BAHAN);
  if (!sh) return ["-"];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return ["-"];
  var data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === kategori && data[i][1] !== "") {
      list.push(data[i][1]);
    }
  }
  return list.length > 0 ? list : ["-"];
}

/**
 * Dapatkan daftar varian menu dari sheet Resep (cache 1x).
 * @return {string[]} Daftar varian produk
 */
function getVarianList() {
  if (CACHED_VARIAN_LIST) return CACHED_VARIAN_LIST;

  var shResep = getSheet(SHEET.RESEP);
  if (!shResep) { CACHED_VARIAN_LIST = ["-"]; return CACHED_VARIAN_LIST; }

  var resepData = shResep.getDataRange().getValues();
  var TOPPINGS = getToppingList();
  var PS = COL.STOCK, PR = COL.RESEP;

  // Bangun stock map: { "Gula Pasir": 500, ... }
  var shStock = getSheet(SHEET.STOCK);
  var stockMap = {};
  if (shStock) {
    var stockData = shStock.getDataRange().getValues();
    for (var si = 1; si < stockData.length; si++) {
      var nama = String(stockData[si][PS.NAMA]).trim();
      var sisa = Number(stockData[si][PS.SISA]) || 0;
      if (nama) stockMap[nama] = sisa;
    }
  }

  // Bangun BOM map: { "Pop Ice - Cokelat": [["Gula Pasir",20], ...], ... }
  var bomMap = {};
  for (var ri = 1; ri < resepData.length; ri++) {
    var menu = String(resepData[ri][PR.MENU]).trim();
    var bahan = String(resepData[ri][PR.BAHAN]).trim();
    var takaran = Number(resepData[ri][PR.TAKARAN]) || 0;
    if (!menu || !bahan) continue;
    if (!bomMap[menu]) bomMap[menu] = [];
    bomMap[menu].push([bahan, takaran]);
  }

  // Kumpulin menu unik, filter topping + stock
  var seen = {};
  var list = [];
  for (var i = 1; i < resepData.length; i++) {
    var name = String(resepData[i][PR.MENU]).trim();
    if (!name || seen[name]) continue;
    seen[name] = true;
    if (TOPPINGS.indexOf(name) >= 0) continue;

    // Cek kecukupan stock buat varian ini
    var bomItems = bomMap[name];
    var stockCukup = true;
    if (bomItems && shStock) {
      for (var bi = 0; bi < bomItems.length; bi++) {
        var bahan = bomItems[bi][0];
        var perlu = bomItems[bi][1];
        var sedia = stockMap[bahan] || 0;
        if (sedia < perlu) {
          stockCukup = false;
          break;
        }
      }
    }

    if (stockCukup) list.push(name);
  }

  CACHED_VARIAN_LIST = list.length > 0 ? list : ["-"];
  return CACHED_VARIAN_LIST;
}

/**
 * Dapatkan daftar topping dari sheet Bahan (cache 1x).
 * @return {string[]} Daftar topping
 */
function getToppingList() {
  if (!CACHED_TOPPING_LIST) CACHED_TOPPING_LIST = getDynamicList("Topping");
  return CACHED_TOPPING_LIST;
}

/**
 * Hapus cache daftar varian & topping.
 */
function clearDynamicCache() {
  CACHED_VARIAN_LIST = null;
  CACHED_TOPPING_LIST = null;
  clearHPPLookupCache();
  PropertiesService.getDocumentProperties().deleteProperty("HPP_CAT_CACHE");
}

// ── Unit Converter (Weight & Volume) ──────────────────────────────────────
/**
 * UnitConverter — normalisasi unit weight & volume.
 * Support: Kg↔Gram, Liter↔ml, Piece↔Piece, unknown = passthrough.
 */
var UnitConverter = {
  _weight: {kg: 1000, kilogram: 1000, gram: 1, g: 1},
  _volume: {liter: 1000, l: 1000, ml: 1, mililiter: 1},

  /**
   * Cari faktor konversi ke base unit (gram atau ml).
   * @param {string} unit - Satuan input
   * @return {number} Faktor konversi (1 = unknown/piece)
   */
  _factor: function(unit) {
    var u = String(unit).trim().toLowerCase();
    if (this._weight[u]) return this._weight[u];
    if (this._volume[u]) return this._volume[u];
    return 1; // piece atau unknown → passthrough
  },

  /**
   * Cari base unit name untuk display.
   * @param {string} unit - Satuan input
   * @return {string} Base unit
   */
  baseUnit: function(unit) {
    var u = String(unit).trim().toLowerCase();
    if (this._weight[u]) return "Gram";
    if (this._volume[u]) return "ml";
    return unit;
  },

  /**
   * Konversi qty dari fromUnit ke toUnit.
   * @param {number} qty - Jumlah
   * @param {string} fromUnit - Satuan asal
   * @param {string} toUnit - Satuan tujuan
   * @return {number} Jumlah dalam satuan tujuan
   */
  convert: function(qty, fromUnit, toUnit) {
    var fromBase = this._factor(fromUnit);
    var toBase = this._factor(toUnit);
    if (fromBase === 0 || toBase === 0) return qty;
    return (qty * fromBase) / toBase;
  },

  /**
   * Konversi ke base unit (Gram / ml / as-is).
   * @param {number} qty - Jumlah
   * @param {string} unit - Satuan asal
   * @return {number} Jumlah dalam base unit
   */
  toBase: function(qty, unit) {
    return qty * this._factor(unit);
  }
};

// ── Bahan Category & Price Maps ─────────────────────────────────────────
/**
 * Dapatkan kategori HPP untuk suatu nama bahan.
 * @param {string} bahanName - Nama bahan
 * @return {string} Kategori HPP ("Bahan Utama", "Topping", "Bahan Pendukung", "Kemasan", atau "Lain-lain")
 */
function getCategoryForBahan(bahanName) {
  var data = getBahanData();
  for (var i = 0; i < data.length; i++) {
    if (data[i][COL.BAHAN.NAMA] === bahanName) {
      var kat = data[i][COL.BAHAN.KATEGORI];
      return KATEGORI_HPP_MAP[kat] || "Lain-lain";
    }
  }
  return "Lain-lain";
}

/**
 * Dapatkan map harga per satuan semua bahan.
 * @return {Object} { "Gula Pasir": 20, "Cup Plastik 18oz": 500, ... }
 */
function getBahanHargaMap() {
  var sh = getSheet(SHEET.BAHAN);
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][COL.BAHAN.NAMA]).trim();
    var harga = Number(data[i][COL.BAHAN.HARGA_PIECE]) || 0;
    if (nama) map[nama] = harga;
  }
  return map;
}

// ── Performance Timer ────────────────────────────────────────────────────────
var _timers = {};

/**
 * Mulai timer dengan label tertentu.
 * @param {string} label - Label timer
 */
function timeStart(label) {
  _timers[label] = new Date().getTime();
}

/**
 * Akhiri timer dan catat durasi ke Logger + audit log.
 * @param {string} label - Label timer
 * @param {boolean} [warnIfSlow=true] - Tampilkan peringatan jika > 10 detik
 */
function timeEnd(label, warnIfSlow) {
  if (warnIfSlow === undefined) warnIfSlow = true;
  if (!_timers[label]) {
    Logger.log("⚠ Timer '" + label + "' belum dimulai.");
    return;
  }
  var elapsed = new Date().getTime() - _timers[label];
  delete _timers[label];
  Logger.log("⏱ " + label + ": " + elapsed + "ms");
  if (elapsed > 10000 && warnIfSlow) {
    Logger.log("⚠ PERINGATAN: " + label + " melebihi 10 detik (" + elapsed + "ms)");
  }
  return elapsed;
}

/**
 * Catat laporan performa ke sheet Audit.
 * @param {string} label - Label operasi
 * @param {number} elapsedMs - Waktu eksekusi dalam ms
 */
function logPerf(label, elapsedMs) {
  Logger.log("⏱ " + label + ": " + elapsedMs + "ms");
}

// ── Color Palette ──────────────────────────────────────────────────────────
/**
 * Dapatkan palet warna untuk styling.
 * @return {Object} Objek warna
 */
function getC() {
  return {
    BLUE:   "#2E86AB", GREEN:  "#27AE60", ORANGE: "#E67E22",
    RED:    "#E74C3C", PURPLE: "#8E44AD", YELLOW: "#F9CA24",
    LIGHT:  "#EAF4FB", WHITE:  "#FFFFFF", DARK:   "#2C3E50",
    LGREEN: "#D5F5E3", LRED:   "#FADBD8", INPUT:  "#FEF9E7"
  };
}

// ── Date/Time Helpers ──────────────────────────────────────────────────────
/**
 * Format tanggal ke DD/MM/YYYY (Asia/Jakarta).
 * @param {Date|string} date - Tanggal atau string
 * @return {string} Tanggal terformat
 */
function fmtDate(date) {
  if (typeof date === "string") return date;
  return Utilities.formatDate(new Date(date), "Asia/Jakarta", "dd/MM/yyyy");
}
/**
 * Format waktu ke HH:mm (Asia/Jakarta).
 * @param {Date} date - Objek tanggal
 * @return {string} Waktu terformat
 */
function fmtTime(date) {
  return Utilities.formatDate(new Date(date), "Asia/Jakarta", "HH:mm");
}
/**
 * Parse string DD/MM/YYYY ke Date object.
 * @param {string} str - String tanggal DD/MM/YYYY
 * @return {Date} Objek Date
 */
function parseDate(str) {
  var parts = str.split("/");
  if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
  return new Date(str);
}

// ── Sheet Safe Accessor ────────────────────────────────────────────────────
/**
 * Dapatkan sheet berdasarkan nama.
 * @param {string} name - Nama sheet
 * @return {Sheet|undefined} Sheet object
 */
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name);
}

// ── Formula Builder ────────────────────────────────────────────────────────
/**
 * Formula builder dengan placeholder {key}.
 * @param {string} tpl - Template formula
 * @param {Object} vars - Variable map
 * @return {string} Formula jadi
 */
function F(tpl, vars) {
  return tpl.replace(/\{(\w+)\}/g, function(_, k) {
    return vars[k] !== undefined ? vars[k] : "{" + k + "}";
  });
}

// ── Zebra Stripe ───────────────────────────────────────────────────────────
/**
 * Terapkan zebra stripe pada baris.
 * @param {Sheet} sh - Sheet object
 * @param {number} r - Nomor baris (1-based)
 * @param {number} i - Index iterasi (0-based)
 * @param {number} [numCols] - Jumlah kolom
 * @param {string} [evenBg] - Warna background genap
 * @return {string} Warna background yang diterapkan
 */
function applyZebraRow(sh, r, i, numCols, evenBg) {
  var C = getC();
  var bg = i % 2 === 0 ? (evenBg || C.LIGHT) : C.WHITE;
  styleData(sh.getRange(r, 1, 1, numCols || 1), bg);
  return bg;
}

// ── LockService Wrapper ────────────────────────────────────────────────────
/**
 * Eksekusi critical section dengan document lock.
 * @param {number} timeout - Maks waktu tunggu lock (ms)
 * @param {Function} fn - Fungsi critical section
 * @return {*} Hasil dari fn
 */
function withLock(timeout, fn) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(timeout || 10000);
    if (isDebug()) Logger.log("🔒 Lock acquired (" + (timeout || 10000) + "ms)");
  } catch (e) {
    if (isDebug()) Logger.log("🔒 Lock TIMEOUT (" + (timeout || 10000) + "ms)");
    throw new Error("Lock timeout (" + (timeout || 10000) + "ms)");
  }
  try {
    return fn();
  } catch (e) {
    throw new Error("Lock execution error: " + e.message);
  } finally {
    lock.releaseLock();
    if (isDebug()) Logger.log("🔒 Lock released");
  }
}

// ── Safe Execution ─────────────────────────────────────────────────────────
/**
 * Eksekusi fungsi dengan error handling.
 * @param {string} name - Nama operasi untuk log
 * @param {Function} fn - Fungsi yang akan dieksekusi
 */
function safeExecute(name, fn) {
  try {
    fn();
  } catch (e) {
    Logger.log("ERROR [" + name + "]: " + e.message);
    notify("Error: " + name + " — " + e.message, true);
  }
}

/**
 * Tampilkan toast notifikasi atau alert error.
 * @param {string} message - Pesan
 * @param {boolean} [isError] - True = alert error, false = toast
 */
function notify(message, isError) {
  try {
    var ui = SpreadsheetApp.getUi();
    if (isError) {
      ui.alert("⚠ Error", message, ui.ButtonSet.OK);
    } else {
      ui.toast(message, "CanvaPOS", 3);
    }
  } catch(e) { /* no UI available */ }
}

// ── Styling ────────────────────────────────────────────────────────────────
/**
 * Set font family Nunito ke seluruh sheet.
 * @param {Sheet} sh - Sheet object
 */
function setSheetFont(sh) {
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily("Nunito");
}

/**
 * Terapkan gaya header (bold, border, background).
 * @param {Range} range - Range header
 * @param {string} bg - Warna background
 * @param {string} [fg] - Warna font (default putih)
 */
function styleHeader(range, bg, fg) {
  var C = getC();
  fg = fg || C.WHITE;
  range.setBackground(bg).setFontColor(fg).setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Terapkan gaya data (border, vertical align).
 * @param {Range} range - Range data
 * @param {string} [bg] - Warna background
 */
function styleData(range, bg) {
  var C = getC();
  range.setBackground(bg || C.WHITE).setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Terapkan gaya baris order POS (dropdown, formula, format).
 * @param {Sheet} sh - Sheet object
 * @param {number} r - Nomor baris
 * @param {number} no - Nomor urut
 * @param {string} bg - Warna background
 */
function styleOrderRow(sh, r, no, bg) {
  var C = getC();
  var fs = 10;
  var varianRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(getVarianList(), true).setAllowInvalid(false).build();
  sh.getRange(r, 1).setValue(no).setHorizontalAlignment("center").setBackground(C.LIGHT)
    .setFontColor(C.DARK).setFontWeight("bold").setFontSize(fs);
  sh.getRange(r, 2).setBackground(bg).setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs)
    .setHorizontalAlignment("left").setDataValidation(varianRule);
  sh.getRange(r, 3).setValue(1).setHorizontalAlignment("center").setBackground(bg)
    .setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs);
  sh.getRange(r, 4).setBackground(bg).setFontColor(C.DARK).setFontWeight("normal")
    .setFontSize(fs).setHorizontalAlignment("left").setWrap(true);
  sh.getRange(r, 5).setFormula(
    F("=IF(B{row}=\"\",\"\",IF(D{row}=\"\",0,LEN(TRIM(D{row}))-LEN(SUBSTITUTE(TRIM(D{row}),\",\",\"\"))+1))", {row: r})
  ).setHorizontalAlignment("center").setBackground(C.LIGHT).setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs);
  sh.getRange(r, 6).setFormula(
    F("=IF(B{row}=\"\",\"\",E{row}*C{row}*{HARGA_TOPPING})", {row: r, HARGA_TOPPING: HARGA_TOPPING})
  ).setNumberFormat('"Rp "#,##0').setBackground(C.LIGHT).setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs).setHorizontalAlignment("right");
  sh.getRange(r, 7).setFormula(
    F("=IF(B{row}=\"\",\"\",C{row}*{HARGA_BASE})", {row: r, HARGA_BASE: HARGA_BASE})
  ).setNumberFormat('"Rp "#,##0').setBackground(C.LIGHT).setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs).setHorizontalAlignment("right");
  sh.getRange(r, 8).setFormula(
    F("=IF(B{row}=\"\",\"\",G{row}+F{row})", {row: r})
  ).setNumberFormat('"Rp "#,##0').setBackground(C.LGREEN).setFontColor(C.DARK).setFontWeight("bold").setFontSize(fs).setHorizontalAlignment("right");
  sh.setRowHeight(r, 28);
}

/**
 * Terapkan format mata uang Rp.
 * @param {Range} range - Range target
 */
function setCurrency(range) {
  range.setNumberFormat('"Rp "#,##0');
}

/**
 * Hapus sheet jika ada.
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} name - Nama sheet
 */
function deleteSheetIfExists(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
}

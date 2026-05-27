// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Google Apps Script (Enhanced with BOM & Event-Driven Stock)
// Jalankan fungsi setupPOS() sekali untuk build semua sheet & tombol
// FIX: semua global pakai var (bukan const) agar reliable di Apps Script V8
// ═══════════════════════════════════════════════════════════════════════════

var HARGA_BASE    = 5000;
var HARGA_TOPPING = 1000;
var HPP_PER_CUP   = 2200;
var HPP_PER_TOP   = 80;
var POS_START_ROW = 7;
var POS_INIT_ROWS = 5;

// --- SISTEM DROPDOWN DINAMIS (CACHED) ---
var CACHED_VARIAN_LIST = null;
var CACHED_TOPPING_LIST = null;

function getDynamicList(kategori) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Bahan");
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

function getVarianList() {
  if (CACHED_VARIAN_LIST) return CACHED_VARIAN_LIST;
  // Ambil daftar menu unik dari tab Resep (col A) — hanya produk, bukan topping
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Resep");
  if (!sh) { CACHED_VARIAN_LIST = ["-"]; return CACHED_VARIAN_LIST; }
  var data = sh.getDataRange().getValues();
  var seen = {};
  var list = [];
  var TOPPINGS = ["Keju","Chocolate","Chocochips","Mesis","Bubuk Oreo","Boba"];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0]).trim();
    if (name && !seen[name] && TOPPINGS.indexOf(name) < 0) {
      seen[name] = true;
      list.push(name);
    }
  }
  CACHED_VARIAN_LIST = list.length > 0 ? list : ["-"];
  return CACHED_VARIAN_LIST;
}

function getToppingList() {
  if (!CACHED_TOPPING_LIST) CACHED_TOPPING_LIST = getDynamicList("Topping");
  return CACHED_TOPPING_LIST;
}

// --- HELPER: Clear Cache untuk refresh data ---
function clearDynamicCache() {
  CACHED_VARIAN_LIST = null;
  CACHED_TOPPING_LIST = null;
}

// Warna — pakai fungsi getter agar selalu terdefinisi saat dipanggil
function getC() {
  return {
    BLUE:   "#2E86AB", GREEN:  "#27AE60", ORANGE: "#E67E22",
    RED:    "#E74C3C", PURPLE: "#8E44AD", YELLOW: "#F9CA24",
    LIGHT:  "#EAF4FB", WHITE:  "#FFFFFF", DARK:   "#2C3E50",
    LGREEN: "#D5F5E3", LRED:   "#FADBD8", INPUT:  "#FEF9E7"
  };
}

// ── Helper: set font Nunito untuk seluruh sheet ───────────────────────────────
function setSheetFont(sh) {
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily("Nunito");
}

// ── Helper: style header row ──────────────────────────────────────────────────
function styleHeader(range, bg, fg) {
  var C = getC();
  fg = fg || C.WHITE;
  range.setBackground(bg)
       .setFontColor(fg)
       .setFontWeight("bold")
       .setFontSize(11)
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
}

// ── Helper: style data row ────────────────────────────────────────────────────
function styleData(range, bg) {
  var C = getC();
  range.setBackground(bg || C.WHITE)
       .setVerticalAlignment("middle")
       .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
}

// ── Helper: style satu baris order POS ───────────────────────────────────────
function styleOrderRow(sh, r, no, bg) {
  var C = getC();
  var fs = 10; // font size konsisten untuk semua baris order

  var varianRule = SpreadsheetApp.newDataValidation()
  //  .requireValueInList(VARIAN_LIST, true)
    .requireValueInList(getVarianList(), true)
    .setAllowInvalid(false).build();

  // Kolom A: No
  sh.getRange(r, 1).setValue(no)
    .setHorizontalAlignment("center").setBackground(C.LIGHT)
    .setFontColor(C.DARK).setFontWeight("bold").setFontSize(fs);
  
  // Kolom B: Product Variant
  sh.getRange(r, 2).setBackground(bg)
    .setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs)
    .setHorizontalAlignment("left").setDataValidation(varianRule);
  
  // Kolom C: Jumlah Cup
  sh.getRange(r, 3).setValue(1).setHorizontalAlignment("center")
    .setBackground(bg).setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs);
  
  // Kolom D: Topping - Jenis
  sh.getRange(r, 4).setBackground(bg)
    .setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs)
    .setHorizontalAlignment("left").setWrap(true);
  
  // Kolom E: Topping - Jumlah
  sh.getRange(r, 5).setFormula(
    "=IF(B"+r+"=\"\",\"\",IF(D"+r+"=\"\",0,LEN(TRIM(D"+r+"))-LEN(SUBSTITUTE(TRIM(D"+r+"),\",\",\"\"))+1))"
  ).setHorizontalAlignment("center").setBackground(C.LIGHT)
    .setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs);
  
  // Kolom F: Topping - Harga Top
  sh.getRange(r, 6).setFormula(
    "=IF(B"+r+"=\"\",\"\",E"+r+"*C"+r+"*"+HARGA_TOPPING+")"
  ).setNumberFormat('"Rp "#,##0').setBackground(C.LIGHT)
    .setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs).setHorizontalAlignment("right");
  
  // Kolom G: Harga Base
  sh.getRange(r, 7).setFormula(
    "=IF(B"+r+"=\"\",\"\",C"+r+"*"+HARGA_BASE+")"
  ).setNumberFormat('"Rp "#,##0').setBackground(C.LIGHT)
    .setFontColor(C.DARK).setFontWeight("normal").setFontSize(fs).setHorizontalAlignment("right");
  
  // Kolom H: Total
  sh.getRange(r, 8).setFormula(
    "=IF(B"+r+"=\"\",\"\",G"+r+"+F"+r+")"
  ).setNumberFormat('"Rp "#,##0').setBackground(C.LGREEN)
    .setFontColor(C.DARK).setFontWeight("bold").setFontSize(fs).setHorizontalAlignment("right");
  
  sh.setRowHeight(r, 28);
}

// ── Helper: currency format ───────────────────────────────────────────────────
function setCurrency(range) {
  range.setNumberFormat('"Rp "#,##0');
}

// ── Helper: hapus sheet jika ada ─────────────────────────────────────────────
function deleteSheetIfExists(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
}

// ═══════════════════════════════════════════════════════════════════════════
// MAIN SETUP — Jalankan setupPOS() sekali, atau step-by-step jika timeout
// ═══════════════════════════════════════════════════════════════════════════
function setupPOS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone("Asia/Jakarta");

  // Buat sheet temp dulu agar tidak pernah 0 sheet
  var tempName = "__temp__";
  deleteSheetIfExists(ss, tempName);
  ss.insertSheet(tempName);

  ["Panduan","POS","Stock","Transaksi","Pendapatan","Pengeluaran","Kas","Bahan","Resep","Sheet1"].forEach(function(n) {
    deleteSheetIfExists(ss, n);
  });

  SpreadsheetApp.flush();
  // ORDER PENTING: Pengeluaran & Kas harus dibangun SEBELUM Pendapatan,
  // karena formula di Pendapatan mereferensi sheet Penguin + Kas
  buildBahan(ss);       SpreadsheetApp.flush();
  buildResep(ss);       SpreadsheetApp.flush();
  buildTransaksi(ss);   SpreadsheetApp.flush();
  buildStock(ss);       SpreadsheetApp.flush();
  buildPengeluaran(ss); SpreadsheetApp.flush();
  buildKas(ss);         SpreadsheetApp.flush();
  buildPendapatan(ss);  SpreadsheetApp.flush();
  buildPOS(ss);         SpreadsheetApp.flush();
  buildPanduan(ss);     SpreadsheetApp.flush();

  deleteSheetIfExists(ss, tempName);

  var order = ["Panduan","POS","Stock","Transaksi","Pendapatan","Pengeluaran","Kas","Bahan","Resep"];
  order.forEach(function(name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });

  ss.setActiveSheet(ss.getSheetByName("POS"));
  clearDynamicCache();
  try {
    SpreadsheetApp.getUi().alert("✅ Setup selesai!\n\nSemua sheet sudah dibuat.\nMulai transaksi dari sheet POS.");
  } catch(e) { /* dipanggil dari non-UI context, skip alert */ }
}

// Jalankan fungsi-fungsi ini satu per satu jika setupPOS() timeout:
function setup_1_Bahan()      { buildBahan(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_2_Resep()      { buildResep(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_3_Transaksi()  { buildTransaksi(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_4_Stock()      { buildStock(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_5_Pengeluaran() { buildPengeluaran(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_5b_Kas()        { buildKas(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_5c_Pendapatan() { buildPendapatan(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_6_POS()        { buildPOS(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_7_Panduan()    { buildPanduan(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
function setup_8_Reorder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var order = ["Panduan","POS","Stock","Transaksi","Pendapatan","Pengeluaran","Kas","Bahan","Resep"];
  order.forEach(function(name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(ss.getSheetByName("POS"));
  try { SpreadsheetApp.getUi().alert("✅ Setup selesai!"); } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — BAHAN
// ═══════════════════════════════════════════════════════════════════════════
function buildBahan(ss) {
  var C = getC();
  var sh = ss.insertSheet("Bahan");
  setSheetFont(sh);
  sh.setTabColor(C.BLUE);

  var headers = ["Kategori","Nama Bahan","Ukuran/Pack","Satuan","Harga Beli","Harga Per Piece"];
  var hRow = sh.getRange(1, 1, 1, headers.length);
  hRow.setValues([headers]);
  styleHeader(hRow, C.BLUE);
  sh.setRowHeight(1, 32);

  var data = [
    // ── Bahan Utama ────────────────────────────────────────────────
    ["Bahan Utama","Pop Ice - Chociato",           "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Cokelat",            "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Duren",              "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Strawberry",         "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Alpukat",            "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Taro",               "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Blueberry",          "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Permen Karet",       "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Lychee",             "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Anggur",             "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Mangga",             "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Melon",              "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Cappuccino",         "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Moccacino",          "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Vanilla Latte",      "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Cookies & Krim",     "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Es Doger",           "Renceng",10, 15000],
    ["Bahan Utama","Pop Ice - Es Teler",           "Renceng",10, 15000],
    ["Bahan Utama","Biji Kopi Robusta",            "Kg",      1,200000],
    ["Bahan Utama","Biji Kopi Arabika",            "Kg",      1,500000],
    ["Bahan Utama","Teh Celup",                    "Pack",  100, 50000],
    // ── Topping ────────────────────────────────────────────────────
    ["Topping","Keju",                "Gram",  250, 20000],
    ["Topping","Chocolate",           "Gram",  250, 20000],
    ["Topping","Chocochips",          "Gram",  250, 15000],
    ["Topping","Mesis",               "Gram",  250, 15000],
    ["Topping","Bubuk Oreo",          "Gram",  250, 15000],
    ["Topping","Boba",                "Gram",  250, 20000],
    // ── Bahan Pendukung ────────────────────────────────────────────
    ["Bahan Pendukung","Gula Pasir",  "Gram", 1000, 20000],
    ["Bahan Pendukung","Susu SKM",    "Gram",  370, 25000],
    ["Bahan Pendukung","Gula Aren",   "Gram",  500, 20000],
    ["Bahan Pendukung","Air (Galon)", "Liter",  19, 22000],
    ["Bahan Pendukung","Es Batu Kristal","Kg",  20, 25000],
    // ── Kemasan ────────────────────────────────────────────────────
    ["Kemasan","Cup Plastik 18oz",     "Pack",   50, 25000],
    ["Kemasan","Tutup Cup Plastik",    "Pack",   50,  5000],
    ["Kemasan","Paper Cup 8oz",        "Pack",   50, 25000],
    ["Kemasan","Tutup Paper Cup",      "Pack",   50,  5000],
    ["Kemasan","Sedotan Boba",         "Pack",  250, 15000],
    ["Kemasan","Sedotan Biasa",        "Pack",  500, 15000]
  ];

  // Tulis kolom A-E (tanpa Harga Per Piece)
  sh.getRange(2, 1, data.length, data[0].length).setValues(data);

  // Kolom F (Harga Per Piece) = =E/D
  data.forEach(function(_, i) {
    var r = i + 2;
    sh.getRange(r, 6).setFormula("=IFERROR(E"+r+"/D"+r+", 0)");
  });

  data.forEach(function(_, i) {
    var r = i + 2;
    var bg = i % 2 === 0 ? C.LIGHT : C.WHITE;
    styleData(sh.getRange(r, 1, 1, headers.length), bg);
    setCurrency(sh.getRange(r, 5, 1, 2));
  });

  [100,200,100,80,120,140].forEach(function(w, i) { sh.setColumnWidth(i+1, w); });
  sh.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — RESEP (BILL OF MATERIALS / BOM)
// ═══════════════════════════════════════════════════════════════════════════
function buildResep(ss) {
  var C = getC();
  var sh = ss.insertSheet("Resep");
  setSheetFont(sh);
  sh.setTabColor(C.PURPLE);

  // Struktur Kolom Relasional Many-to-Many
  // E = Harga Per Piece (VLOOKUP dari Bahan), F = Harga Per Takaran (C × E)
  var headers = ["Nama Menu / Varian", "Nama Bahan Baku", "Takaran (Qty Usage)", "Satuan Penggunaan", "Harga Per Piece", "Harga Per Takaran"];
  var hRow = sh.getRange(1, 1, 1, headers.length);
  hRow.setValues([headers]);
  styleHeader(hRow, C.PURPLE);
  sh.setRowHeight(1, 32);

  // Data Master Resep (BOM) — updated per Request.md
  var data = [
    // POP ICE BLENDER — 18 variants, same BOM composition
    ["Pop Ice - Chociato",       "Pop Ice - Chociato",    1, "Piece"],
    ["Pop Ice - Chociato",       "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Chociato",       "Susu SKM",            30, "Gram"],
    ["Pop Ice - Chociato",       "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Chociato",       "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Chociato",       "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Chociato",       "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Chociato",       "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Cokelat",        "Pop Ice - Cokelat",    1, "Piece"],
    ["Pop Ice - Cokelat",        "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Cokelat",        "Susu SKM",            30, "Gram"],
    ["Pop Ice - Cokelat",        "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Cokelat",        "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Cokelat",        "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Cokelat",        "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Cokelat",        "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Duren",          "Pop Ice - Duren",       1, "Piece"],
    ["Pop Ice - Duren",          "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Duren",          "Susu SKM",            30, "Gram"],
    ["Pop Ice - Duren",          "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Duren",          "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Duren",          "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Duren",          "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Duren",          "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Strawberry",     "Pop Ice - Strawberry",  1, "Piece"],
    ["Pop Ice - Strawberry",     "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Strawberry",     "Susu SKM",            30, "Gram"],
    ["Pop Ice - Strawberry",     "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Strawberry",     "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Strawberry",     "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Strawberry",     "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Strawberry",     "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Alpukat",        "Pop Ice - Alpukat",    1, "Piece"],
    ["Pop Ice - Alpukat",        "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Alpukat",        "Susu SKM",            30, "Gram"],
    ["Pop Ice - Alpukat",        "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Alpukat",        "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Alpukat",        "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Alpukat",        "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Alpukat",        "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Taro",           "Pop Ice - Taro",       1, "Piece"],
    ["Pop Ice - Taro",           "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Taro",           "Susu SKM",            30, "Gram"],
    ["Pop Ice - Taro",           "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Taro",           "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Taro",           "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Taro",           "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Taro",           "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Blueberry",      "Pop Ice - Blueberry",  1, "Piece"],
    ["Pop Ice - Blueberry",      "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Blueberry",      "Susu SKM",            30, "Gram"],
    ["Pop Ice - Blueberry",      "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Blueberry",      "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Blueberry",      "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Blueberry",      "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Blueberry",      "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Permen Karet",   "Pop Ice - Permen Karet",1, "Piece"],
    ["Pop Ice - Permen Karet",   "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Permen Karet",   "Susu SKM",            30, "Gram"],
    ["Pop Ice - Permen Karet",   "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Permen Karet",   "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Permen Karet",   "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Permen Karet",   "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Permen Karet",   "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Lychee",         "Pop Ice - Lychee",     1, "Piece"],
    ["Pop Ice - Lychee",         "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Lychee",         "Susu SKM",            30, "Gram"],
    ["Pop Ice - Lychee",         "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Lychee",         "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Lychee",         "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Lychee",         "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Lychee",         "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Anggur",         "Pop Ice - Anggur",     1, "Piece"],
    ["Pop Ice - Anggur",         "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Anggur",         "Susu SKM",            30, "Gram"],
    ["Pop Ice - Anggur",         "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Anggur",         "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Anggur",         "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Anggur",         "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Anggur",         "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Mangga",         "Pop Ice - Mangga",     1, "Piece"],
    ["Pop Ice - Mangga",         "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Mangga",         "Susu SKM",            30, "Gram"],
    ["Pop Ice - Mangga",         "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Mangga",         "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Mangga",         "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Mangga",         "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Mangga",         "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Melon",          "Pop Ice - Melon",      1, "Piece"],
    ["Pop Ice - Melon",          "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Melon",          "Susu SKM",            30, "Gram"],
    ["Pop Ice - Melon",          "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Melon",          "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Melon",          "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Melon",          "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Melon",          "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Cappuccino",     "Pop Ice - Cappuccino",  1, "Piece"],
    ["Pop Ice - Cappuccino",     "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Cappuccino",     "Susu SKM",            30, "Gram"],
    ["Pop Ice - Cappuccino",     "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Cappuccino",     "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Cappuccino",     "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Cappuccino",     "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Cappuccino",     "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Moccacino",      "Pop Ice - Moccacino",   1, "Piece"],
    ["Pop Ice - Moccacino",      "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Moccacino",      "Susu SKM",            30, "Gram"],
    ["Pop Ice - Moccacino",      "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Moccacino",      "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Moccacino",      "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Moccacino",      "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Moccacino",      "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Vanilla Latte",  "Pop Ice - Vanilla Latte",1,"Piece"],
    ["Pop Ice - Vanilla Latte",  "Gula Pasir",           20, "Gram"],
    ["Pop Ice - Vanilla Latte",  "Susu SKM",             30, "Gram"],
    ["Pop Ice - Vanilla Latte",  "Air (Galon)",         0.15, "Liter"],
    ["Pop Ice - Vanilla Latte",  "Es Batu Kristal",      0.2, "Kg"],
    ["Pop Ice - Vanilla Latte",  "Cup Plastik 18oz",       1, "Piece"],
    ["Pop Ice - Vanilla Latte",  "Tutup Cup Plastik",      1, "Piece"],
    ["Pop Ice - Vanilla Latte",  "Sedotan Biasa",          1, "Piece"],

    ["Pop Ice - Cookies & Krim", "Pop Ice - Cookies & Krim",1,"Piece"],
    ["Pop Ice - Cookies & Krim", "Gula Pasir",           20, "Gram"],
    ["Pop Ice - Cookies & Krim", "Susu SKM",             30, "Gram"],
    ["Pop Ice - Cookies & Krim", "Air (Galon)",         0.15, "Liter"],
    ["Pop Ice - Cookies & Krim", "Es Batu Kristal",      0.2, "Kg"],
    ["Pop Ice - Cookies & Krim", "Cup Plastik 18oz",       1, "Piece"],
    ["Pop Ice - Cookies & Krim", "Tutup Cup Plastik",      1, "Piece"],
    ["Pop Ice - Cookies & Krim", "Sedotan Biasa",          1, "Piece"],

    ["Pop Ice - Es Doger",       "Pop Ice - Es Doger",   1, "Piece"],
    ["Pop Ice - Es Doger",       "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Es Doger",       "Susu SKM",            30, "Gram"],
    ["Pop Ice - Es Doger",       "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Es Doger",       "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Es Doger",       "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Es Doger",       "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Es Doger",       "Sedotan Biasa",         1, "Piece"],

    ["Pop Ice - Es Teler",       "Pop Ice - Es Teler",   1, "Piece"],
    ["Pop Ice - Es Teler",       "Gula Pasir",          20, "Gram"],
    ["Pop Ice - Es Teler",       "Susu SKM",            30, "Gram"],
    ["Pop Ice - Es Teler",       "Air (Galon)",        0.15, "Liter"],
    ["Pop Ice - Es Teler",       "Es Batu Kristal",     0.2, "Kg"],
    ["Pop Ice - Es Teler",       "Cup Plastik 18oz",      1, "Piece"],
    ["Pop Ice - Es Teler",       "Tutup Cup Plastik",     1, "Piece"],
    ["Pop Ice - Es Teler",       "Sedotan Biasa",         1, "Piece"],

    // KOPI TUBRUK
    ["Kopi Tubruk Robusta",   "Biji Kopi Robusta",      15, "Gram"],
    ["Kopi Tubruk Robusta",   "Gula Pasir",            15, "Gram"],
    ["Kopi Tubruk Robusta",   "Air (Galon)",           0.2, "Liter"],
    ["Kopi Tubruk Robusta",   "Paper Cup 8oz",           1, "Piece"],
    ["Kopi Tubruk Robusta",   "Tutup Paper Cup",         1, "Piece"],
    ["Kopi Tubruk Robusta",   "Sedotan Biasa",           1, "Piece"],

    ["Kopi Tubruk Arabika",   "Biji Kopi Arabika",      15, "Gram"],
    ["Kopi Tubruk Arabika",   "Gula Pasir",            15, "Gram"],
    ["Kopi Tubruk Arabika",   "Air (Galon)",           0.2, "Liter"],
    ["Kopi Tubruk Arabika",   "Paper Cup 8oz",           1, "Piece"],
    ["Kopi Tubruk Arabika",   "Tutup Paper Cup",         1, "Piece"],
    ["Kopi Tubruk Arabika",   "Sedotan Biasa",           1, "Piece"],

    // ES TEH ORIGINAL
    ["Es Teh Original",       "Teh Celup",              1, "Piece"],
    ["Es Teh Original",       "Gula Pasir",            20, "Gram"],
    ["Es Teh Original",       "Es Batu Kristal",       0.2, "Kg"],
    ["Es Teh Original",       "Cup Plastik 18oz",        1, "Piece"],
    ["Es Teh Original",       "Tutup Cup Plastik",       1, "Piece"],
    ["Es Teh Original",       "Sedotan Biasa",           1, "Piece"],

    // TOPPINGS
    ["Keju",                  "Keju",                  15, "Gram"],
    ["Chocolate",             "Chocolate",             15, "Gram"],
    ["Chocochips",            "Chocochips",            10, "Gram"],
    ["Mesis",                 "Mesis",                 15, "Gram"],
    ["Bubuk Oreo",            "Bubuk Oreo",            15, "Gram"],
    ["Boba",                  "Boba",                  20, "Gram"]
  ];

  // Tulis seluruh baris data ke dalam sheet Resep (A-D)
  sh.getRange(2, 1, data.length, data[0].length).setValues(data);

  // Menerapkan gaya visual, formula harga, & format angka
  data.forEach(function(_, i) {
    var r = i + 2;
    var bg = i % 2 === 0 ? C.LIGHT : C.WHITE;
    styleData(sh.getRange(r, 1, 1, 6), bg); // style A-F sekaligus
    
    // Format desimal untuk kolom Takaran (C) agar mendukung angka seperti 0.15
    sh.getRange(r, 3).setNumberFormat("#,##0.00");
    
    // Kolom E: VLOOKUP Harga Per Piece dari sheet Bahan
    sh.getRange(r, 5).setFormula('=IFERROR(VLOOKUP(B' + r + ', Bahan!$B$2:$F$100, 5, FALSE), 0)');
    // Kolom F: Harga Per Takaran = Takaran × Harga Per Piece
    sh.getRange(r, 6).setFormula('=IFERROR(C' + r + '*E' + r + ', 0)');
    setCurrency(sh.getRange(r, 5, 1, 2));
  });

  // Set lebar kolom yang proporsional
  [220, 200, 140, 130, 140, 140].forEach(function(w, i) { sh.setColumnWidth(i+1, w); });
  sh.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — STOCK
// ═══════════════════════════════════════════════════════════════════════════
function buildStock(ss) {
  var C = getC();
  var sh = ss.insertSheet("Stock");
  setSheetFont(sh);
  sh.setTabColor(C.GREEN);

  var headers = ["Kategori","Nama Bahan","Satuan","Stok Awal","Terjual","Sisa Stok","Min. Stok","Status"];
  var hRange = sh.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  styleHeader(hRange, C.GREEN);
  sh.setRowHeight(1, 32);

  var data = [];

  data.forEach(function(row, i) {
    var r = i + 2;
    // Kolom: A=Kategori, B=Nama Bahan, C=Satuan, D=Stok Awal, E=Terjual,
    //        F=Sisa Stok, G=Min. Stok, H=Status
    var kat = row[0], nama = row[1], satuan = row[2], stokAwal = row[3], minStok = row[4];
    sh.getRange(r,1).setValue(kat);
    sh.getRange(r,2).setValue(nama);
    sh.getRange(r,3).setValue(satuan);
    sh.getRange(r,4).setValue(stokAwal);

    // var fTerjual;
    // if (kat === "Bahan Utama") {
    //   fTerjual = "=SUMIF(Transaksi!$E$3:$E$5000,B"+r+",Transaksi!$F$3:$F$5000)";
    // } else if (kat === "Topping") {
    //   fTerjual = "=SUMPRODUCT((LEN(Transaksi!$G$3:$G$5000)-LEN(SUBSTITUTE(Transaksi!$G$3:$G$5000,B"+r+",\"\")))/LEN(B"+r+")*(Transaksi!$G$3:$G$5000<>\"\")*Transaksi!$F$3:$F$5000)";
    // } else if (["Cup 18oz","Tutup Cup","Sedotan","Plastik"].indexOf(nama) >= 0) {
    //   fTerjual = "=SUM(Transaksi!$F$3:$F$5000)";
    // } else {
    //   fTerjual = "=0";
    // }
    // sh.getRange(r,5).setFormula(fTerjual);
    sh.getRange(r,6).setValue(stokAwal); // F = Sisa Stok (value, biar bisa di-write langsung)
    sh.getRange(r,7).setValue(minStok);
    sh.getRange(r,8).setFormula("=IF(F"+r+"<=G"+r+",\"⚠ RESTOCK\",\"✓ OK\")");

    var bg = i % 2 === 0 ? C.LIGHT : C.WHITE;
    styleData(sh.getRange(r,1,1,8), bg);
    sh.setRowHeight(r, 24);
  });

  // Conditional formatting untuk semua baris (data dinamis dari Pengeluaran)
  var rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("⚠ RESTOCK")
    .setBackground(C.LRED).setFontColor(C.RED).setBold(true)
    .setRanges([sh.getRange("A2:H1000")])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("✓ OK")
    .setBackground(C.LGREEN).setFontColor(C.GREEN).setBold(true)
    .setRanges([sh.getRange("A2:H1000")])
    .build());
  sh.setConditionalFormatRules(rules);

  [100,200,80,80,80,80,80,110].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — TRANSAKSI
// ═══════════════════════════════════════════════════════════════════════════
function buildTransaksi(ss) {
  var C = getC();
  var sh = ss.insertSheet("Transaksi");
  setSheetFont(sh);
  sh.setTabColor(C.PURPLE);

  sh.getRange("A1:L1").merge()
    .setValue("📋 Data Transaksi")
    .setBackground(C.PURPLE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);

  var headers = ["No. Trx","No. Item","Tanggal","Jam","Kasir","Product Variant",
                 "Jumlah Cup","Topping","Jml Topping","Harga Base","Harga Topping","Total"];
  var hRange = sh.getRange(2,1,1,headers.length);
  hRange.setValues([headers]);
  styleHeader(hRange, C.PURPLE);
  sh.setRowHeight(2, 28);

  // var samples = [
  //   ["TRX-001","23/05/2026","08:30","Kasir 1","Pop Ice - Chociato",      2,"Oreo, Keju",         2,10000, 4000,14000],
  //   ["TRX-002","23/05/2026","09:15","Kasir 1","Pop Ice - Strawberry",    1,"Mesis",              1, 5000, 1000, 6000],
  //   ["TRX-003","23/05/2026","10:00","Kasir 1","Good Day - Freez",        3,"Oreo, Cincau, Jelly",3,15000, 9000,24000]
  // ];
  // samples.forEach(function(row, i) {
  //   var r = i + 3;
  //   sh.getRange(r,1,1,row.length).setValues([row]);
  //   setCurrency(sh.getRange(r,9,1,3));
  //   styleData(sh.getRange(r,1,1,headers.length), i%2===0 ? C.LIGHT : C.WHITE);
  //   sh.setRowHeight(r, 22);
  // });

  [90,60,110,70,90,180,80,220,100,120,130,120].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(2);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — PENDAPATAN (struktur statis, data diisi oleh refreshLaporan)
// ═══════════════════════════════════════════════════════════════════════════
function buildPendapatan(ss) {
  var C = getC();
  var sh = ss.insertSheet("Pendapatan");
  setSheetFont(sh);
  sh.setTabColor(C.RED);

  // ── Judul ──────────────────────────────────────────────────────────────
  sh.getRange("A1:F1").merge()
    .setValue("💰 Laporan Pendapatan & Laba/Rugi")
    .setBackground(C.RED).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 40);

  // ── Tombol Refresh ─────────────────────────────────────────────────────
  sh.getRange("A2:F2").merge()
    .setValue("🔄 Klik menu POS → Refresh Laporan untuk update data")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 22);

  // ── Ringkasan Hari Ini ─────────────────────────────────────────────────
  sh.getRange("A4:F4").merge()
    .setValue("📊 Ringkasan Hari Ini")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(4, 30);

  var ringkasanLabels = [
    "Tanggal","Total Transaksi","Total Cup Terjual",
    "Total Pendapatan","Total HPP","Laba Bersih","Margin (%)"
  ];
  ringkasanLabels.forEach(function(lbl, i) {
    var r = i + 5;
    sh.getRange(r, 1).setValue(lbl).setFontWeight("bold").setBackground(C.LIGHT).setFontSize(10);
    sh.getRange(r, 2).setBackground(C.WHITE).setFontSize(10);
    sh.setRowHeight(r, 24);
  });
  var td = 'TEXT(TODAY(),"DD/MM/YYYY")';
  sh.getRange(5,  2).setFormula('=TEXT(TODAY(),"DD/MM/YYYY")');
  sh.getRange(6,  2).setFormula('=COUNTIF(Transaksi!$C$3:$C$5000,'+td+')');
  sh.getRange(7,  2).setFormula('=SUMIF(Transaksi!$C$3:$C$5000,'+td+',Transaksi!$G$3:$G$5000)');
  sh.getRange(8,  2).setFormula('=SUMIF(Transaksi!$C$3:$C$5000,'+td+',Transaksi!$L$3:$L$5000)')
                     .setNumberFormat('"Rp "#,##0');
  sh.getRange(9,  2).setFormula('=B7*'+HPP_PER_CUP+'+SUMPRODUCT((Transaksi!$C$3:$C$5000='+td+')*Transaksi!$I$3:$I$5000*Transaksi!$G$3:$G$5000)*'+HPP_PER_TOP)
                     .setNumberFormat('"Rp "#,##0');
  sh.getRange(10, 2).setFormula('=B8-B9').setNumberFormat('"Rp "#,##0').setBackground(C.LGREEN);
  sh.getRange(11, 2).setFormula('=IF(B8=0,0,B10/B8*100)').setNumberFormat('0.00"%"');

  // ── Status Kas ─────────────────────────────────────────────────────────
  sh.getRange("A12:C12").merge()
    .setValue("💳 Petty Cash (PC):")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange("D12:E12").merge()
    .setFormula('=IFERROR(LOOKUP(2,1/(Kas!B$5:B="PC"),Kas!F$5:F),0)')
    .setNumberFormat('"Rp "#,##0')
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(12, 24);

  sh.getRange("A13:C13").merge()
    .setValue("🛒 Uang Belanja (UB):")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange("D13:E13").merge()
    .setFormula('=IFERROR(LOOKUP(2,1/(Kas!B$5:B="UB"),Kas!F$5:F),0)')
    .setNumberFormat('"Rp "#,##0')
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(13, 24);

  // ── Rekap Harian — header (data diisi refreshLaporan) ─────────────────
  sh.getRange("A15:F15").merge()
    .setValue("📅 Rekap Harian")
    .setBackground(C.BLUE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(15, 30);

  var hHarian = ["Tanggal","Total Trx","Total Cup","Pendapatan","HPP","Laba Bersih"];
  sh.getRange(16, 1, 1, hHarian.length).setValues([hHarian])
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(16, 26);

  // Tandai baris mulai data harian
  sh.getRange("A17").setValue("").setBackground(C.WHITE);

  // ── Rekap Bulanan — header (data diisi refreshLaporan) ────────────────
  // Posisi awal bulanan akan ditulis dinamis oleh refreshLaporan
  // Simpan marker di named range A13 comment saja, posisi dihitung saat refresh

  [130,90,90,130,130,130].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPER — Hitung HPP per produk dari Resep (BOM) × Harga Bahan
// Returns: { "Nama Produk": totalHargaPerUnit, ... }
// ═══════════════════════════════════════════════════════════════════════════
function getHPPLookup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shBahan = ss.getSheetByName("Bahan");
  var shResep = ss.getSheetByName("Resep");
  if (!shBahan || !shResep) return {};

  // 1. Harga per piece dari Bahan (col F = Harga Per Piece = E/D)
  var bahanData = shBahan.getDataRange().getValues();
  var hargaMap = {};
  for (var i = 1; i < bahanData.length; i++) {
    var name = String(bahanData[i][1]).trim();
    var price = Number(bahanData[i][5]) || 0; // F = Harga Per Piece
    hargaMap[name] = price;
  }

  // 2. HPP per menu dari Resep: sum(qty ingredient × harga ingredient)
  var resepData = shResep.getDataRange().getValues();
  var hppMap = {};
  for (var i = 1; i < resepData.length; i++) {
    var menu = String(resepData[i][0]).trim();
    var bahan = String(resepData[i][1]).trim();
    var takaran = Number(resepData[i][2]) || 0;
    if (!menu || !bahan || !takaran) continue;

    if (!hppMap[menu]) hppMap[menu] = 0;
    hppMap[menu] += takaran * (hargaMap[bahan] || 0);
  }

  return hppMap; // { "Pop Ice - Chociato": 4981, "Keju": 1200, ... }
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — REFRESH LAPORAN
// Baca semua data Transaksi, hitung agregat harian & bulanan pakai BOM HPP
// ═══════════════════════════════════════════════════════════════════════════
function refreshLaporan() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var shTrx = ss.getSheetByName("Transaksi");
  var shPen = ss.getSheetByName("Pendapatan");
  var C    = getC();

  if (!shTrx || !shPen) {
    SpreadsheetApp.getUi().alert("Sheet Transaksi atau Pendapatan tidak ditemukan.");
    return;
  }

  // ── Bangun HPP lookup dari BOM (Resep × Bahan), fallback ke konstanta ─
  var hppLookup = {};
  var hppBomOk = false;
  try {
    hppLookup = getHPPLookup();
    hppBomOk = true;
  } catch(e) {
    Logger.log("getHPPLookup error: " + e.message);
    hppLookup = {};
  }

  // ── Baca semua data Transaksi (mulai baris 3) ──────────────────────────
  var lastRow = shTrx.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert("Belum ada data transaksi.");
    return;
  }

  // Kolom: A=NoTrx, B=NoItem, C=Tanggal, D=Jam, E=Kasir, F=Varian,
  //        G=JmlCup, H=Topping, I=JmlTop, J=HargaBase, K=HargaTop, L=Total
  var data = shTrx.getRange(3, 1, lastRow - 2, 12).getValues();

  // ── Agregat per hari ───────────────────────────────────────────────────
  var hariMap = {}; // key: "DD/MM/YYYY"
  data.forEach(function(row) {
    var tgl = row[2]; // kolom C (Tanggal)
    if (!tgl || tgl === "") return;
    var tglStr = typeof tgl === "string" ? tgl : Utilities.formatDate(new Date(tgl), "Asia/Jakarta", "dd/MM/yyyy");
    if (!hariMap[tglStr]) hariMap[tglStr] = {trx:0, cup:0, pendapatan:0, hppBahan:0};
    hariMap[tglStr].trx++;
    hariMap[tglStr].cup        += Number(row[6])  || 0; // G (JmlCup)
    hariMap[tglStr].pendapatan += Number(row[11]) || 0; // L (Total)

    // Hitung HPP aktual dari BOM per baris transaksi
    var varian = String(row[5]).trim();                     // F (Varian)
    var toppingStr = String(row[7]).trim();                  // H (Topping)
    var qty = Number(row[6]) || 0;                          // G (JmlCup)

    // HPP produk base
    var hppProduk = (hppLookup[varian] || HPP_PER_CUP) * qty;

    // HPP topping
    var hppTopping = 0;
    if (toppingStr) {
      var toppings = toppingStr.split(",");
      for (var ti = 0; ti < toppings.length; ti++) {
        var tName = toppings[ti].trim();
        hppTopping += (hppLookup[tName] || HPP_PER_TOP);
      }
      hppTopping *= qty;
    }

    hariMap[tglStr].hppBahan += hppProduk + hppTopping;
  });

  // ── Baca Pengeluaran per hari & bulan ──────────────────────────────────
  var shPengeluaran = ss.getSheetByName("Pengeluaran");
  if (shPengeluaran && shPengeluaran.getLastRow() >= 4) {
    var penData = shPengeluaran.getRange(4, 1, shPengeluaran.getLastRow() - 3, 7).getValues();
    penData.forEach(function(row) {
      var tgl = row[0]; // kolom A
      if (!tgl || tgl === "") return;
      var tglStr = typeof tgl === "string" ? tgl : Utilities.formatDate(new Date(tgl), "Asia/Jakarta", "dd/MM/yyyy");
      var total = Number(row[6]) || 0; // kolom G
      if (!hariMap[tglStr]) hariMap[tglStr] = {trx:0, cup:0, pendapatan:0, hppBahan:0, pengeluaran:0};
      hariMap[tglStr].pengeluaran = (hariMap[tglStr].pengeluaran || 0) + total;
    });
  }
  // Pastikan semua entry punya field pengeluaran & hppBahan
  Object.keys(hariMap).forEach(function(k) {
    if (!hariMap[k].pengeluaran) hariMap[k].pengeluaran = 0;
    if (!hariMap[k].hppBahan)    hariMap[k].hppBahan = 0;
  });

  // Sort tanggal ascending
  var hariKeys = Object.keys(hariMap).sort(function(a, b) {
    var pa = a.split("/"), pb = b.split("/");
    var da = new Date(pa[2], pa[1]-1, pa[0]);
    var db = new Date(pb[2], pb[1]-1, pb[0]);
    return da - db;
  });

  // ── Agregat per bulan ──────────────────────────────────────────────────
  var bulanMap = {}; // key: "MM/YYYY"
  var bulanNames = ["Januari","Februari","Maret","April","Mei","Juni",
                    "Juli","Agustus","September","Oktober","November","Desember"];
  hariKeys.forEach(function(tgl) {
    var parts  = tgl.split("/");
    var bulanKey = parts[1] + "/" + parts[2]; // MM/YYYY
    if (!bulanMap[bulanKey]) bulanMap[bulanKey] = {trx:0, cup:0, pendapatan:0, hppBahan:0, pengeluaran:0};
    bulanMap[bulanKey].trx        += hariMap[tgl].trx;
    bulanMap[bulanKey].cup        += hariMap[tgl].cup;
    bulanMap[bulanKey].pendapatan += hariMap[tgl].pendapatan;
    bulanMap[bulanKey].hppBahan   += hariMap[tgl].hppBahan;
    bulanMap[bulanKey].pengeluaran += (hariMap[tgl].pengeluaran || 0);
  });

  var bulanKeys = Object.keys(bulanMap).sort(function(a, b) {
    var pa = a.split("/"), pb = b.split("/");
    var da = new Date(pa[1], pa[0]-1, 1);
    var db = new Date(pb[1], pb[0]-1, 1);
    return da - db;
  });

  // ── Update Ringkasan Hari Ini (rows 5-11) dengan BOM HPP ───────────────
  var todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
  var todayData = hariMap[todayStr];
  if (todayData) {
    var hppHI = todayData.hppBahan || 0;
    var labaHI = todayData.pendapatan - hppHI;
    var marginHI = todayData.pendapatan > 0 ? (labaHI / todayData.pendapatan * 100) : 0;
    var labaBg = labaHI >= 0 ? C.LGREEN : C.LRED;
    var labaFg = labaHI >= 0 ? C.GREEN : C.RED;

    shPen.getRange(5, 2).setValue(todayStr).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(6, 2).setValue(todayData.trx).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(7, 2).setValue(todayData.cup).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(8, 2).setValue(todayData.pendapatan).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(9, 2).setValue(hppHI).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(10, 2).setValue(labaHI).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right").setBackground(labaBg).setFontColor(labaFg).setFontWeight("bold");
    shPen.getRange(11, 2).setValue(marginHI).setFontSize(10).setNumberFormat('0.00"%"').setHorizontalAlignment("center");
  } else {
    shPen.getRange(5, 2).setValue(todayStr).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(6, 2).setValue(0).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(7, 2).setValue(0).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(8, 2).setValue(0).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(9, 2).setValue(0).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(10, 2).setValue(0).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right").setBackground(C.LGREEN);
    shPen.getRange(11, 2).setValue(0).setFontSize(10).setNumberFormat('0.00"%"').setHorizontalAlignment("center");
  }

  // ── Tulis Rekap Harian ke sheet Pendapatan ─────────────────────────────
  var HARIAN_START = 17; // baris pertama data harian

  // Hapus data lama
  var lastPenRow = shPen.getLastRow();
  if (lastPenRow >= HARIAN_START) {
    shPen.getRange(HARIAN_START, 1, lastPenRow - HARIAN_START + 1, 6).clearContent().clearFormat();
  }

  hariKeys.forEach(function(tgl, i) {
    var r   = HARIAN_START + i;
    var d   = hariMap[tgl];
    // HPP = BOM-based ingredient cost ONLY (pengeluaran operasional terpisah)
    var hpp  = (d.hppBahan || 0);
    var laba = d.pendapatan - hpp;
    var bg  = i % 2 === 0 ? C.LIGHT : C.WHITE;
    var labaColor = laba >= 0 ? C.LGREEN : C.LRED;
    var labaFont  = laba >= 0 ? C.GREEN  : C.RED;

    shPen.getRange(r, 1).setValue(tgl).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, 2).setValue(d.trx).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, 3).setValue(d.cup).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, 4).setValue(d.pendapatan).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, 5).setValue(hpp).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, 6).setValue(laba).setBackground(labaColor).setFontColor(labaFont)
      .setFontWeight("bold").setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.setRowHeight(r, 22);
  });

  // ── Tulis header Rekap Bulanan ─────────────────────────────────────────
  var BULANAN_HEADER = HARIAN_START + hariKeys.length + 2;

  shPen.getRange(BULANAN_HEADER - 1, 1, 1, 6).merge()
    .setValue("📆 Rekap Bulanan")
    .setBackground(C.GREEN).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  shPen.setRowHeight(BULANAN_HEADER - 1, 30);

  var hBulanan = ["Bulan","Total Trx","Total Cup","Pendapatan","HPP","Laba Bersih"];
  shPen.getRange(BULANAN_HEADER, 1, 1, hBulanan.length).setValues([hBulanan])
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  shPen.setRowHeight(BULANAN_HEADER, 26);

  // ── Tulis Rekap Bulanan ────────────────────────────────────────────────
  bulanKeys.forEach(function(key, i) {
    var r    = BULANAN_HEADER + 1 + i;
    var d    = bulanMap[key];
    var parts = key.split("/");
    var mIdx  = parseInt(parts[0]) - 1;
    var tahun = parts[1];
    var label = bulanNames[mIdx] + " " + tahun;
    // HPP = BOM-based ingredient cost ONLY (pengeluaran operasional terpisah)
    var hpp  = (d.hppBahan || 0);
    var laba = d.pendapatan - hpp;
    var bg   = i % 2 === 0 ? C.LIGHT : C.WHITE;
    var labaColor = laba >= 0 ? C.LGREEN : C.LRED;
    var labaFont  = laba >= 0 ? C.GREEN  : C.RED;

    shPen.getRange(r, 1).setValue(label).setBackground(bg).setFontWeight("bold").setFontSize(10);
    shPen.getRange(r, 2).setValue(d.trx).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, 3).setValue(d.cup).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, 4).setValue(d.pendapatan).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, 5).setValue(hpp).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, 6).setValue(laba).setBackground(labaColor).setFontColor(labaFont)
      .setFontWeight("bold").setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.setRowHeight(r, 24);
  });

  // Hitung total sample untuk diagnostics
  var totalPendapatan = 0, totalHpp = 0, totalLaba = 0;
  hariKeys.forEach(function(tgl) {
    var d = hariMap[tgl];
    totalPendapatan += d.pendapatan;
    totalHpp += (d.hppBahan || 0);
    totalLaba += d.pendapatan - (d.hppBahan || 0);
  });

  SpreadsheetApp.flush();
  try {
    var bomStatus = hppBomOk ? "🧮 BOM HPP: Aktif" : "⚠ BOM HPP: Gagal (pakai konstanta)";
    var diag = bomStatus + "\n" +
      "Total Pendapatan: Rp " + totalPendapatan.toLocaleString("id-ID") + "\n" +
      "Total HPP: Rp " + totalHpp.toLocaleString("id-ID") + "\n" +
      "Total Laba: Rp " + totalLaba.toLocaleString("id-ID") + "\n\n" +
      "Produk di lookup: " + Object.keys(hppLookup).length + " item";
    SpreadsheetApp.getUi().alert(
      "✅ Laporan diperbarui!\n\n" +
      "Rekap Harian: " + hariKeys.length + " hari\n" +
      "Rekap Bulanan: " + bulanKeys.length + " bulan\n\n" +
      diag
    );
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — POS
// ═══════════════════════════════════════════════════════════════════════════
function buildPOS(ss) {
  var C = getC();
  var sh = ss.insertSheet("POS");
  setSheetFont(sh);
  sh.setTabColor(C.ORANGE);

  // ── Header utama ───────────────────────────────────────────────────────
  sh.getRange("A1:H1").merge()
    .setValue("🧋 CanvaPOS")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 44);

  // ── Info kasir & tanggal ───────────────────────────────────────────────
  // Baris 2: Kasir | [B] | Tanggal: | [E:F merged, format panjang] | Jam: | [H]
  sh.getRange("A2").setValue("Kasir:").setFontWeight("bold").setFontSize(10);
  sh.getRange("B2:C2").merge().setValue("Kasir 1").setBackground(C.INPUT).setFontSize(10);
  sh.getRange("D2").setValue("Tanggal:").setFontWeight("bold").setFontSize(10);
  // Format: Ahad, 24 Mei 2026
  sh.getRange("E2:F2").merge()
    .setFormula('=TEXT(TODAY(),"DDDD, D MMMM YYYY")')
    .setBackground(C.LIGHT).setFontSize(10);
  sh.getRange("G2").setValue("Jam:").setFontWeight("bold").setFontSize(10);
  sh.getRange("H2").setFormula('=TEXT(NOW(),"HH:MM")')
    .setBackground(C.LIGHT).setFontSize(10).setHorizontalAlignment("center");
  sh.setRowHeight(2, 26);

  // ── Tombol aksi ───────────────────────────────────────────────────────
  sh.getRange("A3:C3").merge()
    .setValue("💾 SIMPAN TRANSAKSI")
    .setBackground(C.GREEN).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.getRange("D3:F3").merge()
    .setValue("➕ ADD ROW")
    .setBackground(C.BLUE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.getRange("G3:H3").merge()
    .setValue("🗑 CLEAR POS")
    .setBackground(C.RED).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(3, 34);

  // ── Sub-header info harga (langsung baris 4, tanpa spacer & tanpa row kosong) ──
  sh.getRange("A4:H4").merge()
    .setValue("Base: Rp 5.000/cup  |  Topping: +Rp 1.000/jenis  |  Klik menu POS → Pilih Topping untuk tambah topping")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(4, 22);

  // ── Header kolom order (baris 5 & 6 - struktur bertingkat) ────────────
  // Baris 5: Header utama
  // Merge A5:A6 untuk "No"
  sh.getRange("A5:A6").merge()
    .setValue("No")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  // Merge B5:B6 untuk "Product Variant"
  sh.getRange("B5:B6").merge()
    .setValue("Product Variant")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  // Merge C5:C6 untuk "Jumlah Cup"
  sh.getRange("C5:C6").merge()
    .setValue("Jumlah Cup")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  // Merge D5:F5 untuk "Topping"
  sh.getRange("D5:F5").merge()
    .setValue("Topping")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  // Merge G5:G6 untuk "Harga Base"
  sh.getRange("G5:G6").merge()
    .setValue("Harga Base")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  // Merge H5:H6 untuk "Total"
  sh.getRange("H5:H6").merge()
    .setValue("Total")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  sh.setRowHeight(5, 30);
  
  // Baris 6: Sub-header untuk Topping
  sh.getRange("D6").setValue("Jenis")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  sh.getRange("E6").setValue("Jumlah")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  sh.getRange("F6").setValue("Harga Top")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  
  sh.setRowHeight(6, 26);

  // ── Dropdown validasi varian ───────────────────────────────────────────
  // (dihandle di dalam styleOrderRow)

  // ── Baris order awal (5 baris) ─────────────────────────────────────────
  for (var i = 0; i < POS_INIT_ROWS; i++) {
    var r = POS_START_ROW + i;
    var bg = i % 2 === 0 ? C.INPUT : C.WHITE;
    styleOrderRow(sh, r, i + 1, bg);
  }

  // ── Grand Total ────────────────────────────────────────────────────────
  var totalRow = POS_START_ROW + POS_INIT_ROWS;
  sh.getRange(totalRow, 1, 1, 7).merge()
    .setValue("GRAND TOTAL")
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange(totalRow, 8)
    .setFormula("=SUM(H"+POS_START_ROW+":H"+(totalRow-1)+")")
    .setNumberFormat('"Rp "#,##0')
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(13)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(totalRow, 36);

  // Simpan posisi GRAND TOTAL di PropertiesService agar clearPOS reliable
  PropertiesService.getDocumentProperties().setProperty("POS_GRAND_TOTAL_ROW", String(totalRow));

  // ── Lebar kolom ───────────────────────────────────────────────────────
  [40, 200, 90, 220, 70, 120, 120, 130].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(6);

  // ── Assign script ke tombol (via drawing/image tidak bisa di Apps Script) ─
  // Tombol dibuat sebagai sel — user klik sel lalu jalankan macro via menu
  // Kita buat custom menu sebagai alternatif yang reliable
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — PANDUAN
// ═══════════════════════════════════════════════════════════════════════════
function buildPanduan(ss) {
  var C = getC();
  var sh = ss.insertSheet("Panduan");
  setSheetFont(sh);
  sh.setTabColor(C.DARK);

  sh.getRange("A1:D1").merge()
    .setValue("📖 Panduan Penggunaan CanvaPOS")
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 40);

  var panduan = [
    ["","","",""],
    ["🧋 CARA PAKAI","","",""],
    ["1.","Buka sheet POS","",""],
    ["2.","Pilih Product Variant dari dropdown di kolom B","",""],
    ["3.","Isi jumlah cup di kolom C","",""],
    ["4.","Klik menu POS → Pilih Topping untuk menambah topping per baris","",""],
    ["5.","Harga otomatis terhitung di kolom F, G, H","",""],
    ["6.","Klik menu POS → Simpan Transaksi untuk menyimpan ke log","",""],
    ["7.","Klik menu POS → Clear POS untuk reset setelah transaksi selesai","",""],
    ["8.","Klik menu POS → Add Row untuk tambah baris order","",""],
    ["","","",""],
    ["💰 HARGA","","",""],
    ["Pop Ice Blender (all varian)","Rp 5.000/cup","",""],
    ["Kopi Tubruk (Robusta & Arabika)","Rp 5.000/cup","",""],
    ["Es Teh Original","Rp 5.000/cup","",""],
    ["Tambah Topping","Rp 1.000/jenis/cup","",""],
    ["","","",""],
    ["📦 HPP ESTIMASI","","",""],
    ["HPP per cup (tanpa topping)","Rp 2.200","",""],
    ["HPP per jenis topping","Rp 80","",""],
    ["","","",""],
    ["📊 SHEET","","",""],
    ["POS","Input transaksi harian","",""],
    ["Stock","Stok bahan real-time + restock alert","",""],
    ["Transaksi","Log history semua transaksi","",""],
    ["Pendapatan","Laporan laba/rugi harian & bulanan","",""],
    ["Bahan","Master data harga beli bahan","",""],
  ];

  panduan.forEach(function(row, i) {
    var r = i + 2;
    if (row[0] && row[0].match(/^[🧋💰📦📊]/)) {
      sh.getRange(r, 1, 1, 4).merge()
        .setValue(row[0])
        .setBackground(C.ORANGE).setFontColor(C.WHITE)
        .setFontWeight("bold").setFontSize(11);
      sh.setRowHeight(r, 28);
    } else {
      sh.getRange(r, 1).setValue(row[0]).setFontWeight(row[0].match(/^\d\./) ? "normal" : "bold");
      sh.getRange(r, 2).setValue(row[1]);
      sh.setRowHeight(r, 22);
    }
  });

  [40, 400, 200, 100].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
}

// ═══════════════════════════════════════════════════════════════════════════
// CUSTOM MENU — muncul otomatis saat file dibuka
// ═══════════════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧋 POS")
    .addItem("💾 Simpan Transaksi", "simpanTransaksi")
    .addItem("➕ Add Row", "addRowPOS")
    .addItem("🍬 Pilih Topping (baris aktif)", "pilihTopping")
    .addSeparator()
    .addItem("🗑 Clear POS", "clearPOS")
    .addSeparator()
    .addItem("💸 Simpan & Sync Stok (Pengeluaran)", "simpanPengeluaran")
    .addItem('💸 Sinkronisasi Dropdown Pengeluaran', 'syncDropdownPengeluaran')
    .addItem("🔄 Refresh Laporan Pendapatan", "refreshLaporan")
    .addSeparator()
    .addItem("🗑 Hapus Baris Aktif (POS)", "deleteRowPOS")
    .addSeparator()
    .addItem("➕ Tambah Resep / BOM", "showTambahResepDialog")
    .addSeparator()
    .addSeparator()
    .addItem("💰 Top Up PC ke Rp 100.000", "topUpPC")
    .addItem("🛒 Top Up UB (jika < Rp 10.000)", "topUpUB")
    .addItem("📅 Init Saldo Awal PC & UB", "initSaldoKas")
    .addItem("📅 Pilih Tanggal (Pengeluaran)", "showDatePicker")
    .addSeparator()
    .addItem("🔧 Setup Ulang (reset semua)", "setupPOS")
    .addItem("⚡ Install Auto-Fix Trigger", "setupOnEditTrigger")
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — SIMPAN TRANSAKSI
// ═══════════════════════════════════════════════════════════════════════════
function simpanTransaksi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPOS = ss.getSheetByName("POS");
  var shTrx = ss.getSheetByName("Transaksi");

  if (!shPOS || !shTrx) {
    SpreadsheetApp.getUi().alert("Sheet POS atau Transaksi tidak ditemukan. Jalankan Setup dulu.");
    return;
  }

  var kasir   = shPOS.getRange("B2").getValue() || "Kasir 1";
  var tanggal = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
  var jam     = Utilities.formatDate(new Date(), "Asia/Jakarta", "HH:mm");

  // Cari nomor TRX terakhir — scan backwards cari TRX-XXX (skip summary/separator)
  var lastTrxRow = shTrx.getLastRow();
  var nextNo = 1;
  if (lastTrxRow >= 3) {
    for (var scanR = lastTrxRow; scanR >= 3; scanR--) {
      var cellVal = shTrx.getRange(scanR, 1).getValue();
      var match = String(cellVal).match(/TRX-(\d+)/);
      if (match) { nextNo = parseInt(match[1]) + 1; break; }
    }
  }

  // ── Baca SEMUA data POS sekali (getValues force kalkulasi formula) ──
  var lastPOSRow = shPOS.getLastRow();
  var posRange = shPOS.getRange(POS_START_ROW, 1, lastPOSRow - POS_START_ROW + 1, 8);
  var posData = posRange.getValues();
  var C = getC();

  // Kumpulin semua item dulu, baru tulis sekaligus
  var batchRows = [];
  var batchFormats = []; // 0=cup, 1=base, 2=top, 3=total

  for (var ri = 0; ri < posData.length; ri++) {
    var row = posData[ri];
    var varian = String(row[1] || "").trim();
    if (!varian || varian === "" || varian.indexOf("GRAND TOTAL") >= 0) continue;

    var jumlahCup  = Number(row[2]) || 1;
    var topping    = String(row[3] || "");
    var jmlTopping = Number(row[4]) || 0;
    var hargaBase  = Number(row[6]) || 0;
    var hargaTop   = Number(row[5]) || 0;
    var total      = Number(row[7]) || 0;

    var itemNo = batchRows.length + 1; // nomor item dalam 1 transaksi mulai dari 1
    batchRows.push([
      "TRX-" + String(nextNo).padStart(3, "0"),
      itemNo,
      tanggal, jam, kasir, varian, jumlahCup, topping, jmlTopping,
      hargaBase, hargaTop, total
    ]);
    batchFormats.push([jumlahCup, hargaBase, hargaTop, total]);
  }
  // nextNo TIDAK di-increment — semua item dalam batch dapet TRX yang sama

  var saved = batchRows.length;
  if (saved === 0) {
    SpreadsheetApp.getUi().alert("Tidak ada data order di POS. Isi varian dulu ya!");
    return;
  }

  // ── Tulis SEMUA item sekaligus ──────────────────────────────────────────
  var insertRow = shTrx.getLastRow() + 1;
  var dataRange = shTrx.getRange(insertRow, 1, saved, 12);
  dataRange.setValues(batchRows);
  dataRange.offset(0, 9, saved, 3).setNumberFormat('"Rp "#,##0'); // J:L = Harga Base, Harga Top, Total

  // Zebra stripe + border satu kali
  for (var bi = 0; bi < saved; bi++) {
    var r = insertRow + bi;
    var bgColor = (r % 2 === 0) ? "#EAF4FB" : "#FFFFFF";
    shTrx.getRange(r, 1, 1, 12)
      .setBackground(bgColor).setVerticalAlignment("middle")
      .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    shTrx.setRowHeight(r, 22);
  }

  // ── Baris Summary Grand Total di Transaksi ─────────────────────────────
  var totalCup = 0, totalBase = 0, totalTop = 0, totalBayar = 0;
  for (var si = 0; si < saved; si++) {
    var item = batchFormats[si];
    totalCup   += item[0];
    totalBase  += item[1];
    totalTop   += item[2];
    totalBayar += item[3];
  }

  var sumRow = shTrx.getLastRow() + 1;
  // Label (A-F = 6 columns mencakup No.Trx s/d Kasir)
  shTrx.getRange(sumRow, 1, 1, 6).merge()
    .setValue("TOTAL TRANSAKSI — " + tanggal + " " + jam)
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  // Total Cup (G)
  shTrx.getRange(sumRow, 7).setValue(totalCup)
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setHorizontalAlignment("center");
  // Kolom Topping (H-I, kosong)
  shTrx.getRange(sumRow, 8, 1, 2).setBackground(C.DARK);
  // Total Harga Base (J)
  shTrx.getRange(sumRow, 10).setValue(totalBase)
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
  // Total Harga Topping (K)
  shTrx.getRange(sumRow, 11).setValue(totalTop)
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
  // Total Bayar (L)
  shTrx.getRange(sumRow, 12).setValue(totalBayar)
    .setBackground(C.GREEN).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
  shTrx.setRowHeight(sumRow, 28);

  // Baris pemisah kosong setelah summary
  var sepRow = sumRow + 1;
  shTrx.getRange(sepRow, 1, 1, 12).setBackground("#DDDDDD");
  shTrx.setRowHeight(sepRow, 6);

  // ── Clear POS & refresh laporan ────────────────────────────────────────
  stockEngineBOM(shPOS);

  clearPOS();
  SpreadsheetApp.flush();
  refreshLaporan();
}

// ═══════════════════════════════════════════════════════════════════════════
// ENGINE BOM — PEMOTONGAN STOK OTOMATIS (SAFE MODE)
// ═══════════════════════════════════════════════════════════════════════════
function stockEngineBOM(shPOS) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shResep = ss.getSheetByName("Resep");
  var shStock = ss.getSheetByName("Stock");
  if (!shResep || !shStock) return;

  var lastPosRow = shPOS.getLastRow();
  var startRow = typeof POS_START_ROW !== 'undefined' ? POS_START_ROW : 7;
  if (lastPosRow < startRow) return; 
  
  // Ambil data keranjang POS
  var posData = shPOS.getRange(startRow, 2, lastPosRow - startRow + 1, 3).getValues();
  var resepData = shResep.getDataRange().getValues();
  var stockData = shStock.getDataRange().getValues();
  var kebutuhanBahan = {};

  // 1. Kalkulasi Agregat Kebutuhan Bahan berdasarkan QTY Pembelian
  for (var i = 0; i < posData.length; i++) {
    var varian = posData[i][0];
    var qty = parseInt(posData[i][1]) || 0;
    var toppingStr = posData[i][2]; 
    if (!varian || qty === 0) continue;

    var itemsToFind = [varian];
    if (toppingStr) {
      var toppingArray = toppingStr.split(",").map(function(t) { return t.trim(); });
      itemsToFind = itemsToFind.concat(toppingArray);
    }

    // Cocokkan dengan master data di Tab Resep
    for (var k = 0; k < itemsToFind.length; k++) {
      for (var r = 1; r < resepData.length; r++) {
        if (resepData[r][0] === itemsToFind[k]) {
          var namaBahan = resepData[r][1];
          var takaran = parseFloat(resepData[r][2]) || 0;
          if (!kebutuhanBahan[namaBahan]) kebutuhanBahan[namaBahan] = 0;
          kebutuhanBahan[namaBahan] += (takaran * qty);
        }
      }
    }
  }

  // 2. Tembak nilai potongan ke Kolom "Terjual" (Kolom E) di tab Stock
  for (var s = 1; s < stockData.length; s++) {
    var namaBahanStock = stockData[s][1]; 
    if (kebutuhanBahan[namaBahanStock]) {
      // Update E (Terjual)
      var terjualAwal = parseFloat(stockData[s][4]) || 0; 
      var terjualBaru = terjualAwal + kebutuhanBahan[namaBahanStock];
      shStock.getRange(s + 1, 5).setValue(terjualBaru);
      // Update F (Sisa Stok) — kurangi dengan pemakaian
      var sisaAwal = parseFloat(stockData[s][5]) || parseFloat(stockData[s][3]) || 0;
      var sisaBaru = sisaAwal - kebutuhanBahan[namaBahanStock];
      shStock.getRange(s + 1, 6).setValue(sisaBaru);
      stockData[s][5] = sisaBaru;
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — ADD ROW
// ═══════════════════════════════════════════════════════════════════════════
function addRowPOS() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName("POS");
  var C   = getC();

  // Cari baris GRAND TOTAL — deteksi via kolom H yang punya formula SUM
  var lastRow = sh.getLastRow();
  var totalRow = -1;
  for (var r = POS_START_ROW; r <= lastRow; r++) {
    var formula = sh.getRange(r, 8).getFormula();
    if (formula.indexOf("SUM(H") >= 0) { totalRow = r; break; }
  }
  if (totalRow < 0) totalRow = lastRow + 1;

  // Hitung nomor urut: jumlah baris order yang sudah ada
  var orderCount = totalRow - POS_START_ROW;
  var newNo = orderCount + 1;
  var bg = (newNo % 2 === 0) ? C.WHITE : C.INPUT;

  // Insert baris baru sebelum GRAND TOTAL
  sh.insertRowBefore(totalRow);
  var newR = totalRow; // baris baru ada di posisi totalRow setelah insert

  styleOrderRow(sh, newR, newNo, bg);

  // Update formula GRAND TOTAL (sekarang ada di totalRow+1 setelah insert)
  var gtRow = totalRow + 1;
  sh.getRange(gtRow, 8).setFormula(
    "=SUM(H"+POS_START_ROW+":H"+(gtRow-1)+")"
  );
  // Update property
  PropertiesService.getDocumentProperties().setProperty("POS_GRAND_TOTAL_ROW", String(gtRow));

  ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange(newR, 2));
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — CLEAR POS
// ═══════════════════════════════════════════════════════════════════════════
function clearPOS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("POS");
  var C  = getC();

  var lastRow = sh.getLastRow();

  // Hapus SEMUA baris dari POS_START_ROW sampai akhir sheet (bersih total)
  if (lastRow >= POS_START_ROW) {
    sh.deleteRows(POS_START_ROW, lastRow - POS_START_ROW + 1);
  }

  // Rebuild 5 baris order awal
  for (var i = 0; i < POS_INIT_ROWS; i++) {
    var r = POS_START_ROW + i;
    var bg = i % 2 === 0 ? C.INPUT : C.WHITE;
    styleOrderRow(sh, r, i + 1, bg);
  }

  // Rebuild GRAND TOTAL row
  var gtRow = POS_START_ROW + POS_INIT_ROWS;
  sh.getRange(gtRow, 1, 1, 7).merge()
    .setValue("GRAND TOTAL")
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange(gtRow, 8)
    .setFormula("=SUM(H"+POS_START_ROW+":H"+(gtRow-1)+")")
    .setNumberFormat('"Rp "#,##0')
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(13)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(gtRow, 36);

  // Update property
  PropertiesService.getDocumentProperties().setProperty("POS_GRAND_TOTAL_ROW", String(gtRow));
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — PILIH TOPPING (popup checklist via UI)
// ═══════════════════════════════════════════════════════════════════════════
function pilihTopping() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName("POS");
  var ui  = SpreadsheetApp.getUi();

  // Cek baris aktif
  // Pastikan variabel POS_START_ROW sudah terdefinisi di global (biasanya nilainya 7)
  var activeRow = ss.getActiveRange().getRow();
  var startRow = typeof POS_START_ROW !== 'undefined' ? POS_START_ROW : 7;
  
  if (activeRow < startRow) {
    ui.alert("Pilih dulu baris order yang ingin ditambah topping (baris " + startRow + " ke bawah).");
    return;
  }

  // Cek apakah baris ini ada varian-nya
  var varian = sh.getRange(activeRow, 2).getValue();
  if (!varian) {
    ui.alert("Pilih Product Variant dulu di kolom B sebelum menambah topping.");
    return;
  }

  // ─── FIX 1: Panggil daftar Topping secara dinamis dari tab Bahan ───
  var LIST_TOPPING_DINAMIS = getToppingList();

  // Baca topping yang sudah ada
  var existing = sh.getRange(activeRow, 4).getValue() || "";
  var existingList = existing ? existing.split(",").map(function(t) { return t.trim(); }) : [];

  // Tampilkan prompt dengan daftar topping bernomor
  var msg = "Topping tersedia:\n";
  
  // ─── FIX 2: Gunakan LIST_TOPPING_DINAMIS ───
  LIST_TOPPING_DINAMIS.forEach(function(t, i) {
    var check = existingList.indexOf(t) >= 0 ? "✓" : " ";
    msg += "["+check+"] " + (i+1) + ". " + t + "\n";
  });
  msg += "\nKetik nomor topping yang dipilih, pisahkan koma.\nContoh: 1,3,5\n\n";
  msg += "(Kosongkan untuk hapus semua topping)";

  var resp = ui.prompt("🍬 Pilih Topping — Baris " + (activeRow - startRow + 1), msg, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var input = resp.getResponseText().trim();

  if (input === "") {
    sh.getRange(activeRow, 4).setValue("");
    return;
  }

  var selected = [];
  input.split(",").forEach(function(s) {
    var idx = parseInt(s.trim()) - 1;
    // ─── FIX 3: Gunakan LIST_TOPPING_DINAMIS ───
    if (idx >= 0 && idx < LIST_TOPPING_DINAMIS.length) {
      var t = LIST_TOPPING_DINAMIS[idx];
      if (selected.indexOf(t) < 0) selected.push(t);
    }
  });

  sh.getRange(activeRow, 4).setValue(selected.join(", "));
  ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange(activeRow, 4));
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — PENGELUARAN (input langsung di sheet, dropdown dari Bahan)
// ═══════════════════════════════════════════════════════════════════════════
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
  sh.getRange("E2:H2").merge()
    .setValue("ℹ Isi baris baru di bawah → klik menu POS → Simpan & Sync Stok")
    .setBackground(C.LIGHT).setFontColor(C.DARK).setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
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

  // ── Date picker validation (setelah sample data biar ga error) ──────
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true) // izinin data existing
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
function simpanPengeluaran() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var C    = getC();
  var shPen = ss.getSheetByName("Pengeluaran");
  var shStk = ss.getSheetByName("Stock");

  if (!shPen || !shStk) {
    ui.alert("Sheet Pengeluaran atau Stock tidak ditemukan.");
    return;
  }

  var lastRow = shPen.getLastRow();
  if (lastRow < 4) { ui.alert("Belum ada data pengeluaran."); return; }

  var KATEGORI_STOK = ["Bahan Utama","Topping","Kemasan","Bahan Pendukung"];
  var synced = 0, skipped = 0;

  // Baca Bahan untuk konversi satuan (kolom D = ukuran per pack)
  var shBahan = ss.getSheetByName("Bahan");
  var bahanPackMap = {}; // item name → pack size
  if (shBahan) {
    var bahanData = shBahan.getDataRange().getValues();
    for (var bi = 1; bi < bahanData.length; bi++) {
      var bNama = String(bahanData[bi][1]).trim().toLowerCase();
      var bPack = Number(bahanData[bi][3]) || 1; // D = ukuran per pack
      bahanPackMap[bNama] = bPack;
    }
  }

  // Baca semua data pengeluaran sekaligus
  var penData = shPen.getRange(4, 1, lastRow - 3, 8).getValues();
  var stockLastRow = shStk.getLastRow();
  var stockData = stockLastRow > 1
    ? shStk.getRange(2, 1, stockLastRow - 1, 7).getValues()
    : [];

  penData.forEach(function(row, i) {
    var r = i + 4;
    var tanggal  = row[0];
    var kategori = String(row[1]).trim();
    var namaItem = String(row[2]).trim();
    var jumlah   = Number(row[4]) || 0;
    var status   = String(row[7]).trim(); // kolom H = Status Stok

    // Skip baris kosong atau sudah di-sync
    if (!namaItem || !tanggal) return;
    if (status === "✓ Synced") { skipped++; return; }

    // Tandai sudah di-sync
    var bg = i % 2 === 0 ? C.LIGHT : C.WHITE;
    styleData(shPen.getRange(r, 1, 1, 8), bg);
    shPen.getRange(r, 6).setNumberFormat('"Rp "#,##0');

    // Sync ke Stock kalau kategori bahan
    if (KATEGORI_STOK.indexOf(kategori) >= 0 && jumlah > 0) {
      // Konversi satuan: jumlah × pack size dari Bahan
      var packSize = bahanPackMap[namaItem.toLowerCase()] || 1;
      var qtyStok  = jumlah * packSize;

      var found = false;
      for (var j = 0; j < stockData.length; j++) {
        var namaBahan = String(stockData[j][1]).trim().toLowerCase();
        if (namaBahan === namaItem.toLowerCase()) {
          var stokRow = j + 2;
          var stokLama = shStk.getRange(stokRow, 6).getValue() || 0;
          shStk.getRange(stokRow, 6).setValue(stokLama + qtyStok);
          stockData[j][5] = stokLama + qtyStok;
          found = true;
          break;
        }
      }
      if (!found) {
        // Item belum ada di Stock — tambah baris baru
        var satuan = String(row[3]).trim();
        var newRow = shStk.getLastRow() + 1;
        shStk.getRange(newRow, 1).setValue(kategori);
        shStk.getRange(newRow, 2).setValue(namaItem);
        shStk.getRange(newRow, 3).setValue(satuan);
        shStk.getRange(newRow, 4).setValue(qtyStok); // Stok Awal
        shStk.getRange(newRow, 5).setValue(0);        // Terjual
        shStk.getRange(newRow, 6).setValue(qtyStok); // Sisa Stok
        shStk.getRange(newRow, 7).setValue(1);        // Min. Stok (default 1)
        shStk.getRange(newRow, 8).setFormula("=IF(F"+newRow+"<=G"+newRow+",\"⚠ RESTOCK\",\"✓ OK\")");
        var rBg = (newRow % 2 === 0) ? C.LIGHT : C.WHITE;
        styleData(shStk.getRange(newRow, 1, 1, 8), rBg);
        shStk.setRowHeight(newRow, 24);
        stockData.push([kategori, namaItem, satuan, qtyStok, 0, qtyStok, 1]);
      }
    }

    // Tandai status di kolom H
    shPen.getRange(r, 8).setValue("✓ Synced")
      .setFontColor(C.GREEN).setFontWeight("bold").setFontSize(10);
    synced++;
  });

  SpreadsheetApp.flush();
  refreshLaporan();

  var msg = "✅ Selesai!\n\n" +
    "Baris diproses : " + synced + "\n" +
    "Sudah di-sync  : " + skipped + " (dilewati)";
  if (synced > 0) {
    msg += "\n\nItem baru otomatis ditambahkan ke Stock.\n" +
      "Jumlah dikonversi otomatis berdasarkan pack size dari Bahan.";
  }
  try { ui.alert(msg); } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — HAPUS BARIS AKTIF DI POS (support multi-row selection)
// ═══════════════════════════════════════════════════════════════════════════
function deleteRowPOS() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName("POS");
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

  // Validasi: semua baris harus di area order
  if (firstRow < POS_START_ROW || lastSelRow >= gtRow) {
    ui.alert(
      "Pilih baris order yang ingin dihapus.\n" +
      "Baris order dimulai dari baris " + POS_START_ROW +
      " sampai sebelum GRAND TOTAL."
    );
    return;
  }

  // Konfirmasi
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

  // Hapus dari bawah ke atas agar index tidak bergeser
  for (var r = lastSelRow; r >= firstRow; r--) {
    sh.deleteRow(r);
  }

  _renumberAndFixPOS(sh);
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
      // Pastikan formula SUM mencakup semua baris order
      sh.getRange(r, 8).setFormula("=SUM(H" + POS_START_ROW + ":H" + (r - 1) + ")");
      break;
    }

    // Cek apakah ini baris order (ada dropdown di kolom B)
    var colAVal = sh.getRange(r, 1).getValue();
    if (colAVal === "" && sh.getRange(r, 3).getValue() === "") continue; // skip baris kosong aneh

    // Update nomor urut
    sh.getRange(r, 1).setValue(no)
      .setHorizontalAlignment("center").setBackground(C.LIGHT)
      .setFontColor(C.DARK).setFontWeight("bold").setFontSize(10);

    // Pastikan formula di kolom E-H masih benar (referensi baris bisa bergeser)
    sh.getRange(r, 5).setFormula(
    "=IF(B"+r+"=\"\",\"\",IF(D"+r+"=\"\",0,COUNTA(SPLIT(REGEXREPLACE(TRIM(D"+r+"),\"\\s*,\\s*$\",\"\"),\",\"))))"
    );
    // Kolom F = Harga Top, Kolom G = Harga Base (sinkron dengan styleOrderRow)
    sh.getRange(r, 6).setFormula("=IF(B"+r+"=\"\",\"\",E"+r+"*C"+r+"*"+HARGA_TOPPING+")");
    sh.getRange(r, 7).setFormula("=IF(B"+r+"=\"\",\"\",C"+r+"*"+HARGA_BASE+")");
    sh.getRange(r, 8).setFormula("=IF(B"+r+"=\"\",\"\",G"+r+"+F"+r+")");

    no++;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// TRIGGER — onEdit Pengeluaran: dropdown + auto-fill + status stok
// ═══════════════════════════════════════════════════════════════════════════
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
  var shBahan = ss.getSheetByName("Bahan");
  var shStock = ss.getSheetByName("Stock");

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
  var shBahan = ss.getSheetByName("Bahan");
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

// ── onEdit utama: routing ke handler yang sesuai ─────────────────────────
function onEdit(e) {
  if (!e || !e.source) return;
  var shName = e.source.getActiveSheet().getName();

  // ── 1. Jika yang diedit adalah tab Bahan, langsung sinkronkan POS ──
  if (shName === "Bahan") { 
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
}

// ── Fungsi untuk mereset dropdown di Tab POS secara instan ──
function syncDropdownPOS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPOS = ss.getSheetByName("POS");
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

function syncDropdownPengeluaran() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPen = ss.getSheetByName("Pengeluaran");
  var shStock = ss.getSheetByName("Stock");
  
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
// SHEET — KAS (Petty Cash & Uang Belanja tracking)
// ═══════════════════════════════════════════════════════════════════════════
function buildKas(ss) {
  var C = getC();
  var sh = ss.insertSheet("Kas");
  setSheetFont(sh);
  sh.setTabColor("#F1C40F");

  // ── Title ──────────────────────────────────────────────────────────────
  sh.getRange("A1:F1").merge()
    .setValue("💰 Kas Harian — Petty Cash & Uang Belanja")
    .setBackground("#F1C40F").setFontColor(C.DARK)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 40);

  // ── Saldo Saat Ini ────────────────────────────────────────────────────
  sh.getRange("A2:C2").merge()
    .setValue("💳 Petty Cash (PC):")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange("D2:E2").merge()
    .setValue("—")
    .setBackground(C.WHITE).setFontColor(C.GREEN)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("F2").setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(2, 28);

  sh.getRange("A3:C3").merge()
    .setValue("🛒 Uang Belanja (UB):")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange("D3:E3").merge()
    .setValue("—")
    .setBackground(C.WHITE).setFontColor(C.GREEN)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("F3").setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(3, 28);

  // ── Headers ──────────────────────────────────────────────────────────
  var headers = ["Tanggal","Kategori","Jenis","Keterangan","Jumlah","Saldo"];
  sh.getRange(4, 1, 1, headers.length).setValues([headers])
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(4, 28);

  // ── Data kosong — isi via Init Saldo Awal di menu POS ──────────────

  [110,100,130,200,130,130].forEach(function(w, i) { sh.setColumnWidth(i+1, w); });
  sh.setFrozenRows(4);
}

// ── Helper: baca saldo terakhir dari Kas (rekalkulasi dari awal) ──────
function _getSaldoKas(kategori) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Kas");
  if (!sh) return 0;
  var lastRow = sh.getLastRow();
  if (lastRow < 5) return 0;
  var data = sh.getRange(5, 1, lastRow - 4, 5).getValues(); // A-E
  var saldo = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === kategori) {
      var jenis = String(data[i][2]);
      var jumlah = Number(data[i][4]) || 0;
      if (["Saldo Awal","Top Up","Setor"].indexOf(jenis) >= 0) {
        saldo += jumlah;
      } else {
        saldo -= jumlah;
      }
    }
  }
  return saldo;
}

function getSaldoPC() { return _getSaldoKas("PC"); }
function getSaldoUB() { return _getSaldoKas("UB"); }

// ── Helper: catat transaksi ke Kas ────────────────────────────────────
function _catatKas(kategori, jenis, keterangan, jumlah) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Kas");
  if (!sh) return;
  var C = getC();
  var today = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
  var saldoLama = _getSaldoKas(kategori);

  var isDebit = ["Saldo Awal","Top Up","Setor"].indexOf(jenis) >= 0;
  var saldoBaru = isDebit ? saldoLama + Math.abs(jumlah) : saldoLama - Math.abs(jumlah);

  var lastRow = sh.getLastRow() + 1;
  var newRow = [today, kategori, jenis, keterangan, Math.abs(jumlah), saldoBaru];
  sh.getRange(lastRow, 1, 1, newRow.length).setValues([newRow]);
  sh.getRange(lastRow, 5, 1, 2).setNumberFormat('"Rp "#,##0');
  styleData(sh.getRange(lastRow, 1, 1, 6), lastRow % 2 === 0 ? C.LIGHT : C.WHITE);
  sh.setRowHeight(lastRow, 22);

  _updateSaldoDisplay(sh, kategori, saldoBaru);
}

// ── Helper: update display saldo ──────────────────────────────────────
function _updateSaldoDisplay(sh, kategori, saldo) {
  var row = kategori === "PC" ? 2 : 3;
  var C = getC();
  sh.getRange("D" + row + ":E" + row).merge()
    .setValue(saldo)
    .setNumberFormat('"Rp "#,##0')
    .setFontColor(saldo >= 0 ? C.GREEN : C.RED)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
}

// ── Macro: Top Up PC ke Rp 100.000 ─────────────────────────────────────
function topUpPC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Kas");
  if (!sh) { SpreadsheetApp.getUi().alert("Sheet Kas belum ada. Jalankan Setup POS dulu."); return; }
  var saldo = getSaldoPC();
  var kekurangan = 100000 - saldo;
  if (kekurangan <= 0) {
    SpreadsheetApp.getUi().alert("✅ Saldo PC sudah Rp 100.000 atau lebih. Tidak perlu top up.");
    return;
  }
  _catatKas("PC", "Top Up", "Top up PC — tutup kas harian", kekurangan);
  SpreadsheetApp.getUi().alert("✅ PC di-top up Rp " + kekurangan.toLocaleString("id-ID") + "\nSaldo PC sekarang: Rp 100.000");
}

// ── Macro: Top Up UB ke Rp 100.000 ─────────────────────────────────────
function topUpUB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Kas");
  if (!sh) { SpreadsheetApp.getUi().alert("Sheet Kas belum ada. Jalankan Setup POS dulu."); return; }
  var saldo = getSaldoUB();
  if (saldo >= 10000) {
    SpreadsheetApp.getUi().alert("ℹ Saldo UB masih Rp " + saldo.toLocaleString("id-ID") + " (> Rp 10.000).\nTop up hanya diperlukan jika saldo < Rp 10.000.");
    return;
  }
  var kekurangan = 100000 - saldo;
  _catatKas("UB", "Top Up", "Top up UB — saldo menipis", kekurangan);
  SpreadsheetApp.getUi().alert("✅ UB di-top up Rp " + kekurangan.toLocaleString("id-ID") + "\nSaldo UB sekarang: Rp 100.000");
}

// ── Macro: Set Saldo Awal PC & UB (panggil setiap hari baru) ───────────
function initSaldoKas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Kas");
  if (!sh) { SpreadsheetApp.getUi().alert("Sheet Kas belum ada. Jalankan Setup POS dulu."); return; }
  
  // Cek apakah hari ini sudah ada Saldo Awal
  var today = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
  var lastRow = sh.getLastRow();
  var sudahAdaPC = false, sudahAdaUB = false;
  if (lastRow >= 5) {
    var data = sh.getRange(5, 1, lastRow - 4, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      var tgl = String(data[i][0]);
      var jenis = String(data[i][2]);
      if (tgl === today && jenis === "Saldo Awal") {
        if (data[i][1] === "PC") sudahAdaPC = true;
        if (data[i][1] === "UB") sudahAdaUB = true;
      }
    }
  }
  if (sudahAdaPC && sudahAdaUB) {
    SpreadsheetApp.getUi().alert("ℹ Saldo awal hari " + today + " sudah dicatat. Lewati.");
    return;
  }
  
  var msgParts = [];
  if (!sudahAdaPC) {
    _catatKas("PC", "Saldo Awal", "Setoran awal harian", 100000);
    msgParts.push("PC: Rp 100.000");
  }
  if (!sudahAdaUB) {
    _catatKas("UB", "Saldo Awal", "Uang belanja awal", 100000);
    msgParts.push("UB: Rp 100.000");
  }
  SpreadsheetApp.getUi().alert("✅ Saldo awal dicatat:\n" + msgParts.join("\n"));
}

// ═══════════════════════════════════════════════════════════════════════════
// FIX: Install trigger untuk onEdit (fungsi yang hilang)
// ═══════════════════════════════════════════════════════════════════════════
// MACRO — TAMBAH RESEP (form popup)
// ═══════════════════════════════════════════════════════════════════════════
function showTambahResepDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shResep = ss.getSheetByName("Resep");
  var shBahan = ss.getSheetByName("Bahan");
  if (!shResep || !shBahan) {
    SpreadsheetApp.getUi().alert("Sheet Resep atau Bahan tidak ditemukan.");
    return;
  }

  // Ambil daftar menu yang sudah ada
  var resepData = shResep.getDataRange().getValues();
  var menuExists = {};
  for (var i = 1; i < resepData.length; i++) {
    if (resepData[i][0]) menuExists[String(resepData[i][0])] = true;
  }
  var menuList = Object.keys(menuExists).sort();

  // Ambil daftar bahan baku dari sheet Bahan
  var bahanData = shBahan.getDataRange().getValues();
  var bahanList = [];
  for (var j = 1; j < bahanData.length; j++) {
    if (bahanData[j][1]) bahanList.push(String(bahanData[j][1]));
  }
  bahanList.sort();

  var menuOpts = menuList.length > 0 ? menuList.join(",") : "-";
  var bahanOpts = bahanList.join(",");

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Nunito,sans-serif;padding:16px;background:#f9f9f9}' +
    'h2{color:#2C3E50;margin:0 0 12px;font-size:16px;text-align:center}' +
    'label{display:block;font-weight:bold;font-size:12px;color:#555;margin:8px 0 3px}' +
    'input,select{width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box}' +
    'select{background:#fff}' +
    '.row{border:1px solid #ddd;border-radius:6px;padding:10px;margin:8px 0;background:#fff}' +
    '.btn-add{background:#27AE60;color:#fff;border:none;border-radius:4px;padding:6px 12px;font-size:12px;cursor:pointer;margin:4px 2px}' +
    '.btn-submit{background:#E67E22;color:#fff;border:none;border-radius:4px;padding:10px 0;font-size:14px;font-weight:bold;width:100%;cursor:pointer;margin-top:10px}' +
    '.btn-remove{background:#E74C3C;color:#fff;border:none;border-radius:4px;padding:2px 8px;font-size:11px;cursor:pointer;float:right}' +
    '.hint{font-size:11px;color:#999;margin:2px 0}' +
    '</style>' +
    '<h2>➕ Tambah Resep / BOM</h2>' +
    '<form id="frm">' +
    '<label>Nama Menu / Varian</label>' +
    '<select id="menu" onchange="toggleNewMenu()">' +
    '<option value="__new__">+ Tambah Menu Baru</option>' +
    (menuList.map(function(m) { return '<option>' + m + '</option>'; }).join('')) +
    '</select>' +
    '<div id="newMenuDiv" style="display:none">' +
    '<input type="text" id="newMenu" placeholder="Nama menu baru..." style="margin-top:4px">' +
    '</div>' +
    '<div id="ingredients">' +
    '<label>Bahan Baku</label>' +
    '<div class="row" id="row0">' +
    '<select id="bahan0" class="bahan">' +
    '<option value="">— Pilih Bahan —</option>' +
    (bahanList.map(function(b) { return '<option>' + b + '</option>'; }).join('')) +
    '</select>' +
    '<div style="display:flex;gap:6px;margin-top:4px">' +
    '<input type="number" id="qty0" class="qty" placeholder="Takaran" step="0.01" min="0" style="flex:1">' +
    '<select id="satuan0" class="satuan" style="width:100px">' +
    '<option>Gram</option><option>Kg</option><option>Liter</option><option>ml</option><option>Piece</option><option>Sachet</option><option>Pack</option><option>Pcs</option>' +
    '</select>' +
    '</div>' +
    '<button type="button" class="btn-remove" onclick="removeRow(0)" style="display:none">✕</button>' +
    '</div>' +
    '</div>' +
    '<button type="button" class="btn-add" onclick="addRow()">+ Tambah Bahan</button>' +
    '<button type="button" class="btn-submit" onclick="submitForm()">💾 Simpan Resep</button>' +
    '</form>' +
    '<div class="hint">* Pastikan nama bahan baku sudah ada di sheet Bahan</div>' +
    '<script>' +
    'var rowCount = 1;' +
    'function toggleNewMenu(){var v=document.getElementById("menu").value;document.getElementById("newMenuDiv").style.display=v==="__new__"?"block":"none"}' +
    'function addRow(){var c=document.getElementById("ingredients");var r=document.createElement("div");r.className="row";r.id="row"+rowCount;' +
    'r.innerHTML=\'<select id="bahan\'+rowCount+\'" class="bahan\">\'+document.getElementById("bahan0").innerHTML+\'</select>\' +' +
    '\'<div style="display:flex;gap:6px;margin-top:4px"><input type="number" id="qty\'+rowCount+\'" class="qty" placeholder="Takaran" step="0.01" min="0" style="flex:1">\' +' +
    '\'<select id="satuan\'+rowCount+\'" class="satuan" style="width:100px">\'+document.getElementById("satuan0").innerHTML+\'</select></div>\' +' +
    '\'<button type="button" class="btn-remove" onclick="removeRow(\'+rowCount+\')">✕</button>\';' +
    'c.appendChild(r);rowCount++}' +
    'function removeRow(i){var e=document.getElementById("row"+i);if(e){var rows=document.querySelectorAll(".row");if(rows.length>1){e.remove()}else{alert("Minimal 1 bahan baku")}}}' +
    'function submitForm(){var menu=document.getElementById("menu").value;if(menu=="__new__"){menu=document.getElementById("newMenu").value.trim()}if(!menu){alert("Isi nama menu!");return}' +
    'var rows=document.querySelectorAll(".row");var bahan=[],qty=[],satuan=[];var ok=false;' +
    'rows.forEach(function(r){var i=r.id.replace("row","");var b=document.getElementById("bahan"+i).value;var q=document.getElementById("qty"+i).value;var s=document.getElementById("satuan"+i).value;' +
    'if(b&&q){bahan.push(b);qty.push(q);satuan.push(s);ok=true}});if(!ok){alert("Isi minimal 1 bahan baku!");return}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()}).simpanTambahResep(menu,JSON.stringify({bahan:bahan,qty:qty,satuan:satuan}))}' +
    '</script>'
  ).setWidth(450).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "➕ Tambah Resep / BOM");
}

// ── Handler: simpan data dari dialog ke sheet Resep ──────────────────
function simpanTambahResep(namaMenu, jsonData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Resep");
  if (!sh) return;
  var data = JSON.parse(jsonData);
  var C = getC();

  var lastRow = sh.getLastRow();
  var newRows = [];
  for (var i = 0; i < data.bahan.length; i++) {
    newRows.push([namaMenu, data.bahan[i], Number(data.qty[i]), data.satuan[i]]);
  }

  sh.getRange(lastRow + 1, 1, newRows.length, 4).setValues(newRows);

  // Styling
  for (var k = 0; k < newRows.length; k++) {
    var r = lastRow + 1 + k;
    sh.getRange(r, 3).setNumberFormat("#,##0.00");
    styleData(sh.getRange(r, 1, 1, 4), (r % 2 === 0 ? C.LIGHT : C.WHITE));
    sh.setRowHeight(r, 22);
  }

  clearDynamicCache();
  SpreadsheetApp.flush();
}

// ═══════════════════════════════════════════════════════════════════════════
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
// DATE PICKER — HTML dialog untuk kolom Tanggal di Pengeluaran
// ═══════════════════════════════════════════════════════════════════════════
function showDatePicker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeCell = ss.getActiveCell();
  var sheetName = ss.getActiveSheet().getName();
  var col = activeCell.getColumn();

  if (sheetName !== "Pengeluaran" || col !== 1 || activeCell.getRow() < 4) {
    SpreadsheetApp.getUi().alert(
      "Pilih dulu cell di kolom Tanggal (baris 4+) di sheet Pengeluaran,\nlalu klik menu ini."
    );
    return;
  }

  var row = activeCell.getRow();
  showDatePickerForRow(row);
}

// ── Callback dari dialog: set tanggal ke cell ───────────────────────────
function setDateValue(row, dateStr) {
  // dateStr dari <input type="date"> format yyyy-MM-dd
  var parts = dateStr.split("-");
  if (parts.length !== 3) return;
  var d = new Date(parts[0], parts[1] - 1, parts[2]);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Pengeluaran");
  if (!sh) return;
  sh.getRange(row, 1).setValue(d).setNumberFormat("DD/MM/YYYY");
}

// ═══════════════════════════════════════════════════════════════════════════
// TRIGGER — onSelectionChange: auto show date picker di Pengeluaran
// ═══════════════════════════════════════════════════════════════════════════
function onSelectionChange(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (!sh || sh.getName() !== "Pengeluaran") return;

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Hanya untuk kolom Tanggal (A), baris data (4+), cell kosong
  if (col !== 1 || row < 4) return;
  if (e.range.getValue() !== "") return;

  // Buka date picker otomatis
  showDatePickerForRow(row);
}

// ── Date picker untuk row tertentu (dipanggil dari trigger) ─────────────
function showDatePickerForRow(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:sans-serif;text-align:center;padding:20px}' +
    'h3{margin:0 0 16px;color:#333}' +
    'input[type=date]{font-size:18px;padding:10px;border:2px solid #E67E22;' +
    'border-radius:8px;width:100%;box-sizing:border-box}' +
    '.btn{margin-top:16px;background:#E67E22;color:#fff;border:none;' +
    'padding:10px 24px;border-radius:6px;font-size:16px;cursor:pointer}' +
    '.btn:hover{background:#d35400}' +
    '</style>' +
    '<h3>📅 Pilih Tanggal</h3>' +
    '<input type="date" id="picker" autofocus>' +
    '<br><button class="btn" onclick="simpan()">✓ Simpan</button>' +
    '<script>' +
    'function simpan(){' +
    '  var val=document.getElementById("picker").value;' +
    '  if(!val){alert("Pilih tanggal dulu!");return;}' +
    '  google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '    .setDateValue(' + row + ',val);' +
    '}' +
    '</script>'
  )
  .setWidth(300)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, "📅 Pilih Tanggal");
}
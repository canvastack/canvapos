// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Legacy Monolithic File (functions moved to src/)
// All functions have been modularized into src/ subdirectory.
// This file retains the data arrays for reference.
// ═══════════════════════════════════════════════════════════════════════════

// ── Globals moved to src/Config.gs ────────────────────────────────────────
// var HARGA_BASE=5000, HARGA_TOPPING=1000, HPP_PER_CUP=2200, HPP_PER_TOP=80
// var POS_START_ROW=7, POS_INIT_ROWS=5, COL={...}, COLx(), confirmAction()
// ── Helpers moved to src/Helpers.gs ────────────────────────────────────────
// getDynamicList, getVarianList, getToppingList, clearDynamicCache, getC
// setSheetFont, styleHeader, styleData, styleOrderRow, setCurrency, deleteSheetIfExists

// ═══════════════════════════════════════════════════════════════════════════
// MAIN SETUP — moved to src/Builders.gs
// ═══════════════════════════════════════════════════════════════════════════
// setupPOS(), setup_1_Bahan(), setup_2_Resep(), ..., setup_8_Reorder()

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — BAHAN (build function moved to src/Builders.gs)
// Data array kept here for reference (also at src/data/BahanData.gs)
// ═══════════════════════════════════════════════════════════════════════════
var BAHAN_DATA = [
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
    ["Bahan Utama","Biji Kopi Robusta",            "Gram", 1000,200000],
    ["Bahan Utama","Biji Kopi Arabika",            "Gram", 1000,500000],
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

// (Full function moved to src/Builders.gs, data also at src/data/BahanData.gs)
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

// ═══════════════════════════════════════════════════════════════════════════
// REMAINING FUNCTIONS moved to:
//   src/Builders.gs — buildStock(), buildTransaksi(), buildPOS(), buildPanduan(), buildPendapatan()
//   src/Laporan.gs  — getHPPLookup(), refreshLaporan()
//   src/POS.gs      — addRowPOS(), clearPOS(), pilihTopping(), deleteRowPOS(), _renumberAndFixPOS()
//   src/Transaction.gs — getNextTrxNumber(), resetTrxCounter(), simpanTransaksi(), stockEngineBOM()
//   src/Pengeluaran.gs — buildPengeluaran(), simpanPengeluaran(), onEditPengeluaran(), _updateNamaItemDropdown(), _updateStatusStok()
//   src/Kas.gs      — buildKas(), _getSaldoKas(), getSaldoPC(), getSaldoUB(), _catatKas(), _updateSaldoDisplay(), topUpPC(), topUpUB(), initSaldoKas()
//   src/Triggers.gs — onOpen(), onEdit(), syncDropdownPOS(), syncDropdownPengeluaran(), setupOnEditTrigger(), onSelectionChange()
//   src/Dialogs.gs  — showTambahResepDialog(), simpanTambahResep(), showDatePicker(), setDateValue(), showDatePickerForRow()
// ═══════════════════════════════════════════════════════════════════════════

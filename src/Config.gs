// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Config.gs (Global Constants & Config)
// ═══════════════════════════════════════════════════════════════════════════

var HARGA_BASE    = 5000;
var HARGA_TOPPING = 1000;
var HPP_PER_CUP   = 2200;
var HPP_PER_TOP   = 80;
var POS_START_ROW = 7;
var POS_INIT_ROWS = 5;
var PAJAK_PERSEN  = 0; // Default 0% UMKM, bisa diubah manual

// ── Mapping HPP kategori ke Bahan kategori ──────────────────────────────
var KATEGORI_HPP = ["Bahan Utama", "Topping", "Bahan Pendukung", "Kemasan"];
var KATEGORI_HPP_MAP = {
  "Bahan Utama": "Bahan Utama",
  "Topping": "Topping",
  "Bahan Pendukung": "Bahan Pendukung",
  "Kemasan": "Kemasan"
};

// ── Column Index Constants ────────────────────────────────────────────────
// 0-based = index dalam array data hasil getValues()
// 1-based = kolom sheet untuk getRange(row, col) → COL.x.Y + 1
var COL = {
  BAHAN:       { KATEGORI:0, NAMA:1, SATUAN:2, PACK:3, HARGA_BELI:4, HARGA_PIECE:5 },
  RESEP:       { MENU:0, BAHAN:1, TAKARAN:2, SATUAN:3 },
  POS:         { NO:0, VARIAN:1, CUP:2, TOPPING:3, JML_TOP:4, HRG_TOP:5, HRG_BASE:6, TOTAL:7 },
  TRANSAKSI:   { NO_TRX:0, NO_ITEM:1, TGL:2, JAM:3, KASIR:4, VARIAN:5, CUP:6, TOPPING:7, JML_TOP:8, HRG_BASE:9, HRG_TOP:10, TOTAL:11 },
  STOCK:       { KATEGORI:0, NAMA:1, SATUAN:2, STOK_AWAL:3, TERJUAL:4, SISA:5, MIN:6, STATUS:7 },
  PENGELUARAN: { TGL:0, KATEGORI:1, NAMA:2, SATUAN:3, JUMLAH:4, HARGA:5, TOTAL:6, STATUS:7 },
  KAS:         { TGL:0, KATEGORI:1, JENIS:2, KET:3, JUMLAH:4, SALDO:5 },
  PENDAPATAN:  { LABEL:0, TRX:1, CUP:2, REVENUE:3, HPP_BAHAN:4, HPP_TOPPING:5, HPP_SUPPORT:6, HPP_KEMASAN:7, GROSS_PROFIT:8, OPEX:9, EBITDA:10, DEPRESIASI:11, EBIT:12, PAJAK:13, NET_PROFIT:14 },
  ASET:        { NAMA:0, KATEGORI:1, TGL_BELI:2, HARGA:3, UMUR:4, RESIDU:5, DEPRESIASI_BLN:6, AKUMULASI:7, NILAI_BUKU:8 },
};

/**
 * Konversi index 0-based ke 1-based.
 * @param {number} c - Index 0-based
 * @return {number} Index 1-based
 */
function COLx(c) { return c + 1; }

// ── Sheet Name Constants ────────────────────────────────────────────────────
var SHEET = {
  PANDUAN:     "Panduan",
  POS:         "POS",
  STOCK:       "Stock",
  TRANSAKSI:   "Transaksi",
  PENDAPATAN:  "Pendapatan",
  PENGELUARAN: "Pengeluaran",
  BAHAN:       "Bahan",
  RESEP:       "Resep",
  KAS:         "Kas",
  AUDIT:       "Audit",
  ASET:        "Aset"
};

/**
 * Tampilkan dialog konfirmasi Yes/No.
 * @param {string} message - Pesan konfirmasi
 * @param {string} [title] - Judul dialog
 * @return {boolean} True jika user klik Yes
 */
function confirmAction(message, title) {
  if (getEnv() === "development") return true;
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(title || "Konfirmasi", message, ui.ButtonSet.YES_NO);
    return result === ui.Button.YES;
  } catch(e) {
    return true;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// MULTI-BRANCH CONFIG — P3.7
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Environment profiles.
 * @return {Object} Config object for current env
 */
var _envConfig = null;

/**
 * Dapatkan nama environment saat ini dari PropertiesService.
 * Default: "production"
 * @return {string} "production", "staging", atau "development"
 */
function getEnv() {
  try {
    return PropertiesService.getDocumentProperties().getProperty("CANVAPOS_ENV") || "production";
  } catch(e) {
    return "production";
  }
}

/**
 * Set environment.
 * @param {string} env - "production", "staging", atau "development"
 */
function setEnv(env) {
  var valid = ["production", "staging", "development"];
  if (valid.indexOf(env) < 0) {
    SpreadsheetApp.getUi().alert("❌ Environment tidak valid. Pilih: " + valid.join(", "));
    return;
  }
  PropertiesService.getDocumentProperties().setProperty("CANVAPOS_ENV", env);
  _envConfig = null; // reset cache
  try {
    SpreadsheetApp.getUi().alert("✅ Environment diubah ke: " + env);
  } catch(e) {}
  auditLog("Set Environment", env);
}

/**
 * Dapatkan konfigurasi spesifik environment.
 * @return {Object} Environment config { adminEmail, skipConfirm, debug, hppMultiplier }
 */
function getEnvConfig() {
  if (_envConfig) return _envConfig;
  var env = getEnv();
  var base = {
    adminEmail: null,
    skipConfirm: false,
    debug: false,
    label: "Production"
  };
  var profiles = {
    production: {
      adminEmail: null,
      skipConfirm: false,
      debug: false,
      label: "Production"
    },
    staging: {
      adminEmail: null,
      skipConfirm: false,
      debug: true,
      label: "Staging"
    },
    development: {
      adminEmail: null,
      skipConfirm: true,
      debug: true,
      label: "Development"
    }
  };
  _envConfig = profiles[env] || base;
  return _envConfig;
}

/**
 * Cek apakah mode debug aktif.
 * @return {boolean}
 */
function isDebug() {
  return getEnvConfig().debug;
}

/**
 * Setup environment untuk sheet ini.
 * @param {string} env - Environment name
 */
function setupEnv(env) {
  setEnv(env);
  var cfg = getEnvConfig();
  var msg = "🔧 Environment: " + cfg.label + "\n" +
    "Debug: " + (cfg.debug ? "✅ ON" : "❌ OFF") + "\n" +
    "Skip Confirm: " + (cfg.skipConfirm ? "✅ ON" : "❌ OFF") + "\n\n" +
    "Named range & protection akan di-refresh.";
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
  refreshNamedRanges();
  protectAll();
}

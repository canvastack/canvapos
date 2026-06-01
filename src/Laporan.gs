// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Laporan.gs (Reports & Aggregation)
// ═══════════════════════════════════════════════════════════════════════════

// ── Row constants for Pendapatan sheet ──────────────────────────────────
var R_PEND = {
  JUDUL:      1,
  INFO:       2,
  PNL_HEADER: 4,
  PNL_DATE:   5,
  PNL_REVENUE:6,
  PNL_HPP_UTAMA:7,
  PNL_HPP_TOP:8,
  PNL_HPP_SUPPORT:9,
  PNL_HPP_KEMASAN:10,
  PNL_HPP_TOTAL:11,
  PNL_GROSS:  12,
  PNL_OPEX:   13,
  PNL_EBITDA: 14,
  PNL_DEPR:   15,
  PNL_EBIT:   16,
  PNL_PAJAK:  17,
  PNL_NET:    18,
  PNL_MARGIN: 19,
  KAS_PC:     21,
  KAS_UB:     22,
  REKAP_HARIAN_HEADER: 24,
  REKAP_HARIAN_COL: 25,
  REKAP_HARIAN_DATA: 26
};

// ── Label P&L ───────────────────────────────────────────────────────────
var PNL_LABELS = {
  REVENUE:  "💵 Total Pendapatan",
  HPP_UTAMA:"📦 HPP – Bahan Utama",
  HPP_TOP:  "🧀 HPP – Topping",
  HPP_SUP:  "⚙️ HPP – Bahan Pendukung",
  HPP_KEM:  "📋 HPP – Kemasan",
  HPP_TOTAL:"📊 Total HPP (BOM)",
  GROSS:    "📈 Laba Kotor (Gross Profit)",
  OPEX:     "💸 Biaya Operasional (OPEX)",
  EBITDA:   "📊 EBITDA",
  DEPR:     "🏭 Penyusutan (Depresiasi)",
  EBIT:     "📊 EBIT",
  PAJAK:    "🏛️ Pajak (" + PAJAK_PERSEN + "%)",
  NET:      "🏆 Laba Bersih (Net Profit)",
  MARGIN:   "📊 Margin Laba Bersih"
};

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — PENDAPATAN (Multi-level P&L: Revenue → Gross → EBITDA → EBIT → Net)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet Pendapatan (laporan laba/rugi 5-level).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function buildPendapatan(ss) {
  var C = getC();
  var sh = ss.insertSheet("Pendapatan");
  setSheetFont(sh);
  sh.setTabColor(C.RED);
  var CP = COL.PENDAPATAN;
  var R = R_PEND;
  var numPnLCols = 3; // Label + Value Rp + Value %

  // ── Judul ──────────────────────────────────────────────────────────────
  sh.getRange(1, 1, 1, 15).merge()
    .setValue("💰 Laporan Pendapatan & Laba/Rugi (5-Level P&L)")
    .setBackground(C.RED).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(R.JUDUL, 40);
  sh.getRange(2, 1, 1, 15).merge()
    .setValue("🔄 Klik menu POS → Refresh Laporan untuk update data")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(R.INFO, 22);

  // ── P&L Hari Ini (5-Level) ────────────────────────────────────────────
  sh.getRange(R.PNL_HEADER, 1, 1, 3).merge()
    .setValue("📊 P&L Hari Ini")
    .setBackground(C.ORANGE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(R.PNL_HEADER, 30);

  // Tanggal
  var td = 'TEXT(TODAY(),"DD/MM/YYYY")';
  sh.getRange(R.PNL_DATE, 1).setValue("📅 Tanggal").setFontWeight("bold").setBackground(C.LIGHT).setFontSize(10);
  sh.getRange(R.PNL_DATE, 2).setFormula('=' + td).setBackground(C.WHITE).setFontSize(10);
  sh.getRange(R.PNL_DATE, 3).setValue("").setBackground(C.WHITE);
  sh.setRowHeight(R.PNL_DATE, 22);

  // P&L rows: [label, formula/empty, % formula]
  var pnlDefs = [
    {row: R.PNL_REVENUE,      label: PNL_LABELS.REVENUE,   formula: '=SUMIF(TRX_Tgl,'+td+',TRX_Total)', bold: true },
    {row: R.PNL_HPP_UTAMA,    label: PNL_LABELS.HPP_UTAMA, formula: '' },
    {row: R.PNL_HPP_TOP,      label: PNL_LABELS.HPP_TOP,   formula: '' },
    {row: R.PNL_HPP_SUPPORT,  label: PNL_LABELS.HPP_SUP,   formula: '' },
    {row: R.PNL_HPP_KEMASAN,  label: PNL_LABELS.HPP_KEM,   formula: '' },
    {row: R.PNL_HPP_TOTAL,    label: PNL_LABELS.HPP_TOTAL, formula: '=SUM(B' + R.PNL_HPP_UTAMA + ':B' + R.PNL_HPP_KEMASAN + ')', bold: true, bg: C.LIGHT },
    {row: R.PNL_GROSS,        label: PNL_LABELS.GROSS,     formula: '=B' + R.PNL_REVENUE + '-B' + R.PNL_HPP_TOTAL, bold: true },
    {row: R.PNL_OPEX,         label: PNL_LABELS.OPEX,      formula: '' },
    {row: R.PNL_EBITDA,       label: PNL_LABELS.EBITDA,    formula: '=B' + R.PNL_GROSS + '-B' + R.PNL_OPEX, bold: true },
    {row: R.PNL_DEPR,         label: PNL_LABELS.DEPR,      formula: '' },
    {row: R.PNL_EBIT,         label: PNL_LABELS.EBIT,      formula: '=B' + R.PNL_EBITDA + '-B' + R.PNL_DEPR, bold: true },
    {row: R.PNL_PAJAK,        label: PNL_LABELS.PAJAK,     formula: '=B' + R.PNL_EBIT + '*' + PAJAK_PERSEN, bg: C.LIGHT },
    {row: R.PNL_NET,          label: PNL_LABELS.NET,       formula: '=B' + R.PNL_EBIT + '-B' + R.PNL_PAJAK, bold: true, fg: C.GREEN, bg: C.LGREEN },
    {row: R.PNL_MARGIN,       label: PNL_LABELS.MARGIN,    formula: '=IF(B' + R.PNL_REVENUE + '=0,0,B' + R.PNL_NET + '/B' + R.PNL_REVENUE + '*100)' }
  ];

  pnlDefs.forEach(function(pnl) {
    var r = pnl.row;
    sh.getRange(r, 1).setValue(pnl.label).setFontWeight("bold").setBackground(C.LIGHT).setFontSize(10);
    var valRange = sh.getRange(r, 2);
    if (pnl.formula) valRange.setFormula(pnl.formula);
    else valRange.setBackground(C.WHITE).setFontSize(10);
    if (pnl.bold) valRange.setFontWeight("bold");
    if (pnl.fg) valRange.setFontColor(pnl.fg);
    if (pnl.bg) valRange.setBackground(pnl.bg);
    valRange.setNumberFormat('"Rp "#,##0');

    var pctRange = sh.getRange(r, 3);
    if (r === R.PNL_REVENUE || r === R.PNL_NET) {
      pctRange.setValue("%").setFontSize(9).setFontColor(C.GRAY)
        .setHorizontalAlignment("center");
    }
    sh.setRowHeight(r, 24);
  });
  sh.getRange(R.PNL_MARGIN, 2).setNumberFormat('0.00"%"').setBackground(C.WHITE);
  sh.getRange(R.PNL_MARGIN, 3).setValue("");

  // ── Status Kas ─────────────────────────────────────────────────────────
  sh.getRange(R.KAS_PC, 1, 1, 2).merge()
    .setValue("💳 Petty Cash (PC):")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange(R.KAS_PC, 3, 1, 3).merge()
    .setFormula('=IFERROR(LOOKUP(2,1/(Kas!B$5:B="PC"),Kas!F$5:F),0)')
    .setNumberFormat('"Rp "#,##0').setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(R.KAS_PC, 24);

  sh.getRange(R.KAS_UB, 1, 1, 2).merge()
    .setValue("🛒 Uang Belanja (UB):")
    .setBackground(C.LIGHT).setFontColor(C.DARK)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange(R.KAS_UB, 3, 1, 3).merge()
    .setFormula('=IFERROR(LOOKUP(2,1/(Kas!B$5:B="UB"),Kas!F$5:F),0)')
    .setNumberFormat('"Rp "#,##0').setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(R.KAS_UB, 24);

  // ── Rekap Harian ─────────────────────────────────────────────────────────
  sh.getRange(R.REKAP_HARIAN_HEADER, 1, 1, 15).merge()
    .setValue("📅 Rekap Harian")
    .setBackground(C.BLUE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(R.REKAP_HARIAN_HEADER, 30);

  var colHeaders = [
    "Tanggal","Trx","Cup","Pendapatan",
    "HPP Utama","HPP Top","HPP Support","HPP Kemasan",
    "Laba Kotor","OPEX","EBITDA","Depresiasi","EBIT","Pajak","Laba Bersih"
  ];
  var hRange = sh.getRange(R.REKAP_HARIAN_COL, 1, 1, colHeaders.length);
  hRange.setValues([colHeaders])
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(R.REKAP_HARIAN_COL, 26);

  // Marker baris data — biar refreshLaporan tau posisi start
  sh.getRange(R.REKAP_HARIAN_DATA, 1).setValue("").setBackground(C.WHITE);

  // ── Column widths ─────────────────────────────────────────────────────
  [120, 60, 70, 130, 110, 100, 110, 110, 120, 110, 120, 110, 110, 90, 120]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPER — Hitung HPP per produk dari Resep (BOM) × Harga Bahan
// Returns: { "Nama Produk": totalHargaPerUnit, ... }
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Hitung HPP per produk dari BOM × harga bahan.
 * @return {Object} Map { namaProduk: hppPerUnit }
 */
function getHPPLookup() {
  var CACHE_KEY = "HPP_CACHE";
  var TTL_MS = 60 * 60 * 1000; // 1 jam
  var prop = PropertiesService.getDocumentProperties();
  var cached = prop.getProperty(CACHE_KEY);

  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      var age = Date.now() - (parsed.timestamp || 0);
      if (age < TTL_MS && parsed.data) {
        return parsed.data;
      }
    } catch(e) {}
  }

  var shBahan = getSheet(SHEET.BAHAN);
  var shResep = getSheet(SHEET.RESEP);
  if (!shBahan || !shResep) return {};

  var bahanData = shBahan.getDataRange().getValues();
  var hargaMap = {};
  var satuanMap = {};
  for (var i = 1; i < bahanData.length; i++) {
    var name = String(bahanData[i][COL.BAHAN.NAMA]).trim();
    var price = Number(bahanData[i][COL.BAHAN.HARGA_PIECE]) || 0;
    var satuan = String(bahanData[i][COL.BAHAN.SATUAN]).trim();
    hargaMap[name] = price;
    satuanMap[name] = satuan;
  }

  var resepData = shResep.getDataRange().getValues();
  var hppMap = {};
  for (var i = 1; i < resepData.length; i++) {
    var menu = String(resepData[i][COL.RESEP.MENU]).trim();
    var bahan = String(resepData[i][COL.RESEP.BAHAN]).trim();
    var takaran = Number(resepData[i][COL.RESEP.TAKARAN]) || 0;
    var unitResep = String(resepData[i][COL.RESEP.SATUAN]).trim();
    if (!menu || !bahan || !takaran) continue;

    var price = hargaMap[bahan] || 0;
    var unitBahan = satuanMap[bahan] || "";
    // Normalisasi: jika unit Resep ≠ unit Bahan, konversi takaran
    if (unitResep && unitBahan && unitResep.toLowerCase() !== unitBahan.toLowerCase()) {
      takaran = UnitConverter.convert(takaran, unitResep, unitBahan);
    }

    if (!hppMap[menu]) hppMap[menu] = 0;
    hppMap[menu] += takaran * price;
  }

  prop.setProperty(CACHE_KEY, JSON.stringify({data: hppMap, timestamp: Date.now()}));
  return hppMap;
}

/**
 * Hapus cache HPP — panggil setelah edit Bahan/Resep.
 */
function clearHPPLookupCache() {
  PropertiesService.getDocumentProperties().deleteProperty("HPP_CACHE");
}

/**
 * Dapatkan cache map breakdown HPP per kategori per menu.
 * @return {Object} { "MenuName": { "Bahan Utama": 0, "Topping": 0, ... } }
 */
function _getHPPCatCache() {
  var cacheKey = "HPP_CAT_CACHE";
  var prop = PropertiesService.getDocumentProperties();
  var cached = prop.getProperty(cacheKey);
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      var age = Date.now() - (parsed.timestamp || 0);
      if (age < 3600000 && parsed.data) return parsed.data;
    } catch(e) {}
  }
  var shResep = getSheet(SHEET.RESEP);
  var data = shResep ? shResep.getDataRange().getValues() : [];
  var hargaMap = getBahanHargaMap();
  var PR = COL.RESEP;
  var catMap = {};
  for (var i = 1; i < data.length; i++) {
    var menu = String(data[i][PR.MENU]).trim();
    var bahan = String(data[i][PR.BAHAN]).trim();
    var takaran = Number(data[i][PR.TAKARAN]) || 0;
    if (!menu || !bahan) continue;
    var harga = hargaMap[bahan] || 0;
    var cat = getCategoryForBahan(bahan);
    if (cat === "Lain-lain") cat = "Bahan Pendukung";
    if (!catMap[menu]) catMap[menu] = { "Bahan Utama":0, "Topping":0, "Bahan Pendukung":0, "Kemasan":0 };
    catMap[menu][cat] = (catMap[menu][cat] || 0) + (takaran * harga);
  }
  prop.setProperty(cacheKey, JSON.stringify({ data: catMap, timestamp: Date.now() }));
  return catMap;
}

/**
 * Dapatkan total OPEX harian dari sheet Pengeluaran.
 * @param {string} tglStr - Tanggal format "DD/MM/YYYY"
 * @return {number} Total pengeluaran
 */
function _getOPEXPerDay(tglStr) {
  var sh = getSheet(SHEET.PENGELUARAN);
  if (!sh || sh.getLastRow() < 4) return 0;
  var data = sh.getRange(4, 1, sh.getLastRow() - 3, 7).getValues();
  var total = 0;
  for (var i = 0; i < data.length; i++) {
    if (fmtDate(data[i][COL.PENGELUARAN.TGL]) === tglStr &&
        String(data[i][COL.PENGELUARAN.KATEGORI]).trim() === "Operasional") {
      total += Number(data[i][COL.PENGELUARAN.TOTAL]) || 0;
    }
  }
  return total;
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — REFRESH LAPORAN (Multi-level P&L dengan breakdown kategori)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Refresh laporan Pendapatan — P&L 5-level harian & bulanan.
 */
function refreshLaporan() {
  timeStart("refreshLaporan");
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var shTrx = getSheet(SHEET.TRANSAKSI);
  var shPen = getSheet(SHEET.PENDAPATAN);
  var C    = getC();

  if (!shTrx || !shPen) {
    SpreadsheetApp.getUi().alert("Sheet Transaksi atau Pendapatan tidak ditemukan.");
    return;
  }

  var hppCatCache = {};
  var bomOk = false;
  var hariKeys = [];
  var bulanKeys = [];

  withLock(30000, function() {
    var R  = R_PEND;
    var CT = COL.TRANSAKSI;

    try {
      hppCatCache = _getHPPCatCache();
      bomOk = true;
    } catch(e) {
      Logger.log("_getHPPCatCache error: " + e.message);
    }

    var lastRow = shTrx.getLastRow();
    if (lastRow < 3) {
      SpreadsheetApp.getUi().alert("Belum ada data transaksi.");
      return;
    }
    var data = shTrx.getRange(3, 1, lastRow - 2, 12).getValues();

    // ── Agregat per hari dengan breakdown kategori HPP ─────────────────────
    var hariMap = {};
    data.forEach(function(row) {
      var tgl = row[CT.TGL];
      if (!tgl || tgl === "") return;
      var tglStr = fmtDate(tgl);
      if (!hariMap[tglStr]) {
        hariMap[tglStr] = {
          trx:0, cup:0, revenue:0,
          hppUtama:0, hppTop:0, hppSupport:0, hppKemasan:0, hppTotal:0,
          opex:0
        };
      }
      var d = hariMap[tglStr];
      d.trx++;
      d.cup      += Number(row[CT.CUP])    || 0;
      d.revenue  += Number(row[CT.TOTAL])  || 0;

      var varian    = String(row[CT.VARIAN]).trim();
      var toppingStr = String(row[CT.TOPPING]).trim();
      var qty       = Number(row[CT.CUP]) || 0;

      if (bomOk && hppCatCache[varian]) {
        var catHPP = hppCatCache[varian];
        d.hppUtama   += (catHPP["Bahan Utama"] || 0) * qty;
        d.hppTop     += (catHPP["Topping"] || 0) * qty;
        d.hppSupport += (catHPP["Bahan Pendukung"] || 0) * qty;
        d.hppKemasan += (catHPP["Kemasan"] || 0) * qty;
      } else {
        d.hppUtama += HPP_PER_CUP * qty;
      }

      if (toppingStr && bomOk) {
        var toppings = toppingStr.split(",");
        for (var ti = 0; ti < toppings.length; ti++) {
          var tName = toppings[ti].trim();
          if (hppCatCache[tName]) {
            var tCat = hppCatCache[tName];
            d.hppUtama   += (tCat["Bahan Utama"] || 0) * qty;
            d.hppTop     += (tCat["Topping"] || 0) * qty;
            d.hppSupport += (tCat["Bahan Pendukung"] || 0) * qty;
            d.hppKemasan += (tCat["Kemasan"] || 0) * qty;
          } else {
            d.hppTop += HPP_PER_TOP * qty;
          }
        }
      } else if (toppingStr) {
        d.hppTop += toppingStr.split(",").length * HPP_PER_TOP * qty;
      }
    });

    Object.keys(hariMap).forEach(function(k) {
      var d = hariMap[k];
      d.hppTotal = d.hppUtama + d.hppTop + d.hppSupport + d.hppKemasan;
      d.opex = _getOPEXPerDay(k);
    });

    hariKeys = Object.keys(hariMap).sort(function(a, b) {
      var pa = a.split("/"), pb = b.split("/");
      return new Date(pa[2], pa[1]-1, pa[0]) - new Date(pb[2], pb[1]-1, pb[0]);
    });

    // ── Agregat per bulan ──────────────────────────────────────────────────
    var bulanMap = {};
    var bulanNames = ["Januari","Februari","Maret","April","Mei","Juni",
                      "Juli","Agustus","September","Oktober","November","Desember"];
    hariKeys.forEach(function(tgl) {
      var parts  = tgl.split("/");
      var bulanKey = parts[1] + "/" + parts[2];
      if (!bulanMap[bulanKey]) {
        bulanMap[bulanKey] = {
          trx:0, cup:0, revenue:0,
          hppUtama:0, hppTop:0, hppSupport:0, hppKemasan:0, hppTotal:0,
          opex:0
        };
      }
      var src = hariMap[tgl];
      var dst = bulanMap[bulanKey];
      dst.trx        += src.trx;
      dst.cup        += src.cup;
      dst.revenue    += src.revenue;
      dst.hppUtama   += src.hppUtama;
      dst.hppTop     += src.hppTop;
      dst.hppSupport += src.hppSupport;
      dst.hppKemasan += src.hppKemasan;
      dst.hppTotal   += src.hppTotal;
      dst.opex       += src.opex;
    });

    bulanKeys = Object.keys(bulanMap).sort(function(a, b) {
      var pa = a.split("/"), pb = b.split("/");
      return new Date(pa[1], pa[0]-1, 1) - new Date(pb[1], pb[0]-1, 1);
    });

    // ── Update P&L Hari Ini (rows 6-19) ───────────────────────────────────
    var todayStr = fmtDate(new Date());
    var todayData = hariMap[todayStr];
    var def = {trx:0, cup:0, revenue:0, hppUtama:0, hppTop:0, hppSupport:0, hppKemasan:0, hppTotal:0, opex:0};
    var t = todayData || def;
    var depresiasiHarian = getTotalDepresiasi();

    function setPnL(row, value) {
      shPen.getRange(row, 2).setValue(value)
        .setFontSize(10).setBackground(C.WHITE)
        .setHorizontalAlignment("right").setNumberFormat('"Rp "#,##0');
    }

    shPen.getRange(R.PNL_DATE, 2).setFormula('=TEXT(TODAY(),"DD/MM/YYYY")')
      .setFontSize(10).setHorizontalAlignment("center");
    setPnL(R.PNL_REVENUE,      t.revenue);
    setPnL(R.PNL_HPP_UTAMA,    t.hppUtama);
    setPnL(R.PNL_HPP_TOP,      t.hppTop);
    setPnL(R.PNL_HPP_SUPPORT,  t.hppSupport);
    setPnL(R.PNL_HPP_KEMASAN,  t.hppKemasan);

    shPen.getRange(R.PNL_HPP_TOTAL, 2).setFormula(
      '=SUM(B' + R.PNL_HPP_UTAMA + ':B' + R.PNL_HPP_KEMASAN + ')'
    ).setNumberFormat('"Rp "#,##0').setFontWeight("bold").setFontSize(10)
     .setBackground(C.LIGHT).setHorizontalAlignment("right");

    shPen.getRange(R.PNL_GROSS, 2).setFormula(
      '=B' + R.PNL_REVENUE + '-B' + R.PNL_HPP_TOTAL
    ).setNumberFormat('"Rp "#,##0').setFontWeight("bold").setFontSize(10)
     .setBackground(C.WHITE).setHorizontalAlignment("right");

    setPnL(R.PNL_OPEX,  t.opex);

    shPen.getRange(R.PNL_EBITDA, 2).setFormula(
      '=B' + R.PNL_GROSS + '-B' + R.PNL_OPEX
    ).setNumberFormat('"Rp "#,##0').setFontWeight("bold").setFontSize(10)
     .setHorizontalAlignment("right");

    setPnL(R.PNL_DEPR,  depresiasiHarian);

    shPen.getRange(R.PNL_EBIT, 2).setFormula(
      '=B' + R.PNL_EBITDA + '-B' + R.PNL_DEPR
    ).setNumberFormat('"Rp "#,##0').setFontWeight("bold").setFontSize(10)
     .setHorizontalAlignment("right");

    shPen.getRange(R.PNL_PAJAK, 2).setFormula(
      '=B' + R.PNL_EBIT + '*' + PAJAK_PERSEN
    ).setNumberFormat('"Rp "#,##0').setBackground(C.LIGHT).setFontSize(10)
     .setHorizontalAlignment("right");

    shPen.getRange(R.PNL_NET, 2).setFormula(
      '=B' + R.PNL_EBIT + '-B' + R.PNL_PAJAK
    ).setNumberFormat('"Rp "#,##0').setFontWeight("bold").setFontSize(10)
     .setFontColor(C.GREEN).setHorizontalAlignment("right");

    shPen.getRange(R.PNL_MARGIN, 2).setFormula(
      '=IF(B' + R.PNL_REVENUE + '=0,0,B' + R.PNL_NET + '/B' + R.PNL_REVENUE + '*100)'
    ).setNumberFormat('0.00"%"').setFontSize(10).setHorizontalAlignment("center");

    var labaBg = (t.revenue - t.hppTotal - t.opex - depresiasiHarian) >= 0 ? C.LGREEN : C.LRED;
    shPen.getRange(R.PNL_NET, 2).setBackground(labaBg);

    // ── Tulis Rekap Harian (15 kolom) ─────────────────────────────────────
    var DATA_ROW = R.REKAP_HARIAN_DATA;
    var lastPenRow = shPen.getLastRow();
    if (lastPenRow >= DATA_ROW) {
      shPen.getRange(DATA_ROW, 1, lastPenRow - DATA_ROW + 1, 15).clearContent().clearFormat();
    }

    hariKeys.forEach(function(tgl, i) {
      var r = DATA_ROW + i;
      var d = hariMap[tgl];
      var gp     = d.revenue - d.hppTotal;
      var ebitda = gp - d.opex;
      var ebit   = ebitda - depresiasiHarian;
      var pajak  = ebit * PAJAK_PERSEN;
      var net    = ebit - pajak;
      var bg     = applyZebraRow(shPen, r, i, 15);
      var netBg  = net >= 0 ? C.LGREEN : C.LRED;
      var netFg  = net >= 0 ? C.GREEN : C.RED;

      shPen.getRange(r, 1, 1, 15).setValues([[
        tgl, d.trx, d.cup, d.revenue,
        d.hppUtama, d.hppTop, d.hppSupport, d.hppKemasan,
        gp, d.opex, ebitda, depresiasiHarian, ebit, pajak, net
      ]]).setBackground(bg).setFontSize(10);

      shPen.getRange(r, 1).setHorizontalAlignment("center");
      shPen.getRange(r, 2, 1, 2).setHorizontalAlignment("center");
      shPen.getRange(r, 4, 1, 14).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
      shPen.getRange(r, 14).setNumberFormat('"Rp "#,##0.0').setHorizontalAlignment("right");
      shPen.getRange(r, 15).setBackground(netBg).setFontColor(netFg).setFontWeight("bold");
      shPen.setRowHeight(r, 22);
    });

    // ── Rekap Bulanan ─────────────────────────────────────────────────────
    var BULAN_HEADER = DATA_ROW + hariKeys.length + 2;

    shPen.getRange(BULAN_HEADER - 1, 1, 1, 15).merge()
      .setValue("📆 Rekap Bulanan")
      .setBackground(C.GREEN).setFontColor(C.WHITE)
      .setFontWeight("bold").setFontSize(12)
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    shPen.setRowHeight(BULAN_HEADER - 1, 30);

    var hBulan = ["Bulan","Trx","Cup","Pendapatan",
      "HPP Utama","HPP Top","HPP Support","HPP Kemasan",
      "Laba Kotor","OPEX","EBITDA","Depresiasi","EBIT","Pajak","Laba Bersih"];
    shPen.getRange(BULAN_HEADER, 1, 1, hBulan.length).setValues([hBulan])
      .setBackground(C.DARK).setFontColor(C.WHITE)
      .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    shPen.setRowHeight(BULAN_HEADER, 26);

    bulanKeys.forEach(function(key, i) {
      var r = BULAN_HEADER + 1 + i;
      var d = bulanMap[key];
      var parts = key.split("/");
      var mIdx = parseInt(parts[0]) - 1;
      var label = bulanNames[mIdx] + " " + parts[1];
      var gp     = d.revenue - d.hppTotal;
      var ebitda = gp - d.opex;
      var depr   = getTotalDepresiasi(key);
      var ebit   = ebitda - depr;
      var pajak  = ebit * PAJAK_PERSEN;
      var net    = ebit - pajak;
      var bg     = applyZebraRow(shPen, r, i, 15);
      var netBg  = net >= 0 ? C.LGREEN : C.LRED;
      var netFg  = net >= 0 ? C.GREEN : C.RED;

      shPen.getRange(r, 1, 1, 15).setValues([[
        label, d.trx, d.cup, d.revenue,
        d.hppUtama, d.hppTop, d.hppSupport, d.hppKemasan,
        gp, d.opex, ebitda, depr, ebit, pajak, net
      ]]).setBackground(bg).setFontSize(10);

      shPen.getRange(r, 1).setFontWeight("bold");
      shPen.getRange(r, 2, 1, 2).setHorizontalAlignment("center");
      shPen.getRange(r, 4, 1, 14).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
      shPen.getRange(r, 14).setNumberFormat('"Rp "#,##0.0');
      shPen.getRange(r, 15).setBackground(netBg).setFontColor(netFg).setFontWeight("bold");
      shPen.setRowHeight(r, 24);
    });

    SpreadsheetApp.flush();
    refreshDashboard();
  });
  refreshNamedRanges();

  try {
    SpreadsheetApp.getUi().alert(
      "✅ Laporan diperbarui!\n\n" +
      "Rekap Harian: " + hariKeys.length + " hari\n" +
      "Rekap Bulanan: " + bulanKeys.length + " bulan\n" +
      (bomOk ? "🧮 BOM HPP Kategori: Aktif" : "⚠ BOM HPP Kategori: Gagal (pakai konstanta)")
    );
  } catch(e) {}
  timeEnd("refreshLaporan");
}

/**
 * Hitung breakdown HPP per kategori untuk suatu varian + topping.
 * @param {string} varian - Nama menu
 * @param {number} qty - Jumlah cup
 * @param {string[]} [toppingArray] - Array topping (optional)
 * @return {Object} { "Bahan Utama": 0, "Topping": 0, "Bahan Pendukung": 0, "Kemasan": 0, total: 0 }
 */
function getHPPBreakdown(varian, qty, toppingArray) {
  var hppCat = { "Bahan Utama":0, "Topping":0, "Bahan Pendukung":0, "Kemasan":0, total:0 };
  var shResep = getSheet(SHEET.RESEP);
  var hargaMap = getBahanHargaMap();
  if (!shResep) return hppCat;
  var resepData = shResep.getDataRange().getValues();
  var PR = COL.RESEP;

  // Helper: sum BOM ingredients for a menu item
  function sumBOM(menuName) {
    for (var i = 1; i < resepData.length; i++) {
      if (resepData[i][PR.MENU] === menuName) {
        var bahan = String(resepData[i][PR.BAHAN]).trim();
        var takaran = Number(resepData[i][PR.TAKARAN]) || 0;
        var harga = hargaMap[bahan] || 0;
        var cat = getCategoryForBahan(bahan);
        if (cat === "Lain-lain") cat = "Bahan Pendukung";
        hppCat[cat] = (hppCat[cat] || 0) + (takaran * harga);
      }
    }
  }

  sumBOM(varian);

  if (toppingArray && toppingArray.length > 0) {
    toppingArray.forEach(function(t) {
      var tName = t.trim();
      if (tName) sumBOM(tName);
    });
  }

  hppCat.total = (hppCat["Bahan Utama"] + hppCat["Topping"] +
                   hppCat["Bahan Pendukung"] + hppCat["Kemasan"]) * qty;
  hppCat["Bahan Utama"] *= qty;
  hppCat["Topping"] *= qty;
  hppCat["Bahan Pendukung"] *= qty;
  hppCat["Kemasan"] *= qty;

  return hppCat;
}


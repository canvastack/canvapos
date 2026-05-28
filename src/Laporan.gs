// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Laporan.gs (Reports & Aggregation)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — PENDAPATAN (struktur statis, data diisi oleh refreshLaporan)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet Pendapatan (laporan laba/rugi).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
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
  sh.getRange(6,  2).setFormula('=COUNTIF(TRX_Tgl,'+td+')');
  sh.getRange(7,  2).setFormula('=SUMIF(TRX_Tgl,'+td+',TRX_Cup)');
  sh.getRange(8,  2).setFormula('=SUMIF(TRX_Tgl,'+td+',TRX_Total)')
                     .setNumberFormat('"Rp "#,##0');
  sh.getRange(9,  2).setFormula('=B7*'+HPP_PER_CUP+'+SUMPRODUCT((TRX_Tgl='+td+')*TRX_JmlTop*TRX_Cup)*'+HPP_PER_TOP)
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

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — REFRESH LAPORAN
// Baca semua data Transaksi, hitung agregat harian & bulanan pakai BOM HPP
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Refresh laporan Pendapatan (rekap harian & bulanan).
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

  withLock(20000, function() {
    var C = getC();

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
    var tgl = row[COL.TRANSAKSI.TGL];
    if (!tgl || tgl === "") return;
    var tglStr = fmtDate(tgl);
    if (!hariMap[tglStr]) hariMap[tglStr] = {trx:0, cup:0, pendapatan:0, hppBahan:0};
    hariMap[tglStr].trx++;
    hariMap[tglStr].cup        += Number(row[COL.TRANSAKSI.CUP])  || 0;
    hariMap[tglStr].pendapatan += Number(row[COL.TRANSAKSI.TOTAL]) || 0;

    var varian = String(row[COL.TRANSAKSI.VARIAN]).trim();
    var toppingStr = String(row[COL.TRANSAKSI.TOPPING]).trim();
    var qty = Number(row[COL.TRANSAKSI.CUP]) || 0;

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
  var shPengeluaran = getSheet(SHEET.PENGELUARAN);
  if (shPengeluaran && shPengeluaran.getLastRow() >= 4) {
    var penData = shPengeluaran.getRange(4, 1, shPengeluaran.getLastRow() - 3, 7).getValues();
    penData.forEach(function(row) {
      var tgl = row[0]; // kolom A
      if (!tgl || tgl === "") return;
      var tglStr = fmtDate(tgl);
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
  var todayStr = fmtDate(new Date());
  var todayData = hariMap[todayStr];
  if (todayData) {
    var hppHI = todayData.hppBahan || 0;
    var labaHI = todayData.pendapatan - hppHI;
    var marginHI = todayData.pendapatan > 0 ? (labaHI / todayData.pendapatan * 100) : 0;
    var labaBg = labaHI >= 0 ? C.LGREEN : C.LRED;
    var labaFg = labaHI >= 0 ? C.GREEN : C.RED;

    shPen.getRange(5, COLx(COL.PENDAPATAN.LABEL)).setValue(todayStr).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(6, COLx(COL.PENDAPATAN.LABEL)).setValue(todayData.trx).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(7, COLx(COL.PENDAPATAN.LABEL)).setValue(todayData.cup).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(8, COLx(COL.PENDAPATAN.LABEL)).setValue(todayData.pendapatan).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(9, COLx(COL.PENDAPATAN.LABEL)).setValue(hppHI).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(10, COLx(COL.PENDAPATAN.LABEL)).setValue(labaHI).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right").setBackground(labaBg).setFontColor(labaFg).setFontWeight("bold");
    shPen.getRange(11, COLx(COL.PENDAPATAN.LABEL)).setValue(marginHI).setFontSize(10).setNumberFormat('0.00"%"').setHorizontalAlignment("center");
  } else {
    shPen.getRange(5, COLx(COL.PENDAPATAN.LABEL)).setValue(todayStr).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(6, COLx(COL.PENDAPATAN.LABEL)).setValue(0).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(7, COLx(COL.PENDAPATAN.LABEL)).setValue(0).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(8, COLx(COL.PENDAPATAN.LABEL)).setValue(0).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(9, COLx(COL.PENDAPATAN.LABEL)).setValue(0).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(10, COLx(COL.PENDAPATAN.LABEL)).setValue(0).setFontSize(10).setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right").setBackground(C.LGREEN);
    shPen.getRange(11, COLx(COL.PENDAPATAN.LABEL)).setValue(0).setFontSize(10).setNumberFormat('0.00"%"').setHorizontalAlignment("center");
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
    var bg  = applyZebraRow(shPen, r, i, 6);
    var labaColor = laba >= 0 ? C.LGREEN : C.LRED;
    var labaFont  = laba >= 0 ? C.GREEN  : C.RED;

    shPen.getRange(r, COLx(COL.PENDAPATAN.LABEL)).setValue(tgl).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, COLx(COL.PENDAPATAN.TRX)).setValue(d.trx).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, COLx(COL.PENDAPATAN.CUP)).setValue(d.cup).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, COLx(COL.PENDAPATAN.PENDAPATAN)).setValue(d.pendapatan).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, COLx(COL.PENDAPATAN.HPP)).setValue(hpp).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, COLx(COL.PENDAPATAN.LABA)).setValue(laba).setBackground(labaColor).setFontColor(labaFont)
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
    var bg   = applyZebraRow(shPen, r, i, 6);
    var labaColor = laba >= 0 ? C.LGREEN : C.LRED;
    var labaFont  = laba >= 0 ? C.GREEN  : C.RED;

    shPen.getRange(r, COLx(COL.PENDAPATAN.LABEL)).setValue(label).setBackground(bg).setFontWeight("bold").setFontSize(10);
    shPen.getRange(r, COLx(COL.PENDAPATAN.TRX)).setValue(d.trx).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, COLx(COL.PENDAPATAN.CUP)).setValue(d.cup).setBackground(bg).setFontSize(10).setHorizontalAlignment("center");
    shPen.getRange(r, COLx(COL.PENDAPATAN.PENDAPATAN)).setValue(d.pendapatan).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, COLx(COL.PENDAPATAN.HPP)).setValue(hpp).setBackground(bg).setFontSize(10)
      .setNumberFormat('"Rp "#,##0').setHorizontalAlignment("right");
    shPen.getRange(r, COLx(COL.PENDAPATAN.LABA)).setValue(laba).setBackground(labaColor).setFontColor(labaFont)
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
  refreshDashboard();
  });
  refreshNamedRanges();
  timeEnd("refreshLaporan");
}

// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Aset.gs (Fixed Assets & Depreciation Tracking)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Bangun sheet Aset (fixed asset register).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function buildAset(ss) {
  var C = getC();
  var sh = ss.insertSheet("Aset");
  setSheetFont(sh);
  sh.setTabColor("#95A5A6");

  sh.getRange("A1:I1").merge()
    .setValue("📦 Daftar Aset Tetap & Penyusutan")
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);

  // ── Info bar ─────────────────────────────────────────────────────
  sh.getRange("A2:I2").merge()
    .setValue("💡 Klik menu POS → 📦 Tambah Aset Tetap untuk menambah aset. Penyusutan metode garis lurus.")
    .setBackground(C.LIGHT).setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 24);

  // ── Ringkasan Penyusutan ─────────────────────────────────────────
  sh.getRange("A3:C3").merge()
    .setValue("📊 Ringkasan Penyusutan Bulan Ini")
    .setBackground(C.PURPLE).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setVerticalAlignment("middle");
  sh.getRange("D3:E3").merge()
    .setFormula(F("=SUM(G{start}:G{end})", {start: 5, end: 1000}))
    .setNumberFormat('"Rp "#,##0')
    .setBackground(C.PURPLE).setFontColor(C.WHITE)
    .setFontSize(11).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Header ───────────────────────────────────────────────────────
  var headers = ["Nama Aset", "Kategori", "Tgl Beli", "Harga Perolehan",
                 "Umur (bln)", "Nilai Residu", "Penyusutan/bln", "Akum. Penyusutan", "Nilai Buku"];
  var hRange = sh.getRange(4, 1, 1, headers.length);
  hRange.setValues([headers]);
  styleHeader(hRange, C.DARK);
  sh.setRowHeight(4, 28);

  // ── Example row ───────────────────────────────────────────────────
  sh.getRange(5, 1).setValue("Mesin Blender");
  sh.getRange(5, 2).setValue("Peralatan");
  sh.getRange(5, 3).setValue("01/05/2026");
  sh.getRange(5, 4).setValue(350000).setNumberFormat('"Rp "#,##0');
  sh.getRange(5, 5).setValue(24);
  sh.getRange(5, 6).setValue(0).setNumberFormat('"Rp "#,##0');
  sh.getRange(5, 7).setFormula(F("=ROUND((D{row}-F{row})/E{row},0)", {row: 5}))
    .setNumberFormat('"Rp "#,##0');
  sh.getRange(5, 8).setValue(0).setNumberFormat('"Rp "#,##0');
  sh.getRange(5, 9).setFormula(F("=D{row}-H{row}", {row: 5}))
    .setNumberFormat('"Rp "#,##0');
  sh.getRange(5, 5, 1, 6).setBackground(C.INPUT).setFontSize(10);
  sh.setRowHeight(5, 24);

  // ── Column widths ─────────────────────────────────────────────────
  [180, 100, 100, 130, 80, 110, 130, 140, 130]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(4);
}

/**
 * Tambah aset tetap baru via dialog.
 */
function addAset() {
  var html = '<style>' +
    'body{font-family:monospace;padding:16px;color:#333}' +
    'h3{margin:0 0 12px;color:#2C3E50}' +
    'label{display:block;margin:8px 0 4px;font-size:12px;color:#666}' +
    'input{width:100%;padding:6px 8px;border:1px solid #ddd;border-radius:4px;font-size:13px;box-sizing:border-box}' +
    'select{width:100%;padding:6px 8px;border:1px solid #ddd;border-radius:4px;font-size:13px;box-sizing:border-box}' +
    '.btn{margin-top:16px;padding:8px 16px;background:#667eea;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px}' +
    '.btn:hover{background:#5a6fd6}' +
    '</style>' +
    '<h3>📦 Tambah Aset Tetap</h3>' +
    '<form id="frm">' +
    '<label>Nama Aset</label><input name="nama" placeholder="Mesin Blender" required>' +
    '<label>Kategori</label>' +
    '<select name="kategori">' +
    '  <option>Peralatan</option><option>Elektronik</option><option>Furniture</option><option>Kendaraan</option><option>Renovasi</option><option>Lainnya</option>' +
    '</select>' +
    '<label>Tanggal Beli (DD/MM/YYYY)</label><input name="tgl" placeholder="01/05/2026" required>' +
    '<label>Harga Perolehan (Rp)</label><input name="harga" type="number" placeholder="350000" required>' +
    '<label>Umur Ekonomis (bulan)</label><input name="umur" type="number" placeholder="24" required>' +
    '<label>Nilai Residu (Rp)</label><input name="residu" type="number" placeholder="0" value="0">' +
    '</form>' +
    '<button class="btn" onclick="simpan()">💾 Simpan Aset</button>' +
    '<script>' +
    'function simpan(){var f=document.getElementById("frm").elements;' +
    'var d={};for(var i=0;i<f.length;i++){d[f[i].name]=f[i].value}' +
    'if(!d.nama||!d.tgl||!d.harga||!d.umur){alert("Isi semua field!");return}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.addAsetSave(d)}</script>';

  var output = HtmlService.createHtmlOutput(html).setWidth(380).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(output, "📦 Tambah Aset Tetap");
}

/**
 * Simpan aset dari form dialog.
 * @param {Object} data - { nama, kategori, tgl, harga, umur, residu }
 */
function addAsetSave(data) {
  withLock(5000, function() {
    var sh = getSheet(SHEET.ASET);
    if (!sh) { notify("Sheet Aset tidak ditemukan", true); return; }
    var lastRow = sh.getLastRow();
    var r = Math.max(lastRow + 1, 5);
    var CA = COL.ASET;
    sh.getRange(r, COLx(CA.NAMA)).setValue(data.nama);
    sh.getRange(r, COLx(CA.KATEGORI)).setValue(data.kategori);
    sh.getRange(r, COLx(CA.TGL_BELI)).setValue(data.tgl);
    sh.getRange(r, COLx(CA.HARGA)).setValue(Number(data.harga)).setNumberFormat('"Rp "#,##0');
    sh.getRange(r, COLx(CA.UMUR)).setValue(Number(data.umur));
    sh.getRange(r, COLx(CA.RESIDU)).setValue(Number(data.residu)).setNumberFormat('"Rp "#,##0');
    sh.getRange(r, COLx(CA.DEPRESIASI_BLN))
      .setFormula(F("=ROUND((D{row}-F{row})/E{row},0)", {row: r}))
      .setNumberFormat('"Rp "#,##0');
    sh.getRange(r, COLx(CA.AKUMULASI)).setValue(0).setNumberFormat('"Rp "#,##0');
    sh.getRange(r, COLx(CA.NILAI_BUKU))
      .setFormula(F("=D{row}-H{row}", {row: r}))
      .setNumberFormat('"Rp "#,##0');
    sh.getRange(r, 5, 1, 6).setBackground("#FEF9E7").setFontSize(10);
    sh.setRowHeight(r, 24);
    notify("✅ Aset \"" + data.nama + "\" berhasil ditambahkan!");
    auditLog("Tambah Aset", data.nama + " — Rp " + Number(data.harga).toLocaleString("id-ID"));
  });
}

/**
 * Dapatkan total penyusutan per bulan (dari sheet Aset).
 * @param {string} [bulan] - Format "MM/YYYY", default bulan berjalan
 * @return {number} Total penyusutan
 */
function getTotalDepresiasi(bulan) {
  if (!bulan) {
    var d = new Date();
    var mm = ("0" + (d.getMonth() + 1)).slice(-2);
    var yyyy = d.getFullYear();
    bulan = mm + "/" + yyyy;
  }
  var sh = getSheet(SHEET.ASET);
  if (!sh) return 0;
  var lastRow = sh.getLastRow();
  if (lastRow < 5) return 0;
  var data = sh.getRange(5, 1, lastRow - 4, 9).getValues();
  var CA = COL.ASET;
  var total = 0;
  for (var i = 0; i < data.length; i++) {
    var harga = Number(data[i][CA.HARGA]) || 0;
    var umur = Number(data[i][CA.UMUR]) || 1;
    var residu = Number(data[i][CA.RESIDU]) || 0;
    var tglBeli = String(data[i][CA.TGL_BELI]).trim();
    var depresiasiBln = Math.round((harga - residu) / umur);
    if (tglBeli) {
      var parts = tglBeli.split("/");
      if (parts.length === 3) {
        var blnBeli = ("0" + parseInt(parts[1])).slice(-2) + "/" + parts[2];
        if (blnBeli <= bulan) total += depresiasiBln;
      }
    }
  }
  return total;
}

/**
 * Posting penyusutan bulan ini ke jurnal dan update akumulasi.
 */
function postingDepresiasi() {
  if (!confirmAction("Posting penyusutan bulan ini ke akumulasi di sheet Aset?")) return;
  withLock(10000, function() {
    var sh = getSheet(SHEET.ASET);
    if (!sh) { notify("Sheet Aset tidak ditemukan", true); return; }
    var lastRow = sh.getLastRow();
    if (lastRow < 5) { notify("Tidak ada aset untuk di-depresiasi.", true); return; }
    var CA = COL.ASET;
    var data = sh.getRange(5, 1, lastRow - 4, 9).getValues();
    var totalPosted = 0;
    for (var i = 0; i < data.length; i++) {
      if (!String(data[i][CA.NAMA]).trim()) continue;
      var harga = Number(data[i][CA.HARGA]) || 0;
      var umur = Number(data[i][CA.UMUR]) || 1;
      var residu = Number(data[i][CA.RESIDU]) || 0;
      var depresiasiBln = Math.round((harga - residu) / umur);
      if (depresiasiBln <= 0) continue;
      var akumLama = Number(data[i][CA.AKUMULASI]) || 0;
      var r = i + 5;
      sh.getRange(r, COLx(CA.AKUMULASI)).setValue(akumLama + depresiasiBln);
      totalPosted += depresiasiBln;
    }
    notify("✅ Penyusutan diposting: Rp " + totalPosted.toLocaleString("id-ID"));
    auditLog("Posting Depresiasi", "Total: Rp " + totalPosted.toLocaleString("id-ID"));
    refreshLaporan();
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// MIGRATION — Sync Modal Awal from Pengeluaran to Aset + Kas
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Sync data Modal Awal dari Pengeluaran ke Aset (14 aset tetap) + Kas (PC 100rb / UB 1jt).
 * Satu kali jalan — guard mencegah re-run.
 */
function syncModalAwalKeAsetDanKas() {
  if (!confirmAction(
    "🔄 Sync Modal Awal → Aset + Kas\n\n" +
    "• 14 item aset tetap → sheet Aset (dengan penyusutan)\n" +
    "• Petty Cash Rp 100.000 → Kas (PC)\n" +
    "• Saving Budget Rp 1.000.000 → Kas (UB)\n" +
    "• 6 item biaya (Servis, Trashbag) → Rekategori Operasional\n" +
    "• Galon + Isi → split: Galon (Aset 40rb) + Air Isi (OPEX 10rb)\n\n" +
    "Lanjutkan?", "🔄 Sync Modal Awal?"
  )) return;

  withLock(60000, function() {
    var shPen = getSheet(SHEET.PENGELUARAN);
    var shAset = getSheet(SHEET.ASET);
    var shKas = getSheet(SHEET.KAS);
    var C = getC();
    var CA = COL.ASET;

    if (!shPen || !shAset || !shKas) {
      notify("Sheet Pengeluaran / Aset / Kas tidak lengkap.", true);
      return;
    }

    // ── 1. Guard: cek apakah sudah pernah sync ──
    var lastPen = shPen.getLastRow();
    if (lastPen < 4) { notify("Tidak ada data Pengeluaran.", true); return; }
    var penData = shPen.getRange(4, 1, lastPen - 3, 8).getValues();
    for (var i = 0; i < penData.length; i++) {
      if (String(penData[i][1]).trim() === "Modal Awal" &&
          String(penData[i][7]).trim() !== "") {
        notify("✅ Data Modal Awal sudah pernah di-sync. Lewati.");
        return;
      }
    }

    // ── 2. Mapping ──
    var ASSET_MAP = {
      "Mesin Blender":     {k:"Peralatan", u:24, h:350000},
      "Termos":            {k:"Peralatan", u:12, h:135000},
      "Bangku Bakso":      {k:"Furniture", u:24, h:165000},
      "Banner":            {k:"Renovasi",  u:12, h:120000},
      "Kabel 20m":         {k:"Elektronik",u:12, h:87000},
      "Kabel Lampu":       {k:"Elektronik",u:12, h:50000},
      "Lampu 15w":         {k:"Elektronik",u:12, h:20000},
      "Lampu 50w":         {k:"Elektronik",u:12, h:45000},
      "Pisau":             {k:"Peralatan", u:12, h:20000},
      "Botol Mayo":        {k:"Peralatan", u:12, h:40000},
      "Tempat Topping":    {k:"Peralatan", u:12, h:38000},
      "Mika Jelly":        {k:"Peralatan", u:12, h:15000},
      "Laminating Menu":   {k:"Peralatan", u:12, h:15000},
    };
    var OPEX_NAMES = {
      "Servis Kompor": true, "Servis Kabel Rol": true,
      "Servis Mesin": true, "Tambal Ban": true, "Trashbag": true,
    };

    var tglMA = "02/05/2026";
    var asetRows = [];
    var updates = [];

    // ── 3. Process each Modal Awal row ──
    for (var i = 0; i < penData.length; i++) {
      var kat = String(penData[i][1]).trim();
      if (kat !== "Modal Awal") continue;
      var nama = String(penData[i][2]).trim();
      var rowNum = i + 4;

      if (nama === "Petty Cash") {
        updates.push({r:rowNum, c:8, v:"✓ Sync Kas"});
      } else if (nama === "Saving Budget") {
        updates.push({r:rowNum, c:8, v:"✓ Sync Kas"});
      } else if (nama === "Galon + Isi") {
        asetRows.push(["Galon","Peralatan",tglMA,40000,12,0]);
        updates.push({r:rowNum, c:8, v:"✓ Sync Aset + Rekategori"});
      } else if (ASSET_MAP[nama]) {
        var a = ASSET_MAP[nama];
        asetRows.push([nama, a.k, tglMA, a.h, a.u, 0]);
        updates.push({r:rowNum, c:8, v:"✓ Sync Aset"});
      } else if (OPEX_NAMES[nama]) {
        updates.push({r:rowNum, c:2, v:"Operasional"});
        updates.push({r:rowNum, c:8, v:"✓ Rekategori OPEX"});
      } else {
        updates.push({r:rowNum, c:8, v:"⚠ Perlu Review"});
      }
    }

    if (asetRows.length === 0 && updates.length === 0) {
      notify("Tidak ada item Modal Awal yang perlu di-sync.");
      return;
    }

    // ── 4. Write Aset data (hapus example row + tulis batch) ──
    var lastAset = shAset.getLastRow();
    if (lastAset >= 5) shAset.getRange(5, 1, lastAset - 4, 9).clearContent();

    if (asetRows.length > 0) {
      shAset.getRange(5, 1, asetRows.length, 6).setValues(asetRows);
      for (var j = 0; j < asetRows.length; j++) {
        var r = j + 5;
        shAset.getRange(r, COLx(CA.HARGA)).setNumberFormat('"Rp "#,##0');
        shAset.getRange(r, COLx(CA.RESIDU)).setNumberFormat('"Rp "#,##0');
        shAset.getRange(r, COLx(CA.DEPRESIASI_BLN))
          .setFormula(F("=ROUND((D{row}-F{row})/E{row},0)", {row: r}))
          .setNumberFormat('"Rp "#,##0');
        shAset.getRange(r, COLx(CA.AKUMULASI)).setValue(0).setNumberFormat('"Rp "#,##0');
        shAset.getRange(r, COLx(CA.NILAI_BUKU))
          .setFormula(F("=D{row}-H{row}", {row: r}))
          .setNumberFormat('"Rp "#,##0');
        shAset.getRange(r, 1, 1, 9).setBackground("#FEF9E7").setFontSize(10);
        shAset.setRowHeight(r, 24);
      }
    }

    // ── 5. Write Kas entries (tulis langsung, tanpa lock nested) ──
    var saldoPC = _getSaldoKas("PC");
    var saldoPCBaru = saldoPC + 100000;
    var lrKas = shKas.getLastRow() + 1;
    shKas.getRange(lrKas, 1, 1, 6).setValues([
      [tglMA, "PC", "Saldo Awal", "Modal awal 02/05/2026", 100000, saldoPCBaru]
    ]);
    shKas.getRange(lrKas, 5, 1, 2).setNumberFormat('"Rp "#,##0');
    styleData(shKas.getRange(lrKas, 1, 1, 6), lrKas % 2 === 0 ? C.LIGHT : C.WHITE);
    shKas.setRowHeight(lrKas, 22);

    var saldoUB = _getSaldoKas("UB");
    var saldoUBBaru = saldoUB + 1000000;
    lrKas = shKas.getLastRow() + 1;
    shKas.getRange(lrKas, 1, 1, 6).setValues([
      [tglMA, "UB", "Saldo Awal", "Saving budget 02/05/2026", 1000000, saldoUBBaru]
    ]);
    shKas.getRange(lrKas, 5, 1, 2).setNumberFormat('"Rp "#,##0');
    styleData(shKas.getRange(lrKas, 1, 1, 6), lrKas % 2 === 0 ? C.LIGHT : C.WHITE);
    shKas.setRowHeight(lrKas, 22);

    _updateSaldoDisplay(shKas, "PC", _getSaldoKas("PC"));
    _updateSaldoDisplay(shKas, "UB", _getSaldoKas("UB"));

    // ── 6. Append "Air Isi Galon" ke Pengeluaran ──
    var newRowNum = lastPen + 1;
    shPen.getRange(newRowNum, 1, 1, 8).setValues([
      [tglMA, "Operasional", "Air Isi Galon", "Pcs", 1, 10000, 10000, "✎ Impor Sync"]
    ]);
    shPen.getRange(newRowNum, 1).setNumberFormat("DD/MM/YYYY");
    shPen.getRange(newRowNum, 6, 1, 2).setNumberFormat('"Rp "#,##0');

    // ── 7. Apply updates ke Pengeluaran ──
    for (var u = 0; u < updates.length; u++) {
      shPen.getRange(updates[u].r, updates[u].c).setValue(updates[u].v);
    }

    // ── 8. Selesai ──
    SpreadsheetApp.flush();
    refreshLaporan();
    auditLog("Sync Modal Awal", asetRows.length + " aset, PC 100rb, UB 1jt");
    notify("✅ Sync selesai! " + asetRows.length +
      " aset tetap masuk Aset, PC (100rb) & UB (1jt) masuk Kas.\n" +
      "▶ Jalankan menu 📉 Posting Penyusutan untuk posting depresiasi bulan Mei.");
  });
}

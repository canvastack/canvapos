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

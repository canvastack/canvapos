// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Builders.gs (Sheet Builders & Main Setup)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// MAIN SETUP — Jalankan setupPOS() sekali, atau step-by-step jika timeout
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Setup awal — buat semua sheet & isi data awal.
 * @param {string} [env] - Environment: "production", "staging", "development"
 */
function setupPOS(env) {
  if (!confirmAction("⚠️ PERINGATAN: Ini akan MENGHAPUS SEMUA data & membuat ulang semua sheet!\n\nLanjutkan?", "🔧 Setup Ulang Total?")) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone("Asia/Jakarta");

  // Set environment jika disediakan
  if (env) { setEnv(env); }

  // Buat sheet temp dulu agar tidak pernah 0 sheet
  var tempName = "__temp__";
  deleteSheetIfExists(ss, tempName);
  ss.insertSheet(tempName);

  ["Panduan","POS","Stock","Transaksi","Pendapatan","Pengeluaran","Kas","Bahan","Resep","Audit","Sheet1"].forEach(function(n) {
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
  buildAudit(ss);       SpreadsheetApp.flush();
  buildPendapatan(ss);  SpreadsheetApp.flush();
  buildPOS(ss);         SpreadsheetApp.flush();
  buildPanduan(ss);     SpreadsheetApp.flush();

  deleteSheetIfExists(ss, tempName);

  var order = ["Panduan","POS","Stock","Transaksi","Pendapatan","Pengeluaran","Kas","Bahan","Resep","Audit"];
  order.forEach(function(name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });

  ss.setActiveSheet(getSheet(SHEET.POS));
  clearDynamicCache();
  protectAll(); // P3.1 — auto-protect after setup
  setupNamedRanges(); // P3.4 — auto-create named ranges after setup
  auditLog("Setup POS", "Semua sheet dibuat kembali: Bahan, Resep, Stock, Transaksi, Pengeluaran, Kas, Audit, Pendapatan, POS, Panduan");
  try { SpreadsheetApp.getUi().alert("✅ Setup selesai!\n\nSemua sheet sudah dibuat.\nMulai transaksi dari sheet POS."); } catch(e) { /* non-UI context */ }
}

// Jalankan fungsi-fungsi ini satu per satu jika setupPOS() timeout:
/**
 * Setup sheet Bahan saja.
 */
function setup_1_Bahan()      { buildBahan(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet Resep saja.
 */
function setup_2_Resep()      { buildResep(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet Transaksi saja.
 */
function setup_3_Transaksi()  { buildTransaksi(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet Stock saja.
 */
function setup_4_Stock()      { buildStock(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet Pengeluaran saja.
 */
function setup_5_Pengeluaran() { buildPengeluaran(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet Kas saja.
 */
function setup_5b_Kas()        { buildKas(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet Pendapatan saja.
 */
function setup_5c_Pendapatan() { buildPendapatan(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet POS saja.
 */
function setup_6_POS()        { buildPOS(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Setup sheet Panduan saja.
 */
function setup_7_Panduan()    { buildPanduan(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); }
/**
 * Urutkan ulang sheet sesuai urutan default.
 */
function setup_8_Reorder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var order = ["Panduan","POS","Stock","Transaksi","Pendapatan","Pengeluaran","Kas","Bahan","Resep"];
  order.forEach(function(name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(getSheet(SHEET.POS));
  try { SpreadsheetApp.getUi().alert("✅ Setup selesai!"); } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — BAHAN
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet Bahan (master data).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function buildBahan(ss) {
  var C = getC();
  var sh = ss.insertSheet("Bahan");
  setSheetFont(sh);
  sh.setTabColor(C.BLUE);

  var headers = ["Kategori","Nama Bahan","Satuan","Ukuran/Pack","Harga Beli","Harga Per Piece"];
  var hRow = sh.getRange(1, 1, 1, headers.length);
  hRow.setValues([headers]);
  styleHeader(hRow, C.BLUE);
  sh.setRowHeight(1, 32);

  var data = getBahanData();

  // Tulis kolom A-E (tanpa Harga Per Piece)
  sh.getRange(2, 1, data.length, data[0].length).setValues(data);

  // Kolom F (Harga Per Piece) = =E/D
  data.forEach(function(_, i) {
    var r = i + 2;
    sh.getRange(r, 6).setFormula("=IFERROR(E"+r+"/D"+r+", 0)");
  });

  data.forEach(function(_, i) {
    var r = i + 2;
    applyZebraRow(sh, r, i, headers.length);
    setCurrency(sh.getRange(r, 5, 1, 2));
  });

  [100,200,100,80,120,140].forEach(function(w, i) { sh.setColumnWidth(i+1, w); });
  sh.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — RESEP (BILL OF MATERIALS / BOM)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet Resep (Bill of Materials).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
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

  // Data Master Resep (BOM) — generated from templates
  var data = generateBOMData();

  // Tulis seluruh baris data ke dalam sheet Resep (A-D)
  sh.getRange(2, 1, data.length, data[0].length).setValues(data);

  // Menerapkan gaya visual, formula harga, & format angka
  data.forEach(function(_, i) {
    var r = i + 2;
    applyZebraRow(sh, r, i, 6);
    
    sh.getRange(r, 3).setNumberFormat("#,##0.00");
    sh.getRange(r, 5).setFormula(F("=IFERROR(VLOOKUP(B{row}, BAHAN_Lookup, 5, FALSE), 0)", {row: r}));
    sh.getRange(r, 6).setFormula(F("=IFERROR(C{row}*E{row}, 0)", {row: r}));
    setCurrency(sh.getRange(r, 5, 1, 2));
  });

  // Set lebar kolom yang proporsional
  [220, 200, 140, 130, 140, 140].forEach(function(w, i) { sh.setColumnWidth(i+1, w); });
  sh.setFrozenRows(1);
}

/**
 * Sinkronisasi ulang data BOM dari ResepData.gs ke sheet Resep (kolom A-D).
 * Formula kolom E-F tetap aman. Panggil setelah update data BOM.
 */
function syncResepData() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.RESEP);
  if (!sh) return;
  var data = generateBOMData();
  if (!data || !data.length) return;
  sh.getRange(2, 1, data.length, data[0].length).setValues(data);
  // Hapus sisa baris lama (jika data baru lebih pendek)
  var lastRow = sh.getLastRow();
  if (lastRow > data.length + 1) {
    sh.getRange(data.length + 2, 1, lastRow - data.length - 1, 6).clearContent();
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — STOCK
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet Stock (inventory).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
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

    sh.getRange(r,6).setValue(stokAwal); // F = Sisa Stok (value, biar bisa di-write langsung)
    sh.getRange(r,7).setValue(minStok);
    sh.getRange(r,8).setFormula(F("=IF(F{row}<=G{row},\"⚠ RESTOCK\",\"✓ OK\")", {row: r}));

    applyZebraRow(sh, r, i, 8);
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
/**
 * Bangun sheet Transaksi (log transaksi).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
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

  [90,60,110,70,90,180,80,220,100,120,130,120].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(2);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — POS
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet POS (order entry).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
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
    .setFormula(F("=SUM(H{start}:H{end})", {start: POS_START_ROW, end: totalRow - 1}))
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
/**
 * Bangun sheet Panduan (petunjuk penggunaan).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function buildPanduan(ss) {
  var C = getC();
  var sh = ss.insertSheet("Panduan");
  setSheetFont(sh);
  sh.setTabColor(C.DARK);

  // ═══════════════════════════════════════════════════════════════════
  // DASHBOARD SECTION (rows 1-8)
  // ═══════════════════════════════════════════════════════════════════
  sh.getRange("A1:D1").merge()
    .setValue("📊 Dashboard CanvaPOS")
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);

  var labels = [
    ["System Health", "✅ Aktif", "", ""],
    ["Transaksi Terakhir", "", "", ""],
    ["Pendapatan Hari Ini", "", "", ""],
    ["Item Stok Menipis", "", "", ""],
    ["Saldo PC", "", "", ""],
    ["Saldo UB", "", "", ""],
    ["", "", "", ""],
  ];
  labels.forEach(function(row, i) {
    var r = i + 2;
    sh.getRange(r, 1).setValue(row[0])
      .setFontWeight("bold").setFontSize(10)
      .setBackground(C.LIGHT).setVerticalAlignment("middle");
    sh.getRange(r, 2, 1, 3).merge()
      .setValue(row[1] || "")
      .setFontSize(10).setFontColor(C.GREEN).setVerticalAlignment("middle");
    sh.setRowHeight(r, 22);
  });

  // Kas status row with real references
  sh.getRange(9, 1, 1, 4).merge()
    .setValue("📖 Panduan Penggunaan CanvaPOS")
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(9, 40);

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
    ["⚠️ INPUT BAHAN (Gram)","","",""],
    ["Untuk coffee (Kopi Robusta/Arabika)","Input jumlah gram — 1000 = 1Kg","",""],
    ["Contoh: beli 1Kg kopi","Masukkan 1000 di kolom Jumlah (Gram)","",""],
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
    var r = i + 10;
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
// BUILDER — Audit (Hidden Log)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Buat sheet Audit untuk log aktivitas.
 * @param {Spreadsheet} ss - Spreadsheet object
 */
function buildAudit(ss) {
  var C = getC();
  var sh = ss.insertSheet("Audit");
  setSheetFont(sh);
  sh.setTabColor("#95A5A6");
  sh.hideSheet();

  var headers = ["Timestamp", "User", "Action", "Detail"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(C.DARK).setFontColor(C.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 28);
  sh.setFrozenRows(1);

  sh.getRange("A:D").setVerticalAlignment("middle");
  sh.setColumnWidth(1, 180); // Timestamp
  sh.setColumnWidth(2, 120); // User
  sh.setColumnWidth(3, 200); // Action
  sh.setColumnWidth(4, 400); // Detail
}

/**
 * Catat aktivitas ke sheet Audit (hidden).
 * @param {string} action - Nama aksi (e.g. "Simpan Transaksi")
 * @param {string} detail - Detail aksi (e.g. "TRX-001, 3 item, Rp 25.000")
 */
function auditLog(action, detail) {
  if (isDebug()) Logger.log("📋 AUDIT [" + action + "] " + detail);
  var C = getC();
  withLock(5000, function() {
    var sh = getSheet(SHEET.AUDIT);
    if (!sh) return;
    var email = "unknown";
    try { email = Session.getActiveUser().getEmail(); } catch(e) {}
    var lastRow = sh.getLastRow() + 1;
    sh.getRange(lastRow, 1).setValue(fmtDate(new Date()) + " " + fmtTime(new Date()));
    sh.getRange(lastRow, 2).setValue(email);
    sh.getRange(lastRow, 3).setValue(action);
    sh.getRange(lastRow, 4).setValue(detail);
    styleData(sh.getRange(lastRow, 1, 1, 4), lastRow % 2 === 0 ? C.WHITE : C.LIGHT);
  });
}

/**
 * Hapus log audit yang lebih dari 90 hari.
 */
function cleanAuditLog() {
  withLock(10000, function() {
    var sh = getSheet(SHEET.AUDIT);
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 90);
    var deleted = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      var ts = new Date(data[i][0]);
      if (!isNaN(ts.getTime()) && ts < cutoff) {
        sh.deleteRow(i + 1);
        deleted++;
      }
    }
    if (deleted > 0) Logger.log("Audit cleanup: " + deleted + " baris dihapus.");
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// DASHBOARD — P3.2
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Refresh dashboard di sheet Panduan (rows 2-7).
 */
function refreshDashboard() {
  var sh = getSheet(SHEET.PANDUAN);
  if (!sh) return;
  var C = getC();

  // Row 2: System Health
  var allSheets = [SHEET.POS, SHEET.STOCK, SHEET.TRANSAKSI, SHEET.PENDAPATAN,
    SHEET.PENGELUARAN, SHEET.BAHAN, SHEET.RESEP, SHEET.KAS, SHEET.AUDIT];
  var missing = [];
  allSheets.forEach(function(name) {
    if (!getSheet(name)) missing.push(name);
  });
  var health = missing.length === 0 ? "✅ Semua sheet OK" : "⚠ Sheet hilang: " + missing.join(", ");
  sh.getRange("B2:D2").merge().setValue(health)
    .setFontColor(missing.length === 0 ? C.GREEN : C.RED).setFontSize(10);

  // Row 3: Transaksi Terakhir
  var shTrx = getSheet(SHEET.TRANSAKSI);
  if (shTrx) {
    var trxLast = shTrx.getLastRow();
    if (trxLast > 2) {
      var lastData = shTrx.getRange(trxLast, 1, 1, 12).getValues()[0];
      sh.getRange("B3:D3").merge().setValue(
        lastData[0] + " — " + (lastData[COL.TRANSAKSI.VARIAN] || "") + " — Rp " + ((lastData[COL.TRANSAKSI.TOTAL] || 0).toLocaleString("id-ID"))
      ).setFontSize(10).setFontColor(C.DARK);
    }
  }

  // Row 4: Pendapatan Hari Ini (from Pendapatan sheet)
  var shPen = getSheet(SHEET.PENDAPATAN);
  if (shPen) {
    var penVal = shPen.getRange("B8").getValue() || 0;
    sh.getRange("B4:D4").merge().setValue("Rp " + Number(penVal).toLocaleString("id-ID"))
      .setFontSize(10).setFontColor(C.DARK);
  }

  // Row 5: Item Stok Menipis
  var shStock = getSheet(SHEET.STOCK);
  if (shStock) {
    var stockLast = shStock.getLastRow();
    var lowStock = [];
    if (stockLast > 1) {
      var stockData = shStock.getRange(2, 1, stockLast - 1, 8).getValues();
      for (var i = 0; i < stockData.length; i++) {
        var sisa = Number(stockData[i][COL.STOCK.SISA]) || 0;
        var min = Number(stockData[i][COL.STOCK.MIN]) || 0;
        if (min > 0 && sisa <= min) {
          lowStock.push(stockData[i][COL.STOCK.NAMA]);
        }
      }
    }
    var stokMsg = lowStock.length === 0 ? "✅ Semua stok aman" : "⚠ " + lowStock.join(", ");
    sh.getRange("B5:D5").merge().setValue(stokMsg)
      .setFontColor(lowStock.length === 0 ? C.GREEN : C.RED).setFontSize(10);
  }

  // Row 6: Saldo PC
  var saldoPC = getSaldoPC();
  sh.getRange("B6:D6").merge().setValue("Rp " + (saldoPC || 0).toLocaleString("id-ID"))
    .setFontSize(10).setFontColor(saldoPC > 0 ? C.GREEN : C.RED);

  // Row 7: Saldo UB
  var saldoUB = getSaldoUB();
  sh.getRange("B7:D7").merge().setValue("Rp " + (saldoUB || 0).toLocaleString("id-ID"))
    .setFontSize(10).setFontColor(saldoUB >= 10000 ? C.GREEN : C.RED);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET PROTECTION — P3.1
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Proteksi semua formula cell di seluruh sheet.
 * @param {string} [adminEmail] - Email admin yang tetap bisa edit
 */
function protectAll(adminEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var protections = [];

  // Helper: add protection with consistent settings
  function _addProtection(range, desc) {
    try {
      var p = range.protect().setDescription(desc);
      p.removeEditors(p.getEditors());
      if (adminEmail) p.addEditor(adminEmail);
      protections.push(p);
    } catch(e) {
      Logger.log("Protect warning [" + desc + "]: " + e.message);
    }
  }

  // 1. POS — formula columns E-H (skip header & grand total)
  var shPOS = getSheet(SHEET.POS);
  if (shPOS) {
    var posLast = shPOS.getLastRow();
    if (posLast >= POS_START_ROW) {
      _addProtection(
        shPOS.getRange(POS_START_ROW, 5, posLast - POS_START_ROW + 1, 4),
        "POS — Formula (E-H)"
      );
    }
    _addProtection(shPOS.getRange("1:1"), "POS — Header");
    _addProtection(shPOS.getRange("2:6"), "POS — Info Bar");
  }

  // 2. Stock — header + formula (Status column)
  var shStock = getSheet(SHEET.STOCK);
  if (shStock) {
    _addProtection(shStock.getRange("1:1"), "Stock — Header");
    _addProtection(shStock.getRange("H:H"), "Stock — Formula Status");
  }

  // 3. Transaksi — all columns (read-only log)
  var shTrx = getSheet(SHEET.TRANSAKSI);
  if (shTrx) {
    _addProtection(shTrx.getRange("1:1"), "Transaksi — Header");
    _addProtection(shTrx.getDataRange(), "Transaksi — Semua data (read-only)");
  }

  // 4. Pendapatan — header + formula rows (4-11, 16+)
  var shPen = getSheet(SHEET.PENDAPATAN);
  if (shPen) {
    var penLast = shPen.getLastRow();
    if (penLast >= 4) {
      var totalRows = penLast - 4 + 1;
      _addProtection(shPen.getRange(4, 1, totalRows, 10), "Pendapatan — Rekap & Formula");
    }
    _addProtection(shPen.getRange("1:3"), "Pendapatan — Header & Kas Info");
  }

  // 5. Pengeluaran — header + status column
  var shPenEx = getSheet(SHEET.PENGELUARAN);
  if (shPenEx) {
    _addProtection(shPenEx.getRange("1:3"), "Pengeluaran — Header");
    _addProtection(shPenEx.getRange("H:H"), "Pengeluaran — Status Stok");
  }

  // 6. Bahan — header
  var shBahan = getSheet(SHEET.BAHAN);
  if (shBahan) {
    _addProtection(shBahan.getRange("1:1"), "Bahan — Header");
  }

  // 7. Resep — header + formula columns (E, F)
  var shResep = getSheet(SHEET.RESEP);
  if (shResep) {
    _addProtection(shResep.getRange("1:1"), "Resep — Header");
    var resepLast = shResep.getLastRow();
    if (resepLast > 1) {
      _addProtection(shResep.getRange(2, 5, resepLast - 1, 2), "Resep — Formula (E-F)");
    }
  }

  // 8. Kas — header + saldo display
  var shKas = getSheet(SHEET.KAS);
  if (shKas) {
    _addProtection(shKas.getRange("1:4"), "Kas — Header & Info");
  }

  // 9. Panduan — semua (read-only)
  var shPanduan = getSheet(SHEET.PANDUAN);
  if (shPanduan) {
    _addProtection(shPanduan.getDataRange(), "Panduan — Read-only");
  }

  // 10. Audit — semua (read-only)
  var shAudit = getSheet(SHEET.AUDIT);
  if (shAudit) {
    _addProtection(shAudit.getDataRange(), "Audit — Read-only");
  }

  var msg = "✅ Proteksi diterapkan di " + protections.length + " range.\n\n";
  if (adminEmail) msg += "Admin: " + adminEmail + " tetap bisa edit.";
  else msg += "Tidak ada admin — buka proteksi via unprotectAll().";
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

/**
 * Hapus semua proteksi di seluruh sheet.
 * @param {string} [adminEmail] - Email admin yang tetap punya akses setelah unprotect
 */
function unprotectAll(adminEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var removed = 0;

  sheets.forEach(function(sh) {
    var prot = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    prot.forEach(function(p) { p.remove(); removed++; });

    var sheetProt = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    sheetProt.forEach(function(p) { p.remove(); removed++; });
  });

  var msg = "✅ " + removed + " proteksi dihapus.";
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
  auditLog("Unprotect All", removed + " proteksi dihapus");
}

// ═══════════════════════════════════════════════════════════════════════════
// BACKUP — P3.6
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Backup spreadsheet dengan timestamp nama.
 * @return {Spreadsheet} File backup yang baru dibuat
 */
function backupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = fmtDate(new Date()) + " " + fmtTime(new Date()).replace(/:/g, ".");
  var name = ss.getName() + " — Backup " + now;
  var backup;
  try {
    backup = ss.copy(name);
  } catch(e) {
    try { SpreadsheetApp.getUi().alert("❌ Gagal backup: " + e.message); } catch(e2) {}
    return null;
  }
  auditLog("Backup Spreadsheet", "File: " + name);
  try {
    SpreadsheetApp.getUi().alert("✅ Backup berhasil!\n\nNama file:\n" + name);
  } catch(e) {}
  return backup;
}

/**
 * Hapus backup yang lebih dari 7 hari.
 * @param {number} [maxAgeDays=7] - Maksimal umur backup (hari)
 */
function cleanBackups(maxAgeDays) {
  maxAgeDays = maxAgeDays || 7;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents();
  if (!parentFolder.hasNext()) return;

  var folder = parentFolder.next();
  var namePrefix = ss.getName() + " — Backup ";
  var files = folder.getFiles();
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - maxAgeDays);
  var deleted = 0;

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(namePrefix) === 0) {
      if (file.getDateCreated() < cutoff) {
        DriveApp.getFileById(file.getId()).setTrashed(true);
        deleted++;
      }
    }
  }

  if (deleted > 0) {
    Logger.log("Backup cleanup: " + deleted + " file dihapus.");
    auditLog("Clean Backups", deleted + " backup > " + maxAgeDays + " hari dihapus");
  }
}

/**
 * Setup daily backup trigger.
 * @param {number} [hour=2] - Jam (0-23) untuk eksekusi backup
 */
function setupBackupTrigger(hour) {
  hour = (hour !== undefined) ? hour : 2;

  // Hapus trigger lama
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "backupSpreadsheet") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Buat trigger baru
  ScriptApp.newTrigger("backupSpreadsheet")
    .timeBased().everyDays(1).atHour(hour).create();

  // Juga atur cleanup trigger
  triggers = ScriptApp.getProjectTriggers();
  var hasClean = false;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "cleanBackups") hasClean = true;
  });
  if (!hasClean) {
    ScriptApp.newTrigger("cleanBackups")
      .timeBased().everyDays(1).atHour(hour + 1).create();
  }

  var msg = "✅ Backup otomatis diatur setiap hari jam " + hour + ":00.\n" +
    "Backup lama (> 7 hari) akan otomatis dihapus.";
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
  auditLog("Setup Backup Trigger", "Daily at " + hour + ":00");
}

/**
 * Tampilkan daftar semua file backup yang bisa dibuka.
 * Setiap baris menampilkan nama file, tanggal, dan link untuk membuka.
 */
function listBackups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents();
  if (!parentFolder.hasNext()) {
    SpreadsheetApp.getUi().alert("❌ Folder tidak ditemukan.");
    return;
  }

  var folder = parentFolder.next();
  var namePrefix = ss.getName() + " — Backup ";
  var files = folder.getFiles();
  var html = '<style>' +
    'body{font-family:monospace;padding:16px;color:#333}' +
    'h3{margin:0 0 12px;color:#2C3E50}' +
    'table{width:100%;border-collapse:collapse;font-size:13px}' +
    'th{text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;color:#666}' +
    'td{padding:6px 8px;border-bottom:1px solid #eee}' +
    'a{color:#2980b9;text-decoration:none}' +
    'a:hover{text-decoration:underline}' +
    '.empty{padding:20px;text-align:center;color:#999}' +
    '.btn{margin-top:12px;padding:8px 16px;background:#667eea;color:#fff;border:none;border-radius:4px;cursor:pointer}' +
    '</style>' +
    '<h3>📂 Daftar Backup</h3>';

  var count = 0;
  html += '<table><tr><th>No</th><th>Nama File</th><th>Dibuat</th></tr>';
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(namePrefix) === 0) {
      count++;
      var url = file.getUrl();
      var created = Utilities.formatDate(file.getDateCreated(), "Asia/Jakarta", "dd/MM/yyyy HH:mm");
      html += '<tr>' +
        '<td>' + count + '</td>' +
        '<td><a href="' + url + '" target="_blank">' + file.getName().replace(namePrefix, '') + '</a></td>' +
        '<td>' + created + '</td>' +
        '</tr>';
    }
  }
  html += '</table>';

  if (count === 0) {
    html += '<div class="empty">Belum ada backup. Klik menu <b>📀 Backup Sekarang</b> untuk membuat.</div>';
  } else {
    html += '<p style="font-size:12px;color:#666;margin-top:8px">Total: ' + count + ' backup</p>';
  }

  html += '<button class="btn" onclick="google.script.host.close()">Tutup</button>';

  var output = HtmlService.createHtmlOutput(html).setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, "📂 Daftar Backup CanvaPOS");
}

/**
 * Restore dari backup: buka file backup di tab baru.
 */
function restoreBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents();
  if (!parentFolder.hasNext()) {
    SpreadsheetApp.getUi().alert("❌ Folder tidak ditemukan.");
    return;
  }

  var folder = parentFolder.next();
  var namePrefix = ss.getName() + " — Backup ";
  var files = folder.getFiles();
  var backups = [];
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(namePrefix) === 0) {
      backups.push(file);
    }
  }

  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert("❌ Belum ada backup.");
    return;
  }

  // Urutkan dari terbaru
  backups.sort(function(a, b) { return b.getDateCreated() - a.getDateCreated(); });

  // Buka backup terbaru
  var latest = backups[0];
  var url = latest.getUrl();
  var html = '<style>' +
    'body{font-family:sans-serif;padding:20px;text-align:center;color:#333}' +
    'h3{color:#2C3E50}' +
    'p{font-size:14px;line-height:1.6}' +
    '.btn{padding:10px 24px;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:600;margin:4px}' +
    '.btn-primary{background:#2980b9;color:#fff}' +
    '.btn-secondary{background:#95a5a6;color:#fff}' +
    '</style>' +
    '<h3>📀 Restore dari Backup</h3>' +
    '<p>Backup terbaru:<br><b>' + latest.getName() + '</b></p>' +
    '<p style="font-size:13px;color:#666">Dibuat: ' +
    Utilities.formatDate(latest.getDateCreated(), "Asia/Jakarta", "dd/MM/yyyy HH:mm") + '</p>' +
    '<p><button class="btn btn-primary" onclick="window.open(\'' + url + '\',\'_blank\')">📂 Buka Backup</button></p>' +
    '<p style="font-size:12px;color:#999">Atau buka menu <b>📂 Lihat Backup</b> untuk memilih backup lain.</p>' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Tutup</button>';

  var output = HtmlService.createHtmlOutput(html).setWidth(450).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(output, "📀 Restore CanvaPOS");
}

// ═══════════════════════════════════════════════════════════════════════════
// NAMED RANGES — P3.4
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Daftar named range definition.
 * @return {Object[]} Array { name, sheetName, col, startRow }
 */
function _getNamedRangeDefs() {
  return [
    { name: "TRX_Tgl",    sheet: SHEET.TRANSAKSI, col: COL.TRANSAKSI.TGL,     startRow: 3 },
    { name: "TRX_Cup",    sheet: SHEET.TRANSAKSI, col: COL.TRANSAKSI.CUP,     startRow: 3 },
    { name: "TRX_Total",  sheet: SHEET.TRANSAKSI, col: COL.TRANSAKSI.TOTAL,   startRow: 3 },
    { name: "TRX_JmlTop", sheet: SHEET.TRANSAKSI, col: COL.TRANSAKSI.JML_TOP, startRow: 3 },
    { name: "BAHAN_Lookup",sheet: SHEET.BAHAN,     col: COL.BAHAN.NAMA,       startRow: 2, width: 5 }
  ];
}

/**
 * Setup semua named range dengan row count generous (10000).
 * Panggil setelah setupPOS().
 */
function setupNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = _getNamedRangeDefs();

  defs.forEach(function(d) {
    var sh = ss.getSheetByName(d.sheet);
    if (!sh) return;
    try {
      var w = d.width || 1;
      var range = sh.getRange(d.startRow, COLx(d.col), 10000, w);
      ss.setNamedRange(d.name, range);
    } catch(e) {
      Logger.log("⚠ Gagal setup named range " + d.name + ": " + e.message);
    }
  });

  Logger.log("✅ " + defs.length + " named range dibuat.");
}

/**
 * Refresh semua named range ke ukuran data aktual.
 * Panggil setelah simpanTransaksi / simpanPengeluaran.
 */
function refreshNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = _getNamedRangeDefs();

  defs.forEach(function(d) {
    var sh = ss.getSheetByName(d.sheet);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    var numRows = Math.max(1, lastRow - d.startRow + 1);
    numRows = Math.min(numRows, 20000); // safety cap
    try {
      var w = d.width || 1;
      var range = sh.getRange(d.startRow, COLx(d.col), numRows, w);
      ss.setNamedRange(d.name, range);
    } catch(e) {
      Logger.log("⚠ Gagal refresh named range " + d.name + ": " + e.message);
    }
  });
}

/**
 * Hapus semua named range dari project.
 */
function clearNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = ss.getNamedRanges();
  namedRanges.forEach(function(nr) { nr.remove(); });
}

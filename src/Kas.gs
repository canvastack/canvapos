// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Kas.gs (Petty Cash & Uang Belanja)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// SHEET — KAS (Petty Cash & Uang Belanja tracking)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Bangun sheet Kas (Petty Cash & Uang Belanja).
 * @param {Spreadsheet} ss - Spreadsheet object
 */
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
  var sh = getSheet(SHEET.KAS);
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

/**
 * Dapatkan saldo Petty Cash (PC) saat ini.
 * @return {number} Saldo PC
 */
function getSaldoPC() { return _getSaldoKas("PC"); }
/**
 * Dapatkan saldo Uang Belanja (UB) saat ini.
 * @return {number} Saldo UB
 */
function getSaldoUB() { return _getSaldoKas("UB"); }

// ── Helper: catat transaksi ke Kas ────────────────────────────────────
function _catatKas(kategori, jenis, keterangan, jumlah, tgl) {
  withLock(10000, function() {
    var sh = getSheet(SHEET.KAS);
    if (!sh) return;
    var C = getC();
    var today = tgl ? fmtDate(tgl) : fmtDate(new Date());
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
  });
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
/**
 * Top up PC ke Rp 100.000 (closing harian).
 */
function topUpPC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheet(SHEET.KAS);
  if (!sh) { SpreadsheetApp.getUi().alert("Sheet Kas belum ada. Jalankan Setup POS dulu."); return; }
  var saldo = getSaldoPC();
  var kekurangan = 100000 - saldo;
  if (kekurangan <= 0) {
    SpreadsheetApp.getUi().alert("✅ Saldo PC sudah Rp 100.000 atau lebih. Tidak perlu top up.");
    return;
  }
  _catatKas("PC", "Top Up", "Top up PC — tutup kas harian", kekurangan);
  auditLog("Top Up PC", "Rp " + kekurangan.toLocaleString("id-ID") + " — saldo baru Rp 100.000");
  SpreadsheetApp.getUi().alert("✅ PC di-top up Rp " + kekurangan.toLocaleString("id-ID") + "\nSaldo PC sekarang: Rp 100.000");
}

// ── Macro: Top Up UB ke Rp 100.000 ─────────────────────────────────────
/**
 * Top up UB jika saldo di bawah Rp 10.000.
 */
function topUpUB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheet(SHEET.KAS);
  if (!sh) { SpreadsheetApp.getUi().alert("Sheet Kas belum ada. Jalankan Setup POS dulu."); return; }
  var saldo = getSaldoUB();
  if (saldo >= 10000) {
    SpreadsheetApp.getUi().alert("ℹ Saldo UB masih Rp " + saldo.toLocaleString("id-ID") + " (> Rp 10.000).\nTop up hanya diperlukan jika saldo < Rp 10.000.");
    return;
  }
  var kekurangan = 100000 - saldo;
  _catatKas("UB", "Top Up", "Top up UB — saldo menipis", kekurangan);
  auditLog("Top Up UB", "Rp " + kekurangan.toLocaleString("id-ID") + " — saldo baru Rp 100.000");
  SpreadsheetApp.getUi().alert("✅ UB di-top up Rp " + kekurangan.toLocaleString("id-ID") + "\nSaldo UB sekarang: Rp 100.000");
}

// ── Macro: Set Saldo Awal PC & UB (panggil setiap hari baru) ───────────
/**
 * Catat saldo awal PC & UB untuk hari ini.
 */
function initSaldoKas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheet(SHEET.KAS);
  if (!sh) { SpreadsheetApp.getUi().alert("Sheet Kas belum ada. Jalankan Setup POS dulu."); return; }
  
  // Cek apakah hari ini sudah ada Saldo Awal
  var today = fmtDate(new Date());
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
    notify("Saldo awal hari " + today + " sudah dicatat. Lewati.");
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
  auditLog("Init Saldo Kas", msgParts.join(", "));
  notify("✅ Saldo awal dicatat: " + msgParts.join(", "));
}

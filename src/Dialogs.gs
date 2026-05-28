// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — Dialogs.gs (HTML Dialogs & Handlers)
// ═══════════════════════════════════════════════════════════════════════════

// ── Generic Multi-Select Checklist Dialog ─────────────────────────────────
/**
 * Tampilkan dialog checklist multi-select.
 * @param {string} title - Judul dialog
 * @param {string[]} items - Daftar item
 * @param {string[]} selected - Item yang sudah dipilih
 * @param {string} callback - Nama fungsi callback
 */
function showMultiSelectDialog(title, items, selected, callback) {
  if (!items || items.length === 0) {
    SpreadsheetApp.getUi().alert("Tidak ada item untuk dipilih.");
    return;
  }
  var itemOpts = items.map(function(item) {
    var checked = selected.indexOf(item) >= 0 ? "checked" : "";
    return '<label style="display:block;padding:6px 8px;border-bottom:1px solid #eee;cursor:pointer">' +
      '<input type="checkbox" ' + checked + ' value="' + item.replace(/"/g, "&quot;") + '"> ' + item +
      '</label>';
  }).join("");

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Nunito,sans-serif;padding:16px;background:#f9f9f9}' +
    'h3{color:#2C3E50;margin:0 0 12px;font-size:15px;text-align:center}' +
    '.items{max-height:300px;overflow-y:auto;border:1px solid #ddd;border-radius:6px;background:#fff}' +
    '.btn-bar{display:flex;gap:8px;margin-top:12px}' +
    '.btn-ok{flex:1;background:#27AE60;color:#fff;border:none;border-radius:6px;padding:10px;font-size:14px;font-weight:bold;cursor:pointer}' +
    '.btn-cancel{flex:1;background:#ccc;color:#333;border:none;border-radius:6px;padding:10px;font-size:14px;cursor:pointer}' +
    '.select-all{font-size:12px;color:#2E86AB;cursor:pointer;text-align:right;padding:4px 8px}' +
    '</style>' +
    '<h3>' + title + '</h3>' +
    '<div class="select-all" onclick="toggleAll()">☑ Pilih Semua / Hapus Semua</div>' +
    '<div class="items" id="items">' + itemOpts + '</div>' +
    '<div class="btn-bar">' +
    '<button class="btn-ok" onclick="ok()">✓ OK</button>' +
    '<button class="btn-cancel" onclick="cancel()">Batal</button>' +
    '</div>' +
    '<script>' +
    'function ok(){var c=document.querySelectorAll("input:checked");var vals=[];c.forEach(function(x){vals.push(x.value)});' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})["' + callback + '"](vals.join(","));}' +
    'function cancel(){google.script.host.close()}' +
    'function toggleAll(){var c=document.querySelectorAll("input[type=checkbox]");var allChecked=true;c.forEach(function(x){if(!x.checked)allChecked=false});' +
    'c.forEach(function(x){x.checked=!allChecked})}' +
    '</script>'
  ).setWidth(350).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

// ── Handler: callback untuk showMultiSelectDialog ke pilihTopping ────────
function _pilihToppingCallback(selectedStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheet(SHEET.POS);
  var activeRow = ss.getActiveRange().getRow();
  if (!sh || activeRow < POS_START_ROW) return;
  sh.getRange(activeRow, COLx(COL.POS.TOPPING)).setValue(selectedStr);
}

// ── Generic Date Picker ───────────────────────────────────────────────────
/**
 * Tampilkan date picker untuk cell aktif.
 */
function showDatePickerGeneric() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeCell = ss.getActiveCell();
  var row = activeCell.getRow();
  var col = activeCell.getColumn();
  var shName = ss.getActiveSheet().getName();

  if (row < 1 || col < 1) {
    SpreadsheetApp.getUi().alert("Pilih cell yang ingin diisi tanggal.");
    return;
  }
  showDatePickerForRowCol(row, col, shName);
}

/**
 * Tampilkan date picker untuk (row, col) tertentu.
 * @param {number} row - Baris target
 * @param {number} col - Kolom target
 * @param {string} sheetName - Nama sheet
 */
function showDatePickerForRowCol(row, col, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Nunito,sans-serif;text-align:center;padding:20px}' +
    'h3{margin:0 0 16px;color:#333}' +
    'input[type=date]{font-size:18px;padding:10px;border:2px solid #E67E22;' +
    'border-radius:8px;width:100%;box-sizing:border-box}' +
    '.btn{margin-top:16px;background:#E67E22;color:#fff;border:none;' +
    'padding:10px 24px;border-radius:6px;font-size:16px;cursor:pointer}' +
    '.btn-cancel{margin-top:16px;background:#ccc;color:#333;border:none;' +
    'padding:10px 24px;border-radius:6px;font-size:16px;cursor:pointer;margin-left:8px}' +
    '.btn:hover{background:#d35400}' +
    '</style>' +
    '<h3>📅 Pilih Tanggal</h3>' +
    '<input type="date" id="picker" autofocus>' +
    '<br><button class="btn" onclick="simpan()">✓ Simpan</button>' +
    '<button class="btn-cancel" onclick="cancel()">Batal</button>' +
    '<script>' +
    'function simpan(){var v=document.getElementById("picker").value;if(!v){alert("Pilih tanggal dulu!");return;}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.setDateValueGeneric(' + row + ',' + col + ',"" + v);}' +
    'function cancel(){google.script.host.close()}' +
    '</script>'
  ).setWidth(300).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, "📅 Pilih Tanggal");
}

// ── Generic Date Setter — writes to any (row, col) on active sheet ───────
/**
 * Tulis nilai tanggal ke cell (row, col) dari date picker.
 * @param {number} row - Baris target
 * @param {number} col - Kolom target
 * @param {string} dateStr - String tanggal YYYY-MM-DD
 */
function setDateValueGeneric(row, col, dateStr) {
  var parts = dateStr.split("-");
  if (parts.length !== 3) return;
  var d = new Date(parts[0], parts[1] - 1, parts[2]);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sh.getRange(row, col).setValue(d).setNumberFormat("DD/MM/YYYY");
}

// ═══════════════════════════════════════════════════════════════════════════
// MACRO — TAMBAH RESEP (form popup dialog)
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Tampilkan dialog tambah Resep / BOM.
 */
function showTambahResepDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shResep = getSheet(SHEET.RESEP);
  var shBahan = getSheet(SHEET.BAHAN);
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
/**
 * Simpan data resep baru dari dialog ke sheet Resep.
 * @param {string} namaMenu - Nama menu
 * @param {string} jsonData - JSON string { bahan, qty, satuan }
 */
function simpanTambahResep(namaMenu, jsonData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheet(SHEET.RESEP);
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

// ── Legacy: showDatePicker → showDatePickerGeneric (deprecated, keep for menu) ─
/**
 * Legacy date picker — alias ke showDatePickerGeneric.
 */
function showDatePicker() { showDatePickerGeneric(); }

// ═══════════════════════════════════════════════════════════════════════════
// ONBOARDING WIZARD — P3.5
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Tampilkan onboarding wizard 4 langkah.
 */
function showOnboardingWizard() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Nunito,sans-serif;padding:20px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#333;margin:0}' +
    '.card{background:#fff;border-radius:12px;padding:24px;box-shadow:0 8px 32px rgba(0,0,0,0.15);max-width:500px;margin:auto}' +
    'h2{text-align:center;color:#2C3E50;margin:0 0 16px;font-size:18px}' +
    '.steps{display:flex;justify-content:center;gap:4px;margin-bottom:20px}' +
    '.step{width:32px;height:4px;border-radius:2px;background:#ddd;transition:background .3s}' +
    '.step.active{background:#667eea}' +
    '.step.done{background:#27ae60}' +
    '.content{min-height:180px}' +
    '.content h3{color:#2C3E50;margin:0 0 8px;font-size:15px}' +
    '.content p{color:#666;margin:0 0 16px;font-size:13px;line-height:1.5}' +
    '.btn-bar{display:flex;justify-content:space-between;gap:8px;margin-top:16px}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;transition:opacity .2s}' +
    '.btn:hover{opacity:.85}' +
    '.btn-primary{background:#667eea;color:#fff}' +
    '.btn-success{background:#27ae60;color:#fff}' +
    '.btn-outline{background:transparent;color:#667eea;border:1px solid #667eea}' +
    '.btn:disabled{opacity:.5;cursor:default}' +
    '.status{font-size:12px;padding:8px 12px;border-radius:6px;margin:8px 0}' +
    '.status.ok{background:#d4edda;color:#155724}' +
    '.status.warn{background:#fff3cd;color:#856404}' +
    '.status.err{background:#f8d7da;color:#721c24}' +
    '</style>' +
    '<div class="card">' +
    '<h2>🚀 Selamat Datang di CanvaPOS</h2>' +
    '<div class="steps" id="stepBar">' +
    '<div class="step active" id="s0"></div>' +
    '<div class="step" id="s1"></div>' +
    '<div class="step" id="s2"></div>' +
    '<div class="step" id="s3"></div>' +
    '</div>' +
    '<div class="content" id="stepContent">' +
    '<h3>🔧 Setup System</h3>' +
    '<p>Langkah 1 dari 4: Memastikan semua sheet sudah siap.</p>' +
    '<div id="statusCheck"></div>' +
    '</div>' +
    '<div class="btn-bar">' +
    '<div></div>' +
    '<button class="btn btn-primary" id="nextBtn" onclick="nextStep()">Lanjut →</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'var step = 0;' +
    'var steps = [' +
    '  { title: "🔧 Setup System", text: "Memastikan semua sheet sudah siap.", check: "checkSetup" },' +
    '  { title: "📦 Input Stok Awal", text: "Isi stok awal di sheet Stock.", check: "checkStock" },' +
    '  { title: "💰 Initialize Kas", text: "Catat saldo awal PC & UB.", check: "checkKas" },' +
    '  { title: "🎉 Siap Transaksi", text: "Panduan memulai transaksi pertama.", check: "checkReady" }' +
    '];' +
    '' +
    'google.script.run.withSuccessHandler(function(r) {' +
    '  document.getElementById("statusCheck").innerHTML = r;' +
    '}).checkSetup();' +
    '' +
    'function nextStep() {' +
    '  if (step < 3) {' +
    '    step++;' +
    '    updateStep();' +
    '  } else {' +
    '    google.script.host.close();' +
    '  }' +
    '}' +
    '' +
    'function prevStep() {' +
    '  if (step > 0) { step--; updateStep(); }' +
    '}' +
    '' +
    'function runSetup() {' +
    '  google.script.run.withSuccessHandler(function(){ updateStep(); }).setupPOS();' +
    '}' +
    '' +
    'function initKas() {' +
    '  google.script.run.withSuccessHandler(function(){ updateStep(); }).initSaldoKas();' +
    '}' +
    '' +
    'function updateStep() {' +
    '  for (var i = 0; i < 4; i++) {' +
    '    var el = document.getElementById("s" + i);' +
    '    el.className = "step";' +
    '    if (i < step) el.className = "step done";' +
    '    if (i === step) el.className = "step active";' +
    '  }' +
    '  var s = steps[step];' +
    '  document.getElementById("stepContent").innerHTML = ' +
    '    "<h3>" + s.title + "</h3><p>" + s.text + "</p><div id=\"statusCheck\"></div>";' +
    '  var btn = document.getElementById("nextBtn");' +
    '  if (step === 3) { btn.textContent = "✅ Selesai"; btn.className = "btn btn-success"; }' +
    '  else { btn.textContent = "Lanjut →"; btn.className = "btn btn-primary"; }' +
    '  document.getElementById("stepContent").innerHTML += ' +
    '    "<div class=\"btn-bar\" style=\"margin-top:8px\">" +' +
    '    (step > 0 ? "<button class=\"btn btn-outline\" onclick=\"prevStep()\">← Kembali</button>" : "<div></div>") +' +
    '    "</div>";' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    document.getElementById("statusCheck").innerHTML = r;' +
    '  })[s.check]();' +
    '}' +
    '</script>'
  ).setWidth(480).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, "🚀 Onboarding Wizard");
}

/**
 * Step 1: Cek apakah setup sudah jalan.
 * @return {string} HTML status
 */
function checkSetup() {
  var sheets = ["POS","Stock","Transaksi","Pendapatan","Pengeluaran","Bahan","Resep","Kas"];
  var missing = [];
  sheets.forEach(function(name) {
    if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name)) missing.push(name);
  });
  if (missing.length === 0) {
    return '<div class="status ok">✅ Semua sheet sudah siap. Setup OK!</div>' +
      '<div class="status ok">✔ Total: ' + sheets.length + ' sheet terdeteksi.</div>';
  }
  return '<div class="status err">⚠ Sheet belum lengkap. Klik "Setup Ulang" di menu POS.</div>' +
    '<div class="status warn">Hilang: ' + missing.join(", ") + '</div>' +
    '<button class="btn btn-primary" onclick="runSetup()">🔧 Setup Sekarang</button>';
}

/**
 * Step 2: Cek stok.
 * @return {string} HTML status
 */
function checkStock() {
  var sh = getSheet(SHEET.STOCK);
  if (!sh) return '<div class="status err">⚠ Sheet Stock tidak ditemukan.</div>';
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return '<div class="status warn">⚠ Belum ada data stok. Buka sheet Stock dan isi stok awal.</div>';
  var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  var filled = 0;
  for (var i = 0; i < data.length; i++) {
    if (Number(data[i][3]) > 0 || Number(data[i][5]) > 0) filled++;
  }
  var pct = Math.round(filled / data.length * 100);
  var color = pct === 100 ? "ok" : (pct >= 50 ? "warn" : "err");
  return '<div class="status ' + color + '">📦 ' + filled + '/' + data.length + ' item terisi (' + pct + '%)</div>' +
    '<p style="font-size:12px;color:#666">Buka sheet <b>Stock</b> → isi kolom Stok Awal.<br>' +
    'Atau gunakan <b>Pengeluaran</b> → Simpan & Sync Stok untuk isi otomatis.</p>';
}

/**
 * Step 3: Cek Kas.
 * @return {string} HTML status
 */
function checkKas() {
  var saldoPC = getSaldoPC();
  var saldoUB = getSaldoUB();
  var html = '';
  html += (saldoPC > 0)
    ? '<div class="status ok">💰 PC: Rp ' + saldoPC.toLocaleString("id-ID") + '</div>'
    : '<div class="status warn">💰 PC: Belum di-set. Klik "Init Saldo Awal PC & UB"</div>';
  html += (saldoUB > 0)
    ? '<div class="status ok">🛒 UB: Rp ' + saldoUB.toLocaleString("id-ID") + '</div>'
    : '<div class="status warn">🛒 UB: Belum di-set. Klik "Init Saldo Awal PC & UB"</div>';
  if (saldoPC <= 0 || saldoUB <= 0) {
    html += '<button class="btn btn-primary" onclick="initKas()">💰 Init Saldo Sekarang</button>';
  }
  return html;
}

/**
 * Step 4: Siap transaksi.
 * @return {string} HTML status
 */
function checkReady() {
  return '<div class="status ok">🎉 CanvaPOS siap digunakan!</div>' +
    '<p style="font-size:13px;color:#666">Langkah selanjutnya:</p>' +
    '<ol style="font-size:13px;color:#444;line-height:1.8;padding-left:20px">' +
    '<li>Buka sheet <b>POS</b></li>' +
    '<li>Pilih varian produk dari dropdown</li>' +
    '<li>Isi jumlah cup</li>' +
    '<li>Klik menu POS → <b>Simpan Transaksi</b></li>' +
    '</ol>';
}

// ═══════════════════════════════════════════════════════════════════════════
// ENVIRONMENT PICKER — P3.7
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Tampilkan dialog untuk memilih environment.
 */
function showEnvPicker() {
  var env = getEnv();
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Nunito,sans-serif;padding:20px;color:#333}' +
    'h3{color:#2C3E50;margin:0 0 16px;font-size:16px;text-align:center}' +
    'label{display:flex;align-items:center;gap:8px;padding:10px 12px;margin:4px 0;border-radius:6px;cursor:pointer}' +
    'label:hover{background:#f0f0f0}' +
    'label.selected{background:#d4edda;border:1px solid #27ae60}' +
    'input[type=radio]{margin:0}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600}' +
    '.btn-primary{background:#667eea;color:#fff;width:100%;margin-top:12px}' +
    '.btn:hover{opacity:.85}' +
    '.desc{font-size:11px;color:#888;margin-left:24px}' +
    '</style>' +
    '<h3>🌐 Pilih Environment</h3>' +
    '<div id="cur" style="text-align:center;font-size:12px;color:#888;margin-bottom:12px">Current: <b>' + env + '</b></div>' +
    '<label class="' + (env === "production" ? 'selected' : '') + '">' +
    '<input type="radio" name="env" value="production" ' + (env === "production" ? 'checked' : '') + '> ' +
    '<b>⚡ Production</b></label>' +
    '<div class="desc">Konfirmasi ON · Debug OFF · Data asli</div>' +
    '<label class="' + (env === "staging" ? 'selected' : '') + '">' +
    '<input type="radio" name="env" value="staging" ' + (env === "staging" ? 'checked' : '') + '> ' +
    '<b>🧪 Staging</b></label>' +
    '<div class="desc">Konfirmasi ON · Debug ON · Uji coba</div>' +
    '<label class="' + (env === "development" ? 'selected' : '') + '">' +
    '<input type="radio" name="env" value="development" ' + (env === "development" ? 'checked' : '') + '> ' +
    '<b>🔧 Development</b></label>' +
    '<div class="desc">Konfirmasi SKIP · Debug ON · Tanpa hambatan</div>' +
    '<button class="btn btn-primary" onclick="applyEnv()">✅ Terapkan</button>' +
    '<script>' +
    'function applyEnv(){' +
    '  var v = document.querySelector("input[name=env]:checked").value;' +
    '  google.script.run.withSuccessHandler(function(){google.script.host.close();}).setupEnv(v);' +
    '}' +
    '</script>'
  ).setWidth(360).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "🌐 Pilih Environment");
}

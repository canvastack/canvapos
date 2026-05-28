// ═══════════════════════════════════════════════════════════════════════════
// CanvaPOS — ResepData.gs (BOM Data Generator — Template-Based)
// ═══════════════════════════════════════════════════════════════════════════

function generateBOMData() {
  var data = [];

  // ── POP ICE BLENDER — 18 variants ──────────────────────────────────────
  var VARIAN_POP_ICE = [
    "Chociato","Cokelat","Duren","Strawberry","Alpukat","Taro",
    "Blueberry","Permen Karet","Lychee","Anggur","Mangga","Melon",
    "Cappuccino","Moccacino","Vanilla Latte","Cookies & Krim",
    "Es Doger","Es Teler"
  ];
  var POP_ICE_TEMPLATE = [
    ["Gula Pasir",20,"Gram"],["Susu SKM",30,"Gram"],
    ["Air (Galon)",0.15,"Liter"],["Es Batu Kristal",0.2,"Kg"],
    ["Cup Plastik 18oz",1,"Piece"],["Tutup Cup Plastik",1,"Piece"],
    ["Sedotan Biasa",1,"Piece"]
  ];

  VARIAN_POP_ICE.forEach(function(v) {
    var m = "Pop Ice - " + v;
    data.push([m, m, 1, "Piece"]);
    POP_ICE_TEMPLATE.forEach(function(t) { data.push([m, t[0], t[1], t[2]]); });
  });

  // ── KOPI TUBRUK & ES TEH ───────────────────────────────────────────────
  function addBOM(menuName, items) {
    items.forEach(function(t) { data.push([menuName, t[0], t[1], t[2]]); });
  }
  addBOM("Kopi Tubruk Robusta", [
    ["Biji Kopi Robusta",15,"Gram"],["Gula Pasir",15,"Gram"],
    ["Air (Galon)",0.2,"Liter"],["Paper Cup 8oz",1,"Piece"],
    ["Tutup Paper Cup",1,"Piece"],["Sedotan Biasa",1,"Piece"]
  ]);
  addBOM("Kopi Tubruk Arabika", [
    ["Biji Kopi Arabika",15,"Gram"],["Gula Pasir",15,"Gram"],
    ["Air (Galon)",0.2,"Liter"],["Paper Cup 8oz",1,"Piece"],
    ["Tutup Paper Cup",1,"Piece"],["Sedotan Biasa",1,"Piece"]
  ]);
  addBOM("Es Teh Original", [
    ["Teh Celup",1,"Piece"],["Gula Pasir",20,"Gram"],
    ["Es Batu Kristal",0.2,"Kg"],["Cup Plastik 18oz",1,"Piece"],
    ["Tutup Cup Plastik",1,"Piece"],["Sedotan Biasa",1,"Piece"]
  ]);

  // ── TOPPINGS ────────────────────────────────────────────────────────────
  addBOM("Keju",       [["Keju",15,"Gram"]]);
  addBOM("Chocolate",  [["Chocolate",15,"Gram"]]);
  addBOM("Chocochips", [["Chocochips",10,"Gram"]]);
  addBOM("Mesis",      [["Mesis",15,"Gram"]]);
  addBOM("Bubuk Oreo", [["Bubuk Oreo",15,"Gram"]]);
  addBOM("Boba",       [["Boba",20,"Gram"]]);

  return data;
}

// ── Legacy wrapper ──────────────────────────────────────────────────────────
function getResepData() { return generateBOMData(); }

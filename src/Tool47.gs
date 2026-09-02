/**
 * Tool47.gs — Product Page Research Capture.
 * Research columns: id, brand, product, price, url, archetype, description,
 * composition, construction, care, want, slotFront, slotBack, slotSide, slotDetail.
 * BoardImages columns: id, kind, name, note, state.
 *
 * Garment view slots and board images are the tab's own columns (authoritative,
 * cycled inline) and are also mirrored into the Images tab — the shared shot
 * list — so Tool 54 and the Image schedule can see them without reading this
 * tab's internals.
 */

var T47_SLOTS = [
  { k: "slotFront", label: "Front", req: true },
  { k: "slotBack", label: "Back", req: true },
  { k: "slotSide", label: "Side", req: false },
  { k: "slotDetail", label: "Detail", req: false }
];

function t47_get() {
  var cfg = getConfig_();
  var lookups = getLookups_();
  var meta = {
    name: cfg.studentName || "",
    direction: cfg.direction || "",
    archetype: cfg.archetypeClaimed || "",
    combineWhy: cfg.combineWhy || ""
  };
  return JSON.stringify({
    meta: meta,
    garments: readRows_("research"),
    boards: readRows_("boardImages"),
    archetypes: lookups.archetypes || [],
    boardKinds: lookups.boardKinds || [],
    slotStates: lookups.slotStates || ["Planned", "Captured", "On the board"]
  });
}

function t47_save(payloadJson) {
  var payload = JSON.parse(payloadJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (payload.meta) {
      setConfig_({
        studentName: payload.meta.name || "",
        direction: payload.meta.direction || "",
        archetypeClaimed: payload.meta.archetype || "",
        combineWhy: payload.meta.combineWhy || ""
      });
    }

    (payload.deletedGarmentIds || []).forEach(function (id) {
      deleteRowById_("research", id);
      img_deleteForParent_("47 Research", id);
    });
    (payload.deletedBoardIds || []).forEach(function (id) {
      deleteRowById_("boardImages", id);
      img_deleteForParent_("47 Research", id);
    });

    (payload.garments || []).forEach(function (g) {
      writeRowByField_("research", g.id, {
        brand: g.brand || "", product: g.product || "", price: g.price || "", url: g.url || "",
        archetype: g.archetype || "", description: g.description || "", composition: g.composition || "",
        construction: g.construction || "", care: g.care || "", want: g.want || "",
        slotFront: g.slotFront || "Planned", slotBack: g.slotBack || "Planned",
        slotSide: g.slotSide || "Planned", slotDetail: g.slotDetail || "Planned"
      }, []);
      var title = [g.brand, g.product].filter(String).join(" · ") || "Untitled garment";
      T47_SLOTS.forEach(function (s) {
        img_upsertSlot_("47 Research", g.id, s.label, title + " — " + s.label, "Screenshot", g[s.k] || "Planned");
      });
    });

    (payload.boards || []).forEach(function (b) {
      writeRowByField_("boardImages", b.id, {
        kind: b.kind || "", name: b.name || "", note: b.note || "", state: b.state || "Planned"
      }, []);
      img_upsertSlot_("47 Research", b.id, "board", (b.name || b.kind || "Board image"),
        (b.kind || "Detail"), b.state || "Planned");
    });

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t47_get();
}

function t47_addGarment() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_("research");
    var idx = headerIndex_("research");
    var lastCol = sh.getLastColumn();
    var row = [];
    for (var c = 0; c < lastCol; c++) row.push("");
    row[idx.id] = newId_(ID_PREFIX.research);
    row[idx.slotFront] = "Planned"; row[idx.slotBack] = "Planned";
    row[idx.slotSide] = "Planned"; row[idx.slotDetail] = "Planned";
    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t47_get();
}

function t47_addBoard() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_("boardImages");
    var idx = headerIndex_("boardImages");
    var lastCol = sh.getLastColumn();
    var row = [];
    for (var c = 0; c < lastCol; c++) row.push("");
    row[idx.id] = newId_(ID_PREFIX.boardImages);
    row[idx.state] = "Planned";
    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t47_get();
}

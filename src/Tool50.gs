/**
 * Tool50.gs — Finishing & Hardware Spec.
 * Materials columns: id, category, type, item, dimension, qty, unit, source,
 * status, unitValue, lineValue (formula — never write), placement, supplier,
 * substitutedFrom, referenceId, notes, stockN (formula — never write).
 * Finishing columns: id, order, location, operation, spec, thread, settings, referenceId.
 * StockSlip is formula-driven off Materials and is never read or written here —
 * the dialog's own "Stock slip" tab is a live preview computed from Materials,
 * same as the prototype; printing the real tab is the menu's job.
 */

var T50_MATERIALS_SKIP = ["lineValue", "stockN"];

function t50_get() {
  var cfg = getConfig_();
  var lookups = getLookups_();
  var meta = {
    name: cfg.studentName || "", group: cfg.group || "", item: cfg.item || "",
    currency: cfg.currency || "USD",
    slipRequested: cfg.slipRequested || "", slipApprovedBy: cfg.slipApprovedBy || "",
    slipApprovedOn: cfg.slipApprovedOn || ""
  };
  var refs = readRows_("research").map(function (g) {
    return { id: g.id, title: [g.brand, g.product].filter(String).join(" · ") || "Untitled garment" };
  });
  return JSON.stringify({
    meta: meta,
    materials: readRows_("materials"),
    finishing: readRows_("finishing"),
    refs: refs,
    cats: lookups.categories || [],
    types: lookups.types || [],
    sources: lookups.sources || [],
    status: lookups.matStatus || [],
    ops: lookups.finishOps || []
  });
}

function t50_save(payloadJson) {
  var payload = JSON.parse(payloadJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (payload.meta) {
      setConfig_({
        studentName: payload.meta.name || "", group: payload.meta.group || "", item: payload.meta.item || "",
        currency: payload.meta.currency || "USD",
        slipRequested: payload.meta.slipRequested || "", slipApprovedBy: payload.meta.slipApprovedBy || "",
        slipApprovedOn: payload.meta.slipApprovedOn || ""
      });
    }
    (payload.deletedMaterialIds || []).forEach(function (id) { deleteRowById_("materials", id); });
    (payload.deletedFinishingIds || []).forEach(function (id) { deleteRowById_("finishing", id); });

    (payload.materials || []).forEach(function (m) {
      writeRowByField_("materials", m.id, {
        category: m.category || "", type: m.type || "", item: m.item || "", dimension: m.dimension || "",
        qty: m.qty || "", unit: m.unit || "", source: m.source || "", status: m.status || "Not secured",
        unitValue: m.unitValue || "", placement: m.placement || "", supplier: m.supplier || "",
        substitutedFrom: m.substitutedFrom || "", referenceId: m.referenceId || "", notes: m.notes || ""
      }, T50_MATERIALS_SKIP);
    });
    (payload.finishing || []).forEach(function (f, i) {
      writeRowByField_("finishing", f.id, {
        order: i + 1, location: f.location || "", operation: f.operation || "", spec: f.spec || "",
        thread: f.thread || "", settings: f.settings || "", referenceId: f.referenceId || ""
      }, []);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t50_get();
}

function t50_addMaterial() {
  return t50_appendRow_("materials", { status: "Not secured" });
}
function t50_addFinishing() {
  return t50_appendRow_("finishing", {});
}

/** The "Add the usual shirt lines" starter — five common material lines. */
function t50_addStarter() {
  var starter = [
    { category: "Fabric", type: "Main fabric", placement: "" },
    { category: "Fabric", type: "Interlining / fusible", placement: "Collar, stand, cuffs, placket" },
    { category: "Hardware", type: "Button", placement: "Front placket" },
    { category: "Trim", type: "Thread", placement: "" },
    { category: "Label", type: "Brand label", placement: "Centred on back yoke, 2 cm below seam" }
  ];
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    starter.forEach(function (s) {
      t50_appendRowUnlocked_("materials", {
        category: s.category, type: s.type, placement: s.placement, status: "Not secured"
      });
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t50_get();
}

function t50_appendRow_(key, fields) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    t50_appendRowUnlocked_(key, fields);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t50_get();
}

function t50_appendRowUnlocked_(key, fields) {
  var sh = sheet_(key);
  var idx = headerIndex_(key);
  var lastCol = sh.getLastColumn();
  var row = [];
  for (var c = 0; c < lastCol; c++) row.push("");
  row[idx.id] = newId_(ID_PREFIX[key]);
  Object.keys(fields || {}).forEach(function (f) {
    if (idx[f] !== undefined) row[idx[f]] = fields[f];
  });
  sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
}

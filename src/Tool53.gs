/**
 * Tool53.gs — Revision Log.
 * Revisions columns: id, title, kind, predicted, actual, point, digitalVal,
 * physicalVal, delta (formula — never write), why, action, detail,
 * slotPredicted, slotPredictedState, slotActual, slotActualState.
 * "Seed from Tool 52" reads the Build tab directly — same workbook, no file
 * upload needed. Named slots mirror into the Images tab.
 */

var T53_SKIP = ["delta"];

function t53_get() {
  var cfg = getConfig_();
  var lookups = getLookups_();
  return JSON.stringify({
    meta: {
      name: cfg.studentName || "", item: cfg.item || "", route: cfg.route || "",
      unit: cfg.measurementUnit || "cm", closing: cfg.closingStatement || ""
    },
    entries: readRows_("revisions"),
    kinds: lookups.revisionKinds || [],
    actions: lookups.revisionActions || [],
    slotStates: (lookups.slotStates || ["Planned", "Captured", "On the board"]).slice(0, 2)
  });
}

function t53_save(payloadJson) {
  var payload = JSON.parse(payloadJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (payload.meta) {
      setConfig_({
        studentName: payload.meta.name || "", item: payload.meta.item || "",
        route: payload.meta.route || "", measurementUnit: payload.meta.unit || "cm",
        closingStatement: payload.meta.closing || ""
      });
    }
    (payload.deletedIds || []).forEach(function (id) {
      deleteRowById_("revisions", id);
      img_deleteForParent_("53 Revisions", id);
    });
    // Records first, image slots in one batch after — so a slow or failing
    // slot write can never leave records half-saved.
    var slots = [];
    (payload.entries || []).forEach(function (e) {
      writeRowByField_("revisions", e.id, {
        title: e.title || "", kind: e.kind || "", predicted: e.predicted || "", actual: e.actual || "",
        point: e.point || "", digitalVal: e.digitalVal || "", physicalVal: e.physicalVal || "",
        why: e.why || "", action: e.action || "", detail: e.detail || "",
        slotPredicted: (e.slotPredicted || "").trim(), slotPredictedState: e.slotPredictedState || "Planned",
        slotActual: (e.slotActual || "").trim(), slotActualState: e.slotActualState || "Planned"
      }, T53_SKIP);
      var title = e.title || "Untitled divergence";
      if (String(e.slotPredicted || "").trim()) {
        slots.push({ tool: "53 Revisions", parentId: e.id, name: "predicted",
          purpose: title + " — predicted (CLO3D)", type: "Detail",
          status: e.slotPredictedState || "Planned" });
      }
      if (String(e.slotActual || "").trim()) {
        slots.push({ tool: "53 Revisions", parentId: e.id, name: "actual",
          purpose: title + " — actual (garment)", type: "Photograph",
          status: e.slotActualState || "Planned" });
      }
    });
    img_upsertSlots_(slots);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t53_get();
}

function t53_addEntry() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_("revisions");
    var idx = headerIndex_("revisions");
    var lastCol = sh.getLastColumn();
    var row = [];
    for (var c = 0; c < lastCol; c++) row.push("");
    row[idx.id] = newId_(ID_PREFIX.revisions);
    row[idx.slotPredictedState] = "Planned";
    row[idx.slotActualState] = "Planned";
    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t53_get();
}

/** Creates an entry for every deviation the Build tab already recorded. */
function t53_seedFromBuild() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var added = 0;
  try {
    var steps = readRows_("build");
    var sh = sheet_("revisions");
    var idx = headerIndex_("revisions");
    var lastCol = sh.getLastColumn();
    function push(title, kind, actual) {
      var row = [];
      for (var c = 0; c < lastCol; c++) row.push("");
      row[idx.id] = newId_(ID_PREFIX.revisions);
      row[idx.title] = title; row[idx.kind] = kind; row[idx.actual] = actual;
      row[idx.slotPredictedState] = "Planned"; row[idx.slotActualState] = "Planned";
      sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
      added++;
    }
    steps.forEach(function (s, i) {
      var planned = i + 1;
      var actualOrder = parseFloat(s.actualOrder);
      if (actualOrder && actualOrder !== planned) {
        push(s.location + " sewn out of planned order", "Construction sequence",
          "Planned as step " + planned + ", actually done " + actualOrder + ".");
      }
      if (s.unplanned) {
        push(s.location + " was not in the plan", "Construction sequence",
          "This step did not exist in the planned sequence.");
      }
      if (s.attempt === "Redoing a mistake") {
        push(s.location + " redone", "Something else", "Redone after a mistake.");
      }
      String(s.problems || "").split(",").map(function (p) { return p.trim(); }).filter(Boolean).forEach(function (p) {
        var kind = /pattern/i.test(p) ? "Pattern" : (/fabric/i.test(p) ? "Drape / fabric behaviour" : "Something else");
        push(s.location + " — " + p.toLowerCase(), kind, "Recorded during construction: " + p + ".");
      });
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  var out = JSON.parse(t53_get());
  out.added = added;
  return JSON.stringify(out);
}

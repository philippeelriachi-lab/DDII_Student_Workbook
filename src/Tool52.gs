/**
 * Tool52.gs — Construction Log.
 * Build columns: id, plannedOrder, location, operation, pieces, prep, press,
 * thread, settings, unplanned, actualOrder, prepSec, sewSec, pressSec,
 * attempt, problems, photos, startedAt, endedAt.
 * pieces/prep/press/problems are stored as comma-joined strings.
 * Pattern pieces (the master checklist) live in Config.patternPieces,
 * newline-joined — there is no dedicated tab for them.
 */

function t52_splitList_(s) {
  return String(s || "").split(",").map(function (x) { return x.trim(); }).filter(Boolean);
}
function t52_joinList_(a) { return (a || []).join(","); }

function t52_get() {
  var cfg = getConfig_();
  var lookups = getLookups_();
  var meta = {
    name: cfg.studentName || "", item: cfg.item || "", route: cfg.route || "",
    rate: cfg.labourRate || "", currency: cfg.currency || "USD",
    basis: cfg.rateBasis || "", rateSource: cfg.rateSource || "",
    curve: cfg.curveFactor ? parseFloat(cfg.curveFactor) : 0.6,
    lead: cfg.leadTimeOn === true || cfg.leadTimeOn === "TRUE"
  };
  var pieces = String(cfg.patternPieces || "").split("\n").map(function (x) { return x.trim(); }).filter(Boolean);

  var steps = readRows_("build").map(function (s) {
    return {
      id: s.id, plannedOrder: s.plannedOrder, location: s.location || "", operation: s.operation || "",
      pieces: t52_splitList_(s.pieces), prep: t52_splitList_(s.prep), press: t52_splitList_(s.press),
      thread: s.thread || "", settings: s.settings || "", unplanned: !!s.unplanned,
      actualOrder: s.actualOrder || "", prepSec: parseFloat(s.prepSec) || 0, sewSec: parseFloat(s.sewSec) || 0,
      pressSec: parseFloat(s.pressSec) || 0, attempt: s.attempt || "First time",
      problems: t52_splitList_(s.problems), photos: parseInt(s.photos) || 0,
      startedAt: s.startedAt || "", endedAt: s.endedAt || ""
    };
  });
  steps.sort(function (a, b) { return (parseFloat(a.plannedOrder) || 0) - (parseFloat(b.plannedOrder) || 0); });

  return JSON.stringify({
    meta: meta, pieces: pieces, steps: steps,
    seamTypes: lookups.seamTypes || [], prepOps: lookups.prepOps || [], pressOps: lookups.pressOps || [],
    attempts: lookups.attempts || ["First time", "Done it before", "Redoing a mistake"],
    problems: lookups.problems || ["Machine", "Fabric", "Pattern was wrong", "My error", "Ran out of something", "Out of time"]
  });
}

function t52_save(payloadJson) {
  var payload = JSON.parse(payloadJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (payload.meta) {
      setConfig_({
        studentName: payload.meta.name || "", item: payload.meta.item || "", route: payload.meta.route || "",
        labourRate: payload.meta.rate || "", currency: payload.meta.currency || "USD",
        rateBasis: payload.meta.basis || "", rateSource: payload.meta.rateSource || "",
        curveFactor: payload.meta.curve || 0.6, leadTimeOn: !!payload.meta.lead
      });
    }
    if (payload.pieces) setConfig_({ patternPieces: payload.pieces.join("\n") });

    (payload.deletedIds || []).forEach(function (id) { deleteRowById_("build", id); });

    (payload.steps || []).forEach(function (s, i) {
      writeRowByField_("build", s.id, {
        plannedOrder: i + 1, location: s.location || "", operation: s.operation || "",
        pieces: t52_joinList_(s.pieces), prep: t52_joinList_(s.prep), press: t52_joinList_(s.press),
        thread: s.thread || "", settings: s.settings || "", unplanned: !!s.unplanned,
        actualOrder: s.actualOrder || "", prepSec: s.prepSec || 0, sewSec: s.sewSec || 0,
        pressSec: s.pressSec || 0, attempt: s.attempt || "First time",
        problems: t52_joinList_(s.problems), photos: s.photos || 0,
        startedAt: s.startedAt || "", endedAt: s.endedAt || ""
      }, []);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t52_get();
}

function t52_addStep(fieldsJson) {
  var fields = fieldsJson ? JSON.parse(fieldsJson) : {};
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_("build");
    var idx = headerIndex_("build");
    var existing = readRows_("build");
    var maxOrder = existing.reduce(function (m, s) { return Math.max(m, parseFloat(s.plannedOrder) || 0); }, 0);
    var lastCol = sh.getLastColumn();
    var row = [];
    for (var c = 0; c < lastCol; c++) row.push("");
    row[idx.id] = newId_(ID_PREFIX.build);
    row[idx.plannedOrder] = maxOrder + 1;
    row[idx.attempt] = "First time";
    row[idx.unplanned] = !!fields.unplanned;
    if (fields.location) row[idx.location] = fields.location;
    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t52_get();
}

/** Bulk-adds steps (e.g. the standard shirt sequence) with real ids in one pass. */
function t52_addSteps(stepsJson) {
  var newSteps = JSON.parse(stepsJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_("build");
    var idx = headerIndex_("build");
    var existing = readRows_("build");
    var order = existing.reduce(function (m, s) { return Math.max(m, parseFloat(s.plannedOrder) || 0); }, 0);
    var lastCol = sh.getLastColumn();
    newSteps.forEach(function (ns) {
      order++;
      var row = [];
      for (var c = 0; c < lastCol; c++) row.push("");
      row[idx.id] = newId_(ID_PREFIX.build);
      row[idx.plannedOrder] = order;
      row[idx.location] = ns.location || "";
      row[idx.operation] = ns.operation || "";
      row[idx.pieces] = (ns.pieces || []).join(",");
      row[idx.prep] = (ns.prep || []).join(",");
      row[idx.press] = (ns.press || []).join(",");
      row[idx.attempt] = "First time";
      sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t52_get();
}

/** Loads finishing operations from Tool 50 as starting Build steps. */
function t52_fromFinishing() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var added = 0;
  try {
    var fins = readRows_("finishing");
    var sh = sheet_("build");
    var idx = headerIndex_("build");
    var existing = readRows_("build");
    var order = existing.reduce(function (m, s) { return Math.max(m, parseFloat(s.plannedOrder) || 0); }, 0);
    var lastCol = sh.getLastColumn();
    fins.forEach(function (f) {
      order++;
      var row = [];
      for (var c = 0; c < lastCol; c++) row.push("");
      row[idx.id] = newId_(ID_PREFIX.build);
      row[idx.plannedOrder] = order;
      row[idx.location] = f.location || "";
      row[idx.thread] = f.thread || "";
      row[idx.settings] = f.settings || "";
      row[idx.attempt] = "First time";
      sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
      added++;
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  var out = JSON.parse(t52_get());
  out.added = added;
  return JSON.stringify(out);
}

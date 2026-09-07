/**
 * Shared menu, sheet map and helpers for every Tool*.gs file.
 * Sheet names and id prefixes below are the single source of truth —
 * every other file looks a tab up through sheet_(key), never by literal name.
 */

var SHEET_MAP = {
  config: "Config",
  lookups: "Lookups",
  readiness: "Readiness",
  sizing: "Sizing",
  research: "Research",
  boardImages: "BoardImages",
  iterations: "Iterations",
  materials: "Materials",
  finishing: "Finishing",
  stockSlip: "StockSlip",
  pitch: "Pitch",
  build: "Build",
  revisions: "Revisions",
  images: "Images"
};

// Prefixes for newId_(): sz r b i m f p bd v g
var ID_PREFIX = {
  sizing: "sz",
  research: "r",
  boardImages: "bd",
  iterations: "i",
  materials: "m",
  finishing: "f",
  pitch: "p",
  build: "b",
  revisions: "v",
  images: "g"
};

// Tabs that carry record rows with an id column — everything Tidy up / Diagnose sweep.
var ID_BEARING_KEYS = ["sizing", "research", "boardImages", "iterations", "materials",
  "finishing", "pitch", "build", "revisions", "images"];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Tools")
    .addItem("A · Sizing intake", "openToolA")
    .addItem("B · Research capture", "openToolB")
    .addItem("C · Iteration log", "openToolC")
    .addItem("E · Finishing and hardware spec", "openToolE")
    .addItem("F · Pitch planner", "openToolF")
    .addItem("G · Construction log", "openToolG")
    .addItem("H · Revision log", "openToolH")
    .addItem("I · Submission readiness", "openToolI")
    .addSeparator()
    .addItem("Print the stock slip", "printStockSlip")
    .addItem("Image schedule", "openImageSchedule")
    .addSeparator()
    .addItem("Tidy up identifiers", "tidyUpIdentifiers")
    .addItem("Diagnose", "runDiagnose")
    .addItem("Run the A–I migration (once)", "runLetterMigration")
    .addItem("Clean up the old final round", "cleanUpFinalRound")
    .addToUi();
}

/**
 * Looks up a tab by its logical key. Throws a clear, actionable error —
 * naming the tab it expected and listing what actually exists — rather
 * than letting a null sheet reference fail somewhere unrelated later.
 */
function sheet_(key) {
  var name = SHEET_MAP[key];
  if (!name) throw new Error("Unknown sheet key '" + key + "'.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    var found = ss.getSheets().map(function (s) { return s.getName(); }).join(", ");
    throw new Error("Missing tab '" + name + "'. Tabs found: " + found);
  }
  return sh;
}

function toPlain_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || "Asia/Beirut", "yyyy-MM-dd");
  }
  return v;
}

function rowIsBlank_(row) {
  for (var i = 0; i < row.length; i++) {
    var v = row[i];
    if (v !== "" && v !== null && v !== undefined) return false;
  }
  return true;
}

/** Config is key/value/description. Returns {key: value, ...}, dates normalized. */
function getConfig_() {
  var sh = sheet_("config");
  var values = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < values.length; i++) {
    var k = values[i][0];
    if (!k) continue;
    out[String(k)] = toPlain_(values[i][1]);
  }
  return out;
}

/** Merges patch into Config — updates existing keys in place, appends new ones. */
function setConfig_(patch) {
  var sh = sheet_("config");
  var values = sh.getDataRange().getValues();
  var rowOf = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) rowOf[String(values[i][0])] = i + 1;
  }
  Object.keys(patch || {}).forEach(function (k) {
    var v = patch[k];
    if (rowOf[k]) {
      sh.getRange(rowOf[k], 2).setValue(v);
    } else {
      sh.appendRow([k, v, ""]);
    }
  });
}

/**
 * Every dropdown list, one per column, name in row 1. Skips blank columns,
 * trims headers, stops each list at its own last non-blank value.
 */
function getLookups_() {
  var sh = sheet_("lookups");
  var values = sh.getDataRange().getValues();
  if (!values.length) return {};
  var headers = values[0];
  var out = {};
  for (var c = 0; c < headers.length; c++) {
    var name = String(headers[c] || "").trim();
    if (!name) continue;
    var list = [];
    for (var r = 1; r < values.length; r++) {
      var v = values[r][c];
      if (v === "" || v === null || v === undefined) continue;
      list.push(toPlain_(v));
    }
    out[name] = list;
  }
  return out;
}

/**
 * Reads every non-blank record row of a tab as an object keyed by its
 * header row. Rows with content but no id are backfilled right here —
 * "backfill ids on read" — and the id is written back immediately.
 */
function readRows_(key) {
  var sh = sheet_(key);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function (h) { return String(h || "").trim(); });
  var idCol = headers.indexOf("id");
  var out = [];
  var backfills = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (rowIsBlank_(row)) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = toPlain_(row[c]);
    }
    if (idCol >= 0 && !obj.id) {
      var newId = newId_(ID_PREFIX[key] || "x");
      obj.id = newId;
      backfills.push({ row: r + 1, id: newId });
    }
    out.push(obj);
  }
  if (backfills.length) {
    backfills.forEach(function (b) { sh.getRange(b.row, idCol + 1).setValue(b.id); });
  }
  return out;
}

/** Bulk id backfill for one tab. Returns how many rows were fixed. */
function ensureIds_(key) {
  var sh = sheet_(key);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return 0;
  var headers = values[0].map(function (h) { return String(h || "").trim(); });
  var idCol = headers.indexOf("id");
  if (idCol < 0) return 0;
  var fixed = 0;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (rowIsBlank_(row)) continue;
    if (!row[idCol]) {
      sh.getRange(r + 1, idCol + 1).setValue(newId_(ID_PREFIX[key] || "x"));
      fixed++;
    }
  }
  return fixed;
}

function newId_(prefix) {
  return prefix + Utilities.getUuid().replace(/-/g, "").slice(0, 8);
}

/** Header index map for a tab, e.g. {id:0, brand:1, product:2, ...}. */
function headerIndex_(key) {
  var sh = sheet_(key);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idx = {};
  headers.forEach(function (h, i) { var n = String(h || "").trim(); if (n) idx[n] = i; });
  return idx;
}

/**
 * Finds the 1-based sheet row for a given id in a tab, or 0 if not found.
 * A tab holding only its header has nothing to search — asking for a
 * zero-row range there throws, which used to abort whole saves mid-way.
 */
function findRowById_(key, id) {
  var sh = sheet_(key);
  var idx = headerIndex_(key);
  if (idx.id === undefined) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, idx.id + 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

/**
 * Writes fields onto an existing row by id, skipping any column named in
 * `skip` (formula columns a tool must never touch). Appends a new row if
 * the id isn't found yet. Returns the row number written.
 */
function writeRowByField_(key, id, fields, skip) {
  var sh = sheet_(key);
  var idx = headerIndex_(key);
  var skipSet = {};
  (skip || []).forEach(function (s) { skipSet[s] = true; });
  var row = findRowById_(key, id);
  var isNew = !row;
  if (isNew) row = sh.getLastRow() + 1;

  var byCol = {};
  if (isNew && idx.id !== undefined) byCol[idx.id] = id;
  Object.keys(fields).forEach(function (f) {
    if (skipSet[f]) return;
    if (idx[f] === undefined) return;
    byCol[idx[f]] = fields[f];
  });

  // Write contiguous runs of columns in single calls. Skipped columns break a
  // run rather than being written over, so formula columns stay untouched.
  var cols = Object.keys(byCol).map(Number).sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < cols.length) {
    var j = i;
    while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j++;
    var values = [];
    for (var c = i; c <= j; c++) values.push(byCol[cols[c]]);
    sh.getRange(row, cols[i] + 1, 1, values.length).setValues([values]);
    i = j + 1;
  }
  return row;
}

/** Deletes the row matching an id, if any. */
function deleteRowById_(key, id) {
  var row = findRowById_(key, id);
  if (row) sheet_(key).deleteRow(row);
}

/** Finds or creates the Drive folder that holds this workbook's images. */
function imagesFolder_() {
  var cfg = getConfig_();
  if (cfg.imagesFolderId) {
    try {
      return DriveApp.getFolderById(cfg.imagesFolderId);
    } catch (e) {
      // stale id — fall through and recreate
    }
  }
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parents = file.getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var name = SpreadsheetApp.getActiveSpreadsheet().getName() + " — images";
  var existing = parent.getFoldersByName(name);
  var folder = existing.hasNext() ? existing.next() : parent.createFolder(name);
  setConfig_({ imagesFolderId: folder.getId() });
  return folder;
}

function showToolDialog_(templateName, title, width, height) {
  var html = HtmlService.createTemplateFromFile(templateName).evaluate()
    .setWidth(width || 960).setHeight(height || 720);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/** Lets an Editor*.html template pull in shared client-side helpers, e.g. <?!= include_('Shared') ?> */
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function openToolA() { showToolDialog_("EditorA", "Tool A · Sizing Intake", 1040, 760); }
function openToolB() { showToolDialog_("EditorB", "Tool B · Product Page Research Capture", 1040, 760); }
function openToolC() { showToolDialog_("EditorC", "Tool C · Iteration Log", 1040, 760); }
function openToolE() { showToolDialog_("EditorE", "Tool E · Finishing & Hardware Spec", 1060, 780); }
function openToolF() { showToolDialog_("EditorF", "Tool F · Technical Pitch Planner", 980, 760); }
function openToolG() { showToolDialog_("EditorG", "Tool G · Construction Log", 920, 780); }
function openToolH() { showToolDialog_("EditorH", "Tool H · Revision Log", 1000, 760); }
function openToolI() { showToolDialog_("EditorI", "Tool I · Submission Readiness", 960, 780); }

function printStockSlip() {
  var sh = sheet_("stockSlip");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
  SpreadsheetApp.getUi().alert("The stock slip is now the active tab. Use File → Print to print or save it as a PDF.");
}

function openImageSchedule() {
  showToolDialog_("ImageSchedule", "Image schedule", 820, 640);
}

function tidyUpIdentifiers() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var total = 0;
  try {
    ID_BEARING_KEYS.forEach(function (key) { total += ensureIds_(key); });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.getUi().alert("Tidy up identifiers",
    total + " row" + (total === 1 ? "" : "s") + " given an id.",
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function runDiagnose() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = [];
  lines.push("File: " + ss.getName());
  lines.push("");
  lines.push("Tabs:");
  ss.getSheets().forEach(function (sh) {
    lines.push("  " + sh.getName() + (sh.isSheetHidden() ? "  (hidden)" : ""));
  });

  var lookups = getLookups_();
  var names = Object.keys(lookups);
  lines.push("");
  lines.push("Lookup lists (" + names.length + "):");
  names.forEach(function (n) { lines.push("  " + n + ": " + lookups[n].length); });

  var missing = 0;
  ID_BEARING_KEYS.forEach(function (key) {
    var sh;
    try { sh = sheet_(key); } catch (e) { return; }
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return;
    var headers = values[0].map(function (h) { return String(h || "").trim(); });
    var idCol = headers.indexOf("id");
    if (idCol < 0) return;
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (rowIsBlank_(row)) continue;
      if (!row[idCol]) missing++;
    }
  });
  lines.push("");
  lines.push(missing + " row" + (missing === 1 ? "" : "s") + " without an id.");

  SpreadsheetApp.getUi().alert("Diagnose", lines.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}

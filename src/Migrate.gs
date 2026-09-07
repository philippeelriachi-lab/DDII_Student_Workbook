/**
 * Migrate.gs — one-time data migration for the 46–54 → A–I re-lettering.
 *
 * Three sets of values are stored in the workbook itself, not just in code:
 * the Images tab's tool labels, two Config keys named after the old juries,
 * and the Pitch tab's round labels. Renaming them in code alone would orphan
 * every existing row, so this rewrites the stored values to match.
 *
 * Safe to run more than once — anything already migrated is skipped. If a
 * tool was opened between pushing the new code and running this, the new
 * code will have written rows under the new labels alongside the old ones;
 * each pair is merged here rather than left as a duplicate.
 */

var MIG_IMAGE_TOOLS = {
  "47 Research": "B Research",
  "48 Iterations": "C Iterations",
  "53 Revisions": "H Revisions"
};

var MIG_CONFIG_KEYS = {
  jurySelfChecks: "readinessSelfChecks",
  juryRound: "readinessRound"
};

var MIG_ROUND_LABELS = {
  "Session 11 · Midterm jury": "Session 11 · Midterm presentation",
  "Final jury": "Final submission"
};

function runLetterMigration() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var lines = [];
  try {
    lines.push("Images tab");
    lines.push("  " + mig_images_());
    lines.push("");
    lines.push("Config keys");
    lines.push("  " + mig_config_());
    lines.push("");
    lines.push("Pitch rounds");
    lines.push("  " + mig_rounds_());
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  ui.alert("A–I migration", lines.join("\n"), ui.ButtonSet.OK);
}

/**
 * Relabels the Images tab's tool column. Where the same slot already exists
 * under both labels, the two rows are collapsed into one — the row carrying
 * an uploaded file wins, and the old row wins a tie, because it is the one
 * that predates the rename.
 */
function mig_images_() {
  var sh = sheet_("images");
  var idx = headerIndex_("images");
  if (idx.tool === undefined) return "no tool column — nothing to do";
  var last = sh.getLastRow();
  if (last < 2) return "empty — nothing to do";

  var lastCol = sh.getLastColumn();
  var values = sh.getRange(2, 1, last - 1, lastCol).getValues();

  // Every slot already living under a new label, so collisions are visible.
  var newRowAt = {};
  values.forEach(function (row, i) {
    var t = String(row[idx.tool] || "");
    if (mig_isNewLabel_(t)) newRowAt[t + "|" + row[idx.parentId] + "|" + row[idx.name]] = i;
  });

  var relabelled = 0, merged = 0;
  var dropRows = [];
  values.forEach(function (row, i) {
    var oldLabel = String(row[idx.tool] || "");
    var newLabel = MIG_IMAGE_TOOLS[oldLabel];
    if (!newLabel) return;
    var key = newLabel + "|" + row[idx.parentId] + "|" + row[idx.name];
    var twin = newRowAt[key];
    if (twin !== undefined) {
      // Same slot under both labels. Keep whichever holds a real upload.
      var fc = idx.driveFileId;
      var twinHasFile = fc !== undefined && String(values[twin][fc] || "").trim() !== "";
      var mineHasFile = fc !== undefined && String(row[fc] || "").trim() !== "";
      if (twinHasFile && !mineHasFile) {
        dropRows.push(i);
      } else {
        dropRows.push(twin);
        row[idx.tool] = newLabel;
        newRowAt[key] = i;
      }
      merged++;
      return;
    }
    row[idx.tool] = newLabel;
    newRowAt[key] = i;
    relabelled++;
  });

  if (!relabelled && !merged) return "already migrated — nothing to do";

  sh.getRange(2, 1, values.length, lastCol).setValues(values);
  // Bottom up, so earlier indices stay valid as rows disappear.
  dropRows.sort(function (a, b) { return b - a; }).forEach(function (i) {
    sh.deleteRow(i + 2);
  });

  return relabelled + " slot" + (relabelled === 1 ? "" : "s") + " relabelled" +
    (merged ? ", " + merged + " duplicate" + (merged === 1 ? "" : "s") + " merged" : "");
}

function mig_isNewLabel_(t) {
  for (var k in MIG_IMAGE_TOOLS) {
    if (MIG_IMAGE_TOOLS[k] === t) return true;
  }
  return false;
}

/**
 * Renames the two Config keys. jurySelfChecks is a JSON blob of ticks — if
 * both names somehow hold one, they are merged rather than one overwriting
 * the other, with the newly-named blob winning any single disputed tick.
 */
function mig_config_() {
  var sh = sheet_("config");
  var last = sh.getLastRow();
  if (last < 2) return "empty — nothing to do";
  var values = sh.getRange(2, 1, last - 1, 2).getValues();

  var rowOf = {};
  values.forEach(function (r, i) { if (r[0]) rowOf[String(r[0])] = i; });

  var done = [];
  var dropRows = [];
  Object.keys(MIG_CONFIG_KEYS).forEach(function (oldKey) {
    var newKey = MIG_CONFIG_KEYS[oldKey];
    if (rowOf[oldKey] === undefined) return;
    var oldRow = rowOf[oldKey];
    if (rowOf[newKey] !== undefined) {
      var newRow = rowOf[newKey];
      if (oldKey === "jurySelfChecks") {
        values[newRow][1] = JSON.stringify(
          mig_mergeChecks_(values[oldRow][1], values[newRow][1]));
      }
      dropRows.push(oldRow);
    } else {
      values[oldRow][0] = newKey;
    }
    done.push(oldKey + " → " + newKey);
  });

  if (!done.length) return "already migrated — nothing to do";

  sh.getRange(2, 1, values.length, 2).setValues(values);
  dropRows.sort(function (a, b) { return b - a; }).forEach(function (i) {
    sh.deleteRow(i + 2);
  });
  return done.join("; ");
}

function mig_mergeChecks_(oldJson, newJson) {
  var a = {}, b = {};
  try { a = JSON.parse(oldJson || "{}"); } catch (e) { a = {}; }
  try { b = JSON.parse(newJson || "{}"); } catch (e) { b = {}; }
  Object.keys(b).forEach(function (k) { a[k] = b[k]; });
  return a;
}

/**
 * Rewrites the round labels wherever they are stored: the Pitch tab's round
 * column, and the pitchRounds dropdown list on Lookups that feeds it.
 */
function mig_rounds_() {
  var n = 0;
  n += mig_replaceColumn_("pitch", "round", MIG_ROUND_LABELS);
  n += mig_replaceLookupList_("pitchRounds", MIG_ROUND_LABELS);
  return n ? n + " cell" + (n === 1 ? "" : "s") + " rewritten" : "already migrated — nothing to do";
}

function mig_replaceColumn_(key, header, map) {
  var sh = sheet_(key);
  var idx = headerIndex_(key);
  if (idx[header] === undefined) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var rng = sh.getRange(2, idx[header] + 1, last - 1, 1);
  var vals = rng.getValues();
  var n = 0;
  vals.forEach(function (r) {
    var v = map[String(r[0])];
    if (v) { r[0] = v; n++; }
  });
  if (n) rng.setValues(vals);
  return n;
}

function mig_replaceLookupList_(listName, map) {
  var col = mig_lookupColumn_(listName);
  if (!col) return 0;
  var sh = sheet_("lookups");
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var rng = sh.getRange(2, col, last - 1, 1);
  var vals = rng.getValues();
  var n = 0;
  vals.forEach(function (r) {
    var v = map[String(r[0])];
    if (v) { r[0] = v; n++; }
  });
  if (n) rng.setValues(vals);
  return n;
}

/* ---------------------------------------------------------------------------
 * Final-round cleanup.
 *
 * Tool F no longer plans a pitch for the final round, because the final is
 * handed in rather than spoken. Anything left behind from when it did is
 * inert — nothing reads it — but it still sits in the workbook. This removes
 * it. Unlike the migration above this deletes rows, so it counts first, says
 * exactly what it is about to remove, and does nothing without a yes.
 * ------------------------------------------------------------------------- */

// Both spellings, so a workbook cleaned up before the migration ran is caught too.
var MIG_FINAL_LABELS = ["Final submission", "Final jury"];

function cleanUpFinalRound() {
  var ui = SpreadsheetApp.getUi();
  var found = mig_surveyFinal_();

  if (!found.pitchRows && !found.lookupCells && !found.configRows) {
    ui.alert("Final-round cleanup", "Nothing left to remove.", ui.ButtonSet.OK);
    return;
  }

  var what = [];
  if (found.pitchRows) {
    what.push("  " + found.pitchRows + " row" + (found.pitchRows === 1 ? "" : "s") +
      " on the Pitch tab — the sections and questions you planned for the final round");
  }
  if (found.lookupCells) {
    what.push("  the final round's entry in the pitchRounds list on Lookups");
  }
  if (found.configRows) {
    what.push("  the unresolvedFinal key on Config");
  }

  var answer = ui.alert("Final-round cleanup",
    "This permanently deletes:\n\n" + what.join("\n") +
    "\n\nNothing reads any of it any more. There is no undo other than the " +
    "file's own version history. Delete it?",
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var done;
  try {
    done = { pitchRows: mig_deleteFinalPitchRows_(),
             lookupCells: mig_clearFinalLookup_(),
             configRows: mig_deleteConfigKeys_(["unresolvedFinal"]) };
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  ui.alert("Final-round cleanup",
    done.pitchRows + " Pitch row" + (done.pitchRows === 1 ? "" : "s") + " deleted\n" +
    done.lookupCells + " Lookups cell" + (done.lookupCells === 1 ? "" : "s") + " cleared\n" +
    done.configRows + " Config key" + (done.configRows === 1 ? "" : "s") + " removed",
    ui.ButtonSet.OK);
}

function mig_isFinalLabel_(v) {
  return MIG_FINAL_LABELS.indexOf(String(v)) >= 0;
}

/** Counts what the cleanup would remove, without touching anything. */
function mig_surveyFinal_() {
  var out = { pitchRows: 0, lookupCells: 0, configRows: 0 };

  var pitch = sheet_("pitch");
  var pidx = headerIndex_("pitch");
  if (pidx.round !== undefined && pitch.getLastRow() > 1) {
    pitch.getRange(2, pidx.round + 1, pitch.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (mig_isFinalLabel_(r[0])) out.pitchRows++; });
  }

  var col = mig_lookupColumn_("pitchRounds");
  if (col > 0) {
    var lk = sheet_("lookups");
    if (lk.getLastRow() > 1) {
      lk.getRange(2, col, lk.getLastRow() - 1, 1).getValues()
        .forEach(function (r) { if (mig_isFinalLabel_(r[0])) out.lookupCells++; });
    }
  }

  var cfg = sheet_("config");
  if (cfg.getLastRow() > 1) {
    cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (String(r[0]) === "unresolvedFinal") out.configRows++; });
  }
  return out;
}

/** 1-based column of a named list on Lookups, or 0 if there isn't one. */
function mig_lookupColumn_(listName) {
  var sh = sheet_("lookups");
  var lastCol = sh.getLastColumn();
  if (!lastCol) return 0;
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c] || "").trim() === listName) return c + 1;
  }
  return 0;
}

function mig_deleteFinalPitchRows_() {
  var sh = sheet_("pitch");
  var idx = headerIndex_("pitch");
  if (idx.round === undefined) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, idx.round + 1, last - 1, 1).getValues();
  var rows = [];
  vals.forEach(function (r, i) { if (mig_isFinalLabel_(r[0])) rows.push(i + 2); });
  // Bottom up, so the row numbers below stay valid as rows disappear.
  rows.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  return rows.length;
}

/**
 * Blanks the final round's entry in the pitchRounds list. The cell is cleared
 * rather than the row deleted — that column shares its rows with every other
 * lookup list on the tab, so deleting the row would take an unrelated value
 * out of a different list with it.
 */
function mig_clearFinalLookup_() {
  var col = mig_lookupColumn_("pitchRounds");
  if (!col) return 0;
  var sh = sheet_("lookups");
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var rng = sh.getRange(2, col, last - 1, 1);
  var vals = rng.getValues();
  var n = 0;
  vals.forEach(function (r) { if (mig_isFinalLabel_(r[0])) { r[0] = ""; n++; } });
  if (n) rng.setValues(vals);
  return n;
}

function mig_deleteConfigKeys_(keys) {
  var sh = sheet_("config");
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  var rows = [];
  vals.forEach(function (r, i) { if (keys.indexOf(String(r[0])) >= 0) rows.push(i + 2); });
  rows.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  return rows.length;
}

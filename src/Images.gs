/**
 * Images.gs — shared slot registry helpers for Tools 47, 48 and 53.
 * The Images tab is the shot list: one row per slot, whichever tool created it.
 * estMinutes defaults by type, read from Lookups — imageTypes and imageTypeMinutes
 * are two independent lookup lists, paired by position (not by sheet row), so
 * adding a type/minutes pair to Lookups changes the default with no code change.
 */

function imageTypeMinutesMap_() {
  var lookups = getLookups_();
  var types = lookups.imageTypes || [];
  var minutes = lookups.imageTypeMinutes || [];
  var map = {};
  for (var i = 0; i < types.length; i++) {
    var m = parseFloat(minutes[i]);
    map[types[i]] = isFinite(m) ? m : 10;
  }
  return map;
}

function img_all_() {
  return readRows_("images");
}

function img_listForParent_(tool, parentId) {
  return img_all_().filter(function (r) {
    return String(r.tool) === String(tool) && String(r.parentId) === String(parentId);
  });
}

function img_find_(tool, parentId, name) {
  var rows = img_listForParent_(tool, parentId);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].name) === String(name)) return rows[i];
  }
  return null;
}

/** Creates the slot if it doesn't exist yet, or updates it. Returns the resulting row. */
function img_upsertSlot_(tool, parentId, name, purpose, type, status) {
  var existing = img_find_(tool, parentId, name);
  var id = existing ? existing.id : newId_(ID_PREFIX.images);
  var fields = {
    tool: tool,
    parentId: parentId,
    name: name,
    purpose: purpose !== undefined ? purpose : (existing ? existing.purpose : ""),
    type: type !== undefined ? type : (existing ? existing.type : ""),
    status: status !== undefined ? status : (existing ? existing.status : "Planned")
  };
  if (!existing) {
    var minutes = imageTypeMinutesMap_();
    fields.estMinutes = minutes[fields.type] !== undefined ? minutes[fields.type] : 10;
    fields.driveFileId = "";
    fields.notes = "";
  }
  writeRowByField_("images", id, fields, []);
  return img_find_(tool, parentId, name);
}

/**
 * Upserts many slots in one pass. Reads the Images tab once and appends all
 * new rows in a single write — img_upsertSlot_ re-reads the whole tab per
 * slot, which is fine for one slot and far too slow for a whole save.
 * Each entry: {tool, parentId, name, purpose, type, status}.
 */
function img_upsertSlots_(slots) {
  if (!slots || !slots.length) return;
  var sh = sheet_("images");
  var idx = headerIndex_("images");
  var minutes = imageTypeMinutesMap_();
  var byKey = {};
  img_all_().forEach(function (r) {
    byKey[r.tool + "|" + r.parentId + "|" + r.name] = r;
  });
  var lastCol = sh.getLastColumn();
  var appends = [];
  slots.forEach(function (s) {
    var k = s.tool + "|" + s.parentId + "|" + s.name;
    var found = byKey[k];
    if (found) {
      if (found.purpose !== s.purpose || found.type !== s.type || found.status !== s.status) {
        writeRowByField_("images", found.id, { purpose: s.purpose, type: s.type, status: s.status }, []);
      }
      return;
    }
    var row = [];
    for (var c = 0; c < lastCol; c++) row.push("");
    row[idx.id] = newId_(ID_PREFIX.images);
    row[idx.tool] = s.tool;
    row[idx.parentId] = s.parentId;
    row[idx.name] = s.name;
    row[idx.purpose] = s.purpose || "";
    row[idx.type] = s.type || "";
    row[idx.status] = s.status || "Planned";
    row[idx.estMinutes] = minutes[s.type] !== undefined ? minutes[s.type] : 10;
    appends.push(row);
    byKey[k] = { id: row[idx.id], purpose: s.purpose, type: s.type, status: s.status };
  });
  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, lastCol).setValues(appends);
  }
}

/** Cycles a slot's status through the given list, creating the slot first if needed. */
function img_cycleSlot_(tool, parentId, name, purpose, type, states) {
  var existing = img_find_(tool, parentId, name);
  var current = existing ? existing.status : states[0];
  var i = states.indexOf(current);
  var next = states[(i + 1 + states.length) % states.length];
  return img_upsertSlot_(tool, parentId, name, purpose, type, next).status;
}

/** Removes every slot registered against one parent record (e.g. a deleted garment). */
function img_deleteForParent_(tool, parentId) {
  img_listForParent_(tool, parentId).forEach(function (r) { deleteRowById_("images", r.id); });
}

/** Counts Planned slots, optionally restricted to a set of tool labels. Used by Tool 54. */
function img_plannedCount_(tools) {
  var set = null;
  if (tools) { set = {}; tools.forEach(function (t) { set[t] = true; }); }
  var rows = img_all_().filter(function (r) { return r.status === "Planned"; });
  if (set) rows = rows.filter(function (r) { return set[r.tool]; });
  return rows.length;
}

/** Data for the "Image schedule" menu dialog: every slot, sorted by status. */
function imageSchedule_get() {
  var rows = img_all_();
  var order = { "Planned": 0, "Captured": 1, "On the board": 2 };
  rows.sort(function (a, b) {
    var oa = order[a.status] !== undefined ? order[a.status] : 9;
    var ob = order[b.status] !== undefined ? order[b.status] : 9;
    if (oa !== ob) return oa - ob;
    return String(a.tool || "").localeCompare(String(b.tool || ""));
  });
  return JSON.stringify({ rows: rows });
}

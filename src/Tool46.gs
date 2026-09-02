/**
 * Tool46.gs — Sizing Intake.
 * Sizing columns: id, group, num, label, hint, avatarField,
 * reading1, reading2, agreed, delta (formula — never write), enteredInCLO, note.
 */

var T46_GROUP_WHY = {
  "Heights and verticals": "Enter these first. Avatar circumferences shift when heights change, so entering in any other order means chasing your own corrections. Verify this group before you start the next.",
  "Girths": "Take each one level and firm without compressing. Tape tension is the single most common source of disagreement between partners.",
  "Widths and lengths": "These depend on landmarks rather than on the tape. Where you disagree, it is almost always because you found the shoulder point or the neck base differently.",
  "Stated rather than measured": "Read off a label or judged, not taken with a tape. Record where it came from.",
  "Added for this garment": "Points this particular item needs that the standard set does not cover."
};

var T46_META_KEYS = ["subjectType", "subjectId", "statedSize", "measurementUnit", "tolerance",
  "reader1", "reader2", "sizingDate", "sizingNotes"];

function t46_get() {
  var cfg = getConfig_();
  var meta = {};
  T46_META_KEYS.forEach(function (k) { meta[k] = cfg[k] || ""; });
  if (!meta.tolerance) meta.tolerance = "0.5";
  if (!meta.measurementUnit) meta.measurementUnit = "cm";
  if (!meta.sizingDate) {
    var d = new Date();
    meta.sizingDate = Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Beirut", "yyyy-MM-dd");
  }

  var points = readRows_("sizing");
  var groups = [];
  var byGroup = {};
  points.forEach(function (p) {
    if (!byGroup[p.group]) {
      byGroup[p.group] = { g: p.group, why: T46_GROUP_WHY[p.group] || "", pts: [] };
      groups.push(byGroup[p.group]);
    }
    byGroup[p.group].pts.push(p);
  });
  groups.forEach(function (gr) {
    gr.pts.sort(function (a, b) { return (parseFloat(a.num) || 0) - (parseFloat(b.num) || 0); });
  });

  return JSON.stringify({ meta: meta, groups: groups });
}

function t46_save(payloadJson) {
  var payload = JSON.parse(payloadJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (payload.meta) setConfig_(payload.meta);
    (payload.points || []).forEach(function (p) {
      writeRowByField_("sizing", p.id, {
        reading1: p.reading1 || "",
        reading2: p.reading2 || "",
        agreed: p.agreed || "",
        enteredInCLO: !!p.enteredInCLO,
        note: p.note || ""
      }, ["delta"]);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t46_get();
}

/** "Added for this garment" — appends a fresh Sizing row for a point the standard set doesn't cover. */
function t46_addPoint(fieldsJson) {
  var fields = JSON.parse(fieldsJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_("sizing");
    var idx = headerIndex_("sizing");
    var pts = readRows_("sizing");
    var maxNum = pts.reduce(function (m, p) { return Math.max(m, parseFloat(p.num) || 0); }, 0);
    var lastCol = sh.getLastColumn();
    var rowArr = [];
    for (var c = 0; c < lastCol; c++) rowArr.push("");
    rowArr[idx.id] = newId_(ID_PREFIX.sizing);
    rowArr[idx.group] = "Added for this garment";
    rowArr[idx.num] = Math.round((maxNum + 1) * 10) / 10;
    rowArr[idx.label] = fields.label || "";
    rowArr[idx.hint] = fields.hint || "";
    rowArr[idx.avatarField] = fields.avatarField || "";
    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([rowArr]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t46_get();
}

/** Clears every reading, keeping the points and the subject details. */
function t46_clearReadings() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    readRows_("sizing").forEach(function (p) {
      writeRowByField_("sizing", p.id,
        { reading1: "", reading2: "", agreed: "", enteredInCLO: false, note: p.note || "" },
        ["delta"]);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t46_get();
}

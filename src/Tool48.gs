/**
 * Tool48.gs — Iteration Log.
 * Iterations columns: id, order, title, session, method, referenceId, tried,
 * revealed, decision, includeInPitch, slotFront, slotBack, slotSide, slotPattern.
 * References are read live from the Research tab (same workbook) — no file
 * upload needed. Slots mirror into the Images tab like Tool 47's.
 */

var T48_SLOTS = [
  { k: "slotFront", label: "Front", req: true, type: "Screenshot" },
  { k: "slotBack", label: "Back", req: false, type: "Screenshot" },
  { k: "slotSide", label: "Side", req: false, type: "Screenshot" },
  { k: "slotPattern", label: "Pattern", req: false, type: "Pattern export" }
];

function t48_get() {
  var cfg = getConfig_();
  var lookups = getLookups_();
  var meta = { name: cfg.studentName || "", group: cfg.group || "" };

  var refs = readRows_("research").map(function (g) {
    return { id: g.id, title: [g.brand, g.product].filter(String).join(" · ") || "Untitled garment" };
  });

  var iterations = readRows_("iterations");
  iterations.sort(function (a, b) { return (parseFloat(a.order) || 0) - (parseFloat(b.order) || 0); });

  return JSON.stringify({
    meta: meta, iterations: iterations, refs: refs,
    methods: lookups.methods || ["Block manipulation", "3D-first extraction", "Both"],
    decisions: lookups.decisions || ["Pursue", "Park — may return", "Ruled out"],
    slotStates: lookups.slotStates || ["Planned", "Captured", "On the board"]
  });
}

function t48_save(payloadJson) {
  var payload = JSON.parse(payloadJson);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (payload.meta) {
      setConfig_({ studentName: payload.meta.name || "", group: payload.meta.group || "" });
    }
    (payload.deletedIds || []).forEach(function (id) {
      deleteRowById_("iterations", id);
      img_deleteForParent_("48 Iterations", id);
    });
    (payload.iterations || []).forEach(function (e, i) {
      writeRowByField_("iterations", e.id, {
        order: i + 1, title: e.title || "", session: e.session || "", method: e.method || "",
        referenceId: e.referenceId || "", tried: e.tried || "", revealed: e.revealed || "",
        decision: e.decision || "", includeInPitch: e.includeInPitch !== false,
        slotFront: e.slotFront || "Planned", slotBack: e.slotBack || "Planned",
        slotSide: e.slotSide || "Planned", slotPattern: e.slotPattern || "Planned"
      }, []);
      var name = e.title || "Untitled iteration";
      T48_SLOTS.forEach(function (s) {
        img_upsertSlot_("48 Iterations", e.id, s.label, name + " — " + s.label, s.type, e[s.k] || "Planned");
      });
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t48_get();
}

function t48_addIteration() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_("iterations");
    var idx = headerIndex_("iterations");
    var existing = readRows_("iterations");
    var maxOrder = existing.reduce(function (m, e) { return Math.max(m, parseFloat(e.order) || 0); }, 0);
    var lastCol = sh.getLastColumn();
    var row = [];
    for (var c = 0; c < lastCol; c++) row.push("");
    row[idx.id] = newId_(ID_PREFIX.iterations);
    row[idx.order] = maxOrder + 1;
    row[idx.includeInPitch] = true;
    row[idx.slotFront] = "Planned"; row[idx.slotBack] = "Planned";
    row[idx.slotSide] = "Planned"; row[idx.slotPattern] = "Planned";
    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return t48_get();
}

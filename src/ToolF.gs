/**
 * ToolF.gs — Technical Pitch Planner.
 * Pitch columns: id, round, kind, order, title, say, show, plannedSec,
 * lastActualSec, rehearsals.
 * One row per section (kind "Section", fixed structural template below) or
 * per expected question (kind "Question"). `round` is written using the
 * exact label from Lookups.pitchRounds so the column stays dropdown-safe,
 * even though the section structure itself is fixed course content, not a
 * lookup list.
 */

// The final round is a submission, not a presentation — nothing is spoken, so
// there is no pitch to plan for it. Tool I checks the submission document
// instead. Lookups.pitchRounds may still carry a third label; it goes unused.
var TF_ROUND_KEYS = ["pitch", "midterm"];
var TF_ROUND_TEMPLATE = {
  pitch: {
    label: "Session 7 · Direction pitch", budget: 180, unresolvedKey: "unresolvedPitch",
    secs: [
      { t: "Direction, in one sentence", m: 20, p: "What you are making. Not how you got there — that is next." },
      { t: "Where it came from", m: 30, p: "Which board image, and what specifically you took from it." },
      { t: "The range you explored", m: 45, p: "Show the compiled set. What you tried, and what you ruled out." },
      { t: "Why this one", m: 45, p: "The evidential argument. What it does that the others did not." },
      { t: "What I need to find out", m: 20, p: "One question you are carrying into development." }
    ]
  },
  midterm: {
    label: "Session 11 · Midterm presentation", budget: 480, unresolvedKey: "unresolvedMidterm",
    secs: [
      { t: "Direction and reference lineage", m: 60, p: "Where this came from and what it is. The room has not seen your board." },
      { t: "Iteration range and the narrowing", m: 75, p: "What you explored and why you chose this. Show the range, not just the winner." },
      { t: "The resolved item", m: 90, p: "Worn and hanger renders. Walk the room around it." },
      { t: "Fit development", m: 60, p: "What the pose testing showed and what you changed because of it." },
      { t: "Materials, hardware and finishing", m: 75, p: "Your spec. What you chose, why, and where it came from." },
      { t: "Production intention", m: 60, p: "What you plan to build, in what, and whether you have it." }
    ]
  }
};

function tF_roundLabel_(key, lookupLabels) {
  var i = TF_ROUND_KEYS.indexOf(key);
  if (lookupLabels && lookupLabels[i]) return lookupLabels[i];
  return TF_ROUND_TEMPLATE[key].label;
}

function tF_get() {
  var cfg = getConfig_();
  var lookups = getLookups_();
  var labels = lookups.pitchRounds || [];
  var rows = readRows_("pitch");

  var rounds = {};
  TF_ROUND_KEYS.forEach(function (key) {
    var tmpl = TF_ROUND_TEMPLATE[key];
    var label = tF_roundLabel_(key, labels);
    var secRows = rows.filter(function (r) { return r.round === label && r.kind === "Section"; });
    var qRows = rows.filter(function (r) { return r.round === label && r.kind === "Question"; })
      .sort(function (a, b) { return (parseFloat(a.order) || 0) - (parseFloat(b.order) || 0); });

    var secs = tmpl.secs.map(function (s, i) {
      var row = secRows.filter(function (r) { return (parseFloat(r.order) || 0) === i; })[0];
      return {
        t: s.t, p: s.p,
        say: row ? row.say || "" : "",
        show: row ? row.show || "" : "",
        plannedSec: row && row.plannedSec !== "" && row.plannedSec !== undefined ? row.plannedSec : s.m,
        lastActualSec: row ? row.lastActualSec || "" : ""
      };
    });
    var rehearsals = secRows.reduce(function (m, r) { return Math.max(m, parseFloat(r.rehearsals) || 0); }, 0);

    rounds[key] = {
      label: label, budget: tmpl.budget, secs: secs, rehearsals: rehearsals,
      unres: cfg[tmpl.unresolvedKey] || "",
      qs: qRows.map(function (r) { return { id: r.id, title: r.title || "", say: r.say || "" }; })
    };
  });

  return JSON.stringify({
    meta: { name: cfg.studentName || "", item: cfg.item || "" },
    rounds: rounds
  });
}

function tF_save(payloadJson) {
  var payload = JSON.parse(payloadJson);
  var lookups = getLookups_();
  var labels = lookups.pitchRounds || [];
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (payload.meta) {
      setConfig_({ studentName: payload.meta.name || "", item: payload.meta.item || "" });
    }
    (payload.deletedQuestionIds || []).forEach(function (id) { deleteRowById_("pitch", id); });

    TF_ROUND_KEYS.forEach(function (key) {
      var r = payload.rounds && payload.rounds[key];
      if (!r) return;
      var label = tF_roundLabel_(key, labels);
      var tmpl = TF_ROUND_TEMPLATE[key];
      var cfgPatch = {};
      cfgPatch[tmpl.unresolvedKey] = r.unres || "";
      setConfig_(cfgPatch);

      var existingSections = readRows_("pitch").filter(function (x) { return x.round === label && x.kind === "Section"; });
      (r.secs || []).forEach(function (s, i) {
        var row = existingSections.filter(function (x) { return (parseFloat(x.order) || 0) === i; })[0];
        var id = row ? row.id : newId_(ID_PREFIX.pitch);
        writeRowByField_("pitch", id, {
          round: label, kind: "Section", order: i, title: tmpl.secs[i].t,
          say: s.say || "", show: s.show || "", plannedSec: s.plannedSec || 0,
          lastActualSec: s.lastActualSec !== undefined ? s.lastActualSec : "",
          rehearsals: r.rehearsals || 0
        }, []);
      });

      (r.qs || []).forEach(function (q, i) {
        var id = q.id || newId_(ID_PREFIX.pitch);
        writeRowByField_("pitch", id, {
          round: label, kind: "Question", order: i, title: q.title || "", say: q.say || ""
        }, []);
      });
    });

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return tF_get();
}

/** Records one completed rehearsal run's per-section actual times for a round. */
function tF_recordRun(key, timesJson) {
  var times = JSON.parse(timesJson);
  var lookups = getLookups_();
  var label = tF_roundLabel_(key, lookups.pitchRounds || []);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var existingSections = readRows_("pitch").filter(function (x) { return x.round === label && x.kind === "Section"; });
    var rehearsals = existingSections.reduce(function (m, r) { return Math.max(m, parseFloat(r.rehearsals) || 0); }, 0) + 1;
    var tmpl = TF_ROUND_TEMPLATE[key];
    tmpl.secs.forEach(function (s, i) {
      var row = existingSections.filter(function (x) { return (parseFloat(x.order) || 0) === i; })[0];
      var id = row ? row.id : newId_(ID_PREFIX.pitch);
      writeRowByField_("pitch", id, {
        round: label, kind: "Section", order: i, title: s.t,
        say: row ? row.say || "" : "", show: row ? row.show || "" : "",
        plannedSec: row && row.plannedSec !== "" ? row.plannedSec : s.m,
        lastActualSec: times[i] !== undefined ? Math.round(times[i]) : (row ? row.lastActualSec || "" : ""),
        rehearsals: rehearsals
      }, []);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return tF_get();
}

/**
 * ToolI.gs — Submission Readiness.
 * Reads sibling tabs directly — no file loading. Verified checks compute
 * their own found/needed counts from Sizing, Research, Iterations,
 * Materials, Finishing, Build, Revisions and Pitch; self-declared checks
 * are the student's own say-so. Neither Readiness nor StockSlip is ever
 * written to — this tool only reads.
 *
 * Self-declared answers, the round toggle and the "what's not ready" note
 * have no dedicated columns anywhere else, so they live in Config under
 * readinessSelfChecks (a JSON blob), readinessRound and readinessGap.
 */

var TI_CHECKS = [
  { g: "Body and research", items: [
    { id: "avatar", t: "Avatar built from measurements you took", s: "Every point agreed and entered", rounds: ["mid", "fin"], auto: true },
    { id: "refs", t: "Reference garments captured", s: "Five, with composition and construction", rounds: ["mid", "fin"], auto: true },
    { id: "iters", t: "Iteration range recorded", s: "Both development methods evidenced", rounds: ["mid", "fin"], auto: true },
    { id: "explain", t: "I can explain how research became this garment", s: "Without reading from a slide", rounds: ["mid", "fin"] }
  ]},
  { g: "Fit and pattern", items: [
    { id: "poses", t: "Fit tested across all four poses", s: "A-pose, T-pose, arms forward, Pose 1", rounds: ["mid", "fin"] },
    { id: "fix", t: "Corrections made in 2D and verified", s: "Not by softening the fabric", rounds: ["mid", "fin"] },
    { id: "patcut", t: "Pattern resolved enough to cut", s: "Named pieces, seam allowance, notches, grainlines", rounds: ["mid"] },
    { id: "patprint", t: "Patterns printed at true scale and verified", s: "Test square measured", rounds: ["fin"] }
  ]},
  { g: "Specification", items: [
    { id: "dim", t: "Every material line is dimensioned", s: "", rounds: ["mid", "fin"], auto: true },
    { id: "trace", t: "Every material line traces to a reference", s: "", rounds: ["mid", "fin"], auto: true },
    { id: "secured", t: "Material secured — collected or ordered", s: "An unsigned slip is not secured", rounds: ["mid"], crit: true, auto: true },
    { id: "unavail", t: "Nothing left unavailable", s: "Anything refused needs a substitute chosen", rounds: ["mid"], crit: true, auto: true },
    { id: "slip", t: "Stock slip approved, if using stock", s: "", rounds: ["mid", "fin"], auto: true },
    { id: "tag", t: "Tag specified with a placement", s: "Required by E5", rounds: ["mid", "fin"], auto: true },
    { id: "finops", t: "Finishing operations specified", s: "Numbers, not adjectives", rounds: ["mid", "fin"], auto: true }
  ]},
  { g: "Scope and buildability", items: [
    { id: "scope", t: "Buildable in the two studio sessions", s: "Count your pieces and be honest", rounds: ["mid"], crit: true },
    { id: "ability", t: "Construction is within what I have shown I can sew", s: "Or I have a plan for the hard part", rounds: ["mid"], crit: true }
  ]},
  { g: "Construction and documentation", items: [
    { id: "conlog", t: "Construction log maintained during sewing", s: "Not reconstructed afterwards", rounds: ["fin"], auto: true },
    { id: "rate", t: "Labour rate researched, with basis and source", s: "A number without an assumption is a guess", rounds: ["fin"], auto: true },
    { id: "revlog", t: "Revision log has substantive entries", s: "Where prediction and outcome diverged", rounds: ["fin"], auto: true },
    { id: "closing", t: "Closing statement written", s: "What this taught you about trusting the simulation", rounds: ["fin"], auto: true },
    { id: "sewplan", t: "Sewing plan followable by someone else", s: "Could they build it without you", rounds: ["fin"] },
    { id: "techpack", t: "Tech pack complete", s: "Every template section populated", rounds: ["fin"] }
  ]},
  // The midterm is presented aloud, so these read tool F's plan. The final is
  // handed in, so the same group becomes a checklist over the submission
  // document — the tool cannot see inside a slide file, so every one of those
  // is the student's own say-so.
  { g: "Presentation", gFin: "Submission document", items: [
    { id: "planned", t: "Pitch planned within the time allowed", s: "", rounds: ["mid"], auto: true },
    { id: "rehearsed", t: "Rehearsed aloud", s: "Standing, at full volume, timed", rounds: ["mid"], auto: true },
    { id: "unres", t: "What is unresolved is named in the pitch", s: "The room can only help with problems you admit", rounds: ["mid"], auto: true },
    { id: "questions", t: "I have prepared for the questions I expect", s: "", rounds: ["mid"], auto: true },
    { id: "images", t: "Full image set — worn and hanger", s: "Defined angles, not just the good side", rounds: ["mid"] },
    { id: "feedback", t: "I can take feedback without defending", s: "Questions after, not during", rounds: ["mid"] },
    { id: "garment", t: "Garment finished and handed in", s: "", rounds: ["fin"] },
    { id: "docdirection", t: "Direction and reference lineage are in the document", s: "Which board images, and what you took from each", rounds: ["fin"] },
    { id: "dociter", t: "The iteration range and the narrowing are shown", s: "What you explored, not only what you chose", rounds: ["fin"] },
    { id: "docfit", t: "Fit development is documented", s: "Pose testing, the corrections, and whether they held", rounds: ["fin"] },
    { id: "docspec", t: "Materials, hardware and finishing specification included", s: "With sources and dimensions, not adjectives", rounds: ["fin"] },
    { id: "docbuild", t: "Digital to physical is documented", s: "Cutting, construction, and what the sewing plan turned out to be", rounds: ["fin"] },
    { id: "docrev", t: "Where prediction and outcome diverged is in the document", s: "Your revision log, shown rather than summarised", rounds: ["fin"] },
    { id: "docgarment", t: "The finished garment is documented", s: "Worn and hanger, at the defined angles", rounds: ["fin"] },
    { id: "docnext", t: "Closing statement and what you would do next included", s: "Not an apology. A direction.", rounds: ["fin"] },
    { id: "docorder", t: "The document reads start to finish, in order", s: "Someone who was not there can follow it without you", rounds: ["fin"] },
    { id: "docfile", t: "Exported as one slide file and checked at print size", s: "Text legible on a hard copy, images not pixelated", rounds: ["fin"] }
  ]}
];

function tI_num_(v) { var n = parseFloat(v); return isFinite(n) ? n : null; }

function tI_auto_(id, ctx) {
  switch (id) {
    case "avatar": {
      var pts = ctx.sizing;
      var ag = pts.filter(function (p) { return tI_num_(p.agreed) !== null; }).length;
      var en = pts.filter(function (p) { return !!p.enteredInCLO; }).length;
      return { ok: ag >= 18 && en >= 18, label: ag + " agreed · " + en + " entered" };
    }
    case "refs": {
      var c = ctx.research.length;
      return { ok: c >= 5, label: c + " garments" };
    }
    case "iters": {
      var e = ctx.iterations;
      var methods = ctx.lookups.methods || [];
      var b = e.some(function (x) { return x.method === methods[0] || x.method === methods[2]; });
      var d3 = e.some(function (x) { return x.method === methods[1] || x.method === methods[2]; });
      return { ok: e.length >= 5 && b && d3, label: e.length + " iterations" + ((b && d3) ? "" : ", one method only") };
    }
    case "dim": {
      var m = ctx.materials;
      var k = m.filter(function (x) { return String(x.dimension || "").trim(); }).length;
      return { ok: m.length > 0 && k === m.length, label: k + "/" + m.length };
    }
    case "trace": {
      var m2 = ctx.materials;
      var k2 = m2.filter(function (x) { return String(x.referenceId || "").trim(); }).length;
      return { ok: m2.length > 0 && k2 === m2.length, label: k2 + "/" + m2.length };
    }
    case "secured": {
      var m3 = ctx.materials;
      var k3 = m3.filter(function (x) { return x.status === "In hand" || x.status === "Requested or ordered"; }).length;
      return { ok: m3.length > 0 && k3 === m3.length, label: k3 + "/" + m3.length };
    }
    case "unavail": {
      var u = ctx.materials.filter(function (x) { return String(x.status || "").indexOf("Unavailable") === 0; }).length;
      return { ok: u === 0, label: u === 0 ? "none" : u + " line" + (u === 1 ? "" : "s") };
    }
    case "slip": {
      var st = ctx.materials.filter(function (x) { return x.source === "CSB stock"; }).length;
      if (!st) return { ok: true, label: "no stock" };
      var ok = String(ctx.config.slipApprovedOn || "").trim() !== "";
      return { ok: ok, label: ok ? "approved" : "not signed" };
    }
    case "tag": {
      var okTag = ctx.materials.some(function (x) { return x.category === "Label" && String(x.placement || "").trim(); });
      return { ok: okTag, label: okTag ? "found" : "missing" };
    }
    case "finops": {
      var f = ctx.finishing;
      return { ok: f.length >= 4, label: f.length + " ops" };
    }
    case "conlog": {
      var s = ctx.build.filter(function (x) { return x.actualOrder; }).length;
      return { ok: s >= 8, label: s + " steps logged" };
    }
    case "rate": {
      var r = tI_num_(ctx.config.labourRate) > 0 && String(ctx.config.rateBasis || "").trim() && String(ctx.config.rateSource || "").trim();
      return { ok: !!r, label: r ? (ctx.config.labourRate + " " + (ctx.config.currency || "")) : "incomplete" };
    }
    case "revlog": {
      var e2 = ctx.revisions.filter(function (x) { return String(x.why || "").trim() && String(x.action || "").trim(); });
      return { ok: e2.length >= 3, label: e2.length + " complete" };
    }
    case "closing": {
      var okC = String(ctx.config.closingStatement || "").trim().length > 40;
      return { ok: okC, label: okC ? "written" : "too short" };
    }
    case "planned": {
      var rd = ctx.pitchRound;
      if (!rd) return { ok: false, label: "not planned" };
      var budget = TF_ROUND_TEMPLATE.midterm.budget;
      var t = rd.secs.reduce(function (a, b) { return a + (parseInt(b.plannedSec) || 0); }, 0);
      return { ok: t > 0 && t <= budget, label: Math.floor(t / 60) + ":" + String(t % 60).padStart(2, "0") };
    }
    case "rehearsed": {
      var rd2 = ctx.pitchRound;
      var c2 = rd2 ? rd2.rehearsals : 0;
      return { ok: c2 >= 2, label: c2 + " run" + (c2 === 1 ? "" : "s") };
    }
    case "unres": {
      var rd3 = ctx.pitchRound;
      var okU = rd3 && String(rd3.unres || "").trim().length > 20;
      return { ok: !!okU, label: okU ? "named" : "not named" };
    }
    case "questions": {
      var rd4 = ctx.pitchRound;
      var c4 = rd4 ? rd4.qs.filter(function (q) { return String(q.title || "").trim(); }).length : 0;
      return { ok: c4 >= 3, label: c4 + " prepared" };
    }
  }
  return null;
}

function tI_context_(round) {
  var lookups = getLookups_();
  var config = getConfig_();
  // Only the midterm is presented aloud. The final is handed in as a document,
  // so there is no pitch to read for it — its presentation group is replaced
  // by the submission-document checklist, which is entirely self-declared.
  var pitchRound = { secs: [], rehearsals: 0, unres: "", qs: [] };
  if (round !== "fin") {
    var labels = lookups.pitchRounds || [];
    var label = labels[TF_ROUND_KEYS.indexOf("midterm")] || TF_ROUND_TEMPLATE.midterm.label;
    var pitchRows = readRows_("pitch");
    var secRows = pitchRows.filter(function (r) { return r.round === label && r.kind === "Section"; });
    var qRows = pitchRows.filter(function (r) { return r.round === label && r.kind === "Question"; });
    pitchRound = {
      secs: secRows.map(function (r) { return { plannedSec: r.plannedSec }; }),
      rehearsals: secRows.reduce(function (m, r) { return Math.max(m, parseFloat(r.rehearsals) || 0); }, 0),
      unres: config.unresolvedMidterm || "",
      qs: qRows.map(function (r) { return { title: r.title }; })
    };
  }
  return {
    round: round, lookups: lookups, config: config,
    sizing: readRows_("sizing"), research: readRows_("research"), iterations: readRows_("iterations"),
    materials: readRows_("materials"), finishing: readRows_("finishing"), build: readRows_("build"),
    revisions: readRows_("revisions"), pitchRound: pitchRound
  };
}

function tI_shotCount_() {
  var rows = img_all_();
  var n = 0, where = [];
  var boardCount = rows.filter(function (r) { return r.tool === "B Research" && r.status === "Planned"; }).length;
  if (boardCount) { n += boardCount; where.push(boardCount + " on the board"); }
  var compCount = rows.filter(function (r) { return r.tool === "C Iterations" && r.status === "Planned"; }).length;
  if (compCount) { n += compCount; where.push(compCount + " for the compilation"); }
  var revCount = rows.filter(function (r) { return r.tool === "H Revisions" && r.status === "Planned"; }).length;
  if (revCount) { n += revCount; where.push(revCount + " revision pairs"); }
  return { n: n, where: where };
}

function tI_get(round) {
  round = round === "fin" ? "fin" : "mid";
  var cfg = getConfig_();
  var ctx = tI_context_(round);
  var self = {};
  try { self = JSON.parse(cfg.readinessSelfChecks || "{}"); } catch (e) { self = {}; }

  var groups = TI_CHECKS.map(function (g) {
    var items = g.items.filter(function (i) { return i.rounds.indexOf(round) >= 0; }).map(function (i) {
      if (i.auto) {
        var r = tI_auto_(i.id, ctx);
        if (r) return { id: i.id, t: i.t, s: i.s, crit: !!i.crit, mode: "auto", ok: r.ok, label: r.label };
      }
      return { id: i.id, t: i.t, s: i.s, crit: !!i.crit, mode: "self", ok: !!self[i.id] };
    });
    return { g: (round === "fin" && g.gFin) ? g.gFin : g.g, items: items };
  }).filter(function (g) { return g.items.length; });

  return JSON.stringify({
    round: round,
    meta: { name: cfg.studentName || "", group: cfg.group || "", item: cfg.item || "", gap: cfg.readinessGap || "" },
    groups: groups,
    shots: tI_shotCount_()
  });
}

function tI_toggleSelf(id) {
  var cfg = getConfig_();
  var self = {};
  try { self = JSON.parse(cfg.readinessSelfChecks || "{}"); } catch (e) { self = {}; }
  self[id] = !self[id];
  setConfig_({ readinessSelfChecks: JSON.stringify(self) });
  return tI_get(cfg.readinessRound === "fin" ? "fin" : "mid");
}

function tI_setRound(round) {
  round = round === "fin" ? "fin" : "mid";
  setConfig_({ readinessRound: round });
  return tI_get(round);
}

function tI_saveMeta(payloadJson) {
  var payload = JSON.parse(payloadJson);
  setConfig_({
    studentName: payload.name || "", group: payload.group || "", item: payload.item || "",
    readinessGap: payload.gap || ""
  });
  return tI_get(payload.round === "fin" ? "fin" : "mid");
}

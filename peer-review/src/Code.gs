/**
 * Digital Design II — Tool D: Peer Review.
 *
 * Standalone from the student workbook, and deliberately so. Every other tool
 * is copied once per student; this one is shared — ten students write into one
 * spreadsheet and the instructor compiles what they wrote about each other.
 * A per-student copy would isolate exactly the data that has to pool.
 *
 * Students get the web app URL and never open the spreadsheet, which is why
 * the deployment executes as the owner rather than as the user.
 *
 * Sheet geometry lives in Sheets.gs and is the single source of truth. The
 * Responses tab carries a title block above its header row, and its Index
 * column is a sheet-owned formula filled the full height of the sheet — both
 * of which this file has to respect when it appends. See lastFilledRow_().
 */

var STUDENTS = [
  'Abou Akrouche, Ammoon',
  'Al Assemi, Mohammad Kheir',
  'Al Ghosh, Amin',
  'Al Zain, Al Hasan',
  'Hadjian, Mariebelle',
  'Hamzeh, Ali',
  'Hassan, Marwa',
  'Shehan, Abdullah',
  'Sleiman, Salem',
  'Zeineldeen, Kamal'
];

var ARCHETYPES = ['Work shirt', 'Dress shirt', 'Overshirt', 'Camp collar',
                  'Western', 'Tunic', 'Band collar', 'Something else'];

/**
 * Question sets, one per round. Session 16 has no round — the final rehearsal
 * is verbal only.
 *
 * Round labels are matched character for character by the Participation sheet's
 * COUNTIFS. Changing one silently zeroes that column. Same for KEY_QUESTION.
 */
var ROUNDS = {
  s4: {
    label: 'Session 4 — Reference board',
    fields: [
      {t: 'Direction, in my words',
       h: "One sentence. If you can't write it, that is the finding — say so.", k: 'text'},
      {t: 'Is the silhouette legible?', k: 'choice', opts: ['Yes', 'Partly', 'No'],
       follow: 'What made it readable, or what stopped it?', short: true},
      {t: 'Which archetype is this?', k: 'archetype',
       elseAsk: 'Name the archetype you see.',
       follow: 'Is it consistent, or are two archetypes mixed without a reason?', short: true},
      {t: 'Coverage gaps',
       h: 'What is missing that you would need to work from this board — a view, a detail, a material?', k: 'text'},
      {t: 'Strongest reference, and why', k: 'text'},
      {t: "One thing I'd add", k: 'text',
       h: "If a brand or a specific garment comes to mind that isn't on the board, name it. A reference you can point them to is worth more than a note that something is missing."}
    ]
  },
  s7: {
    label: 'Session 7 — Direction pitch',
    fields: [
      {t: 'Direction, in my words', h: 'One sentence, as you understood it from the pitch.', k: 'text'},
      {t: 'Does the chosen iteration deliver on the direction claimed?', k: 'choice',
       opts: ['Yes', 'Partly', 'Not yet'], follow: 'Where does it hold, and where does it slip?'},
      {t: 'Strongest element', k: 'text', short: true},
      {t: "What's unresolved that they didn't name",
       h: 'The useful part of this form. What did you see that they walked past?', k: 'text'},
      {t: "One thing I'd add", k: 'text', short: true}
    ]
  },
  s11: {
    label: 'Session 11 — Midterm presentation',
    fields: [
      {t: 'Direction, in my words', k: 'text', short: true},
      {t: 'Does the resolved item deliver on the direction pitched at Session 7?', k: 'choice',
       opts: ['Yes', 'Partly', 'Departed — explained', 'Departed — unexplained'],
       follow: 'What changed, and was the change argued or just made?'},
      {t: 'Strongest resolved element',
       h: 'Pattern, finishing, hardware, imagery — name which.', k: 'text', short: true},
      {t: 'Weakest resolved element', k: 'text', short: true},
      {t: "What's unresolved that they didn't name", k: 'text'},
      {t: "One thing I'd fix before cutting fabric", k: 'text', short: true}
    ]
  }
};

/**
 * The one question that appears exactly once in every round, so one row
 * carrying it equals one completed review. Duplicate detection, the progress
 * line, Package's "reviews received" and every Participation count all key on
 * this string. Do not change it.
 */
var KEY_QUESTION = 'Direction, in my words';

/* ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Peer Review — Digital Design II')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getConfig() {
  return {students: STUDENTS, archetypes: ARCHETYPES, rounds: ROUNDS};
}

/**
 * Appends one review as one row per question — long format, which is what
 * lets the Package sheet compile a student's feedback with plain formulas.
 *
 * payload = {round, date, presenter, reviewer, answers: [{question, response}]}
 */
function submitReview(payload) {
  if (!payload || !payload.reviewer || !payload.presenter) {
    throw new Error('Choose your name and the presenter before submitting.');
  }
  if (payload.reviewer === payload.presenter) {
    throw new Error('You cannot review your own work.');
  }
  if (STUDENTS.indexOf(payload.reviewer) < 0 || STUDENTS.indexOf(payload.presenter) < 0) {
    throw new Error('That name is not on the class list.');
  }
  var round = ROUNDS[payload.round];
  if (!round) throw new Error('Unknown review round.');
  var answers = (payload.answers || []).filter(function (a) {
    return a && String(a.response).trim() !== '';
  });
  if (answers.length < 2) {
    throw new Error('Answer at least two questions before submitting.');
  }

  // Script lock, not document lock: ten students submit through one web app
  // running as the owner, so the contention is between concurrent executions
  // of this script, not between editors of the spreadsheet.
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = responsesSheet_();

    if (hasSubmitted_(sh, round.label, payload.presenter, payload.reviewer)) {
      throw new Error('You have already submitted a review of ' + payload.presenter +
                      ' for ' + round.label + '.');
    }

    var start = lastFilledRow_(sh) + 1;
    if (start + answers.length - 1 > LAST_DATA_ROW) {
      throw new Error('The collection sheet is full. Tell your instructor — ' +
                      'nothing has been lost, but this review was not saved.');
    }

    var stamp = new Date();
    var body = answers.map(function (a) {
      return [round.label, payload.date, payload.presenter, payload.reviewer,
              a.question, a.response];
    });
    var stamps = answers.map(function () { return [stamp]; });

    // A–F then H, deliberately skipping G. The Index column is a formula the
    // sheet owns; writing anything into it — an empty string included — would
    // destroy it for that row and drop the review out of the Package sheet.
    sh.getRange(start, COL.round, body.length, 6).setValues(body);
    sh.getRange(start, COL.submitted, stamps.length, 1).setValues(stamps);
    SpreadsheetApp.flush();
    return {ok: true, count: body.length};
  } finally {
    lock.releaseLock();
  }
}

/** True if this reviewer already filed this round for this presenter. */
function hasSubmitted_(sh, roundLabel, presenter, reviewer) {
  var rows = filledRows_(sh);
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r[COL.round - 1] === roundLabel &&
        r[COL.presenting - 1] === presenter &&
        r[COL.reviewer - 1] === reviewer &&
        r[COL.question - 1] === KEY_QUESTION) {
      return true;
    }
  }
  return false;
}

/**
 * How many reviews this person has filed for a round, and who is left. Drives
 * the progress line — only answerable because the data is pooled.
 */
function remainingFor(reviewer, roundKey) {
  var round = ROUNDS[roundKey];
  if (!round || !reviewer) return {done: 0, total: 0, remaining: []};

  var done = {};
  filledRows_(responsesSheet_()).forEach(function (r) {
    if (r[COL.round - 1] === round.label &&
        r[COL.reviewer - 1] === reviewer &&
        r[COL.question - 1] === KEY_QUESTION) {
      done[r[COL.presenting - 1]] = true;
    }
  });

  var others = STUDENTS.filter(function (s) { return s !== reviewer; });
  return {
    done: others.filter(function (s) { return done[s]; }).length,
    total: others.length,
    remaining: others.filter(function (s) { return !done[s]; })
  };
}

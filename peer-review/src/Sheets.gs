/**
 * Sheets.gs — the shape of the collection spreadsheet, and the one-time build.
 *
 * The Responses tab is not a plain table. It carries a title block above the
 * header row, and its Index column is a formula filled the whole height of the
 * sheet so a new row indexes itself. Two consequences the rest of the code
 * lives with:
 *
 *   getLastRow() is useless here — the filled-down formula makes it report the
 *   bottom of the sheet whether or not a single review has been submitted.
 *   lastFilledRow_() reads column C instead.
 *
 *   The script never writes column G. Writing anything there, "" included,
 *   replaces the formula and drops that review out of the Package sheet.
 *
 * setup() builds all four tabs and is safe to run again: it repairs structure
 * and formulas without touching submitted reviews.
 */

var SHEET = 'Responses';
var HEADER_ROW = 5;
var FIRST_DATA_ROW = 6;
var LAST_DATA_ROW = 3000;
var PACKAGE_ENTRIES = 200;

var HEADERS = ['Round', 'Date', 'Presenting', 'Reviewer', 'Question', 'Response', 'Index', 'Submitted'];
var COL = {round: 1, date: 2, presenting: 3, reviewer: 4,
           question: 5, response: 6, index: 7, submitted: 8};

var INK = '#000000';
var MUTED = '#5c5c5c';
var RULE = '#9a9a9a';
var FONT = 'Helvetica Neue';

function responsesSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh) throw new Error("The '" + SHEET + "' tab is missing. Run setup() from the script editor.");
  return sh;
}

/** The last row carrying an actual submission, or HEADER_ROW if there are none. */
function lastFilledRow_(sh) {
  var vals = sh.getRange(FIRST_DATA_ROW, COL.presenting,
                         LAST_DATA_ROW - FIRST_DATA_ROW + 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return FIRST_DATA_ROW + i;
  }
  return HEADER_ROW;
}

/** Every submitted row, columns A–F. Empty array when nothing has been filed. */
function filledRows_(sh) {
  var last = lastFilledRow_(sh);
  if (last < FIRST_DATA_ROW) return [];
  return sh.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, 6).getValues();
}

/* ------------------------------------------------------------------ */

/**
 * Run once from the editor, and again any time the sheets need repairing.
 * Submitted reviews are never touched.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var did = [];
  did.push(buildResponses_(ss));
  did.push(buildParticipation_(ss));   // Package's dropdown reads its roster, so it comes first
  did.push(buildPackage_(ss));
  did.push(buildHowTo_(ss));

  // A fresh spreadsheet arrives with an empty default tab.
  var stray = ss.getSheetByName('Sheet1');
  if (stray && ss.getSheets().length > 1 && stray.getLastRow() === 0) {
    ss.deleteSheet(stray);
    did.push('removed the empty default tab');
  }

  ['How to use', SHEET, 'Package', 'Participation'].forEach(function (name, i) {
    var t = ss.getSheetByName(name);
    if (t) { ss.setActiveSheet(t); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(ss.getSheetByName('How to use'));
  SpreadsheetApp.flush();
  return did.join('\n') + '\n\nDeploy as a web app next: Execute as Me, Access Anyone with the link.';
}

/** Tab order is set once at the end of setup(), so creation order is free. */
function sheetNamed_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** A new sheet is 1000 rows by 26 columns. Grow it before writing past that. */
function ensureSize_(sh, rows, cols) {
  var haveRows = sh.getMaxRows();
  if (haveRows < rows) sh.insertRowsAfter(haveRows, rows - haveRows);
  var haveCols = sh.getMaxColumns();
  if (haveCols < cols) sh.insertColumnsAfter(haveCols, cols - haveCols);
}

function titleBlock_(sh, title, note) {
  sh.getRange('A1').setValue(title).setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue(note).setFontSize(10).setFontColor(MUTED);
}

/* ---------- Responses ---------- */

function buildResponses_(ss) {
  var sh = sheetNamed_(ss, SHEET);
  ensureSize_(sh, LAST_DATA_ROW, HEADERS.length);
  sh.getRange(1, 1, 4, HEADERS.length).clearContent();
  titleBlock_(sh, 'Collected peer reviews',
    'Written by the peer review web app, one row per question. Do not sort, reorder, or delete rows in the middle — ' +
    'the Package sheet reads this in order and the Index column in G must stay filled down.');

  sh.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setFontColor(INK)
    .setBorder(null, null, true, null, null, null, RULE, SpreadsheetApp.BorderStyle.SOLID);

  // Index: each row counts how many of the selected student's rows precede it,
  // which is what Package's MATCH walks. Written here, owned by the sheet.
  var n = LAST_DATA_ROW - FIRST_DATA_ROW + 1;
  var idx = [];
  for (var r = FIRST_DATA_ROW; r <= LAST_DATA_ROW; r++) {
    idx.push(['=IF($C' + r + '="","",IF($C' + r + '=Package!$B$3,COUNTIF($C$' +
              FIRST_DATA_ROW + ':$C' + r + ',Package!$B$3),""))']);
  }
  sh.getRange(FIRST_DATA_ROW, COL.index, n, 1).setFormulas(idx);

  sh.getRange(FIRST_DATA_ROW, COL.submitted, n, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setFrozenRows(HEADER_ROW);
  [30, 11, 24, 24, 40, 70, 8, 18].forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w * 7);
  });
  sh.getRange(1, 1, LAST_DATA_ROW, HEADERS.length).setFontFamily(FONT).setVerticalAlignment('top');
  sh.getRange(FIRST_DATA_ROW, COL.response, n, 1).setWrap(true);
  return 'Responses: header row ' + HEADER_ROW + ', data rows ' + FIRST_DATA_ROW + '–' + LAST_DATA_ROW +
         ', Index formula filled down';
}

/* ---------- Package ---------- */

function buildPackage_(ss) {
  var sh = sheetNamed_(ss, 'Package');
  ensureSize_(sh, HEADER_ROW + PACKAGE_ENTRIES, 5);
  var chosen = sh.getRange('B3').getValue();   // keep whoever is selected

  titleBlock_(sh, 'Feedback package',
    'Choose a student. Every review collected for them compiles below, ready to print or export as their package.');

  sh.getRange('A3').setValue('Student').setFontColor(MUTED).setFontSize(10);
  sh.getRange('D3').setValue('Reviews received').setFontColor(MUTED).setFontSize(10);
  sh.getRange('E3').setFormula('=IF($B$3="","",COUNTIFS(Responses!$C$' + FIRST_DATA_ROW + ':$C$' + LAST_DATA_ROW +
    ',$B$3,Responses!$E$' + FIRST_DATA_ROW + ':$E$' + LAST_DATA_ROW + ',"' + KEY_QUESTION + '"))');

  // The roster on Participation is the list, rather than a comma-separated
  // literal — every name contains a comma, so a literal list splits each one
  // into two useless half-entries.
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName('Participation').getRange('A6:A15'), true)
    .setAllowInvalid(false)
    .setHelpText('Pick a student from the class list.')
    .build();
  sh.getRange('B3').setDataValidation(rule).setFontWeight('bold');
  if (chosen) sh.getRange('B3').setValue(chosen);

  sh.getRange(HEADER_ROW, 1, 1, 5).setValues([['#', 'Round', 'Reviewer', 'Question', 'Response']])
    .setFontWeight('bold')
    .setBorder(null, null, true, null, null, null, RULE, SpreadsheetApp.BorderStyle.SOLID);

  var nums = [], rows = [];
  for (var i = 1; i <= PACKAGE_ENTRIES; i++) {
    var r = HEADER_ROW + i;
    nums.push([i]);
    rows.push(['A', 'D', 'E', 'F'].map(function (c) {
      return '=IFERROR(INDEX(Responses!$' + c + '$' + FIRST_DATA_ROW + ':$' + c + '$' + LAST_DATA_ROW +
             ',MATCH($A' + r + ',Responses!$G$' + FIRST_DATA_ROW + ':$G$' + LAST_DATA_ROW + ',0)),"")';
    }));
  }
  sh.getRange(FIRST_DATA_ROW, 1, PACKAGE_ENTRIES, 1).setValues(nums).setFontColor(MUTED);
  sh.getRange(FIRST_DATA_ROW, 2, PACKAGE_ENTRIES, 4).setFormulas(rows);

  sh.setFrozenRows(HEADER_ROW);
  [5, 30, 24, 38, 78].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 7); });
  sh.getRange(1, 1, HEADER_ROW + PACKAGE_ENTRIES, 5).setFontFamily(FONT).setVerticalAlignment('top');
  sh.getRange(FIRST_DATA_ROW, 5, PACKAGE_ENTRIES, 1).setWrap(true);
  return 'Package: ' + PACKAGE_ENTRIES + ' entries, student dropdown reading the Participation roster';
}

/* ---------- Participation ---------- */

function buildParticipation_(ss) {
  var sh = sheetNamed_(ss, 'Participation');
  ensureSize_(sh, FIRST_DATA_ROW + STUDENTS.length + 8, 6);
  titleBlock_(sh, 'Reviews given and received',
    'Reviews completed for others feed Participation & Engagement under Professional Practice. ' +
    'Counts are of whole reviews, not individual answers.');

  sh.getRange(HEADER_ROW, 1, 1, 6).setValues([['Student', 'Given\nSession 4', 'Given\nSession 7',
    'Given\nSession 11', 'Given\ntotal', 'Received\ntotal']])
    .setFontWeight('bold').setWrap(true)
    .setBorder(null, null, true, null, null, null, RULE, SpreadsheetApp.BorderStyle.SOLID);

  var A = '$A$' + FIRST_DATA_ROW + ':$A$' + LAST_DATA_ROW;
  var C = '$C$' + FIRST_DATA_ROW + ':$C$' + LAST_DATA_ROW;
  var D = '$D$' + FIRST_DATA_ROW + ':$D$' + LAST_DATA_ROW;
  var E = '$E$' + FIRST_DATA_ROW + ':$E$' + LAST_DATA_ROW;
  var order = ['s4', 's7', 's11'];

  var names = [], formulas = [];
  STUDENTS.forEach(function (name, i) {
    var r = FIRST_DATA_ROW + i;
    names.push([name]);
    var row = order.map(function (k) {
      return '=COUNTIFS(Responses!' + D + ',$A' + r + ',Responses!' + A + ',"' +
             ROUNDS[k].label + '",Responses!' + E + ',"' + KEY_QUESTION + '")';
    });
    row.push('=SUM(B' + r + ':D' + r + ')');
    row.push('=COUNTIFS(Responses!' + C + ',$A' + r + ',Responses!' + E + ',"' + KEY_QUESTION + '")');
    formulas.push(row);
  });
  sh.getRange(FIRST_DATA_ROW, 1, STUDENTS.length, 1).setValues(names);
  sh.getRange(FIRST_DATA_ROW, 2, STUDENTS.length, 5).setFormulas(formulas).setHorizontalAlignment('center');

  var noteRow = FIRST_DATA_ROW + STUDENTS.length + 2;
  sh.getRange(noteRow, 1).setValue('Notes').setFontWeight('bold');
  [['•  Counts key on the response to "' + KEY_QUESTION + '", which appears once in every round. ' +
    'One counted review equals one completed form.'],
   ["•  Reviews given is the figure that matters for Professional Practice. Reviews received tells you whether a student's work was seen by enough peers."],
   ['•  With ten students, a full round is nine reviews given and nine received per student.'],
   ['•  This column A is also the class list the Package dropdown reads. Keep it matching STUDENTS in Code.gs.']
  ].forEach(function (line, i) {
    sh.getRange(noteRow + 1 + i, 1).setValue(line[0]).setFontSize(10).setFontColor(MUTED);
  });

  sh.setColumnWidth(1, 28 * 7);
  for (var c = 2; c <= 6; c++) sh.setColumnWidth(c, 14 * 7);
  sh.getRange(1, 1, noteRow + 5, 6).setFontFamily(FONT);
  return 'Participation: ' + STUDENTS.length + ' students, three rounds';
}

/* ---------- How to use ---------- */

function buildHowTo_(ss) {
  var sh = sheetNamed_(ss, 'How to use');
  ensureSize_(sh, 30, 1);
  sh.getRange(1, 1, 30, 1).clearContent();
  sh.getRange('A1').setValue('Peer review collection').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Digital Design II · Tool D').setFontSize(10).setFontColor(MUTED);

  var lines = [
    '',
    '',
    '1.  Students open the peer review web app link. They pick their own name, the round, and who they are reviewing.',
    '2.  Submitting writes straight into Responses — one row per question. Nobody but you opens this spreadsheet.',
    '3.  The form refuses a review of yourself, and a second review of the same person for the same round.',
    '4.  On Package, choose a student. Every review collected for them compiles below, ready to print or export.',
    '5.  Participation counts how many reviews each student completed for others — the evidence for Participation & Engagement.',
    '',
    'Notes',
    '•  Do not sort or reorder Responses, and do not delete rows in the middle of it. Package reads it in order, and the Index formula in column G must stay filled down.',
    '•  Package shows up to ' + PACKAGE_ENTRIES + ' entries — a whole term for one student at nine reviewers and six questions a round.',
    '•  Capacity is rows ' + FIRST_DATA_ROW + ' to ' + LAST_DATA_ROW + ' on Responses. A submission past that is refused with a message rather than written where no formula can see it.',
    '•  The class list lives in STUDENTS in Code.gs, and is mirrored in column A of Participation. Change both together, then run setup() again.',
    '•  setup() is safe to re-run. It repairs structure and formulas and leaves submitted reviews alone.'
  ];
  lines.forEach(function (t, i) {
    if (t) sh.getRange(3 + i, 1).setValue(t).setFontSize(t === 'Notes' ? 11 : 10)
      .setFontWeight(t === 'Notes' ? 'bold' : 'normal')
      .setFontColor(t === 'Notes' ? INK : MUTED);
  });
  sh.setColumnWidth(1, 110 * 7);
  sh.getRange(1, 1, 30, 1).setFontFamily(FONT);
  return 'How to use: rewritten for the web app';
}

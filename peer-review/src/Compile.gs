/**
 * Compile.gs — turns the pooled reviews about one student into a document
 * you can hand them.
 *
 * The Package sheet does this on screen for one student at a time. This does
 * it as a Google Doc: pick a round (or the whole term), pick a student, get a
 * file in the same Drive folder as the spreadsheet, grouped by reviewer and
 * ready to print or share.
 *
 * Reads only. Nothing here writes to Responses.
 */

/** Round choices for the dialog, in course order, with a whole-term option. */
function compile_options() {
  var rounds = Object.keys(ROUNDS).map(function (k) {
    return {key: k, label: ROUNDS[k].label};
  });
  rounds.unshift({key: 'all', label: 'All rounds — the whole term'});
  return {students: STUDENTS, rounds: rounds};
}

/**
 * payload = {student, roundKey, showReviewers}
 * Returns {url, name, reviews, answers} for the dialog to link to.
 */
function compile_build(payload) {
  var student = payload && payload.student;
  var roundKey = (payload && payload.roundKey) || 'all';
  var named = !(payload && payload.showReviewers === false);

  if (!student) throw new Error('Choose a student first.');
  if (STUDENTS.indexOf(student) < 0) throw new Error('That name is not on the class list.');
  if (roundKey !== 'all' && !ROUNDS[roundKey]) throw new Error('Unknown review round.');

  var wanted = roundKey === 'all' ? null : ROUNDS[roundKey].label;
  var rows = filledRows_(responsesSheet_()).filter(function (r) {
    return r[COL.presenting - 1] === student &&
           (!wanted || r[COL.round - 1] === wanted);
  });

  if (!rows.length) {
    throw new Error('No reviews collected for ' + student +
      (wanted ? ' at ' + wanted : ' yet') + '.');
  }

  // Group by round, then by reviewer, keeping the order the sheet holds them
  // in — that is submission order, which is the only order that means anything.
  var groups = [];
  var seen = {};
  rows.forEach(function (r) {
    var round = r[COL.round - 1];
    var reviewer = r[COL.reviewer - 1];
    var key = round + ' ' + reviewer;
    if (!seen[key]) {
      seen[key] = {round: round, reviewer: reviewer, date: r[COL.date - 1], answers: []};
      groups.push(seen[key]);
    }
    seen[key].answers.push({q: r[COL.question - 1], a: r[COL.response - 1]});
  });

  var doc = compile_write_(student, wanted, groups, named);
  return {
    url: doc.getUrl(),
    name: doc.getName(),
    reviews: groups.length,
    answers: rows.length
  };
}

function compile_write_(student, roundLabel, groups, named) {
  var tz = Session.getScriptTimeZone() || 'Asia/Beirut';
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var name = 'Peer review - ' + student + ' - ' + (roundLabel || 'whole term') + ' - ' + stamp;

  var doc = DocumentApp.create(name);
  var body = doc.getBody();
  body.setMarginTop(56).setMarginBottom(56).setMarginLeft(64).setMarginRight(64);

  var title = body.appendParagraph(student);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE)
       .setFontFamily('Helvetica Neue').setForegroundColor('#000000');

  var strap = body.appendParagraph('Digital Design II - Peer review - ' +
    (roundLabel || 'all rounds') + ' - compiled ' + stamp);
  strap.setFontFamily('Helvetica Neue').setFontSize(9)
       .setForegroundColor('#5c5c5c').setSpacingAfter(4);

  var count = body.appendParagraph(groups.length + ' review' + (groups.length === 1 ? '' : 's') +
    (named ? ' from named peers' : ', anonymised'));
  count.setFontFamily('Helvetica Neue').setFontSize(9)
       .setForegroundColor('#5c5c5c').setSpacingAfter(14);

  body.appendHorizontalRule();

  var lastRound = null;
  groups.forEach(function (g, i) {
    // Only label the round when the document spans more than one.
    if (!roundLabel && g.round !== lastRound) {
      lastRound = g.round;
      var rh = body.appendParagraph(g.round);
      rh.setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setFontFamily('Helvetica Neue').setForegroundColor('#000000')
        .setSpacingBefore(20).setSpacingAfter(2);
    }

    var who = named ? g.reviewer : 'Reviewer ' + (i + 1);
    var rev = body.appendParagraph(who);
    rev.setHeading(DocumentApp.ParagraphHeading.HEADING2)
       .setFontFamily('Helvetica Neue').setFontSize(12)
       .setForegroundColor('#000000').setSpacingBefore(16).setSpacingAfter(6);

    g.answers.forEach(function (qa) {
      var q = body.appendParagraph(qa.q);
      q.setHeading(DocumentApp.ParagraphHeading.NORMAL)
       .setFontFamily('Helvetica Neue').setFontSize(9).setBold(true)
       .setForegroundColor('#5c5c5c').setSpacingBefore(9).setSpacingAfter(1);

      var a = body.appendParagraph(String(qa.a || '-'));
      a.setHeading(DocumentApp.ParagraphHeading.NORMAL)
       .setFontFamily('Helvetica Neue').setFontSize(11).setBold(false)
       .setForegroundColor('#000000').setSpacingAfter(2);
    });
  });

  var foot = body.appendParagraph(
    'Developmental and non-graded. These are your peers reading your work, not an ' +
    'assessment of it. Where several of them say the same thing, that is the signal.');
  foot.setFontFamily('Helvetica Neue').setFontSize(9)
      .setForegroundColor('#5c5c5c').setSpacingBefore(24);

  doc.saveAndClose();

  // Sit the document beside the spreadsheet rather than loose in My Drive.
  // A failed move is not worth losing the document over.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) DriveApp.getFileById(doc.getId()).moveTo(parents.next());
  } catch (e) {
    // leave it in My Drive
  }
  return doc;
}

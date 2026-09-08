/**
 * Menu.gs — the instructor's side of the tool.
 *
 * Students never open this spreadsheet; they only ever see the web app. This
 * menu is for you: compile a student's feedback into a document, rebuild the
 * sheets, and check what the deployment is actually set to.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Peer Review')
    .addItem('Compile a feedback package…', 'openCompileDialog')
    .addSeparator()
    .addItem('Rebuild the sheets', 'rebuildSheets')
    .addItem('Deployment and link', 'showDeployment')
    .addToUi();
}

function openCompileDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Compile')
    .setWidth(520).setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(html, 'Compile a feedback package');
}

/**
 * setup() is safe to re-run — it repairs structure and formulas and leaves
 * submitted reviews alone — but it is still the kind of thing worth being
 * asked about before it runs.
 */
function rebuildSheets() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert('Rebuild the sheets',
    'This rewrites the four tabs: headers, formulas, the student dropdown, ' +
    'column widths and the notes.\n\n' +
    'Submitted reviews are not touched. Nothing is deleted.\n\nGo ahead?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  ui.alert('Rebuild the sheets', setup(), ui.ButtonSet.OK);
}

/** What the class link is, and what the deployment ought to be set to. */
function showDeployment() {
  var ui = SpreadsheetApp.getUi();
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }

  var lines = [];
  if (url) {
    lines.push('The link to give the class:');
    lines.push('');
    lines.push(url);
    lines.push('');
    lines.push('Give them this, never the spreadsheet.');
  } else {
    lines.push('Not deployed yet.');
    lines.push('');
    lines.push('In the script editor: Deploy, New deployment, then pick Web app ' +
               'from the gear beside "Select type".');
  }
  lines.push('');
  lines.push('Settings it needs:');
  lines.push('  Execute as       Me');
  lines.push('  Who has access   Anyone with the link');
  lines.push('');
  lines.push('Execute-as-me is what lets students write into a sheet they cannot open.');
  lines.push('');
  lines.push('After any clasp push, the deployment keeps serving the old code until ' +
             'you edit it and set Version to New version.');

  ui.alert('Deployment and link', lines.join('\n'), ui.ButtonSet.OK);
}

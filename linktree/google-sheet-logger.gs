/**
 * School of Social link page -> Google Sheet logger.
 *
 * Setup (once):
 * 1. Create a blank Google Sheet (e.g. "School of Social Leads").
 * 2. Extensions > Apps Script. Delete what's there, paste this whole file, Save.
 * 3. Deploy > New deployment > type "Web app".
 *    Execute as: Me. Who has access: Anyone. Deploy, approve access.
 * 4. Copy the Web app URL (ends in /exec) and send it to Claude.
 *
 * Tabs are created automatically on the first request:
 *   Summary       live counts (visits, clicks, applications, qualified, unqualified, budgets)
 *   Applications  one row per person, updated as they move through the form
 *   Clicks        one row per visit or tap on the page
 */

var APP_HEADERS = [
  'ID', 'Started', 'Last update', 'Furthest step', 'Status', 'Clicked "Pick a time"',
  'Name', 'Email', 'Platform', 'Account', 'Role', 'Audience', 'Goal', 'Biggest blocker',
  'Start timeline', 'Monthly budget'
];
var CLICK_HEADERS = ['Time', 'Event', 'Visitor ID', 'Referrer', 'Device'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    setup_(ss);
    if (d.type === 'click') {
      ss.getSheetByName('Clicks').appendRow([new Date(), d.event, d.visitor || '', d.ref || '', d.device || '']);
    } else if (d.type === 'app') {
      upsertApp_(ss.getSheetByName('Applications'), d);
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function upsertApp_(sh, d) {
  var now = new Date();
  var found = sh.getRange('A:A').createTextFinder(String(d.id)).matchEntireCell(true).findNext();
  var row = found ? found.getRow() : sh.getLastRow() + 1;
  var cur = found ? sh.getRange(row, 1, 1, APP_HEADERS.length).getValues()[0] : APP_HEADERS.map(function () { return ''; });
  var f = d.fields || {};
  var next = [
    d.id,
    cur[1] || now,
    now,
    Math.max(Number(cur[3]) || 0, Number(d.step) || 0),
    d.status || cur[4] || 'Started',
    d.booked ? 'Yes' : (cur[5] || 'No'),
    pick_(f.name, cur[6]), pick_(f.email, cur[7]), pick_(f.platform, cur[8]), pick_(f.handle, cur[9]),
    pick_(f.role, cur[10]), pick_(f.audience, cur[11]), pick_(f.goal, cur[12]), pick_(f.blocker, cur[13]),
    pick_(f.timeline, cur[14]), pick_(f.budget, cur[15])
  ];
  sh.getRange(row, 1, 1, next.length).setValues([next]);
}

function pick_(v, old) { return (v === undefined || v === null || v === '') ? old : v; }

function setup_(ss) {
  if (!ss.getSheetByName('Applications')) {
    var a = ss.insertSheet('Applications');
    a.appendRow(APP_HEADERS); a.setFrozenRows(1); a.getRange(1, 1, 1, APP_HEADERS.length).setFontWeight('bold');
  }
  if (!ss.getSheetByName('Clicks')) {
    var c = ss.insertSheet('Clicks');
    c.appendRow(CLICK_HEADERS); c.setFrozenRows(1); c.getRange(1, 1, 1, CLICK_HEADERS.length).setFontWeight('bold');
  }
  if (!ss.getSheetByName('Summary')) {
    var s = ss.insertSheet('Summary', 0);
    var rows = [
      ['Metric', 'Count'],
      ['Page visits', '=COUNTIF(Clicks!B:B,"Page visit")'],
      ['Unique visitors', '=IFERROR(COUNTUNIQUE(FILTER(Clicks!C2:C,Clicks!B2:B="Page visit")),0)'],
      ['Tapped "Join the community"', '=COUNTIF(Clicks!B:B,"Community")'],
      ['Tapped "Book a consultation"', '=COUNTIF(Clicks!B:B,"Book a consultation")'],
      ['Tapped "Work directly with Arturo"', '=COUNTIF(Clicks!B:B,"Work directly with Arturo")'],
      ['Copied email', '=COUNTIF(Clicks!B:B,"Email")'],
      ['Social icon taps', '=COUNTIF(Clicks!B:B,"Instagram")+COUNTIF(Clicks!B:B,"TikTok")+COUNTIF(Clicks!B:B,"YouTube")+COUNTIF(Clicks!B:B,"Facebook")'],
      ['', ''],
      ['Applications started (step 1 done)', '=COUNTA(Applications!A2:A)'],
      ['Finished all 4 steps', '=COUNTIF(Applications!E:E,"Qualified")+COUNTIF(Applications!E:E,"Unqualified")'],
      ['Qualified', '=COUNTIF(Applications!E:E,"Qualified")'],
      ['Unqualified (sent to community)', '=COUNTIF(Applications!E:E,"Unqualified")'],
      ['Qualified who clicked "Pick a time"', '=COUNTIFS(Applications!E:E,"Qualified",Applications!F:F,"Yes")'],
      ['', ''],
      ['Budget: Under $100', '=COUNTIF(Applications!P:P,"Under $100")'],
      ['Budget: $100 to $500', '=COUNTIF(Applications!P:P,"$100 to $500")'],
      ['Budget: $500 to $1,000', '=COUNTIF(Applications!P:P,"$500 to $1,000")'],
      ['Budget: $1,000 to $2,000', '=COUNTIF(Applications!P:P,"$1,000 to $2,000")'],
      ['Budget: $2,000 to $3,000', '=COUNTIF(Applications!P:P,"$2,000 to $3,000")'],
      ['Budget: $3,000 to $5,000', '=COUNTIF(Applications!P:P,"$3,000 to $5,000")'],
      ['Budget: $5,000+', '=COUNTIF(Applications!P:P,"$5,000+")']
    ];
    s.getRange(1, 1, rows.length, 2).setValues(rows);
    s.getRange('A1:B1').setFontWeight('bold'); s.setColumnWidth(1, 300);
    var def = ss.getSheetByName('Sheet1'); if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  }
}

/** Run once from the editor to create the tabs before any traffic. */
function setupNow() { setup_(SpreadsheetApp.getActiveSpreadsheet()); }

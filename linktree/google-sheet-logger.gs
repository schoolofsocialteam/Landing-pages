/**
 * School of Social link page -> Google Sheet logger.
 *
 * Setup (once). The "School of Social CRM" sheet is already built, so no need to run setupNow.
 * 1. Open the sheet. Extensions > Apps Script. Delete what's there, paste this whole file, Save.
 * 2. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 * 3. Approve access when Google asks (Advanced > Go to project, since it's your own script).
 * 4. Copy the Web app URL (ends in /exec) and send it to Claude.
 * (setupNow only rebuilds missing tabs, e.g. on a fresh blank sheet.)
 *
 * Tabs:
 *   Dashboard     live totals, funnel and budget chart
 *   Applications  one row per person, updated as they move through the form (newest at the bottom)
 *   Clicks        one row per page visit or tap
 */

var BRAND = { charcoal: '#222222', cobalt: '#1E4EDB', highlight: '#FFF3A4', paper: '#FAF9F6', white: '#FFFFFF', rule: '#E7E4DC', muted: '#5B5D66' };
var BUDGETS = ['Under $100', '$100 to $500', '$500 to $1,000', '$1,000 to $2,000', '$2,000 to $3,000', '$3,000 to $5,000', '$5,000+'];

var APP_HEADERS = [
  'Status', 'First name', 'Last name', 'Email', 'Platform', 'Account', 'Monthly budget', 'Start timeline',
  'Role', 'Audience', 'Goal', 'Biggest blocker', 'Furthest step', 'Clicked "Pick a time"', 'Started', 'Last update', 'ID'
];
var COL = {}; APP_HEADERS.forEach(function (h, i) { COL[h] = i; });
var CLICK_HEADERS = ['Time', 'Event', 'Visitor ID', 'Device', 'Referrer'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    setup_(ss);
    if (d.type === 'click') {
      ss.getSheetByName('Clicks').appendRow([new Date(), d.event, d.visitor || '', d.device || '', d.ref || '']);
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
  var idCol = COL['ID'] + 1;
  var found = sh.getRange(1, idCol, sh.getMaxRows(), 1).createTextFinder(String(d.id)).matchEntireCell(true).findNext();
  var row = found ? found.getRow() : sh.getLastRow() + 1;
  var cur = found ? sh.getRange(row, 1, 1, APP_HEADERS.length).getValues()[0] : APP_HEADERS.map(function () { return ''; });
  var f = d.fields || {};
  var v = cur.slice();
  function set(h, val) { if (val !== undefined && val !== null && val !== '') v[COL[h]] = val; }
  set('Status', d.status || cur[COL['Status']] || 'Started');
  set('First name', f.first); set('Last name', f.last); set('Email', f.email);
  set('Platform', f.platform); set('Account', f.handle); set('Monthly budget', f.budget);
  set('Start timeline', f.timeline); set('Role', f.role); set('Audience', f.audience);
  set('Goal', f.goal); set('Biggest blocker', f.blocker);
  v[COL['Furthest step']] = Math.max(Number(cur[COL['Furthest step']]) || 0, Number(d.step) || 0) + ' of 4';
  v[COL['Clicked "Pick a time"']] = d.booked ? 'Yes' : (cur[COL['Clicked "Pick a time"']] || 'No');
  v[COL['Started']] = cur[COL['Started']] || now;
  v[COL['Last update']] = now;
  v[COL['ID']] = d.id;
  sh.getRange(row, 1, 1, v.length).setValues([v]);
}

function header_(sh, headers, widths) {
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(BRAND.charcoal).setFontColor(BRAND.white).setFontWeight('bold')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  if (sh.getMaxColumns() > headers.length) sh.deleteColumns(headers.length + 1, sh.getMaxColumns() - headers.length);
  sh.getRange(2, 1, sh.getMaxRows() - 1, headers.length).setVerticalAlignment('middle').setFontFamily('Inter');
  sh.getRange(1, 1, 1, headers.length).setFontFamily('Inter');
  sh.setHiddenGridlines(true);
}

function setup_(ss) {
  if (!ss.getSheetByName('Applications')) {
    var a = ss.insertSheet('Applications');
    header_(a, APP_HEADERS, [110, 120, 120, 220, 120, 160, 140, 130, 180, 110, 220, 320, 110, 140, 150, 150, 160]);
    a.getRange('A:A').setFontWeight('bold').setHorizontalAlignment('center');
    a.getRange('O:P').setNumberFormat('mmm d, yyyy h:mm am/pm');
    a.getRange('L:L').setWrap(true);
    a.getRange(2, 1, a.getMaxRows() - 1, APP_HEADERS.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
      .setFirstRowColor(BRAND.white).setSecondRowColor(BRAND.paper);
    var status = a.getRange('A2:A');
    a.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Qualified').setBackground(BRAND.cobalt).setFontColor(BRAND.white).setRanges([status]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Unqualified').setBackground(BRAND.highlight).setFontColor(BRAND.charcoal).setRanges([status]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Started').setBackground(BRAND.rule).setFontColor(BRAND.muted).setRanges([status]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setFontColor(BRAND.cobalt).setBold(true).setRanges([a.getRange('N2:N')]).build()
    ]);
    a.getRange('A1').setNote('Started = finished step 1 only. Qualified = $2,000+/mo and not "Just exploring". Unqualified = sent to the community.');
  }
  if (!ss.getSheetByName('Clicks')) {
    var c = ss.insertSheet('Clicks');
    header_(c, CLICK_HEADERS, [170, 220, 120, 100, 300]);
    c.getRange('A:A').setNumberFormat('mmm d, yyyy h:mm am/pm');
    c.getRange(2, 1, c.getMaxRows() - 1, CLICK_HEADERS.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
      .setFirstRowColor(BRAND.white).setSecondRowColor(BRAND.paper);
  }
  if (!ss.getSheetByName('Dashboard')) buildDashboard_(ss);
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
}

function buildDashboard_(ss) {
  var s = ss.insertSheet('Dashboard', 0);
  s.setHiddenGridlines(true);
  s.getRange('A1:H60').setFontFamily('Inter').setBackground(BRAND.paper).setVerticalAlignment('middle');
  s.setColumnWidth(1, 24); s.setColumnWidth(2, 300); s.setColumnWidth(3, 110); s.setColumnWidth(4, 24);
  [5, 6, 7, 8].forEach(function (c) { s.setColumnWidth(c, 110); });

  s.getRange('B2').setValue('School of Social / Link page').setFontSize(20).setFontWeight('bold').setFontColor(BRAND.charcoal);
  s.getRange('B3').setValue('Live totals. Updates on its own as people visit and apply.').setFontColor(BRAND.muted);

  function section(row, title) {
    s.getRange(row, 2, 1, 2).merge().setValue(title.toUpperCase())
      .setFontColor(BRAND.cobalt).setFontWeight('bold').setFontSize(10);
  }
  function rows(start, list) {
    var r = s.getRange(start, 2, list.length, 2);
    r.setValues(list).setBackground(BRAND.white);
    s.getRange(start, 3, list.length, 1).setFontWeight('bold').setHorizontalAlignment('right').setFontSize(13);
    for (var i = 0; i < list.length; i++) s.setRowHeight(start + i, 30);
    r.setBorder(null, null, true, null, null, true, BRAND.rule, SpreadsheetApp.BorderStyle.SOLID);
  }

  section(5, 'Traffic');
  rows(6, [
    ['Page visits', '=COUNTIF(Clicks!B:B,"Page visit")'],
    ['Unique visitors', '=IFERROR(COUNTUNIQUE(FILTER(Clicks!C2:C,Clicks!B2:B="Page visit")),0)'],
    ['Tapped "Join the community"', '=COUNTIF(Clicks!B:B,"Community")'],
    ['Tapped "Book a consultation"', '=COUNTIF(Clicks!B:B,"Book a consultation")'],
    ['Tapped "Work directly with Arturo"', '=COUNTIF(Clicks!B:B,"Work directly with Arturo")'],
    ['Social icon taps', '=COUNTIF(Clicks!B:B,"Instagram")+COUNTIF(Clicks!B:B,"TikTok")+COUNTIF(Clicks!B:B,"YouTube")+COUNTIF(Clicks!B:B,"Facebook")'],
    ['Copied email', '=COUNTIF(Clicks!B:B,"Email")']
  ]);

  section(14, '1:1 application funnel');
  rows(15, [
    ['Opened the application', '=COUNTIF(Clicks!B:B,"Work directly with Arturo")'],
    ['Finished step 1 (contact info)', '=COUNTA(Applications!Q2:Q)'],
    ['Finished all 4 steps', '=COUNTIF(Applications!A:A,"Qualified")+COUNTIF(Applications!A:A,"Unqualified")'],
    ['Qualified', '=COUNTIF(Applications!A:A,"Qualified")'],
    ['Unqualified (sent to community)', '=COUNTIF(Applications!A:A,"Unqualified")'],
    ['Qualified who clicked "Pick a time"', '=COUNTIFS(Applications!A:A,"Qualified",Applications!N:N,"Yes")'],
    ['Qualified rate', '=IFERROR(C18/C17,0)']
  ]);
  s.getRange('C21').setNumberFormat('0%');
  s.getRange('B18:C18').setBackground(BRAND.cobalt).setFontColor(BRAND.white);
  s.getRange('B19:C19').setBackground(BRAND.highlight);

  section(23, 'Monthly budget picked');
  rows(24, BUDGETS.map(function (b) { return [b, '=COUNTIF(Applications!G:G,"' + b + '")']; }));

  var chart = s.newChart().asBarChart()
    .addRange(s.getRange('B24:C30'))
    .setPosition(5, 5, 0, 0)
    .setOption('title', 'How much people are ready to invest per month')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [BRAND.cobalt])
    .setOption('backgroundColor', BRAND.paper)
    .setOption('width', 460).setOption('height', 300)
    .build();
  s.insertChart(chart);
}

/** Run once from the editor to build the tabs before any traffic. */
function setupNow() { setup_(SpreadsheetApp.getActiveSpreadsheet()); }

/** Optional: rebuild the dashboard if you change it by accident. */
function rebuildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Dashboard'); if (s) ss.deleteSheet(s);
  buildDashboard_(ss);
}

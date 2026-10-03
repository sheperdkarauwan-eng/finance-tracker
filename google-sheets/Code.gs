/**
 * Finance tracker -> Google Sheets sync + dashboard.
 *
 * Setup (once):
 *   1. Create a blank Google Sheet, then Extensions > Apps Script.
 *   2. Replace everything in Code.gs with this file and save.
 *   3. Back in the sheet, reload. Use the "Finance tracker" menu > "Build dashboard".
 *      Approve the permissions prompt. It shows your sync key when done.
 *   4. In Apps Script: Deploy > New deployment > Web app.
 *      Execute as: Me. Who has access: Anyone. Copy the Web app URL.
 *   5. In the app: Settings > Google Sheets sync, paste the URL and the key.
 */

var THEME = {
  bg: '#0f1424', panel: '#161c31', grid: '#2a3150', text: '#e6e9f2', muted: '#8d96b3',
  green: '#3ddc97', pink: '#ff5c8a', purple: '#9b8cff', gold: '#f5c542', blue: '#4f9dff', orange: '#ff9f43'
};
var MONEY = '€#,##0.00;-€#,##0.00;€0.00';

// Raw data tabs written by the app on every sync. Column 0 of any tab listed in DATE_COLS is a date.
var TABS = {
  Transactions: ['Date', 'Type', 'Category', 'Amount', 'Account', 'Note', 'Source', 'Needs review', 'Id'],
  Budgets: ['Category', 'Monthly budget'],
  Expected: ['Description', 'Amount'],
  Recurring: ['Name', 'Type', 'Cycle', 'Amount', 'Monthly cost', 'Billing day', 'Account', 'Note'],
  Savings: ['Date', 'Account', 'Amount', 'Note'],
  Loans: ['Date', 'Person', 'Type', 'Amount', 'Outstanding change', 'Account', 'Note'],
  Tours: ['Date', 'Tour', 'Type', 'Category', 'Amount', 'Owed change', 'Account', 'Note']
};
var DATE_COLS = { Transactions: true, Savings: true, Loans: true, Tours: true };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Finance tracker')
    .addItem('Build dashboard', 'setup')
    .addItem('Show sync key', 'showKey')
    .addToUi();
}

function showKey() {
  SpreadsheetApp.getUi().alert('Sync key', getOrCreateKey_(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function getOrCreateKey_() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('SYNC_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('SYNC_KEY', key);
  }
  return key;
}

// ---------- Web app endpoint ----------

function doGet() {
  return json_({ ok: true, message: 'Finance tracker sync is running. The app sends data with POST.' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'busy, try again' });
  try {
    var data = JSON.parse(e.postData.contents);
    if (!data.key || data.key !== PropertiesService.getScriptProperties().getProperty('SYNC_KEY')) {
      return json_({ ok: false, error: 'wrong sync key' });
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName('Dashboard')) setupSheets_(ss);
    writeTab_(ss, 'Transactions', data.transactions);
    writeTab_(ss, 'Budgets', data.budgets);
    writeTab_(ss, 'Expected', data.expected);
    writeTab_(ss, 'Recurring', data.recurring);
    writeTab_(ss, 'Savings', data.savings);
    writeTab_(ss, 'Loans', data.loans);
    writeTab_(ss, 'Tours', data.tours);
    var settings = ss.getSheetByName('Settings');
    settings.getRange('B1').setValue(Number(data.salary) || 0);
    settings.getRange('B2').setValue(new Date());
    return json_({ ok: true, rows: (data.transactions || []).length });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function toDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : s;
}

function writeTab_(ss, name, rows) {
  var sh = ss.getSheetByName(name);
  var width = TABS[name].length;
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  rows = (rows || []).map(function (r) {
    var row = r.slice(0, width);
    while (row.length < width) row.push('');
    if (DATE_COLS[name]) row[0] = toDate_(row[0]);
    return row;
  });
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
}

// ---------- Building the workbook ----------

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets_(ss);
  var key = getOrCreateKey_();
  ss.getSheetByName('Dashboard').activate();
  try {
    SpreadsheetApp.getUi().alert('Dashboard built',
      'Your sync key:\n\n' + key + '\n\nNext: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone), then paste the URL and this key into the app under Settings > Google Sheets sync.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('Sync key: ' + key);
  }
}

function resetSheet_(ss, name) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();
  return sh;
}

function setupSheets_(ss) {
  ss.setSpreadsheetLocale('en_GB');
  ss.setSpreadsheetTimeZone('Europe/Rome');
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var headers = TABS[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(THEME.panel).setFontColor(THEME.text);
    sh.setFrozenRows(1);
    if (DATE_COLS[name]) sh.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  });
  ss.getSheetByName('Transactions').getRange('D2:D').setNumberFormat(MONEY);
  ss.getSheetByName('Budgets').getRange('B2:B').setNumberFormat(MONEY);
  ss.getSheetByName('Expected').getRange('B2:B').setNumberFormat(MONEY);
  ss.getSheetByName('Recurring').getRange('D2:E').setNumberFormat(MONEY);
  ss.getSheetByName('Savings').getRange('C2:C').setNumberFormat(MONEY);
  ss.getSheetByName('Loans').getRange('D2:E').setNumberFormat(MONEY);
  ss.getSheetByName('Tours').getRange('E2:F').setNumberFormat(MONEY);

  var settings = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  settings.getRange('A1:A2').setValues([['Salary'], ['Last synced']]).setFontWeight('bold');
  settings.getRange('B1').setNumberFormat(MONEY);
  settings.getRange('B2').setNumberFormat('yyyy-mm-dd hh:mm');

  var dash = resetSheet_(ss, 'Dashboard');
  var calc = resetSheet_(ss, 'Calc');
  buildCalc_(calc);
  buildDashboard_(ss, dash, calc);

  ['Sheet1', 'Foglio1'].forEach(function (n) {
    var blank = ss.getSheetByName(n);
    if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  });
  ss.setActiveSheet(ss.getSheetByName('Dashboard'));
  ss.moveActiveSheet(1);
}

// Helper tab feeding the dashboard and charts. Hidden after building.
function buildCalc_(c) {

  c.getRange('A1:B2').setValues([
    ['Month start', '=DATE(YEAR(Dashboard!C3),MONTH(Dashboard!C3),1)'],
    ['Month end', '=EOMONTH(B1,0)']
  ]);
  c.getRange('B1:B2').setNumberFormat('yyyy-mm-dd');

  // Category breakdown: rows 5..14 (header + 8 categories + Uncategorized)
  c.getRange('A4:C4').setValues([['Category', 'Budget', 'Actual']]);
  var rows = [];
  for (var i = 0; i < 9; i++) {
    var r = 5 + i;
    rows.push([
      '=IF(Budgets!A' + (i + 2) + '="","",Budgets!A' + (i + 2) + ')',
      '=IF(A' + r + '="","",Budgets!B' + (i + 2) + ')',
      '=IF(A' + r + '="","",SUMIFS(Transactions!$D:$D,Transactions!$B:$B,"expense",Transactions!$C:$C,A' + r + ',Transactions!$A:$A,">="&$B$1,Transactions!$A:$A,"<="&$B$2))'
    ]);
  }
  c.getRange(5, 1, 9, 3).setFormulas(rows);
  c.getRange('E4:F4').setValues([['Category', 'Actual']]);
  var pie = [];
  for (var p = 0; p < 9; p++) pie.push(['=A' + (5 + p), '=C' + (5 + p)]);
  c.getRange(5, 5, 9, 2).setFormulas(pie);

  // Daily spending: rows 17..47
  c.getRange('A16:C16').setValues([['Day', 'Spent', 'Daily budget']]);
  var daily = [];
  for (var d = 0; d < 31; d++) {
    var dr = 17 + d;
    daily.push([
      d === 0 ? '=B1' : '=IF(A' + (dr - 1) + '="","",IF(A' + (dr - 1) + '+1<=$B$2,A' + (dr - 1) + '+1,""))',
      '=IF(A' + dr + '="","",SUMIFS(Transactions!$D:$D,Transactions!$B:$B,"expense",Transactions!$A:$A,A' + dr + '))',
      '=IF(A' + dr + '="","",SUM(Budgets!$B:$B)/DAY($B$2))'
    ]);
  }
  c.getRange(17, 1, 31, 3).setFormulas(daily);
  c.getRange('A17:A47').setNumberFormat('d mmm');

  // Last 6 months: rows 51..56
  c.getRange('A50:C50').setValues([['Month', 'Income', 'Expenses']]);
  var months = [];
  for (var m = 0; m < 6; m++) {
    var mr = 51 + m;
    months.push([
      '=EDATE($B$1,' + (m - 5) + ')',
      '=SUMIFS(Transactions!$D:$D,Transactions!$B:$B,"income",Transactions!$A:$A,">="&A' + mr + ',Transactions!$A:$A,"<="&EOMONTH(A' + mr + ',0))',
      '=SUMIFS(Transactions!$D:$D,Transactions!$B:$B,"expense",Transactions!$A:$A,">="&A' + mr + ',Transactions!$A:$A,"<="&EOMONTH(A' + mr + ',0))'
    ]);
  }
  c.getRange(51, 1, 6, 3).setFormulas(months);
  c.getRange('A51:A56').setNumberFormat('mmm yyyy');
  c.getRange('B5:C13').setNumberFormat(MONEY);
  c.getRange('B17:C47').setNumberFormat(MONEY);
  c.getRange('B51:C56').setNumberFormat(MONEY);
  c.hideSheet();
}

function buildDashboard_(ss, d, calc) {
  d.setHiddenGridlines(true);
  d.getRange('A1:N90').setBackground(THEME.bg).setFontColor(THEME.text).setFontFamily('Arial');
  d.setColumnWidth(1, 16);
  for (var col = 2; col <= 13; col++) d.setColumnWidth(col, 110);
  d.setColumnWidth(14, 16);

  // Title + month picker
  d.getRange('B1:M1').merge().setFormula('=UPPER(TEXT(Calc!B1,"mmmm yyyy"))')
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center');
  d.setRowHeight(1, 46);
  d.getRange('B3').setValue('Month').setFontColor(THEME.muted);
  d.getRange('C3').setFormula('=EOMONTH(TODAY(),-1)+1').setNumberFormat('mmmm yyyy')
    .setBackground(THEME.panel).setFontWeight('bold');
  d.getRange('D3:H3').merge().setValue('Type any date in that month to switch. Data refreshes on every sync from the app.')
    .setFontColor(THEME.muted).setFontSize(9);
  d.getRange('J3').setValue('Last synced').setFontColor(THEME.muted);
  d.getRange('K3:L3').merge().setFormula('=IF(Settings!B2="","never",Settings!B2)').setNumberFormat('d mmm, hh:mm');

  // KPI boxes (row 5 label, row 6 value), two columns each
  var kpis = [
    ['SALARY', '=Settings!B1', THEME.blue],
    ['BUDGETED + EXPECTED', '=SUM(Budgets!B:B)+SUM(Expected!B:B)', THEME.pink],
    ['LEFT TO ALLOCATE', '=Settings!B1-SUM(Budgets!B:B)-SUM(Expected!B:B)', THEME.green],
    ['INCOME', '=SUMIFS(Transactions!D:D,Transactions!B:B,"income",Transactions!A:A,">="&Calc!B1,Transactions!A:A,"<="&Calc!B2)', THEME.purple],
    ['TOTAL SPENT', '=SUMIFS(Transactions!D:D,Transactions!B:B,"expense",Transactions!A:A,">="&Calc!B1,Transactions!A:A,"<="&Calc!B2)', THEME.gold],
    ['NET', '=H6-J6', THEME.orange]
  ];
  kpis.forEach(function (k, i) {
    var c0 = 2 + i * 2;
    var label = d.getRange(5, c0, 1, 2).merge().setValue(k[0]);
    var value = d.getRange(6, c0, 1, 2).merge();
    value.setFormula(k[1]);
    label.setFontColor(k[2]).setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center').setBackground(THEME.panel);
    value.setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center').setBackground(THEME.panel).setNumberFormat(MONEY);
    d.getRange(5, c0, 2, 2).setBorder(true, true, true, true, false, false, k[2], SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  d.setRowHeight(6, 40);

  // Section: EXPENSES (B..E), rows 9..
  sectionHeader_(d, 'B9:E9', 'EXPENSES', THEME.pink);
  tableHeader_(d, 'B10:E10', ['Category', 'Budget', 'Actual', 'Left'], THEME.pink);
  var exp = [];
  for (var i = 0; i < 9; i++) {
    var cr = 5 + i, r = 11 + i;
    exp.push(['=Calc!A' + cr, '=Calc!B' + cr, '=Calc!C' + cr, '=IF(B' + r + '="","",C' + r + '-D' + r + ')']);
  }
  d.getRange(11, 2, 9, 4).setFormulas(exp);
  d.getRange('B20:E20').setFormulas([['Total', '=SUM(C11:C19)', '=SUM(D11:D19)', '=C20-D20']]).setFontWeight('bold');
  d.getRange('C11:E20').setNumberFormat(MONEY);
  box_(d, 'B10:E20', THEME.pink);

  // Section: INCOME (G..H)
  sectionHeader_(d, 'G9:H9', 'INCOME', THEME.green);
  d.getRange('G10').setFormula(
    '=IFERROR(QUERY(Transactions!A:G,"select G, sum(D) where B = \'income\' and A >= date \'"&TEXT(Calc!B1,"yyyy-mm-dd")&"\' and A <= date \'"&TEXT(Calc!B2,"yyyy-mm-dd")&"\' group by G label G \'Source\', sum(D) \'Actual\'",1),{"Source","Actual"})');
  tableHeader_(d, 'G10:H10', null, THEME.green);
  d.getRange('H11:H20').setNumberFormat(MONEY);
  box_(d, 'G10:H20', THEME.green);

  // Section: SUBSCRIPTIONS & UTILITIES (J..M)
  sectionHeader_(d, 'J9:M9', 'SUBSCRIPTIONS & UTILITIES', THEME.purple);
  d.getRange('J10').setFormula(
    '=IFERROR(QUERY(Recurring!A:F,"select A, B, F, E where A is not null order by E desc label A \'Name\', B \'Type\', F \'Billing day\', E \'Monthly\'",1),{"Name","Type","Billing day","Monthly"})');
  tableHeader_(d, 'J10:M10', null, THEME.purple);
  d.getRange('M11:M19').setNumberFormat(MONEY);
  d.getRange('J20:M20').setValues([['Total', '', '', '']]);
  d.getRange('M20').setFormula('=SUM(Recurring!E:E)').setNumberFormat(MONEY);
  d.getRange('J20:M20').setFontWeight('bold');
  box_(d, 'J10:M20', THEME.purple);

  // Section: LOANS OWED TO YOU (B..E)
  sectionHeader_(d, 'B23:E23', 'LOANS (OWED TO YOU)', THEME.gold);
  d.getRange('B24').setFormula(
    '=IFERROR(QUERY(Loans!A:E,"select B, sum(E) where B is not null group by B order by sum(E) desc label B \'Person\', sum(E) \'Outstanding\'",1),{"Person","Outstanding"})');
  tableHeader_(d, 'B24:E24', null, THEME.gold);
  d.getRange('C25:C33').setNumberFormat(MONEY);
  d.getRange('B34:C34').setValues([['Total', '']]);
  d.getRange('C34').setFormula('=SUM(Loans!E:E)').setNumberFormat(MONEY);
  d.getRange('B34:C34').setFontWeight('bold');
  box_(d, 'B24:E34', THEME.gold);

  // Section: SAVINGS (G..H)
  sectionHeader_(d, 'G23:H23', 'SAVINGS', THEME.blue);
  d.getRange('G24').setFormula(
    '=IFERROR(QUERY(Savings!A:C,"select B, sum(C) where B is not null group by B label B \'Account\', sum(C) \'Total\'",1),{"Account","Total"})');
  tableHeader_(d, 'G24:H24', null, THEME.blue);
  d.getRange('H25:H33').setNumberFormat(MONEY);
  d.getRange('G34:H34').setValues([['Total', '']]);
  d.getRange('H34').setFormula('=SUM(Savings!C:C)').setNumberFormat(MONEY);
  d.getRange('G34:H34').setFontWeight('bold');
  box_(d, 'G24:H34', THEME.blue);

  // Section: TOURS (J..M)
  sectionHeader_(d, 'J23:M23', 'TOURS (STILL OWED)', THEME.orange);
  d.getRange('J24').setFormula(
    '=IFERROR(QUERY(Tours!A:F,"select B, sum(F) where B is not null group by B order by sum(F) desc label B \'Tour\', sum(F) \'Still owed\'",1),{"Tour","Still owed"})');
  tableHeader_(d, 'J24:M24', null, THEME.orange);
  d.getRange('K25:K33').setNumberFormat(MONEY);
  d.getRange('J34:K34').setValues([['Total', '']]);
  d.getRange('K34').setFormula('=SUM(Tours!F:F)').setNumberFormat(MONEY);
  d.getRange('J34:K34').setFontWeight('bold');
  box_(d, 'J24:M34', THEME.orange);

  // Over-budget "Left" cells turn pink
  var rule = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
    .setFontColor(THEME.pink).setRanges([d.getRange('E11:E20'), d.getRange('F6'), d.getRange('L6')]).build();
  d.setConditionalFormatRules([rule]);

  buildCharts_(d, calc);
}

function sectionHeader_(sh, a1, text, color) {
  sh.getRange(a1).merge().setValue(text).setFontColor(color).setFontSize(14).setFontWeight('bold');
}

function tableHeader_(sh, a1, labels, color) {
  var r = sh.getRange(a1);
  if (labels) r.setValues([labels]);
  r.setFontWeight('bold').setFontColor(color).setBackground(THEME.panel).setFontSize(9);
}

function box_(sh, a1, color) {
  sh.getRange(a1).setBorder(true, true, true, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID);
}

function chartStyle_(builder, title) {
  return builder
    .setOption('title', title)
    .setOption('backgroundColor', THEME.bg)
    .setOption('titleTextStyle', { color: THEME.text, fontSize: 13, bold: true })
    .setOption('legend', { position: 'bottom', textStyle: { color: THEME.muted } })
    .setOption('hAxis', { textStyle: { color: THEME.muted }, gridlines: { color: THEME.grid } })
    .setOption('vAxis', { textStyle: { color: THEME.muted }, gridlines: { color: THEME.grid }, format: '€#,##0' })
    .setOption('width', 440)
    .setOption('height', 280);
}

function buildCharts_(d, calc) {

  d.getRange('B37:M37').merge().setValue('CHARTS').setFontSize(14).setFontWeight('bold').setFontColor(THEME.text);

  var budgetVsActual = chartStyle_(d.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(calc.getRange('A4:C13')).setNumHeaders(1), 'Budget vs actual')
    .setOption('colors', [THEME.green, THEME.pink])
    .setPosition(39, 2, 0, 0).build();
  d.insertChart(budgetVsActual);

  var breakdown = chartStyle_(d.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(calc.getRange('E4:F13')).setNumHeaders(1), 'Spending by category')
    .setOption('pieHole', 0.5)
    .setOption('pieSliceBorderColor', THEME.bg)
    .setOption('legend', { position: 'right', textStyle: { color: THEME.muted } })
    .setOption('colors', [THEME.green, THEME.blue, THEME.purple, THEME.gold, THEME.pink, THEME.orange, '#38bdf8', '#c084fc', '#8d96b3'])
    .setPosition(39, 8, 0, 0).build();
  d.insertChart(breakdown);

  var dailyChart = chartStyle_(d.newChart().setChartType(Charts.ChartType.AREA)
    .addRange(calc.getRange('A16:C47')).setNumHeaders(1), 'Daily spending vs daily budget')
    .setOption('colors', [THEME.pink, THEME.green])
    .setOption('areaOpacity', 0.2)
    .setOption('width', 900)
    .setPosition(55, 2, 0, 0).build();
  d.insertChart(dailyChart);

  var sixMonths = chartStyle_(d.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(calc.getRange('A50:C56')).setNumHeaders(1), 'Income vs expenses (6 months)')
    .setOption('colors', [THEME.green, THEME.pink])
    .setOption('width', 900)
    .setPosition(71, 2, 0, 0).build();
  d.insertChart(sixMonths);
}

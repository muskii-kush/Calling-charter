/**
 * Calling Cost & Service Dashboard — Google Apps Script backend
 * Bind this project to a Google Sheets copy of the source workbook.
 * Required tabs: Sheet1, Sheet3, Team wise, vendor wise.
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.dashboardData = safeJson_(getDashboardData());
  return template.evaluate()
    .setTitle('Calling Cost & Service Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Bind this Apps Script project to the Google Sheets copy of the calling workbook.');
  var costsSheet = requireSheet_(ss, 'Sheet3');
  var usesSheet = requireSheet_(ss, 'Sheet1');
  var teamsSheet = requireSheet_(ss, 'Team wise');
  var vendorsSheet = requireSheet_(ss, 'vendor wise');

  // Sheet3: vendor license/channel allocation snapshot; first three rows contain rate cards and headers.
  var costValues = costsSheet.getRange('A4:M18').getValues();
  var costs = costValues.filter(function(r) { return r[0] !== '' && r[0] != null; }).map(function(r) {
    return {
      bu: String(r[0] || ''),
      dialer: String(r[1] || ''),
      function: String(r[3] || ''),
      month: r[4] instanceof Date ? Utilities.formatDate(r[4], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(r[4] || ''),
      agents: num_(r[5]), supervisors: num_(r[6]), admins: num_(r[7]),
      primary: num_(r[8]), secondary: num_(r[9]), mobile: num_(r[10]), dids: num_(r[11]),
      total: num_(r[12])
    };
  });
  var ozoneRates = costsSheet.getRange('F1:L1').getValues()[0].map(num_);
  var exotelRates = costsSheet.getRange('F2:L2').getValues()[0].map(num_);

  // Sheet1: team/use-case inventory. POCs are intentionally not sent to the browser.
  var useValues = usesSheet.getRange('A4:P65').getValues();
  var uses = useValues.filter(function(r) { return r[0] !== '' && r[0] != null; }).map(function(r) {
    return {
      id: r[0], market: String(r[1] || ''), subBU: String(r[2] || ''),
      team: String(r[3] || ''), subTeam: String(r[4] || ''),
      onlyCalling: String(r[5] || ''), agents: num_(r[6]),
      centralScope: String(r[7] || ''), autoLMS: String(r[8] || ''),
      type: String(r[9] || ''), dialer: String(r[10] || ''),
      mechanism: String(r[11] || ''), provider: String(r[12] || ''),
      licenseCost: String(r[14] || ''), monthly: isNum_(r[15]) ? Number(r[15]) : null
    };
  });
  var teamList = teamsSheet.getRange('B2:B15').getValues().map(function(r) { return r[0]; }).filter(Boolean).map(String);
  var vendors = vendorsSheet.getRange('A2:A9').getValues().map(function(r) { return r[0]; }).filter(Boolean).map(String);
  return {
    source: ss.getName(), snapshotDate: '5 Sep 2026', costs: costs, uses: uses,
    teamList: teamList, vendors: vendors,
    rates: { Exotel: exotelRates, Ozonetel: ozoneRates }
  };
}

function requireSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing required worksheet tab: ' + name);
  return sheet;
}
function num_(value) {
  var n = Number(value);
  return isFinite(n) ? n : 0;
}
function isNum_(value) { return typeof value === 'number' && isFinite(value); }
function safeJson_(value) {
  return JSON.stringify(value).replace(/[<>&\u2028\u2029]/g, function(c) {
    return {'<':'\\u003c','>':'\\u003e','&':'\\u0026','\u2028':'\\u2028','\u2029':'\\u2029'}[c];
  });
}

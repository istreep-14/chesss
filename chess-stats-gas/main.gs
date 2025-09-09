function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Chess Stats')
    .addItem('Setup Games Sheet', 'setupGamesSheet')
    .addSeparator()
    .addItem('Fetch Month…', 'promptFetchMonth')
    .addItem('Fetch Range…', 'promptFetchRange')
    .addToUi();
}

function setupGamesSheet() {
  getOrCreateGamesSheet_();
}

function promptFetchMonth() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Fetch Month', 'Enter: username, YYYY-MM (e.g., erik, 2024-08)', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var parts = r1.getResponseText().split(',').map(function(s){ return s.trim(); });
  if (parts.length !== 2) {
    ui.alert('Please enter values like: username, 2024-08');
    return;
  }
  var username = parts[0];
  var ym = parts[1];
  var y = Number(ym.split('-')[0]);
  var m = Number(ym.split('-')[1]);
  fetchAndAppend_(username, y, m);
}

function promptFetchRange() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Fetch Range', 'Enter: username, YYYY-MM..YYYY-MM (e.g., erik, 2024-01..2024-03)', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var parts = r1.getResponseText().split(',').map(function(s){ return s.trim(); });
  if (parts.length !== 2 || parts[1].indexOf('..') === -1) {
    ui.alert('Please enter values like: username, 2024-01..2024-03');
    return;
  }
  var username = parts[0];
  var range = parts[1].split('..');
  var y1 = Number(range[0].split('-')[0]);
  var m1 = Number(range[0].split('-')[1]);
  var y2 = Number(range[1].split('-')[0]);
  var m2 = Number(range[1].split('-')[1]);

  var ym1 = y1 * 12 + (m1 - 1);
  var ym2 = y2 * 12 + (m2 - 1);
  for (var ym = ym1; ym <= ym2; ym++) {
    var year = Math.floor(ym / 12);
    var month = (ym % 12) + 1;
    fetchAndAppend_(username, year, month);
  }
}

function fetchAndAppend_(username, year, month) {
  var sheet = getOrCreateGamesSheet_();
  var rows = fetchChessDotComMonth_(username, year, month);
  appendRows_(sheet, rows);
}
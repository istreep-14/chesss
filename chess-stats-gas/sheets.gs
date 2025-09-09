function getOrCreateGamesSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(PROJECT.sheetTitle);
  if (!sheet) {
    sheet = ss.insertSheet(PROJECT.sheetTitle);
  }
  sheet.setFrozenRows(1);
  setupHeaders_(sheet);
  return sheet;
}

function setupHeaders_(sheet) {
  var titles = HEADERS.map(function(h) { return h.title; });
  var range = sheet.getRange(1, 1, 1, titles.length);
  range.setValues([titles]);
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');

  for (var c = 0; c < HEADERS.length; c++) {
    var header = HEADERS[c];
    if (header.width) sheet.setColumnWidth(c + 1, header.width);
    if (header.color) sheet.getRange(1, c + 1).setBackground(header.color);
    sheet.setColumnHidden(c + 1, header.visible === false);
  }
}

function buildKeyToColumnMap_() {
  var map = {};
  for (var i = 0; i < HEADERS.length; i++) {
    map[HEADERS[i].key] = { index: i + 1, letter: columnNumberToLetter_(i + 1) };
  }
  return map;
}

function columnNumberToLetter_(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function replaceFormulaPlaceholders_(template, keyToCol, rowNumber) {
  if (!template) return '';
  return String(template)
    .replace(/\$\{row\}/g, String(rowNumber))
    .replace(/\$\{col\(([^)]+)\)\}/g, function(_m, key) {
      var info = keyToCol[key];
      return info ? info.letter : 'A';
    });
}

function getExistingUrlSet_(sheet) {
  var set = Object.create(null);
  var keyToCol = buildKeyToColumnMap_();
  var urlInfo = keyToCol['url'];
  if (!urlInfo) return set;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;
  var range = sheet.getRange(2, urlInfo.index, lastRow - 1, 1);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v) set[String(v)] = true;
  }
  return set;
}

function appendRows_(sheet, rows) {
  if (!rows || !rows.length) return;

  setupHeaders_(sheet);

  // Dedupe by URL to avoid re-adding games
  var existing = getExistingUrlSet_(sheet);
  var filtered = rows.filter(function(r) { return r && r.url && !existing[r.url]; });
  if (!filtered.length) return;

  var startRow = sheet.getLastRow() + 1;
  var keyToCol = buildKeyToColumnMap_();

  var height = filtered.length;
  var width = HEADERS.length;
  var values = new Array(height);
  var formulas = new Array(height);
  for (var i = 0; i < height; i++) {
    values[i] = new Array(width);
    formulas[i] = new Array(width);
    var rowObj = filtered[i];
    for (var c = 0; c < HEADERS.length; c++) {
      var header = HEADERS[c];
      var rowNumber = startRow + i;
      if (header.formula) {
        var f = replaceFormulaPlaceholders_(header.formula, keyToCol, rowNumber);
        formulas[i][c] = f ? '=' + f : '';
        values[i][c] = '';
      } else {
        var fromKey = header.from || header.key;
        var v = rowObj[fromKey];
        values[i][c] = v === undefined || v === null ? '' : v;
        formulas[i][c] = '';
      }
    }
  }

  var range = sheet.getRange(startRow, 1, height, width);
  range.setValues(values);
  range.setFormulas(formulas);
}
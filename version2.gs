/**
 * Version 2 bootstrap for Chess.com -> Google Sheets
 *
 * Creates a Headers sheet to choose fields (enable + order), and a Games sheet
 * to write raw game data (no formulas) based on the selected headers.
 */

/** Sheet names */
var SHEET_HEADERS_NAME = 'Headers';
var SHEET_GAMES_NAME = 'Games';

/** Menu label */
var MENU_NAME = 'Version 2';

/**
 * Adds a custom menu to the spreadsheet when opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(MENU_NAME)
    .addItem('Setup sheets and headers', 'setupVersion2')
    .addSeparator()
    .addItem('Fetch games (prompt)', 'runFetchGamesPrompt')
    .addToUi();
}

/**
 * Ensures the sheets exist and populates the Headers sheet with the
 * full catalog of selectable fields from the Chess.com API and PGN.
 */
function setupVersion2() {
  var headersSheet = getOrCreateSheet_(SHEET_HEADERS_NAME);
  var gamesSheet = getOrCreateSheet_(SHEET_GAMES_NAME);

  // Prepare Headers sheet
  headersSheet.clear();
  var headerRow = [['Enabled', 'Order', 'Field', 'Source', 'Display Name', 'Description', 'Example', 'Input']];
  headersSheet.getRange(1, 1, 1, headerRow[0].length).setValues(headerRow);
  headersSheet.setFrozenRows(1);
  headersSheet.autoResizeColumns(1, headerRow[0].length);

  var catalog = buildHeaderCatalog_();
  var values = catalog.map(function(entry) {
    var display = prettifyFieldName_(entry.field, entry.source);
    return [false, '', entry.field, entry.source, display, '', '', ''];
  });
  if (values.length > 0) {
    headersSheet.getRange(2, 1, values.length, 8).setValues(values);
    headersSheet.getRange(2, 1, values.length, 1).insertCheckboxes();
  }

  // Prepare Games sheet (empty now; columns will be created on fetch)
  gamesSheet.clear();
  gamesSheet.setFrozenRows(1);
}

/**
 * Prompts for username and year-month, then fetches games to the Games sheet
 * using the currently enabled and ordered headers in the Headers sheet.
 */
function runFetchGamesPrompt() {
  var ui = SpreadsheetApp.getUi();

  var userResp = ui.prompt('Chess.com username', 'Enter the Chess.com username (e.g., magnuscarlsen):', ui.ButtonSet.OK_CANCEL);
  if (userResp.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var username = String(userResp.getResponseText() || '').trim();
  if (!username) {
    ui.alert('Username is required.');
    return;
  }

  var ymResp = ui.prompt('Year and Month', 'Enter year-month in YYYY-MM format (e.g., 2024-08):', ui.ButtonSet.OK_CANCEL);
  if (ymResp.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var ymText = String(ymResp.getResponseText() || '').trim();
  var ymMatch = /^(\d{4})-(\d{1,2})$/.exec(ymText);
  if (!ymMatch) {
    ui.alert('Invalid format. Please use YYYY-MM (e.g., 2024-08).');
    return;
  }
  var year = ymMatch[1];
  var month = ymMatch[2];

  fetchGamesToSheet_(username, year, month);
}

/**
 * Fetches monthly games for a user and writes selected fields to the Games sheet.
 * @param {string} username Chess.com username
 * @param {string|number} year Four-digit year (YYYY)
 * @param {string|number} month One- or two-digit month (1-12 or 01-12)
 */
function fetchGamesToSheet_(username, year, month) {
  var ui = SpreadsheetApp.getUi();
  var headersSheet = getOrCreateSheet_(SHEET_HEADERS_NAME);
  var gamesSheet = getOrCreateSheet_(SHEET_GAMES_NAME);

  var selectedHeaders = readSelectedHeaders_(headersSheet);
  if (selectedHeaders.length === 0) {
    ui.alert('No headers are enabled in the Headers sheet. Please enable at least one.');
    return;
  }

  var normalizedMonth = String(month);
  if (normalizedMonth.length === 1) {
    normalizedMonth = '0' + normalizedMonth;
  }
  var normalizedYear = String(year);

  var url = 'https://api.chess.com/pub/player/' + encodeURIComponent(username) + '/games/' + normalizedYear + '/' + normalizedMonth;
  var response;
  try {
    response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {
    ui.alert('Request failed: ' + e);
    return;
  }

  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    ui.alert('Request failed with status ' + code + '\nURL: ' + url + '\nBody: ' + response.getContentText());
    return;
  }

  var json;
  try {
    json = JSON.parse(response.getContentText());
  } catch (e2) {
    ui.alert('Failed to parse response JSON.');
    return;
  }

  var games = (json && json.games) || [];
  if (!Array.isArray(games) || games.length === 0) {
    ui.alert('No games found for ' + username + ' ' + normalizedYear + '-' + normalizedMonth + '.');
    // Clear Games sheet except header row if present
    gamesSheet.clear();
    return;
  }

  // Prepare header row in Games sheet using display names if provided
  var headerRow = [selectedHeaders.map(function(h) { return h.displayName || h.field; })];
  gamesSheet.clear();
  gamesSheet.getRange(1, 1, 1, headerRow[0].length).setValues(headerRow);
  gamesSheet.setFrozenRows(1);

  // Build rows
  var allRows = [];
  for (var i = 0; i < games.length; i++) {
    var game = games[i];
    var pgnText = (game && game.pgn) ? String(game.pgn) : '';
    var pgnTags = parsePgnTags_(pgnText);
    var pgnMoves = parsePgnMoves_(pgnText);

    var row = selectedHeaders.map(function(h) {
      if (h.source === 'json') {
        return deepGet_(game, h.field);
      }
      if (h.source === 'pgn') {
        return (pgnTags[h.field] != null ? pgnTags[h.field] : '');
      }
      if (h.source === 'pgn_moves') {
        return pgnMoves;
      }
      return '';
    });
    allRows.push(row);
  }

  // Write all rows
  if (allRows.length > 0) {
    gamesSheet.getRange(2, 1, allRows.length, selectedHeaders.length).setValues(allRows);
  }
}

/**
 * Reads selected headers from the Headers sheet, sorting by Order ascending.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} headersSheet
 * @return {Array<{field:string, source:string, order:number, displayName:string}>}
 */
function readSelectedHeaders_(headersSheet) {
  var lastRow = headersSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  var lastCol = headersSheet.getLastColumn();
  var headerNames = headersSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  function colIndexByName_(name) {
    for (var i = 0; i < headerNames.length; i++) {
      if (String(headerNames[i]).trim() === name) return i;
    }
    return -1;
  }
  var colEnabled = colIndexByName_('Enabled');
  var colOrder = colIndexByName_('Order');
  var colField = colIndexByName_('Field');
  var colSource = colIndexByName_('Source');
  var colDisplay = colIndexByName_('Display Name');
  var range = headersSheet.getRange(2, 1, lastRow - 1, lastCol);
  var values = range.getValues();
  var selected = [];
  for (var i = 0; i < values.length; i++) {
    var enabled = colEnabled >= 0 ? values[i][colEnabled] : values[i][0];
    var orderRaw = colOrder >= 0 ? values[i][colOrder] : values[i][1];
    var field = String((colField >= 0 ? values[i][colField] : values[i][2]) || '').trim();
    var source = String((colSource >= 0 ? values[i][colSource] : values[i][3]) || '').trim();
    var displayName = String((colDisplay >= 0 ? values[i][colDisplay] : '') || '').trim();
    if (!enabled || !field || !source) {
      continue;
    }
    var order = parseInt(orderRaw, 10);
    if (isNaN(order)) {
      order = Number.POSITIVE_INFINITY;
    }
    selected.push({ field: field, source: source, order: order, displayName: displayName, rowIndex: i });
  }
  // Stable sort: primary by order ascending, secondary by original row index
  selected.sort(function(a, b) {
    if (a.order !== b.order) {
      return a.order - b.order;
    }
    return a.rowIndex - b.rowIndex;
  });
  return selected;
}

/**
 * Safely retrieves a nested property from an object using dot-separated path.
 * Supports keys that include special characters like '@' via bracket access.
 * @param {Object} obj
 * @param {string} path dot-separated path, e.g., "white.username" or "white.@id"
 * @return {*} value or empty string when missing
 */
function deepGet_(obj, path) {
  if (obj == null || !path) {
    return '';
  }
  var segments = String(path).split('.');
  var cursor = obj;
  for (var i = 0; i < segments.length; i++) {
    var key = segments[i];
    if (cursor != null && Object.prototype.hasOwnProperty.call(cursor, key)) {
      cursor = cursor[key];
    } else {
      return '';
    }
  }
  // Flatten objects/arrays to JSON string to keep sheet values raw (no formulas)
  if (cursor != null && typeof cursor === 'object') {
    try {
      return JSON.stringify(cursor);
    } catch (e) {
      return '';
    }
  }
  return cursor != null ? cursor : '';
}

/**
 * Extracts PGN tags into a map.
 * @param {string} pgnText
 * @return {!Object<string,string>}
 */
function parsePgnTags_(pgnText) {
  var tags = {};
  if (!pgnText) {
    return tags;
  }
  var re = /^\[([A-Za-z0-9_]+)\s+"([\s\S]*?)"\]\s*$/gm;
  var match;
  while ((match = re.exec(pgnText)) !== null) {
    var tagName = match[1];
    var tagValue = match[2];
    tags[tagName] = tagValue;
  }
  return tags;
}

/**
 * Returns the SAN moves block from PGN (text after the blank line following tags).
 * @param {string} pgnText
 * @return {string}
 */
function parsePgnMoves_(pgnText) {
  if (!pgnText) {
    return '';
  }
  var parts = String(pgnText).split(/\r?\n\r?\n/);
  if (parts.length < 2) {
    return '';
  }
  // Join any additional sections in case of embedded comments
  var moves = parts.slice(1).join('\n\n').trim();
  return moves;
}

/**
 * Ensures a sheet exists; returns it. Creates at end if missing.
 * @param {string} name
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    return sheet;
  }
  return ss.insertSheet(name);
}

/**
 * Catalog of selectable headers.
 * - source = 'json' means read from game JSON via the Field path
 * - source = 'pgn' means read from PGN tag [Field "..."]
 * - source = 'pgn_moves' means the SAN moves text block
 * @return {Array<{field:string, source:'json'|'pgn'|'pgn_moves'}>}
 */
function buildHeaderCatalog_() {
  var fields = [];

  // JSON fields (Chess.com monthly games API)
  [
    'url',
    'rated',
    'time_class',
    'rules',
    'time_control',
    'end_time',
    'tcn',
    'uuid',
    'initial_setup',
    'fen',
    'white.username',
    'white.rating',
    'white.result',
    'white.uuid',
    'white.@id',
    'black.username',
    'black.rating',
    'black.result',
    'black.uuid',
    'black.@id',
    'accuracies.white',
    'accuracies.black',
    'pgn'
  ].forEach(function(path) {
    fields.push({ field: path, source: 'json' });
  });

  // PGN tag fields commonly present in Chess.com PGNs
  [
    'Event',
    'Site',
    'Date',
    'Round',
    'White',
    'Black',
    'Result',
    'WhiteElo',
    'BlackElo',
    'TimeControl',
    'ECO',
    'Opening',
    'Termination',
    'CurrentPosition',
    'UTCDate',
    'UTCTime',
    'StartTime',
    'EndTime',
    'FEN',
    'SetUp',
    'Variant'
  ].forEach(function(tag) {
    fields.push({ field: tag, source: 'pgn' });
  });

  // Moves block (SAN notation after the PGN tags)
  fields.push({ field: 'Moves', source: 'pgn_moves' });

  return fields;
}

/**
 * Generates a human-friendly display name for a field path.
 * Examples:
 *  - "white.username" -> "White Username"
 *  - "accuracies.white" -> "Accuracies White"
 *  - PGN tag sources keep the tag as-is
 * @param {string} field
 * @param {string} source
 * @return {string}
 */
function prettifyFieldName_(field, source) {
  if (!field) return '';
  if (source === 'pgn') {
    return field;
  }
  var parts = String(field)
    .replace(/[@\[\]]/g, '')
    .split(/[._]/)
    .filter(function(p) { return p && p.length; });
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    parts[i] = p.charAt(0).toUpperCase() + p.slice(1).replace(/([a-z])([A-Z])/g, '$1 $2');
  }
  return parts.join(' ');
}


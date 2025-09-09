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
    .addItem('Setup Config sheet', 'setupConfig')
    .addItem('Run configured fetch (V2)', 'runConfiguredFetchV2')
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
    // Default Display Name to the field; leave Description, Example, and Input blank for user to fill
    return [false, '', entry.field, entry.source, (entry.displayName || entry.field), (entry.description || ''), (entry.example || ''), ''];
  });
  if (values.length > 0) {
    headersSheet.getRange(2, 1, values.length, 8).setValues(values);
    headersSheet.getRange(2, 1, values.length, 1).insertCheckboxes();
  }

  // Prepare Games sheet (empty now; columns will be created on fetch)
  gamesSheet.clear();
  gamesSheet.setFrozenRows(1);
}

// Opening Info utilities removed

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

  // Prepare header row in Games sheet using the selected display names (fallback to field)
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
    var derivedReg = null; // lazy init so we do it once per fetch

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
      if (h.source === 'derived') {
        if (!derivedReg) derivedReg = getDerivedRegistry_();
        var def = derivedReg[h.field];
        if (def && typeof def.compute === 'function') {
          try {
            var v = def.compute(game, pgnTags, pgnMoves);
            // Normalize objects to JSON for safety
            if (v != null && typeof v === 'object') {
              try { return JSON.stringify(v); } catch (e) { return ''; }
            }
            return (v != null ? v : '');
          } catch (e) {
            return '';
          }
        }
        return '';
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
 * @return {Array<{field:string, source:string, order:number}>}
 */
function readSelectedHeaders_(headersSheet) {
  var lastRow = headersSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  // Columns: Enabled, Order, Field, Source, Display Name, Description, Example, Input
  var range = headersSheet.getRange(2, 1, lastRow - 1, 8);
  var values = range.getValues();
  var selected = [];
  for (var i = 0; i < values.length; i++) {
    var enabled = values[i][0];
    var orderRaw = values[i][1];
    var field = String(values[i][2] || '').trim();
    var source = String(values[i][3] || '').trim();
    var displayName = String(values[i][4] || '').trim();
    // Description (5), Example (6) are informational; Input (7) optional for future parsing
    var input = String(values[i][7] || '').trim();
    if (!enabled || !field || !source) {
      continue;
    }
    var order = parseInt(orderRaw, 10);
    if (isNaN(order)) {
      order = Number.POSITIVE_INFINITY;
    }
    selected.push({ field: field, source: source, order: order, rowIndex: i, displayName: displayName, input: input });
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
    // Fallback: if there is no tag section, treat the entire text as moves
    return String(pgnText).trim();
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

  // Derived fields (computed in code via registry)
  var derived = getDerivedRegistry_();
  Object.keys(derived).forEach(function(key) {
    var def = derived[key] || {};
    fields.push({
      field: key,
      source: 'derived',
      displayName: def.displayName || key,
      description: def.description || '',
      example: def.example || ''
    });
  });

  return fields;
}


/**
 * Registry of derived fields computed in code (no spreadsheet formulas).
 * Each entry defines a compute(game, pgnTags, pgnMoves) function returning a scalar value.
 * @return {!Object<string,{displayName:string, description:string, example:*, compute:function(Object,Object,string):*}>}
 */
function getDerivedRegistry_() {
  // Helper to parse time control strings like "300+0", "600+5", or "600"
  function parseTimeControlString_(tc) {
    var s = String(tc || '').trim();
    if (!s || s === '-') return { initialSec: null, incrementSec: null };
    var parts = s.split('+');
    var initialSec = parseInt(parts[0], 10);
    var incrementSec = parts.length > 1 ? parseInt(parts[1], 10) : 0;
    if (!isFinite(initialSec)) initialSec = null;
    if (!isFinite(incrementSec)) incrementSec = null;
    return { initialSec: initialSec, incrementSec: incrementSec };
  }

  function classifySpeed_(initialSec, incrementSec) {
    if (initialSec === '' || initialSec == null) return '';
    var inc = incrementSec || 0;
    var base = Number(initialSec) + Number(inc) * 40;
    if (base < 180) return 'bullet';
    if (base < 480) return 'blitz';
    if (base < 1500) return 'rapid';
    return 'classical';
  }

  // Compute functions
  var registry = {
    result_numeric: {
      displayName: 'Result (Numeric)',
      description: '1 (white win), 0.5 (draw), 0 (black win) from PGN Result',
      example: 1,
      compute: function(game, pgnTags, pgnMoves) {
        var r = String((pgnTags && pgnTags['Result']) || '').trim();
        if (r === '1-0') return 1;
        if (r === '0-1') return 0;
        if (r === '1/2-1/2') return 0.5;
        return '';
      }
    },
    moves_count: {
      displayName: 'Moves (count)',
      description: 'Approximate number of full moves in PGN',
      example: 32,
      compute: function(game, pgnTags, pgnMoves) {
        var text = String(pgnMoves || '');
        if (!text) return '';
        var matches = text.match(/\b\d+\./g);
        return matches ? matches.length : '';
      }
    },
    plies: {
      displayName: 'Plies',
      description: 'Approximate number of half-moves (plies)',
      example: 64,
      compute: function(game, pgnTags, pgnMoves) {
        var moves = registry.moves_count.compute(game, pgnTags, pgnMoves);
        if (moves === '' || moves == null) return '';
        // Roughly two plies per move (may be off by 1 for unfinished last move)
        return Number(moves) * 2;
      }
    },
    initial_seconds: {
      displayName: 'InitialSec',
      description: 'Initial time (seconds) parsed from time control',
      example: 300,
      compute: function(game, pgnTags, pgnMoves) {
        var tc = (game && game.time_control) || (pgnTags && pgnTags['TimeControl']) || '';
        return parseTimeControlString_(tc).initialSec;
      }
    },
    increment_seconds: {
      displayName: 'Increment',
      description: 'Increment (seconds) parsed from time control',
      example: 0,
      compute: function(game, pgnTags, pgnMoves) {
        var tc = (game && game.time_control) || (pgnTags && pgnTags['TimeControl']) || '';
        return parseTimeControlString_(tc).incrementSec;
      }
    },
    speed_class: {
      displayName: 'SpeedClass',
      description: 'bullet / blitz / rapid / classical derived from time control',
      example: 'blitz',
      compute: function(game, pgnTags, pgnMoves) {
        var tc = (game && game.time_control) || (pgnTags && pgnTags['TimeControl']) || '';
        var parsed = parseTimeControlString_(tc);
        return classifySpeed_(parsed.initialSec, parsed.incrementSec);
      }
    },
    accuracy_diff: {
      displayName: 'AccuracyDiff',
      description: 'WhiteAccuracy - BlackAccuracy from PGN tags (if present)',
      example: 3.2,
      compute: function(game, pgnTags, pgnMoves) {
        var w = parseFloat((pgnTags && pgnTags['WhiteAccuracy']) || '');
        var b = parseFloat((pgnTags && pgnTags['BlackAccuracy']) || '');
        if (!isFinite(w) || !isFinite(b)) return '';
        // Keep one decimal like Chess.com UI typically shows
        return Math.round((w - b) * 10) / 10;
      }
    }
  };

  // -------- Additional helpers for new derived fields --------
  function formatLocalDateTime_(dateObj) {
    try {
      var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
      return Utilities.formatDate(dateObj, tz || 'UTC', 'yyyy-MM-dd HH:mm:ss');
    } catch (e) {
      return '';
    }
  }

  function epochSecToDate_(epochSec) {
    if (epochSec == null || epochSec === '' || !isFinite(Number(epochSec))) return null;
    try { return new Date(Number(epochSec) * 1000); } catch (e) { return null; }
  }

  function parseHmsToSeconds_(s) {
    var t = String(s || '').trim();
    if (!t) return null;
    // Support fractional seconds like 0:02:59.9
    var rawParts = t.split(':');
    if (rawParts.length < 2) {
      var nOnly = parseFloat(t);
      return isFinite(nOnly) ? nOnly : null;
    }
    var secondsPart = rawParts.pop();
    var minutesPart = rawParts.pop();
    var hoursPart = rawParts.length > 0 ? rawParts.pop() : '0';
    var seconds = parseFloat(String(secondsPart).replace(/[^0-9\.]/g, ''));
    var minutes = parseInt(minutesPart, 10);
    var hours = parseInt(hoursPart, 10);
    if (!isFinite(seconds) || !isFinite(minutes) || !isFinite(hours)) return null;
    return hours * 3600 + minutes * 60 + seconds;
  }

  function computeGameLengthSecondsFromPgn_(pgnTags) {
    if (!pgnTags) return '';
    var start = parseHmsToSeconds_(pgnTags['StartTime']);
    var end = parseHmsToSeconds_(pgnTags['EndTime']);
    if (start == null || end == null) return '';
    var diff = end - start;
    if (diff < 0) diff += 24 * 3600; // handle midnight rollover
    return diff;
  }

  function extractClocksFromPgn_(pgnText) {
    var re = /\[%clk\s+([^\]]+)\]/g;
    var clocks = [];
    if (!pgnText) return clocks;
    var m;
    while ((m = re.exec(String(pgnText))) !== null) {
      clocks.push(m[1]);
    }
    return clocks;
  }

  function clocksToSecondsList_(clocks) {
    if (!clocks || !clocks.length) return [];
    return clocks.map(function(c){
      var n = parseHmsToSeconds_(c);
      return (typeof n === 'number' && isFinite(n)) ? n : null;
    });
  }

  // --- SAN moves helpers ---
  function stripCurlyComments_(s) {
    return String(s || '').replace(/\{[^}]*\}/g, ' ');
  }

  function stripSemicolonComments_(s) {
    return String(s || '').replace(/;[^\n]*/g, '');
  }

  function stripNagAnnotations_(s) {
    return String(s || '').replace(/\$\d+/g, '');
  }

  function removeMoveNumbers_(s) {
    return String(s || '').replace(/\b\d+\.(?:\.\.)?/g, ' ');
  }

  function normalizeWhitespace_(s) {
    return String(s || '').replace(/\s+/g, ' ').trim();
  }

  function extractSanPliesFromMoves_(movesText) {
    var t = String(movesText || '');
    if (!t) return [];
    // Remove comments and annotations
    t = stripCurlyComments_(t);
    t = stripSemicolonComments_(t);
    t = stripNagAnnotations_(t);
    // Remove move numbers and results
    t = removeMoveNumbers_(t);
    t = t.replace(/\b(1-0|0-1|1\/2-1\/2|\*)\b/g, ' ');
    // Normalize and split
    t = normalizeWhitespace_(t);
    if (!t) return [];
    var tokens = t.split(' ');
    // Filter ellipses and empties
    tokens = tokens.filter(function(tok) { return tok && tok !== '...' && tok !== '..'; });
    return tokens;
  }

  function stripBracketTags_(s) {
    return String(s || '').replace(/\[%[^\]]*\]/g, ' ');
  }

  function normalizeNumberedMovesText_(movesText) {
    var t = String(movesText || '');
    if (!t) return '';
    t = stripCurlyComments_(t);
    t = stripSemicolonComments_(t);
    t = stripNagAnnotations_(t);
    t = stripBracketTags_(t); // remove [%clk], [%eval], etc.
    // Remove result tokens
    t = t.replace(/\b(1-0|0-1|1\/2-1\/2|\*)\b/g, ' ');
    // Normalize  "1..." to "1." style
    t = t.replace(/\b(\d+)\.\.\./g, '$1.');
    // Collapse whitespace
    t = normalizeWhitespace_(t);
    return t;
  }

  function buildMoveDurations_(clockSecondsList, baseSec, incSec) {
    if (!clockSecondsList || !clockSecondsList.length) return [];
    var can = (baseSec != null && incSec != null && isFinite(baseSec) && isFinite(incSec));
    if (!can) return [];
    // Split by side: even indexes are white (0-based), odd are black
    var white = [];
    var black = [];
    for (var i = 0; i < clockSecondsList.length; i++) {
      var v = clockSecondsList[i];
      if (typeof v !== 'number' || !isFinite(v)) { if (i % 2 === 0) white.push(null); else black.push(null); continue; }
      if (i % 2 === 0) white.push(v); else black.push(v);
    }
    function durationsFor(seq) {
      var out = [];
      for (var i = 0; i < seq.length; i++) {
        var curr = seq[i];
        if (typeof curr !== 'number' || !isFinite(curr)) { out.push(''); continue; }
        if (i === 0) {
          var d0 = Math.max(0, baseSec - curr + incSec);
          out.push(Math.round(d0 * 100) / 100);
        } else {
          var prev = seq[i-1];
          if (typeof prev !== 'number' || !isFinite(prev)) { out.push(''); continue; }
          var di = Math.max(0, prev - curr + incSec);
          out.push(Math.round(di * 100) / 100);
        }
      }
      return out;
    }
    var wDur = durationsFor(white);
    var bDur = durationsFor(black);
    var merged = [];
    var maxLen = Math.max(wDur.length, bDur.length);
    for (var k = 0; k < maxLen; k++) {
      if (k < wDur.length) merged.push(wDur[k]);
      if (k < bDur.length) merged.push(bDur[k]);
    }
    return merged;
  }

  function listToBracedString_(arr) {
    if (!arr || !arr.length) return '';
    return '{' + arr.map(function(x){ return (x === '' || x == null) ? '' : String(x); }).join(', ') + '}';
  }

  // -------- New derived entries --------
  registry.base_seconds = {
    displayName: 'Base time (s)',
    description: 'Initial time (seconds) parsed from time control',
    example: 300,
    compute: function(game, pgnTags, pgnMoves) {
      var tc = (game && game.time_control) || (pgnTags && pgnTags['TimeControl']) || '';
      return parseTimeControlString_(tc).initialSec;
    }
  };

  registry.increment_seconds.displayName = 'Increment (s)';

  registry.end_time_formatted = {
    displayName: 'End Time (Local)',
    description: 'Game end time formatted in spreadsheet time zone',
    example: '2025-09-08 14:23:45',
    compute: function(game) {
      var d = epochSecToDate_(game && game.end_time);
      return d ? formatLocalDateTime_(d) : '';
    }
  };

  registry.end_year = { displayName: 'End Year', description: 'YYYY from end_time', example: 2025, compute: function(game){ var d = epochSecToDate_(game && game.end_time); return d ? d.getFullYear() : ''; } };
  registry.end_month = { displayName: 'End Month', description: '1-12 from end_time', example: 9, compute: function(game){ var d = epochSecToDate_(game && game.end_time); return d ? (d.getMonth()+1) : ''; } };
  registry.end_day = { displayName: 'End Day', description: '1-31 from end_time', example: 8, compute: function(game){ var d = epochSecToDate_(game && game.end_time); return d ? d.getDate() : ''; } };
  registry.end_hour = { displayName: 'End Hour', description: '0-23 from end_time', example: 14, compute: function(game){ var d = epochSecToDate_(game && game.end_time); return d ? d.getHours() : ''; } };
  registry.end_minute = { displayName: 'End Minute', description: '0-59 from end_time', example: 23, compute: function(game){ var d = epochSecToDate_(game && game.end_time); return d ? d.getMinutes() : ''; } };
  registry.end_second = { displayName: 'End Second', description: '0-59 from end_time', example: 45, compute: function(game){ var d = epochSecToDate_(game && game.end_time); return d ? d.getSeconds() : ''; } };
  registry.end_millisecond = { displayName: 'End Milliseconds', description: '0-999 from end_time', example: 123, compute: function(game){ var d = epochSecToDate_(game && game.end_time); return d ? d.getMilliseconds() : ''; } };

  registry.game_length_seconds = {
    displayName: 'GameLength (s)',
    description: 'Derived from PGN tags EndTime - StartTime',
    example: 420,
    compute: function(game, pgnTags) { return computeGameLengthSecondsFromPgn_(pgnTags); }
  };

  registry.start_time_derived_local = {
    displayName: 'Start Time (Local, derived)',
    description: 'End Time minus GameLength (local time zone)',
    example: '2025-09-08 14:16:45',
    compute: function(game, pgnTags) {
      var d = epochSecToDate_(game && game.end_time);
      var len = computeGameLengthSecondsFromPgn_(pgnTags);
      if (!d || !isFinite(len)) return '';
      return formatLocalDateTime_(new Date(d.getTime() - Number(len) * 1000));
    }
  };

  registry.moves_san_list = {
    displayName: 'Moves (SAN list)',
    description: 'List of SAN plies (no comments/clock/NAG/move numbers)',
    example: '{e4, e5, Nf3, Nc6, ...}',
    compute: function(game, pgnTags, pgnMoves) {
      var plies = extractSanPliesFromMoves_(pgnMoves);
      return listToBracedString_(plies);
    }
  };

  registry.moves_list_numbered = {
    displayName: 'Moves (numbered)',
    description: 'Numbered SAN movetext (comments/NAG/clock tags removed)',
    example: '1. e4 e5 2. Nf3 Nc6',
    compute: function(game, pgnTags, pgnMoves) {
      return normalizeNumberedMovesText_(pgnMoves);
    }
  };

  registry.clocks_list = {
    displayName: 'Clocks',
    description: 'Clock tags extracted from PGN, in original format',
    example: '{5:00, 5:00, 4:58, ...}',
    compute: function(game, pgnTags, pgnMoves) {
      var pgn = (pgnMoves && String(pgnMoves)) || (game && game.pgn) || '';
      var clocks = extractClocksFromPgn_(pgn);
      return listToBracedString_(clocks);
    }
  };

  registry.clock_seconds_list = {
    displayName: 'Clock Seconds',
    description: 'Clock tags converted to seconds',
    example: '{300, 300, 298, ...}',
    compute: function(game, pgnTags, pgnMoves) {
      var pgn = (pgnMoves && String(pgnMoves)) || (game && game.pgn) || '';
      var clocks = extractClocksFromPgn_(pgn);
      var secs = clocksToSecondsList_(clocks);
      return listToBracedString_(secs);
    }
  };

  registry.move_times_seconds = {
    displayName: 'Clock Seconds_Incriment',
    description: 'Per-ply durations including increment (legacy label)',
    example: '{2, 2, 3, ...}',
    compute: function(game, pgnTags, pgnMoves) {
      var pgn = (pgnMoves && String(pgnMoves)) || (game && game.pgn) || '';
      var clocks = extractClocksFromPgn_(pgn);
      var secs = clocksToSecondsList_(clocks);
      var tc = (game && game.time_control) || (pgnTags && pgnTags['TimeControl']) || '';
      var parsed = parseTimeControlString_(tc);
      var durations = buildMoveDurations_(secs, parsed.initialSec, parsed.incrementSec);
      return listToBracedString_(durations);
    }
  };

  registry.reason = {
    displayName: 'Reason',
    description: 'Termination reason from PGN tag',
    example: 'Time forfeit',
    compute: function(game, pgnTags) { return (pgnTags && pgnTags['Termination']) || ''; }
  };

  registry.format = {
    displayName: 'Format',
    description: 'Format derived from rules/time_class (e.g., blitz, rapid, live960, daily 960)',
    example: 'blitz',
    compute: function(game, pgnTags) {
      var rules = (game && game.rules) || '';
      var timeClass = (game && game.time_class) || '';
      if (rules === 'chess') return timeClass || '';
      if (rules === 'chess960') return (timeClass === 'daily') ? 'daily 960' : 'live960';
      return rules || '';
    }
  };

  registry.opening_url = {
    displayName: 'Opening URL',
    description: 'From PGN ECOUrl/OpeningUrl or JSON opening_url',
    example: 'https://www.chess.com/openings/...',
    compute: function(game, pgnTags) {
      var v = (pgnTags && (pgnTags['ECOUrl'] || pgnTags['OpeningUrl'])) || (game && game.opening_url) || '';
      return v || '';
    }
  };

  registry.endboard_url = {
    displayName: 'Endboard URL',
    description: 'Image URL for final FEN (service-dependent)',
    example: 'https://www.chess.com/dynboard?fen=...'
    ,compute: function(game) {
      var fen = (game && game.fen) || '';
      if (!fen) return '';
      // Generic fallback compatible with chess.com dynamic board
      return 'https://www.chess.com/dynboard?fen=' + encodeURIComponent(String(fen));
    }
  };

  registry.rating_difference = {
    displayName: 'RatingDiff (opp - mine)',
    description: 'Black rating minus White rating (no player context in V2)',
    example: 35,
    compute: function(game) {
      var w = game && game.white && game.white.rating;
      var b = game && game.black && game.black.rating;
      if (!isFinite(w) || !isFinite(b)) return '';
      return Number(b) - Number(w);
    }
  };

  return registry;
}


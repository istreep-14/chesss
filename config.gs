// Config sheet support: setup, read, and run configured fetch/recalc

/** @const {string} */
var CONFIG_SHEET_NAME = 'Config';

/**
 * Creates or resets the Config sheet with controls for username, archive selection,
 * write mode, and data group toggles. Safe to re-run.
 */
function setupConfig() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  }
  sheet.clear();

  // Column widths for readability
  sheet.setColumnWidth(1, 220); // A
  sheet.setColumnWidth(2, 240); // B
  sheet.setColumnWidth(3, 200); // C

  // Titles
  sheet.getRange('A1').setValue('General Settings');
  sheet.getRange('A1').setFontWeight('bold');

  var settingsRows = [
    ['Username', ''],
    ['Write Mode', 'Append'],
    ['Archives Mode', 'Count'],
    ['Count (N)', 2],
    ['Count Direction', 'Newest'],
    ['Range Start (YYYY-MM)', ''],
    ['Range End (YYYY-MM)', '']
  ];
  sheet.getRange(2, 1, settingsRows.length, 2).setValues(settingsRows);

  // Data validation for dropdowns
  var dvWrite = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Append', 'Clear then Append'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B3').setDataValidation(dvWrite);

  var dvMode = SpreadsheetApp.newDataValidation()
    .requireValueInList(['All', 'Count', 'Range'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B4').setDataValidation(dvMode);

  var dvDir = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Newest', 'Oldest'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B6').setDataValidation(dvDir);

  // Data Groups section
  var tableStartRow = 10;
  sheet.getRange(tableStartRow, 1).setValue('Data Groups');
  sheet.getRange(tableStartRow, 1).setFontWeight('bold');
  sheet.getRange(tableStartRow + 1, 1, 1, 3).setValues([[
    'Data Group', 'Calculate New', 'Recalculate'
  ]]);
  sheet.getRange(tableStartRow + 1, 1, 1, 3).setFontWeight('bold');

  var groups = [
    ['Archive API', true, false],
    ['Derived', true, false]
  ];
  sheet.getRange(tableStartRow + 2, 1, groups.length, 3).setValues(groups);

  // Checkboxes for booleans
  var cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(tableStartRow + 2, 2, groups.length, 2).setDataValidation(cbRule);
}

/**
 * Reads the Config sheet into a structured object. Ensures defaults.
 * @return {{ username: string,
 *            writeMode: 'append'|'clear',
 *            selection: { type: 'all'|'count'|'range', count?: number, direction?: 'newest'|'oldest', start?: string, end?: string },
 *            groups: { archiveApi: { calculateNew: boolean, recalculate: boolean }, derived: { calculateNew: boolean, recalculate: boolean } } }}
 */
function readConfig_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) throw new Error('Config sheet not found. Run setupConfig() first.');

  // Read key-value settings in columns A:B from row 2 until first blank key
  var settings = {};
  for (var r = 2; r < 200; r++) {
    var key = String(sheet.getRange(r, 1).getValue() || '').trim();
    var val = sheet.getRange(r, 2).getValue();
    if (!key) break;
    settings[key] = val;
  }

  var username = String(settings['Username'] || '').trim();
  var writeModeRaw = String(settings['Write Mode'] || 'Append').toLowerCase();
  var writeMode = writeModeRaw.indexOf('clear') !== -1 ? 'clear' : 'append';
  var archivesModeRaw = String(settings['Archives Mode'] || 'Count').toLowerCase();
  var selType = 'count';
  if (archivesModeRaw === 'all') selType = 'all';
  else if (archivesModeRaw === 'range') selType = 'range';

  var count = Number(settings['Count (N)'] || 2);
  if (!isFinite(count) || count < 0) count = 0;
  var dirRaw = String(settings['Count Direction'] || 'Newest').toLowerCase();
  var direction = dirRaw === 'oldest' ? 'oldest' : 'newest';
  var start = String(settings['Range Start (YYYY-MM)'] || '').trim();
  var end = String(settings['Range End (YYYY-MM)'] || '').trim();

  // Locate Data Groups header row
  var headerRow = -1;
  var lastRow = sheet.getLastRow();
  for (var rr = 1; rr <= Math.min(lastRow, 200); rr++) {
    if (String(sheet.getRange(rr, 1).getValue() || '').trim() === 'Data Groups') {
      headerRow = rr + 1; // header titles are next row
      break;
    }
  }
  if (headerRow === -1) throw new Error('Data Groups header not found in Config. Run setupConfig() again.');

  // Read table rows until blank Data Group name
  var groups = { archiveApi: { calculateNew: true, recalculate: false }, derived: { calculateNew: true, recalculate: false } };
  for (var tr = headerRow + 1; tr <= lastRow; tr++) {
    var name = String(sheet.getRange(tr, 1).getValue() || '').trim();
    if (!name) break;
    var calc = !!sheet.getRange(tr, 2).getValue();
    var recal = !!sheet.getRange(tr, 3).getValue();
    if (/^archive api$/i.test(name)) groups.archiveApi = { calculateNew: calc, recalculate: recal };
    if (/^derived$/i.test(name)) groups.derived = { calculateNew: calc, recalculate: recal };
  }

  return {
    username: username,
    writeMode: writeMode,
    selection: { type: selType, count: count, direction: direction, start: start, end: end },
    groups: groups
  };
}

/**
 * Utility: clear all data rows but keep the header row and column structure.
 */
function clearDataKeepHeader_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
}

/**
 * Returns all archive URLs for a given username.
 * @param {string} username
 * @return {string[]}
 */
function getArchivesForUsername(username) {
  var base = 'https://api.chess.com/pub/player/' + encodeURIComponent(username);
  var profileResp = UrlFetchApp.fetch(base, { muteHttpExceptions: true });
  if (profileResp.getResponseCode() !== 200) {
    throw new Error('Player not found or API error for username: ' + username);
  }
  var archivesResp = UrlFetchApp.fetch(base + '/games/archives', { muteHttpExceptions: true });
  if (archivesResp.getResponseCode() !== 200) {
    throw new Error('Archives fetch failed: ' + archivesResp.getResponseCode());
  }
  var archives = JSON.parse(archivesResp.getContentText());
  return archives.archives || [];
}

/**
 * Extracts YYYY-MM from a Chess.com archive URL, or empty string if not found.
 * @param {string} url
 * @return {string}
 */
function monthFromArchiveUrl_(url) {
  var m = String(url || '').match(/\/(\d{4})\/(\d{2})\s*$/);
  if (!m) return '';
  return m[1] + '-' + m[2];
}

/**
 * Lexicographic compare for YYYY-MM strings.
 */
function compareYm_(a, b) {
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

/**
 * Runs fetch/append per Config and triggers configured processors.
 * - Clears sheet first if selected
 * - Selects archives based on All/Count/Range and direction
 * - Appends new rows; skips existing by URL
 * - Runs data groups for new rows per Calculate New
 * - Optionally recalculates across existing rows per Recalculate
 */
function runConfiguredFetchV2() {
  var cfg = readConfig_();
  if (!cfg.username) throw new Error('Username is empty in Config.');

  var headersSheet = getOrCreateSheet_(SHEET_HEADERS_NAME);
  var gamesSheet = getOrCreateSheet_(SHEET_GAMES_NAME);
  var selectedHeaders = readSelectedHeaders_(headersSheet);
  if (selectedHeaders.length === 0) {
    throw new Error('No headers are enabled in the Headers sheet.');
  }

  // If not fetching (Archive API Calculate New is off), only run recalc toggles
  if (!(cfg.groups && cfg.groups.archiveApi && cfg.groups.archiveApi.calculateNew)) {
    if (cfg.groups && cfg.groups.derived && cfg.groups.derived.recalculate) {
      recalcDerivedV2_(headersSheet, gamesSheet, selectedHeaders);
    }
    return;
  }

  // Prepare header row expectations
  var expectedHeaderRow = [selectedHeaders.map(function(h) { return h.displayName || h.field; })];

  if (cfg.writeMode === 'clear') {
    gamesSheet.clear();
    gamesSheet.getRange(1, 1, 1, expectedHeaderRow[0].length).setValues(expectedHeaderRow);
    gamesSheet.setFrozenRows(1);
  } else {
    // Append mode: verify header row matches current selection
    var lastCol = gamesSheet.getLastColumn();
    if (lastCol > 0) {
      var currentHeader = gamesSheet.getRange(1, 1, 1, lastCol).getValues();
      var eq = (currentHeader[0].length === expectedHeaderRow[0].length) && currentHeader[0].every(function(v, i) { return String(v) === String(expectedHeaderRow[0][i]); });
      if (!eq && gamesSheet.getLastRow() > 1) {
        throw new Error('Header mismatch for append. Use "Clear then Append" or update Headers to match existing columns.');
      }
      if (!eq) {
        gamesSheet.clear();
        gamesSheet.getRange(1, 1, 1, expectedHeaderRow[0].length).setValues(expectedHeaderRow);
        gamesSheet.setFrozenRows(1);
      }
    } else {
      gamesSheet.getRange(1, 1, 1, expectedHeaderRow[0].length).setValues(expectedHeaderRow);
      gamesSheet.setFrozenRows(1);
    }
  }

  // For dedupe we require the 'url' field to be included
  var urlHeaderIndex = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'url'; });
  if (urlHeaderIndex === -1) {
    throw new Error('Dedupe requires the "url" field in Headers. Enable it before appending.');
  }

  // Build archive list
  var archives = getArchivesForUsername(cfg.username);
  var selected = archives.slice();
  if (cfg.selection.type === 'count') {
    var n = Math.max(0, Math.floor(cfg.selection.count || 0));
    if (n > 0) {
      selected = (cfg.selection.direction === 'oldest') ? selected.slice(0, n) : selected.slice(-n);
    } else {
      selected = [];
    }
  } else if (cfg.selection.type === 'range') {
    var startYm = String(cfg.selection.start || '').trim();
    var endYm = String(cfg.selection.end || '').trim();
    if (!startYm || !endYm) throw new Error('Range Start/End must be provided for Archives Mode = Range.');
    selected = selected.filter(function(u) {
      var ym = monthFromArchiveUrl_(u);
      return ym && compareYm_(startYm, ym) <= 0 && compareYm_(ym, endYm) <= 0;
    });
  }
  if (!selected.length) return;

  // Existing URL set from sheet
  var existingUrlSet = readExistingUrlsV2_(gamesSheet, urlHeaderIndex);

  // Build and append in chunks per archive
  for (var i = 0; i < selected.length; i++) {
    var url = selected[i];
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) continue;
    var data;
    try { data = JSON.parse(resp.getContentText()); } catch (e) { continue; }
    var games = (data && data.games) || [];
    if (!games.length) continue;

    var rows = [];
    for (var g = 0; g < games.length; g++) {
      var game = games[g];
      var gameUrl = (game && game.url) ? String(game.url) : '';
      if (!gameUrl) continue;
      if (existingUrlSet.has(gameUrl)) continue;

      var pgnText = (game && game.pgn) ? String(game.pgn) : '';
      var pgnTags = parsePgnTags_(pgnText);
      var pgnMoves = parsePgnMoves_(pgnText);
      var derivedReg = null;
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
          if (!(cfg.groups && cfg.groups.derived && cfg.groups.derived.calculateNew)) return '';
          if (!derivedReg) derivedReg = getDerivedRegistry_();
          var def = derivedReg[h.field];
          if (def && typeof def.compute === 'function') {
            try {
              var v = def.compute(game, pgnTags, pgnMoves);
              if (v != null && typeof v === 'object') { try { return JSON.stringify(v); } catch (e) { return ''; } }
              return (v != null ? v : '');
            } catch (e) { return ''; }
          }
          return '';
        }
        return '';
      });
      rows.push(row);
      existingUrlSet.add(gameUrl);
    }
    if (rows.length) {
      appendRowsV2_(gamesSheet, rows, selectedHeaders.length);
    }
  }

  // Recalculate derived across sheet if requested
  if (cfg.groups && cfg.groups.derived && cfg.groups.derived.recalculate) {
    recalcDerivedV2_(headersSheet, gamesSheet, selectedHeaders);
  }
}

/**
 * Runs recalculation tasks across all rows based on the config toggles.
 */
function recalcDerivedV2_(headersSheet, gamesSheet, selectedHeaders) {
  // Identify which columns are derived in current selection
  var derivedIndices = [];
  for (var i = 0; i < selectedHeaders.length; i++) {
    if (selectedHeaders[i].source === 'derived') derivedIndices.push(i);
  }
  if (!derivedIndices.length) return;

  var lastRow = gamesSheet.getLastRow();
  var lastCol = gamesSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  // Build mapping for inputs we may use
  var idxUrl = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'url'; });
  var idxTimeControl = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'time_control'; });
  var idxPgn = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'pgn'; });
  var idxMoves = selectedHeaders.findIndex(function(h) { return h.source === 'pgn_moves'; });

  var values = gamesSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var derivedReg = getDerivedRegistry_();

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    // Minimal inputs for compute()
    var game = {};
    if (idxTimeControl !== -1) game.time_control = row[idxTimeControl];
    var pgnText = idxPgn !== -1 ? String(row[idxPgn] || '') : '';
    var pgnTags = parsePgnTags_(pgnText);
    var pgnMoves = idxMoves !== -1 ? String(row[idxMoves] || '') : '';

    var changed = false;
    for (var di = 0; di < derivedIndices.length; di++) {
      var c = derivedIndices[di];
      var h = selectedHeaders[c];
      var def = derivedReg[h.field];
      var newVal = '';
      if (def && typeof def.compute === 'function') {
        try {
          var v = def.compute(game, pgnTags, pgnMoves);
          if (v != null && typeof v === 'object') { try { newVal = JSON.stringify(v); } catch (e) { newVal = ''; } }
          else newVal = (v != null ? v : '');
        } catch (e) { newVal = ''; }
      }
      if (row[c] !== newVal) { row[c] = newVal; changed = true; }
    }
    if (changed) values[r] = row;
  }

  gamesSheet.getRange(2, 1, values.length, lastCol).setValues(values);
}

function readExistingUrlsV2_(gamesSheet, urlHeaderIndex) {
  var set = new Set();
  var lastRow = gamesSheet.getLastRow();
  var lastCol = gamesSheet.getLastColumn();
  if (lastRow < 2 || urlHeaderIndex === -1) return set;
  var values = gamesSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < values.length; i++) {
    var u = String(values[i][urlHeaderIndex] || '').trim();
    if (u) set.add(u);
  }
  return set;
}

function appendRowsV2_(gamesSheet, rows, width) {
  if (!rows || !rows.length) return;
  var start = gamesSheet.getLastRow() + 1;
  gamesSheet.getRange(start, 1, rows.length, width).setValues(rows);
}

/**
 * Normalizes a Chess.com game object like normalizeGame(), but uses a provided
 * username instead of the global PLAYER_USERNAME.
 * @param {Object} game
 * @param {string} username
 * @return {Object}
 */
function normalizeGameForUsername_(game, username) {
  var playerLower = String(username || '').toLowerCase();
  var color = '';
  var side = null;
  var opponent = null;

  var white = game.white || {};
  var black = game.black || {};
  var whiteUser = (white.username || '').toLowerCase();
  var blackUser = (black.username || '').toLowerCase();

  if (whiteUser === playerLower) {
    color = 'white';
    side = white;
    opponent = black;
  } else if (blackUser === playerLower) {
    color = 'black';
    side = black;
    opponent = white;
  } else {
    return { isPlayerInGame: false };
  }

  var rated = !!game.rated;
  var timeClass = game.time_class || '';
  var timeControl = game.time_control || '';
  var rules = game.rules || '';
  var format = '';
  var result = determineResult(side, opponent);
  var reason = determineReason(side, opponent);
  var resultValue = determineResultValue(result);

  var fen = game.fen || '';
  var fenEncoded = encodeURIComponent(fen);
  var imageUrl = ENDBOARD_BASE_URL + '?fen=' + fenEncoded + '&board=' + encodeURIComponent(ENDBOARD_BOARD) + '&piece=' + encodeURIComponent(ENDBOARD_PIECE) + '&size=' + encodeURIComponent(String(ENDBOARD_SIZE));

  var endTime = game.end_time ? new Date(game.end_time * 1000) : new Date();

  var eco = '';
  var openingUrl = '';
  var pgnText = game.pgn || '';
  var movesSan = '';
  var clocksStr = '';
  var parsedMoves = [];
  if (game.pgn) {
    try {
      var pgn = String(game.pgn).replace(/\r\n/g, '\n');
      pgnText = pgn;
      var ecoMatch = pgn.match(/^\[ECO\s+"([^"]+)"\]/m);
      if (ecoMatch && ecoMatch[1]) eco = ecoMatch[1];
      var ecoUrlMatch = pgn.match(/^\[(?:ECOUrl|OpeningUrl)\s+"([^"]+)"\]/m);
      if (ecoUrlMatch && ecoUrlMatch[1]) openingUrl = ecoUrlMatch[1];
      var sanClockParsed = parsePgnToSanAndClocks(pgn);
      if (sanClockParsed) {
        movesSan = sanClockParsed.movesSan || movesSan;
        clocksStr = sanClockParsed.clocks || clocksStr;
      }
      var movetext = extractMovetextFromPgn(pgn);
      parsedMoves = movetext ? splitMovesWithClocks(movetext) : [];
    } catch (e) {}
  }
  if (!eco && game.eco && !/^https?:\/\//i.test(String(game.eco))) eco = String(game.eco);
  if (!openingUrl && game.opening_url) openingUrl = String(game.opening_url);
  if (eco && /^https?:\/\//i.test(eco)) { if (!openingUrl) openingUrl = eco; eco = ''; }

  if (rules === 'chess') {
    format = timeClass || '';
  } else if (rules === 'chess960') {
    format = (timeClass === 'daily') ? 'daily 960' : 'live960';
  } else {
    format = rules || '';
  }

  var gameId = '';
  if (game.url) {
    var idMatch = String(game.url).match(/\/(\d+)(?:\?.*)?$/);
    if (idMatch && idMatch[1]) gameId = idMatch[1];
  }

  var tcParsedFinal = parseTimeControlToBaseInc(timeControl, timeClass);
  var baseSecondsFinal = tcParsedFinal.baseSeconds;
  var incSecondsFinal = tcParsedFinal.incrementSeconds;
  var derivedFinal = buildClockSecondsAndMoveTimes(parsedMoves, baseSecondsFinal, incSecondsFinal);
  var clockSecondsStrFinal = derivedFinal.clockSecondsStr;
  var moveTimesStrFinal = derivedFinal.moveTimesStr;
  var clockSecondsIncrementStrFinal = moveTimesStrFinal;

  return {
    isPlayerInGame: true,
    timestamp: endTime,
    url: game.url || '',
    gameId: gameId,
    rated: rated,
    timeClass: timeClass,
    timeControl: timeControl,
    rules: rules,
    format: format,
    color: color,
    opponentUsername: opponent && opponent.username ? opponent.username : 'Unknown',
    opponentRating: opponent && opponent.rating ? opponent.rating : '',
    myRating: side && side.rating ? side.rating : '',
    result: determineResult(side, opponent),
    resultValue: determineResultValue(result),
    reason: determineReason(side, opponent),
    fen: fen,
    imageUrl: imageUrl,
    eco: eco,
    openingUrl: openingUrl,
    pgn: pgnText,
    movesSan: movesSan,
    clocks: clocksStr,
    baseSeconds: baseSecondsFinal,
    incrementSeconds: incSecondsFinal,
    clockSeconds: clockSecondsStrFinal,
    clockSecondsIncrement: clockSecondsIncrementStrFinal,
    moveTimes: moveTimesStrFinal
  };
}


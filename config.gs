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
    ['Derived', true, false],
    ['Opening Info', false, false]
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
  var groups = { archiveApi: { calculateNew: true, recalculate: false }, derived: { calculateNew: true, recalculate: false }, openingInfo: { calculateNew: false, recalculate: false } };
  for (var tr = headerRow + 1; tr <= lastRow; tr++) {
    var name = String(sheet.getRange(tr, 1).getValue() || '').trim();
    if (!name) break;
    var calc = !!sheet.getRange(tr, 2).getValue();
    var recal = !!sheet.getRange(tr, 3).getValue();
    if (/^archive api$/i.test(name)) groups.archiveApi = { calculateNew: calc, recalculate: recal };
    if (/^derived$/i.test(name)) groups.derived = { calculateNew: calc, recalculate: recal };
    if (/^opening info$/i.test(name)) groups.openingInfo = { calculateNew: calc, recalculate: recal };
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
// (moved) getArchivesForUsername -> fetchConfigured.gs

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
// (moved) runConfiguredFetchV2 -> fetchConfigured.gs

/**
 * Runs recalculation tasks across all rows based on the config toggles.
 */
// (moved) recalcDerivedV2_ -> fetchConfigured.gs

// (moved) readExistingUrlsV2_ -> fetchConfigured.gs

// (moved) appendRowsV2_ -> fetchConfigured.gs

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


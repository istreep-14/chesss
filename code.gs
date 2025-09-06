// =====================
// Configuration
// =====================
// Update values here to control behavior without touching the rest of the code.
const PLAYER_USERNAME = 'frankscobey';
const SHEET_ID = '12zoMMkrkZz9WmhL4ds9lo8-91o-so337dOKwU3_yK3k';
const SHEET_NAME = 'Games';

// How many recent monthly archives to scan on each sync run
const RECENT_ARCHIVES_TO_SCAN = 2;

// Global batch size for all batch-processing functions
const BATCH_SIZE = 200;
// Backfill batch limits per run (legacy vars kept for compatibility)
const BACKFILL_MOVES_CLOCKS_BATCH = BATCH_SIZE;    // for backfillMovesAndClocks()
const BACKFILL_ECO_OPENING_BATCH = BATCH_SIZE;     // for backfillEcoAndOpening()
const BACKFILL_CALLBACK_FIELDS_BATCH = BATCH_SIZE; // for backfillCallbackFields()

// Throttling to be gentle with APIs and Sheets writes
const THROTTLE_APPEND_EVERY_N_ROWS = 100;   // sleep after this many buffered rows
const THROTTLE_SLEEP_MS = 200;              // ms to sleep when throttling

// API base URLs
const CHESS_COM_API_BASE = 'https://api.chess.com/pub';
const CHESS_COM_CALLBACK_BASE = 'https://www.chess.com/callback';

// If false, the general callback backfill will NOT write these callback-derived
// header columns. They are handled by a dedicated function instead.
const ENABLE_CALLBACK_HEADERS_IN_FULL_BACKFILL = false;

// Endboard image settings
const ENDBOARD_BASE_URL = 'https://www.chess.com/dynboard';
const ENDBOARD_BOARD = 'brown';
const ENDBOARD_PIECE = 'neo';
const ENDBOARD_SIZE = 3; // 1..10

// Trigger interval minutes for the installer helper
const TRIGGER_INTERVAL_MINUTES = 15;

// Elo K-factor to use for pregame formula estimation for established players
const ELO_K_ESTABLISHED = 10;

// =====================
// Lichess configuration
// =====================
// Provide your Lichess API token via Script Properties (recommended):
//   setLichessToken('your_token_here')
// Or set LICHESS_TOKEN_DEFAULT below (less secure). Script Properties takes precedence.
const LICHESS_API_BASE = 'https://lichess.org';
const LICHESS_TOKEN_DEFAULT = '';
function setLichessToken(token) {
  PropertiesService.getScriptProperties().setProperty('LICHESS_TOKEN', String(token || ''));
}
function getLichessToken_() {
  var p = PropertiesService.getScriptProperties().getProperty('LICHESS_TOKEN');
  var t = (p && String(p)) || String(LICHESS_TOKEN_DEFAULT || '');
  return t.trim();
}
function getLichessAuthHeaders_(accept) {
  var headers = {};
  if (accept) headers['Accept'] = accept;
  var tok = getLichessToken_();
  if (tok) headers['Authorization'] = 'Bearer ' + tok;
  return headers;
}

function setupSheet() {
  ensureSheet();
}

// Run this periodically via time-driven trigger (e.g., every 5–15 minutes)
function syncRecentGames() {
  const sheet = ensureSheet();
  const existingUrls = loadExistingUrls(sheet);
  const archives = getArchives();
  if (!archives || archives.length === 0) return;

  const recentArchives = archives.slice(-RECENT_ARCHIVES_TO_SCAN);
  const rowsToAppend = [];

  for (var i = 0; i < recentArchives.length; i++) {
    var monthUrl = recentArchives[i];
    var monthData = fetchJson(monthUrl);
    if (!monthData || !monthData.games || monthData.games.length === 0) continue;

    for (var j = 0; j < monthData.games.length; j++) {
      var game = monthData.games[j];
      if (!game || !game.url) continue;
      if (existingUrls.has(game.url)) continue;

      var normalized = normalizeGame(game);
      if (!normalized || !normalized.isPlayerInGame) continue;

      rowsToAppend.push(buildRow(normalized));
      existingUrls.add(game.url);
    }
  }

  appendRows(sheet, rowsToAppend);
  // Grouped processing for just-added rows
  if (rowsToAppend.length) {
    var startRow = (sheet.getLastRow() - rowsToAppend.length) + 1;
    processNewRowsGrouped(startRow, rowsToAppend.length, { includeCallback: true });
  }
}

/**
 * Batch-ingests new games (non-callback), respecting a limit. Returns number
 * of rows appended in this run.
 * @param {number=} limit Optional cap on rows appended this run (defaults to BATCH_SIZE)
 * @return {number}
 */
function ingestNewGamesBatch(limit) {
  const sheet = ensureSheet();
  const existingUrls = loadExistingUrls(sheet);
  const archives = getArchives();
  if (!archives || archives.length === 0) return 0;

  const recentArchives = archives.slice(-RECENT_ARCHIVES_TO_SCAN);
  const rowsToAppend = [];
  var effectiveLimit = (typeof limit === 'number' && isFinite(limit) && limit >= 0) ? limit : BATCH_SIZE;

  for (var i = 0; i < recentArchives.length; i++) {
    var monthUrl = recentArchives[i];
    var monthData = fetchJson(monthUrl);
    if (!monthData || !monthData.games || monthData.games.length === 0) continue;

    for (var j = 0; j < monthData.games.length; j++) {
      if (rowsToAppend.length >= effectiveLimit) break;
      var game = monthData.games[j];
      if (!game || !game.url) continue;
      if (existingUrls.has(game.url)) continue;

      var normalized = normalizeGame(game);
      if (!normalized || !normalized.isPlayerInGame) continue;

      rowsToAppend.push(buildRow(normalized));
      existingUrls.add(game.url);
    }
    if (rowsToAppend.length >= effectiveLimit) break;
  }

  appendRows(sheet, rowsToAppend);
  // Grouped processing for just-added rows in this batch
  if (rowsToAppend.length) {
    var startRow = (sheet.getLastRow() - rowsToAppend.length) + 1;
    processNewRowsGrouped(startRow, rowsToAppend.length, { includeCallback: true });
  }
  return rowsToAppend.length;
}

/**
 * Runs non-callback ingestion in repeated batches until no more rows are appended.
 * Batch size is controlled by BATCH_SIZE.
 */
function runNonCallbackBatchToCompletion() {
  while (true) {
    var appended = ingestNewGamesBatch(BATCH_SIZE);
    if (appended < BATCH_SIZE) break;
    if (THROTTLE_APPEND_EVERY_N_ROWS > 0) Utilities.sleep(THROTTLE_SLEEP_MS);
  }
}

// One-time (or repeated) backfill of your entire archive history
// If you have many games, this might need multiple runs due to execution time limits.
function backfillAllGamesOnce() {
  const sheet = ensureSheet();
  const existingUrls = loadExistingUrls(sheet);
  const archives = getArchives();
  if (!archives || archives.length === 0) return;

  const rowsToAppend = [];

  for (var i = 0; i < archives.length; i++) {
    var monthUrl = archives[i];
    var monthData = fetchJson(monthUrl);
    if (!monthData || !monthData.games || monthData.games.length === 0) continue;

    for (var j = 0; j < monthData.games.length; j++) {
      var game = monthData.games[j];
      if (!game || !game.url) continue;
      if (existingUrls.has(game.url)) continue;

      var normalized = normalizeGame(game);
      if (!normalized || !normalized.isPlayerInGame) continue;

      rowsToAppend.push(buildRow(normalized));
      existingUrls.add(game.url);

      // Optional: light throttling to be gentle with API/sheet
      if (THROTTLE_APPEND_EVERY_N_ROWS > 0 && (rowsToAppend.length % THROTTLE_APPEND_EVERY_N_ROWS) === 0) {
        Utilities.sleep(THROTTLE_SLEEP_MS);
      }
    }

    // Optional: write in batches per month to avoid large in-memory arrays
    appendRows(sheet, rowsToAppend);
    rowsToAppend.length = 0;
  }
  // After a full backfill, run non-callback grouped processing for all rows
  processAllRowsGrouped({ includeCallback: false });
}

function getArchives() {
  const profileResp = UrlFetchApp.fetch(
    CHESS_COM_API_BASE + '/player/' + encodeURIComponent(PLAYER_USERNAME),
    { muteHttpExceptions: true }
  );
  if (profileResp.getResponseCode() !== 200) {
    throw new Error('Player not found or API error: ' + profileResp.getResponseCode());
  }

  const archivesResp = UrlFetchApp.fetch(
    CHESS_COM_API_BASE + '/player/' + encodeURIComponent(PLAYER_USERNAME) + '/games/archives',
    { muteHttpExceptions: true }
  );
  if (archivesResp.getResponseCode() !== 200) {
    throw new Error('Archives fetch failed: ' + archivesResp.getResponseCode());
  }

  const archives = JSON.parse(archivesResp.getContentText());
  return archives.archives || [];
}

function fetchJson(url) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;
  try {
    return JSON.parse(resp.getContentText());
  } catch (e) {
    return null;
  }
}

/**
 * Ensures the sheet exists and has all required headers (initial, legacy migrations,
 * ECO/Opening URL, Rating change, and extended callback-derived metadata).
 * Unified setup function for sheet creation and header management.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The sheet reference
 */
function ensureSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'URL',
      'Game ID',
      'Rated',
      'Time class',
      'Time control',
      'Base time (s)',
      'Increment (s)',
      'Rules',
      'Format',
      'Color',
      'Opponent',
      'Opponent rating',
      'Opponent rating pregame',
      'Opponent rating change',
      'My rating',
      'My rating pregame',
      'My rating pregame_derived',
      'My rating change',
      'Result',
      'Result_Value',
      'Reason',
      'FEN',
      'Endboard URL',
      'ECO',
      'Opening URL',
      'Moves (SAN)',
      'Clocks',
      'Clock Seconds',
      'Move Times'
    ]);
    return sheet;
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return sheet;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Legacy: rename "Moves" to "Moves (SAN)" and add "Clocks" after it
  var movesSanIdx = headers.indexOf('Moves (SAN)');
  var clocksIdx = headers.indexOf('Clocks');
  var legacyMovesIdx = headers.indexOf('Moves');
  if (movesSanIdx === -1 && legacyMovesIdx !== -1) {
    sheet.getRange(1, legacyMovesIdx + 1).setValue('Moves (SAN)');
    movesSanIdx = legacyMovesIdx;
  }
  if (clocksIdx === -1) {
    var insertAfter = (movesSanIdx !== -1) ? (movesSanIdx + 1) : lastCol;
    sheet.insertColumnsAfter(insertAfter, 1);
    sheet.getRange(1, insertAfter + 1).setValue('Clocks');
  }

  // Ensure ECO and Opening URL exist
  var updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  var hasEco = headers.indexOf('ECO') !== -1;
  var hasOpeningUrl = headers.indexOf('Opening URL') !== -1;
  var toAdd = [];
  if (!hasEco) toAdd.push('ECO');
  if (!hasOpeningUrl) toAdd.push('Opening URL');
  if (toAdd.length) {
    sheet.insertColumnsAfter(updatedLastCol, toAdd.length);
    sheet.getRange(1, updatedLastCol + 1, 1, toAdd.length).setValues([toAdd]);
  }

  // Skip adding legacy "Rating change" column

  // Ensure Moves (SAN) and Clocks exist if sheet had neither
  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  if (headers.indexOf('Moves (SAN)') === -1 || headers.indexOf('Clocks') === -1) {
    var needMovesSan = headers.indexOf('Moves (SAN)') === -1;
    var needClocks = headers.indexOf('Clocks') === -1;
    var addList = [];
    if (needMovesSan) addList.push('Moves (SAN)');
    if (needClocks) addList.push('Clocks');
    sheet.insertColumnsAfter(updatedLastCol, addList.length);
    sheet.getRange(1, updatedLastCol + 1, 1, addList.length).setValues([addList]);
  }

  // Ensure Base/Increment appear immediately after Time control
  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  var timeControlIdx = headers.indexOf('Time control');
  if (timeControlIdx !== -1) {
    var baseIdx = headers.indexOf('Base time (s)');
    var incIdx = headers.indexOf('Increment (s)');
    var insertAfter = timeControlIdx + 1; // 1-based column is timeControlIdx+1, we insert after -> +1
    var toInsert = [];
    if (baseIdx === -1) toInsert.push('Base time (s)');
    if (incIdx === -1) toInsert.push('Increment (s)');
    if (toInsert.length) {
      sheet.insertColumnsAfter(insertAfter, toInsert.length);
      sheet.getRange(1, insertAfter + 1, 1, toInsert.length).setValues([toInsert]);
    }
  }

  // Ensure Clock Seconds and Move Times appear immediately after Clocks
  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  var clocksHeaderIdx = headers.indexOf('Clocks');
  if (clocksHeaderIdx !== -1) {
    var clkSecIdx = headers.indexOf('Clock Seconds');
    var moveTimesIdx = headers.indexOf('Move Times');
    var afterClocks = clocksHeaderIdx + 1;
    var toInsert2 = [];
    if (clkSecIdx === -1) toInsert2.push('Clock Seconds');
    if (moveTimesIdx === -1) toInsert2.push('Move Times');
    if (toInsert2.length) {
      sheet.insertColumnsAfter(afterClocks, toInsert2.length);
      sheet.getRange(1, afterClocks + 1, 1, toInsert2.length).setValues([toInsert2]);
    }
  }

  // Ensure Result_Value appears immediately after Result
  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  var resultIdx = headers.indexOf('Result');
  if (resultIdx !== -1) {
    var resultValueIdx = headers.indexOf('Result_Value');
    if (resultValueIdx === -1) {
      var afterResult = resultIdx + 1;
      sheet.insertColumnsAfter(afterResult, 1);
      sheet.getRange(1, afterResult + 1).setValue('Result_Value');
    }
  }

  // Ensure rating pregame and rating change columns exist near ratings
  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  var oppRatingIdx = headers.indexOf('Opponent rating');
  if (oppRatingIdx !== -1) {
    var oppPregameIdx = headers.indexOf('Opponent rating pregame');
    var oppChangeIdx = headers.indexOf('Opponent rating change');
    var oppPregameFormulaIdx = headers.indexOf('Opponent rating pregame_formula');
    var oppPregameSelectedIdx = headers.indexOf('Opponent rating pregame_selected');
    var oppPregameSourceIdx = headers.indexOf('Opponent rating pregame_source');
    var toInsertOpp = [];
    if (oppPregameIdx === -1) toInsertOpp.push('Opponent rating pregame');
    if (oppPregameFormulaIdx === -1) toInsertOpp.push('Opponent rating pregame_formula');
    if (oppPregameSelectedIdx === -1) toInsertOpp.push('Opponent rating pregame_selected');
    if (oppPregameSourceIdx === -1) toInsertOpp.push('Opponent rating pregame_source');
    if (oppChangeIdx === -1) toInsertOpp.push('Opponent rating change');
    if (toInsertOpp.length) {
      var afterOpp = oppRatingIdx + 1;
      sheet.insertColumnsAfter(afterOpp, toInsertOpp.length);
      sheet.getRange(1, afterOpp + 1, 1, toInsertOpp.length).setValues([toInsertOpp]);
    }
  }

  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  var myRatingIdx = headers.indexOf('My rating');
  if (myRatingIdx !== -1) {
    var myPregameIdx = headers.indexOf('My rating pregame');
    var myChangeIdx = headers.indexOf('My rating change');
    var myPregameDerivedIdx = headers.indexOf('My rating pregame_derived');
    var myPregameFormulaIdx = headers.indexOf('My rating pregame_formula');
    var myPregameSelectedIdx = headers.indexOf('My rating pregame_selected');
    var myPregameSourceIdx = headers.indexOf('My rating pregame_source');
    var toInsertMy = [];
    if (myPregameIdx === -1) toInsertMy.push('My rating pregame');
    if (myPregameDerivedIdx === -1) toInsertMy.push('My rating pregame_derived');
    if (myPregameFormulaIdx === -1) toInsertMy.push('My rating pregame_formula');
    if (myPregameSelectedIdx === -1) toInsertMy.push('My rating pregame_selected');
    if (myPregameSourceIdx === -1) toInsertMy.push('My rating pregame_source');
    if (myChangeIdx === -1) toInsertMy.push('My rating change');
    if (toInsertMy.length) {
      var afterMy = myRatingIdx + 1;
      sheet.insertColumnsAfter(afterMy, toInsertMy.length);
      sheet.getRange(1, afterMy + 1, 1, toInsertMy.length).setValues([toInsertMy]);
    }
  }

  // Ensure audit column for rating delta mismatch exists
  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  if (headers.indexOf('Rating delta mismatch') === -1) {
    var lastColNow = sheet.getLastColumn();
    sheet.insertColumnsAfter(lastColNow, 1);
    sheet.getRange(1, lastColNow + 1).setValue('Rating delta mismatch');
  }

  // Ensure additional callback-derived columns exist
  updatedLastCol = sheet.getLastColumn();
  headers = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
  var extraHeaders = [
    'White rating change',
    'Black rating change',
    'Opponent membership code',
    'Opponent membership level',
    'Opponent country',
    'Opponent avatar URL',
    'My membership code',
    'My membership level'
  ];
  var missing = [];
  for (var i = 0; i < extraHeaders.length; i++) {
    if (headers.indexOf(extraHeaders[i]) === -1) missing.push(extraHeaders[i]);
  }
  if (missing.length) {
    var last2 = sheet.getLastColumn();
    sheet.insertColumnsAfter(last2, missing.length);
    sheet.getRange(1, last2 + 1, 1, missing.length).setValues([missing]);
  }

  // Remove redundant/duplicate columns if they exist
  var headersToRemove = [
    'Winner color',
    'Game end reason',
    'Result message',
    'Move timestamps',
    'Move list',
    'Last move',
    'Rating change',
    'Is live game',
    'Is abortable',
    'Is analyzable',
    'Is resignable',
    'Is checkmate',
    'Is stalemate',
    'Is finished',
    'Can send trophy',
    'Changes players rating',
    'Allow vacation',
    'Game UUID',
    'Turn color',
    'Ply count',
    'Initial setup',
    'Type name'
  ];
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var removeCols = [];
  for (var r = 0; r < headersToRemove.length; r++) {
    var idx = headers.indexOf(headersToRemove[r]);
    if (idx !== -1) removeCols.push(idx + 1); // 1-based
  }
  if (removeCols.length) {
    // Sort descending to avoid shifting indexes while deleting
    removeCols.sort(function(a, b) { return b - a; });
    for (var d = 0; d < removeCols.length; d++) {
      sheet.deleteColumn(removeCols[d]);
    }
  }

  return sheet;
}

function loadExistingUrls(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  // URL is column 2 (B)
  const urlValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const set = new Set();
  for (var i = 0; i < urlValues.length; i++) {
    var v = urlValues[i][0];
    if (v) set.add(String(v));
  }
  return set;
}

function appendRows(sheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  const numCols = rows[0].length;
  sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
}

function buildRow(details) {
  return [
    details.timestamp,                // Timestamp
    details.url,                      // URL
    details.gameId || '',             // Game ID
    details.rated ? 'Yes' : 'No',     // Rated
    details.timeClass,                // Time class
    details.timeControl,              // Time control
    (details.baseSeconds != null ? details.baseSeconds : ''), // Base time (s)
    (details.incrementSeconds != null ? details.incrementSeconds : ''), // Increment (s)
    details.rules,                    // Rules
    details.format || '',             // Format
    details.color,                    // Color
    details.opponentUsername,         // Opponent
    details.opponentRating || '',     // Opponent rating
    '',                               // Opponent rating pregame (backfilled)
    '',                               // Opponent rating change (backfilled)
    details.myRating || '',           // My rating
    '',                               // My rating pregame (backfilled)
    '',                               // My rating pregame_derived (backfilled)
    '',                               // My rating change (backfilled)
    details.result,                   // Result
    (details.resultValue != null ? details.resultValue : ''), // Result_Value
    details.reason,                   // Reason
    details.fen,                      // FEN
    details.imageUrl,                 // Endboard URL
    details.eco || '',                // ECO
    details.openingUrl || '',         // Opening URL
    details.movesSan || '',           // Moves (SAN list)
    details.clocks || '',             // Clocks list
    details.clockSeconds || '',       // Clock Seconds
    details.moveTimes || ''           // Move Times
  ];
}

function normalizeGame(game) {
  const playerLower = String(PLAYER_USERNAME || '').toLowerCase();

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

  // Parse PGN tags for ECO (code only) and Opening URL distinctly
  var eco = '';
  var openingUrl = '';
  var movesSan = '';
  var clocksStr = '';
  var parsedMoves = [];
  if (game.pgn) {
    try {
      var pgn = String(game.pgn).replace(/\r\n/g, '\n');
      // Extract tags like [ECO "B01"] and [ECOUrl "..."] anchored per line
      var ecoMatch = pgn.match(/^\[ECO\s+"([^"]+)"\]/m);
      if (ecoMatch && ecoMatch[1]) eco = ecoMatch[1];
      var ecoUrlMatch = pgn.match(/^\[(?:ECOUrl|OpeningUrl)\s+"([^"]+)"\]/m);
      if (ecoUrlMatch && ecoUrlMatch[1]) openingUrl = ecoUrlMatch[1];
      // Also derive SAN moves and Clocks from the PGN movetext for normal flow
      var sanClockParsed = parsePgnToSanAndClocks(pgn);
      if (sanClockParsed) {
        movesSan = sanClockParsed.movesSan || movesSan;
        clocksStr = sanClockParsed.clocks || clocksStr;
      }
      var movetext = extractMovetextFromPgn(pgn);
      parsedMoves = movetext ? splitMovesWithClocks(movetext) : [];
    } catch (e) {
      // ignore PGN parse issues
    }
  }
  // Fallback to API fields if PGN lacked them
  if (!eco && game.eco && !/^https?:\/\//i.test(String(game.eco))) eco = String(game.eco);
  if (!openingUrl && game.opening_url) openingUrl = String(game.opening_url);
  // If ECO accidentally contains a URL, move it to Opening URL and clear ECO
  if (eco && /^https?:\/\//i.test(eco)) {
    if (!openingUrl) openingUrl = eco;
    eco = '';
  }

  // Derive format per rules
  if (rules === 'chess') {
    format = timeClass || '';
  } else if (rules === 'chess960') {
    format = (timeClass === 'daily') ? 'daily 960' : 'live960';
  } else {
    format = rules || '';
  }

  // Extract game ID from URL (numbers after the last '/')
  var gameId = '';
  if (game.url) {
    var idMatch = String(game.url).match(/\/(\d+)(?:\?.*)?$/);
    if (idMatch && idMatch[1]) gameId = idMatch[1];
  }

  // Time control-derived base and increment
  var tcParsedFinal = parseTimeControlToBaseInc(timeControl, timeClass);
  var baseSecondsFinal = tcParsedFinal.baseSeconds;
  var incSecondsFinal = tcParsedFinal.incrementSeconds;
  // Derived strings
  var derivedFinal = buildClockSecondsAndMoveTimes(parsedMoves, baseSecondsFinal, incSecondsFinal);
  var clockSecondsStrFinal = derivedFinal.clockSecondsStr;
  var moveTimesStrFinal = derivedFinal.moveTimesStr;

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
    result: result,
    resultValue: resultValue,
    reason: reason,
    fen: fen,
    imageUrl: imageUrl,
    eco: eco,
    openingUrl: openingUrl,
    movesSan: movesSan,
    clocks: clocksStr,
    baseSeconds: baseSecondsFinal,
    incrementSeconds: incSecondsFinal,
    clockSeconds: clockSecondsStrFinal,
    moveTimes: moveTimesStrFinal
  };
}

function determineResultValue(result) {
  if (result === 'won') return 1;
  if (result === 'drew') return 0.5;
  if (result === 'lost') return 0;
  return '';
}

function determineResult(side, opponent) {
  if (!side || !opponent) return 'unknown';
  if (side.result === 'win') return 'won';
  if (opponent.result === 'win') return 'lost';

  var drawResults = {
    'agreed': true,
    'repetition': true,
    'stalemate': true,
    'insufficient': true,
    '50move': true,
    'timevsinsufficient': true
  };
  if (drawResults[side.result] || drawResults[opponent.result]) return 'drew';

  return 'unknown';
}

function determineReason(side, opponent) {
  if (!side || !opponent) return 'unknown';

  var reasons = [
    ['checkmated', 'checkmate'],
    ['agreed', 'agreement'],
    ['repetition', 'repetition'],
    ['timeout', 'timeout'],
    ['resigned', 'resignation'],
    ['stalemate', 'stalemate'],
    ['insufficient', 'insufficient material'],
    ['50move', '50-move rule'],
    ['abandoned', 'abandonment'],
    ['kingofthehill', 'opponent king reached the hill'],
    ['threecheck', 'three check'],
    ['timevsinsufficient', 'timeout vs insufficient material'],
    ['bughousepartnerlose', 'bughouse partner lost']
  ];

  for (var i = 0; i < reasons.length; i++) {
    var code = reasons[i][0];
    var text = reasons[i][1];
    if (side.result === code || opponent.result === code) return text;
  }
  return 'unknown';
}

// Optional: helper to create a time-driven trigger for syncRecentGames()
function installTriggerEvery15Minutes() {
  ScriptApp.newTrigger('syncRecentGames')
    .timeBased()
    .everyMinutes(TRIGGER_INTERVAL_MINUTES)
    .create();
}

/**
 * Ensure the header row contains "Moves (SAN)" and "Clocks" columns.
 * If an older header had a single "Moves" column, convert it to "Moves (SAN)"
 * and insert a new "Clocks" column after it.
 */
// ensureMovesClocksHeaders has been superseded by ensureSheet()

/**
 * Fetches live game callback data for a given Chess.com gameId and tries to
 * obtain a PGN string. Returns empty string if not found.
 * @param {string} gameId
 * @return {string}
 */
function fetchCallbackGamePgn(gameId) {
  if (!gameId) return '';
  var url = CHESS_COM_CALLBACK_BASE + '/live/game/' + encodeURIComponent(String(gameId));
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return '';
  var text = resp.getContentText();
  if (!text) return '';
  try {
    var data = JSON.parse(text);
    if (data && data.pgn) return String(data.pgn);
    if (data && data.game && data.game.pgn) return String(data.game.pgn);
  } catch (e) {
    // Not JSON; fall through to heuristic PGN extraction from text
  }
  // Heuristic: look for typical PGN start and end markers in raw text
  var pgnMatch = text.match(/\[Event[^]*?(?:\n\n|\r\n\r\n)[^]*$/);
  if (pgnMatch && pgnMatch[0]) return String(pgnMatch[0]);
  return '';
}

/**
 * Fetches the raw callback JSON for a given Chess.com gameId.
 * @param {string} gameId
 * @return {Object|null}
 */
function fetchCallbackGameData(gameId) {
  if (!gameId) return null;
  var url = CHESS_COM_CALLBACK_BASE + '/live/game/' + encodeURIComponent(String(gameId));
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;
  var text = resp.getContentText();
  if (!text) return null;
  try {
    var data = JSON.parse(text);
    return data && typeof data === 'object' ? data : null;
  } catch (e) {
    return null;
  }
}

/**
 * Given a PGN string, extract movetext and return bracketed lists:
 * {san1, san2, ...} and {clk1, clk2, ...}
 * @param {string} pgn
 * @return {{ movesSan: string, clocks: string }}
 */
function parsePgnToSanAndClocks(pgn) {
  var movesSan = '';
  var clocksStr = '';
  if (!pgn) return { movesSan: movesSan, clocks: clocksStr };

  var normalized = String(pgn).replace(/\r\n/g, '\n');
  var headerEnd = normalized.indexOf('\n\n');
  var body = '';
  if (headerEnd !== -1) {
    body = normalized.substring(headerEnd + 2);
  } else {
    var splitIndex = normalized.lastIndexOf(']\n');
    if (splitIndex !== -1) body = normalized.substring(splitIndex + 2);
  }
  if (!body) return { movesSan: movesSan, clocks: clocksStr };

  body = body.replace(/^\s+/, '');
  body = body.replace(/\s+(1-0|0-1|1\/2-1\/2|\*)\s*$/m, '');

  var parsed = splitMovesWithClocks(body);
  if (parsed && parsed.length) {
    var sanList = parsed.map(function(m) { return m.san; });
    var clkList = parsed.map(function(m) { return m.clock; });
    movesSan = sanList.length ? '{' + sanList.join(', ') + '}' : '';
    clocksStr = clkList.length ? '{' + clkList.join(', ') + '}' : '';
  }
  return { movesSan: movesSan, clocks: clocksStr };
}

/**
 * Extracts movetext (body after headers) from a PGN string.
 * @param {string} pgn
 * @return {string}
 */
function extractMovetextFromPgn(pgn) {
  if (!pgn) return '';
  var normalized = String(pgn).replace(/\r\n/g, '\n');
  var headerEnd = normalized.indexOf('\n\n');
  if (headerEnd !== -1) return normalized.substring(headerEnd + 2);
  var splitIndex = normalized.lastIndexOf(']\n');
  if (splitIndex !== -1) return normalized.substring(splitIndex + 2);
  return normalized;
}

/**
 * Parse Time control to base and increment seconds, per rules:
 * - If time control contains '/', treat as daily; return nulls (skip base/inc).
 * - If time class is 'daily', also skip base/inc.
 * - If value contains '+', base is before '+', increment is after '+', both in seconds.
 * - Else, entire value is base seconds and increment is 0.
 * @param {string} timeControl
 * @param {string} timeClass
 * @return {{baseSeconds: (number|null), incrementSeconds: (number|null)}}
 */
function parseTimeControlToBaseInc(timeControl, timeClass) {
  var tc = String(timeControl || '').trim();
  var cls = String(timeClass || '').trim().toLowerCase();
  if (!tc) return { baseSeconds: null, incrementSeconds: null };
  if (cls === 'daily' || tc.indexOf('/') !== -1) {
    return { baseSeconds: null, incrementSeconds: null };
  }
  var base = null;
  var inc = null;
  if (tc.indexOf('+') !== -1) {
    var parts = tc.split('+');
    base = toIntSafe(parts[0]);
    inc = toIntSafe(parts[1]);
  } else {
    base = toIntSafe(tc);
    inc = 0;
  }
  return {
    baseSeconds: isFinite(base) ? base : null,
    incrementSeconds: isFinite(inc) ? inc : null
  };
}

/**
 * Convert string to integer safely.
 * @param {string} s
 * @return {number}
 */
function toIntSafe(s) {
  var n = parseInt(String(s).replace(/[^0-9\-]/g, ''), 10);
  return isFinite(n) ? n : NaN;
}

/**
 * Build bracketed lists for Clock Seconds and Move Times.
 * Clock Seconds mirrors the Clocks list but in seconds (numbers).
 * Move Times is computed per side using base/increment rules:
 *  - White first move: base - clock(white 1st) + inc, if base/inc known
 *  - Subsequent same-side move: prevClock(side) - currClock(side) + inc
 * If base/inc are null (daily), produce empty string for Move Times. Always
 * produce Clock Seconds if clocks are available.
 * @param {Array<{color:string, clockSeconds:number}>} parsedMoves
 * @param {number|null} baseSeconds
 * @param {number|null} incSeconds
 * @return {{clockSecondsStr: string, moveTimesStr: string}}
 */
function buildClockSecondsAndMoveTimes(parsedMoves, baseSeconds, incSeconds) {
  if (!parsedMoves || !parsedMoves.length) return { clockSecondsStr: '', moveTimesStr: '' };
  var clkSecs = parsedMoves.map(function(m) {
    var v = m && isFinite(m.clockSeconds) ? m.clockSeconds : NaN;
    return isFinite(v) ? v : '';
  });
  var clockSecondsStr = clkSecs.length ? '{' + clkSecs.join(', ') + '}' : '';

  var canComputeMoves = (baseSeconds != null && incSeconds != null);
  if (!canComputeMoves) return { clockSecondsStr: clockSecondsStr, moveTimesStr: '' };

  var lastClockByColor = { white: null, black: null };
  var haveSeenFirstByColor = { white: false, black: false };
  var moveDurations = [];
  for (var i = 0; i < parsedMoves.length; i++) {
    var m = parsedMoves[i];
    var clr = m.color === 'black' ? 'black' : 'white';
    var curr = isFinite(m.clockSeconds) ? m.clockSeconds : NaN;
    var duration = NaN;
    if (!haveSeenFirstByColor[clr]) {
      if (isFinite(baseSeconds) && isFinite(curr) && isFinite(incSeconds)) {
        duration = (baseSeconds - curr + incSeconds);
      }
      haveSeenFirstByColor[clr] = true;
    } else {
      var prev = lastClockByColor[clr];
      if (isFinite(prev) && isFinite(curr) && isFinite(incSeconds)) {
        duration = (prev - curr + incSeconds);
      }
    }
    if (isFinite(duration)) {
      var rounded = Math.round(duration * 100) / 100;
      moveDurations.push(rounded);
    } else {
      moveDurations.push('');
    }
    lastClockByColor[clr] = isFinite(curr) ? curr : lastClockByColor[clr];
  }
  var moveTimesStr = moveDurations.length ? '{' + moveDurations.join(', ') + '}' : '';
  return { clockSecondsStr: clockSecondsStr, moveTimesStr: moveTimesStr };
}

/**
 * Backfills missing "Moves (SAN)" and "Clocks" for existing rows by
 * querying https://www.chess.com/callback/live/game/{game-id}.
 * Optionally provide a limit of how many missing rows to process this run.
 *
 * @param {number=} limit Optional max number of rows to process
 */
function backfillMovesAndClocks(limit) {
  const sheet = ensureSheet();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Build header map to locate columns robustly
  const lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerToIndex = {};
  for (var c = 0; c < headers.length; c++) {
    headerToIndex[headers[c]] = c + 1; // 1-based
  }

  var colGameId = headerToIndex['Game ID'] || 3;
  var colUrl = headerToIndex['URL'] || 2;
  var colMovesSan = headerToIndex['Moves (SAN)'];
  var colClocks = headerToIndex['Clocks'];
  var colClockSeconds = headerToIndex['Clock Seconds'];
  var colMoveTimes = headerToIndex['Move Times'];
  var colTimeClass = headerToIndex['Time class'];
  var colTimeControl = headerToIndex['Time control'];
  var colBase = headerToIndex['Base time (s)'];
  var colInc = headerToIndex['Increment (s)'];
  if (!colMovesSan || !colClocks) return; // headers must exist

  var numRows = lastRow - 1;
  var gameIds = sheet.getRange(2, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(2, colUrl, numRows, 1).getValues();
  var movesSanVals = sheet.getRange(2, colMovesSan, numRows, 1).getValues();
  var clocksVals = sheet.getRange(2, colClocks, numRows, 1).getValues();
  var clockSecondsVals = colClockSeconds ? sheet.getRange(2, colClockSeconds, numRows, 1).getValues() : null;
  var moveTimesVals = colMoveTimes ? sheet.getRange(2, colMoveTimes, numRows, 1).getValues() : null;
  var timeClassVals = colTimeClass ? sheet.getRange(2, colTimeClass, numRows, 1).getValues() : null;
  var timeControlVals = colTimeControl ? sheet.getRange(2, colTimeControl, numRows, 1).getValues() : null;
  var baseVals = colBase ? sheet.getRange(2, colBase, numRows, 1).getValues() : null;
  var incVals = colInc ? sheet.getRange(2, colInc, numRows, 1).getValues() : null;

  var toProcessCount = 0;
  for (var r = 0; r < numRows; r++) {
    var hasMoves = String(movesSanVals[r][0] || '').trim() !== '';
    var hasClocks = String(clocksVals[r][0] || '').trim() !== '';
    if (!hasMoves || !hasClocks) toProcessCount++;
  }
  var effectiveLimit = toProcessCount;
  if (typeof limit === 'number' && isFinite(limit) && limit >= 0) {
    effectiveLimit = Math.min(toProcessCount, limit);
  } else if (typeof BACKFILL_MOVES_CLOCKS_BATCH === 'number' && BACKFILL_MOVES_CLOCKS_BATCH > 0) {
    effectiveLimit = Math.min(toProcessCount, BACKFILL_MOVES_CLOCKS_BATCH);
  }

  for (var i = 0, processed = 0; i < numRows && processed < effectiveLimit; i++) {
    var currentMoves = String(movesSanVals[i][0] || '').trim();
    var currentClocks = String(clocksVals[i][0] || '').trim();
    if (currentMoves && currentClocks) continue;

    var gid = String(gameIds[i][0] || '').trim();
    if (!gid) {
      var url = String(urls[i][0] || '');
      var idMatch = url.match(/\/(\d+)(?:\?.*)?$/);
      if (idMatch && idMatch[1]) gid = idMatch[1];
    }
    if (!gid) continue;

    var pgn = fetchCallbackGamePgn(gid);
    if (!pgn) continue;
    var parsed = parsePgnToSanAndClocks(pgn);
    var newMovesSan = parsed.movesSan || currentMoves;
    var newClocks = parsed.clocks || currentClocks;
    movesSanVals[i][0] = newMovesSan;
    clocksVals[i][0] = newClocks;

    // Compute derived fields if columns exist
    if (clockSecondsVals || moveTimesVals || baseVals || incVals) {
      var timeClass = timeClassVals ? String(timeClassVals[i][0] || '') : '';
      var timeControl = timeControlVals ? String(timeControlVals[i][0] || '') : '';
      var tcParsed = parseTimeControlToBaseInc(timeControl, timeClass);
      var baseSeconds = tcParsed.baseSeconds;
      var incSeconds = tcParsed.incrementSeconds;
      if (baseVals) baseVals[i][0] = (baseSeconds != null ? baseSeconds : baseVals[i][0]);
      if (incVals) incVals[i][0] = (incSeconds != null ? incSeconds : incVals[i][0]);

      var movetext = extractMovetextFromPgn(pgn);
      var parsedMoves = movetext ? splitMovesWithClocks(movetext) : [];
      var derived = buildClockSecondsAndMoveTimes(parsedMoves, baseSeconds, incSeconds);
      if (clockSecondsVals) clockSecondsVals[i][0] = derived.clockSecondsStr || clockSecondsVals[i][0];
      if (moveTimesVals) moveTimesVals[i][0] = derived.moveTimesStr || moveTimesVals[i][0];
    }
    processed++;
  }

  // Batch update the entire two columns at once
  sheet.getRange(2, colMovesSan, numRows, 1).setValues(movesSanVals);
  sheet.getRange(2, colClocks, numRows, 1).setValues(clocksVals);
  if (colClockSeconds && clockSecondsVals) sheet.getRange(2, colClockSeconds, numRows, 1).setValues(clockSecondsVals);
  if (colMoveTimes && moveTimesVals) sheet.getRange(2, colMoveTimes, numRows, 1).setValues(moveTimesVals);
  if (colBase && baseVals) sheet.getRange(2, colBase, numRows, 1).setValues(baseVals);
  if (colInc && incVals) sheet.getRange(2, colInc, numRows, 1).setValues(incVals);
  return processed;
}

/**
 * Extract ECO code and Opening URL from a PGN string.
 * @param {string} pgn
 * @return {{ eco: string, openingUrl: string }}
 */
function parseEcoAndOpeningFromPgn(pgn) {
  var eco = '';
  var openingUrl = '';
  if (!pgn) return { eco: eco, openingUrl: openingUrl };
  try {
    var normalized = String(pgn).replace(/\r\n/g, '\n');
    var ecoMatch = normalized.match(/^\[ECO\s+"([^"]+)"\]/m);
    if (ecoMatch && ecoMatch[1]) eco = ecoMatch[1];
    var urlMatch = normalized.match(/^\[(?:ECOUrl|OpeningUrl)\s+"([^"]+)"\]/m);
    if (urlMatch && urlMatch[1]) openingUrl = urlMatch[1];
  } catch (e) {}
  return { eco: eco, openingUrl: openingUrl };
}

/**
 * Backfills missing ECO and Opening URL for existing rows by querying
 * https://www.chess.com/callback/live/game/{game-id}. Will not overwrite
 * non-empty values unless ECO appears to be a URL (bad old data).
 *
 * @param {number=} limit Optional max number of rows to process this run
 */
function backfillEcoAndOpening(limit) {
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerToIndex = {};
  for (var c = 0; c < headers.length; c++) headerToIndex[headers[c]] = c + 1;

  var colGameId = headerToIndex['Game ID'] || 3;
  var colUrl = headerToIndex['URL'] || 2;
  var colEco = headerToIndex['ECO'];
  var colOpeningUrl = headerToIndex['Opening URL'];
  if (!colEco || !colOpeningUrl) return;

  var numRows = lastRow - 1;
  var gameIds = sheet.getRange(2, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(2, colUrl, numRows, 1).getValues();
  var ecos = sheet.getRange(2, colEco, numRows, 1).getValues();
  var openingUrls = sheet.getRange(2, colOpeningUrl, numRows, 1).getValues();

  var candidates = [];
  for (var r = 0; r < numRows; r++) {
    var ecoVal = String(ecos[r][0] || '').trim();
    var openVal = String(openingUrls[r][0] || '').trim();
    var ecoLooksWrong = ecoVal && /^https?:\/\//i.test(ecoVal);
    if (!ecoVal || !openVal || ecoLooksWrong) candidates.push(r);
  }

  if (typeof limit === 'number' && isFinite(limit) && limit >= 0) {
    candidates = candidates.slice(0, limit);
  } else if (typeof BACKFILL_ECO_OPENING_BATCH === 'number' && BACKFILL_ECO_OPENING_BATCH > 0) {
    candidates = candidates.slice(0, BACKFILL_ECO_OPENING_BATCH);
  }

  var processed = 0;
  for (var i = 0; i < candidates.length; i++) {
    var idx = candidates[i];
    var gid = String(gameIds[idx][0] || '').trim();
    if (!gid) {
      var url = String(urls[idx][0] || '');
      var idMatch = url.match(/\/(\d+)(?:\?.*)?$/);
      if (idMatch && idMatch[1]) gid = idMatch[1];
    }
    if (!gid) continue;

    var pgn = fetchCallbackGamePgn(gid);
    if (!pgn) continue;
    var parsed = parseEcoAndOpeningFromPgn(pgn);

    // Only write when empty or obviously wrong
    var wrote = false;
    if (!String(ecos[idx][0] || '').trim() || /^https?:\/\//i.test(String(ecos[idx][0] || ''))) {
      var newEco = parsed.eco || ecos[idx][0];
      if (newEco !== ecos[idx][0]) { ecos[idx][0] = newEco; wrote = true; }
    }
    if (!String(openingUrls[idx][0] || '').trim()) {
      var newUrl = parsed.openingUrl || openingUrls[idx][0];
      if (newUrl !== openingUrls[idx][0]) { openingUrls[idx][0] = newUrl; wrote = true; }
    }
    if (wrote) processed++;
  }

  sheet.getRange(2, colEco, numRows, 1).setValues(ecos);
  sheet.getRange(2, colOpeningUrl, numRows, 1).setValues(openingUrls);
  return processed;
}

/**
 * Estimate pregame ratings using Elo with fixed K, based on postgame ratings and result.
 * Writes to: My rating pregame_formula, Opponent rating pregame_formula.
 * Does not overwrite existing non-empty values.
 */
function backfillPregameFormulaInRange(sheet, startRow, numRows, K) {
  var headerToIndex = getHeaderToIndexMap_(sheet);
  var colMyRating = headerToIndex['My rating'];
  var colOppRating = headerToIndex['Opponent rating'];
  var colResult = headerToIndex['Result'];
  var colMyPregameF = headerToIndex['My rating pregame_formula'];
  var colOppPregameF = headerToIndex['Opponent rating pregame_formula'];
  if (!colMyRating || !colOppRating || !colResult || !colMyPregameF || !colOppPregameF) return 0;

  function init(col) { return sheet.getRange(startRow, col, numRows, 1).getValues(); }
  var myPostVals = init(colMyRating);
  var oppPostVals = init(colOppRating);
  var resultVals = init(colResult);
  var myPreFVals = init(colMyPregameF);
  var oppPreFVals = init(colOppPregameF);

  function resultToScore(r) {
    var s = String(r || '').toLowerCase();
    if (s === 'won') return 1;
    if (s === 'drew') return 0.5;
    if (s === 'lost') return 0;
    return null;
  }

  function estimate(RpostMe, RpostOpp, S, Kval) {
    var sumPost = RpostMe + RpostOpp;
    var d = RpostOpp - RpostMe; // start from post diff
    for (var i = 0; i < 25; i++) {
      var E = 1 / (1 + Math.pow(10, d / 400));
      var RpreMe = RpostMe - Kval * (S - E);
      var RpreOpp = sumPost - RpreMe;
      var newD = RpreOpp - RpreMe;
      if (Math.abs(newD - d) < 0.01) { d = newD; break; }
      d = newD;
    }
    var E2 = 1 / (1 + Math.pow(10, d / 400));
    var RpreMe2 = RpostMe - Kval * (S - E2);
    var RpreOpp2 = sumPost - RpreMe2;
    return { my: Math.round(RpreMe2), opp: Math.round(RpreOpp2) };
  }

  var wrote = 0;
  for (var i = 0; i < numRows; i++) {
    var haveMy = String(myPreFVals[i][0] || '').trim() !== '';
    var haveOpp = String(oppPreFVals[i][0] || '').trim() !== '';
    if (haveMy && haveOpp) continue;

    var Rm = Number(myPostVals[i][0]);
    var Ro = Number(oppPostVals[i][0]);
    var S = resultToScore(resultVals[i][0]);
    if (!isFinite(Rm) || !isFinite(Ro) || S == null) continue;

    var est = estimate(Rm, Ro, S, K);
    if (!haveMy) { myPreFVals[i][0] = est.my; wrote++; }
    if (!haveOpp) { oppPreFVals[i][0] = est.opp; wrote++; }
  }

  if (wrote > 0) {
    sheet.getRange(startRow, colMyPregameF, numRows, 1).setValues(myPreFVals);
    sheet.getRange(startRow, colOppPregameF, numRows, 1).setValues(oppPreFVals);
  }
  return wrote;
}

/**
 * Backfills various fields from the callback JSON for rows.
 * Populates: rating deltas, winner, end reason/message, timestamps, movelist,
 * last move, base/increment, flags, identifiers, turn color, ply count,
 * initial setup, type name, opponent/my membership and opponent metadata.
 *
 * @param {number=} limit Optional max number of rows to process this run
 */
function backfillCallbackFields(limit) {
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Header map
  const lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerToIndex = {};
  for (var c = 0; c < headers.length; c++) headerToIndex[headers[c]] = c + 1;

  var colGameId = headerToIndex['Game ID'] || 3;
  var colUrl = headerToIndex['URL'] || 2;
  var colColor = headerToIndex['Color'];
  var colMyRating = headerToIndex['My rating'];
  var colOppRating = headerToIndex['Opponent rating'];
  var colResult = headerToIndex['Result'];
  var colResultValue = headerToIndex['Result_Value'];

  // Resolve new columns
  var colWhiteDelta = headerToIndex['White rating change'];
  var colBlackDelta = headerToIndex['Black rating change'];
  var colMyDelta = headerToIndex['My rating change'];
  var colOppDelta = headerToIndex['Opponent rating change'];
  var colMyPregame = headerToIndex['My rating pregame'];
  var colOppPregame = headerToIndex['Opponent rating pregame'];
  var colMyPregameDerived = headerToIndex['My rating pregame_derived'];
  var colMyPregameFormula = headerToIndex['My rating pregame_formula'];
  var colOppPregameFormula = headerToIndex['Opponent rating pregame_formula'];
  var colMyPregameSelected = headerToIndex['My rating pregame_selected'];
  var colMyPregameSource = headerToIndex['My rating pregame_source'];
  var colOppPregameSelected = headerToIndex['Opponent rating pregame_selected'];
  var colOppPregameSource = headerToIndex['Opponent rating pregame_source'];
  var colDeltaMismatch = headerToIndex['Rating delta mismatch'];
  var colWinnerColor = headerToIndex['Winner color'];
  var colEndReason = headerToIndex['Game end reason'];
  var colResultMsg = headerToIndex['Result message'];
  var colMoveTimestamps = headerToIndex['Move timestamps'];
  var colMoveList = headerToIndex['Move list'];
  var colLastMove = headerToIndex['Last move'];
  var colBase = headerToIndex['Base time (s)'];
  var colInc = headerToIndex['Increment (s)'];
  var colIsLive = headerToIndex['Is live game'];
  var colIsAbortable = headerToIndex['Is abortable'];
  var colIsAnalyzable = headerToIndex['Is analyzable'];
  var colIsResignable = headerToIndex['Is resignable'];
  var colIsCheckmate = headerToIndex['Is checkmate'];
  var colIsStalemate = headerToIndex['Is stalemate'];
  var colIsFinished = headerToIndex['Is finished'];
  var colCanSendTrophy = headerToIndex['Can send trophy'];
  var colChangesPlayersRating = headerToIndex['Changes players rating'];
  var colAllowVacation = headerToIndex['Allow vacation'];
  var colUuid = headerToIndex['Game UUID'];
  var colTurnColor = headerToIndex['Turn color'];
  var colPlyCount = headerToIndex['Ply count'];
  var colInitialSetup = headerToIndex['Initial setup'];
  var colTypeName = headerToIndex['Type name'];
  var colOppMembershipCode = headerToIndex['Opponent membership code'];
  var colOppMembershipLevel = headerToIndex['Opponent membership level'];
  var colOppCountry = headerToIndex['Opponent country'];
  var colOppAvatar = headerToIndex['Opponent avatar URL'];
  var colMyMembershipCode = headerToIndex['My membership code'];
  var colMyMembershipLevel = headerToIndex['My membership level'];

  var numRows = lastRow - 1;
  var gameIds = sheet.getRange(2, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(2, colUrl, numRows, 1).getValues();
  var colors = colColor ? sheet.getRange(2, colColor, numRows, 1).getValues() : null;
  var myRatings = colMyRating ? sheet.getRange(2, colMyRating, numRows, 1).getValues() : null;
  var oppRatings = colOppRating ? sheet.getRange(2, colOppRating, numRows, 1).getValues() : null;
  var resultVals = colResult ? sheet.getRange(2, colResult, numRows, 1).getValues() : null;

  // Prepare arrays for batch update where columns exist
  function initColumn(col) {
    return col ? sheet.getRange(2, col, numRows, 1).getValues() : null;
  }
  var whiteDeltaVals = initColumn(colWhiteDelta);
  var blackDeltaVals = initColumn(colBlackDelta);
  var myDeltaVals = initColumn(colMyDelta);
  var oppDeltaVals = initColumn(colOppDelta);
  var myPregameVals = initColumn(colMyPregame);
  var oppPregameVals = initColumn(colOppPregame);
  var myPregameDerivedVals = initColumn(colMyPregameDerived);
  var myPregameFormulaVals = initColumn(colMyPregameFormula);
  var oppPregameFormulaVals = initColumn(colOppPregameFormula);
  var myPregameSelectedVals = initColumn(colMyPregameSelected);
  var myPregameSourceVals = initColumn(colMyPregameSource);
  var oppPregameSelectedVals = initColumn(colOppPregameSelected);
  var oppPregameSourceVals = initColumn(colOppPregameSource);
  var deltaMismatchVals = initColumn(colDeltaMismatch);
  var resultValueVals = initColumn(colResultValue);
  var winnerVals = initColumn(colWinnerColor);
  var endReasonVals = initColumn(colEndReason);
  var resultMsgVals = initColumn(colResultMsg);
  var tsVals = initColumn(colMoveTimestamps);
  var moveListVals = initColumn(colMoveList);
  var lastMoveVals = initColumn(colLastMove);
  var baseVals = initColumn(colBase);
  var incVals = initColumn(colInc);
  var isLiveVals = initColumn(colIsLive);
  var isAbortableVals = initColumn(colIsAbortable);
  var isAnalyzableVals = initColumn(colIsAnalyzable);
  var isResignableVals = initColumn(colIsResignable);
  var isCheckmateVals = initColumn(colIsCheckmate);
  var isStalemateVals = initColumn(colIsStalemate);
  var isFinishedVals = initColumn(colIsFinished);
  var canSendTrophyVals = initColumn(colCanSendTrophy);
  var changesPlayersRatingVals = initColumn(colChangesPlayersRating);
  var allowVacationVals = initColumn(colAllowVacation);
  var uuidVals = initColumn(colUuid);
  var turnVals = initColumn(colTurnColor);
  var plyVals = initColumn(colPlyCount);
  var setupVals = initColumn(colInitialSetup);
  var typeNameVals = initColumn(colTypeName);
  var oppMembershipCodeVals = initColumn(colOppMembershipCode);
  var oppMembershipLevelVals = initColumn(colOppMembershipLevel);
  var oppCountryVals = initColumn(colOppCountry);
  var oppAvatarVals = initColumn(colOppAvatar);
  var myMembershipCodeVals = initColumn(colMyMembershipCode);
  var myMembershipLevelVals = initColumn(colMyMembershipLevel);

  // Build candidate list: rows missing any targeted derived rating fields
  var candidates = [];
  for (var r = 0; r < numRows; r++) {
    var needs = false;
    if (myDeltaVals && String(myDeltaVals[r][0] || '').trim() === '') needs = true;
    if (!needs && oppDeltaVals && String(oppDeltaVals[r][0] || '').trim() === '') needs = true;
    if (!needs && myPregameVals && String(myPregameVals[r][0] || '').trim() === '') needs = true;
    if (!needs && oppPregameVals && String(oppPregameVals[r][0] || '').trim() === '') needs = true;
    if (!needs && resultValueVals && String(resultValueVals[r][0] || '').trim() === '') needs = true;
    if (needs) candidates.push(r);
  }

  if (typeof limit === 'number' && isFinite(limit) && limit >= 0) {
    candidates = candidates.slice(0, limit);
  } else if (typeof BACKFILL_CALLBACK_FIELDS_BATCH === 'number' && BACKFILL_CALLBACK_FIELDS_BATCH > 0) {
    candidates = candidates.slice(0, BACKFILL_CALLBACK_FIELDS_BATCH);
  }

  var processed = 0;
  for (var k = 0; k < candidates.length; k++) {
    var i = candidates[k];
    var gid = String(gameIds[i][0] || '').trim();
    if (!gid) {
      var url = String(urls[i][0] || '');
      var idMatch = url.match(/\/(\d+)(?:\?.*)?$/);
      if (idMatch && idMatch[1]) gid = idMatch[1];
    }
    if (!gid) continue;

    var data = fetchCallbackGameData(gid);
    if (!data || !data.game) continue;
    var g = data.game;
    var players = data.players || {};
    var top = players.top || {};
    var bottom = players.bottom || {};

    var wrote = false;
    if (ENABLE_CALLBACK_HEADERS_IN_FULL_BACKFILL) {
      if (whiteDeltaVals && String(whiteDeltaVals[i][0] || '').trim() === '' && (g.ratingChangeWhite != null)) { whiteDeltaVals[i][0] = g.ratingChangeWhite; wrote = true; }
      if (blackDeltaVals && String(blackDeltaVals[i][0] || '').trim() === '' && (g.ratingChangeBlack != null)) { blackDeltaVals[i][0] = g.ratingChangeBlack; wrote = true; }
    }
    if (winnerVals && String(winnerVals[i][0] || '').trim() === '' && g.colorOfWinner) { winnerVals[i][0] = g.colorOfWinner; wrote = true; }
    if (endReasonVals && String(endReasonVals[i][0] || '').trim() === '' && g.gameEndReason) { endReasonVals[i][0] = g.gameEndReason; wrote = true; }
    if (resultMsgVals && String(resultMsgVals[i][0] || '').trim() === '' && g.resultMessage) { resultMsgVals[i][0] = g.resultMessage; wrote = true; }
    if (tsVals && String(tsVals[i][0] || '').trim() === '' && g.moveTimestamps) { tsVals[i][0] = g.moveTimestamps; wrote = true; }
    if (moveListVals && String(moveListVals[i][0] || '').trim() === '' && g.moveList) { moveListVals[i][0] = g.moveList; wrote = true; }
    if (lastMoveVals && String(lastMoveVals[i][0] || '').trim() === '' && g.lastMove) { lastMoveVals[i][0] = g.lastMove; wrote = true; }
    if (baseVals && String(baseVals[i][0] || '').trim() === '' && (g.baseTime1 != null)) { baseVals[i][0] = g.baseTime1; wrote = true; }
    if (incVals && String(incVals[i][0] || '').trim() === '' && (g.timeIncrement1 != null)) { incVals[i][0] = g.timeIncrement1; wrote = true; }
    if (isLiveVals && String(isLiveVals[i][0] || '').trim() === '') { isLiveVals[i][0] = g.isLiveGame; wrote = true; }
    if (isAbortableVals && String(isAbortableVals[i][0] || '').trim() === '') { isAbortableVals[i][0] = g.isAbortable; wrote = true; }
    if (isAnalyzableVals && String(isAnalyzableVals[i][0] || '').trim() === '') { isAnalyzableVals[i][0] = g.isAnalyzable; wrote = true; }
    if (isResignableVals && String(isResignableVals[i][0] || '').trim() === '') { isResignableVals[i][0] = g.isResignable; wrote = true; }
    if (isCheckmateVals && String(isCheckmateVals[i][0] || '').trim() === '') { isCheckmateVals[i][0] = g.isCheckmate; wrote = true; }
    if (isStalemateVals && String(isStalemateVals[i][0] || '').trim() === '') { isStalemateVals[i][0] = g.isStalemate; wrote = true; }
    if (isFinishedVals && String(isFinishedVals[i][0] || '').trim() === '') { isFinishedVals[i][0] = g.isFinished; wrote = true; }
    if (canSendTrophyVals && String(canSendTrophyVals[i][0] || '').trim() === '') { canSendTrophyVals[i][0] = g.canSendTrophy; wrote = true; }
    if (changesPlayersRatingVals && String(changesPlayersRatingVals[i][0] || '').trim() === '') { changesPlayersRatingVals[i][0] = g.changesPlayersRating; wrote = true; }
    if (allowVacationVals && String(allowVacationVals[i][0] || '').trim() === '') { allowVacationVals[i][0] = g.allowVacation; wrote = true; }
    if (uuidVals && String(uuidVals[i][0] || '').trim() === '' && g.uuid) { uuidVals[i][0] = g.uuid; wrote = true; }
    if (turnVals && String(turnVals[i][0] || '').trim() === '' && g.turnColor) { turnVals[i][0] = g.turnColor; wrote = true; }
    if (plyVals && String(plyVals[i][0] || '').trim() === '' && (g.plyCount != null)) { plyVals[i][0] = g.plyCount; wrote = true; }
    if (setupVals && String(setupVals[i][0] || '').trim() === '' && (g.initialSetup != null)) { setupVals[i][0] = g.initialSetup; wrote = true; }
    if (typeNameVals && String(typeNameVals[i][0] || '').trim() === '' && g.typeName) { typeNameVals[i][0] = g.typeName; wrote = true; }

    // Opponent vs My membership based on player color in our row
    var ourColor = colors ? String(colors[i][0] || '').toLowerCase() : '';
    var opponentObj = (ourColor === 'white') ? top : bottom; // if we were white, opponent is top (black)
    var myObj = (ourColor === 'white') ? bottom : top;

    if (ENABLE_CALLBACK_HEADERS_IN_FULL_BACKFILL) {
      if (oppMembershipCodeVals && String(oppMembershipCodeVals[i][0] || '').trim() === '' && opponentObj.membershipCode) { oppMembershipCodeVals[i][0] = opponentObj.membershipCode; wrote = true; }
      if (oppMembershipLevelVals && String(oppMembershipLevelVals[i][0] || '').trim() === '' && (opponentObj.membershipLevel != null)) { oppMembershipLevelVals[i][0] = opponentObj.membershipLevel; wrote = true; }
      if (oppCountryVals && String(oppCountryVals[i][0] || '').trim() === '' && opponentObj.countryName) { oppCountryVals[i][0] = opponentObj.countryName; wrote = true; }
      if (oppAvatarVals && String(oppAvatarVals[i][0] || '').trim() === '' && opponentObj.avatarUrl) { oppAvatarVals[i][0] = opponentObj.avatarUrl; wrote = true; }
      if (myMembershipCodeVals && String(myMembershipCodeVals[i][0] || '').trim() === '' && myObj.membershipCode) { myMembershipCodeVals[i][0] = myObj.membershipCode; wrote = true; }
      if (myMembershipLevelVals && String(myMembershipLevelVals[i][0] || '').trim() === '' && (myObj.membershipLevel != null)) { myMembershipLevelVals[i][0] = myObj.membershipLevel; wrote = true; }
    }

    // Compute My/Opponent rating change and pregame ratings
    var myDelta = (ourColor === 'white') ? g.ratingChangeWhite : g.ratingChangeBlack;
    var oppDelta = (ourColor === 'white') ? g.ratingChangeBlack : g.ratingChangeWhite;
    if (myDeltaVals && String(myDeltaVals[i][0] || '').trim() === '' && (myDelta != null)) { myDeltaVals[i][0] = myDelta; wrote = true; }
    if (oppDeltaVals && String(oppDeltaVals[i][0] || '').trim() === '' && (oppDelta != null)) { oppDeltaVals[i][0] = oppDelta; wrote = true; }

    var myPost = myRatings ? Number(myRatings[i][0]) : NaN;
    var oppPost = oppRatings ? Number(oppRatings[i][0]) : NaN;
    var myPre = (isFinite(myPost) && isFinite(Number(myDelta))) ? (myPost - Number(myDelta)) : '';
    var oppPre = (isFinite(oppPost) && isFinite(Number(oppDelta))) ? (oppPost - Number(oppDelta)) : '';
    if (myPregameVals && String(myPregameVals[i][0] || '').trim() === '' && myPre !== '') { myPregameVals[i][0] = myPre; wrote = true; }
    if (oppPregameVals && String(oppPregameVals[i][0] || '').trim() === '' && oppPre !== '') { oppPregameVals[i][0] = oppPre; wrote = true; }

    // Compute selected pregame ratings and sources
    var myCb = myPregameVals ? Number(myPregameVals[i][0]) : NaN;
    var myLast = myPregameDerivedVals ? Number(myPregameDerivedVals[i][0]) : NaN;
    var myEst = myPregameFormulaVals ? Number(myPregameFormulaVals[i][0]) : NaN;
    var selMy = '';
    var srcMy = '';
    if (isFinite(myCb)) { selMy = myCb; srcMy = 'cb'; }
    else if (isFinite(myLast) && isFinite(myEst)) { selMy = (Math.abs(myLast - myEst) <= 3) ? myLast : myEst; srcMy = (Math.abs(myLast - myEst) <= 3) ? 'last' : 'est'; }
    else if (isFinite(myEst)) { selMy = myEst; srcMy = 'est'; }
    else if (isFinite(myLast)) { selMy = myLast; srcMy = 'last_only'; }
    else { selMy = ''; srcMy = 'none'; }
    if (myPregameSelectedVals && String(myPregameSelectedVals[i][0] || '').trim() === '' && selMy !== '') { myPregameSelectedVals[i][0] = selMy; wrote = true; }
    if (myPregameSourceVals && String(myPregameSourceVals[i][0] || '').trim() === '' && srcMy) { myPregameSourceVals[i][0] = srcMy; wrote = true; }

    var oppCb = oppPregameVals ? Number(oppPregameVals[i][0]) : NaN;
    var oppEst = oppPregameFormulaVals ? Number(oppPregameFormulaVals[i][0]) : NaN;
    var selOpp = '';
    var srcOpp = '';
    if (isFinite(oppCb)) { selOpp = oppCb; srcOpp = 'cb'; }
    else if (isFinite(oppPost) && isFinite(Number(oppDelta))) { selOpp = (oppPost - Number(oppDelta)); srcOpp = 'cb_via_delta'; }
    else if (isFinite(oppPost) && isFinite(Number(myDelta))) { selOpp = (oppPost + Number(myDelta)); srcOpp = 'derived_from_my_delta'; }
    else if (isFinite(oppEst)) { selOpp = oppEst; srcOpp = 'formula'; }
    else { selOpp = ''; srcOpp = 'none'; }
    if (oppPregameSelectedVals && String(oppPregameSelectedVals[i][0] || '').trim() === '' && selOpp !== '') { oppPregameSelectedVals[i][0] = selOpp; wrote = true; }
    if (oppPregameSourceVals && String(oppPregameSourceVals[i][0] || '').trim() === '' && srcOpp) { oppPregameSourceVals[i][0] = srcOpp; wrote = true; }

    if (deltaMismatchVals && isFinite(Number(myDelta)) && isFinite(Number(oppDelta))) {
      var sum = Number(myDelta) + Number(oppDelta);
      if (String(deltaMismatchVals[i][0] || '').trim() === '' && sum !== 0) { deltaMismatchVals[i][0] = sum; wrote = true; }
    }

    // Compute Result_Value from Result text
    if (resultValueVals && resultVals) {
      var r = String(resultVals[i][0] || '').toLowerCase();
      var rv = (r === 'won') ? 1 : (r === 'drew') ? 0.5 : (r === 'lost') ? 0 : '';
      if (String(resultValueVals[i][0] || '').trim() === '' && rv !== '') { resultValueVals[i][0] = rv; wrote = true; }
    }

    if (wrote) processed++;
  }

  // Batch write columns that exist
  function setCol(col, vals) {
    if (col && vals) sheet.getRange(2, col, numRows, 1).setValues(vals);
  }
  setCol(colWhiteDelta, whiteDeltaVals);
  setCol(colBlackDelta, blackDeltaVals);
  setCol(colMyDelta, myDeltaVals);
  setCol(colOppDelta, oppDeltaVals);
  setCol(colMyPregame, myPregameVals);
  setCol(colOppPregame, oppPregameVals);
  setCol(colMyPregameSelected, myPregameSelectedVals);
  setCol(colMyPregameSource, myPregameSourceVals);
  setCol(colOppPregameSelected, oppPregameSelectedVals);
  setCol(colOppPregameSource, oppPregameSourceVals);
  setCol(colDeltaMismatch, deltaMismatchVals);
  setCol(colResultValue, resultValueVals);
  setCol(colWinnerColor, winnerVals);
  setCol(colEndReason, endReasonVals);
  setCol(colResultMsg, resultMsgVals);
  setCol(colMoveTimestamps, tsVals);
  setCol(colMoveList, moveListVals);
  setCol(colLastMove, lastMoveVals);
  setCol(colBase, baseVals);
  setCol(colInc, incVals);
  setCol(colIsLive, isLiveVals);
  setCol(colIsAbortable, isAbortableVals);
  setCol(colIsAnalyzable, isAnalyzableVals);
  setCol(colIsResignable, isResignableVals);
  setCol(colIsCheckmate, isCheckmateVals);
  setCol(colIsStalemate, isStalemateVals);
  setCol(colIsFinished, isFinishedVals);
  setCol(colCanSendTrophy, canSendTrophyVals);
  setCol(colChangesPlayersRating, changesPlayersRatingVals);
  setCol(colAllowVacation, allowVacationVals);
  setCol(colUuid, uuidVals);
  setCol(colTurnColor, turnVals);
  setCol(colPlyCount, plyVals);
  setCol(colInitialSetup, setupVals);
  setCol(colTypeName, typeNameVals);
  setCol(colOppMembershipCode, oppMembershipCodeVals);
  setCol(colOppMembershipLevel, oppMembershipLevelVals);
  setCol(colOppCountry, oppCountryVals);
  setCol(colOppAvatar, oppAvatarVals);
  setCol(colMyMembershipCode, myMembershipCodeVals);
  setCol(colMyMembershipLevel, myMembershipLevelVals);
  return processed;
}
/**
 * Convenience: backfill both ECO/Opening and Moves/Clocks.
 * @param {number=} limit Optional max rows for each pass
 */
function backfillAllMetadata(limit) {
  backfillEcoAndOpening(limit);
  backfillMovesAndClocks(limit);
}

/**
 * Updates ONLY the dedicated callback header fields in batches:
 *  - White rating change
 *  - Black rating change
 *  - Opponent membership code/level/country/avatar URL
 *  - My membership code/level
 * Returns number of rows processed this run.
 * @param {number=} limit
 * @return {number}
 */
function backfillCallbackHeaderFields(limit) {
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerToIndex = {};
  for (var c = 0; c < headers.length; c++) headerToIndex[headers[c]] = c + 1;

  var colGameId = headerToIndex['Game ID'] || 3;
  var colUrl = headerToIndex['URL'] || 2;
  var colColor = headerToIndex['Color'];
  var colWhiteDelta = headerToIndex['White rating change'];
  var colBlackDelta = headerToIndex['Black rating change'];
  var colOppMembershipCode = headerToIndex['Opponent membership code'];
  var colOppMembershipLevel = headerToIndex['Opponent membership level'];
  var colOppCountry = headerToIndex['Opponent country'];
  var colOppAvatar = headerToIndex['Opponent avatar URL'];
  var colMyMembershipCode = headerToIndex['My membership code'];
  var colMyMembershipLevel = headerToIndex['My membership level'];

  var numRows = lastRow - 1;
  var gameIds = sheet.getRange(2, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(2, colUrl, numRows, 1).getValues();
  var colors = colColor ? sheet.getRange(2, colColor, numRows, 1).getValues() : null;

  function initColumn(col) { return col ? sheet.getRange(2, col, numRows, 1).getValues() : null; }
  var whiteDeltaVals = initColumn(colWhiteDelta);
  var blackDeltaVals = initColumn(colBlackDelta);
  var oppMembershipCodeVals = initColumn(colOppMembershipCode);
  var oppMembershipLevelVals = initColumn(colOppMembershipLevel);
  var oppCountryVals = initColumn(colOppCountry);
  var oppAvatarVals = initColumn(colOppAvatar);
  var myMembershipCodeVals = initColumn(colMyMembershipCode);
  var myMembershipLevelVals = initColumn(colMyMembershipLevel);

  // Build candidate list: rows missing any targeted callback header fields
  var candidates = [];
  for (var r = 0; r < numRows; r++) {
    var needs = false;
    if (whiteDeltaVals && String(whiteDeltaVals[r][0] || '').trim() === '') needs = true;
    if (!needs && blackDeltaVals && String(blackDeltaVals[r][0] || '').trim() === '') needs = true;
    if (!needs && oppMembershipCodeVals && String(oppMembershipCodeVals[r][0] || '').trim() === '') needs = true;
    if (!needs && oppMembershipLevelVals && String(oppMembershipLevelVals[r][0] || '').trim() === '') needs = true;
    if (!needs && oppCountryVals && String(oppCountryVals[r][0] || '').trim() === '') needs = true;
    if (!needs && oppAvatarVals && String(oppAvatarVals[r][0] || '').trim() === '') needs = true;
    if (!needs && myMembershipCodeVals && String(myMembershipCodeVals[r][0] || '').trim() === '') needs = true;
    if (!needs && myMembershipLevelVals && String(myMembershipLevelVals[r][0] || '').trim() === '') needs = true;
    if (needs) candidates.push(r);
  }

  if (typeof limit === 'number' && isFinite(limit) && limit >= 0) {
    candidates = candidates.slice(0, limit);
  } else if (typeof BATCH_SIZE === 'number' && BATCH_SIZE > 0) {
    candidates = candidates.slice(0, BATCH_SIZE);
  }

  var processed = 0;
  for (var k = 0; k < candidates.length; k++) {
    var i = candidates[k];
    var gid = String(gameIds[i][0] || '').trim();
    if (!gid) {
      var url = String(urls[i][0] || '');
      var idMatch = url.match(/\/(\d+)(?:\?.*)?$/);
      if (idMatch && idMatch[1]) gid = idMatch[1];
    }
    if (!gid) continue;

    var data = fetchCallbackGameData(gid);
    if (!data || !data.game) continue;
    var g = data.game;
    var players = data.players || {};
    var top = players.top || {};
    var bottom = players.bottom || {};

    if (whiteDeltaVals && String(whiteDeltaVals[i][0] || '').trim() === '' && (g.ratingChangeWhite != null)) whiteDeltaVals[i][0] = g.ratingChangeWhite;
    if (blackDeltaVals && String(blackDeltaVals[i][0] || '').trim() === '' && (g.ratingChangeBlack != null)) blackDeltaVals[i][0] = g.ratingChangeBlack;

    var ourColor = colors ? String(colors[i][0] || '').toLowerCase() : '';
    var opponentObj = (ourColor === 'white') ? top : bottom;
    var myObj = (ourColor === 'white') ? bottom : top;

    if (oppMembershipCodeVals && String(oppMembershipCodeVals[i][0] || '').trim() === '' && opponentObj.membershipCode) oppMembershipCodeVals[i][0] = opponentObj.membershipCode;
    if (oppMembershipLevelVals && String(oppMembershipLevelVals[i][0] || '').trim() === '' && (opponentObj.membershipLevel != null)) oppMembershipLevelVals[i][0] = opponentObj.membershipLevel;
    if (oppCountryVals && String(oppCountryVals[i][0] || '').trim() === '' && opponentObj.countryName) oppCountryVals[i][0] = opponentObj.countryName;
    if (oppAvatarVals && String(oppAvatarVals[i][0] || '').trim() === '' && opponentObj.avatarUrl) oppAvatarVals[i][0] = opponentObj.avatarUrl;
    if (myMembershipCodeVals && String(myMembershipCodeVals[i][0] || '').trim() === '' && myObj.membershipCode) myMembershipCodeVals[i][0] = myObj.membershipCode;
    if (myMembershipLevelVals && String(myMembershipLevelVals[i][0] || '').trim() === '' && (myObj.membershipLevel != null)) myMembershipLevelVals[i][0] = myObj.membershipLevel;

    processed++;
  }

  function setCol(col, vals) { if (col && vals) sheet.getRange(2, col, numRows, 1).setValues(vals); }
  setCol(colWhiteDelta, whiteDeltaVals);
  setCol(colBlackDelta, blackDeltaVals);
  setCol(colOppMembershipCode, oppMembershipCodeVals);
  setCol(colOppMembershipLevel, oppMembershipLevelVals);
  setCol(colOppCountry, oppCountryVals);
  setCol(colOppAvatar, oppAvatarVals);
  setCol(colMyMembershipCode, myMembershipCodeVals);
  setCol(colMyMembershipLevel, myMembershipLevelVals);

  return processed;
}

/**
 * Runs the dedicated callback header-field backfill repeatedly until completion.
 */
function runCallbackHeadersBatchToCompletion() {
  while (true) {
    var processed = backfillCallbackHeaderFields(BATCH_SIZE);
    if (processed < BATCH_SIZE) break;
    Utilities.sleep(THROTTLE_SLEEP_MS);
  }
}

/**
 * Runs remaining callback-derived metadata in repeated batches until complete.
 * This includes Moves/Clocks, ECO/Opening URL, and other callback fields
 * (excluding the dedicated header subset).
 */
function runCallbackOthersBatchToCompletion() {
  // Moves & Clocks
  while (true) {
    var a = backfillMovesAndClocks(BATCH_SIZE);
    if (a < BATCH_SIZE) break;
    Utilities.sleep(THROTTLE_SLEEP_MS);
  }
  // Skip ECO & Opening and other non-callback-only fields here by design
}

/**
 * Runs the general callback-derived rating/metadata backfill repeatedly until completion.
 */
function runCallbackFieldsBatchToCompletion() {
  while (true) {
    var n = backfillCallbackFields(BATCH_SIZE);
    if (n < BATCH_SIZE) break;
    Utilities.sleep(THROTTLE_SLEEP_MS);
  }
}

/**
 * Splits a movetext string that contains move numbers and clock annotations
 * like: "1. e4 {[%clk 0:02:59.9]} 1... e5 {[%clk 0:02:59.1]} 2. Nf3 {[%clk 0:02:59.1]}"
 * into a structured list of moves.
 *
 * Rules:
 * - "n." indicates White's move number n
 * - "n..." indicates Black's move number n
 * - Each move is followed by a clock annotation in the form {[%clk H:MM:SS(.t)?]}
 * - The dot in the time is fractional seconds and must not be confused with move numbers
 *
 * @param {string} movetext Raw movetext containing numbers, SAN and {[%clk ...]} annotations
 * @return {Array<{ moveNumber: number, color: 'white'|'black', san: string, clock: string, clockSeconds: number }>} Parsed moves
 */
function splitMovesWithClocks(movetext) {
  if (!movetext || typeof movetext !== 'string') return [];

  // Global regex to capture: moveNumber, (optional) "..." for black, SAN token, and clock time
  // Example match: 1. e4 {[%clk 0:02:59.9]}
  //                1... e5 {[%clk 0:02:59.1]}
  var moveRegex = /(\d+)\.(\.\.)?\s*([^\s\{]+)\s*\{\s*\[%clk\s+([^\]]+)\]\s*\}/g;

  var results = [];
  var match;
  while ((match = moveRegex.exec(movetext)) !== null) {
    var moveNumber = parseInt(match[1], 10);
    var isBlack = !!match[2];
    var san = match[3];
    var clockRaw = String(match[4]).trim();

    var clockSeconds = parseClockToSeconds(clockRaw);

    results.push({
      moveNumber: isFinite(moveNumber) ? moveNumber : null,
      color: isBlack ? 'black' : 'white',
      san: san,
      clock: clockRaw,
      clockSeconds: clockSeconds
    });
  }

  return results;
}

/**
 * Converts a clock string like "0:02:59.9" or "1:00:00" to seconds as a number.
 * Accepts H:MM:SS(.t)? or MM:SS(.t)? formats.
 * @param {string} clock
 * @return {number} seconds (floating)
 */
function parseClockToSeconds(clock) {
  if (!clock) return NaN;
  var parts = String(clock).split(':');
  if (parts.length < 2) return NaN;

  // Parse from right: seconds (may contain fraction), minutes, optional hours
  var secondsPart = parts.pop();
  var minutesPart = parts.pop();
  var hoursPart = parts.length > 0 ? parts.pop() : '0';

  var seconds = parseFloat(secondsPart.replace(/[^0-9\.]/g, ''));
  var minutes = parseInt(minutesPart, 10);
  var hours = parseInt(hoursPart, 10);

  if (!isFinite(seconds) || !isFinite(minutes) || !isFinite(hours)) return NaN;
  return (hours * 3600) + (minutes * 60) + seconds;
}

/**
 * Backfill "My rating pregame_derived" as the latest prior "My rating" for the same Format.
 * For each row, find the most recent earlier row (by Timestamp) whose Format equals this row's Format,
 * and set derived = that row's My rating. Only fills empty cells, never overwrites.
 * @param {number=} limit Optional max rows to process
 * @return {number} number of rows updated
 */
function backfillMyRatingPregameDerived(limit) {
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerToIndex = {};
  for (var c = 0; c < headers.length; c++) headerToIndex[headers[c]] = c + 1;

  var colTimestamp = headerToIndex['Timestamp'] || 1;
  var colFormat = headerToIndex['Format'];
  var colMyRating = headerToIndex['My rating'];
  var colDerived = headerToIndex['My rating pregame_derived'];
  if (!colFormat || !colMyRating || !colDerived) return 0;

  var numRows = lastRow - 1;
  var tsVals = sheet.getRange(2, colTimestamp, numRows, 1).getValues();
  var fmtVals = sheet.getRange(2, colFormat, numRows, 1).getValues();
  var myVals = sheet.getRange(2, colMyRating, numRows, 1).getValues();
  var drvVals = sheet.getRange(2, colDerived, numRows, 1).getValues();

  // Build indexes of rows by format with timestamps to allow binary search of prior max<current
  var formatToRows = {};
  for (var i = 0; i < numRows; i++) {
    var f = String(fmtVals[i][0] || '').trim();
    if (!formatToRows[f]) formatToRows[f] = [];
    var t = tsVals[i][0];
    var timeMs = (t instanceof Date) ? t.getTime() : (new Date(t)).getTime();
    formatToRows[f].push({ idx: i, time: isFinite(timeMs) ? timeMs : 0 });
  }
  for (var key in formatToRows) {
    formatToRows[key].sort(function(a, b) { return a.time - b.time; });
  }

  // Create mapping from row index to its position within sorted list for its format
  var formatToSortedIndexes = {};
  for (var fKey in formatToRows) {
    var arr = formatToRows[fKey];
    var posByIdx = {};
    for (var p = 0; p < arr.length; p++) posByIdx[arr[p].idx] = p;
    formatToSortedIndexes[fKey] = posByIdx;
  }

  // Determine candidate rows: derived empty
  var candidates = [];
  for (var r = 0; r < numRows; r++) {
    if (String(drvVals[r][0] || '').trim() === '') candidates.push(r);
  }
  if (typeof limit === 'number' && isFinite(limit) && limit >= 0) {
    candidates = candidates.slice(0, limit);
  } else if (typeof BATCH_SIZE === 'number' && BATCH_SIZE > 0) {
    candidates = candidates.slice(0, BATCH_SIZE);
  }

  var updated = 0;
  for (var k = 0; k < candidates.length; k++) {
    var rowIdx = candidates[k];
    var f = String(fmtVals[rowIdx][0] || '').trim();
    var arr = formatToRows[f] || [];
    var pos = (formatToSortedIndexes[f] && formatToSortedIndexes[f][rowIdx] != null) ? formatToSortedIndexes[f][rowIdx] : -1;
    if (pos <= 0) continue; // no earlier row
    // find prior position with valid My rating
    for (var q = pos - 1; q >= 0; q--) {
      var priorIdx = arr[q].idx;
      var priorRating = Number(myVals[priorIdx][0]);
      if (isFinite(priorRating)) {
        drvVals[rowIdx][0] = priorRating;
        updated++;
        break;
      }
    }
  }

  if (updated > 0) sheet.getRange(2, colDerived, numRows, 1).setValues(drvVals);
  return updated;
}

/**
 * Runs backfillMyRatingPregameDerived repeatedly until completion.
 */
function runMyRatingPregameDerivedToCompletion() {
  while (true) {
    var n = backfillMyRatingPregameDerived(BATCH_SIZE);
    if (n < BATCH_SIZE) break;
    Utilities.sleep(THROTTLE_SLEEP_MS);
  }
}

// =====================
// Grouped, range-based processors
// =====================

/**
 * Run grouped processors for a contiguous range of rows.
 * startRow is 1-based, and should point to the first data row (>=2).
 * @param {number} startRow First row to process (1-based)
 * @param {number} numRows Number of rows to process
 * @param {{ includeCallback?: boolean, includeMovesClocks?: boolean, includeEcoOpening?: boolean, includeDerived?: boolean }} options
 */
function processNewRowsGrouped(startRow, numRows, options) {
  var sheet = ensureSheet();
  if (!isFinite(startRow) || !isFinite(numRows) || numRows <= 0) return;
  var opts = options || {};
  var doCallback = !!opts.includeCallback; // include callback-derived fields
  var doMovesClocks = opts.includeMovesClocks != null ? !!opts.includeMovesClocks : true;
  var doEcoOpening = opts.includeEcoOpening != null ? !!opts.includeEcoOpening : true;
  var doDerived = opts.includeDerived != null ? !!opts.includeDerived : true;

  // Moves & Clocks (from callback PGN) are safe and useful for recent rows
  if (doMovesClocks) backfillMovesAndClocksInRange(sheet, startRow, numRows);

  // ECO & Opening URL
  if (doEcoOpening) backfillEcoAndOpeningInRange(sheet, startRow, numRows);

  // General callback-derived fields (rating pregame/deltas, result value, etc.)
  if (doCallback) backfillCallbackFieldsInRange(sheet, startRow, numRows);

  // Derived rating pregame_derived depends on prior rows, compute for target rows
  if (doDerived) backfillMyRatingPregameDerivedForRows(sheet, startRow, numRows);

  // Formula-based pregame estimates (Elo) for established players
  backfillPregameFormulaInRange(sheet, startRow, numRows, ELO_K_ESTABLISHED);
}

/**
 * Run grouped processors for the entire sheet. Useful after a full backfill ingest.
 * @param {{ includeCallback?: boolean }} options
 */
function processAllRowsGrouped(options) {
  var sheet = ensureSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var count = lastRow - 1;
  var opts = options || {};
  var doCallback = !!opts.includeCallback;

  backfillMovesAndClocksInRange(sheet, 2, count);
  backfillEcoAndOpeningInRange(sheet, 2, count);
  if (doCallback) backfillCallbackFieldsInRange(sheet, 2, count);
  backfillMyRatingPregameDerivedForRows(sheet, 2, count);
  backfillPregameFormulaInRange(sheet, 2, count, ELO_K_ESTABLISHED);
}

// ---- Helpers (range-limited variants) ----

function getHeaderToIndexMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var c = 0; c < headers.length; c++) map[headers[c]] = c + 1;
  return map;
}

/**
 * Range-limited variant of backfillMovesAndClocks.
 */
function backfillMovesAndClocksInRange(sheet, startRow, numRows) {
  var headerToIndex = getHeaderToIndexMap_(sheet);
  var colGameId = headerToIndex['Game ID'] || 3;
  var colUrl = headerToIndex['URL'] || 2;
  var colMovesSan = headerToIndex['Moves (SAN)'];
  var colClocks = headerToIndex['Clocks'];
  var colClockSeconds = headerToIndex['Clock Seconds'];
  var colMoveTimes = headerToIndex['Move Times'];
  var colTimeClass = headerToIndex['Time class'];
  var colTimeControl = headerToIndex['Time control'];
  var colBase = headerToIndex['Base time (s)'];
  var colInc = headerToIndex['Increment (s)'];
  if (!colMovesSan || !colClocks) return 0;

  var gameIds = sheet.getRange(startRow, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(startRow, colUrl, numRows, 1).getValues();
  var movesSanVals = sheet.getRange(startRow, colMovesSan, numRows, 1).getValues();
  var clocksVals = sheet.getRange(startRow, colClocks, numRows, 1).getValues();
  var clockSecondsVals = colClockSeconds ? sheet.getRange(startRow, colClockSeconds, numRows, 1).getValues() : null;
  var moveTimesVals = colMoveTimes ? sheet.getRange(startRow, colMoveTimes, numRows, 1).getValues() : null;
  var timeClassVals = colTimeClass ? sheet.getRange(startRow, colTimeClass, numRows, 1).getValues() : null;
  var timeControlVals = colTimeControl ? sheet.getRange(startRow, colTimeControl, numRows, 1).getValues() : null;
  var baseVals = colBase ? sheet.getRange(startRow, colBase, numRows, 1).getValues() : null;
  var incVals = colInc ? sheet.getRange(startRow, colInc, numRows, 1).getValues() : null;

  var processed = 0;
  for (var i = 0; i < numRows; i++) {
    var currentMoves = String(movesSanVals[i][0] || '').trim();
    var currentClocks = String(clocksVals[i][0] || '').trim();
    if (currentMoves && currentClocks) continue;

    var gid = String(gameIds[i][0] || '').trim();
    if (!gid) {
      var url = String(urls[i][0] || '');
      var idMatch = url.match(/\/(\d+)(?:\?.*)?$/);
      if (idMatch && idMatch[1]) gid = idMatch[1];
    }
    if (!gid) continue;

    var pgn = fetchCallbackGamePgn(gid);
    if (!pgn) continue;
    var parsed = parsePgnToSanAndClocks(pgn);
    var newMovesSan = parsed.movesSan || currentMoves;
    var newClocks = parsed.clocks || currentClocks;
    movesSanVals[i][0] = newMovesSan;
    clocksVals[i][0] = newClocks;

    if (clockSecondsVals || moveTimesVals || baseVals || incVals) {
      var timeClass = timeClassVals ? String(timeClassVals[i][0] || '') : '';
      var timeControl = timeControlVals ? String(timeControlVals[i][0] || '') : '';
      var tcParsed = parseTimeControlToBaseInc(timeControl, timeClass);
      var baseSeconds = tcParsed.baseSeconds;
      var incSeconds = tcParsed.incrementSeconds;
      if (baseVals) baseVals[i][0] = (baseSeconds != null ? baseSeconds : baseVals[i][0]);
      if (incVals) incVals[i][0] = (incSeconds != null ? incSeconds : incVals[i][0]);

      var movetext = extractMovetextFromPgn(pgn);
      var parsedMoves = movetext ? splitMovesWithClocks(movetext) : [];
      var derived = buildClockSecondsAndMoveTimes(parsedMoves, baseSeconds, incSeconds);
      if (clockSecondsVals) clockSecondsVals[i][0] = derived.clockSecondsStr || clockSecondsVals[i][0];
      if (moveTimesVals) moveTimesVals[i][0] = derived.moveTimesStr || moveTimesVals[i][0];
    }
    processed++;
  }

  sheet.getRange(startRow, colMovesSan, numRows, 1).setValues(movesSanVals);
  sheet.getRange(startRow, colClocks, numRows, 1).setValues(clocksVals);
  if (colClockSeconds && clockSecondsVals) sheet.getRange(startRow, colClockSeconds, numRows, 1).setValues(clockSecondsVals);
  if (colMoveTimes && moveTimesVals) sheet.getRange(startRow, colMoveTimes, numRows, 1).setValues(moveTimesVals);
  if (colBase && baseVals) sheet.getRange(startRow, colBase, numRows, 1).setValues(baseVals);
  if (colInc && incVals) sheet.getRange(startRow, colInc, numRows, 1).setValues(incVals);
  return processed;
}

/**
 * Range-limited variant of backfillEcoAndOpening.
 */
function backfillEcoAndOpeningInRange(sheet, startRow, numRows) {
  var headerToIndex = getHeaderToIndexMap_(sheet);
  var colGameId = headerToIndex['Game ID'] || 3;
  var colUrl = headerToIndex['URL'] || 2;
  var colEco = headerToIndex['ECO'];
  var colOpeningUrl = headerToIndex['Opening URL'];
  if (!colEco || !colOpeningUrl) return 0;

  var gameIds = sheet.getRange(startRow, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(startRow, colUrl, numRows, 1).getValues();
  var ecos = sheet.getRange(startRow, colEco, numRows, 1).getValues();
  var openingUrls = sheet.getRange(startRow, colOpeningUrl, numRows, 1).getValues();

  var processed = 0;
  for (var i = 0; i < numRows; i++) {
    var ecoVal = String(ecos[i][0] || '').trim();
    var openVal = String(openingUrls[i][0] || '').trim();
    var ecoLooksWrong = ecoVal && /^https?:\/\//i.test(ecoVal);
    if (!(ecoLooksWrong || !ecoVal || !openVal)) continue;

    var gid = String(gameIds[i][0] || '').trim();
    if (!gid) {
      var url = String(urls[i][0] || '');
      var idMatch = url.match(/\/(\d+)(?:\?.*)?$/);
      if (idMatch && idMatch[1]) gid = idMatch[1];
    }
    if (!gid) continue;

    var pgn = fetchCallbackGamePgn(gid);
    if (!pgn) continue;
    var parsed = parseEcoAndOpeningFromPgn(pgn);
    var wrote = false;
    if (!ecoVal || /^https?:\/\//i.test(ecoVal)) {
      var newEco = parsed.eco || ecos[i][0];
      if (newEco !== ecos[i][0]) { ecos[i][0] = newEco; wrote = true; }
    }
    if (!openVal) {
      var newUrl = parsed.openingUrl || openingUrls[i][0];
      if (newUrl !== openingUrls[i][0]) { openingUrls[i][0] = newUrl; wrote = true; }
    }
    if (wrote) processed++;
  }

  sheet.getRange(startRow, colEco, numRows, 1).setValues(ecos);
  sheet.getRange(startRow, colOpeningUrl, numRows, 1).setValues(openingUrls);
  return processed;
}

/**
 * Range-limited variant of backfillCallbackFields.
 */
function backfillCallbackFieldsInRange(sheet, startRow, numRows) {
  var headerToIndex = getHeaderToIndexMap_(sheet);
  var colGameId = headerToIndex['Game ID'] || 3;
  var colUrl = headerToIndex['URL'] || 2;
  var colColor = headerToIndex['Color'];
  var colMyRating = headerToIndex['My rating'];
  var colOppRating = headerToIndex['Opponent rating'];
  var colResult = headerToIndex['Result'];
  var colResultValue = headerToIndex['Result_Value'];
  var colMyDelta = headerToIndex['My rating change'];
  var colOppDelta = headerToIndex['Opponent rating change'];
  var colMyPregame = headerToIndex['My rating pregame'];
  var colOppPregame = headerToIndex['Opponent rating pregame'];
  var colMyPregameDerived = headerToIndex['My rating pregame_derived'];
  var colMyPregameFormula = headerToIndex['My rating pregame_formula'];
  var colOppPregameFormula = headerToIndex['Opponent rating pregame_formula'];
  var colMyPregameSelected = headerToIndex['My rating pregame_selected'];
  var colMyPregameSource = headerToIndex['My rating pregame_source'];
  var colOppPregameSelected = headerToIndex['Opponent rating pregame_selected'];
  var colOppPregameSource = headerToIndex['Opponent rating pregame_source'];
  var colDeltaMismatch = headerToIndex['Rating delta mismatch'];

  if (!colMyDelta && !colOppDelta && !colMyPregame && !colOppPregame && !colResultValue && !colMyPregameSelected && !colOppPregameSelected) return 0;

  function init(col) { return col ? sheet.getRange(startRow, col, numRows, 1).getValues() : null; }
  var gameIds = sheet.getRange(startRow, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(startRow, colUrl, numRows, 1).getValues();
  var colors = init(colColor);
  var myRatings = init(colMyRating);
  var oppRatings = init(colOppRating);
  var resultVals = init(colResult);
  var myDeltaVals = init(colMyDelta);
  var oppDeltaVals = init(colOppDelta);
  var myPregameVals = init(colMyPregame);
  var oppPregameVals = init(colOppPregame);
  var resultValueVals = init(colResultValue);
  var myPregameDerivedVals = init(colMyPregameDerived);
  var myPregameFormulaVals = init(colMyPregameFormula);
  var oppPregameFormulaVals = init(colOppPregameFormula);
  var myPregameSelectedVals = init(colMyPregameSelected);
  var myPregameSourceVals = init(colMyPregameSource);
  var oppPregameSelectedVals = init(colOppPregameSelected);
  var oppPregameSourceVals = init(colOppPregameSource);
  var deltaMismatchVals = init(colDeltaMismatch);

  var processed = 0;
  for (var i = 0; i < numRows; i++) {
    var needs = false;
    if (myDeltaVals && String(myDeltaVals[i][0] || '').trim() === '') needs = true;
    if (!needs && oppDeltaVals && String(oppDeltaVals[i][0] || '').trim() === '') needs = true;
    if (!needs && myPregameVals && String(myPregameVals[i][0] || '').trim() === '') needs = true;
    if (!needs && oppPregameVals && String(oppPregameVals[i][0] || '').trim() === '') needs = true;
    if (!needs && resultValueVals && String(resultValueVals[i][0] || '').trim() === '') needs = true;
    if (!needs) continue;

    var gid = String(gameIds[i][0] || '').trim();
    if (!gid) {
      var url = String(urls[i][0] || '');
      var idMatch = url.match(/\/(\d+)(?:\?.*)?$/);
      if (idMatch && idMatch[1]) gid = idMatch[1];
    }
    if (!gid) continue;

    var data = fetchCallbackGameData(gid);
    if (!data || !data.game) continue;
    var g = data.game;

    var ourColor = colors ? String(colors[i][0] || '').toLowerCase() : '';
    var myDelta = (ourColor === 'white') ? g.ratingChangeWhite : g.ratingChangeBlack;
    var oppDelta = (ourColor === 'white') ? g.ratingChangeBlack : g.ratingChangeWhite;
    if (myDeltaVals && String(myDeltaVals[i][0] || '').trim() === '' && (myDelta != null)) myDeltaVals[i][0] = myDelta;
    if (oppDeltaVals && String(oppDeltaVals[i][0] || '').trim() === '' && (oppDelta != null)) oppDeltaVals[i][0] = oppDelta;

    var myPost = myRatings ? Number(myRatings[i][0]) : NaN;
    var oppPost = oppRatings ? Number(oppRatings[i][0]) : NaN;
    var myPre = (isFinite(myPost) && isFinite(Number(myDelta))) ? (myPost - Number(myDelta)) : '';
    var oppPre = (isFinite(oppPost) && isFinite(Number(oppDelta))) ? (oppPost - Number(oppDelta)) : '';
    if (myPregameVals && String(myPregameVals[i][0] || '').trim() === '' && myPre !== '') myPregameVals[i][0] = myPre;
    if (oppPregameVals && String(oppPregameVals[i][0] || '').trim() === '' && oppPre !== '') oppPregameVals[i][0] = oppPre;

    // Compute selected values
    var myCb = myPregameVals ? Number(myPregameVals[i][0]) : NaN;
    var myLast = myPregameDerivedVals ? Number(myPregameDerivedVals[i][0]) : NaN;
    var myEst = myPregameFormulaVals ? Number(myPregameFormulaVals[i][0]) : NaN;
    var selMy = '';
    var srcMy = '';
    if (isFinite(myCb)) { selMy = myCb; srcMy = 'cb'; }
    else if (isFinite(myLast) && isFinite(myEst)) { selMy = (Math.abs(myLast - myEst) <= 3) ? myLast : myEst; srcMy = (Math.abs(myLast - myEst) <= 3) ? 'last' : 'est'; }
    else if (isFinite(myEst)) { selMy = myEst; srcMy = 'est'; }
    else if (isFinite(myLast)) { selMy = myLast; srcMy = 'last_only'; }
    else { selMy = ''; srcMy = 'none'; }
    if (myPregameSelectedVals && String(myPregameSelectedVals[i][0] || '').trim() === '' && selMy !== '') myPregameSelectedVals[i][0] = selMy;
    if (myPregameSourceVals && String(myPregameSourceVals[i][0] || '').trim() === '' && srcMy) myPregameSourceVals[i][0] = srcMy;

    var oppCb = oppPregameVals ? Number(oppPregameVals[i][0]) : NaN;
    var oppEst = oppPregameFormulaVals ? Number(oppPregameFormulaVals[i][0]) : NaN;
    var selOpp = '';
    var srcOpp = '';
    if (isFinite(oppCb)) { selOpp = oppCb; srcOpp = 'cb'; }
    else if (isFinite(oppPost) && isFinite(Number(oppDelta))) { selOpp = (oppPost - Number(oppDelta)); srcOpp = 'cb_via_delta'; }
    else if (isFinite(oppPost) && isFinite(Number(myDelta))) { selOpp = (oppPost + Number(myDelta)); srcOpp = 'derived_from_my_delta'; }
    else if (isFinite(oppEst)) { selOpp = oppEst; srcOpp = 'formula'; }
    else { selOpp = ''; srcOpp = 'none'; }
    if (oppPregameSelectedVals && String(oppPregameSelectedVals[i][0] || '').trim() === '' && selOpp !== '') oppPregameSelectedVals[i][0] = selOpp;
    if (oppPregameSourceVals && String(oppPregameSourceVals[i][0] || '').trim() === '' && srcOpp) oppPregameSourceVals[i][0] = srcOpp;

    if (deltaMismatchVals && isFinite(Number(myDelta)) && isFinite(Number(oppDelta))) {
      var sum = Number(myDelta) + Number(oppDelta);
      if (String(deltaMismatchVals[i][0] || '').trim() === '' && sum !== 0) deltaMismatchVals[i][0] = sum;
    }

    if (resultValueVals && resultVals) {
      var r = String(resultVals[i][0] || '').toLowerCase();
      var rv = (r === 'won') ? 1 : (r === 'drew') ? 0.5 : (r === 'lost') ? 0 : '';
      if (String(resultValueVals[i][0] || '').trim() === '' && rv !== '') resultValueVals[i][0] = rv;
    }
    processed++;
  }

  function setCol(col, vals) { if (col && vals) sheet.getRange(startRow, col, numRows, 1).setValues(vals); }
  setCol(colMyDelta, myDeltaVals);
  setCol(colOppDelta, oppDeltaVals);
  setCol(colMyPregame, myPregameVals);
  setCol(colOppPregame, oppPregameVals);
  setCol(colMyPregameSelected, myPregameSelectedVals);
  setCol(colMyPregameSource, myPregameSourceVals);
  setCol(colOppPregameSelected, oppPregameSelectedVals);
  setCol(colOppPregameSource, oppPregameSourceVals);
  setCol(colDeltaMismatch, deltaMismatchVals);
  setCol(colResultValue, resultValueVals);
  return processed;
}

/**
 * Fill My rating pregame_derived for only the specified rows, using global history.
 */
function backfillMyRatingPregameDerivedForRows(sheet, startRow, numRows) {
  var headerToIndex = getHeaderToIndexMap_(sheet);
  var colTimestamp = headerToIndex['Timestamp'] || 1;
  var colFormat = headerToIndex['Format'];
  var colMyRating = headerToIndex['My rating'];
  var colDerived = headerToIndex['My rating pregame_derived'];
  if (!colFormat || !colMyRating || !colDerived) return 0;

  var lastRow = sheet.getLastRow();
  var total = lastRow - 1;
  if (total <= 0) return 0;

  var tsVals = sheet.getRange(2, colTimestamp, total, 1).getValues();
  var fmtVals = sheet.getRange(2, colFormat, total, 1).getValues();
  var myVals = sheet.getRange(2, colMyRating, total, 1).getValues();
  var drvVals = sheet.getRange(2, colDerived, total, 1).getValues();

  var formatToRows = {};
  for (var i = 0; i < total; i++) {
    var f = String(fmtVals[i][0] || '').trim();
    if (!formatToRows[f]) formatToRows[f] = [];
    var t = tsVals[i][0];
    var timeMs = (t instanceof Date) ? t.getTime() : (new Date(t)).getTime();
    formatToRows[f].push({ idx: i, time: isFinite(timeMs) ? timeMs : 0 });
  }
  for (var key in formatToRows) {
    formatToRows[key].sort(function(a, b) { return a.time - b.time; });
  }
  var formatToSortedIndexes = {};
  for (var fKey in formatToRows) {
    var arr = formatToRows[fKey];
    var posByIdx = {};
    for (var p = 0; p < arr.length; p++) posByIdx[arr[p].idx] = p;
    formatToSortedIndexes[fKey] = posByIdx;
  }

  var updated = 0;
  var startIdx0 = startRow - 2; // convert to 0-based index within data rows
  for (var k = 0; k < numRows; k++) {
    var rowIdx = startIdx0 + k;
    if (rowIdx < 0 || rowIdx >= total) continue;
    if (String(drvVals[rowIdx][0] || '').trim() !== '') continue;
    var f = String(fmtVals[rowIdx][0] || '').trim();
    var arr = formatToRows[f] || [];
    var pos = (formatToSortedIndexes[f] && formatToSortedIndexes[f][rowIdx] != null) ? formatToSortedIndexes[f][rowIdx] : -1;
    if (pos <= 0) continue;
    for (var q = pos - 1; q >= 0; q--) {
      var priorIdx = arr[q].idx;
      var priorRating = Number(myVals[priorIdx][0]);
      if (isFinite(priorRating)) {
        drvVals[rowIdx][0] = priorRating;
        updated++;
        break;
      }
    }
  }
  if (updated > 0) sheet.getRange(2, colDerived, total, 1).setValues(drvVals);
  return updated;
}

// =====================
// Lichess API helpers and orchestrator
// =====================

/**
 * Import a PGN into Lichess. Returns { id, url }.
 * @param {string} pgn
 * @param {{ source?: string }=} options
 * @return {{ id: string, url: string }}
 */
function lichessImportGameFromPgn(pgn, options) {
  if (!pgn || typeof pgn !== 'string') throw new Error('lichessImportGameFromPgn: PGN is required');
  var url = LICHESS_API_BASE + '/api/import';
  var params = options || {};
  var body = 'pgn=' + encodeURIComponent(pgn);
  if (params.source) body += '&source=' + encodeURIComponent(String(params.source));
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/x-www-form-urlencoded',
    payload: body,
    headers: getLichessAuthHeaders_('application/json')
  });
  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Lichess import failed: ' + code + ' ' + resp.getContentText());
  }
  var json = {};
  try { json = JSON.parse(resp.getContentText()); } catch (e) {}
  var gameUrl = String(json.url || '');
  var id = String(json.id || (gameUrl ? gameUrl.split('/').pop() : ''));
  if (!id) throw new Error('Lichess import succeeded but no game id in response');
  return { id: id, url: (gameUrl || (LICHESS_API_BASE + '/' + id)) };
}

/**
 * Export a single game PGN by Lichess game ID.
 * @param {string} gameId
 * @param {{ moves?: boolean, clocks?: boolean, evals?: boolean }=} options
 * @return {string} PGN
 */
function lichessExportGamePgn(gameId, options) {
  if (!gameId) throw new Error('lichessExportGamePgn: gameId required');
  var opts = options || {};
  // The .pgn endpoint returns text/plain PGN. Optional query params include moves, evals, clocks
  var qs = [];
  if (opts.moves != null) qs.push('moves=' + (opts.moves ? 'true' : 'false'));
  if (opts.clocks != null) qs.push('clocks=' + (opts.clocks ? 'true' : 'false'));
  if (opts.evals != null) qs.push('evals=' + (opts.evals ? 'true' : 'false'));
  var url = LICHESS_API_BASE + '/game/export/' + encodeURIComponent(String(gameId)) + '.pgn' + (qs.length ? ('?' + qs.join('&')) : '');
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: getLichessAuthHeaders_('text/plain')
  });
  var code = resp.getResponseCode();
  if (code !== 200) throw new Error('Lichess export failed: ' + code + ' ' + resp.getContentText());
  return String(resp.getContentText() || '');
}

/**
 * Cloud-evaluate a single FEN using Lichess Cloud Eval.
 * @param {string} fen
 * @param {{ multiPv?: number, syzygy?: number, variant?: string }=} options
 * @return {Object} JSON response
 */
function lichessCloudEvalFen(fen, options) {
  if (!fen) throw new Error('lichessCloudEvalFen: FEN required');
  var url = LICHESS_API_BASE + '/api/cloud-eval';
  var params = options || {};
  var qs = ['fen=' + encodeURIComponent(String(fen))];
  if (isFinite(params.multiPv)) qs.push('multiPv=' + Number(params.multiPv));
  if (isFinite(params.syzygy)) qs.push('syzygy=' + Number(params.syzygy));
  if (params.variant) qs.push('variant=' + encodeURIComponent(String(params.variant)));
  var resp = UrlFetchApp.fetch(url + '?' + qs.join('&'), {
    method: 'get',
    muteHttpExceptions: true,
    headers: getLichessAuthHeaders_('application/json')
  });
  var code = resp.getResponseCode();
  if (code !== 200) throw new Error('Cloud eval failed: ' + code + ' ' + resp.getContentText());
  try { return JSON.parse(resp.getContentText()); } catch (e) { return {}; }
}

/**
 * Cloud-evaluate multiple FENs. Returns array of { fen, eval } objects.
 * @param {Array<string>} fens
 * @param {{ multiPv?: number, syzygy?: number, variant?: string, throttleMs?: number }=} options
 * @return {Array<{ fen: string, eval: Object }>} results
 */
function lichessCloudEvalFens(fens, options) {
  var list = Array.isArray(fens) ? fens : [];
  var opts = options || {};
  var throttle = isFinite(opts.throttleMs) ? Number(opts.throttleMs) : 150;
  var results = [];
  for (var i = 0; i < list.length; i++) {
    var fen = String(list[i] || '').trim();
    if (!fen) continue;
    var ev = lichessCloudEvalFen(fen, opts);
    results.push({ fen: fen, eval: ev });
    if (throttle > 0) Utilities.sleep(throttle);
  }
  return results;
}

/**
 * OPTIONAL: Request external engine analysis. Note: This endpoint is intended for
 * Lichess external engine operators and is not available to regular users.
 * This function is provided for completeness and will throw unless explicitly allowed.
 * @param {string} gameId
 * @param {{ allow?: boolean }=} options
 * @return {Object}
 */
function lichessRequestExternalEngineAnalysis(gameId, options) {
  var opts = options || {};
  if (!opts.allow) throw new Error('External engine analysis requires operator privileges and is disabled. Set options.allow=true if you know what you are doing.');
  if (!gameId) throw new Error('lichessRequestExternalEngineAnalysis: gameId required');
  // Best-effort call; path subject to Lichess operator program requirements.
  var url = LICHESS_API_BASE + '/api/external-engine/analyse';
  var payload = 'gameId=' + encodeURIComponent(String(gameId));
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    headers: getLichessAuthHeaders_('application/json')
  });
  var code = resp.getResponseCode();
  if (code !== 200 && code !== 202) {
    throw new Error('External engine analysis request failed: ' + code + ' ' + resp.getContentText());
  }
  try { return JSON.parse(resp.getContentText()); } catch (e) { return { status: code }; }
}

/**
 * Orchestrator: Import PGN into Lichess, export it back, and analyze via Cloud Eval.
 * Optionally attempts external engine analysis if enabled.
 * @param {string} pgn
 * @param {{ fens?: Array<string>, fenExtractorName?: string, cloudEval?: { multiPv?: number, syzygy?: number, variant?: string, throttleMs?: number }, external?: { enabled?: boolean, allow?: boolean }, exportOptions?: { moves?: boolean, clocks?: boolean, evals?: boolean } }=} options
 * @return {{ gameId: string, url: string, exportedPgn: string, cloudEvaluations: Array<{ fen: string, eval: Object }>, externalAnalysis?: Object }}
 */
function lichessImportExportAndAnalyze(pgn, options) {
  var opts = options || {};
  var imported = lichessImportGameFromPgn(pgn, { source: 'Apps Script' });
  var exportedPgn = lichessExportGamePgn(imported.id, opts.exportOptions);

  // Determine FENs to evaluate
  var fens = [];
  if (Array.isArray(opts.fens) && opts.fens.length) {
    fens = opts.fens.slice();
  } else if (opts.fenExtractorName && typeof this[opts.fenExtractorName] === 'function') {
    try {
      var extracted = this[opts.fenExtractorName](exportedPgn);
      if (Array.isArray(extracted)) fens = extracted;
    } catch (e) {}
  }
  // Cloud eval the positions if provided
  var cloudEvaluations = fens.length ? lichessCloudEvalFens(fens, (opts.cloudEval || {})) : [];

  // Optional external engine analysis (requires operator access)
  var externalAnalysis = null;
  if (opts.external && (opts.external.enabled || opts.external.allow)) {
    externalAnalysis = lichessRequestExternalEngineAnalysis(imported.id, { allow: !!opts.external.allow });
  }

  return {
    gameId: imported.id,
    url: imported.url,
    exportedPgn: exportedPgn,
    cloudEvaluations: cloudEvaluations,
    externalAnalysis: externalAnalysis
  };
}

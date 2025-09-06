const PLAYER_USERNAME = 'frankscobey';
const SHEET_ID = '12zoMMkrkZz9WmhL4ds9lo8-91o-so337dOKwU3_yK3k';
const SHEET_NAME = 'Games';

// Optional: how many recent monthly archives to scan on each sync
const RECENT_ARCHIVES_TO_SCAN = 2;

function setupSheet() {
  const sheet = getOrCreateSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'URL',
      'Game ID',
      'Rated',
      'Time class',
      'Time control',
      'Rules',
      'Format',
      'Color',
      'Opponent',
      'Opponent rating',
      'My rating',
      'Result',
      'Reason',
      'FEN',
      'Endboard URL',
      'ECO',
      'Opening URL',
      'Moves (SAN)',
      'Clocks'
    ]);
  }
}

// Run this periodically via time-driven trigger (e.g., every 5–15 minutes)
function syncRecentGames() {
  const sheet = getOrCreateSheet();
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
}

// One-time (or repeated) backfill of your entire archive history
// If you have many games, this might need multiple runs due to execution time limits.
function backfillAllGamesOnce() {
  const sheet = getOrCreateSheet();
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
      if ((rowsToAppend.length % 100) === 0) Utilities.sleep(200);
    }

    // Optional: write in batches per month to avoid large in-memory arrays
    appendRows(sheet, rowsToAppend);
    rowsToAppend.length = 0;
  }
}

function getArchives() {
  const profileResp = UrlFetchApp.fetch(
    'https://api.chess.com/pub/player/' + encodeURIComponent(PLAYER_USERNAME),
    { muteHttpExceptions: true }
  );
  if (profileResp.getResponseCode() !== 200) {
    throw new Error('Player not found or API error: ' + profileResp.getResponseCode());
  }

  const archivesResp = UrlFetchApp.fetch(
    'https://api.chess.com/pub/player/' + encodeURIComponent(PLAYER_USERNAME) + '/games/archives',
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

function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
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
    details.rules,                    // Rules
    details.format || '',             // Format
    details.color,                    // Color
    details.opponentUsername,         // Opponent
    details.opponentRating || '',     // Opponent rating
    details.myRating || '',           // My rating
    details.result,                   // Result
    details.reason,                   // Reason
    details.fen,                      // FEN
    details.imageUrl,                 // Endboard URL
    details.eco || '',                // ECO
    details.openingUrl || '',         // Opening URL
    details.movesSan || '',           // Moves (SAN list)
    details.clocks || ''              // Clocks list
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

  var fen = game.fen || '';
  var fenEncoded = encodeURIComponent(fen);
  var imageUrl = 'https://www.chess.com/dynboard?fen=' + fenEncoded + '&board=brown&piece=neo&size=3';

  var endTime = game.end_time ? new Date(game.end_time * 1000) : new Date();

  // Parse PGN tags for ECO and Opening URL only (skip heavy movetext parsing here)
  var eco = game.eco || '';
  var openingUrl = '';
  var movesSan = '';
  var clocksStr = '';
  if (game.pgn) {
    try {
      var pgn = String(game.pgn).replace(/\r\n/g, '\n');
      // Extract tags like [ECO "B01"] and [ECOUrl "..."] anchored per line
      var ecoMatch = pgn.match(/^\[ECO\s+"([^"]+)"\]/m);
      if (ecoMatch && ecoMatch[1] && !eco) eco = ecoMatch[1];
      var ecoUrlMatch = pgn.match(/^\[(?:ECOUrl|OpeningUrl)\s+"([^"]+)"\]/m);
      if (ecoUrlMatch && ecoUrlMatch[1]) openingUrl = ecoUrlMatch[1];
    } catch (e) {
      // ignore PGN parse issues
    }
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
    reason: reason,
    fen: fen,
    imageUrl: imageUrl,
    eco: eco,
    openingUrl: openingUrl,
    movesSan: movesSan,
    clocks: clocksStr
  };
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
    .everyMinutes(15)
    .create();
}

/**
 * Ensure the header row contains "Moves (SAN)" and "Clocks" columns.
 * If an older header had a single "Moves" column, convert it to "Moves (SAN)"
 * and insert a new "Clocks" column after it.
 */
function ensureMovesClocksHeaders() {
  const sheet = getOrCreateSheet();
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var movesSanIdx = headers.indexOf('Moves (SAN)');
  var clocksIdx = headers.indexOf('Clocks');
  var legacyMovesIdx = headers.indexOf('Moves');

  if (movesSanIdx !== -1 && clocksIdx !== -1) {
    return; // already up to date
  }

  if (legacyMovesIdx !== -1) {
    // Rename legacy "Moves" to "Moves (SAN)" and insert "Clocks" after it
    sheet.getRange(1, legacyMovesIdx + 1).setValue('Moves (SAN)');
    sheet.insertColumnsAfter(legacyMovesIdx + 1, 1);
    sheet.getRange(1, legacyMovesIdx + 2).setValue('Clocks');
    return;
  }

  // Neither new nor legacy headers found; append both at the end
  const currentLastCol = sheet.getLastColumn();
  sheet.insertColumnsAfter(currentLastCol, 2);
  sheet.getRange(1, currentLastCol + 1, 1, 2).setValues([
    ['Moves (SAN)', 'Clocks']
  ]);
}

/**
 * Fetches live game callback data for a given Chess.com gameId and tries to
 * obtain a PGN string. Returns empty string if not found.
 * @param {string} gameId
 * @return {string}
 */
function fetchCallbackGamePgn(gameId) {
  if (!gameId) return '';
  var url = 'https://www.chess.com/callback/live/game/' + encodeURIComponent(String(gameId));
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
 * Backfills missing "Moves (SAN)" and "Clocks" for existing rows by
 * querying https://www.chess.com/callback/live/game/{game-id}.
 * Optionally provide a limit of how many missing rows to process this run.
 *
 * @param {number=} limit Optional max number of rows to process
 */
function backfillMovesAndClocks(limit) {
  const sheet = getOrCreateSheet();
  ensureMovesClocksHeaders();

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
  if (!colMovesSan || !colClocks) return; // headers must exist

  var numRows = lastRow - 1;
  var gameIds = sheet.getRange(2, colGameId, numRows, 1).getValues();
  var urls = sheet.getRange(2, colUrl, numRows, 1).getValues();
  var movesSanVals = sheet.getRange(2, colMovesSan, numRows, 1).getValues();
  var clocksVals = sheet.getRange(2, colClocks, numRows, 1).getValues();

  var toProcessCount = 0;
  for (var r = 0; r < numRows; r++) {
    var hasMoves = String(movesSanVals[r][0] || '').trim() !== '';
    var hasClocks = String(clocksVals[r][0] || '').trim() !== '';
    if (!hasMoves || !hasClocks) toProcessCount++;
  }
  if (typeof limit === 'number' && isFinite(limit) && limit >= 0) {
    // noop, we will enforce during loop
  } else {
    limit = toProcessCount; // process all missing by default
  }

  for (var i = 0, processed = 0; i < numRows && processed < limit; i++) {
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
    movesSanVals[i][0] = parsed.movesSan || currentMoves;
    clocksVals[i][0] = parsed.clocks || currentClocks;
    processed++;
  }

  // Batch update the entire two columns at once
  sheet.getRange(2, colMovesSan, numRows, 1).setValues(movesSanVals);
  sheet.getRange(2, colClocks, numRows, 1).setValues(clocksVals);
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

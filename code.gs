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
      'Rated',
      'Time class',
      'Time control',
      'Rules',
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
      'Moves'
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
    details.rated ? 'Yes' : 'No',     // Rated
    details.timeClass,                // Time class
    details.timeControl,              // Time control
    details.rules,                    // Rules
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
    details.moves || ''               // Moves (PGN movetext)
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
  var result = determineResult(side, opponent);
  var reason = determineReason(side, opponent);

  var fen = game.fen || '';
  var fenEncoded = encodeURIComponent(fen);
  var imageUrl = 'https://www.chess.com/dynboard?fen=' + fenEncoded + '&board=brown&piece=neo&size=3';

  var endTime = game.end_time ? new Date(game.end_time * 1000) : new Date();

  // Parse PGN tags for ECO, Opening URL, and get movetext (moves)
  var eco = '';
  var openingUrl = '';
  var moves = '';
  if (game.pgn) {
    try {
      var pgn = String(game.pgn);
      // Extract tags like [ECO "B01"] and [OpeningUrl "..."] or [Opening "..."]
      var ecoMatch = pgn.match(/\n\[ECO\s+"([^"]+)"\]/);
      if (ecoMatch && ecoMatch[1]) eco = ecoMatch[1];
      var openingUrlMatch = pgn.match(/\n\[OpeningUrl\s+"([^"]+)"\]/);
      if (openingUrlMatch && openingUrlMatch[1]) openingUrl = openingUrlMatch[1];
      // Extract movetext: content after the last closing bracket line of headers
      var splitIndex = pgn.lastIndexOf(']\n');
      if (splitIndex !== -1) {
        var afterHeaders = pgn.substring(splitIndex + 2);
        // Remove leading newlines/spaces
        afterHeaders = afterHeaders.replace(/^\s+/, '');
        // Remove result token at end (e.g., 1-0, 0-1, 1/2-1/2, *), if present
        moves = afterHeaders.replace(/\s+(1-0|0-1|1\/2-1\/2|\*)\s*$/m, '');
      }
    } catch (e) {
      // ignore PGN parse issues
    }
  }

  // Prefer ECO-based URL for Opening URL when ECO is available
  if (eco) {
    openingUrl = 'https://www.365chess.com/eco/' + encodeURIComponent(eco);
  } else if (!openingUrl && game.opening_url) {
    // Fallback to any opening URL provided by the source if ECO is missing
    openingUrl = game.opening_url;
  }

  return {
    isPlayerInGame: true,
    timestamp: endTime,
    url: game.url || '',
    rated: rated,
    timeClass: timeClass,
    timeControl: timeControl,
    rules: rules,
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
    moves: moves
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

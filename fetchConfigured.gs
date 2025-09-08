// Fetch pipeline for configured runs (V2)

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

// Back-compat: legacy menu handler name used in other files
function runConfiguredFetch() {
  return runConfiguredFetchV2();
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
  // Additional JSON fields needed for derived computations during recalc
  var idxEndTime = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'end_time'; });
  var idxRules = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'rules'; });
  var idxTimeClass = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'time_class'; });
  var idxFen = selectedHeaders.findIndex(function(h) { return h.source === 'json' && h.field === 'fen'; });

  var values = gamesSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var derivedReg = getDerivedRegistry_();

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    // Minimal inputs for compute()
    var game = {};
    if (idxTimeControl !== -1) game.time_control = row[idxTimeControl];
    if (idxEndTime !== -1) game.end_time = row[idxEndTime];
    if (idxRules !== -1) game.rules = row[idxRules];
    if (idxTimeClass !== -1) game.time_class = row[idxTimeClass];
    if (idxFen !== -1) game.fen = row[idxFen];
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


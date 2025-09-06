// =====================
// Opening Info Tab (Apps Script)
// =====================
// Creates a separate sheet tab with Opening/ECO info, Lichess links, and
// Opening Explorer summaries for the position after White's first move.

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Opening Info')
      .addItem('Build for selected row', 'buildOpeningInfoForSelectedRow')
      .addToUi();
  } catch (e) {
    // UI may not be available in some contexts
  }
}

function buildOpeningInfoForSelectedRow() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var games = ss.getSheetByName(SHEET_NAME);
  if (!games) throw new Error('Sheet not found: ' + SHEET_NAME);

  var active = games.getActiveRange();
  if (!active) throw new Error('Select a row in the "' + SHEET_NAME + '" sheet.');
  var row = active.getRow();
  if (row <= 1) throw new Error('Select a data row (row > 1).');

  var headerToIndex = getHeaderToIndexMap_(games);
  var colUrl = headerToIndex['URL'] || 2;
  var colGameId = headerToIndex['Game ID'] || 3;
  var colEco = headerToIndex['ECO'];
  var colOpeningUrl = headerToIndex['Opening URL'];
  var colMovesSan = headerToIndex['Moves (SAN)'];

  var sourceUrl = String(games.getRange(row, colUrl).getValue() || '');
  var gameId = String(games.getRange(row, colGameId).getValue() || '');
  var ecoExisting = colEco ? String(games.getRange(row, colEco).getValue() || '') : '';
  var openingUrlExisting = colOpeningUrl ? String(games.getRange(row, colOpeningUrl).getValue() || '') : '';
  var movesSanCell = colMovesSan ? String(games.getRange(row, colMovesSan).getValue() || '') : '';

  if (!gameId) {
    // Try extract from URL
    var idMatch = sourceUrl.match(/\/(\d+)(?:\?.*)?$/);
    if (idMatch && idMatch[1]) gameId = idMatch[1];
  }
  if (!gameId) throw new Error('Game ID not found in selected row.');

  // Fetch PGN via Chess.com callback, then parse ECO/opening and first move
  var pgn = fetchCallbackGamePgn(gameId);
  if (!pgn) throw new Error('PGN not available for gameId ' + gameId);

  var parsedTags = parseEcoOpeningVariationFromPgn_(pgn);
  var eco = parsedTags.eco || ecoExisting || '';
  var openingName = parsedTags.opening || '';
  var variation = parsedTags.variation || '';
  var openingUrl = parsedTags.openingUrl || openingUrlExisting || '';

  var firstSan = extractFirstWhiteSanFromPgn_(pgn);
  if (!firstSan && movesSanCell) firstSan = extractFirstWhiteSanFromMovesCell_(movesSanCell);
  var firstUci = firstSan ? sanFromStartToUci_(firstSan) : '';

  // Import to Lichess to get a permanent analysis URL
  var lichessLink = '';
  try {
    var imported = lichessImportGameFromPgn(pgn, { source: 'Apps Script' });
    lichessLink = imported && imported.url ? String(imported.url) : '';
  } catch (e) {
    lichessLink = '';
  }

  // Fetch Opening Explorer summaries (Masters + Lichess) for position after first move if available
  var mastersSummary = null;
  var lichessSummary = null;
  if (firstUci) {
    mastersSummary = fetchExplorerSummaryAfterPlay_('master', [firstUci]);
    lichessSummary = fetchExplorerSummaryAfterPlay_('lichess', [firstUci]);
  }

  // Write to the Opening Info tab
  var out = ensureOpeningInfoSheet_();
  out.clearContents();

  // Header row with compact fields
  out.getRange(1, 1, 1, 10).setValues([[
    'Source URL',
    'Game ID',
    'Opening Name',
    'ECO',
    'Variation',
    'Opening URL',
    'Lichess Game URL',
    'First move (SAN)',
    'First move (UCI)',
    'Note'
  ]]);
  out.getRange(2, 1, 1, 10).setValues([[
    sourceUrl,
    gameId,
    openingName,
    eco,
    variation,
    openingUrl,
    lichessLink,
    firstSan || '',
    firstUci || '',
    (firstUci ? '' : 'First move could not be parsed to UCI; explorer stats limited')
  ]]);

  var r = 4;
  // Masters block
  if (mastersSummary) {
    r = writeExplorerBlock_(out, r, 'Masters DB Summary', mastersSummary);
  } else {
    out.getRange(r++, 1).setValue('Masters DB Summary: (unavailable)');
  }

  r++;
  // Lichess block
  if (lichessSummary) {
    r = writeExplorerBlock_(out, r, 'Lichess DB Summary', lichessSummary);
  } else {
    out.getRange(r++, 1).setValue('Lichess DB Summary: (unavailable)');
  }
}

function ensureOpeningInfoSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var name = 'Opening Info';
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function parseEcoOpeningVariationFromPgn_(pgn) {
  var res = { eco: '', opening: '', variation: '', openingUrl: '' };
  if (!pgn) return res;
  try {
    var normalized = String(pgn).replace(/\r\n/g, '\n');
    var eco = normalized.match(/^\[ECO\s+"([^"]+)"\]/m);
    if (eco && eco[1]) res.eco = eco[1];
    var opening = normalized.match(/^\[Opening\s+"([^"]+)"\]/m);
    if (opening && opening[1]) res.opening = opening[1];
    var variation = normalized.match(/^\[Variation\s+"([^"]+)"\]/m);
    if (variation && variation[1]) res.variation = variation[1];
    var urlTag = normalized.match(/^\[(?:ECOUrl|OpeningUrl)\s+"([^"]+)"\]/m);
    if (urlTag && urlTag[1]) res.openingUrl = urlTag[1];
  } catch (e) {}
  return res;
}

function extractFirstWhiteSanFromPgn_(pgn) {
  if (!pgn) return '';
  var normalized = String(pgn).replace(/\r\n/g, '\n');
  var headerEnd = normalized.indexOf('\n\n');
  var body = headerEnd !== -1 ? normalized.substring(headerEnd + 2) : normalized;
  body = body.replace(/^\s+/, '');
  var m = body.match(/1\.\s*([^\s\{\)]+)(?:\s|\{|\)|$)/);
  if (!m || !m[1]) return '';
  return sanitizeSanToken_(m[1]);
}

function extractFirstWhiteSanFromMovesCell_(movesCell) {
  // movesCell like: {e4, e5, Nf3, Nc6, Bb5, ...}
  var s = String(movesCell || '').trim();
  if (!s || s[0] !== '{' || s[s.length - 1] !== '}') return '';
  var inner = s.substring(1, s.length - 1);
  var parts = inner.split(',').map(function(t) { return String(t).trim(); });
  return parts.length ? sanitizeSanToken_(parts[0]) : '';
}

function sanitizeSanToken_(san) {
  // Remove trailing check/mate and NAG symbols
  return String(san || '').replace(/[+#!?]+$/g, '');
}

function sanFromStartToUci_(san) {
  var t = String(san || '').trim();
  if (!t) return '';
  // Castling not possible on first move
  // Knight moves: Nf3, Nc3, Na3, Nh3, Nd2
  if (/^N[a-h][1-8]$/.test(t)) {
    var dest = t.substring(1);
    var map = {
      'a3': 'b1a3', 'c2': 'b4?','c3': 'b1c3', 'd2': 'b1d2',
      'f2': 'g4?', 'f3': 'g1f3', 'h2': 'g4?', 'h3': 'g1h3'
    };
    if (map[dest]) return (map[dest] === 'b4?' || map[dest] === 'g4?') ? '' : map[dest];
    // Conservative: handle known legal ones
    if (dest === 'c3') return 'b1c3';
    if (dest === 'a3') return 'b1a3';
    if (dest === 'f3') return 'g1f3';
    if (dest === 'h3') return 'g1h3';
    if (dest === 'd2') return 'b1d2';
    return '';
  }
  // Pawn pushes: e4, e3, etc.
  if (/^[a-h][1-8]$/.test(t)) {
    var file = t.charAt(0);
    var rank = t.charAt(1);
    if (rank === '3' || rank === '4') return file + '2' + file + rank;
    return '';
  }
  // Captures on move 1 are impossible; others unsupported
  return '';
}

function fetchExplorerSummaryAfterPlay_(which, uciMoves) {
  // which: 'master' or 'lichess'
  var base = (which === 'master') ? 'https://explorer.lichess.ovh/master' : 'https://explorer.lichess.ovh/lichess';
  var params = [];
  if (uciMoves && uciMoves.length) params.push('play=' + encodeURIComponent(uciMoves.join(',')));
  params.push('moves=12');
  params.push('topGames=0');
  if (which === 'lichess') {
    params.push('variant=standard');
    params.push('speeds=blitz,rapid,classical');
    params.push('ratings=1600,1800,2000,2200');
  }
  var url = base + '?' + params.join('&');
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'Accept': 'application/json' } });
  if (resp.getResponseCode() !== 200) return null;
  var json = {};
  try { json = JSON.parse(resp.getContentText()); } catch (e) { return null; }
  return summarizeExplorerJson_(json);
}

function summarizeExplorerJson_(json) {
  if (!json || typeof json !== 'object') return null;
  var total = Number(json.white || 0) + Number(json.draws || 0) + Number(json.black || 0);
  var sideToMove = 'black'; // after White's first move
  function pct(n) { return total > 0 ? Math.round((n / total) * 1000) / 10 : 0; }
  function scoreForSideToMove(white, draws, black) {
    if (sideToMove === 'white') return ((white + draws * 0.5) / (white + draws + black)) * 100;
    return ((black + draws * 0.5) / (white + draws + black)) * 100;
  }
  var moves = Array.isArray(json.moves) ? json.moves.slice() : [];
  var top = moves.map(function(m) {
    var w = Number(m.white || 0), d = Number(m.draws || 0), b = Number(m.black || 0);
    var tot = w + d + b;
    var score = tot > 0 ? Math.round(scoreForSideToMove(w, d, b) * 10) / 10 : 0;
    return {
      san: String(m.san || ''),
      uci: String(m.uci || ''),
      games: tot,
      score: score,
      avg: (m.averageRating != null ? Number(m.averageRating) : null),
      last: (m.lastPlayed || null)
    };
  });
  // Sort by games desc
  top.sort(function(a, b) { return b.games - a.games; });
  return {
    totals: {
      games: total,
      white: Number(json.white || 0),
      draws: Number(json.draws || 0),
      black: Number(json.black || 0),
      whitePct: pct(Number(json.white || 0)),
      drawPct: pct(Number(json.draws || 0)),
      blackPct: pct(Number(json.black || 0))
    },
    topMoves: top.slice(0, 12)
  };
}

function writeExplorerBlock_(sheet, startRow, title, summary) {
  var r = startRow;
  sheet.getRange(r++, 1).setValue(title);
  sheet.getRange(r, 1, 1, 7).setValues([[
    'Total games', 'White', 'Draws', 'Black', 'White %', 'Draw %', 'Black %'
  ]]);
  sheet.getRange(r + 1, 1, 1, 7).setValues([[
    summary.totals.games,
    summary.totals.white,
    summary.totals.draws,
    summary.totals.black,
    summary.totals.whitePct,
    summary.totals.drawPct,
    summary.totals.blackPct
  ]]);
  r += 3;
  sheet.getRange(r++, 1).setValue('Top moves from this position (side to move: Black)');
  sheet.getRange(r, 1, 1, 5).setValues([[ 'SAN', 'UCI', 'Games', 'Score% (STM)', 'Avg Elo' ]]);
  var rows = summary.topMoves.map(function(m) {
    return [ m.san, m.uci, m.games, m.score, (m.avg != null ? m.avg : '') ];
  });
  if (rows.length) sheet.getRange(r + 1, 1, rows.length, 5).setValues(rows);
  r += (rows.length + 2);
  return r;
}


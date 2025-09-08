// Opening Info Data Group & Derived Headers

// Lightweight in-run memoization to avoid repeated API calls per game
var OPENINFO_MEMO = {
  first: {}, // by key -> { firstSan, firstUci, afterFirstFen }
  explorer: {}, // by key -> { masters: summary, lichdb: summary }
  cloud: {}, // by key+fen -> { evalText, depth, pv }
  import: {} // by key -> { url }
};

/** Returns whether Opening Info computations are enabled in Config. */
function isOpeningInfoEnabled_() {
  try {
    var cfg = readConfig_();
    var g = cfg && cfg.groups && cfg.groups.openingInfo;
    return !!(g && (g.calculateNew || g.recalculate));
  } catch (e) {
    return false;
  }
}

/** Build derived registry entries for Opening Info. Merged into getDerivedRegistry_(). */
function getOpeningInfoDerivedRegistry_() {
  var reg = {};

  function guard(fn) {
    return function(game, pgnTags, pgnMoves) {
      if (!isOpeningInfoEnabled_()) return '';
      try { return fn(game, pgnTags, pgnMoves); } catch (e) { return ''; }
    };
  }

  // Basic opening outputs
  reg.openinfo_opening_name = {
    displayName: 'Opening Name', description: 'PGN [Opening] tag', example: 'Sicilian Defense',
    compute: guard(function(game, pgnTags) { return (pgnTags && pgnTags['Opening']) || ''; })
  };
  reg.openinfo_eco = {
    displayName: 'ECO', description: 'PGN [ECO] code', example: 'B90',
    compute: guard(function(game, pgnTags) { return (pgnTags && pgnTags['ECO']) || ''; })
  };
  reg.openinfo_variation = {
    displayName: 'Variation', description: 'PGN [Variation] tag', example: 'Najdorf Variation',
    compute: guard(function(game, pgnTags) { return (pgnTags && pgnTags['Variation']) || ''; })
  };
  reg.openinfo_opening_url = {
    displayName: 'Opening URL', description: 'From PGN ECOUrl/OpeningUrl or JSON opening_url', example: 'https://www.chess.com/openings/...',
    compute: guard(function(game, pgnTags) {
      var v = (pgnTags && (pgnTags['ECOUrl'] || pgnTags['OpeningUrl'])) || (game && game.opening_url) || '';
      return v || '';
    })
  };

  // First move SAN/UCI
  reg.openinfo_first_san = {
    displayName: 'First move (SAN)', description: 'White first move SAN parsed from PGN', example: 'e4',
    compute: guard(function(game) {
      var key = openingKey_(game);
      var base = ensureOpeningBase_(game);
      OPENINFO_MEMO.first[key] = base;
      return base.firstSan || '';
    })
  };
  reg.openinfo_first_uci = {
    displayName: 'First move (UCI)', description: 'White first move UCI from start', example: 'e2e4',
    compute: guard(function(game) {
      var key = openingKey_(game);
      var base = ensureOpeningBase_(game);
      OPENINFO_MEMO.first[key] = base;
      return base.firstUci || '';
    })
  };

  // Lichess import link
  reg.openinfo_lichess_game_url = {
    displayName: 'Lichess Game URL', description: 'Anonymous import URL for PGN', example: 'https://lichess.org/study/..',
    compute: guard(function(game) {
      var key = openingKey_(game);
      if (OPENINFO_MEMO.import[key]) return OPENINFO_MEMO.import[key].url || '';
      var pgn = (game && game.pgn) || '';
      if (!pgn) return '';
      var imp = lichessImportGameFromPgnOpening_(pgn);
      OPENINFO_MEMO.import[key] = { url: imp && imp.url ? String(imp.url) : '' };
      return OPENINFO_MEMO.import[key].url || '';
    })
  };

  // Explorer summaries (Masters and Lichess)
  function getExplorer_(game) {
    var key = openingKey_(game);
    if (OPENINFO_MEMO.explorer[key]) return OPENINFO_MEMO.explorer[key];
    var base = ensureOpeningBase_(game);
    if (!base.firstUci) { OPENINFO_MEMO.explorer[key] = { masters: null, lichdb: null }; return OPENINFO_MEMO.explorer[key]; }
    var masters = fetchExplorerSummaryOpening_('master', [base.firstUci]);
    var lichdb = fetchExplorerSummaryOpening_('lichess', [base.firstUci]);
    OPENINFO_MEMO.explorer[key] = { masters: masters, lichdb: lichdb };
    return OPENINFO_MEMO.explorer[key];
  }

  function top1_(summary) {
    if (!summary || !summary.topMoves || !summary.topMoves.length) return null;
    return summary.topMoves[0];
  }

  // Masters totals
  reg.openinfo_masters_total_games = { displayName: 'Masters Total Games', description: 'Explorer total games after 1st move', example: 12345,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.masters && ex.masters.totals ? ex.masters.totals.games : ''; }) };
  reg.openinfo_masters_white_pct = { displayName: 'Masters White %', description: 'Explorer white%', example: 35.2,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.masters && ex.masters.totals ? ex.masters.totals.whitePct : ''; }) };
  reg.openinfo_masters_draw_pct = { displayName: 'Masters Draw %', description: 'Explorer draw%', example: 32.1,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.masters && ex.masters.totals ? ex.masters.totals.drawPct : ''; }) };
  reg.openinfo_masters_black_pct = { displayName: 'Masters Black %', description: 'Explorer black%', example: 32.7,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.masters && ex.masters.totals ? ex.masters.totals.blackPct : ''; }) };
  reg.openinfo_masters_top1_san = { displayName: 'Masters Top1 SAN', description: 'Most played reply SAN', example: '...Nf6',
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.masters); return t ? t.san : ''; }) };
  reg.openinfo_masters_top1_uci = { displayName: 'Masters Top1 UCI', description: 'Most played reply UCI', example: 'g8f6',
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.masters); return t ? t.uci : ''; }) };
  reg.openinfo_masters_top1_games = { displayName: 'Masters Top1 Games', description: 'Games on top move', example: 5432,
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.masters); return t ? t.games : ''; }) };
  reg.openinfo_masters_top1_score = { displayName: 'Masters Top1 Score% (STM)', description: 'Score for side to move', example: 47.3,
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.masters); return t ? t.score : ''; }) };
  reg.openinfo_masters_top1_avg_elo = { displayName: 'Masters Top1 Avg Elo', description: 'Average rating if available', example: 2440,
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.masters); return t && t.avg != null ? t.avg : ''; }) };

  // Lichess DB totals
  reg.openinfo_lichdb_total_games = { displayName: 'Lichess DB Total Games', description: 'Explorer total games after 1st move', example: 98765,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.games : ''; }) };
  reg.openinfo_lichdb_white_pct = { displayName: 'Lichess DB White %', description: 'Explorer white%', example: 49.8,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.whitePct : ''; }) };
  reg.openinfo_lichdb_draw_pct = { displayName: 'Lichess DB Draw %', description: 'Explorer draw%', example: 6.1,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.drawPct : ''; }) };
  reg.openinfo_lichdb_black_pct = { displayName: 'Lichess DB Black %', description: 'Explorer black%', example: 44.1,
    compute: guard(function(game){ var ex = getExplorer_(game); return ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.blackPct : ''; }) };
  reg.openinfo_lichdb_top1_san = { displayName: 'Lichess DB Top1 SAN', description: 'Most played reply SAN', example: '...Nc6',
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.lichdb); return t ? t.san : ''; }) };
  reg.openinfo_lichdb_top1_uci = { displayName: 'Lichess DB Top1 UCI', description: 'Most played reply UCI', example: 'b8c6',
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.lichdb); return t ? t.uci : ''; }) };
  reg.openinfo_lichdb_top1_games = { displayName: 'Lichess DB Top1 Games', description: 'Games on top move', example: 12345,
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.lichdb); return t ? t.games : ''; }) };
  reg.openinfo_lichdb_top1_score = { displayName: 'Lichess DB Top1 Score% (STM)', description: 'Score for side to move', example: 51.2,
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.lichdb); return t ? t.score : ''; }) };
  reg.openinfo_lichdb_top1_avg_elo = { displayName: 'Lichess DB Top1 Avg Elo', description: 'Average rating if available', example: 1840,
    compute: guard(function(game){ var ex = getExplorer_(game); var t = top1_(ex.lichdb); return t && t.avg != null ? t.avg : ''; }) };

  // Cloud eval
  function getEval_(game, fen) {
    if (!fen) return { evalText: '', depth: '', pv: '' };
    var key = openingKey_(game) + '|fen|' + String(fen);
    if (OPENINFO_MEMO.cloud[key]) return OPENINFO_MEMO.cloud[key];
    var json = lichessCloudEvalFenOpening_(fen);
    var s = summarizeCloudEvalOpening_(json);
    OPENINFO_MEMO.cloud[key] = s;
    return s;
  }

  reg.openinfo_eval_start_cp = { displayName: 'Eval Start (cp/mate)', description: 'Lichess Cloud Eval for start position', example: 0,
    compute: guard(function(game){ var s = getEval_(game, getStartFenOpening_()); return s.evalText; }) };
  reg.openinfo_eval_start_depth = { displayName: 'Eval Start Depth', description: 'Depth', example: 33,
    compute: guard(function(game){ var s = getEval_(game, getStartFenOpening_()); return s.depth; }) };
  reg.openinfo_eval_start_pv = { displayName: 'Eval Start PV (UCI)', description: 'Principal variation', example: 'e2e4 e7e5 g1f3',
    compute: guard(function(game){ var s = getEval_(game, getStartFenOpening_()); return s.pv; }) };

  reg.openinfo_eval_after_first_cp = { displayName: 'Eval After 1st (cp/mate)', description: 'After white first move', example: 22,
    compute: guard(function(game){ var base = ensureOpeningBase_(game); var s = getEval_(game, base.afterFirstFen); return s.evalText; }) };
  reg.openinfo_eval_after_first_depth = { displayName: 'Eval After 1st Depth', description: 'Depth', example: 32,
    compute: guard(function(game){ var base = ensureOpeningBase_(game); var s = getEval_(game, base.afterFirstFen); return s.depth; }) };
  reg.openinfo_eval_after_first_pv = { displayName: 'Eval After 1st PV (UCI)', description: 'Principal variation', example: 'g8f6 b1c3',
    compute: guard(function(game){ var base = ensureOpeningBase_(game); var s = getEval_(game, base.afterFirstFen); return s.pv; }) };

  reg.openinfo_eval_final_cp = { displayName: 'Eval Final (cp/mate)', description: 'Final FEN from game JSON', example: -150,
    compute: guard(function(game){ var fen = (game && game.fen) || ''; var s = getEval_(game, fen); return s.evalText; }) };
  reg.openinfo_eval_final_depth = { displayName: 'Eval Final Depth', description: 'Depth', example: 28,
    compute: guard(function(game){ var fen = (game && game.fen) || ''; var s = getEval_(game, fen); return s.depth; }) };
  reg.openinfo_eval_final_pv = { displayName: 'Eval Final PV (UCI)', description: 'Principal variation', example: '...',
    compute: guard(function(game){ var fen = (game && game.fen) || ''; var s = getEval_(game, fen); return s.pv; }) };

  // Note
  reg.openinfo_note = {
    displayName: 'Opening Info Note', description: 'Warnings/notes from parser', example: 'First move could not be parsed to UCI',
    compute: guard(function(game) {
      var base = ensureOpeningBase_(game);
      if (!base.firstUci) return 'First move could not be parsed to UCI; explorer stats limited';
      return '';
    })
  };

  return reg;
}

// ---------- Helpers ----------

function openingKey_(game) {
  var u = (game && game.url) || '';
  if (u) return String(u);
  var p = (game && game.pgn) || '';
  if (!p) return String(Math.random());
  return 'pgn:' + String(p).slice(0, 64);
}

function ensureOpeningBase_(game) {
  var key = openingKey_(game);
  if (OPENINFO_MEMO.first[key]) return OPENINFO_MEMO.first[key];
  var pgn = (game && game.pgn) || '';
  var firstSan = extractFirstWhiteSanFromPgnOpening_(pgn);
  var firstUci = firstSan ? sanFromStartToUciOpening_(firstSan) : '';
  var afterFirstFen = firstUci ? fenAfterFirstMoveFromStartOpening_(firstUci) : '';
  var out = { firstSan: firstSan || '', firstUci: firstUci || '', afterFirstFen: afterFirstFen || '' };
  OPENINFO_MEMO.first[key] = out;
  return out;
}

function extractFirstWhiteSanFromPgnOpening_(pgn) {
  if (!pgn) return '';
  var normalized = String(pgn).replace(/\r\n/g, '\n');
  var headerEnd = normalized.indexOf('\n\n');
  var body = headerEnd !== -1 ? normalized.substring(headerEnd + 2) : normalized;
  body = body.replace(/^\s+/, '');
  var m = body.match(/1\.\s*([^\s\{\)]+)(?:\s|\{|\)|$)/);
  if (!m || !m[1]) return '';
  return sanitizeSanTokenOpening_(m[1]);
}

function sanitizeSanTokenOpening_(san) {
  return String(san || '').replace(/[+#!?]+$/g, '');
}

function sanFromStartToUciOpening_(san) {
  var t = String(san || '').trim();
  if (!t) return '';
  // Knight moves: Nf3, Nc3, Na3, Nh3, Nd2
  if (/^N[a-h][1-8]$/.test(t)) {
    var dest = t.substring(1);
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
  // Captures/castling unsupported on move 1 in this simplified mapper
  return '';
}

function getStartFenOpening_() {
  return 'rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1';
}

function fenAfterFirstMoveFromStartOpening_(uci) {
  var t = String(uci || '').trim();
  if (!t || t.length !== 4) return '';
  var from = t.substring(0, 2);
  var to = t.substring(2, 4);
  var fileFrom = from.charAt(0), rankFrom = from.charAt(1);
  var fileTo = to.charAt(0), rankTo = to.charAt(1);

  var board = {};
  'abcdefgh'.split('').forEach(function(f) { board[f + '2'] = 'P'; board[f + '7'] = 'p'; });
  board['b1'] = 'N'; board['g1'] = 'N'; board['b8'] = 'n'; board['g8'] = 'n';

  var moving = board[from];
  if (!moving) return '';
  if (moving === 'P') {
    if (rankFrom !== '2') return '';
    if (!(rankTo === '3' || rankTo === '4')) return '';
    if (fileFrom !== fileTo) return '';
  } else if (moving === 'N') {
    var legalKnightTargets = { 'b1': { 'a3': true, 'c3': true }, 'g1': { 'f3': true, 'h3': true } };
    if (!legalKnightTargets[from] || !legalKnightTargets[from][to]) return '';
  } else {
    return '';
  }

  board[to] = moving;
  delete board[from];

  function pieceAt(sq) { return board[sq] || null; }
  function makeRank(rank) {
    var empties = 0, out = '';
    'abcdefgh'.split('').forEach(function(f) {
      var p = pieceAt(f + rank);
      if (p) { if (empties) { out += String(empties); empties = 0; } out += p; }
      else { empties++; }
    });
    if (empties) out += String(empties);
    return out;
  }
  var placement = [ '8','7','6','5','4','3','2','1' ].map(makeRank).join('/');
  var stm = 'b';
  var castling = 'KQkq';
  var ep = '-';
  if (moving === 'P' && rankFrom === '2' && rankTo === '4') ep = fileFrom + '3';
  var halfmove = 0;
  var fullmove = 1;
  return placement + ' ' + stm + ' ' + castling + ' ' + ep + ' ' + halfmove + ' ' + fullmove;
}

function fetchExplorerSummaryOpening_(which, uciMoves) {
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
  return summarizeExplorerJsonOpening_(json);
}

function summarizeExplorerJsonOpening_(json) {
  if (!json || typeof json !== 'object') return null;
  var total = Number(json.white || 0) + Number(json.draws || 0) + Number(json.black || 0);
  var sideToMove = 'black';
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
    return { san: String(m.san || ''), uci: String(m.uci || ''), games: tot, score: score, avg: (m.averageRating != null ? Number(m.averageRating) : null), last: (m.lastPlayed || null) };
  });
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

function lichessCloudEvalFenOpening_(fen) {
  if (!fen) return null;
  var url = 'https://lichess.org/api/cloud-eval?multiPv=1&fen=' + encodeURIComponent(String(fen));
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'Accept': 'application/json' } });
  if (resp.getResponseCode() !== 200) return null;
  try { return JSON.parse(resp.getContentText()); } catch (e) { return null; }
}

function summarizeCloudEvalOpening_(json) {
  if (!json || !json.pvs || !json.pvs.length) return { evalText: '', depth: '', pv: '' };
  var pv = json.pvs[0] || {};
  var depth = (pv.depth != null) ? pv.depth : '';
  var evalText = '';
  if (pv.mate != null) {
    evalText = 'M' + pv.mate;
  } else if (pv.cp != null) {
    evalText = String(pv.cp);
  }
  var pvMoves = Array.isArray(pv.moves) ? pv.moves.join(' ') : (pv.moves || '');
  return { evalText: evalText, depth: depth, pv: pvMoves };
}

function lichessImportGameFromPgnOpening_(pgn) {
  if (!pgn) return null;
  var url = 'https://lichess.org/api/import';
  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { pgn: String(pgn), source: 'Apps Script' },
    muteHttpExceptions: true
  };
  var resp = UrlFetchApp.fetch(url, options);
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) return null;
  try { return JSON.parse(resp.getContentText()); } catch (e) { return null; }
}


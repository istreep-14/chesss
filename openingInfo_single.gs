// Single-PGN runner for Opening Info derived fields

/** Parse PGN header tags into a simple map. */
function parsePgnTagsSimple_(pgn) {
  if (!pgn) return {};
  var tags = {};
  var lines = String(pgn).replace(/\r\n/g, '\n').split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line || line.charAt(0) !== '[') break;
    var m = line.match(/^\[(\w+)\s+"([\s\S]*?)"\]$/);
    if (m && m[1]) tags[m[1]] = m[2] || '';
  }
  return tags;
}

/** Internal: get explorer summaries for a game based on first UCI. */
function getExplorerForSinglePgn_(game) {
  var key = openingKey_(game);
  if (OPENINFO_MEMO.explorer[key]) return OPENINFO_MEMO.explorer[key];
  var base = ensureOpeningBase_(game);
  if (!base.firstUci) { OPENINFO_MEMO.explorer[key] = { masters: null, lichdb: null }; return OPENINFO_MEMO.explorer[key]; }
  var masters = fetchExplorerSummaryOpening_('master', [base.firstUci]);
  var lichdb = fetchExplorerSummaryOpening_('lichess', [base.firstUci]);
  OPENINFO_MEMO.explorer[key] = { masters: masters, lichdb: lichdb };
  return OPENINFO_MEMO.explorer[key];
}

function top1Single_(summary) {
  if (!summary || !summary.topMoves || !summary.topMoves.length) return null;
  return summary.topMoves[0];
}

/** Internal: get cloud eval summary for a given FEN and game memo key. */
function getEvalForSinglePgn_(game, fen) {
  if (!fen) return { evalText: '', depth: '', pv: '' };
  var key = openingKey_(game) + '|fen|' + String(fen);
  if (OPENINFO_MEMO.cloud[key]) return OPENINFO_MEMO.cloud[key];
  var json = lichessCloudEvalFenOpening_(fen);
  var s = summarizeCloudEvalOpening_(json);
  OPENINFO_MEMO.cloud[key] = s;
  return s;
}

/**
 * Compute all Opening Info fields for a single PGN.
 * Returns an object keyed by the same field ids used in the data group.
 */
function computeOpeningInfoForPgn(pgn) {
  var pgnTags = parsePgnTagsSimple_(pgn);
  var game = { pgn: String(pgn || ''), url: '', fen: '' };

  // Basic opening outputs
  var opening_name = (pgnTags && pgnTags['Opening']) || '';
  var eco = (pgnTags && pgnTags['ECO']) || '';
  var variation = (pgnTags && pgnTags['Variation']) || '';
  var opening_url = ((pgnTags && (pgnTags['ECOUrl'] || pgnTags['OpeningUrl'])) || (game && game.opening_url) || '') || '';

  // First move SAN/UCI
  var base = ensureOpeningBase_(game);
  OPENINFO_MEMO.first[openingKey_(game)] = base;

  // Lichess import link
  var importUrl = '';
  try {
    var imp = lichessImportGameFromPgnOpening_(game.pgn);
    importUrl = imp && imp.url ? String(imp.url) : '';
    OPENINFO_MEMO.import[openingKey_(game)] = { url: importUrl };
  } catch (e) {}

  // Explorer summaries
  var ex = getExplorerForSinglePgn_(game);

  // Cloud evals
  var evalStart = getEvalForSinglePgn_(game, getStartFenOpening_());
  var evalAfterFirst = getEvalForSinglePgn_(game, base.afterFirstFen);
  var evalFinal = getEvalForSinglePgn_(game, (game && game.fen) || '');

  // Note
  var note = base.firstUci ? '' : 'First move could not be parsed to UCI; explorer stats limited';

  return {
    openinfo_opening_name: opening_name,
    openinfo_eco: eco,
    openinfo_variation: variation,
    openinfo_opening_url: opening_url,
    openinfo_first_san: base.firstSan || '',
    openinfo_first_uci: base.firstUci || '',
    openinfo_lichess_game_url: importUrl,

    openinfo_masters_total_games: ex.masters && ex.masters.totals ? ex.masters.totals.games : '',
    openinfo_masters_white_pct: ex.masters && ex.masters.totals ? ex.masters.totals.whitePct : '',
    openinfo_masters_draw_pct: ex.masters && ex.masters.totals ? ex.masters.totals.drawPct : '',
    openinfo_masters_black_pct: ex.masters && ex.masters.totals ? ex.masters.totals.blackPct : '',
    openinfo_masters_top1_san: (function(){ var t = top1Single_(ex.masters); return t ? t.san : ''; })(),
    openinfo_masters_top1_uci: (function(){ var t = top1Single_(ex.masters); return t ? t.uci : ''; })(),
    openinfo_masters_top1_games: (function(){ var t = top1Single_(ex.masters); return t ? t.games : ''; })(),
    openinfo_masters_top1_score: (function(){ var t = top1Single_(ex.masters); return t ? t.score : ''; })(),
    openinfo_masters_top1_avg_elo: (function(){ var t = top1Single_(ex.masters); return (t && t.avg != null) ? t.avg : ''; })(),

    openinfo_lichdb_total_games: ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.games : '',
    openinfo_lichdb_white_pct: ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.whitePct : '',
    openinfo_lichdb_draw_pct: ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.drawPct : '',
    openinfo_lichdb_black_pct: ex.lichdb && ex.lichdb.totals ? ex.lichdb.totals.blackPct : '',
    openinfo_lichdb_top1_san: (function(){ var t = top1Single_(ex.lichdb); return t ? t.san : ''; })(),
    openinfo_lichdb_top1_uci: (function(){ var t = top1Single_(ex.lichdb); return t ? t.uci : ''; })(),
    openinfo_lichdb_top1_games: (function(){ var t = top1Single_(ex.lichdb); return t ? t.games : ''; })(),
    openinfo_lichdb_top1_score: (function(){ var t = top1Single_(ex.lichdb); return t ? t.score : ''; })(),
    openinfo_lichdb_top1_avg_elo: (function(){ var t = top1Single_(ex.lichdb); return (t && t.avg != null) ? t.avg : ''; })(),

    openinfo_eval_start_cp: evalStart.evalText,
    openinfo_eval_start_depth: evalStart.depth,
    openinfo_eval_start_pv: evalStart.pv,
    openinfo_eval_after_first_cp: evalAfterFirst.evalText,
    openinfo_eval_after_first_depth: evalAfterFirst.depth,
    openinfo_eval_after_first_pv: evalAfterFirst.pv,
    openinfo_eval_final_cp: evalFinal.evalText,
    openinfo_eval_final_depth: evalFinal.depth,
    openinfo_eval_final_pv: evalFinal.pv,

    openinfo_note: note
  };
}


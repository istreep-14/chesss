function fetchChessDotComMonth_(username, year, month) {
  var mm = month < 10 ? '0' + month : String(month);
  var url = 'https://api.chess.com/pub/player/' + encodeURIComponent(username) + '/games/' + year + '/' + mm;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Failed to fetch: ' + url + ' (' + resp.getResponseCode() + ')');
  }
  var json = JSON.parse(resp.getContentText());
  var games = json.games || [];
  var rows = [];
  for (var i = 0; i < games.length; i++) {
    var g = games[i];

    var white = g.white && (g.white.username || g.white);
    var black = g.black && (g.black.username || g.black);
    var userLower = String(username).toLowerCase();
    var color = (white && String(white).toLowerCase() === userLower) ? 'white' : ((black && String(black).toLowerCase() === userLower) ? 'black' : '');
    var opponent = color === 'white' ? (black || '') : (color === 'black' ? (white || '') : '');

    var endTime = g.end_time ? new Date(Number(g.end_time) * 1000) : null;
    var endTimeIso = endTime ? Utilities.formatDate(endTime, 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'") : '';

    var resultSimple = computeSimpleResult_(g, color);

    var eco = extractPgnTag_(g.pgn, 'ECO');
    var termination = extractPgnTag_(g.pgn, 'Termination') || g.termination || '';

    rows.push({
      source: 'chess.com',
      username: username,
      url: g.url || '',
      end_time: endTimeIso,
      opponent: opponent,
      user_color: color,
      result_simple: resultSimple,
      time_control: g.time_control || '',
      eco: eco || '',
      termination: String(termination || ''),
      pgn: g.pgn || ''
    });
  }
  return rows;
}

function computeSimpleResult_(game, userColor) {
  if (!userColor) return '';
  var whiteResult = game.white && game.white.result;
  var blackResult = game.black && game.black.result;

  function isDraw(r) {
    return r === 'agreed' || r === 'stalemate' || r === 'repetition' || r === 'insufficient' || r === 'timevsinsufficient' || r === '50move' || r === 'draw';
  }

  if (isDraw(whiteResult) || isDraw(blackResult)) return 'draw';

  if (userColor === 'white') {
    if (whiteResult === 'win') return 'win';
    if (blackResult === 'win') return 'loss';
  } else if (userColor === 'black') {
    if (blackResult === 'win') return 'win';
    if (whiteResult === 'win') return 'loss';
  }
  return '';
}

function extractPgnTag_(pgn, tag) {
  if (!pgn) return '';
  var re = new RegExp('\\[' + tag + ' "([^"]*)"\\]');
  var m = pgn.match(re);
  return m ? m[1] : '';
}
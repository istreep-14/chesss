// Apps Script utility to parse opening and analysis fields from a PGN export and
// populate OpeningFamily, OpeningVariation, OpeningSub1, OpeningSub2, OpeningECO
// plus WhiteAccuracy, BlackAccuracy, Result, and extended metrics computed from [%eval]s
// in a Google Sheet. This relies solely on PGN tags: [Opening "..."] and [ECO "..."]

const SCRIPT_PROP_MY_USERNAME_KEY = 'MY_USERNAME';

// Column indices (1-based)
const COL = {
  ANALYZED_PGN: 5,  // Column E: PGN that includes [Opening] and [ECO] tags
  OPEN_FAM: 6,      // Column F: OpeningFamily
  OPEN_VAR: 7,      // Column G: OpeningVariation
  OPEN_SUB1: 8,     // Column H: OpeningSub1
  OPEN_SUB2: 9,     // Column I: OpeningSub2
  OPEN_ECO: 10,     // Column J: OpeningECO
  WHITE_ACC: 11,    // Column K: WhiteAccuracy
  BLACK_ACC: 12,    // Column L: BlackAccuracy
  RESULT: 13,       // Column M: Result (e.g., 1-0, 0-1, 1/2-1/2)
  WHITE_ACPL: 14,   // Column N: WhiteACPL (approx from eval deltas)
  BLACK_ACPL: 15,   // Column O: BlackACPL
  WHITE_BEST_PCT: 16, // Column P: WhiteBestPct
  BLACK_BEST_PCT: 17, // Column Q: BlackBestPct
  WHITE_INACC: 18,  // Column R: WhiteInaccuracies
  WHITE_MIST: 19,   // Column S: WhiteMistakes
  WHITE_BLUN: 20,   // Column T: WhiteBlunders
  BLACK_INACC: 21,  // Column U: BlackInaccuracies
  BLACK_MIST: 22,   // Column V: BlackMistakes
  BLACK_BLUN: 23,   // Column W: BlackBlunders
  WHITE_MISSED_WINS: 24, // Column X: WhiteMissedWins (approx)
  BLACK_MISSED_WINS: 25, // Column Y: BlackMissedWins (approx)
  THEORY_LIKE_PLY: 26,    // Column Z: TheoryLikePly (heuristic)
  MY_USERNAME: 27,       // Column AA: MyUsername
  OPP_USERNAME: 28,      // Column AB: OppUsername
  MY_COLOR: 29,          // Column AC: MyColor (White/Black)
  MY_RATING: 30,         // Column AD: MyRating
  OPP_RATING: 31,        // Column AE: OppRating
  TIME_CONTROL: 32,      // Column AF: TimeControl (e.g., 300+0)
  INITIAL_SEC: 33,       // Column AG: InitialSec
  INCREMENT: 34,         // Column AH: Increment
  SPEED_CLASS: 35,       // Column AI: SpeedClass (bullet/blitz/rapid/classical)
  TERMINATION: 36,       // Column AJ: Termination
  MOVES_COUNT: 37,       // Column AK: MovesCount
  PLIES: 38,             // Column AL: Plies
  MY_AVG_MOVE_SEC: 39,   // Column AM: MyAvgMoveTimeSec
  OPP_AVG_MOVE_SEC: 40,  // Column AN: OppAvgMoveTimeSec
  MY_TP_LE10: 41,        // Column AO: MyTimePressureMoves<=10s
  MY_TP_LE5: 42,         // Column AP: MyTimePressureMoves<=5s
  MY_ACPL_OPEN: 43,      // Column AQ: MyACPL_Opening
  MY_ACPL_MID: 44,       // Column AR: MyACPL_Midgame
  MY_ACPL_END: 45,       // Column AS: MyACPL_Endgame
  OPP_ACPL_OPEN: 46,     // Column AT: OppACPL_Opening
  OPP_ACPL_MID: 47,      // Column AU: OppACPL_Midgame
  OPP_ACPL_END: 48,      // Column AV: OppACPL_Endgame
  MY_RESULT_NUM: 49,     // Column AW: MyResultNum (1/0.5/0)
  ACCURACY_DIFF: 50,     // Column AX: MyAccuracyMinusOpp
  RATING_DIFF: 51,       // Column AY: OppRatingMinusMyRating
  MAX_LEAD_CP_MY: 52,    // Column AZ: MaxLeadCpMy
  MIN_LEAD_CP_MY: 53,    // Column BA: MinLeadCpMy
  MAX_SWING_CP: 54,      // Column BB: MaxSingleMoveSwingCp
  FIRST_INACC_PLY_MY: 55,// Column BC: FirstInaccPlyMy
  FIRST_MIST_PLY_MY: 56, // Column BD: FirstMistPlyMy
  FIRST_BLUN_PLY_MY: 57, // Column BE: FirstBlunPlyMy
  FIRST_INACC_PLY_OPP: 58,// Column BF: FirstInaccPlyOpp
  FIRST_MIST_PLY_OPP: 59,// Column BG: FirstMistPlyOpp
  FIRST_BLUN_PLY_OPP: 60,// Column BH: FirstBlunPlyOpp
  COMEBACK: 61,          // Column BI: Comeback (1/0)
  CONVERSION: 62,        // Column BJ: Conversion (1/0)
  SAVE: 63,              // Column BK: Save (1/0)
  ADV_TIME_MY: 64,       // Column BL: AdvantageTimeSecMy (>= +50cp)
  ADV_TIME_OPP: 65,      // Column BM: AdvantageTimeSecOpp
  MAX_BLUNDER_STREAK_MY: 66, // Column BN
  MAX_BLUNDER_STREAK_OPP: 67, // Column BO
  OPENING_END_PLY: 68,   // Column BP
  ENDGAME_START_PLY: 69, // Column BQ
  EVAL_AT_MOVE20_MY: 70, // Column BR
  EVAL_AT_MOVE30_MY: 71  // Column BS
};

// Ensure the header row has expected titles for the opening columns
function ensureOpeningHeaders_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), COL.EVAL_AT_MOVE30_MY)).getValues()[0];

  // Expand columns if needed
  if (sheet.getMaxColumns() < COL.EVAL_AT_MOVE30_MY) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL.EVAL_AT_MOVE30_MY - sheet.getMaxColumns());
  }

  const expected = {
    [COL.OPEN_FAM - 1]: 'OpeningFamily',
    [COL.OPEN_VAR - 1]: 'OpeningVariation',
    [COL.OPEN_SUB1 - 1]: 'OpeningSub1',
    [COL.OPEN_SUB2 - 1]: 'OpeningSub2',
    [COL.OPEN_ECO - 1]: 'OpeningECO',
    [COL.WHITE_ACC - 1]: 'WhiteAccuracy',
    [COL.BLACK_ACC - 1]: 'BlackAccuracy',
    [COL.RESULT - 1]: 'Result',
    [COL.WHITE_ACPL - 1]: 'WhiteACPL',
    [COL.BLACK_ACPL - 1]: 'BlackACPL',
    [COL.WHITE_BEST_PCT - 1]: 'WhiteBestPct',
    [COL.BLACK_BEST_PCT - 1]: 'BlackBestPct',
    [COL.WHITE_INACC - 1]: 'WhiteInaccuracies',
    [COL.WHITE_MIST - 1]: 'WhiteMistakes',
    [COL.WHITE_BLUN - 1]: 'WhiteBlunders',
    [COL.BLACK_INACC - 1]: 'BlackInaccuracies',
    [COL.BLACK_MIST - 1]: 'BlackMistakes',
    [COL.BLACK_BLUN - 1]: 'BlackBlunders',
    [COL.WHITE_MISSED_WINS - 1]: 'WhiteMissedWins',
    [COL.BLACK_MISSED_WINS - 1]: 'BlackMissedWins',
    [COL.THEORY_LIKE_PLY - 1]: 'TheoryLikePly',
    [COL.MY_USERNAME - 1]: 'MyUsername',
    [COL.OPP_USERNAME - 1]: 'OppUsername',
    [COL.MY_COLOR - 1]: 'MyColor',
    [COL.MY_RATING - 1]: 'MyRating',
    [COL.OPP_RATING - 1]: 'OppRating',
    [COL.TIME_CONTROL - 1]: 'TimeControl',
    [COL.INITIAL_SEC - 1]: 'InitialSec',
    [COL.INCREMENT - 1]: 'Increment',
    [COL.SPEED_CLASS - 1]: 'SpeedClass',
    [COL.TERMINATION - 1]: 'Termination',
    [COL.MOVES_COUNT - 1]: 'MovesCount',
    [COL.PLIES - 1]: 'Plies',
    [COL.MY_AVG_MOVE_SEC - 1]: 'MyAvgMoveTimeSec',
    [COL.OPP_AVG_MOVE_SEC - 1]: 'OppAvgMoveTimeSec',
    [COL.MY_TP_LE10 - 1]: 'MyTPMoves<=10s',
    [COL.MY_TP_LE5 - 1]: 'MyTPMoves<=5s',
    [COL.MY_ACPL_OPEN - 1]: 'MyACPL_Open',
    [COL.MY_ACPL_MID - 1]: 'MyACPL_Mid',
    [COL.MY_ACPL_END - 1]: 'MyACPL_End',
    [COL.OPP_ACPL_OPEN - 1]: 'OppACPL_Open',
    [COL.OPP_ACPL_MID - 1]: 'OppACPL_Mid',
    [COL.OPP_ACPL_END - 1]: 'OppACPL_End',
    [COL.MY_RESULT_NUM - 1]: 'MyResultNum',
    [COL.ACCURACY_DIFF - 1]: 'MyAccuracyMinusOpp',
    [COL.RATING_DIFF - 1]: 'OppRatingMinusMyRating',
    [COL.MAX_LEAD_CP_MY - 1]: 'MaxLeadCpMy',
    [COL.MIN_LEAD_CP_MY - 1]: 'MinLeadCpMy',
    [COL.MAX_SWING_CP - 1]: 'MaxSingleMoveSwingCp',
    [COL.FIRST_INACC_PLY_MY - 1]: 'FirstInaccPlyMy',
    [COL.FIRST_MIST_PLY_MY - 1]: 'FirstMistPlyMy',
    [COL.FIRST_BLUN_PLY_MY - 1]: 'FirstBlunPlyMy',
    [COL.FIRST_INACC_PLY_OPP - 1]: 'FirstInaccPlyOpp',
    [COL.FIRST_MIST_PLY_OPP - 1]: 'FirstMistPlyOpp',
    [COL.FIRST_BLUN_PLY_OPP - 1]: 'FirstBlunPlyOpp',
    [COL.COMEBACK - 1]: 'Comeback',
    [COL.CONVERSION - 1]: 'Conversion',
    [COL.SAVE - 1]: 'Save',
    [COL.ADV_TIME_MY - 1]: 'AdvantageTimeSecMy',
    [COL.ADV_TIME_OPP - 1]: 'AdvantageTimeSecOpp',
    [COL.MAX_BLUNDER_STREAK_MY - 1]: 'MaxBlunderStreakMy',
    [COL.MAX_BLUNDER_STREAK_OPP - 1]: 'MaxBlunderStreakOpp',
    [COL.OPENING_END_PLY - 1]: 'OpeningEndPly',
    [COL.ENDGAME_START_PLY - 1]: 'EndgameStartPly',
    [COL.EVAL_AT_MOVE20_MY - 1]: 'EvalAtMove20My',
    [COL.EVAL_AT_MOVE30_MY - 1]: 'EvalAtMove30My'
  };

  let mutated = false;
  Object.keys(expected).forEach(function(indexStr) {
    const idx = Number(indexStr);
    if ((headers[idx] || '').toString().trim() !== expected[idx]) {
      headers[idx] = expected[idx];
      mutated = true;
    }
  });

  if (mutated) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

// Main entry: read PGN from column E and fill opening columns F-J, analysis K-M, metrics N-Z, and insights AA-AV
function updateOpeningsFromAnalyzedPgnSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

  ensureOpeningHeaders_();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // nothing to do

  const rowCount = lastRow - 1;

  // Read PGNs and existing opening values to avoid wiping non-empty rows unnecessarily
  const pgnValues = sheet.getRange(2, COL.ANALYZED_PGN, rowCount, 1).getValues();
  const existingOpenings = sheet.getRange(2, COL.OPEN_FAM, rowCount, 5).getValues();
  const existingAnalysis = sheet.getRange(2, COL.WHITE_ACC, rowCount, 3).getValues();
  const existingMetrics = sheet.getRange(2, COL.WHITE_ACPL, rowCount, (COL.THEORY_LIKE_PLY - COL.WHITE_ACPL + 1)).getValues();
  const existingInsights = sheet.getRange(2, COL.MY_USERNAME, rowCount, (COL.EVAL_AT_MOVE30_MY - COL.MY_USERNAME + 1)).getValues();

  const openingOutputs = new Array(rowCount);
  const analysisOutputs = new Array(rowCount);
  const metricsOutputs = new Array(rowCount);
  const insightsOutputs = new Array(rowCount);

  for (let i = 0; i < rowCount; i++) {
    const pgnText = (pgnValues[i][0] || '').toString();
    if (!pgnText) {
      // keep existing values if PGN is empty
      openingOutputs[i] = existingOpenings[i];
      analysisOutputs[i] = existingAnalysis[i];
      metricsOutputs[i] = existingMetrics[i];
      insightsOutputs[i] = existingInsights[i];
      continue;
    }

    const opening = parseOpeningFieldsFromPgn(pgnText);
    const analysis = parseAnalysisFieldsFromPgn(pgnText);
    const evals = parseEvalSequenceFromPgn_(pgnText);
    const metrics = computeEvalMetrics_(evals);
    const clocks = parseClocksSequenceFromPgn_(pgnText);
    const players = parsePlayersAndGameTagsFromPgn_(pgnText);
    const my = determineMyOrientation_(players.whiteName, players.blackName);
    const movesInfo = estimateMovesAndPlies_(evals, pgnText);
    const timeMetrics = computeTimeMetrics_(clocks, my.side, players.initialSec);
    const phaseAcpl = computeAcplByPhase_(evals, metrics ? metrics.theoryLikePly : 0);
    const extra = computeExtraGameInsights_(pgnText, evals, my, analysis, players, clocks);
    // If no opening parsed, keep existing
    if (!opening || !(opening.family || opening.variation || opening.sub1 || opening.sub2 || opening.eco)) {
      openingOutputs[i] = existingOpenings[i];
    } else {
      openingOutputs[i] = [
        opening.family || '',
        opening.variation || '',
        opening.sub1 || '',
        opening.sub2 || '',
        opening.eco || ''
      ];
    }

    if (!analysis || !(analysis.whiteAccuracy || analysis.blackAccuracy || analysis.result)) {
      analysisOutputs[i] = existingAnalysis[i];
    } else {
      analysisOutputs[i] = [
        analysis.whiteAccuracy || '',
        analysis.blackAccuracy || '',
        analysis.result || ''
      ];
    }

    if (!metrics) {
      metricsOutputs[i] = existingMetrics[i];
    } else {
      metricsOutputs[i] = [
        metrics.whiteAcpl, metrics.blackAcpl,
        metrics.whiteBestPct, metrics.blackBestPct,
        metrics.whiteInacc, metrics.whiteMist, metrics.whiteBlun,
        metrics.blackInacc, metrics.blackMist, metrics.blackBlun,
        metrics.whiteMissedWins, metrics.blackMissedWins,
        metrics.theoryLikePly
      ];
    }

    // Insights row
    insightsOutputs[i] = [
      my.username || '',
      my.oppUsername || '',
      my.side ? (my.side === 'w' ? 'White' : 'Black') : '',
      extra.myRating != null ? extra.myRating : '',
      extra.oppRating != null ? extra.oppRating : '',
      players.timeControl || '',
      players.initialSec != null ? players.initialSec : '',
      players.increment != null ? players.increment : '',
      players.speedClass || '',
      players.termination || '',
      movesInfo.moves != null ? movesInfo.moves : '',
      movesInfo.plies != null ? movesInfo.plies : '',
      timeMetrics.myAvgSec != null ? round1_(timeMetrics.myAvgSec) : '',
      timeMetrics.oppAvgSec != null ? round1_(timeMetrics.oppAvgSec) : '',
      timeMetrics.myTPLe10 != null ? timeMetrics.myTPLe10 : '',
      timeMetrics.myTPLe5 != null ? timeMetrics.myTPLe5 : '',
      phaseAcpl.myOpen != null ? phaseAcpl.myOpen : '',
      phaseAcpl.myMid != null ? phaseAcpl.myMid : '',
      phaseAcpl.myEnd != null ? phaseAcpl.myEnd : '',
      phaseAcpl.oppOpen != null ? phaseAcpl.oppOpen : '',
      phaseAcpl.oppMid != null ? phaseAcpl.oppMid : '',
      phaseAcpl.oppEnd != null ? phaseAcpl.oppEnd : '',
      extra.myResultNum != null ? extra.myResultNum : '',
      extra.accDiff != null ? extra.accDiff : '',
      extra.ratingDiff != null ? extra.ratingDiff : '',
      extra.maxLeadMy != null ? extra.maxLeadMy : '',
      extra.minLeadMy != null ? extra.minLeadMy : '',
      extra.maxSwing != null ? extra.maxSwing : '',
      extra.firstInaccMy != null ? extra.firstInaccMy : '',
      extra.firstMistMy != null ? extra.firstMistMy : '',
      extra.firstBlunMy != null ? extra.firstBlunMy : '',
      extra.firstInaccOpp != null ? extra.firstInaccOpp : '',
      extra.firstMistOpp != null ? extra.firstMistOpp : '',
      extra.firstBlunOpp != null ? extra.firstBlunOpp : '',
      extra.comeback != null ? extra.comeback : '',
      extra.conversion != null ? extra.conversion : '',
      extra.save != null ? extra.save : '',
      extra.advTimeMy != null ? extra.advTimeMy : '',
      extra.advTimeOpp != null ? extra.advTimeOpp : '',
      extra.maxBlunderStreakMy != null ? extra.maxBlunderStreakMy : '',
      extra.maxBlunderStreakOpp != null ? extra.maxBlunderStreakOpp : '',
      extra.openingEndPly != null ? extra.openingEndPly : '',
      extra.endgameStartPly != null ? extra.endgameStartPly : '',
      extra.evalAtMove20My != null ? extra.evalAtMove20My : '',
      extra.evalAtMove30My != null ? extra.evalAtMove30My : ''
    ];
  }

  sheet.getRange(2, COL.OPEN_FAM, rowCount, 5).setValues(openingOutputs);
  sheet.getRange(2, COL.WHITE_ACC, rowCount, 3).setValues(analysisOutputs);
  sheet.getRange(2, COL.WHITE_ACPL, rowCount, (COL.THEORY_LIKE_PLY - COL.WHITE_ACPL + 1)).setValues(metricsOutputs);
  sheet.getRange(2, COL.MY_USERNAME, rowCount, (COL.OPP_ACPL_END - COL.MY_USERNAME + 1)).setValues(insightsOutputs);
}

// Extracts Opening/ECO tags and splits into family/variation/sub-variations
function parseOpeningFieldsFromPgn(pgnText) {
  if (!pgnText) return { family: '', variation: '', sub1: '', sub2: '', eco: '' };

  const openingName = extractPgnTag_(pgnText, 'Opening');
  const eco = extractPgnTag_(pgnText, 'ECO');

  const parts = splitOpeningName_(openingName);
  return {
    family: parts.family || '',
    variation: parts.variation || '',
    sub1: (parts.subs && parts.subs[0]) || '',
    sub2: (parts.subs && parts.subs[1]) || '',
    eco: eco || ''
  };
}

// Extract WhiteAccuracy/BlackAccuracy/Result from PGN tags, if present
function parseAnalysisFieldsFromPgn(pgnText) {
  if (!pgnText) return { whiteAccuracy: '', blackAccuracy: '', result: '' };
  const whiteAcc = extractPgnTag_(pgnText, 'WhiteAccuracy');
  const blackAcc = extractPgnTag_(pgnText, 'BlackAccuracy');
  const result = extractPgnTag_(pgnText, 'Result');
  return {
    whiteAccuracy: whiteAcc || '',
    blackAccuracy: blackAcc || '',
    result: result || ''
  };
}

// Parse a sequence of engine evals from the PGN (in move order). Returns an array of
// { side: 'w'|'b', cp: number, isMate: boolean, matePly: number|null }
function parseEvalSequenceFromPgn_(pgnText) {
  const re = /\[%eval\s+([^\]]+)\]/g;
  const evals = [];
  let m;
  let plyIndex = 0; // 0-based, even=white, odd=black
  while ((m = re.exec(pgnText)) !== null) {
    const token = (m[1] || '').trim();
    let cp = null;
    let isMate = false;
    let matePly = null;
    if (token.indexOf('#') !== -1) {
      // Mate form like #5 or #-3
      isMate = true;
      const sign = token.indexOf('#-') !== -1 ? -1 : 1;
      const numMatch = token.match(/#-?(\d+)/);
      matePly = numMatch ? parseInt(numMatch[1], 10) : null;
      // Represent mate as very large cp with sign from white's perspective
      cp = sign * 10000;
    } else {
      // Numeric pawns value, convert to cp
      const val = parseFloat(token);
      if (isFinite(val)) cp = Math.round(val * 100);
    }
    if (cp !== null) {
      evals.push({ side: (plyIndex % 2 === 0 ? 'w' : 'b'), cp: cp, isMate: isMate, matePly: matePly });
      plyIndex++;
    }
  }
  return evals;
}

// Compute ACPL-like metrics and mistake counts from eval sequence
function computeEvalMetrics_(evals) {
  if (!evals || !evals.length) return null;

  const clampCpForMate = function(cp) {
    // Clamp to avoid exploding averages
    if (cp > 2000) return 2000;
    if (cp < -2000) return -2000;
    return cp;
  };

  let whiteLossSum = 0, blackLossSum = 0;
  let whiteMoves = 0, blackMoves = 0;
  let whiteBest = 0, blackBest = 0;
  let whiteInacc = 0, whiteMist = 0, whiteBlun = 0;
  let blackInacc = 0, blackMist = 0, blackBlun = 0;
  let whiteMissedWins = 0, blackMissedWins = 0;

  let theoryLikePly = 0;
  for (let i = 0; i < evals.length; i++) {
    const after = clampCpForMate(evals[i].cp);
    const before = (i > 0) ? clampCpForMate(evals[i - 1].cp) : null;
    const mover = evals[i].side; // 'w' or 'b'

    // Theory-like heuristic: early plies with small swing and near equal
    if (i === theoryLikePly) {
      const delta = (before === null) ? 0 : (after - before);
      const nearEqual = Math.abs(after) <= 20; // <= 0.20 pawns
      const smallSwing = Math.abs(delta) <= 30; // <= 0.30 pawns
      if (nearEqual && smallSwing) theoryLikePly++; // continue streak
    }

    if (before === null) continue; // cannot compute loss on very first move

    const beforePersp = (mover === 'w') ? before : -before;
    const afterPersp = (mover === 'w') ? after : -after;
    const loss = Math.max(0, beforePersp - afterPersp);

    // Classify thresholds (centipawns)
    const isBest = loss <= 10; // within 0.10 pawns
    const isInacc = loss > 50 && loss <= 100; // >0.50 and <=1.00 pawns
    const isMist = loss > 100 && loss <= 300; // >1.00 and <=3.00 pawns
    const isBlun = loss > 300; // >3.00 pawns

    if (mover === 'w') {
      whiteMoves++;
      whiteLossSum += loss;
      if (isBest) whiteBest++;
      if (isInacc) whiteInacc++;
      if (isMist) whiteMist++;
      if (isBlun) whiteBlun++;

      // Missed win: had big advantage, then dropped below a safer threshold
      if (beforePersp >= 300 && afterPersp <= 100) whiteMissedWins++;
    } else {
      blackMoves++;
      blackLossSum += loss;
      if (isBest) blackBest++;
      if (isInacc) blackInacc++;
      if (isMist) blackMist++;
      if (isBlun) blackBlun++;

      if (beforePersp >= 300 && afterPersp <= 100) blackMissedWins++;
    }
  }

  const whiteAcpl = whiteMoves ? Math.round(whiteLossSum / whiteMoves) : '';
  const blackAcpl = blackMoves ? Math.round(blackLossSum / blackMoves) : '';
  const whiteBestPct = whiteMoves ? Math.round((whiteBest / whiteMoves) * 1000) / 10 : '';
  const blackBestPct = blackMoves ? Math.round((blackBest / blackMoves) * 1000) / 10 : '';

  return {
    whiteAcpl: whiteAcpl,
    blackAcpl: blackAcpl,
    whiteBestPct: whiteBestPct,
    blackBestPct: blackBestPct,
    whiteInacc: whiteInacc,
    whiteMist: whiteMist,
    whiteBlun: whiteBlun,
    blackInacc: blackInacc,
    blackMist: blackMist,
    blackBlun: blackBlun,
    whiteMissedWins: whiteMissedWins,
    blackMissedWins: blackMissedWins,
    theoryLikePly: theoryLikePly
  };
}

// ======= Insights helpers =======

function setMyUsernameInteractive() {
  const username = Browser.inputBox('Enter your Lichess username (stored in Script Properties):');
  if (username && username !== 'cancel') {
    PropertiesService.getScriptProperties().setProperty(SCRIPT_PROP_MY_USERNAME_KEY, username.trim());
    SpreadsheetApp.getActive().toast('Username saved.');
  }
}

function determineMyOrientation_(whiteName, blackName) {
  const myUsername = (PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_MY_USERNAME_KEY) || '').trim();
  const w = (whiteName || '').toLowerCase();
  const b = (blackName || '').toLowerCase();
  const me = (myUsername || '').toLowerCase();
  let side = '';
  let username = '';
  let oppUsername = '';
  let rating = null;
  let oppRating = null;
  if (me && w === me) {
    side = 'w'; username = whiteName; oppUsername = blackName;
  } else if (me && b === me) {
    side = 'b'; username = blackName; oppUsername = whiteName;
  }
  return { side: side, username: username, oppUsername: oppUsername, rating: rating, oppRating: oppRating };
}

function parsePlayersAndGameTagsFromPgn_(pgnText) {
  const whiteName = extractPgnTag_(pgnText, 'White');
  const blackName = extractPgnTag_(pgnText, 'Black');
  const whiteEloStr = extractPgnTag_(pgnText, 'WhiteElo');
  const blackEloStr = extractPgnTag_(pgnText, 'BlackElo');
  const timeControl = extractPgnTag_(pgnText, 'TimeControl');
  const termination = extractPgnTag_(pgnText, 'Termination');

  const tc = parseTimeControl_(timeControl);
  const speedClass = classifySpeed_(tc.initialSec, tc.increment);

  return {
    whiteName: whiteName,
    blackName: blackName,
    whiteElo: parseIntSafe_(whiteEloStr),
    blackElo: parseIntSafe_(blackEloStr),
    timeControl: timeControl,
    initialSec: tc.initialSec,
    increment: tc.increment,
    speedClass: speedClass,
    termination: termination
  };
}

function parseIntSafe_(s) {
  const n = parseInt((s || '').toString(), 10);
  return isFinite(n) ? n : null;
}

function parseTimeControl_(tc) {
  // Formats: "300+0", "600+5", "600" (no increment), "-" (no clock)
  if (!tc || tc === '-') return { initialSec: null, increment: null };
  const parts = tc.split('+');
  const initialSec = parseIntSafe_(parts[0]);
  const increment = parts.length > 1 ? parseIntSafe_(parts[1]) : 0;
  return { initialSec: initialSec, increment: increment };
}

function classifySpeed_(initialSec, increment) {
  if (initialSec == null) return '';
  const inc = increment || 0;
  const base = initialSec + inc * 40;
  if (base < 180) return 'bullet';
  if (base < 480) return 'blitz';
  if (base < 1500) return 'rapid';
  return 'classical';
}

function parseClocksSequenceFromPgn_(pgnText) {
  const re = /\[%clk\s+([0-9:]+)\]/g;
  const arr = [];
  let m;
  while ((m = re.exec(pgnText)) !== null) {
    arr.push(parseClockStrToSec_(m[1]));
  }
  return arr; // clocks per ply in move order (White then Black repeating)
}

function parseClockStrToSec_(s) {
  const parts = (s || '').split(':').map(function(p){return parseInt(p,10)||0;});
  if (parts.length === 3) return parts[0]*3600 + parts[1]*60 + parts[2];
  if (parts.length === 2) return parts[0]*60 + parts[1];
  return parseInt(s,10)||0;
}

function estimateMovesAndPlies_(evals, pgnText) {
  const plies = (evals && evals.length) ? evals.length : estimatePliesFromPgn_(pgnText);
  const moves = Math.ceil((plies || 0) / 2);
  return { plies: plies, moves: moves };
}

function estimatePliesFromPgn_(pgnText) {
  // Fallback: count move numbers like "1.", "2." etc and multiply by 2 approximately
  const nums = (pgnText.match(/\n\d+\./g) || []).length;
  return nums * 2;
}

function computeTimeMetrics_(clocks, mySide, initialSec) {
  const result = { myAvgSec: null, oppAvgSec: null, myTPLe10: null, myTPLe5: null };
  if (!clocks || !clocks.length) return result;
  // Build per-side sequences
  const whiteClocks = [];
  const blackClocks = [];
  for (let i = 0; i < clocks.length; i++) {
    if (i % 2 === 0) whiteClocks.push(clocks[i]);
    else blackClocks.push(clocks[i]);
  }
  const comp = function(seq, init) {
    let times = [];
    for (let i = 0; i < seq.length; i++) {
      if (i === 0) {
        if (init != null && seq[i] != null) times.push(Math.max(0, init - seq[i]));
      } else if (seq[i-1] != null && seq[i] != null) {
        times.push(Math.max(0, seq[i-1] - seq[i]));
      }
    }
    if (!times.length) return { avg: null, median: null, tp10: 0, tp5: 0 };
    const avg = times.reduce(function(a,b){return a+b;},0) / times.length;
    const sorted = times.slice().sort(function(a,b){return a-b;});
    const mid = Math.floor(sorted.length/2);
    const median = sorted.length % 2 ? sorted[mid] : (sorted[mid-1]+sorted[mid])/2;
    return {
      avg: avg,
      median: median,
      tp10: seq.filter(function(s){return s != null && s <= 10;}).length,
      tp5: seq.filter(function(s){return s != null && s <= 5;}).length
    };
  };
  const white = comp(whiteClocks, initialSec);
  const black = comp(blackClocks, initialSec);
  if (mySide === 'w') {
    result.myAvgSec = white.avg; result.oppAvgSec = black.avg;
    result.myTPLe10 = white.tp10; result.myTPLe5 = white.tp5;
  } else if (mySide === 'b') {
    result.myAvgSec = black.avg; result.oppAvgSec = white.avg;
    result.myTPLe10 = black.tp10; result.myTPLe5 = black.tp5;
  }
  return result;
}

function computeAcplByPhase_(evals, theoryLikePly) {
  const res = { myOpen: null, myMid: null, myEnd: null, oppOpen: null, oppMid: null, oppEnd: null };
  const mySide = determineMyOrientation_(null, null).side; // may be '' if not set; we compute white/black then map
  if (!evals || !evals.length) return res;
  const openEnd = Math.max(0, theoryLikePly || 0);
  const endStart = Math.max(openEnd, evals.length - 16);

  function acplForRange(start, end, side) {
    let lossSum = 0, moves = 0;
    for (let i = Math.max(1, start); i < Math.min(end, evals.length); i++) {
      const mover = evals[i].side;
      if (side === mover) {
        const before = clampCp_(evals[i-1].cp);
        const after = clampCp_(evals[i].cp);
        const beforePersp = (mover === 'w') ? before : -before;
        const afterPersp = (mover === 'w') ? after : -after;
        const loss = Math.max(0, beforePersp - afterPersp);
        lossSum += loss; moves++;
      }
    }
    return moves ? Math.round(lossSum / moves) : null;
  }

  function clampCp_(cp) { if (cp > 2000) return 2000; if (cp < -2000) return -2000; return cp; }

  const wOpen = acplForRange(0, openEnd, 'w');
  const bOpen = acplForRange(0, openEnd, 'b');
  const wMid = acplForRange(openEnd, endStart, 'w');
  const bMid = acplForRange(openEnd, endStart, 'b');
  const wEnd = acplForRange(endStart, evals.length, 'w');
  const bEnd = acplForRange(endStart, evals.length, 'b');

  if (mySide === 'w') {
    res.myOpen = wOpen; res.myMid = wMid; res.myEnd = wEnd;
    res.oppOpen = bOpen; res.oppMid = bMid; res.oppEnd = bEnd;
  } else if (mySide === 'b') {
    res.myOpen = bOpen; res.myMid = bMid; res.myEnd = bEnd;
    res.oppOpen = wOpen; res.oppMid = wMid; res.oppEnd = wEnd;
  } else {
    // Unknown orientation: leave nulls
  }
  return res;
}

function round1_(x) { return Math.round(x * 10) / 10; }

function computeExtraGameInsights_(pgnText, evals, my, analysis, players, clocks) {
  // Map ratings to my/opp
  let myRating = null, oppRating = null;
  if (my.side === 'w') { myRating = players.whiteElo; oppRating = players.blackElo; }
  else if (my.side === 'b') { myRating = players.blackElo; oppRating = players.whiteElo; }

  // Result numeric
  const resultTag = extractPgnTag_(pgnText, 'Result');
  let myResultNum = null;
  if (my.side) {
    if (resultTag === '1-0') myResultNum = (my.side === 'w') ? 1 : 0;
    else if (resultTag === '0-1') myResultNum = (my.side === 'b') ? 1 : 0;
    else if (resultTag === '1/2-1/2') myResultNum = 0.5;
  }

  // Accuracy diff
  const wAcc = parseFloat(analysis.whiteAccuracy || '') || null;
  const bAcc = parseFloat(analysis.blackAccuracy || '') || null;
  let accDiff = null;
  if (wAcc != null && bAcc != null) accDiff = (my.side === 'w') ? (wAcc - bAcc) : (bAcc - wAcc);

  // Rating diff (opp - mine)
  let ratingDiff = null;
  if (myRating != null && oppRating != null) ratingDiff = (oppRating - myRating);

  // Lead/swing, first mistakes, blunder streaks
  let maxLeadMy = null, minLeadMy = null, maxSwing = 0;
  let firstInaccMy = null, firstMistMy = null, firstBlunMy = null;
  let firstInaccOpp = null, firstMistOpp = null, firstBlunOpp = null;
  let streakMy = 0, streakOpp = 0, maxStreakMy = 0, maxStreakOpp = 0;
  const thInacc = 50, thMist = 100, thBlun = 300;

  for (let i = 0; i < evals.length; i++) {
    const cp = evals[i].cp;
    const side = evals[i].side;
    const myLead = (my.side === 'w') ? cp : -cp;
    maxLeadMy = (maxLeadMy == null) ? myLead : Math.max(maxLeadMy, myLead);
    minLeadMy = (minLeadMy == null) ? myLead : Math.min(minLeadMy, myLead);
    if (i > 0) maxSwing = Math.max(maxSwing, Math.abs(evals[i].cp - evals[i-1].cp));

    if (i > 0) {
      const before = evals[i-1].cp;
      const after = evals[i].cp;
      const mover = side;
      const beforePersp = (mover === 'w') ? before : -before;
      const afterPersp = (mover === 'w') ? after : -after;
      const loss = Math.max(0, beforePersp - afterPersp);
      const isInacc = loss > thInacc && loss <= thMist;
      const isMist = loss > thMist && loss <= thBlun;
      const isBlun = loss > thBlun;

      if (my.side && mover === my.side) {
        if (isInacc && firstInaccMy == null) firstInaccMy = i + 1;
        if (isMist && firstMistMy == null) firstMistMy = i + 1;
        if (isBlun && firstBlunMy == null) firstBlunMy = i + 1;
        if (isBlun) { streakMy++; maxStreakMy = Math.max(maxStreakMy, streakMy); } else { streakMy = 0; }
      } else {
        if (isInacc && firstInaccOpp == null) firstInaccOpp = i + 1;
        if (isMist && firstMistOpp == null) firstMistOpp = i + 1;
        if (isBlun && firstBlunOpp == null) firstBlunOpp = i + 1;
        if (isBlun) { streakOpp++; maxStreakOpp = Math.max(maxStreakOpp, streakOpp); } else { streakOpp = 0; }
      }
    }
  }

  // Comeback/Conversion/Save
  let comeback = 0, conversion = 0, save = 0;
  if (my.side) {
    const hadLead = (maxLeadMy != null && maxLeadMy >= 200);
    const hadDeficit = (minLeadMy != null && minLeadMy <= -200);
    if (hadLead && myResultNum === 1) conversion = 1;
    if (hadDeficit && myResultNum === 0.5) save = 1;
    if (hadDeficit && myResultNum === 1) comeback = 1;
  }

  // Advantage time (>= +50 cp) by my side using clocks
  let advTimeMy = null, advTimeOpp = null;
  if (clocks && clocks.length && evals && evals.length) {
    const whiteClocks = [], blackClocks = [];
    for (let i = 0; i < clocks.length; i++) { if (i % 2 === 0) whiteClocks.push(clocks[i]); else blackClocks.push(clocks[i]); }
    function sumIntervals(seq) {
      let sum = 0;
      for (let i = 1; i < seq.length; i++) if (seq[i-1] != null && seq[i] != null) sum += Math.max(0, seq[i-1] - seq[i]);
      return sum;
    }
    // Count time while lead for the side on move >= +50
    function timeWhileLeading(side) {
      let total = 0;
      if (side === 'w') {
        for (let i = 1; i < whiteClocks.length; i++) {
          const cp = evals[i*2 - 1] ? evals[i*2 - 1].cp : null; // before white's i-th move, after black's move
          if (cp != null && cp >= 50) total += Math.max(0, whiteClocks[i-1] - whiteClocks[i]);
        }
      } else {
        for (let i = 1; i < blackClocks.length; i++) {
          const cp = evals[i*2] ? evals[i*2].cp : null; // before black's i-th move, after white's move
          if (cp != null && -cp >= 50) total += Math.max(0, blackClocks[i-1] - blackClocks[i]);
        }
      }
      return total;
    }
    advTimeMy = my.side ? timeWhileLeading(my.side) : null;
    advTimeOpp = my.side ? timeWhileLeading(my.side === 'w' ? 'b' : 'w') : null;
  }

  // Phase boundaries & checkpoints
  const openingEndPly = (evals && evals.length) ? Math.max(0, analysis && analysis.theoryLikePly ? analysis.theoryLikePly : 0) : null;
  const endgameStartPly = (evals && evals.length) ? Math.max(openingEndPly || 0, evals.length - 16) : null;
  function evalAtMove(moveNum) {
    const ply = (moveNum * 2) - 1; // after black move
    if (evals && evals[ply]) {
      const cp = evals[ply].cp;
      return (my.side === 'w') ? cp : -cp;
    }
    return null;
  }

  const evalAt20 = evalAtMove(20);
  const evalAt30 = evalAtMove(30);

  return {
    myRating: myRating,
    oppRating: oppRating,
    myResultNum: myResultNum,
    accDiff: accDiff,
    ratingDiff: ratingDiff,
    maxLeadMy: maxLeadMy,
    minLeadMy: minLeadMy,
    maxSwing: maxSwing,
    firstInaccMy: firstInaccMy,
    firstMistMy: firstMistMy,
    firstBlunMy: firstBlunMy,
    firstInaccOpp: firstInaccOpp,
    firstMistOpp: firstMistOpp,
    firstBlunOpp: firstBlunOpp,
    comeback: comeback,
    conversion: conversion,
    save: save,
    advTimeMy: advTimeMy,
    advTimeOpp: advTimeOpp,
    maxBlunderStreakMy: maxStreakMy,
    maxBlunderStreakOpp: maxStreakOpp,
    openingEndPly: openingEndPly,
    endgameStartPly: endgameStartPly,
    evalAtMove20My: evalAt20,
    evalAtMove30My: evalAt30
  };
}

// Reads a PGN tag like [TagName "Value"] and returns the Value
function extractPgnTag_(pgnText, tagName) {
  // Match e.g. [Opening "Sicilian Defense: Najdorf Variation, English Attack"]
  const re = new RegExp('\\[' + tagName + '\\s+"([^"\\\\]*(?:\\\\.[^"\\\\]*)*)"\\]');
  const match = pgnText.match(re);
  return match ? match[1] : '';
}

// Split an opening name into family / variation / sub-variations
// Example: "Sicilian Defense: Najdorf Variation, English Attack"
//  -> family: "Sicilian Defense"
//  -> variation: "Najdorf Variation"
//  -> subs: ["English Attack"]
function splitOpeningName_(name) {
  if (!name) return { family: '', variation: '', subs: [] };

  // Separate at the first colon, if present
  const colonIndex = name.indexOf(':');
  if (colonIndex === -1) {
    return { family: name.trim(), variation: '', subs: [] };
  }

  const family = name.slice(0, colonIndex).trim();
  const rest = name.slice(colonIndex + 1).trim();
  if (!rest) return { family: family, variation: '', subs: [] };

  // Split rest by commas into variation and deeper sub-variations
  const parts = rest.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
  const variation = parts.shift() || '';
  return { family: family, variation: variation, subs: parts };
}

// Simple test runner in Logs
function test_parseOpeningFieldsFromPgn() {
  const sample = '[Event "Casual game"]\n' +
    '[Site "https://lichess.org/AbCdEf12"]\n' +
    '[Date "2025.09.06"]\n' +
    '[Opening "Sicilian Defense: Najdorf Variation, English Attack"]\n' +
    '[ECO "B90"]\n\n' +
    '1. e4 c5 2. Nf3 d6 3. d4 cxd4 4. Nxd4 Nf6 5. Nc3 a6 6. Be3 e6 7. f3';

  const res = parseOpeningFieldsFromPgn(sample);
  Logger.log(JSON.stringify(res, null, 2));
  // Expected:
  // {
  //   "family": "Sicilian Defense",
  //   "variation": "Najdorf Variation",
  //   "sub1": "English Attack",
  //   "sub2": "",
  //   "eco": "B90"
  // }
}


// Apps Script utility to parse opening and analysis fields from a PGN export and
// populate OpeningFamily, OpeningVariation, OpeningSub1, OpeningSub2, OpeningECO
// plus WhiteAccuracy, BlackAccuracy, Result, and extended metrics computed from [%eval]s
// in a Google Sheet. This relies solely on PGN tags: [Opening "..."] and [ECO "..."]

const SHEET_NAME = 'Games';

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
  THEORY_LIKE_PLY: 26    // Column Z: TheoryLikePly (heuristic)
};

// Ensure the header row has expected titles for the opening columns
function ensureOpeningHeaders_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), COL.THEORY_LIKE_PLY)).getValues()[0];

  // Expand columns if needed
  if (sheet.getMaxColumns() < COL.THEORY_LIKE_PLY) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL.THEORY_LIKE_PLY - sheet.getMaxColumns());
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
    [COL.THEORY_LIKE_PLY - 1]: 'TheoryLikePly'
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

// Main entry: read PGN from column E and fill opening columns F-J, analysis K-M, and metrics N-Z
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

  const openingOutputs = new Array(rowCount);
  const analysisOutputs = new Array(rowCount);
  const metricsOutputs = new Array(rowCount);

  for (let i = 0; i < rowCount; i++) {
    const pgnText = (pgnValues[i][0] || '').toString();
    if (!pgnText) {
      // keep existing values if PGN is empty
      openingOutputs[i] = existingOpenings[i];
      analysisOutputs[i] = existingAnalysis[i];
      metricsOutputs[i] = existingMetrics[i];
      continue;
    }

    const opening = parseOpeningFieldsFromPgn(pgnText);
    const analysis = parseAnalysisFieldsFromPgn(pgnText);
    const evals = parseEvalSequenceFromPgn_(pgnText);
    const metrics = computeEvalMetrics_(evals);
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
  }

  sheet.getRange(2, COL.OPEN_FAM, rowCount, 5).setValues(openingOutputs);
  sheet.getRange(2, COL.WHITE_ACC, rowCount, 3).setValues(analysisOutputs);
  sheet.getRange(2, COL.WHITE_ACPL, rowCount, (COL.THEORY_LIKE_PLY - COL.WHITE_ACPL + 1)).setValues(metricsOutputs);
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


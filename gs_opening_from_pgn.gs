// Apps Script utility to parse opening fields from a PGN export and
// populate OpeningFamily, OpeningVariation, OpeningSub1, OpeningSub2, OpeningECO
// in a Google Sheet. This relies solely on PGN tags: [Opening "..."] and [ECO "..."]

const SHEET_NAME = 'Games';

// Column indices (1-based)
const COL = {
  ANALYZED_PGN: 5,  // Column E: PGN that includes [Opening] and [ECO] tags
  OPEN_FAM: 6,      // Column F: OpeningFamily
  OPEN_VAR: 7,      // Column G: OpeningVariation
  OPEN_SUB1: 8,     // Column H: OpeningSub1
  OPEN_SUB2: 9,     // Column I: OpeningSub2
  OPEN_ECO: 10      // Column J: OpeningECO
};

// Ensure the header row has expected titles for the opening columns
function ensureOpeningHeaders_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), COL.OPEN_ECO)).getValues()[0];

  // Expand columns if needed
  if (sheet.getMaxColumns() < COL.OPEN_ECO) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL.OPEN_ECO - sheet.getMaxColumns());
  }

  const expected = {
    [COL.OPEN_FAM - 1]: 'OpeningFamily',
    [COL.OPEN_VAR - 1]: 'OpeningVariation',
    [COL.OPEN_SUB1 - 1]: 'OpeningSub1',
    [COL.OPEN_SUB2 - 1]: 'OpeningSub2',
    [COL.OPEN_ECO - 1]: 'OpeningECO'
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

// Main entry: read PGN from column E and fill opening columns F-J
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

  const outputs = new Array(rowCount);

  for (let i = 0; i < rowCount; i++) {
    const pgnText = (pgnValues[i][0] || '').toString();
    if (!pgnText) {
      // keep existing values if PGN is empty
      outputs[i] = existingOpenings[i];
      continue;
    }

    const opening = parseOpeningFieldsFromPgn(pgnText);
    // If no opening parsed, keep existing
    if (!opening || !(opening.family || opening.variation || opening.sub1 || opening.sub2 || opening.eco)) {
      outputs[i] = existingOpenings[i];
    } else {
      outputs[i] = [
        opening.family || '',
        opening.variation || '',
        opening.sub1 || '',
        opening.sub2 || '',
        opening.eco || ''
      ];
    }
  }

  sheet.getRange(2, COL.OPEN_FAM, rowCount, 5).setValues(outputs);
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


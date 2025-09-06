## PGN-based Opening Extraction for Google Sheets

This project provides Apps Script utilities to import/analyze games on Lichess and parse opening information directly from the PGN export. It populates opening columns in a `Games` sheet using only the `[Opening]` and `[ECO]` tags.

### Files
- `code.gs`: Main workflow (batch import/analyze via Lichess)
- `gs_opening_from_pgn.gs`: PGN tag parser and sheet updater

### Sheet Setup
Create a sheet named `Games` with at least these columns:

1. `PGN`
2. `LichessGameId`
3. `LichessUrl`
4. `AnalysisStatus`
5. `AnalyzedPGN` (PGN with tags from Lichess export)
6. `OpeningFamily`
7. `OpeningVariation`
8. `OpeningSub1`
9. `OpeningSub2`
10. `OpeningECO`
11. `WhiteAccuracy`
12. `BlackAccuracy`
13. `Result`

`gs_opening_from_pgn.gs` will ensure the Opening columns (6–10) exist.

### Usage: Fill Opening & Analysis Columns From PGN
1. Export or fetch the analyzed PGN from Lichess and place it in column E (`AnalyzedPGN`).
2. In Apps Script, run `updateOpeningsFromAnalyzedPgnSheet()`.
3. The script will parse opening fields:
   - `OpeningFamily`: text before the first `:`
   - `OpeningVariation`: first item after `:`
   - `OpeningSub1`, `OpeningSub2`: subsequent comma-separated items
   - `OpeningECO`: from the `[ECO "..."]` tag
   And analysis fields:
   - `WhiteAccuracy`: from `[WhiteAccuracy "..."]`
   - `BlackAccuracy`: from `[BlackAccuracy "..."]`
   - `Result`: from `[Result "..."]` (e.g., `1-0`, `0-1`, `1/2-1/2`)

Example mapping:

Input PGN tags:
```
[Opening "Sicilian Defense: Najdorf Variation, English Attack"]
[ECO "B90"]
```

Output columns:
```
OpeningFamily:   Sicilian Defense
OpeningVariation: Najdorf Variation
OpeningSub1:     English Attack
OpeningSub2:     (empty)
OpeningECO:      B90
WhiteAccuracy:   92.4
BlackAccuracy:   78.1
Result:          1-0
```

### Notes
- Lichess determines `[Opening]` and `[ECO]` from the position (transposition-aware), not just the move order.
- If the PGN lacks those tags, existing opening columns are left unchanged.
- For very large PGNs, consider storing them outside the main sheet to avoid cell size limits.

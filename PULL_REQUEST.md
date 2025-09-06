## Title
Add PGN-based opening extraction for Games sheet (OpeningFamily/Variation/Sub1/Sub2/ECO)

## Summary
This PR adds a standalone Apps Script file that parses the `[Opening]` and `[ECO]` PGN tags from the analyzed PGN and fills the following columns in the `Games` sheet:
- OpeningFamily (F)
- OpeningVariation (G)
- OpeningSub1 (H)
- OpeningSub2 (I)
- OpeningECO (J)

It is transposition-aware because the tags are produced by Lichess from the position (PGN → FEN → opening), not the move order.

## Files Added
- `gs_opening_from_pgn.gs` – Functions to parse PGN tags and write opening columns

## User-facing Changes
- New function `updateOpeningsFromAnalyzedPgnSheet()` to backfill opening columns from `AnalyzedPGN` in column E.

## How to Use
1) Ensure the `Games` sheet has:
   - Column E: `AnalyzedPGN` (PGN text that includes `[Opening]` and `[ECO]` tags)
   - Columns F–J will be created/updated as: OpeningFamily, OpeningVariation, OpeningSub1, OpeningSub2, OpeningECO
2) Run `updateOpeningsFromAnalyzedPgnSheet()`.
3) The script preserves existing opening values if no PGN is present or tags are missing.

## Implementation Notes
- Parsing rules:
  - Split `[Opening]` value at the first `:` → Family vs. rest
  - Split the rest by `,` → Variation (first), then Sub-variations (remaining)
  - ECO is read from `[ECO]`
- No external dependencies. Uses Apps Script `SpreadsheetApp` only.

## Testing
- Added a simple `test_parseOpeningFieldsFromPgn()` function logging expected output for:
  `"Sicilian Defense: Najdorf Variation, English Attack"` with ECO `B90`.

## Limitations
- Relies on PGN containing `[Opening]` and `[ECO]` tags. If missing, columns remain unchanged.
- Very long PGNs can exceed cell limits; consider storing large PGNs in a separate sheet or Drive file.

## Checklist
- [x] New file added: `gs_opening_from_pgn.gs`
- [x] Backfill function implemented and documented
- [x] No breaking changes to existing flows

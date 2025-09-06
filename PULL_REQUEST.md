## Title
Add PGN-based opening and analysis extraction for Games sheet

## Summary
This PR adds a standalone Apps Script file that parses the `[Opening]`, `[ECO]`, and analysis-related PGN tags from the analyzed PGN and fills the following columns in the `Games` sheet:
- OpeningFamily (F)
- OpeningVariation (G)
- OpeningSub1 (H)
- OpeningSub2 (I)
- OpeningECO (J)
- WhiteAccuracy (K)
- BlackAccuracy (L)
- Result (M)
- WhiteACPL (N)
- BlackACPL (O)
- WhiteBestPct (P)
- BlackBestPct (Q)
- WhiteInaccuracies (R)
- WhiteMistakes (S)
- WhiteBlunders (T)
- BlackInaccuracies (U)
- BlackMistakes (V)
- BlackBlunders (W)
- WhiteMissedWins (X)
- BlackMissedWins (Y)
- TheoryLikePly (Z)

It is transposition-aware because the tags are produced by Lichess from the position (PGN → FEN → opening), not the move order.

## Files Added
- `gs_opening_from_pgn.gs` – Functions to parse PGN tags and write opening, analysis, and metrics columns

## User-facing Changes
- New function `updateOpeningsFromAnalyzedPgnSheet()` to backfill opening, analysis, and metrics columns from `AnalyzedPGN` in column E.

## How to Use
1) Ensure the `Games` sheet has:
   - Column E: `AnalyzedPGN` (PGN text from Lichess export that includes tags and eval comments)
   - Columns F–Z will be created/updated as above
2) Run `updateOpeningsFromAnalyzedPgnSheet()`.
3) The script preserves existing values if no PGN is present or tags are missing.

## Implementation Notes
- Opening parsing
  - Split `[Opening]` at first `:` → Family vs. rest; split rest by `,` → Variation then Sub-variations
  - ECO is read from `[ECO]`
- Analysis parsing
  - Accuracy values are read from `[WhiteAccuracy]`, `[BlackAccuracy]`
  - Result is read from `[Result]`
- Metrics parsing (from `[%eval ...]` comments)
  - ACPL: average non-negative loss (centipawns) relative to mover’s perspective
  - BestPct: share of moves with loss ≤ 10 cp
  - Inaccuracy/Mistake/Blunder thresholds: >50/100/300 cp
  - MissedWins: had ≥ +300 cp then moved to ≤ +100 cp
  - TheoryLikePly: early consecutive plies with |eval| ≤ 20 cp and |delta| ≤ 30 cp
- Mate evals are clamped for stability; no external dependencies.

## Testing
- Included `test_parseOpeningFieldsFromPgn()` which logs expected output for a known opening (`B90`).

## Limitations
- Relies on PGN containing the relevant tags/evals. If missing, columns remain unchanged.
- Heuristics for metrics are approximations (Chrome extension parity not guaranteed).
- Very long PGNs can exceed cell limits; consider storing large PGNs in a separate sheet or Drive file.

## Checklist
- [x] New file added: `gs_opening_from_pgn.gs`
- [x] Backfill function implemented and documented
- [x] No breaking changes to existing flows

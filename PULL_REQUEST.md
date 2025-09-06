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
- MyUsername (AA)
- OppUsername (AB)
- MyColor (AC)
- MyRating (AD)
- OppRating (AE)
- TimeControl (AF)
- InitialSec (AG)
- Increment (AH)
- SpeedClass (AI)
- Termination (AJ)
- MovesCount (AK)
- Plies (AL)
- MyAvgMoveTimeSec (AM)
- OppAvgMoveTimeSec (AN)
- MyTPMoves<=10s (AO)
- MyTPMoves<=5s (AP)
- MyACPL_Open (AQ)
- MyACPL_Mid (AR)
- MyACPL_End (AS)
- OppACPL_Open (AT)
- OppACPL_Mid (AU)
- OppACPL_End (AV)

It is transposition-aware because the tags are produced by Lichess from the position (PGN → FEN → opening), not the move order.

## Files Added
- `gs_opening_from_pgn.gs` – Functions to parse PGN tags and write opening, analysis, metrics, and insights columns

## User-facing Changes
- New function `updateOpeningsFromAnalyzedPgnSheet()` to backfill opening, analysis, metrics, and insights columns from `AnalyzedPGN` in column E.

## How to Use
1) Ensure the `Games` sheet has:
   - Column E: `AnalyzedPGN` (PGN text from Lichess export that includes tags and eval/clock comments)
   - Columns F–AV will be created/updated as above
2) Run `setMyUsernameInteractive()` once to set your username (for side-aware metrics).
3) Run `updateOpeningsFromAnalyzedPgnSheet()`.
4) The script preserves existing values if no PGN is present or tags are missing.

## Implementation Notes
- Opening parsing
  - Split `[Opening]` at first `:` → Family vs. rest; split rest by `,` → Variation then Sub-variations
  - ECO is read from `[ECO]`
- Analysis parsing
  - Accuracy values are read from `[WhiteAccuracy]`, `[BlackAccuracy]`
  - Result is read from `[Result]`
- Metrics parsing (from `[%eval ...]` comments)
  - ACPL: average non-negative loss (centipawns) relative to mover’s perspective; mates clamped
  - BestPct: share of moves with loss ≤ 10 cp; Inaccuracy/Mistake/Blunder thresholds: >50/100/300 cp
  - MissedWins: had ≥ +300 cp then moved to ≤ +100 cp
  - TheoryLikePly: early consecutive plies with |eval| ≤ 20 cp and |delta| ≤ 30 cp
- Insights parsing
  - Players, ratings, time control, termination: from PGN tags (`White`, `Black`, `WhiteElo`, `BlackElo`, `TimeControl`, `Termination`)
  - Speed class: from initial + 40×increment heuristic
  - Avg move times and time pressure: derived from `[%clk]` deltas per side
  - Phase ACPL: opening/midgame/endgame split using `TheoryLikePly` and last 16 plies for endgame

## Testing
- Included `test_parseOpeningFieldsFromPgn()` which logs expected output for a known opening (`B90`).

## Limitations
- Relies on PGN containing the relevant tags/evals/clocks. If missing, columns remain unchanged.
- Heuristics approximate Chess.com Insights and the referenced extension; parity not guaranteed.
- Very long PGNs can exceed cell limits; consider storing large PGNs in a separate sheet or Drive file.

## Checklist
- [x] New file added: `gs_opening_from_pgn.gs`
- [x] Backfill function implemented and documented
- [x] No breaking changes to existing flows

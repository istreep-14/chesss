## Legacy (v1) scripts and docs

The current project is Version 2 (see `version2.gs` and `config.gs`). The files below are from an earlier iteration and are kept for reference only. Version 2 does not depend on or invoke them unless you explicitly include them in your Apps Script project.

What's legacy:
- `code.gs`: Older scripts/utilities from the previous iteration; not used by V2.
- `gs_opening_from_pgn.gs`: Parses Lichess-analyzed PGN to fill opening, accuracy, and metrics columns in a sheet.
- `openingInfo.gs`: Builds an "Opening Info" tab and explorer summaries for a selected game.

Notes about Version 2:
- V2 lives in `version2.gs` and `config.gs`.
- V2 uses a `Headers` sheet to choose fields and a `Config` sheet to control fetching and recalculation.
- V2 fetches Chess.com monthly archives and writes raw data without spreadsheet formulas.

Where to find legacy docs:
- The previous README content for the PGN-based workflow is preserved below the V2 section in `README.md`.


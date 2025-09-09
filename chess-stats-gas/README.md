# Chess Stats (Google Apps Script)

Minimal Apps Script to fetch new games from Chess.com into a Google Sheet with configurable headers (visibility, widths, colors) and simple deduping by URL.

## Setup

1) Create a new Google Sheet (or open an existing one).
2) In the Sheet, go to Extensions → Apps Script.
3) Create these files and paste the code from this repo:
   - `config.gs`
   - `sheets.gs`
   - `chesscom.gs`
   - `main.gs`
4) Save the project. The first run will prompt you to authorize permissions.

Notes:
- If using clasp or an exported project, include `appsscript.json` (provided) to set scopes and V8.
- No external libs are needed; uses `UrlFetchApp` and `SpreadsheetApp` built-ins.

## Usage

- In the Sheet, from the toolbar choose Chess Stats → "Setup Games Sheet" to create/refresh the header row and formatting.
- Fetch a single month: Chess Stats → "Fetch Month…" (enter: `username, YYYY-MM`)
- Fetch a range of months: Chess Stats → "Fetch Range…" (enter: `username, YYYY-MM..YYYY-MM`)
- The script dedupes by `url`, so re-fetching a month will only add new games.

## Configure Headers

Open `config.gs` and edit the `HEADERS` array:
- `visible`: show/hide columns
- `width`: column width in pixels
- `color`: header background color
- `from`: source field taken from fetched data
- `formula`: optional formula string (placeholders supported: `${row}`, `${col(key)}`)

Example: To hide PGN, set `visible: false`. To add a calculated column, add an entry with a `formula` and no `from`.

## Customization Tips

- Add other sources by creating a new `fetch...` function returning row objects matching `from` keys.
- Extend dedupe: change `getExistingUrlSet_()` to use another key if needed.
- Date/timezone: `end_time` is stored in UTC ISO; adjust formatting in the Sheet if desired.

## Troubleshooting

- Authorization prompts: run any menu item once to grant permissions.
- Empty adds: If nothing is added, either there were no games or they already exist (same `url`).
- Rate limits: Chess.com APIs are public but may throttle; try smaller ranges.
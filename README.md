# Excel Color Extractor

Extracts exactly what is present in an Excel workbook — cell values and cell
fill colors — per sheet, and reports honest coverage/accuracy statistics on
the extraction itself.

**This tool never assumes what a color means.** Different hotels use
different, inconsistent color-coding schemes with no shared legend, so the
extractor only ever reports raw facts ("cell A3 has value 'Occupied' and color
#FF0000") plus an auto-discovered inventory of the distinct colors found in
each sheet. It never invents a label, never fills in a blank or unresolved
color with a guessed default, and never merges data across sheets. Human
labels can optionally be attached afterward, per sheet, per workbook — never
inferred automatically, and never shared between workbooks.

## Project layout

```
backend/    FastAPI service, extraction engine, MongoDB access
frontend/   React (Vite) UI
```

## Prerequisites

- Python 3.10+
- Node.js 18+
- A local MongoDB server

### Installing & running MongoDB locally

If you don't already have MongoDB installed:

- **Windows**: install "MongoDB Community Server" from mongodb.com, or via
  `winget install MongoDB.Server`. It installs as a Windows service
  (`MongoDB`) listening on `mongodb://localhost:27017` by default.
- **macOS**: `brew tap mongodb/brew && brew install mongodb-community && brew services start mongodb-community`
- **Linux**: follow your distro's MongoDB Community Edition install instructions,
  then `sudo systemctl start mongod`.

Verify it's running:

```bash
mongosh --eval "db.runCommand({ ping: 1 })"
```

No database or collections need to be created ahead of time — the backend
creates the `excel_color_extractor` database and its `workbooks` /
`extractions` collections on first write.

## Backend setup

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS/Linux
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

The API is now at `http://localhost:8000` (interactive docs at
`http://localhost:8000/docs`). Uploaded files are stored under
`backend/storage/`.

Environment variables (optional):

- `MONGO_URI` (default `mongodb://localhost:27017`)
- `MONGO_DB_NAME` (default `excel_color_extractor`)
- `STORAGE_DIR` (default `backend/storage`)

## Frontend setup

```bash
cd frontend
npm install
npm run dev
```

The UI is now at `http://localhost:5173` and proxies `/api/*` requests to the
backend at `http://localhost:8000`.

## API summary

| Method | Path | Purpose |
|---|---|---|
| POST | `/workbooks/upload` | Upload an `.xlsx`/`.xls` file, get immediate metadata |
| POST | `/workbooks/{id}/extract` | Run extraction across all sheets |
| GET | `/workbooks/{id}` | Metadata + per-sheet coverage/color-inventory summaries |
| GET | `/workbooks/{id}/sheets/{sheet_name}` | Full raw extracted data for one sheet |
| POST | `/workbooks/{id}/sheets/{sheet_name}/labels` | Attach a human label to a color, for this sheet only |
| GET | `/workbooks/{id}/sheets/{sheet_name}/report?format=csv\|json\|xlsx` | Downloadable per-sheet report |
| GET | `/workbooks/{id}/report?format=csv\|json\|xlsx` | Downloadable whole-workbook report |

## How color resolution works (and where it deliberately gives up)

For every used cell, the extractor:

1. Checks the cell's **direct fill**. If it has an explicit RGB (or a legacy
   indexed palette color that isn't a system/auto slot), that's reported as
   resolved.
2. If there's no direct fill, it checks whether the cell is covered by a
   **conditional formatting** rule. Simple `cellIs` rules with literal
   operands (`greaterThan 100`, `equal "Closed"`, etc.) are evaluated against
   the cell's actual value. If exactly one rule matches, its fill is used.
3. Anything that can't be resolved with certainty is reported as such, with a
   reason — never guessed:
   - **Theme colors** (`type="theme"`) — flagged as
     "exact shade not resolvable without further theme-XML lookup".
   - **Auto colors** — flagged as "no explicit RGB specified".
   - **Indexed system slots** (64/65) — flagged as "system/auto, not an
     explicit color".
   - **Formula-based / scale-based conditional formatting** (`expression`,
     `colorScale`, `dataBar`, `iconSet`, or a `cellIs` rule whose operand
     isn't a literal) — flagged as "requires evaluation that was not
     performed".
   - **Multiple conditional formatting rules matching the same cell** —
     flagged as ambiguous rather than picking one arbitrarily.
4. **Merged cells**: the value is reported once, against the top-left anchor
   cell, with a `merged_with` list of the other cell references in the merge.
5. Sheets with no formatting at all produce a valid report with 0% color
   coverage — not an error.

Every count in the coverage summary (`total_cells`, `colored_cells`,
`blank_cells`, `unresolved_cells`, `ambiguous_cells`) is a direct tally from
this per-cell process, and every report bundles that summary alongside the
raw data it was computed from.

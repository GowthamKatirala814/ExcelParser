# Excel Color Extractor

Extracts exactly what is present in an Excel workbook — cell values and cell
fill colors — per sheet, and reports honest coverage/accuracy statistics on
the extraction itself. Optionally, it can turn that raw extraction into a
clean, business-friendly **structured JSON** using an LLM (Gemini) — but only
when you explicitly ask for it.

**This tool never assumes what a color means.** Different hotels use
different, inconsistent color-coding schemes with no shared legend, so the
raw extractor only ever reports facts ("cell A3 has value 'Occupied' and
color #FF0000") plus an auto-discovered inventory of the distinct colors
found in each sheet. It never invents a label, never fills in a blank or
unresolved color with a guessed default, and never merges data across sheets.
Human labels can optionally be attached afterward, per sheet, per color.

## Two stages

```
Excel (.xlsx/.xls)
      ↓  openpyxl  (instant, no AI, no token cost)
Raw extraction  — values, fill colors, fonts, borders, comments, formulas,
      ↓           merged cells, structural blocks + coverage stats
Structured JSON — only on demand (Gemini): legend/entity/calendar detection,
                  date ranges computed deterministically from the raw cells
```

- **Plain extraction is always instant and free.** It runs on upload/extract
  with zero AI involvement.
- **AI structuring is strictly opt-in.** It only runs when you click
  "Generate & Download" (or "Regenerate") for a sheet or the whole workbook.
  Results are cached per sheet, so re-downloading the same format never
  re-spends tokens — only an explicit "Regenerate" does.

## Project layout

```
backend/    FastAPI service, extraction engine, LLM structuring, MongoDB access
frontend/   React (Vite) UI
```

## Prerequisites

- Python 3.10+
- Node.js 18+
- A local MongoDB server
- (Optional, only for AI structuring) one or more Google Gemini API keys

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
creates the `excel_color_extractor` database and its `workbooks`,
`extractions`, and `structured` collections on first write.

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

### Configuration (environment variables)

All configuration lives in `backend/.env` (git-ignored). Copy the template
and fill in real values:

```bash
cp .env.example .env    # then edit backend/.env
```

Plain extraction needs no configuration. AI structuring needs at least one
Gemini key.

| Variable | Default | Purpose |
|---|---|---|
| `MONGO_URI` | `mongodb://localhost:27017` | MongoDB connection string |
| `MONGO_DB_NAME` | `excel_color_extractor` | Database name |
| `STORAGE_DIR` | `backend/storage` | Where uploaded files are stored |
| `GEMINI_API_KEY` | *(none)* | A single Gemini key |
| `GEMINI_API_KEYS` | *(none)* | Comma-separated list of keys |
| `GEMINI_API_KEY_1`, `_2`, … | *(none)* | Numbered keys (any count) |
| `GEMINI_MODEL` | `gemini-flash-latest` | Model used for structuring |
| `GEMINI_KEY_COOLDOWN_SECONDS` | `60` | Cooldown before retrying a rate-limited key |
| `LLM_PROVIDER` | `gemini` | Which LLM backs the structuring stage |
| `LLM_MAX_INPUT_BYTES` | `15000000` | Max payload size sent to the LLM |

**Multiple API keys / automatic rotation:** any combination of
`GEMINI_API_KEY`, `GEMINI_API_KEYS`, and `GEMINI_API_KEY_N` is merged into one
rotation pool. If a key hits its quota, rate limit, or a model-access error,
the request automatically retries with the next available key; the failed key
is put on a short cooldown. Adding another key is a one-line change in
`backend/.env` — no code change. A single key is enough to run the app.

> Note: `gemini-2.5-flash` returns HTTP 404 ("no longer available to new
> users") on some newer API keys/projects even though it appears in the model
> catalog. `gemini-flash-latest` (the default) avoids that restriction.

## Frontend setup

```bash
cd frontend
npm install
npm run dev
```

The UI is now at `http://localhost:5173` and proxies `/api/*` requests to the
backend at `http://localhost:8000`.

## Using it

1. **Upload** an `.xlsx`/`.xls` file, then click **Upload**.
2. Review the **workbook** summary (file name, size, per-sheet row/column counts).
3. Click **Extract all sheets** — instant, no AI. You get per-sheet coverage
   stats, a per-color inventory (with inline, per-color labeling), and a
   per-cell table.
4. **Download** via the single "Format" dropdown + Download button (per sheet
   and for the whole workbook):
   - **CSV / JSON / Raw JSON** — available immediately, no AI.
   - **Structured JSON (AI)** — if not generated yet the button reads
     "Generate & Download" (runs the AI, then downloads); if already generated
     it downloads the cached copy instantly. A separate **Regenerate** button
     forces a fresh AI run.

## API summary

| Method | Path | Purpose |
|---|---|---|
| POST | `/workbooks/upload` | Upload an `.xlsx`/`.xls` file, get immediate metadata |
| POST | `/workbooks/{id}/extract` | Run plain extraction across all sheets (no AI) |
| GET | `/workbooks/{id}` | Metadata + per-sheet coverage/color-inventory summaries |
| GET | `/workbooks/{id}/sheets/{sheet}` | Full raw extracted data for one sheet |
| POST | `/workbooks/{id}/sheets/{sheet}/labels` | Attach a human label to a color, for this sheet only |
| GET | `/workbooks/{id}/sheets/{sheet}/report?format=csv\|json` | Downloadable per-sheet report |
| GET | `/workbooks/{id}/report?format=csv\|json` | Downloadable whole-workbook report (zip) |
| GET | `/workbooks/{id}/raw.json` | Raw extraction JSON, whole workbook |
| GET | `/workbooks/{id}/sheets/{sheet}/raw.json` | Raw extraction JSON, one sheet |
| POST | `/workbooks/{id}/structured/generate?force=` | Generate structured JSON for the workbook (on demand) |
| POST | `/workbooks/{id}/sheets/{sheet}/structured/generate?force=` | Generate structured JSON for one sheet (on demand) |
| GET | `/workbooks/{id}/structured/progress` | Per-sheet structuring status |
| GET | `/workbooks/{id}/structured` | Download combined structured JSON, whole workbook |
| GET | `/workbooks/{id}/sheets/{sheet}/structured` | Download structured JSON, one sheet |

`force=false` (default) reuses any already-generated (cached) sheet;
`force=true` regenerates from scratch.

## How color resolution works (and where it deliberately gives up)

For every used cell, the extractor:

1. Checks the cell's **direct fill**. An explicit RGB, a resolvable **theme
   color** (resolved against the workbook's own `theme` XML with the standard
   tint/shade formula), or a legacy indexed palette color that isn't a
   system/auto slot is reported as resolved.
2. If there's no direct fill, it checks whether the cell is covered by a
   **conditional formatting** rule. Simple `cellIs` rules with literal
   operands (`greaterThan 100`, `equal "Closed"`, etc.) and `containsText`
   rules are evaluated against the cell's actual value, in the workbook's own
   rule-priority order. The first rule that definitely matches wins.
3. Anything that can't be resolved with certainty is reported as such, with a
   reason — never guessed (auto colors, indexed system slots, unresolvable
   theme references, formula/scale-based conditional formatting, or a rule
   whose operand isn't a literal).
4. **Merged cells**: the value is reported once against the top-left anchor
   cell, with a `merged_with` list of the other cell references.
5. Sheets with no formatting at all produce a valid report with 0% color
   coverage — not an error.

Every count in the coverage summary (`total_cells`, `colored_cells`,
`blank_cells`, `unresolved_cells`, `ambiguous_cells`) is a direct tally from
this per-cell process.

## How AI structuring stays trustworthy

When structured JSON is requested, the LLM is used only to identify **meaning**
that genuinely needs judgment — a legend/key section, an entity master table,
and which rows are entities in a calendar/availability grid. It is **not**
asked to transcribe or count grid cells. Every actual date range and status
value in the output is computed deterministically by code, directly from the
already-trusted raw cells, so the LLM cannot introduce an off-by-a-day or a
dropped/merged range. After generation, the result is validated against the
raw extraction (date bounds, value completeness) and any inconsistencies are
surfaced as warnings rather than silently accepted. Each sheet is processed
independently, so one sheet failing (quota, timeout, malformed response)
never affects the others, and the raw extraction is always downloadable
regardless.

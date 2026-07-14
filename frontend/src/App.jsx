import { useState } from "react";
import UploadPanel from "./components/UploadPanel.jsx";
import SheetTab from "./components/SheetTab.jsx";
import DownloadControl from "./components/DownloadControl.jsx";
import {
  extractWorkbook,
  generateSheetStructuring,
  generateStructuring,
  getStructuringProgress,
  rawJsonUrl,
  structuredJsonUrl,
  workbookReportUrl,
} from "./api.js";

const POLL_INTERVAL_MS = 3000;
const MAX_POLLS = 200;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const isBusy = (s) => s === "pending" || s === "processing";

function anySheetBusy(progress) {
  return Object.values(progress?.sheets || {}).some((s) => isBusy(s.status));
}
function workbookAllSettled(progress) {
  const sheets = Object.values(progress?.sheets || {});
  return sheets.length > 0 && sheets.every((s) => s.status === "completed" || s.status === "failed");
}
function workbookAnyGenerated(progress) {
  return Object.values(progress?.sheets || {}).some(
    (s) => s.status === "completed" || s.status === "failed"
  );
}

export default function App() {
  const [workbook, setWorkbook] = useState(null);
  const [extractSummary, setExtractSummary] = useState(null);
  const [extracting, setExtracting] = useState(false);
  const [extractError, setExtractError] = useState(null);
  const [activeSheet, setActiveSheet] = useState(null);
  const [progress, setProgress] = useState(null);

  const handleUploaded = (wb) => {
    setWorkbook(wb);
    setExtractSummary(null);
    setExtractError(null);
    setActiveSheet(null);
    setProgress(null);
  };

  // One-shot fetch of current structuring status (no looping). Used after
  // extract so the UI shows "not started" for every sheet without kicking
  // off any AI work.
  const refreshProgress = async (workbookId) => {
    try {
      setProgress(await getStructuringProgress(workbookId));
    } catch {
      /* leave prior progress in place */
    }
  };

  const handleExtract = async () => {
    if (!workbook) return;
    setExtracting(true);
    setExtractError(null);
    try {
      const result = await extractWorkbook(workbook.id);
      setExtractSummary(result);
      setActiveSheet(Object.keys(result.sheets)[0] || null);
      await refreshProgress(workbook.id); // shows all "not started"; no AI call
    } catch (e) {
      setExtractError(e.message);
    } finally {
      setExtracting(false);
    }
  };

  // Triggers on-demand generation, then polls until the requested scope has
  // settled, updating `progress` each round so the UI reflects it live.
  // Returns { ready } once done. This is the ONLY path that spends tokens.
  const generateAndWait = async ({ scope, sheetName, force }) => {
    if (scope === "sheet") {
      await generateSheetStructuring(workbook.id, sheetName, force);
    } else {
      await generateStructuring(workbook.id, force);
    }

    for (let i = 0; i < MAX_POLLS; i++) {
      const data = await getStructuringProgress(workbook.id);
      setProgress(data);

      if (scope === "sheet") {
        const status = data.sheets?.[sheetName]?.status;
        if (!isBusy(status)) return { ready: status === "completed" };
      } else if (workbookAllSettled(data)) {
        return { ready: workbookAnyGenerated(data) };
      }
      await sleep(POLL_INTERVAL_MS);
    }
    return { ready: false };
  };

  // Structured status for the whole-workbook download control.
  const workbookStructured = {
    ready: workbookAllSettled(progress),
    busy: anySheetBusy(progress),
    failed: false,
    error: null,
  };

  const workbookUrls = workbook
    ? {
        csv: workbookReportUrl(workbook.id, "csv"),
        json: workbookReportUrl(workbook.id, "json"),
        raw: rawJsonUrl(workbook.id),
        structured: structuredJsonUrl(workbook.id),
      }
    : {};

  return (
    <div className="app">
      <header>
        <h1>Excel Color Extractor</h1>
        <p className="tagline">
          Extracts the values and fill colors present in a workbook, per sheet, with honest
          coverage stats. Colors are never assumed to mean anything — you label them yourself.
        </p>
      </header>

      <UploadPanel onUploaded={handleUploaded} />

      {workbook && (
        <section>
          <h2>2. Workbook</h2>
          <ul className="summary-list">
            <li>File: {workbook.filename}</li>
            <li>Size: {(workbook.size / 1024).toFixed(1)} KB</li>
            <li>Sheets: {workbook.sheet_names.length}</li>
          </ul>
          <table>
            <thead>
              <tr>
                <th>Sheet</th>
                <th>Rows</th>
                <th>Columns</th>
              </tr>
            </thead>
            <tbody>
              {workbook.sheet_names.map((name) => (
                <tr key={name}>
                  <td>{name}</td>
                  <td>{workbook.sheet_dimensions[name]?.rows}</td>
                  <td>{workbook.sheet_dimensions[name]?.columns}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>
      )}

      {workbook && (
        <section>
          <h2>3. Extract</h2>
          <button onClick={handleExtract} disabled={extracting}>
            {extracting ? "Processing…" : "Extract all sheets"}
          </button>
          {extractError && <p className="error">Error: {extractError}</p>}
          {extractSummary && (
            <p className="status-line">
              Done. Total processing time:{" "}
              {extractSummary.total_processing_time_seconds.toFixed(3)}s.
            </p>
          )}
        </section>
      )}

      {extractSummary && (
        <section>
          <h2>4. Results</h2>
          <div className="sheet-nav">
            {Object.keys(extractSummary.sheets).map((name) => (
              <button
                key={name}
                className={name === activeSheet ? "secondary active" : "secondary"}
                onClick={() => setActiveSheet(name)}
              >
                {name}
              </button>
            ))}
          </div>

          {activeSheet && (
            <SheetTab
              key={activeSheet}
              workbookId={workbook.id}
              sheetName={activeSheet}
              summary={extractSummary.sheets[activeSheet]}
              structuring={progress?.sheets?.[activeSheet]}
              onGenerate={(force) =>
                generateAndWait({ scope: "sheet", sheetName: activeSheet, force })
              }
            />
          )}
        </section>
      )}

      {extractSummary && (
        <section>
          <h2>5. Whole-workbook download</h2>
          <DownloadControl
            label="Download whole workbook"
            urls={workbookUrls}
            structured={workbookStructured}
            onGenerate={() => generateAndWait({ scope: "workbook", force: false })}
            onRegenerate={() => generateAndWait({ scope: "workbook", force: true })}
          />
          {progress && (workbookAnyGenerated(progress) || anySheetBusy(progress)) && (
            <p className="status-line">{progress.overall?.message}</p>
          )}
        </section>
      )}
    </div>
  );
}

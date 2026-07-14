import { useEffect, useState } from "react";
import {
  addLabel,
  getSheet,
  sheetRawJsonUrl,
  sheetReportUrl,
  sheetStructuredJsonUrl,
} from "../api.js";
import CoverageCard from "./CoverageCard.jsx";
import ColorInventoryTable from "./ColorInventoryTable.jsx";
import DataTable from "./DataTable.jsx";
import DownloadControl from "./DownloadControl.jsx";

export default function SheetTab({ workbookId, sheetName, summary, structuring, onGenerate }) {
  const [sheetData, setSheetData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  const load = async () => {
    setLoading(true);
    setError(null);
    try {
      setSheetData(await getSheet(workbookId, sheetName));
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [workbookId, sheetName]);

  const handleLabel = async (colorKey, label) => {
    try {
      await addLabel(workbookId, sheetName, colorKey, label);
      await load();
    } catch (e) {
      alert(`Could not save label: ${e.message}`);
    }
  };

  if (loading) return <p>Loading sheet data…</p>;
  if (error) return <p className="error">Error: {error}</p>;
  if (!sheetData) return null;

  const status = structuring?.status || "not_started";
  const structured = {
    ready: status === "completed",
    busy: status === "pending" || status === "processing",
    failed: status === "failed",
    error: structuring?.error,
  };

  const urls = {
    csv: sheetReportUrl(workbookId, sheetName, "csv"),
    json: sheetReportUrl(workbookId, sheetName, "json"),
    raw: sheetRawJsonUrl(workbookId, sheetName),
    structured: sheetStructuredJsonUrl(workbookId, sheetName),
  };

  return (
    <div>
      <h3>{sheetName}</h3>
      <CoverageCard summary={summary} />

      <h3>Colors and labels</h3>
      <p className="note">
        Labels are saved per color — every cell of that color shares the label.
      </p>
      <ColorInventoryTable inventory={sheetData.color_inventory} onLabel={handleLabel} />

      <h3>Cells</h3>
      <DataTable rows={sheetData.rows} />

      <DownloadControl
        label="Download this sheet"
        urls={urls}
        structured={structured}
        onGenerate={() => onGenerate(false)}
        onRegenerate={() => onGenerate(true)}
      />

      {status === "completed" && structuring?.validation_warnings?.length > 0 && (
        <div className="note">
          <p>Validation warnings (structured JSON is still downloadable):</p>
          <ul className="summary-list">
            {structuring.validation_warnings.map((w, i) => (
              <li key={i}>{w}</li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

import { useState } from "react";

// One Format dropdown + one Download button per scope (this sheet / whole
// workbook). CSV, JSON and Raw JSON download immediately. "Structured JSON
// (AI)" only calls the AI when the user asks: if it isn't generated yet the
// button reads "Generate & Download" and triggers generation first, then
// downloads; if it's already generated it downloads the cached result with
// no AI call. A separate "Regenerate" button forces a fresh AI run.

function triggerDownload(url) {
  const a = document.createElement("a");
  a.href = url;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
}

const FORMATS = [
  { value: "csv", label: "CSV" },
  { value: "json", label: "JSON" },
  { value: "raw", label: "Raw JSON" },
  { value: "structured", label: "Structured JSON (AI)" },
];

export default function DownloadControl({ label, urls, structured, onGenerate, onRegenerate }) {
  const [format, setFormat] = useState("csv");
  const [working, setWorking] = useState(false);

  const isStructured = format === "structured";
  const busy = working || (isStructured && structured.busy);

  const buttonLabel = busy
    ? "Generating…"
    : isStructured && !structured.ready
    ? "Generate & Download"
    : "Download";

  const handleDownload = async () => {
    if (!isStructured) {
      triggerDownload(urls[format]);
      return;
    }
    if (structured.ready) {
      triggerDownload(urls.structured);
      return;
    }
    setWorking(true);
    try {
      const { ready } = await onGenerate(); // force=false; reuses cached sheets
      if (ready) triggerDownload(urls.structured);
      else alert("AI generation did not complete. See the status message.");
    } catch (e) {
      alert(`Generation failed: ${e.message}`);
    } finally {
      setWorking(false);
    }
  };

  const handleRegenerate = async () => {
    setWorking(true);
    try {
      const { ready } = await onRegenerate(); // force=true
      if (ready) triggerDownload(urls.structured);
    } catch (e) {
      alert(`Regeneration failed: ${e.message}`);
    } finally {
      setWorking(false);
    }
  };

  return (
    <div className="downloads">
      <p className="downloads-label">{label}</p>
      <div className="download-group">
        <label>
          Format:{" "}
          <select value={format} onChange={(e) => setFormat(e.target.value)} disabled={busy}>
            {FORMATS.map((f) => (
              <option key={f.value} value={f.value}>
                {f.label}
              </option>
            ))}
          </select>
        </label>
        <button onClick={handleDownload} disabled={busy}>
          {buttonLabel}
        </button>
        {isStructured && structured.ready && !busy && (
          <button className="secondary" onClick={handleRegenerate}>
            Regenerate
          </button>
        )}
      </div>
      {isStructured && (
        <p className="note">Uses AI — only generated when requested.</p>
      )}
      {isStructured && structured.failed && structured.error && (
        <p className="error">{structured.error}</p>
      )}
    </div>
  );
}

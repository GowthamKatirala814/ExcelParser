import { useState } from "react";
import { uploadWorkbook } from "../api.js";

export default function UploadPanel({ onUploaded }) {
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState(null);

  const handleUpload = async () => {
    if (!file) return;
    setUploading(true);
    setError(null);
    try {
      const workbook = await uploadWorkbook(file);
      onUploaded(workbook);
    } catch (e) {
      setError(e.message);
    } finally {
      setUploading(false);
    }
  };

  return (
    <section>
      <h2>1. Upload a workbook</h2>
      <div className="field-row">
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setFile(e.target.files[0] || null)}
          disabled={uploading}
        />
        <button onClick={handleUpload} disabled={!file || uploading}>
          {uploading ? "Uploading…" : "Upload"}
        </button>
      </div>
      {error && <p className="error">Error: {error}</p>}
    </section>
  );
}

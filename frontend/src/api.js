const BASE = "/api";

async function handle(response) {
  if (!response.ok) {
    let detail = response.statusText;
    try {
      const body = await response.json();
      detail = body.detail || detail;
    } catch (_) {
      // ignore
    }
    throw new Error(detail);
  }
  return response;
}

export async function uploadWorkbook(file) {
  const formData = new FormData();
  formData.append("file", file);
  const res = await handle(await fetch(`${BASE}/workbooks/upload`, { method: "POST", body: formData }));
  return res.json();
}

export async function extractWorkbook(workbookId) {
  const res = await handle(await fetch(`${BASE}/workbooks/${workbookId}/extract`, { method: "POST" }));
  return res.json();
}

export async function getWorkbook(workbookId) {
  const res = await handle(await fetch(`${BASE}/workbooks/${workbookId}`));
  return res.json();
}

export async function getSheet(workbookId, sheetName) {
  const res = await handle(
    await fetch(`${BASE}/workbooks/${workbookId}/sheets/${encodeURIComponent(sheetName)}`)
  );
  return res.json();
}

export async function addLabel(workbookId, sheetName, colorKey, label) {
  const res = await handle(
    await fetch(`${BASE}/workbooks/${workbookId}/sheets/${encodeURIComponent(sheetName)}/labels`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ color_key: colorKey, label }),
    })
  );
  return res.json();
}

export function sheetReportUrl(workbookId, sheetName, format) {
  return `${BASE}/workbooks/${workbookId}/sheets/${encodeURIComponent(sheetName)}/report?format=${format}`;
}

export function workbookReportUrl(workbookId, format) {
  return `${BASE}/workbooks/${workbookId}/report?format=${format}`;
}

export function rawJsonUrl(workbookId) {
  return `${BASE}/workbooks/${workbookId}/raw.json`;
}

export function structuredJsonUrl(workbookId) {
  return `${BASE}/workbooks/${workbookId}/structured`;
}

export function sheetRawJsonUrl(workbookId, sheetName) {
  return `${BASE}/workbooks/${workbookId}/sheets/${encodeURIComponent(sheetName)}/raw.json`;
}

export function sheetStructuredJsonUrl(workbookId, sheetName) {
  return `${BASE}/workbooks/${workbookId}/sheets/${encodeURIComponent(sheetName)}/structured`;
}

export async function getStructuringProgress(workbookId) {
  const res = await handle(await fetch(`${BASE}/workbooks/${workbookId}/structured/progress`));
  return res.json();
}

export async function generateStructuring(workbookId, force = false) {
  const res = await handle(
    await fetch(`${BASE}/workbooks/${workbookId}/structured/generate?force=${force}`, {
      method: "POST",
    })
  );
  return res.json();
}

export async function generateSheetStructuring(workbookId, sheetName, force = false) {
  const res = await handle(
    await fetch(
      `${BASE}/workbooks/${workbookId}/sheets/${encodeURIComponent(
        sheetName
      )}/structured/generate?force=${force}`,
      { method: "POST" }
    )
  );
  return res.json();
}

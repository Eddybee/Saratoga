const API = "http://localhost:8000";

export async function uploadPdf(file) {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch(`${API}/upload-pdf`, { method: "POST", body: form });
  if (!res.ok) throw new Error((await res.json()).detail || "Upload failed");
  return res.json();
}

export async function getPageImageUrl(pdfId, page) {
  return `${API}/pdf/${pdfId}/page/${page}`;
}

export async function extractAoi(payload, options = {}) {
  const res = await fetch(`${API}/extract-aoi`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    signal: options.signal,
  });
  if (!res.ok) throw new Error((await res.json()).detail || "Extraction failed");
  return res.json();
}

export async function getExtractions() {
  const res = await fetch(`${API}/extractions`);
  return res.json();
}

export async function deleteExtraction(id) {
  await fetch(`${API}/extractions/${id}`, { method: "DELETE" });
}

export async function clearExtractions() {
  await fetch(`${API}/extractions`, { method: "DELETE" });
}

export async function exportExcel(extractions) {
  const res = await fetch(`${API}/export-excel`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ extractions }),
  });
  if (!res.ok) throw new Error("Export failed");
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "plan_extractions.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

export async function healthCheck() {
  try {
    const res = await fetch(`${API}/health`);
    if (!res.ok) return false;
    const data = await res.json();
    return Boolean(data.llm);
  } catch {
    return false;
  }
}

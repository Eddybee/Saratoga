import { useState, useRef, useCallback, useEffect } from "react";
import "./App.css";
import {
  uploadPdf,
  getPageImageUrl,
  extractAoi,
  deleteExtraction,
  clearExtractions,
  exportExcel,
  healthCheck,
} from "./api";

function formatUsd(value) {
  if (!Number.isFinite(value) || value <= 0) return "$0.00";
  if (value < 0.01) return "<$0.01";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function formatTokens(value) {
  if (!Number.isFinite(value) || value <= 0) return "0";
  return new Intl.NumberFormat("en-US").format(value);
}

function formatExtractionRoute(method) {
  if (!method) return "Unknown";
  if (method === "llm_gpt5_mini") return "GPT-5 Mini Vision";
  if (method === "llm_gpt5_mini_raw") return "GPT-5 Mini Vision (raw)";
  return method.replace(/_/g, " ");
}

function getLlmDisplayTable(ext) {
  const headers = ["Item", "Dimensions", "Notes"];
  const table = Array.isArray(ext.table_data) ? ext.table_data[0] : null;

  if (Array.isArray(table) && table.length > 1) {
    return [
      headers,
      ...table.slice(1).map((row) => {
        const cells = Array.isArray(row) ? row.map((cell) => `${cell ?? ""}`.trim()) : ["", "", `${row ?? ""}`.trim()];
        if (cells.length < 3) return [...cells, ...Array(3 - cells.length).fill("")];
        if (cells.length > 3) return [cells[0], cells[1], cells.slice(2).filter(Boolean).join(" | ")];
        return cells;
      }),
    ];
  }

  const fallbackText = `${ext.cleaned_text || ext.raw_text || ""}`.trim();
  return fallbackText ? [headers, ["", "", fallbackText]] : [headers];
}

/* ───── Toast helper ───── */
function Toast({ message, type, onDone }) {
  useEffect(() => {
    const t = setTimeout(onDone, 3000);
    return () => clearTimeout(t);
  }, [onDone]);
  return <div className={`toast ${type}`}>{message}</div>;
}

/* ───── Label prompt modal ───── */
function LabelModal({ onSubmit, onCancel, onDiscard }) {
  const [value, setValue] = useState("");
  const ref = useRef();
  useEffect(() => ref.current?.focus(), []);
  const handleSubmit = (e) => {
    e.preventDefault();
    onSubmit(value.trim());
  };
  return (
    <div className="label-modal-overlay" onClick={onDiscard}>
      <form className="label-modal" onClick={(e) => e.stopPropagation()} onSubmit={handleSubmit}>
        <h3>Label this extraction</h3>
        <input
          ref={ref}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder='e.g. "Floor Area", "Window Schedule"'
        />
        <div className="label-modal-actions">
          <button type="button" className="btn-cancel" onClick={onDiscard}>
            ✕ Cancel
          </button>
          <button type="button" className="btn-outline" onClick={onCancel}>
            Skip
          </button>
          <button type="submit" className="btn-primary">
            Extract
          </button>
        </div>
      </form>
    </div>
  );
}

/* ───── Extraction result card ───── */
function ExtractionCard({ ext, onDelete }) {
  const [expanded, setExpanded] = useState(false);
  const [showRaw, setShowRaw] = useState(false);
  const [copied, setCopied] = useState(false);
  const llmUsage = ext.llm_usage || {};
  const hasUsage = Boolean(
    llmUsage.input_tokens || llmUsage.output_tokens || llmUsage.estimated_cost_usd
  );
  const llmTable = getLlmDisplayTable(ext);
  const routeLabel = formatExtractionRoute(ext.extraction_method);

  const confClass =
    ext.avg_confidence >= 70
      ? "confidence-high"
      : ext.avg_confidence >= 40
      ? "confidence-mid"
      : "confidence-low";

  const displayText = showRaw
    ? ext.raw_text
    : ext.cleaned_text || ext.raw_text;

  const handleCopy = () => {
    navigator.clipboard.writeText(displayText || "");
    setCopied(true);
    setTimeout(() => setCopied(false), 1500);
  };

  return (
    <div className={`extraction-card ${expanded ? "expanded" : ""}`}>
      <div className="extraction-card-header">
        <div className="extraction-card-copy">
          <div className="extraction-label">{ext.label || "Unlabeled"}</div>
          <div className="extraction-meta">
            <span>Page {ext.page + 1}</span>
            <span className={`confidence-badge ${confClass}`}>
              {ext.avg_confidence}% conf
            </span>
            {ext.extraction_method && (
              <span className="method-badge" title={`Extraction method: ${ext.extraction_method}`}>
                {routeLabel}
              </span>
            )}
          </div>
          <div className="extraction-chip-row">
            <span className="engine-pill vision">Vision LLM</span>
            <span className="route-pill vision">{routeLabel}</span>
            {hasUsage && (
              <span className="usage-pill compact">{formatUsd(llmUsage.estimated_cost_usd)} run</span>
            )}
          </div>
        </div>
        <div className="extraction-actions">
          <button
            className="extraction-action-btn"
            onClick={handleCopy}
            title="Copy text"
          >
            {copied ? "✓" : "📋"}
          </button>
          <button
            className="extraction-action-btn"
            onClick={() => setExpanded((v) => !v)}
            title={expanded ? "Collapse" : "Expand"}
          >
            {expanded ? "▲" : "▼"}
          </button>
          <button className="extraction-delete" onClick={() => onDelete(ext.id)}>
            ✕
          </button>
        </div>
      </div>

      {ext.cropped_preview && (
        <div className={`extraction-preview ${expanded ? "preview-expanded" : ""}`}>
          <img src={`data:image/png;base64,${ext.cropped_preview}`} alt="crop" />
        </div>
      )}

      {hasUsage && (
        <div className="usage-strip">
          <div className="usage-pill">
            <span>Cost</span>
            <strong>{formatUsd(llmUsage.estimated_cost_usd)}</strong>
          </div>
          <div className="usage-pill">
            <span>Input</span>
            <strong>{formatTokens(llmUsage.input_tokens)} tok</strong>
          </div>
          <div className="usage-pill">
            <span>Output</span>
            <strong>{formatTokens(llmUsage.output_tokens)} tok</strong>
          </div>
        </div>
      )}

      {/* LLM Table view — render as a proper HTML table */}
      {llmTable && (
        <div className="extraction-table-wrap">
          <table className="extraction-table">
            {llmTable.length > 0 && (
              <thead>
                <tr>
                  {llmTable[0].map((cell, ci) => (
                    <th key={ci}>{cell}</th>
                  ))}
                </tr>
              </thead>
            )}
            <tbody>
              {llmTable.slice(1).map((row, ri) => (
                <tr key={ri}>
                  {row.map((cell, ci) => (
                    <td key={ci}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Raw / Cleaned toggle + full text (shown when expanded or if no structured lines) */}
      {(expanded || !ext.structured_lines?.length) && (
        <div className="extraction-text-section">
          <div className="text-toggle">
            <button
              className={`toggle-btn ${!showRaw ? "active" : ""}`}
              onClick={() => setShowRaw(false)}
            >
              Cleaned
            </button>
            <button
              className={`toggle-btn ${showRaw ? "active" : ""}`}
              onClick={() => setShowRaw(true)}
            >
              Raw
            </button>
          </div>
          <div className={`extraction-text ${expanded ? "text-expanded" : ""}`}>
            {displayText || <em style={{ color: "#aaa" }}>No text detected</em>}
          </div>
        </div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════
   Main App
   ═══════════════════════════════════════════ */
export default function App() {
  // PDF state
  const [pdf, setPdf] = useState(null); // { pdf_id, filename, pages, page_sizes }
  const [currentPage, setCurrentPage] = useState(0);
  const [pageUrl, setPageUrl] = useState(null);

  // Selection state
  const [isDrawing, setIsDrawing] = useState(false);
  const [selStart, setSelStart] = useState(null);
  const [selRect, setSelRect] = useState(null);
  const [pendingRect, setPendingRect] = useState(null); // waiting for label

  // Extractions
  const [extractions, setExtractions] = useState([]);

  // UI state
  const [loading, setLoading] = useState(null); // string message or null
  const [canCancelExtraction, setCanCancelExtraction] = useState(false);
  const [toast, setToast] = useState(null);
  const [backendOk, setBackendOk] = useState(true);

  const imgRef = useRef();
  const imageStageRef = useRef();
  const fileInputRef = useRef();
  const selStartRef = useRef(null);
  const selRectRef = useRef(null);
  const extractionAbortRef = useRef(null);

  // Check backend health on mount
  useEffect(() => {
    healthCheck().then(setBackendOk);
  }, []);

  // Load page image when pdf or page changes
  useEffect(() => {
    if (!pdf) return;
    getPageImageUrl(pdf.pdf_id, currentPage).then(setPageUrl);
  }, [pdf, currentPage]);

  /* ── Upload handler ── */
  const handleUpload = useCallback(async (file) => {
    if (!file || !file.name.toLowerCase().endsWith(".pdf")) {
      setToast({ message: "Please select a PDF file", type: "error" });
      return;
    }
    setLoading("Uploading PDF…");
    try {
      const result = await uploadPdf(file);
      setPdf(result);
      setCurrentPage(0);
      setExtractions([]);
      setToast({ message: `Loaded ${result.filename} (${result.pages} pages)`, type: "success" });
    } catch (err) {
      setToast({ message: err.message, type: "error" });
    } finally {
      setLoading(null);
    }
  }, []);

  /* ── Drag & drop ── */
  const [dragging, setDragging] = useState(false);

  const onDragOver = (e) => {
    e.preventDefault();
    setDragging(true);
  };
  const onDragLeave = () => setDragging(false);
  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    handleUpload(file);
  };

  /* ── Rectangle drawing on page image ── */
  const getRelativePos = (e) => {
    const rect = imageStageRef.current.getBoundingClientRect();
    return {
      x: Math.min(Math.max(e.clientX - rect.left, 0), rect.width),
      y: Math.min(Math.max(e.clientY - rect.top, 0), rect.height),
    };
  };

  const buildSelectionRect = (start, end) => ({
    x: Math.min(start.x, end.x),
    y: Math.min(start.y, end.y),
    width: Math.abs(end.x - start.x),
    height: Math.abs(end.y - start.y),
  });

  const finishSelection = (rect) => {
    setIsDrawing(false);
    if (rect && rect.width > 10 && rect.height > 10) {
      setPendingRect(rect);
    }
    setSelRect(null);
    setSelStart(null);
    selStartRef.current = null;
    selRectRef.current = null;
  };

  const onMouseDown = (e) => {
    if (e.button !== 0) return;
    const pos = getRelativePos(e);
    selStartRef.current = pos;
    selRectRef.current = null;
    setSelStart(pos);
    setSelRect(null);
    setIsDrawing(true);
  };

  const onMouseMove = (e) => {
    if (!isDrawing || !selStartRef.current) return;
    const pos = getRelativePos(e);
    const nextRect = buildSelectionRect(selStartRef.current, pos);
    selRectRef.current = nextRect;
    setSelRect(nextRect);
  };

  const onMouseUp = (e) => {
    if (!isDrawing) return;
    let nextRect = selRectRef.current;
    if (e && selStartRef.current) {
      nextRect = buildSelectionRect(selStartRef.current, getRelativePos(e));
    }
    finishSelection(nextRect);
  };

  const onMouseLeave = () => {
    if (!isDrawing) return;
    finishSelection(selRectRef.current);
  };

  const cancelExtractionRequest = () => {
    if (!extractionAbortRef.current) return;
    extractionAbortRef.current.abort();
    extractionAbortRef.current = null;
    setCanCancelExtraction(false);
    setLoading(null);
    setToast({ message: "Extraction cancelled", type: "info" });
  };

  /* ── Submit AOI extraction ── */
  const submitExtraction = useCallback(
    async (label) => {
      if (!pendingRect || !pdf) return;
      const stageRect = imageStageRef.current.getBoundingClientRect();
      const displayW = stageRect.width;
      const displayH = stageRect.height;

      // Phase 1: Convert pixel selection → percentages (0-100)
      const pct_x = (pendingRect.x / displayW) * 100;
      const pct_y = (pendingRect.y / displayH) * 100;
      const pct_w = (pendingRect.width / displayW) * 100;
      const pct_h = (pendingRect.height / displayH) * 100;

      const payload = {
        pdf_id: pdf.pdf_id,
        page: currentPage,
        pct_x,
        pct_y,
        pct_w,
        pct_h,
        // Legacy fields for backward compat
        x: pendingRect.x,
        y: pendingRect.y,
        width: pendingRect.width,
        height: pendingRect.height,
        image_width: displayW,
        image_height: displayH,
        label: label || "",
      };
      const controller = new AbortController();
      extractionAbortRef.current = controller;
      setPendingRect(null);
      setLoading("🤖 Sending to GPT-5 Mini Vision…");
      setCanCancelExtraction(true);
      try {
        const result = await extractAoi(payload, { signal: controller.signal });
        setExtractions((prev) => [...prev, result]);
        setToast({ message: "Text extracted!", type: "success" });
      } catch (err) {
        if (err.name === "AbortError") {
          return;
        }
        setToast({ message: err.message, type: "error" });
      } finally {
        extractionAbortRef.current = null;
        setCanCancelExtraction(false);
        setLoading(null);
      }
    },
    [pendingRect, pdf, currentPage]
  );

  /* ── Delete / Clear / Export ── */
  const handleDelete = async (id) => {
    try {
      await deleteExtraction(id);
      setExtractions((prev) => prev.filter((e) => e.id !== id));
    } catch {
      setToast({ message: "Delete failed", type: "error" });
    }
  };

  const handleClear = async () => {
    try {
      await clearExtractions();
      setExtractions([]);
    } catch {
      setToast({ message: "Clear failed", type: "error" });
    }
  };

  const handleExport = async () => {
    if (!extractions.length) return;
    setLoading("Exporting to Excel…");
    try {
      await exportExcel(extractions);
      setToast({ message: "Excel downloaded!", type: "success" });
    } catch (err) {
      setToast({ message: err.message, type: "error" });
    } finally {
      setLoading(null);
    }
  };

  // Markers for current page
  const pageMarkers = extractions.filter(
    (e) => e.pdf_id === pdf?.pdf_id && e.page === currentPage
  );
  const llmExtractions = extractions.filter((e) => e.extraction_method?.startsWith("llm")).length;
  const totalEstimatedCost = extractions.reduce(
    (sum, ext) => sum + (ext.llm_usage?.estimated_cost_usd || 0),
    0
  );
  const totalInputTokens = extractions.reduce(
    (sum, ext) => sum + (ext.llm_usage?.input_tokens || 0),
    0
  );
  const currentPageSize = pdf?.page_sizes?.[currentPage];

  /* ════════════════════ Render ════════════════════ */
  return (
    <div className="app">
      {/* Header */}
      <header className="app-header">
        <div className="brand-block">
          <div className="brand-kicker">Vision-first construction capture</div>
          <h1>
            <span>Saratoga</span> Plan Extractor
          </h1>
        </div>
        <div className="header-actions">
          {pdf && (
            <div className="header-status-group">
              <div className="header-stat">
                <span>Vision runs</span>
                <strong>{llmExtractions}</strong>
              </div>
              <div className="header-stat">
                <span>Est. cost</span>
                <strong>{formatUsd(totalEstimatedCost)}</strong>
              </div>
            </div>
          )}
          {!backendOk && (
            <span className="backend-alert">
              Backend offline. Start the FastAPI server on :8000.
            </span>
          )}
          {pdf && (
            <button
              className="btn-outline btn-sm btn-icon"
              onClick={() => fileInputRef.current?.click()}
            >
              📄 New PDF
            </button>
          )}
        </div>
      </header>

      {/* Body */}
      <div className="app-body">
        {!pdf ? (
          /* ── Upload Screen ── */
          <div className="upload-screen">
            <div className="upload-shell">
              <div className="upload-copy">
                <span className="section-kicker">Blueprint intelligence workspace</span>
                <h2>Pull schedules, notes, and quantities out of plan sheets faster.</h2>
                <p>
                  Upload a PDF, box the exact region you care about, and let the vision model return
                  structured text, schedules, and spreadsheet-ready results.
                </p>
                <div className="upload-highlights">
                  <div className="upload-highlight">
                    <strong>Vision-first extraction</strong>
                    <span>Built for messy schedules, room legends, and dense plan notes.</span>
                  </div>
                  <div className="upload-highlight">
                    <strong>Region-based workflow</strong>
                    <span>Extract only what matters instead of parsing an entire set at once.</span>
                  </div>
                  <div className="upload-highlight">
                    <strong>Excel-ready output</strong>
                    <span>Review every crop before exporting a clean handoff to estimating.</span>
                  </div>
                </div>
              </div>
              <div
                className={`upload-zone ${dragging ? "dragging" : ""}`}
                onDragOver={onDragOver}
                onDragLeave={onDragLeave}
                onDrop={onDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <div className="icon">📄</div>
                <h2>Upload a Construction Plan</h2>
                <p>Drag and drop a PDF here</p>
                <div className="or">or</div>
                <button className="btn-primary" type="button">
                  Browse Files
                </button>
                <div className="upload-hint">
                  Best results: high-resolution plan sheets with schedules, legends, room data, or notes.
                </div>
              </div>
            </div>
          </div>
        ) : (
          /* ── Viewer ── */
          <div className="viewer">
            <div className="viewer-main">
              {/* Toolbar */}
              <div className="viewer-toolbar">
                <div className="toolbar-group">
                  <div className="toolbar-copy">
                    <span className="pdf-name">{pdf.filename}</span>
                    {currentPageSize && (
                      <span className="pdf-dimensions">
                        {currentPageSize[0]} × {currentPageSize[1]} px
                      </span>
                    )}
                  </div>
                </div>
                <div className="toolbar-group">
                  <div className="page-nav">
                    <button
                      className="btn-outline btn-sm"
                      disabled={currentPage === 0}
                      onClick={() => setCurrentPage((p) => p - 1)}
                    >
                      ◀
                    </button>
                    <span>
                      Page {currentPage + 1} / {pdf.pages}
                    </span>
                    <button
                      className="btn-outline btn-sm"
                      disabled={currentPage >= pdf.pages - 1}
                      onClick={() => setCurrentPage((p) => p + 1)}
                    >
                      ▶
                    </button>
                  </div>
                </div>
                <div className="toolbar-group">
                  <span className="viewer-instruction">
                    Draw a rectangle to send that crop through the vision model.
                  </span>
                </div>
              </div>

              {/* Canvas */}
              <div className="canvas-container">
                <div className="page-wrapper">
                  <div className="page-image-stage" ref={imageStageRef}>
                    {pageUrl && (
                      <img
                        ref={imgRef}
                        src={pageUrl}
                        alt={`Page ${currentPage + 1}`}
                        draggable={false}
                      />
                    )}
                    <div
                      className="selection-overlay"
                      onMouseDown={onMouseDown}
                      onMouseMove={onMouseMove}
                      onMouseUp={onMouseUp}
                      onMouseLeave={onMouseLeave}
                    >
                      {/* Active selection rectangle */}
                      {selRect && (
                        <div
                          className="selection-rect"
                          style={{
                            left: selRect.x,
                            top: selRect.y,
                            width: selRect.width,
                            height: selRect.height,
                          }}
                        />
                      )}
                      {/* Previous extraction markers on this page */}
                      {pageMarkers.map((ext) => {
                        const r = ext.region || {};
                        const markerStyle = r.pct_x !== undefined
                          ? {
                              left: `${r.pct_x}%`,
                              top: `${r.pct_y}%`,
                              width: `${r.pct_w}%`,
                              height: `${r.pct_h}%`,
                            }
                          : {
                              left: r.x,
                              top: r.y,
                              width: r.width ?? r.w,
                              height: r.height ?? r.h,
                            };
                        return (
                          <div key={ext.id} className="aoi-marker-layer" style={markerStyle}>
                            <div className="aoi-marker">
                              <span className="aoi-label">
                                {ext.label || ext.id}
                              </span>
                            </div>
                            <button
                              type="button"
                              className="aoi-marker-delete"
                              title="Delete this extraction"
                              onClick={(event) => {
                                event.stopPropagation();
                                handleDelete(ext.id);
                              }}
                            >
                              ✕
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* Side Panel */}
            <div className="side-panel">
              <div className="panel-header">
                <div className="panel-heading">
                  <h3>Extractions</h3>
                  <p>Review structured captures before export.</p>
                </div>
                <span className="count">{extractions.length}</span>
              </div>
              <div className="panel-summary">
                <div className="summary-tile">
                  <span>This page</span>
                  <strong>{pageMarkers.length}</strong>
                </div>
                <div className="summary-tile">
                  <span>Vision input</span>
                  <strong>{formatTokens(totalInputTokens)}</strong>
                </div>
                <div className="summary-tile">
                  <span>Est. spend</span>
                  <strong>{formatUsd(totalEstimatedCost)}</strong>
                </div>
              </div>
              <div className="panel-body">
                {extractions.length === 0 ? (
                  <div className="panel-empty">
                    <div className="icon">🔍</div>
                    <p>
                      Draw a rectangle on the plan
                      <br />
                      to extract a schedule, note block, or quantity callout.
                    </p>
                  </div>
                ) : (
                  [...extractions].reverse().map((ext) => (
                    <ExtractionCard
                      key={ext.id}
                      ext={ext}
                      onDelete={handleDelete}
                    />
                  ))
                )}
              </div>
              {extractions.length > 0 && (
                <div className="panel-footer">
                  <div className="panel-footer-copy">
                    <span>{llmExtractions} vision captures ready for handoff</span>
                    <strong>{formatUsd(totalEstimatedCost)} estimated run cost</strong>
                  </div>
                  <div className="panel-footer-actions">
                    <button className="btn-danger btn-sm" onClick={handleClear}>
                      Clear All
                    </button>
                    <button className="btn-primary btn-sm btn-icon" onClick={handleExport}>
                      📥 Export Excel
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".pdf"
        style={{ display: "none" }}
        onChange={(e) => {
          handleUpload(e.target.files[0]);
          e.target.value = "";
        }}
      />

      {/* Label modal */}
      {pendingRect && (
        <LabelModal
          onSubmit={submitExtraction}
          onCancel={() => submitExtraction("")}
          onDiscard={() => setPendingRect(null)}
        />
      )}

      {/* Loading overlay */}
      {loading && (
        <div className="loading-overlay">
          <div className="spinner" />
          <p>{loading}</p>
          {canCancelExtraction && (
            <button type="button" className="btn-outline btn-sm loading-cancel-btn" onClick={cancelExtractionRequest}>
              Cancel Extraction
            </button>
          )}
        </div>
      )}

      {/* Toast notification */}
      {toast && (
        <Toast
          message={toast.message}
          type={toast.type}
          onDone={() => setToast(null)}
        />
      )}
    </div>
  );
}

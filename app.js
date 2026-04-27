import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const COLUMNS = ["Name","Email","Mobile","PINCode","BillAddress1","BillAddress2","ShipAddress1","ShipAddress2","ShipPINCode","EcommType","GSTIN","OrderID","MasterSKU","Qty"];
const CONCURRENCY = 3;

function splitAddr(addr) {
  if (!addr || addr.length <= 100) return [addr || "", ""];
  let cut = 100;
  while (cut > 0 && addr[cut] !== "," && addr[cut] !== " ") cut--;
  if (cut === 0) cut = 100;
  return [addr.slice(0, cut).trim(), addr.slice(cut).replace(/^[, ]+/, "").trim()];
}

function enforceAddrLimits(row) {
  const [b1, b2] = splitAddr((row.BillAddress1 || "") + (row.BillAddress2 ? ", " + row.BillAddress2 : ""));
  const [s1, s2] = splitAddr((row.ShipAddress1 || "") + (row.ShipAddress2 ? ", " + row.ShipAddress2 : ""));
  return { ...row, BillAddress1: b1, BillAddress2: b2, ShipAddress1: s1, ShipAddress2: s2 };
}

const SYSTEM_PROMPT = `You are a Purchase Order data extraction engine. Extract data from the PO PDF and return ONLY a valid JSON array — no markdown, no explanation, no extra text.

Each element represents one line item. Extract:
- Name: Buyer company name from the TOP of the PO letterhead
- Email: Delivery address email
- Mobile: Phone at delivery address, else ""
- PINCode: PIN code from buyer's address at top of PO (number)
- BillAddress1: Street/locality from BUYER'S address block at TOP of PO. No PIN code. Max 100 chars. Split overflow into BillAddress2.
- BillAddress2: Overflow from BillAddress1 if >100 chars, else "". No PIN code.
- ShipAddress1: Street/locality from the "Delivery Address" section ONLY. No PIN code. Max 100 chars. Split overflow into ShipAddress2.
- ShipAddress2: Overflow from ShipAddress1 if >100 chars, else "". No PIN code.
- ShipPINCode: PIN code from delivery address (number)
- EcommType: Always "B2B"
- GSTIN: Delivery address GSTN number
- OrderID: PO number (number)
- MasterSKU: Material description of the line item
- Qty: Quantity (number)

RULES:
1. BillAddress = buyer letterhead address at top. NOT vendor. NOT delivery.
2. ShipAddress = ONLY the "Delivery Address" block.
3. No PIN codes inside any address field ever.
4. All address fields max 100 chars, break at comma/space.

Return ONLY JSON array: [{"Name":"...","Email":"...","Mobile":"","PINCode":160071,"BillAddress1":"...","BillAddress2":"","ShipAddress1":"...","ShipAddress2":"","ShipPINCode":560062,"EcommType":"B2B","GSTIN":"...","OrderID":5000650406,"MasterSKU":"...","Qty":1}]`;

const pdfToBase64 = (file) => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = () => res(r.result.split(",")[1]);
  r.onerror = () => rej(new Error("Failed to read file"));
  r.readAsDataURL(file);
});

// Robust JSON extractor — handles common AI formatting issues
const extractJSON = (raw) => {
  // 1. Strip markdown fences
  let text = raw.replace(/```json|```/g, "").trim();
  // 2. Find the outermost [ ... ] array
  const start = text.indexOf("[");
  const end = text.lastIndexOf("]");
  if (start === -1 || end === -1) throw new Error("No JSON array found in response");
  text = text.slice(start, end + 1);
  // 3. Try direct parse first
  try { return JSON.parse(text); } catch (_) {}
  // 4. Fix common issues: trailing commas before } or ]
  text = text
    .replace(/,\s*([}\]])/g, "$1")       // trailing commas
    .replace(/([{,]\s*)(\w+)\s*:/g, '$1"$2":') // unquoted keys
    .replace(/:\s*'([^']*)'/g, ': "$1"') // single-quoted values
    .replace(/[ -]/g, " ");   // control characters
  try { return JSON.parse(text); } catch (_) {}
  // 5. Last resort: extract individual objects
  const objects = [];
  const objRegex = /\{[^{}]+\}/g;
  let match;
  while ((match = objRegex.exec(text)) !== null) {
    try { objects.push(JSON.parse(match[0])); } catch (_) {}
  }
  if (objects.length) return objects;
  throw new Error("Could not parse response as JSON");
};

const callAPI = async (base64, retries = 3) => {
  for (let attempt = 0; attempt <= retries; attempt++) {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 1024,
        system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: [
          { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } },
          { type: "text", text: "Extract all line items and return as a JSON array." }
        ]}]
      })
    });
    const data = await response.json();
    if (response.ok) return data;
    const errMsg = data.error?.message || `HTTP ${response.status}`;
    if (attempt < retries && (response.status === 529 || response.status === 429 || response.status >= 500)) {
      await new Promise(r => setTimeout(r, 2000 * Math.pow(2, attempt)));
      continue;
    }
    throw new Error(errMsg);
  }
};

export default function POExtractor() {
  const [rows, setRows] = useState([]);
  const [queue, setQueue] = useState([]);
  const [dragging, setDragging] = useState(false);
  const [compareIds, setCompareIds] = useState(null);
  const [compareFileName, setCompareFileName] = useState("");
  const fileRef = useRef();
  const compareRef = useRef();

  const updateQueue = (name, update) =>
    setQueue(prev => prev.map(q => q.name === name ? { ...q, ...update } : q));

  const processSingleFile = async (file, base64) => {
    updateQueue(file.name, { status: "processing", error: "" });
    try {
      const data = await callAPI(base64);
      const text = data.content.map(b => b.text || "").join("").trim();
      const parsed = extractJSON(text);
      const safe = parsed.map(enforceAddrLimits);
      setRows(prev => [...prev, ...safe]);
      updateQueue(file.name, { status: "done", count: safe.length });
    } catch (e) {
      updateQueue(file.name, { status: "error", error: e.message });
    }
  };

  const runBatch = useCallback(async (files) => {
    // Pre-read all PDFs to base64 in parallel
    const b64Map = new Map();
    await Promise.all(files.map(async f => {
      try { b64Map.set(f.name, await pdfToBase64(f)); }
      catch (_) { b64Map.set(f.name, null); }
    }));

    // Process with concurrency pool
    let idx = 0;
    const worker = async () => {
      while (idx < files.length) {
        const file = files[idx++];
        const b64 = b64Map.get(file.name);
        if (!b64) { updateQueue(file.name, { status: "error", error: "Failed to read PDF" }); continue; }
        await processSingleFile(file, b64);
      }
    };
    await Promise.all(Array.from({ length: Math.min(CONCURRENCY, files.length) }, worker));
  }, []);

  const addFiles = useCallback((newFiles) => {
    const pdfs = Array.from(newFiles).filter(f => f.type === "application/pdf");
    if (!pdfs.length) return;
    setQueue(prev => [...prev, ...pdfs.map(f => ({ name: f.name, file: f, status: "pending", count: 0, error: "" }))]);
    runBatch(pdfs);
  }, [runBatch]);

  const retryFailed = useCallback(() => {
    setQueue(prev => {
      const failed = prev.filter(q => q.status === "error");
      if (!failed.length) return prev;
      const updated = prev.map(q => q.status === "error" ? { ...q, status: "pending", error: "" } : q);
      runBatch(failed.map(q => q.file));
      return updated;
    });
  }, [runBatch]);

  const loadCompareFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const orderIdx = (data[0] || []).findIndex(h => h === "OrderID");
        if (orderIdx === -1) { alert("No OrderID column found."); return; }
        const ids = new Set(data.slice(1).map(r => String(r[orderIdx] || "").trim()).filter(Boolean));
        setCompareIds(ids);
        setCompareFileName(file.name);
      } catch (e) { alert("Failed to read file: " + e.message); }
    };
    reader.readAsArrayBuffer(file);
  };

  const updateCell = (ri, col, val) => setRows(prev => prev.map((r, i) => i === ri ? { ...r, [col]: val } : r));
  const deleteRow = (idx) => setRows(prev => prev.filter((_, i) => i !== idx));
  const clearAll = () => { setRows([]); setQueue([]); setCompareIds(null); setCompareFileName(""); };

  const exportExcel = (data = rows, filename = "LeadsCreation_PO_Export.xlsx") => {
    const ws = XLSX.utils.json_to_sheet(data, { header: COLUMNS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "LeadsCreation");
    XLSX.writeFile(wb, filename);
  };

  const doneCount = queue.filter(q => q.status === "done").length;
  const errorCount = queue.filter(q => q.status === "error").length;
  const isProcessing = queue.some(q => q.status === "processing" || q.status === "pending");
  const progressPct = queue.length ? Math.round((doneCount / queue.length) * 100) : 0;
  const missingRows = compareIds ? rows.filter(r => !compareIds.has(String(r.OrderID).trim())) : [];
  const missingPOIds = [...new Set(missingRows.map(r => String(r.OrderID)))];

  return (
    <div style={{ minHeight: "100vh", background: "#0f0f13", fontFamily: "'IBM Plex Mono', monospace", color: "#e8e6e0" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Fraunces:wght@600&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { height: 5px; width: 5px; }
        ::-webkit-scrollbar-track { background: #14141c; }
        ::-webkit-scrollbar-thumb { background: #2e2e3e; border-radius: 3px; }
        .cell-input { background: transparent; border: none; color: #d8d6d0; font-family: 'IBM Plex Mono', monospace; font-size: 11px; width: 100%; padding: 7px 8px; outline: none; }
        .cell-input:focus { background: rgba(201,168,76,0.06); border-radius: 3px; }
        .dropzone { border: 1.5px dashed #252530; border-radius: 10px; padding: 30px 20px; text-align: center; cursor: pointer; transition: all 0.2s; }
        .dropzone:hover, .dropzone.over { border-color: #c9a84c; background: rgba(201,168,76,0.04); }
        .btn { padding: 8px 18px; border: none; border-radius: 5px; font-family: 'IBM Plex Mono', monospace; font-size: 11px; font-weight: 600; cursor: pointer; transition: all 0.15s; letter-spacing: 0.5px; }
        .btn-gold { background: #c9a84c; color: #0a0a0e; }
        .btn-gold:hover { background: #e0c060; }
        .btn-ghost { background: transparent; color: #555; border: 1px solid #222230; }
        .btn-ghost:hover { border-color: #444; color: #999; }
        th { background: #111118; color: #c9a84c; font-size: 9.5px; font-weight: 600; letter-spacing: 1.2px; text-transform: uppercase; padding: 10px 8px; white-space: nowrap; border-right: 1px solid #1a1a24; position: sticky; top: 0; z-index: 2; }
        td { border-bottom: 1px solid #161620; border-right: 1px solid #161620; min-width: 120px; max-width: 220px; vertical-align: middle; }
        tr:nth-child(even) td { background: rgba(255,255,255,0.01); }
        tr:hover td { background: rgba(201,168,76,0.03) !important; }
        .del-btn { background: transparent; border: none; color: #3a3a4a; cursor: pointer; padding: 4px 8px; font-size: 13px; transition: color 0.15s; }
        .del-btn:hover { color: #e05555; }
        .file-row { display: flex; align-items: flex-start; gap: 10px; padding: 9px 14px; border-bottom: 1px solid #0f0f16; }
        .file-row:last-child { border-bottom: none; }
        .file-row.active { background: rgba(201,168,76,0.04); }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes fadeUp { from { opacity:0; transform:translateY(6px); } to { opacity:1; transform:translateY(0); } }
        .fade-up { animation: fadeUp 0.25s ease-out; }
        .spinner { display: inline-block; animation: spin 0.8s linear infinite; }
      `}</style>

      {/* Header */}
      <div style={{ background: "#09090d", borderBottom: "1px solid #161620", padding: "16px 28px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          <div style={{ fontFamily: "Fraunces, serif", fontSize: 19, fontWeight: 600, color: "#c9a84c" }}>PO → Leads Converter</div>
          <div style={{ fontSize: 9.5, color: "#383848", marginTop: 3, letterSpacing: 1.2, textTransform: "uppercase" }}>Batch Upload · AI Extract · Export Excel</div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {rows.length > 0 && <>
            <span style={{ fontSize: 11, color: "#444" }}>{rows.length} rows</span>
            <button className="btn btn-ghost"
              style={compareIds ? { borderColor: "#c9a84c", color: "#c9a84c" } : {}}
              onClick={() => compareRef.current.click()}>
              {compareIds ? `⚡ ${missingPOIds.length} missing` : "⚡ Compare Excel"}
            </button>
            <input ref={compareRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => { loadCompareFile(e.target.files[0]); e.target.value = ""; }} />
            <button className="btn btn-ghost" onClick={clearAll}>Clear All</button>
            <button className="btn btn-gold" onClick={() => exportExcel()}>↓ Export Excel</button>
          </>}
        </div>
      </div>

      <div style={{ padding: "22px 28px", display: "flex", flexDirection: "column", gap: 16 }}>

        {/* Upload + Queue */}
        <div style={{ display: "flex", gap: 16, alignItems: "stretch" }}>
          <div style={{ flex: queue.length > 0 ? "0 0 260px" : 1 }}>
            <div className={`dropzone ${dragging ? "over" : ""}`}
              style={{ height: "100%", minHeight: 130, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 8 }}
              onDragOver={e => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={e => { e.preventDefault(); setDragging(false); addFiles(e.dataTransfer.files); }}
              onClick={() => fileRef.current.click()}>
              <input ref={fileRef} type="file" accept=".pdf" multiple style={{ display: "none" }} onChange={e => { addFiles(e.target.files); e.target.value = ""; }} />
              <div style={{ fontSize: 24, opacity: 0.3 }}>📂</div>
              <div style={{ fontSize: 12, color: "#aaa" }}>Drop <strong style={{ color: "#c9a84c" }}>multiple PO PDFs</strong></div>
              <div style={{ fontSize: 10, color: "#383848" }}>or click to browse · {CONCURRENCY} at a time</div>
            </div>
          </div>

          {queue.length > 0 && (
            <div className="fade-up" style={{ flex: 1, border: "1px solid #161620", borderRadius: 10, overflow: "hidden", display: "flex", flexDirection: "column" }}>
              <div style={{ background: "#0a0a10", padding: "9px 14px", borderBottom: "1px solid #161620", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 10, color: "#444", letterSpacing: 1, textTransform: "uppercase" }}>
                  {doneCount}/{queue.length} done · {rows.length} rows{errorCount > 0 ? ` · ${errorCount} failed` : ""}
                </span>
                {isProcessing
                  ? <span style={{ fontSize: 10, color: "#c9a84c" }}><span className="spinner">⟳</span> {queue.filter(q => q.status === "processing").length} running</span>
                  : errorCount > 0
                    ? <button className="btn btn-ghost" style={{ fontSize: 10, padding: "4px 10px", borderColor: "#e05555", color: "#e05555" }} onClick={retryFailed}>↺ Retry {errorCount} failed</button>
                    : <span style={{ fontSize: 10, color: "#48b478" }}>✓ complete</span>
                }
              </div>
              <div style={{ height: 2, background: "#111118" }}>
                <div style={{ height: "100%", width: `${progressPct}%`, background: "#c9a84c", transition: "width 0.4s ease" }} />
              </div>
              <div style={{ overflowY: "auto", flex: 1, maxHeight: 230 }}>
                {queue.map((item, i) => (
                  <div key={i} className={`file-row ${item.status === "processing" ? "active" : ""}`}>
                    <div style={{ width: 16, textAlign: "center", flexShrink: 0, paddingTop: 1 }}>
                      {item.status === "pending"    && <span style={{ color: "#2a2a38" }}>○</span>}
                      {item.status === "processing" && <span className="spinner" style={{ color: "#c9a84c", display: "inline-block" }}>⟳</span>}
                      {item.status === "done"       && <span style={{ color: "#48b478" }}>✓</span>}
                      {item.status === "error"      && <span style={{ color: "#e05555" }}>✕</span>}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 11, color: item.status === "error" ? "#e05555" : "#aaa", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.name}</div>
                      <div style={{ fontSize: 10, marginTop: 2 }}>
                        {item.status === "pending"    && <span style={{ color: "#383848" }}>Waiting…</span>}
                        {item.status === "processing" && <span style={{ color: "#c9a84c" }}>Extracting…</span>}
                        {item.status === "done"       && <span style={{ color: "#48b478" }}>{item.count} row{item.count !== 1 ? "s" : ""} extracted</span>}
                        {item.status === "error"      && <span style={{ color: "#e05555", wordBreak: "break-word", whiteSpace: "normal" }}>{item.error || "Unknown error"}</span>}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* Empty state */}
        {rows.length === 0 && queue.length === 0 && (
          <div style={{ marginTop: 8, textAlign: "center" }}>
            <div style={{ fontSize: 10, color: "#1e1e28", letterSpacing: 1.2, textTransform: "uppercase", marginBottom: 10 }}>Fields extracted per line item</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 7, justifyContent: "center" }}>
              {COLUMNS.map(c => <span key={c} style={{ padding: "3px 10px", border: "1px solid #161620", borderRadius: 20, fontSize: 10, color: "#2a2a38" }}>{c}</span>)}
            </div>
          </div>
        )}

        {/* Missing POs panel */}
        {compareIds && rows.length > 0 && (
          <div className="fade-up" style={{ border: `1px solid ${missingPOIds.length > 0 ? "rgba(201,168,76,0.3)" : "rgba(72,180,120,0.25)"}`, borderRadius: 8, overflow: "hidden" }}>
            <div style={{ background: missingPOIds.length > 0 ? "rgba(201,168,76,0.06)" : "rgba(72,180,120,0.06)", padding: "10px 16px", borderBottom: `1px solid ${missingPOIds.length > 0 ? "rgba(201,168,76,0.15)" : "rgba(72,180,120,0.15)"}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div>
                <div style={{ fontSize: 12, fontWeight: 600, color: missingPOIds.length > 0 ? "#c9a84c" : "#48b478" }}>
                  {missingPOIds.length > 0 ? `${missingPOIds.length} PO${missingPOIds.length !== 1 ? "s" : ""} not in ${compareFileName}` : `All POs already present in ${compareFileName}`}
                </div>
                <div style={{ fontSize: 10, color: "#444", marginTop: 2 }}>Compared against {compareIds.size} POs in reference file</div>
              </div>
              {missingPOIds.length > 0 && (
                <button className="btn btn-gold" style={{ fontSize: 10, padding: "6px 14px" }} onClick={() => exportExcel(missingRows, "Missing_POs_Export.xlsx")}>↓ Export Missing Only</button>
              )}
            </div>
            {missingPOIds.length > 0 && (
              <div style={{ padding: "12px 16px", display: "flex", flexWrap: "wrap", gap: 8 }}>
                {missingPOIds.map(id => (
                  <span key={id} style={{ padding: "3px 10px", borderRadius: 4, background: "rgba(201,168,76,0.08)", border: "1px solid rgba(201,168,76,0.2)", fontSize: 11, color: "#c9a84c", fontFamily: "monospace" }}>{id}</span>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Data table */}
        {rows.length > 0 && (
          <div className="fade-up">
            <div style={{ borderRadius: 8, border: "1px solid #161620", overflow: "hidden" }}>
              <div style={{ overflowX: "auto", maxHeight: "50vh", overflowY: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                  <thead>
                    <tr>
                      <th style={{ minWidth: 36 }}>#</th>
                      {COLUMNS.map(col => <th key={col}>{col}</th>)}
                      <th style={{ minWidth: 40 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row, ri) => {
                      const isMissing = compareIds && !compareIds.has(String(row.OrderID).trim());
                      return (
                        <tr key={ri}>
                          <td style={{ textAlign: "center", fontSize: 10, padding: "0 8px", background: isMissing ? "rgba(201,168,76,0.08)" : "#0c0c12", color: isMissing ? "#c9a84c" : "#2a2a3a" }}>
                            {isMissing ? "⚡" : ri + 1}
                          </td>
                          {COLUMNS.map(col => (
                            <td key={col}>
                              <input className="cell-input" value={row[col] ?? ""} onChange={e => updateCell(ri, col, e.target.value)} />
                            </td>
                          ))}
                          <td style={{ textAlign: "center", background: "#0c0c12", padding: "0 4px" }}>
                            <button className="del-btn" onClick={() => deleteRow(ri)}>✕</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
            <div style={{ marginTop: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 10, color: "#2a2a38" }}>Click any cell to edit · Drop more PDFs to append rows</span>
              <div style={{ display: "flex", gap: 10 }}>
                <button className="btn btn-ghost" onClick={clearAll}>Clear All</button>
                <button className="btn btn-gold" onClick={() => exportExcel()}>↓ Export to Excel</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

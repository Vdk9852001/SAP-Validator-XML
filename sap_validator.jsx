import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── CSS Variables (matching reference dashboard) ────────────────────────────
const CSS = `
  :root {
    --teal:#0f766e;--teal-bg:#ccfbf1;
    --bg:#f4f6fa;--surface:#fff;--surface2:#f0f2f7;
    --border:#dde1ec;--border2:#c8cde0;--text:#1a1f36;--muted:#6b728e;
    --pass:#16a34a;--pass-bg:#dcfce7;--fail:#dc2626;--fail-bg:#fee2e2;
    --warn:#d97706;--warn-bg:#fef3c7;--info:#2563eb;--info-bg:#dbeafe;
    --accent:#4f46e5;--accent-light:#eef2ff;--accent-mid:rgba(79,70,229,.12);
    --shadow:0 1px 4px rgba(0,0,0,.08);--shadow-md:0 4px 16px rgba(0,0,0,.1);
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:13px; }
  ::-webkit-scrollbar{width:5px;height:5px}
  ::-webkit-scrollbar-track{background:transparent}
  ::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
  table{width:100%;border-collapse:collapse;font-size:12px}
  th{background:var(--surface2);color:var(--muted);font-size:10px;font-weight:700;
    text-transform:uppercase;letter-spacing:.05em;padding:9px 12px;text-align:left;
    border-bottom:1px solid var(--border)}
  td{padding:9px 12px;border-top:1px solid var(--border)}
  @keyframes spin{to{transform:rotate(360deg)}}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
  @keyframes sli{from{opacity:0;transform:translateX(16px)}to{opacity:1;transform:none}}
  @keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
`;

// ─── XML Parser ───────────────────────────────────────────────────────────────
function parseXMLWorkbook(xmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "application/xml");
  const NS = "urn:schemas-microsoft-com:office:spreadsheet";

  const getAttr = (el, name) =>
    el.getAttribute(`ss:${name}`) || el.getAttributeNS(NS, name) || "";

  function parseRow(rowEl, numCols = 120) {
    const result = Array(numCols).fill("");
    let col = 0;
    for (const cell of rowEl.children) {
      if (cell.localName !== "Cell") continue;
      const idx = getAttr(cell, "Index");
      if (idx) col = parseInt(idx) - 1;
      const dataEl = [...cell.children].find(c => c.localName === "Data");
      if (col < numCols) result[col] = dataEl ? (dataEl.textContent || "").trim() : "";
      col++;
    }
    return result;
  }

  const sheets = {};
  doc.querySelectorAll("Worksheet").forEach(ws => {
    const name = getAttr(ws, "Name") || "";
    const table = ws.querySelector("Table");
    if (!table) return;
    const rowEls = [...table.children].filter(c => c.localName === "Row");
    if (rowEls.length < 4) return;

    // Find field name row: SAP codes like ARBPL, WERKS (3-20 chars, uppercase)
    let fieldRowIdx = -1, fields = [];
    for (let i = 0; i < Math.min(rowEls.length, 10); i++) {
      const vals = parseRow(rowEls[i]).filter(v => v);
      if (vals.length >= 2 && vals[0] && /^[A-Z][A-Z0-9_]{2,19}$/.test(vals[0]) &&
          vals.filter(v => /^[A-Z][A-Z0-9_]{1,19}$/.test(v)).length > vals.length * 0.5) {
        fieldRowIdx = i; fields = parseRow(rowEls[i]); break;
      }
    }
    if (fieldRowIdx === -1) { sheets[name] = { fields: [], data: [] }; return; }

    // Skip description rows (very long text)
    let dataStart = fieldRowIdx + 1;
    while (dataStart < rowEls.length) {
      const first = parseRow(rowEls[dataStart])[0] || "";
      if (first.length > 60 || first.includes("\n")) dataStart++;
      else break;
    }

    const data = [];
    for (let i = dataStart; i < rowEls.length; i++) {
      const vals = parseRow(rowEls[i]);
      if (!vals[0]) continue;
      const rec = {};
      fields.forEach((f, idx) => { if (f) rec[f] = vals[idx] || ""; });
      data.push(rec);
    }
    sheets[name] = { fields: fields.filter(Boolean), data };
  });
  return sheets;
}

// ─── Excel Parser (SheetJS) ───────────────────────────────────────────────────
async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const headers = rows.length ? Object.keys(rows[0]) : [];
        const cleaned = rows.map(row => {
          const rec = {};
          headers.forEach(h => { rec[h] = String(row[h] ?? "").trim(); });
          return rec;
        });
        resolve({ headers, data: cleaned });
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── Dynamic Header Matcher ───────────────────────────────────────────────────
// Tries to auto-map Excel headers ↔ XML SAP field codes using heuristics
function buildDynamicFieldMap(excelHeaders, xmlFields) {
  const map = {}; // excelHeader → xmlField or null

  // Known SAP field → friendly name mappings
  const SAP_LABELS = {
    ARBPL:"Work Center", WERKS:"Plant", KTEXT:"Description", VERAN:"Person Responsible",
    VERWE:"Work Center Category", PLANV:"Usage", STEUS:"Control Key", KTSCH:"Standard Value Key",
    KOSTL:"Cost Center", AUFNR:"Order", MATNR:"Material", KUNNR:"Customer", LIFNR:"Vendor",
    BUKRS:"Company Code", VKORG:"Sales Org", VTWEG:"Distribution Channel", SPART:"Division",
    EKORG:"Purchasing Org", WERKS2:"Plant", LGORT:"Storage Location", CHARG:"Batch",
    MEINH:"Unit", MEINS:"Base Unit", ZIEME:"Unit of Issue", BRGEW:"Gross Weight",
    NTGEW:"Net Weight", VOLUM:"Volume", GEWEI:"Weight Unit", VOLEH:"Volume Unit",
    MATKL:"Material Group", MTART:"Material Type", MBRSH:"Industry Sector",
    PSTAT:"Maint Status", ERSDA:"Created On", ERNAM:"Created By",
    LAEDA:"Last Changed", AENAM:"Changed By", PRCTR:"Profit Center",
    KOSTL2:"Cost Center", LSTAR:"Activity Type", BEGDA:"Valid From", ENDDA:"Valid To",
    CANUM:"Capacity", KAPAR:"Capacity Category", SPRAS:"Language",
    STAND:"Standard Value Key", PRVBE:"Supply Area", PLANR:"Planner Group",
    VERSA:"Formula", NGRAD:"Utilization %", BEGZT:"Start Time", ENDZT:"End Time",
    PAUSE:"Break Duration", AZNOR:"Normal Capacity", RGEKZ:"Backflush",
  };

  // Build reverse: friendly → SAP code
  const labelToSap = {};
  Object.entries(SAP_LABELS).forEach(([code, label]) => {
    labelToSap[label.toLowerCase()] = code;
    labelToSap[label.toLowerCase().replace(/\s+/g, "")] = code;
    labelToSap[label.toLowerCase().replace(/[^a-z0-9]/g, "")] = code;
  });

  const norm = s => (s || "").toLowerCase().replace(/[^a-z0-9]/g, "");

  excelHeaders.forEach(header => {
    const h = norm(header);
    // 1. Direct match in XML fields
    if (xmlFields.includes(header)) { map[header] = header; return; }
    // 2. Case-insensitive SAP code match
    const direct = xmlFields.find(f => norm(f) === h);
    if (direct) { map[header] = direct; return; }
    // 3. Friendly label → SAP code → XML field
    const sapCode = labelToSap[h] || labelToSap[header.toLowerCase()];
    if (sapCode && xmlFields.includes(sapCode)) { map[header] = sapCode; return; }
    // 4. Partial match: XML field label contains header words
    const words = h.split(/\s+/).filter(w => w.length > 3);
    const partial = xmlFields.find(f => {
      const fLabel = norm(SAP_LABELS[f] || f);
      return words.length > 0 && words.every(w => fLabel.includes(w));
    });
    if (partial) { map[header] = partial; return; }
    // 5. No match
    map[header] = null;
  });

  return map;
}

// ─── Auto-detect join keys ────────────────────────────────────────────────────
function detectJoinKeys(xmlFields, excelHeaders, fieldMap) {
  // Priority: known key field patterns
  const KEY_PATTERNS = [
    ["ARBPL","WERKS"], ["MATNR","WERKS"], ["KUNNR","VKORG"], ["LIFNR","EKORG"],
    ["AUFNR"], ["MATNR"], ["KUNNR"], ["LIFNR"], ["ARBPL"],
    ["BUKRS","SAKNR"], ["PRCTR","KOKRS"], ["KOSTL","KOKRS"],
  ];
  for (const pattern of KEY_PATTERNS) {
    if (pattern.every(k => xmlFields.includes(k))) {
      return pattern;
    }
  }
  // Fallback: first 1-2 fields that look like IDs (short, likely alphanumeric)
  const candidates = xmlFields.filter(f => f.length <= 6).slice(0, 2);
  return candidates.length ? candidates : [xmlFields[0]].filter(Boolean);
}

// ─── Validation Engine ────────────────────────────────────────────────────────
function runValidation(xmlSheets, excelData, excelHeaders, joinKeys, fieldMap) {
  // Pick the primary XML sheet (first with data)
  const sheetName = Object.keys(xmlSheets).find(k => xmlSheets[k].data.length > 0);
  if (!sheetName) return { error: "No data sheets found in XML", records: [], summary: {} };

  const xmlData = xmlSheets[sheetName].data;
  const xmlFields = xmlSheets[sheetName].fields;

  // Map join keys: Excel column → XML field
  const excelJoinCols = joinKeys.map(xmlKey => {
    const excelCol = excelHeaders.find(h => fieldMap[h] === xmlKey) || xmlKey;
    return { xmlKey, excelCol };
  });

  // Build XML lookup by composite key
  const xmlMap = {};
  xmlData.forEach(rec => {
    const key = joinKeys.map(k => (rec[k] || "").toString().toUpperCase().trim()).join("||");
    xmlMap[key] = rec;
  });

  // All mapped field pairs (excel col → xml field, both exist and mapped)
  const fieldPairs = excelHeaders
    .map(h => ({ excelCol: h, xmlField: fieldMap[h] }))
    .filter(p => p.xmlField && xmlFields.includes(p.xmlField) &&
                 !joinKeys.includes(p.xmlField)); // exclude join key fields from diff

  // Validate each Excel record
  const records = [];
  let totalFieldDiffs = 0;
  const excelKeys = new Set();

  excelData.forEach(excelRec => {
    const key = excelJoinCols.map(({ excelCol }) =>
      (excelRec[excelCol] || "").toString().toUpperCase().trim()
    ).join("||");
    excelKeys.add(key);

    const xmlRec = xmlMap[key];
    const fieldChecks = [];
    let hasDiff = false;

    if (xmlRec) {
      fieldPairs.forEach(({ excelCol, xmlField }) => {
        const excelVal = (excelRec[excelCol] || "").trim();
        const xmlVal = (xmlRec[xmlField] || "").trim();
        if (excelVal === "" && xmlVal === "") return;
        const match = excelVal.toLowerCase() === xmlVal.toLowerCase();
        fieldChecks.push({ excelCol, xmlField, excelVal, xmlVal, match });
        if (!match) { hasDiff = true; totalFieldDiffs++; }
      });
    }

    records.push({
      key,
      excelKeyDisplay: excelJoinCols.map(({ excelCol }) => excelRec[excelCol] || "").join(" · "),
      excelRec, xmlRec: xmlRec || null,
      status: !xmlRec ? "missing_xml" : hasDiff ? "mismatch" : "matched",
      fieldChecks,
    });
  });

  // XML-only records
  const xmlOnly = xmlData.filter(rec => {
    const key = joinKeys.map(k => (rec[k] || "").toString().toUpperCase().trim()).join("||");
    return !excelKeys.has(key);
  });

  const matched = records.filter(r => r.status === "matched").length;
  const missingXML = records.filter(r => r.status === "missing_xml").length;
  const mismatch = records.filter(r => r.status === "mismatch").length;

  return {
    sheetName, xmlFields, fieldPairs, excelJoinCols,
    records, xmlOnly,
    summary: {
      total: records.length, matched, mismatch, missingXML,
      xmlOnly: xmlOnly.length, fieldDiffs: totalFieldDiffs,
      matchRate: records.length ? Math.round((matched / records.length) * 100) : 0,
    },
    fieldMap,
  };
}

// ─── Excel Export ─────────────────────────────────────────────────────────────
function exportToExcel(results, joinKeys) {
  if (!results) return;
  const { records, summary, fieldPairs, excelJoinCols } = results;

  // Summary sheet
  const summaryRows = [
    ["SAP Migration Validation Report"],
    ["Generated", new Date().toLocaleString()],
    [],
    ["SUMMARY", ""],
    ["Total Records", summary.total],
    ["Matched", summary.matched],
    ["Mismatches (Field Diffs)", summary.mismatch],
    ["Missing in XML", summary.missingXML],
    ["XML Only (Not in Post-Load)", summary.xmlOnly],
    ["Match Rate", summary.matchRate + "%"],
    [],
    ["JOIN KEYS USED"],
    ...excelJoinCols.map(({ excelCol, xmlKey }) => [`Excel: ${excelCol}`, `XML: ${xmlKey}`]),
  ];

  // Detail sheet: all records with field checks
  const detailHeaders = [
    "Key",
    ...excelJoinCols.map(e => `Key: ${e.excelCol}`),
    "Status",
    ...fieldPairs.flatMap(p => [`Excel: ${p.excelCol}`, `XML: ${p.xmlField}`, `Match?`]),
  ];

  const detailRows = records.map(r => {
    const keyVals = excelJoinCols.map(({ excelCol }) => r.excelRec[excelCol] || "");
    const fieldVals = fieldPairs.flatMap(p => {
      const check = r.fieldChecks.find(c => c.excelCol === p.excelCol && c.xmlField === p.xmlField);
      return check ? [check.excelVal, check.xmlVal, check.match ? "YES" : "NO"] : ["", r.xmlRec ? (r.xmlRec[p.xmlField] || "") : "N/A", ""];
    });
    return [r.key, ...keyVals, r.status.toUpperCase(), ...fieldVals];
  });

  // Mismatches sheet
  const mismatchRows = [
    ["KEY", "EXCEL COLUMN", "XML FIELD", "EXCEL VALUE", "XML VALUE"],
    ...records.flatMap(r =>
      r.fieldChecks.filter(c => !c.match).map(c => [
        r.excelKeyDisplay, c.excelCol, c.xmlField, c.excelVal, c.xmlVal
      ])
    ),
  ];

  // XML-only sheet
  const xmlOnlyHeaders = results.xmlFields.slice(0, 20);
  const xmlOnlyRows = results.xmlOnly.map(rec => xmlOnlyHeaders.map(f => rec[f] || ""));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), "Summary");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([detailHeaders, ...detailRows]), "All Records");
  if (mismatchRows.length > 1)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mismatchRows), "Mismatches");
  if (xmlOnlyRows.length)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[...xmlOnlyHeaders], ...xmlOnlyRows]), "XML Only");

  XLSX.writeFile(wb, `SAP_Validation_${Date.now()}.xlsx`);
}

// ─── AI Column Matching ───────────────────────────────────────────────────────
async function aiMatchColumns(excelHeaders, xmlFields) {
  const prompt = `You are an SAP data migration expert. Given these Excel post-load column headers and SAP XML field codes, return a JSON mapping of Excel header → SAP field code. Only map where confident. Return ONLY valid JSON object, no explanation.

Excel headers: ${JSON.stringify(excelHeaders)}
SAP XML fields: ${JSON.stringify(xmlFields)}

Rules:
- Map "Work Center" → ARBPL, "Plant" → WERKS, "Description" → KTEXT, etc.
- Use your SAP knowledge to match by meaning, not just string similarity
- If no match, set value to null
- Return: {"Excel Header": "SAP_FIELD" or null, ...}`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await response.json();
  const text = data.content?.find(b => b.type === "text")?.text || "{}";
  try {
    const clean = text.replace(/```json|```/g, "").trim();
    return JSON.parse(clean);
  } catch { return {}; }
}

// ─── Fuzzy Search Across Both Files ──────────────────────────────────────────
function fuzzySearchBothFiles(query, xmlSheets, excelData, excelHeaders, maxResults = 12) {
  const q = (query || "").toLowerCase().trim();
  if (!q) return { xmlHits: [], excelHits: [] };
  const xmlHits = [], excelHits = [];

  const primarySheet = Object.entries(xmlSheets || {}).find(([, s]) => s.data.length > 0);
  if (primarySheet) {
    const [sheetName, { data, fields }] = primarySheet;
    data.forEach((rec, rowIdx) => {
      const matchingFields = fields.filter(f => {
        const val = (rec[f] || "").toLowerCase();
        return val && (val === q || val.includes(q) || q.includes(val));
      });
      if (matchingFields.length > 0) xmlHits.push({ rowIdx, sheetName, rec, matchingFields });
    });
  }

  excelData.forEach((rec, rowIdx) => {
    const matchingFields = excelHeaders.filter(h => {
      const val = (rec[h] || "").toLowerCase();
      return val && (val === q || val.includes(q) || q.includes(val));
    });
    if (matchingFields.length > 0) excelHits.push({ rowIdx, rec, matchingFields });
  });

  const score = (hit) => {
    let s = hit.matchingFields.length;
    hit.matchingFields.forEach(f => {
      const val = (hit.rec[f] || "").toLowerCase();
      if (val === q) s += 10;
      else if (val.startsWith(q)) s += 4;
    });
    return s;
  };

  return {
    xmlHits: xmlHits.sort((a, b) => score(b) - score(a)).slice(0, maxResults),
    excelHits: excelHits.sort((a, b) => score(b) - score(a)).slice(0, maxResults),
  };
}

function HL({ text, query }) {
  const t = String(text || ""), q = (query || "").toLowerCase();
  if (!q) return <>{t}</>;
  const idx = t.toLowerCase().indexOf(q);
  if (idx === -1) return <>{t}</>;
  return <>{t.slice(0, idx)}<mark style={{background:"#fef08a",borderRadius:2,padding:"0 1px"}}>{t.slice(idx, idx + q.length)}</mark>{t.slice(idx + q.length)}</>;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
const esc = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const fmt = n => Number(n || 0).toLocaleString();

// ─── Sub-components ───────────────────────────────────────────────────────────
function Badge({ type, children }) {
  const styles = {
    pass: { background: "var(--pass-bg)", color: "var(--pass)", border: "1px solid rgba(22,163,74,.3)" },
    fail: { background: "var(--fail-bg)", color: "var(--fail)", border: "1px solid rgba(220,38,38,.3)" },
    warn: { background: "var(--warn-bg)", color: "var(--warn)", border: "1px solid rgba(217,119,6,.3)" },
    key:  { background: "var(--accent-light)", color: "var(--accent)", border: "1px solid rgba(79,70,229,.3)" },
    info: { background: "var(--info-bg)", color: "var(--info)", border: "1px solid rgba(37,99,235,.3)" },
  };
  return (
    <span style={{
      ...styles[type], fontSize: 10, fontWeight: 700, padding: "2px 8px",
      borderRadius: 20, display: "inline-block",
    }}>{children}</span>
  );
}

function Card({ value, label, type }) {
  const colors = { ok: "var(--pass)", warn: "var(--fail)", blue: "var(--info)", "": "var(--text)" };
  return (
    <div style={{
      background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 10,
      padding: "13px 14px", boxShadow: "var(--shadow)", flex: 1, minWidth: 100,
    }}>
      <div style={{ fontSize: 22, fontWeight: 700, color: colors[type || ""] }}>{value}</div>
      <div style={{ fontSize: 10, color: "var(--muted)", marginTop: 2, fontWeight: 500 }}>{label}</div>
    </div>
  );
}

function Spinner() {
  return <span style={{
    width: 12, height: 12, border: "2px solid currentColor", borderTopColor: "transparent",
    borderRadius: "50%", display: "inline-block", animation: "spin .7s linear infinite",
  }} />;
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  // File state
  const [xmlSheets, setXmlSheets] = useState(null);
  const [xmlFileName, setXmlFileName] = useState("");
  const [excelData, setExcelData] = useState([]);
  const [excelHeaders, setExcelHeaders] = useState([]);
  const [excelFileName, setExcelFileName] = useState("");

  // Mapping & join key state
  const [fieldMap, setFieldMap] = useState({});       // excelHeader → xmlField | null
  const [joinKeys, setJoinKeys] = useState([]);        // XML field codes used as join key
  const [aiMappingLoading, setAiMappingLoading] = useState(false);
  const [showJoinKeyEditor, setShowJoinKeyEditor] = useState(false);
  const [pendingJoinKey, setPendingJoinKey] = useState("");

  // Validation state
  const [results, setResults] = useState(null);
  const [validating, setValidating] = useState(false);

  // UI state
  const [activeTab, setActiveTab] = useState("records");
  const [filterStatus, setFilterStatus] = useState("all");
  const [search, setSearch] = useState("");
  const [expandedRows, setExpandedRows] = useState({});
  const [toast, setToast] = useState(null);
  const [showMappingEditor, setShowMappingEditor] = useState(false);

  // ── Smart Search & Learn state ─────────────────────────────────────────────
  const [searchModal, setSearchModal] = useState(null); // { excelRec, excelKeyDisplay }
  const [smartQuery, setSmartQuery] = useState("");
  const [smartResults, setSmartResults] = useState(null);
  const [learnedMappings, setLearnedMappings] = useState([]); // [{excelField, xmlField, label}]
  const [confirmedMatch, setConfirmedMatch] = useState(null); // {excelRec, xmlRec, fields}

  const xmlRef = useRef(), excelRef = useRef();

  const showToast = (msg, type = "info") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 4000);
  };

  // ── Smart Search & Learn ────────────────────────────────────────────────────
  const runSmartSearch = useCallback((q) => {
    setSmartQuery(q);
    if (!q.trim()) { setSmartResults(null); return; }
    const r = fuzzySearchBothFiles(q, xmlSheets, excelData, excelHeaders);
    setSmartResults(r);
    setConfirmedMatch(null);
  }, [xmlSheets, excelData, excelHeaders]);

  const applyLearnedKey = useCallback((excelRec, xmlRec, excelKeyField, xmlKeyField) => {
    // 1. Update fieldMap so excelKeyField → xmlKeyField
    setFieldMap(prev => ({ ...prev, [excelKeyField]: xmlKeyField }));
    // 2. Add xmlKeyField to join keys if not already there
    setJoinKeys(prev => prev.includes(xmlKeyField) ? prev : [...prev, xmlKeyField]);
    // 3. Record the learned mapping for the legend
    const label = `${excelKeyField} → ${xmlKeyField}`;
    setLearnedMappings(prev =>
      prev.find(m => m.label === label) ? prev : [...prev, { excelField: excelKeyField, xmlField: xmlKeyField, label }]
    );
    setConfirmedMatch({ excelRec, xmlRec, excelKeyField, xmlKeyField });
    showToast(`✓ Learned: "${excelKeyField}" maps to "${xmlKeyField}". Re-run validation to apply.`, "success");
  }, []);

  // ── Load XML ────────────────────────────────────────────────────────────────
  const handleXML = useCallback(e => {
    const file = e.target.files[0];
    if (!file) return;
    setXmlFileName(file.name);
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const sheets = parseXMLWorkbook(ev.target.result);
        setXmlSheets(sheets);
        setResults(null);
        showToast(`XML parsed: ${Object.keys(sheets).length} sheets found`, "success");
      } catch (err) { showToast("XML parse error: " + err.message, "error"); }
    };
    reader.readAsText(file);
  }, []);

  // ── Load Excel ──────────────────────────────────────────────────────────────
  const handleExcel = useCallback(async e => {
    const file = e.target.files[0];
    if (!file) return;
    setExcelFileName(file.name);
    try {
      const parsed = await parseExcelFile(file);
      setExcelData(parsed.data);
      setExcelHeaders(parsed.headers);
      setResults(null);
      showToast(`Excel loaded: ${parsed.data.length} rows, ${parsed.headers.length} columns`, "success");
    } catch (err) { showToast("Excel parse error: " + err.message, "error"); }
  }, []);

  // ── Auto-build field map when both files loaded ─────────────────────────────
  useEffect(() => {
    if (!xmlSheets || !excelHeaders.length) return;
    const primarySheet = Object.values(xmlSheets).find(s => s.data.length > 0);
    if (!primarySheet) return;
    const autoMap = buildDynamicFieldMap(excelHeaders, primarySheet.fields);
    setFieldMap(autoMap);
    const detected = detectJoinKeys(primarySheet.fields, excelHeaders, autoMap);
    setJoinKeys(detected);
  }, [xmlSheets, excelHeaders]);

  // ── AI Mapping ──────────────────────────────────────────────────────────────
  const handleAIMapping = useCallback(async () => {
    if (!xmlSheets || !excelHeaders.length) return;
    setAiMappingLoading(true);
    try {
      const primarySheet = Object.values(xmlSheets).find(s => s.data.length > 0);
      const aiMap = await aiMatchColumns(excelHeaders, primarySheet?.fields || []);
      setFieldMap(prev => {
        const merged = { ...prev };
        Object.entries(aiMap).forEach(([h, v]) => {
          if (!merged[h] && v) merged[h] = v;
        });
        return merged;
      });
      showToast("AI mapping applied successfully", "success");
    } catch (err) { showToast("AI mapping failed: " + err.message, "error"); }
    setAiMappingLoading(false);
  }, [xmlSheets, excelHeaders]);

  // ── Validate ────────────────────────────────────────────────────────────────
  const handleValidate = useCallback(() => {
    if (!xmlSheets || !excelData.length || !joinKeys.length) return;
    setValidating(true);
    setResults(null);
    setTimeout(() => {
      const r = runValidation(xmlSheets, excelData, excelHeaders, joinKeys, fieldMap);
      setResults(r);
      setValidating(false);
      setActiveTab("records");
      showToast(`Validation complete: ${r.summary.matchRate}% match rate`, r.summary.matchRate === 100 ? "success" : "warn");
    }, 300);
  }, [xmlSheets, excelData, excelHeaders, joinKeys, fieldMap]);

  // ── Filtered records ────────────────────────────────────────────────────────
  const filteredRecords = results ? results.records.filter(r => {
    if (filterStatus !== "all" && r.status !== filterStatus) return false;
    if (search && !r.excelKeyDisplay.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  }) : [];

  const primarySheet = xmlSheets ? Object.entries(xmlSheets).find(([, s]) => s.data.length > 0) : null;
  const xmlFields = primarySheet ? primarySheet[1].fields : [];
  const mappedCount = Object.values(fieldMap).filter(Boolean).length;

  // ─── Render ─────────────────────────────────────────────────────────────────
  return (
    <>
      <style>{CSS}</style>
      <div style={{ display: "flex", flexDirection: "column", height: "100vh", background: "var(--bg)" }}>

        {/* ── HEADER ── */}
        <header style={{
          background: "var(--surface)", borderBottom: "1px solid var(--border)",
          padding: "0 20px", display: "flex", alignItems: "center",
          justifyContent: "space-between", height: 54, gap: 10,
          boxShadow: "var(--shadow)", flexShrink: 0,
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{
              width: 34, height: 34,
              background: "linear-gradient(135deg,#4f46e5,#7c3aed)",
              borderRadius: 9, display: "flex", alignItems: "center", justifyContent: "center",
            }}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2">
                <ellipse cx="12" cy="6" rx="8" ry="3"/>
                <path d="M4 6v4c0 1.66 3.58 3 8 3s8-1.34 8-3V6"/>
                <path d="M4 10v4c0 1.66 3.58 3 8 3s8-1.34 8-3v-4"/>
                <path d="M4 14v4c0 1.66 3.58 3 8 3s8-1.34 8-3v-4"/>
              </svg>
            </div>
            <div>
              <div style={{ fontSize: 15, fontWeight: 700 }}>
                Genpact <span style={{ color: "var(--accent)" }}>SAP</span> Validator
              </div>
              <div style={{ fontSize: 10, color: "var(--muted)", marginTop: 1 }}>
                Post-Load Migration Validation
              </div>
            </div>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 7, flexWrap: "wrap" }}>
            {results && (
              <span style={{
                fontSize: 11, fontWeight: 600, padding: "3px 10px", borderRadius: 20,
                background: results.summary.matchRate === 100 ? "var(--pass-bg)" : "var(--warn-bg)",
                color: results.summary.matchRate === 100 ? "var(--pass)" : "var(--warn)",
              }}>
                Match Rate: {results.summary.matchRate}%
              </span>
            )}
            {learnedMappings.length > 0 && (
              <span style={{
                fontSize: 11, fontWeight: 600, padding: "3px 10px", borderRadius: 20,
                background: "var(--teal-bg)", color: "var(--teal)",
                border: "1px solid rgba(15,118,110,.25)",
              }}>
                🧠 {learnedMappings.length} learned mapping{learnedMappings.length > 1 ? "s" : ""}
              </span>
            )}
            <span style={{
              fontSize: 11, fontWeight: 600, padding: "3px 10px", borderRadius: 20,
              background: "var(--accent-light)", color: "var(--accent)",
              border: "1px solid rgba(79,70,229,.2)",
            }}>
              {mappedCount}/{excelHeaders.length} fields mapped
            </span>
            {results && (
              <button onClick={() => exportToExcel(results, joinKeys)} style={{
                background: "var(--surface2)", color: "var(--text)", border: "1px solid var(--border)",
                padding: "6px 14px", borderRadius: 7, fontSize: 12, cursor: "pointer", fontWeight: 500,
                display: "flex", alignItems: "center", gap: 5,
              }}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
                Download Excel
              </button>
            )}
            <button onClick={handleValidate}
              disabled={!xmlSheets || !excelData.length || validating || !joinKeys.length}
              style={{
                background: "var(--accent)", color: "#fff",
                border: "none", padding: "6px 14px", borderRadius: 7, fontSize: 12,
                cursor: "pointer", fontWeight: 500, opacity: (!xmlSheets || !excelData.length) ? 0.5 : 1,
                display: "flex", alignItems: "center", gap: 5,
              }}>
              {validating ? <><Spinner /> Validating...</> : <>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
                </svg>
                Run Validation
              </>}
            </button>
          </div>
        </header>

        {/* ── BODY ── */}
        <div style={{ display: "grid", gridTemplateColumns: "270px 1fr", flex: 1, overflow: "hidden" }}>

          {/* ── SIDEBAR ── */}
          <aside style={{
            background: "var(--surface)", borderRight: "1px solid var(--border)",
            display: "flex", flexDirection: "column", overflow: "hidden",
          }}>
            <div style={{ padding: "14px 15px 6px", fontSize: 10, fontWeight: 700, color: "var(--muted)", letterSpacing: ".08em", textTransform: "uppercase" }}>
              Files & Configuration
            </div>

            <div style={{ flex: 1, overflowY: "auto", padding: "0 8px 8px" }}>

              {/* XML Upload */}
              <div style={{
                background: "var(--surface2)", border: "1px solid var(--border)",
                borderRadius: 9, padding: 12, marginBottom: 8,
              }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 8 }}>
                  Step 1 — XML Template
                </div>
                <button onClick={() => xmlRef.current.click()} style={{
                  display: "block", background: "var(--accent)", color: "#fff",
                  border: "none", padding: "7px 12px", borderRadius: 7, fontSize: 12,
                  fontWeight: 600, cursor: "pointer", width: "100%", marginBottom: 6,
                }}>Browse XML file…</button>
                <input ref={xmlRef} type="file" accept=".xml" onChange={handleXML} style={{ display: "none" }} />
                <div style={{ fontSize: 11, color: xmlFileName ? "var(--pass)" : "var(--muted)", marginTop: 4 }}>
                  {xmlFileName ? `✓ ${xmlFileName}` : "No file selected"}
                </div>
                {xmlSheets && (
                  <div style={{ marginTop: 8, display: "flex", flexWrap: "wrap", gap: 4 }}>
                    {Object.entries(xmlSheets).map(([name, s]) => (
                      <div key={name} style={{
                        background: s.data.length ? "var(--teal-bg)" : "var(--surface)",
                        color: s.data.length ? "var(--teal)" : "var(--muted)",
                        border: "1px solid", borderColor: s.data.length ? "rgba(15,118,110,.25)" : "var(--border)",
                        borderRadius: 5, padding: "2px 7px", fontSize: 10, fontWeight: 600,
                      }}>
                        {name} <span style={{ fontWeight: 400, color: "var(--muted)" }}>({s.data.length})</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Excel Upload */}
              <div style={{
                background: "var(--surface2)", border: "1px solid var(--border)",
                borderRadius: 9, padding: 12, marginBottom: 8,
              }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 8 }}>
                  Step 2 — Post-Load Excel (.xlsx)
                </div>
                <button onClick={() => excelRef.current.click()} style={{
                  display: "block", background: "var(--accent)", color: "#fff",
                  border: "none", padding: "7px 12px", borderRadius: 7, fontSize: 12,
                  fontWeight: 600, cursor: "pointer", width: "100%", marginBottom: 6,
                }}>Browse Excel file…</button>
                <input ref={excelRef} type="file" accept=".xlsx,.xls" onChange={handleExcel} style={{ display: "none" }} />
                <div style={{ fontSize: 11, color: excelFileName ? "var(--pass)" : "var(--muted)", marginTop: 4 }}>
                  {excelFileName ? `✓ ${excelFileName} (${fmt(excelData.length)} rows)` : "No file selected"}
                </div>
              </div>

              {/* Join Key Panel */}
              {xmlSheets && excelHeaders.length > 0 && (
                <div style={{
                  background: "var(--surface)", border: "1px solid var(--border)",
                  borderRadius: 9, padding: 12, marginBottom: 8, boxShadow: "var(--shadow)",
                }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: ".06em" }}>
                      Join Keys
                    </div>
                    <button onClick={() => setShowJoinKeyEditor(v => !v)} style={{
                      background: "var(--accent-light)", color: "var(--accent)",
                      border: "1px solid rgba(79,70,229,.25)", padding: "3px 10px",
                      borderRadius: 6, fontSize: 11, fontWeight: 600, cursor: "pointer",
                    }}>
                      Edit
                    </button>
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 6 }}>
                    {joinKeys.map((k, i) => (
                      <span key={k} style={{
                        background: "var(--accent-light)", color: "var(--accent)",
                        border: "1px solid rgba(79,70,229,.25)", borderRadius: 6,
                        padding: "3px 10px", fontSize: 12, fontWeight: 600,
                        display: "flex", alignItems: "center", gap: 3,
                      }}>
                        {i > 0 && <span style={{ color: "var(--muted)", marginRight: 3 }}>+</span>}
                        {k}
                      </span>
                    ))}
                    {!joinKeys.length && (
                      <span style={{ fontSize: 11, color: "var(--warn)" }}>⚠ No join keys set</span>
                    )}
                  </div>
                  {showJoinKeyEditor && (
                    <div>
                      <div style={{ fontSize: 10, color: "var(--muted)", marginBottom: 6 }}>
                        Click to remove. Add from XML fields:
                      </div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 8, maxHeight: 120, overflowY: "auto" }}>
                        {joinKeys.map(k => (
                          <span key={k} onClick={() => setJoinKeys(prev => prev.filter(x => x !== k))}
                            style={{
                              background: "var(--accent)", color: "#fff", borderRadius: 5,
                              padding: "2px 8px", fontSize: 11, fontWeight: 600, cursor: "pointer",
                              display: "flex", alignItems: "center", gap: 4,
                            }}>
                            {k} <span style={{ opacity: .7 }}>×</span>
                          </span>
                        ))}
                      </div>
                      <div style={{ fontSize: 10, color: "var(--muted)", marginBottom: 4 }}>Add XML field as key:</div>
                      <div style={{ display: "flex", gap: 4 }}>
                        <input value={pendingJoinKey}
                          onChange={e => setPendingJoinKey(e.target.value.toUpperCase())}
                          onKeyDown={e => {
                            if (e.key === "Enter" && pendingJoinKey) {
                              setJoinKeys(prev => prev.includes(pendingJoinKey) ? prev : [...prev, pendingJoinKey]);
                              setPendingJoinKey("");
                            }
                          }}
                          placeholder="e.g. ARBPL" style={{
                            flex: 1, background: "var(--surface2)", border: "1px solid var(--border)",
                            color: "var(--text)", padding: "5px 8px", borderRadius: 6, fontSize: 11,
                          }} />
                        <button onClick={() => {
                          if (pendingJoinKey) {
                            setJoinKeys(prev => prev.includes(pendingJoinKey) ? prev : [...prev, pendingJoinKey]);
                            setPendingJoinKey("");
                          }
                        }} style={{
                          background: "var(--accent)", color: "#fff", border: "none",
                          borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer",
                        }}>Add</button>
                      </div>
                      <div style={{ marginTop: 6, display: "flex", flexWrap: "wrap", gap: 3, maxHeight: 100, overflowY: "auto" }}>
                        {xmlFields.filter(f => !joinKeys.includes(f)).slice(0, 30).map(f => (
                          <span key={f} onClick={() => setJoinKeys(prev => [...prev, f])}
                            title="Click to add as join key"
                            style={{
                              background: "var(--surface2)", border: "1px solid var(--border)",
                              borderRadius: 4, padding: "1px 6px", fontSize: 10, cursor: "pointer",
                              color: "var(--muted)",
                            }}>{f}</span>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* Field Mapping Panel */}
              {xmlSheets && excelHeaders.length > 0 && (
                <div style={{
                  background: "var(--surface)", border: "1px solid var(--border)",
                  borderRadius: 9, padding: 12, boxShadow: "var(--shadow)",
                }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: ".06em" }}>
                      Field Mapping ({mappedCount}/{excelHeaders.length})
                    </div>
                    <div style={{ display: "flex", gap: 4 }}>
                      <button onClick={handleAIMapping} disabled={aiMappingLoading} style={{
                        background: "#7c3aed20", color: "#7c3aed", border: "1px solid #7c3aed40",
                        padding: "3px 8px", borderRadius: 6, fontSize: 10, fontWeight: 600, cursor: "pointer",
                      }}>
                        {aiMappingLoading ? <Spinner /> : "🤖 AI"}
                      </button>
                      <button onClick={() => setShowMappingEditor(v => !v)} style={{
                        background: "var(--accent-light)", color: "var(--accent)",
                        border: "1px solid rgba(79,70,229,.25)", padding: "3px 8px",
                        borderRadius: 6, fontSize: 10, fontWeight: 600, cursor: "pointer",
                      }}>Edit</button>
                    </div>
                  </div>

                  {showMappingEditor ? (
                    <div style={{ maxHeight: 300, overflowY: "auto" }}>
                      {excelHeaders.map(h => (
                        <div key={h} style={{
                          display: "flex", alignItems: "center", gap: 6,
                          padding: "5px 0", borderBottom: "1px solid var(--border)",
                        }}>
                          <div style={{ flex: 1, fontSize: 11, color: "var(--text)", fontWeight: 500, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}
                            title={h}>{h}</div>
                          <select value={fieldMap[h] || ""}
                            onChange={e => setFieldMap(prev => ({ ...prev, [h]: e.target.value || null }))}
                            style={{
                              background: "var(--surface2)", border: "1px solid var(--border)",
                              color: fieldMap[h] ? "var(--accent)" : "var(--muted)",
                              padding: "3px 5px", borderRadius: 5, fontSize: 10, width: 90,
                            }}>
                            <option value="">— skip —</option>
                            {xmlFields.map(f => <option key={f} value={f}>{f}</option>)}
                          </select>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div style={{ maxHeight: 160, overflowY: "auto" }}>
                      {excelHeaders.map(h => (
                        <div key={h} style={{
                          display: "flex", alignItems: "center", justifyContent: "space-between",
                          padding: "3px 0", fontSize: 11,
                        }}>
                          <span style={{ color: "var(--text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{h}</span>
                          <span style={{ color: fieldMap[h] ? "var(--accent)" : "var(--muted)", fontWeight: 600, fontSize: 10, marginLeft: 4, flexShrink: 0 }}>
                            {fieldMap[h] || "—"}
                          </span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          </aside>

          {/* ── MAIN CONTENT ── */}
          <main style={{ overflowY: "auto", padding: 22 }}>

            {/* Welcome / no results */}
            {!results && !validating && (
              <div style={{
                display: "flex", flexDirection: "column", alignItems: "center",
                justifyContent: "center", height: "100%", gap: 20, textAlign: "center",
              }}>
                <h2 style={{ fontSize: 22, fontWeight: 700 }}>
                  Genpact <span style={{ color: "var(--accent)" }}>SAP</span> Validator
                </h2>
                <p style={{ color: "var(--muted)", fontSize: 13, maxWidth: 520, lineHeight: 1.7 }}>
                  Upload your SAP LTMC XML migration template and post-load Excel extract.
                  The tool automatically detects all headers, maps SAP field codes, and validates
                  every record field-by-field. Works with any SAP object type.
                </p>
                <div style={{
                  background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 12,
                  padding: 22, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14,
                  minWidth: 480, boxShadow: "var(--shadow)",
                }}>
                  {[
                    { title: "XML Template", hint: "SAP LTMC SpreadsheetML XML file" },
                    { title: "Post-Load Excel", hint: "SAPUI5 export as .xlsx" },
                  ].map(({ title, hint }) => (
                    <div key={title} style={{
                      background: "var(--surface2)", border: "2px dashed var(--border)",
                      borderRadius: 10, padding: 18, textAlign: "center",
                    }}>
                      <h4 style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>
                        ⬆ {title}
                      </h4>
                      <div style={{ fontSize: 10, color: "var(--muted)", lineHeight: 1.5 }}>{hint}</div>
                    </div>
                  ))}
                </div>
                <div style={{
                  background: "var(--accent-light)", border: "1px solid rgba(79,70,229,.2)",
                  borderRadius: 9, padding: "10px 16px", fontSize: 12, color: "var(--accent)",
                  maxWidth: 520, lineHeight: 1.7,
                }}>
                  <b>How it works:</b> Upload both files → the tool auto-maps all column headers
                  to SAP field codes using heuristics + Claude AI → validate → download Excel report.
                  Works with Work Centers, Materials, Customers, Vendors, and all 40+ SAP objects.
                </div>
              </div>
            )}

            {validating && (
              <div style={{
                display: "flex", alignItems: "center", gap: 9, padding: "10px 14px",
                borderRadius: 8, marginBottom: 12, fontSize: 12, lineHeight: 1.55,
                background: "var(--accent-light)", border: "1px solid rgba(79,70,229,.2)", color: "var(--accent)",
              }}>
                <Spinner /> <span>Running validation…</span>
              </div>
            )}

            {/* Results */}
            {results && (
              <>
                {/* Header */}
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 18, gap: 12 }}>
                  <div>
                    <div style={{ fontSize: 18, fontWeight: 700 }}>
                      {results.sheetName} Validation
                    </div>
                    <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 3 }}>
                      {xmlFileName} vs {excelFileName} — {new Date().toLocaleString()}
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                    <button onClick={() => exportToExcel(results, joinKeys)} style={{
                      background: "var(--accent-light)", color: "var(--accent)",
                      border: "1px solid rgba(79,70,229,.25)", padding: "6px 13px",
                      borderRadius: 7, fontSize: 12, fontWeight: 600, cursor: "pointer",
                      display: "flex", alignItems: "center", gap: 5,
                    }}>
                      ↓ Download Excel
                    </button>
                    <span style={{
                      fontSize: 12, fontWeight: 700, padding: "5px 16px", borderRadius: 20,
                      background: results.summary.matchRate === 100 ? "var(--pass-bg)" : results.summary.matchRate >= 80 ? "var(--warn-bg)" : "var(--fail-bg)",
                      color: results.summary.matchRate === 100 ? "var(--pass)" : results.summary.matchRate >= 80 ? "var(--warn)" : "var(--fail)",
                    }}>
                      {results.summary.matchRate === 100 ? "PASS" : results.summary.matchRate >= 80 ? "WARNING" : "FAIL"}
                    </span>
                  </div>
                </div>

                {/* Join Key Info Bar */}
                <div style={{
                  display: "inline-flex", alignItems: "center", gap: 7,
                  background: "var(--accent-light)", border: "1px solid rgba(79,70,229,.2)",
                  borderRadius: 7, padding: "5px 12px", fontSize: 11, color: "var(--accent)",
                  marginBottom: 10, fontWeight: 500, flexWrap: "wrap",
                }}>
                  <span>Join keys:</span>
                  {results.excelJoinCols.map(({ excelCol, xmlKey }, i) => (
                    <span key={i}>
                      {i > 0 && <span style={{ margin: "0 4px", opacity: .5 }}>+</span>}
                      <b>{xmlKey}</b>
                      <span style={{ opacity: .6, marginLeft: 3 }}>(Excel: {excelCol})</span>
                    </span>
                  ))}
                </div>

                {/* Summary Cards */}
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 20 }}>
                  <Card value={fmt(results.summary.total)} label="Total Records" />
                  <Card value={fmt(results.summary.matched)} label="Matched" type="ok" />
                  <Card value={fmt(results.summary.mismatch)} label="Field Diffs" type={results.summary.mismatch ? "warn" : "ok"} />
                  <Card value={fmt(results.summary.missingXML)} label="Missing in XML" type={results.summary.missingXML ? "warn" : "ok"} />
                  <Card value={fmt(results.summary.xmlOnly)} label="XML Only" type={results.summary.xmlOnly ? "warn" : ""} />
                  <Card value={results.summary.matchRate + "%"} label="Match Rate" type="blue" />
                </div>

                {/* Field Mapping Summary */}
                <div style={{
                  background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 10,
                  padding: "12px 16px", marginBottom: 20, boxShadow: "var(--shadow)",
                }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: ".07em", marginBottom: 10 }}>
                    Field Mapping ({results.fieldPairs.length} fields validated)
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {results.fieldPairs.map(({ excelCol, xmlField }) => (
                      <span key={excelCol} style={{
                        display: "inline-flex", flexDirection: "column",
                        background: "var(--surface2)", border: "1px solid rgba(79,70,229,.3)",
                        borderRadius: 6, padding: "3px 9px", margin: 2, fontSize: 11,
                        lineHeight: 1.4, fontWeight: 500, color: "var(--accent)",
                      }}>
                        {excelCol}
                        <small style={{ fontSize: 9, color: "var(--muted)", fontWeight: 400 }}>→ {xmlField}</small>
                      </span>
                    ))}
                    {excelHeaders.filter(h => !results.fieldMap[h] && !joinKeys.includes(results.fieldMap[h])).slice(0, 8).map(h => (
                      <span key={h} style={{
                        display: "inline-flex", flexDirection: "column",
                        background: "var(--warn-bg)", border: "1px solid rgba(217,119,6,.3)",
                        borderRadius: 6, padding: "3px 9px", margin: 2, fontSize: 11,
                        lineHeight: 1.4, fontWeight: 500, color: "var(--warn)",
                      }}>
                        {h}
                        <small style={{ fontSize: 9, fontWeight: 400 }}>not mapped</small>
                      </span>
                    ))}
                  </div>
                </div>

                {/* Tabs */}
                <div style={{ borderBottom: "1px solid var(--border)", marginBottom: 0, display: "flex" }}>
                  {[
                    { key: "records", label: `All Records (${results.records.length})` },
                    { key: "mismatch", label: `Mismatches (${results.summary.mismatch})` },
                    { key: "missing", label: `⚠ Missing in XML (${results.summary.missingXML})` },
                    { key: "xmlonly", label: `XML Only (${results.summary.xmlOnly})` },
                  ].map(t => (
                    <button key={t.key} onClick={() => setActiveTab(t.key)} style={{
                      background: activeTab === t.key ? "var(--surface)" : "transparent",
                      color: activeTab === t.key ? "var(--accent)" : "var(--muted)",
                      border: "none", borderBottom: activeTab === t.key ? "2px solid var(--accent)" : "2px solid transparent",
                      padding: "8px 16px", fontSize: 12, cursor: "pointer", fontWeight: 700,
                      letterSpacing: ".03em",
                    }}>{t.label}</button>
                  ))}
                </div>

                {/* Records Table */}
                <div style={{
                  background: "var(--surface)", border: "1px solid var(--border)",
                  borderRadius: "0 0 10px 10px", overflow: "hidden", marginBottom: 22, boxShadow: "var(--shadow)",
                }}>
                  {/* Filter bar */}
                  <div style={{ display: "flex", gap: 8, padding: "10px 14px", borderBottom: "1px solid var(--border)", flexWrap: "wrap", alignItems: "center" }}>
                    <input placeholder="Search by key..." value={search}
                      onChange={e => setSearch(e.target.value)} style={{
                        background: "var(--surface2)", border: "1px solid var(--border)",
                        color: "var(--text)", padding: "6px 10px", borderRadius: 7, fontSize: 12, flex: 1, minWidth: 150,
                      }} />
                    {activeTab === "records" && ["all", "matched", "mismatch", "missing_xml"].map(s => (
                      <button key={s} onClick={() => setFilterStatus(s)} style={{
                        background: filterStatus === s ? "var(--accent)" : "var(--surface2)",
                        color: filterStatus === s ? "#fff" : "var(--muted)",
                        border: "1px solid var(--border)", borderRadius: 6, padding: "5px 12px",
                        fontSize: 11, cursor: "pointer", fontWeight: 700,
                      }}>
                        {s === "all" ? "All" : s === "matched" ? "Matched" : s === "mismatch" ? "Diffs" : "Missing"}
                      </button>
                    ))}
                    <span style={{ fontSize: 11, color: "var(--muted)" }}>
                      {filteredRecords.length} results
                    </span>
                  </div>

                  {/* Table */}
                  <div style={{ maxHeight: 480, overflowY: "auto" }}>
                    <table>
                      <thead>
                        <tr>
                          <th>#</th>
                          <th>Status</th>
                          <th>Key</th>
                          {results.excelJoinCols.map(({ excelCol }) => <th key={excelCol}>{excelCol}</th>)}
                          <th>Diffs</th>
                          <th></th>
                        </tr>
                      </thead>
                      <tbody>
                        {(activeTab === "records" ? filteredRecords :
                          activeTab === "mismatch" ? results.records.filter(r => r.status === "mismatch") :
                          activeTab === "missing" ? results.records.filter(r => r.status === "missing_xml") :
                          []
                        ).map((r, i) => {
                          const isExp = !!expandedRows[r.key + i];
                          const diffs = r.fieldChecks.filter(c => !c.match);
                          return (
                            <>
                              <tr key={r.key + i} onClick={() => setExpandedRows(prev => ({ ...prev, [r.key + i]: !prev[r.key + i] }))}
                                style={{ cursor: "pointer", background: i % 2 === 0 ? "var(--surface)" : "var(--surface2)" }}>
                                <td style={{ color: "var(--muted)", fontSize: 11 }}>{String(i + 1).padStart(3, "0")}</td>
                                <td>
                                  <Badge type={r.status === "matched" ? "pass" : r.status === "mismatch" ? "warn" : "fail"}>
                                    {r.status === "matched" ? "✓ Matched" : r.status === "mismatch" ? "⚠ Diff" : "✗ Not Found"}
                                  </Badge>
                                </td>
                                <td style={{ fontWeight: 600 }}>{r.excelKeyDisplay}</td>
                                {results.excelJoinCols.map(({ excelCol }) => (
                                  <td key={excelCol} style={{ color: "var(--muted)", fontSize: 11, fontFamily: "monospace" }}>
                                    {r.excelRec[excelCol] || "—"}
                                  </td>
                                ))}
                                <td>
                                  {diffs.length > 0 && (
                                    <span style={{ color: "var(--fail)", fontWeight: 700, fontSize: 11 }}>
                                      {diffs.length} diff{diffs.length > 1 ? "s" : ""}
                                    </span>
                                  )}
                                </td>
                                <td style={{ color: "var(--muted)", fontSize: 12 }}>{isExp ? "▲" : "▼"}</td>
                              </tr>
                              {isExp && (
                                <tr key={r.key + i + "exp"}>
                                  <td colSpan={5 + results.excelJoinCols.length} style={{ padding: 0, background: "var(--surface2)" }}>
                                    <div style={{ padding: "10px 14px 14px 22px" }}>
                                      {r.status === "missing_xml" ? (
                                        <div style={{ fontSize: 11, color: "var(--fail)", padding: "8px 0" }}>
                                          ✗ This record exists in post-load Excel but was NOT found in the XML template.
                                          Key: <b>{r.excelKeyDisplay}</b>
                                        </div>
                                      ) : r.fieldChecks.length === 0 ? (
                                        <div style={{ fontSize: 11, color: "var(--muted)" }}>No field diffs.</div>
                                      ) : (
                                        <table style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 8, overflow: "hidden" }}>
                                          <thead>
                                            <tr>
                                              <th style={{ background: "var(--fail-bg)", color: "var(--fail)", border: "none" }}>Excel Column</th>
                                              <th style={{ background: "var(--fail-bg)", color: "var(--fail)", border: "none" }}>SAP Field</th>
                                              <th style={{ background: "var(--fail-bg)", color: "var(--fail)", border: "none" }}>Excel Value</th>
                                              <th style={{ background: "var(--fail-bg)", color: "var(--fail)", border: "none" }}>XML Value</th>
                                              <th style={{ background: "var(--fail-bg)", color: "var(--fail)", border: "none" }}>Match?</th>
                                            </tr>
                                          </thead>
                                          <tbody>
                                            {r.fieldChecks.map((c, ci) => (
                                              <tr key={ci} style={{ background: c.match ? "var(--surface)" : "rgba(220,38,38,.04)" }}>
                                                <td style={{ fontSize: 11, color: "var(--muted)" }}>{c.excelCol}</td>
                                                <td style={{ fontSize: 11, fontFamily: "monospace", color: "var(--accent)" }}>{c.xmlField}</td>
                                                <td style={{ fontSize: 11, fontFamily: "monospace", color: "var(--info)" }}>{c.excelVal || <span style={{ color: "var(--muted)" }}>—</span>}</td>
                                                <td style={{ fontSize: 11, fontFamily: "monospace", color: c.match ? "var(--pass)" : "var(--fail)", fontWeight: c.match ? 400 : 600 }}>{c.xmlVal || <span style={{ color: "var(--muted)" }}>—</span>}</td>
                                                <td><Badge type={c.match ? "pass" : "fail"}>{c.match ? "✓" : "✗"}</Badge></td>
                                              </tr>
                                            ))}
                                          </tbody>
                                        </table>
                                      )}
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </>
                          );
                        })}

                        {activeTab === "xmlonly" && results.xmlOnly.map((rec, i) => (
                          <tr key={i} style={{ background: i % 2 === 0 ? "var(--surface)" : "var(--surface2)" }}>
                            <td style={{ color: "var(--muted)", fontSize: 11 }}>{i + 1}</td>
                            <td><Badge type="info">XML Only</Badge></td>
                            <td colSpan={3 + results.excelJoinCols.length} style={{ fontSize: 12, color: "var(--text)", fontFamily: "monospace" }}>
                              {joinKeys.map(k => rec[k] || "").join(" · ")}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}
          </main>
        </div>
      </div>

      {/* ── Smart Search & Learn Modal ── */}
      {searchModal && (
        <div style={{
          position:"fixed",inset:0,background:"rgba(15,20,40,.45)",display:"flex",
          alignItems:"center",justifyContent:"center",zIndex:300,
        }} onClick={e => { if(e.target===e.currentTarget){setSearchModal(null);setSmartResults(null);setSmartQuery("");setConfirmedMatch(null);} }}>
          <div style={{
            background:"var(--surface)",border:"1px solid var(--border)",borderRadius:14,
            width:720,maxHeight:"87vh",display:"flex",flexDirection:"column",
            overflow:"hidden",boxShadow:"var(--shadow-md)",
          }}>
            {/* Modal Header */}
            <div style={{
              padding:"16px 20px",borderBottom:"1px solid var(--border)",
              display:"flex",alignItems:"center",justifyContent:"space-between",
              background:"var(--surface2)",flexShrink:0,
            }}>
              <div>
                <div style={{fontSize:14,fontWeight:700}}>🔍 Smart Key Search &amp; Learn</div>
                <div style={{fontSize:11,color:"var(--muted)",marginTop:2}}>
                  Missing key: <b style={{color:"var(--fail)"}}>{searchModal.excelKeyDisplay}</b>
                </div>
              </div>
              <button onClick={() => {setSearchModal(null);setSmartResults(null);setSmartQuery("");setConfirmedMatch(null);}}
                style={{background:"none",border:"none",color:"var(--muted)",fontSize:20,cursor:"pointer",borderRadius:5,lineHeight:1,padding:"0 3px"}}>×</button>
            </div>

            <div style={{overflowY:"auto",padding:20,flex:1}}>

              {/* Info banner */}
              <div style={{
                background:"var(--accent-light)",border:"1px solid rgba(79,70,229,.2)",
                borderRadius:9,padding:"11px 14px",marginBottom:16,fontSize:12,
                color:"var(--accent)",lineHeight:1.7,
              }}>
                <b>How it works:</b> Type any value from the record below (e.g. work center ID, name, plant).
                The tool will search <b>both files</b> for matching records. When you find the XML record that
                corresponds to this Excel row, click <b>"This is the match"</b> — the tool will learn
                which field to join on and update the mapping automatically.
              </div>

              {/* Excel record being searched */}
              <div style={{marginBottom:14}}>
                <div style={{fontSize:10,fontWeight:700,color:"var(--muted)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:6}}>
                  Excel Record (Not Found in XML)
                </div>
                <div style={{
                  background:"var(--fail-bg)",border:"1px solid rgba(220,38,38,.2)",
                  borderRadius:8,padding:"10px 14px",display:"flex",flexWrap:"wrap",gap:8,
                }}>
                  {Object.entries(searchModal.excelRec).filter(([,v])=>v).slice(0,12).map(([k,v]) => (
                    <span key={k} style={{
                      display:"inline-flex",flexDirection:"column",background:"var(--surface)",
                      border:"1px solid var(--border)",borderRadius:6,padding:"3px 9px",fontSize:11,
                    }}>
                      <small style={{fontSize:9,color:"var(--muted)"}}>{k}</small>
                      <b style={{color:"var(--text)"}}>{String(v).slice(0,30)}</b>
                    </span>
                  ))}
                </div>
              </div>

              {/* Search box */}
              <div style={{marginBottom:14}}>
                <div style={{fontSize:10,fontWeight:700,color:"var(--muted)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:6}}>
                  Search in both files
                </div>
                <input
                  autoFocus
                  value={smartQuery}
                  onChange={e => runSmartSearch(e.target.value)}
                  placeholder="Type a value to search across XML and Excel (e.g. 'BBURNI', 'USP1', 'BURN IN')"
                  style={{
                    width:"100%",background:"var(--surface2)",border:"1px solid var(--border)",
                    color:"var(--text)",padding:"9px 12px",borderRadius:8,fontSize:13,
                    outline:"none",
                  }}
                />
                {/* Quick-fill chips from the excel record */}
                <div style={{marginTop:6,display:"flex",flexWrap:"wrap",gap:5}}>
                  <span style={{fontSize:10,color:"var(--muted)",alignSelf:"center"}}>Quick search:</span>
                  {Object.entries(searchModal.excelRec).filter(([,v])=>v&&String(v).length<=30).slice(0,8).map(([k,v]) => (
                    <button key={k} onClick={() => runSmartSearch(String(v))} style={{
                      background:"var(--surface2)",border:"1px solid var(--border)",
                      borderRadius:5,padding:"2px 8px",fontSize:10,cursor:"pointer",color:"var(--text)",
                    }}>{String(v)}</button>
                  ))}
                </div>
              </div>

              {/* Confirmed match banner */}
              {confirmedMatch && (
                <div style={{
                  background:"var(--pass-bg)",border:"1px solid rgba(22,163,74,.3)",
                  borderRadius:9,padding:"11px 14px",marginBottom:14,fontSize:12,color:"var(--pass)",lineHeight:1.7,
                }}>
                  <b>✓ Mapping learned!</b> The tool now knows <code style={{background:"var(--surface2)",padding:"1px 5px",borderRadius:4,color:"var(--accent)"}}>{confirmedMatch.excelKeyField}</code> (Excel)
                  maps to <code style={{background:"var(--surface2)",padding:"1px 5px",borderRadius:4,color:"var(--accent)"}}>{confirmedMatch.xmlKeyField}</code> (XML).
                  <br/>Close this dialog and click <b>Run Validation</b> to see the updated results.
                </div>
              )}

              {/* Search results */}
              {smartResults && (
                <>
                  {/* XML hits */}
                  <div style={{marginBottom:16}}>
                    <div style={{fontSize:10,fontWeight:700,color:"var(--muted)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:8}}>
                      XML Template matches ({smartResults.xmlHits.length})
                    </div>
                    {smartResults.xmlHits.length === 0 ? (
                      <div style={{fontSize:12,color:"var(--muted)",padding:"8px 0"}}>No matches found in XML.</div>
                    ) : smartResults.xmlHits.map((hit, hi) => (
                      <div key={hi} style={{
                        background:"var(--surface2)",border:"1px solid var(--border)",
                        borderRadius:9,padding:"10px 14px",marginBottom:6,
                      }}>
                        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:10}}>
                          <div style={{flex:1}}>
                            <div style={{fontSize:10,color:"var(--muted)",marginBottom:5}}>
                              XML row {hit.rowIdx+1} · matched on:{" "}
                              {hit.matchingFields.map(f => (
                                <span key={f} style={{
                                  background:"var(--pass-bg)",color:"var(--pass)",fontSize:10,
                                  fontWeight:700,padding:"1px 6px",borderRadius:4,marginRight:3,
                                }}>{f}: <HL text={String(hit.rec[f])} query={smartQuery}/></span>
                              ))}
                            </div>
                            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                              {Object.entries(hit.rec).filter(([,v])=>v).slice(0,10).map(([k,v]) => (
                                <span key={k} style={{
                                  display:"inline-flex",flexDirection:"column",background:"var(--surface)",
                                  border:`1px solid ${hit.matchingFields.includes(k)?"rgba(22,163,74,.4)":"var(--border)"}`,
                                  borderRadius:6,padding:"3px 9px",fontSize:11,
                                }}>
                                  <small style={{fontSize:9,color:"var(--muted)"}}>{k}</small>
                                  <b style={{color:hit.matchingFields.includes(k)?"var(--pass)":"var(--text)"}}>
                                    <HL text={String(v).slice(0,25)} query={smartQuery}/>
                                  </b>
                                </span>
                              ))}
                            </div>
                          </div>
                          <div style={{flexShrink:0}}>
                            <button onClick={() => {
                              // Find which field in XML matches which field in Excel record
                              // Try to identify the join field pair
                              const xmlMatchField = hit.matchingFields[0];
                              // Find the Excel column whose value matches this XML field value
                              const xmlVal = (hit.rec[xmlMatchField]||"").toLowerCase().trim();
                              const excelMatchField = Object.keys(searchModal.excelRec).find(k =>
                                (searchModal.excelRec[k]||"").toLowerCase().trim() === xmlVal
                              ) || Object.keys(searchModal.excelRec)[0];
                              applyLearnedKey(searchModal.excelRec, hit.rec, excelMatchField, xmlMatchField);
                            }} style={{
                              background:"var(--pass)",color:"#fff",border:"none",
                              borderRadius:7,padding:"7px 14px",fontSize:12,fontWeight:600,
                              cursor:"pointer",whiteSpace:"nowrap",
                            }}>✓ This is the match</button>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Excel hits */}
                  {smartResults.excelHits.length > 0 && (
                    <div>
                      <div style={{fontSize:10,fontWeight:700,color:"var(--muted)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:8}}>
                        Excel Post-Load matches ({smartResults.excelHits.length})
                      </div>
                      {smartResults.excelHits.slice(0,4).map((hit, hi) => (
                        <div key={hi} style={{
                          background:"var(--info-bg)",border:"1px solid rgba(37,99,235,.2)",
                          borderRadius:9,padding:"10px 14px",marginBottom:6,fontSize:11,
                        }}>
                          <div style={{fontSize:10,color:"var(--info)",marginBottom:4}}>
                            Excel row {hit.rowIdx+1} · matched on:{" "}
                            {hit.matchingFields.map(f => (
                              <span key={f} style={{fontWeight:700,marginRight:4}}>{f}: <HL text={String(hit.rec[f])} query={smartQuery}/></span>
                            ))}
                          </div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                            {Object.entries(hit.rec).filter(([,v])=>v).slice(0,8).map(([k,v]) => (
                              <span key={k} style={{background:"var(--surface)",border:"1px solid var(--border)",borderRadius:5,padding:"2px 7px",fontSize:10}}>
                                <span style={{color:"var(--muted)"}}>{k}: </span>
                                <HL text={String(v).slice(0,20)} query={smartQuery}/>
                              </span>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  {smartResults.xmlHits.length === 0 && smartResults.excelHits.length === 0 && (
                    <div style={{
                      background:"var(--warn-bg)",border:"1px solid rgba(217,119,6,.25)",
                      borderRadius:9,padding:"12px 16px",fontSize:12,color:"var(--warn)",
                    }}>
                      ⚠ No matches found for "<b>{smartQuery}</b>" in either file.
                      Try a different value — perhaps a partial ID, description, or plant code.
                    </div>
                  )}
                </>
              )}

              {/* Learned mappings legend */}
              {learnedMappings.length > 0 && (
                <div style={{marginTop:16,borderTop:"1px solid var(--border)",paddingTop:14}}>
                  <div style={{fontSize:10,fontWeight:700,color:"var(--muted)",textTransform:"uppercase",letterSpacing:".06em",marginBottom:8}}>
                    Learned mappings this session ({learnedMappings.length})
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                    {learnedMappings.map((m, i) => (
                      <span key={i} style={{
                        background:"var(--teal-bg)",color:"var(--teal)",
                        border:"1px solid rgba(15,118,110,.25)",borderRadius:6,
                        padding:"3px 10px",fontSize:11,fontWeight:600,
                      }}>✓ {m.label}</span>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div style={{
          position: "fixed", bottom: 20, right: 20, zIndex: 999,
          background: "var(--surface)", border: "1px solid var(--border)", borderRadius: 10,
          padding: "11px 16px", fontSize: 12, maxWidth: 320, boxShadow: "var(--shadow-md)",
          display: "flex", alignItems: "flex-start", gap: 8, animation: "sli .22s ease",
          borderLeft: `3px solid ${toast.type === "success" ? "var(--pass)" : toast.type === "error" ? "var(--fail)" : "var(--info)"}`,
        }}>
          <span style={{ flex: 1, lineHeight: 1.5 }}>{toast.msg}</span>
        </div>
      )}
    </>
  );
}

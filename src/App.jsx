import { useState } from "react";
import * as XLSX from "xlsx";
import { EXCEL_FILES } from "./excelFiles";
import {
  BUCKET,
  supabase,
  SUPABASE_ANON_KEY,
  SUPABASE_URL,
} from "./supabaseClient";

// ─── Comparison Logic ────────────────────────────────────────────────────────

/**
 * Builds a lookup map from a resources array using `key` as the identifier.
 * @param {Array} resources
 * @returns {Object} map of key → resource
 */
function buildResourceMap(resources) {
  return resources.reduce((map, resource) => {
    map[resource.key] = resource;
    return map;
  }, {});
}

/**
 * Returns true if any of the tracked fields differ between two resources.
 * @param {Object} oldResource
 * @param {Object} newResource
 * @returns {boolean}
 */
function hasChanged(oldResource, newResource) {
  return (
    oldResource.value !== newResource.value ||
    oldResource.keyGroup !== newResource.keyGroup ||
    oldResource.isActive !== newResource.isActive
  );
}

/**
 * Compares two parsed JSON objects and returns two separate lists:
 *   - changed: resources whose key exists in both but has at least one modified field
 *   - added:   resources whose key is not present in the old JSON at all
 *
 * @param {Object} oldJson - Parsed old JSON
 * @param {Object} newJson - Parsed new JSON
 * @returns {{ changed: { data: { resources: Array } }, added: { data: { resources: Array } } }}
 */
function compareResources(oldJson, newJson) {
  const oldResources = oldJson?.data?.resources ?? [];
  const newResources = newJson?.data?.resources ?? [];

  const oldMap = buildResourceMap(oldResources);

  const changed = [];
  const added = [];

  for (const newResource of newResources) {
    const oldResource = oldMap[newResource.key];

    if (!oldResource) {
      // Case A: key does not exist in old → added resource
      added.push(newResource);
    } else if (hasChanged(oldResource, newResource)) {
      // Case B: key exists but fields changed → modified resource
      changed.push(newResource);
    }
    // Case C: nothing changed → skip
  }

  return {
    changed: { data: { resources: changed } },
    added: { data: { resources: added } },
  };
}

// ─── Table Logic ────────────────────────────────────────────────────────────

const COMPANY_PREFIXES = {
  AB: ["AB; ABWM; ABAPP"],
  BCP: ["FEP; FEPWM; FEPAPP"],
};

/**
 * Expands a list of resources into table rows.
 * Each resource produces one row per company prefix.
 * Row shape: [prefix, keyGroup, key, value]
 *
 * @param {Array} resources
 * @param {'AB'|'BCP'} company
 * @returns {Array<[string, string, string, string]>}
 */
function buildTableRows(resources, company) {
  const prefixes = COMPANY_PREFIXES[company] ?? [];
  return resources.flatMap((r) =>
    prefixes.map((prefix) => [
      prefix,
      r.keyGroup ?? "",
      r.key ?? "",
      r.value ?? "",
    ]),
  );
}

/**
 * Converts rows to TSV string (tab-separated values).
 * Pastes directly into Excel as separate columns.
 * @param {Array<string[]>} rows
 * @returns {string}
 */
function rowsToTsv(rows) {
  return rows.map((row) => row.join("\t")).join("\n");
}

/**
 * Returns the 0-based index of the next empty row in a worksheet.
 * Falls back to 0 if the sheet has no content yet.
 * @param {Object} ws - SheetJS worksheet
 * @returns {number}
 */
function getNextRow(ws) {
  const ref = ws["!ref"];
  if (!ref) return 0;
  const range = XLSX.utils.decode_range(ref);
  return range.e.r + 1; // one row below the last occupied row
}

/**
 * Fetches an Excel file from Supabase Storage, appends changedRows to the
 * "Modified" sheet and addedRows to the "Added" sheet (PT) or
 * "Modified EN" / "Added EN" (EN), then re-uploads the file in place.
 *
 * @param {string}   fileName    - Filename in the bucket (e.g. "BCP_Cards.xlsx")
 * @param {Array}    changedRows
 * @param {Array}    addedRows
 * @param {'PT'|'EN'} language
 * @returns {Promise<{ error: string|null }>}
 */
async function appendRowsToSupabase(
  fileName,
  changedRows,
  addedRows,
  language,
) {
  // 1. Download existing file using the authenticated storage endpoint.
  const authDownloadUrl = `${SUPABASE_URL}/storage/v1/object/authenticated/${BUCKET}/${fileName}`;
  console.log("[Export] Downloading from:", authDownloadUrl);

  const response = await fetch(authDownloadUrl, {
    headers: {
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
      apikey: SUPABASE_ANON_KEY,
      "Cache-Control": "no-cache, no-store",
      Pragma: "no-cache",
    },
    cache: "no-store",
  });

  console.log(
    "[Export] Download status:",
    response.status,
    response.statusText,
  );

  if (!response.ok) {
    const body = await response.text();
    console.error("[Export] Download error body:", body);
    return {
      error: `Download failed: ${response.status} ${response.statusText} — ${body}`,
    };
  }

  // 2. Read the workbook
  const buffer = await response.arrayBuffer();
  console.log("[Export] Downloaded buffer size (bytes):", buffer.byteLength);

  const wb = XLSX.read(buffer, { type: "buffer" });
  console.log("[Export] Sheets found in workbook:", wb.SheetNames);

  // PT → sheets 1 & 2 | EN → sheets 3 & 4
  const modifiedSheetName = language === "EN" ? "Modified EN" : "Modified";
  const addedSheetName = language === "EN" ? "Added EN" : "Added";

  function appendToSheet(sheetName, rows) {
    if (!rows.length) return;
    if (!wb.SheetNames.includes(sheetName)) {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    } else {
      const ws = wb.Sheets[sheetName];
      const startRow = getNextRow(ws);
      XLSX.utils.sheet_add_aoa(ws, rows, { origin: { r: startRow, c: 0 } });
    }
  }

  appendToSheet(modifiedSheetName, changedRows);
  appendToSheet(addedSheetName, addedRows);

  // 3. Serialize and re-upload (upsert = overwrite in place)
  const data = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  console.log("[Export] Serialized workbook size (bytes):", data.byteLength);

  const uploadBlob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const { error: uploadError } = await supabase.storage
    .from(BUCKET)
    .upload(fileName, uploadBlob, {
      upsert: true,
      contentType: uploadBlob.type,
    });

  if (uploadError) {
    console.error("[Export] Upload error:", uploadError);
    return { error: `Upload failed: ${uploadError.message}` };
  }

  console.log("[Export] Upload successful.");
  return { error: null };
}

// ─── UI Logic ────────────────────────────────────────────────────────────────

const styles = {
  container: {
    maxWidth: "960px",
    margin: "0 auto",
    padding: "32px 16px",
    fontFamily: "system-ui, sans-serif",
    color: "#1a1a1a",
  },
  heading: {
    fontSize: "22px",
    fontWeight: "700",
    marginBottom: "8px",
  },
  subheading: {
    fontSize: "13px",
    color: "#666",
    marginBottom: "32px",
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "24px",
    marginBottom: "24px",
  },
  label: {
    display: "block",
    fontSize: "13px",
    fontWeight: "600",
    marginBottom: "8px",
    color: "#444",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  textarea: {
    width: "100%",
    minHeight: "240px",
    padding: "12px",
    fontFamily: "monospace",
    fontSize: "13px",
    border: "1px solid #d1d5db",
    borderRadius: "6px",
    resize: "vertical",
    boxSizing: "border-box",
    backgroundColor: "#fafafa",
    lineHeight: "1.5",
    outline: "none",
  },
  textareaError: {
    borderColor: "#ef4444",
    backgroundColor: "#fff5f5",
  },
  textareaReadonly: {
    backgroundColor: "#f0fdf4",
    borderColor: "#86efac",
    cursor: "default",
  },
  errorText: {
    fontSize: "12px",
    color: "#ef4444",
    marginTop: "6px",
  },
  actions: {
    display: "flex",
    gap: "12px",
    alignItems: "center",
    marginBottom: "24px",
  },
  button: {
    padding: "10px 24px",
    fontSize: "14px",
    fontWeight: "600",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    backgroundColor: "#2563eb",
    color: "#fff",
    transition: "background 0.15s",
  },
  buttonDisabled: {
    backgroundColor: "#93c5fd",
    cursor: "not-allowed",
  },
  buttonSecondary: {
    padding: "10px 20px",
    fontSize: "14px",
    fontWeight: "600",
    border: "1px solid #d1d5db",
    borderRadius: "6px",
    cursor: "pointer",
    backgroundColor: "#fff",
    color: "#374151",
    transition: "background 0.15s",
  },
  copySuccess: {
    fontSize: "13px",
    color: "#16a34a",
  },
  resultSection: {
    marginTop: "8px",
  },
  resultHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  badge: {
    fontSize: "12px",
    fontWeight: "600",
    padding: "2px 10px",
    borderRadius: "99px",
    backgroundColor: "#dcfce7",
    color: "#166534",
  },
  select: {
    padding: "10px 12px",
    fontSize: "14px",
    fontWeight: "600",
    border: "1px solid #d1d5db",
    borderRadius: "6px",
    backgroundColor: "#fff",
    color: "#1a1a1a",
    cursor: "pointer",
    outline: "none",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: "13px",
    fontFamily: "monospace",
  },
  th: {
    textAlign: "left",
    padding: "8px 12px",
    backgroundColor: "#f1f5f9",
    borderBottom: "2px solid #e2e8f0",
    fontSize: "12px",
    fontWeight: "700",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    color: "#475569",
    fontFamily: "system-ui, sans-serif",
  },
  td: {
    padding: "7px 12px",
    borderBottom: "1px solid #e2e8f0",
    verticalAlign: "top",
    color: "#1a1a1a",
  },
  trEven: {
    backgroundColor: "#f8fafc",
  },
  tabBar: {
    display: "flex",
    gap: "0",
    borderBottom: "2px solid #e2e8f0",
    marginBottom: "32px",
  },
  tab: {
    padding: "10px 24px",
    fontSize: "14px",
    fontWeight: "600",
    border: "none",
    borderBottom: "2px solid transparent",
    marginBottom: "-2px",
    background: "none",
    cursor: "pointer",
    color: "#64748b",
    transition: "color 0.15s, border-color 0.15s",
  },
  tabActive: {
    color: "#2563eb",
    borderBottomColor: "#2563eb",
  },
  modalOverlay: {
    position: "fixed",
    inset: 0,
    backgroundColor: "rgba(0,0,0,0.4)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 1000,
  },
  modalBox: {
    backgroundColor: "#fff",
    borderRadius: "10px",
    padding: "28px 32px",
    maxWidth: "400px",
    width: "90%",
    boxShadow: "0 8px 32px rgba(0,0,0,0.18)",
  },
  modalTitle: {
    fontSize: "16px",
    fontWeight: "700",
    marginBottom: "8px",
    color: "#1a1a1a",
  },
  modalBody: {
    fontSize: "14px",
    color: "#64748b",
    marginBottom: "24px",
    lineHeight: "1.5",
  },
  modalActions: {
    display: "flex",
    gap: "10px",
    justifyContent: "flex-end",
  },
  buttonDanger: {
    padding: "9px 20px",
    fontSize: "14px",
    fontWeight: "600",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    backgroundColor: "#dc2626",
    color: "#fff",
  },
};

function ConfirmModal({ fileName, onConfirm, onCancel }) {
  return (
    <div style={styles.modalOverlay}>
      <div style={styles.modalBox}>
        <p style={styles.modalTitle}>Remover todo o conteúdo?</p>
        <p style={styles.modalBody}>
          Todas as folhas de <strong>{fileName}.xlsx</strong> serão esvaziadas.
          Esta ação não pode ser revertida.
        </p>
        <div style={styles.modalActions}>
          <button style={styles.buttonSecondary} onClick={onCancel}>
            Cancelar
          </button>
          <button style={styles.buttonDanger} onClick={onConfirm}>
            Sim, remover
          </button>
        </div>
      </div>
    </div>
  );
}

function FilesScreen() {
  const [downloading, setDownloading] = useState({}); // { [name]: bool }
  const [resetting, setResetting] = useState({}); // { [name]: bool }
  const [confirmReset, setConfirmReset] = useState(null); // name | null
  const [errors, setErrors] = useState({}); // { [name]: string }

  async function handleDownloadFile(name) {
    const fileName = `${name}.xlsx`;
    setDownloading((d) => ({ ...d, [name]: true }));
    setErrors((e) => ({ ...e, [name]: null }));

    const authDownloadUrl = `${SUPABASE_URL}/storage/v1/object/authenticated/${BUCKET}/${fileName}`;
    const response = await fetch(authDownloadUrl, {
      headers: {
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        apikey: SUPABASE_ANON_KEY,
      },
      cache: "no-store",
    });

    setDownloading((d) => ({ ...d, [name]: false }));

    if (!response.ok) {
      setErrors((e) => ({
        ...e,
        [name]: `Error ${response.status}: ${response.statusText}`,
      }));
      return;
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function handleResetFile(name) {
    const fileName = `${name}.xlsx`;
    setResetting((r) => ({ ...r, [name]: true }));
    setErrors((e) => ({ ...e, [name]: null }));

    const authDownloadUrl = `${SUPABASE_URL}/storage/v1/object/authenticated/${BUCKET}/${fileName}`;
    const response = await fetch(authDownloadUrl, {
      headers: {
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        apikey: SUPABASE_ANON_KEY,
      },
      cache: "no-store",
    });

    if (!response.ok) {
      setErrors((e) => ({
        ...e,
        [name]: `Reset failed: ${response.status} ${response.statusText}`,
      }));
      setResetting((r) => ({ ...r, [name]: false }));
      return;
    }

    const buffer = await response.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "buffer" });

    // Clear only the diff sheets — all other sheets (e.g. "Capa") are left untouched
    const DIFF_SHEETS = ["Modified", "Added", "Modified EN", "Added EN"];
    for (const sheetName of wb.SheetNames) {
      if (DIFF_SHEETS.includes(sheetName)) {
        wb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([]);
      }
    }

    const data = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const uploadBlob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const { error: uploadError } = await supabase.storage
      .from(BUCKET)
      .upload(fileName, uploadBlob, {
        upsert: true,
        contentType: uploadBlob.type,
      });

    setResetting((r) => ({ ...r, [name]: false }));

    if (uploadError) {
      setErrors((e) => ({
        ...e,
        [name]: `Reset failed: ${uploadError.message}`,
      }));
    }
  }

  return (
    <div>
      {confirmReset && (
        <ConfirmModal
          fileName={confirmReset}
          onCancel={() => setConfirmReset(null)}
          onConfirm={() => {
            const name = confirmReset;
            setConfirmReset(null);
            handleResetFile(name);
          }}
        />
      )}
      <p style={{ ...styles.subheading, marginBottom: "24px" }}>
        Download the latest version of each Excel file directly from Supabase
        Storage.
      </p>
      <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
        {EXCEL_FILES.map((name) => (
          <div
            key={name}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "14px 20px",
              border: "1px solid #e2e8f0",
              borderRadius: "8px",
              backgroundColor: "#f8fafc",
            }}
          >
            <span
              style={{
                fontSize: "14px",
                fontWeight: "600",
                fontFamily: "monospace",
              }}
            >
              {name}.xlsx
            </span>
            <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
              {errors[name] && (
                <span style={{ fontSize: "12px", color: "#dc2626" }}>
                  {errors[name]}
                </span>
              )}
              <button
                style={{
                  ...styles.buttonSecondary,
                  ...(resetting[name] ? styles.buttonDisabled : {}),
                }}
                onClick={() => setConfirmReset(name)}
                disabled={!!resetting[name] || !!downloading[name]}
              >
                {resetting[name] ? "Resetting…" : "Reset"}
              </button>
              <button
                style={{
                  ...styles.buttonSecondary,
                  ...(downloading[name] ? styles.buttonDisabled : {}),
                }}
                onClick={() => handleDownloadFile(name)}
                disabled={!!downloading[name] || !!resetting[name]}
              >
                {downloading[name] ? "Downloading…" : "Download latest"}
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function UrlToolScreen() {
  const [input, setInput] = useState("");
  const [output, setOutput] = useState("");
  const [error, setError] = useState("");
  const [mode, setMode] = useState("decode"); // 'decode' | 'encode'
  const [copyDone, setCopyDone] = useState(false);

  function handleConvert() {
    setError("");
    setOutput("");
    const raw = input.trim();
    if (!raw) return;

    if (mode === "decode") {
      try {
        const decoded = decodeURIComponent(raw);
        try {
          const parsed = JSON.parse(decoded);
          setOutput(JSON.stringify(parsed, null, 2));
        } catch {
          // Not JSON — just show the decoded string as-is
          setOutput(decoded);
        }
      } catch {
        setError("Invalid URL-encoded string.");
      }
    } else {
      // encode mode: accept raw string or pretty JSON
      try {
        // Try to re-minify if it's valid JSON
        const minified = JSON.stringify(JSON.parse(raw));
        setOutput(encodeURIComponent(minified));
      } catch {
        // Not JSON — encode as-is
        setOutput(encodeURIComponent(raw));
      }
    }
  }

  function handleCopy() {
    if (!output) return;
    navigator.clipboard.writeText(output).then(() => {
      setCopyDone(true);
      setTimeout(() => setCopyDone(false), 2000);
    });
  }

  function handleClear() {
    setInput("");
    setOutput("");
    setError("");
  }

  return (
    <div>
      <p style={{ ...styles.subheading, marginBottom: "24px" }}>
        Decode a URL-encoded string (e.g. query params) into readable JSON, or
        encode JSON back.
      </p>

      {/* Mode toggle */}
      <div style={{ display: "flex", gap: "8px", marginBottom: "16px" }}>
        {["decode", "encode"].map((m) => (
          <button
            key={m}
            style={{
              ...styles.buttonSecondary,
              ...(mode === m
                ? {
                    backgroundColor: "#2563eb",
                    color: "#fff",
                    borderColor: "#2563eb",
                  }
                : {}),
            }}
            onClick={() => {
              setMode(m);
              setOutput("");
              setError("");
            }}
          >
            {m === "decode" ? "Decode" : "Encode"}
          </button>
        ))}
      </div>

      {/* Input */}
      <div style={{ marginBottom: "12px" }}>
        <label style={styles.label}>
          {mode === "decode" ? "URL-encoded input" : "JSON / plain text input"}
        </label>
        <textarea
          style={{
            ...styles.textarea,
            minHeight: "160px",
            ...(error ? styles.textareaError : {}),
          }}
          value={input}
          onChange={(e) => {
            setInput(e.target.value);
            setError("");
          }}
          placeholder={
            mode === "decode"
              ? "%7B%22key%22%3A%22value%22%7D"
              : '{"key": "value"}'
          }
          spellCheck={false}
        />
        {error && <p style={styles.errorText}>{error}</p>}
      </div>

      {/* Actions */}
      <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
        <button style={styles.button} onClick={handleConvert}>
          {mode === "decode" ? "Decode" : "Encode"}
        </button>
        <button style={styles.buttonSecondary} onClick={handleClear}>
          Clear
        </button>
      </div>

      {/* Output */}
      {output && (
        <div>
          <div style={{ ...styles.resultHeader, marginBottom: "8px" }}>
            <label style={styles.label}>Output</label>
            <button style={styles.buttonSecondary} onClick={handleCopy}>
              {copyDone ? "Copied!" : "Copy"}
            </button>
          </div>
          <pre
            style={{
              backgroundColor: "#f8fafc",
              border: "1px solid #e2e8f0",
              borderRadius: "6px",
              padding: "16px",
              fontSize: "13px",
              fontFamily: "monospace",
              overflowX: "auto",
              whiteSpace: "pre-wrap",
              wordBreak: "break-all",
              margin: 0,
              color: "#1a1a1a",
            }}
          >
            {output}
          </pre>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [activeTab, setActiveTab] = useState("comparator");
  const [oldInput, setOldInput] = useState("");
  const [newInput, setNewInput] = useState("");
  const [changedTableRows, setChangedTableRows] = useState([]);
  const [addedTableRows, setAddedTableRows] = useState([]);
  const [company, setCompany] = useState("BCP");
  const [language, setLanguage] = useState("PT");
  const [error, setError] = useState("");
  const [oldError, setOldError] = useState("");
  const [newError, setNewError] = useState("");
  const [copySuccess, setCopySuccess] = useState(null); // 'changed' | 'added' | 'table-changed' | 'table-added' | null
  const [changedCount, setChangedCount] = useState(null);
  const [addedCount, setAddedCount] = useState(null);
  const [selectedFile, setSelectedFile] = useState("");
  const [isUploading, setIsUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState(null); // { ok: bool, message: string } | null

  const isCompareDisabled = newInput.trim() === "";
  const isExportDisabled =
    !selectedFile ||
    isUploading ||
    (changedTableRows.length === 0 && addedTableRows.length === 0);

  function validateJson(raw, setFieldError) {
    try {
      const parsed = JSON.parse(raw);
      setFieldError("");
      return parsed;
    } catch {
      setFieldError("Invalid JSON");
      return null;
    }
  }

  function handleCompare() {
    setError("");
    setChangedTableRows([]);
    setAddedTableRows([]);
    setChangedCount(null);
    setAddedCount(null);
    setCopySuccess(null);

    // Old JSON is optional — if empty, all new resources are treated as added
    let oldJson = { data: { resources: [] } };
    if (oldInput.trim() !== "") {
      const parsed = validateJson(oldInput, setOldError);
      if (!parsed) {
        setError("Invalid JSON input");
        return;
      }
      if (!Array.isArray(parsed?.data?.resources)) {
        setError('Expected structure: { "data": { "resources": [...] } }');
        return;
      }
      oldJson = parsed;
    } else {
      setOldError("");
    }

    const newJson = validateJson(newInput, setNewError);
    if (!newJson) {
      setError("Invalid JSON input");
      return;
    }
    if (!Array.isArray(newJson?.data?.resources)) {
      setError('Expected structure: { "data": { "resources": [...] } }');
      return;
    }

    const { changed, added } = compareResources(oldJson, newJson);
    setChangedCount(changed.data.resources.length);
    setAddedCount(added.data.resources.length);

    setChangedTableRows(buildTableRows(changed.data.resources, company));
    setAddedTableRows(buildTableRows(added.data.resources, company));
  }

  function handleCopyTable(type) {
    const rows = type === "table-changed" ? changedTableRows : addedTableRows;
    if (!rows.length) return;
    navigator.clipboard.writeText(rowsToTsv(rows)).then(() => {
      setCopySuccess(type);
      setTimeout(() => setCopySuccess(null), 2000);
    });
  }

  async function handleDownload() {
    if (!selectedFile) return;
    const fileName = `${selectedFile}.xlsx`;
    const authDownloadUrl = `${SUPABASE_URL}/storage/v1/object/authenticated/${BUCKET}/${fileName}`;
    const response = await fetch(authDownloadUrl, {
      headers: {
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        apikey: SUPABASE_ANON_KEY,
      },
      cache: "no-store",
    });
    if (!response.ok) {
      setUploadStatus({
        ok: false,
        message: `Download failed: ${response.status} ${response.statusText}`,
      });
      return;
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function handleExportToExcel() {
    if (!selectedFile) return;
    setIsUploading(true);
    setUploadStatus(null);

    const { error } = await appendRowsToSupabase(
      `${selectedFile}.xlsx`,
      changedTableRows,
      addedTableRows,
      language,
    );

    setIsUploading(false);
    if (error) {
      setUploadStatus({ ok: false, message: error });
    } else {
      setUploadStatus({
        ok: true,
        message: `${selectedFile}.xlsx updated successfully in Supabase.`,
      });
    }
  }

  function handleOldChange(e) {
    setOldInput(e.target.value);
    setOldError("");
    setError("");
  }

  function handleNewChange(e) {
    setNewInput(e.target.value);
    setNewError("");
    setError("");
  }

  return (
    <div style={styles.container}>
      <h1 style={styles.heading}>JSON Resource Comparator</h1>

      {/* Tab bar */}
      <div style={styles.tabBar}>
        <button
          style={{
            ...styles.tab,
            ...(activeTab === "comparator" ? styles.tabActive : {}),
          }}
          onClick={() => setActiveTab("comparator")}
        >
          Comparator
        </button>
        <button
          style={{
            ...styles.tab,
            ...(activeTab === "files" ? styles.tabActive : {}),
          }}
          onClick={() => setActiveTab("files")}
        >
          Files
        </button>
        <button
          style={{
            ...styles.tab,
            ...(activeTab === "url" ? styles.tabActive : {}),
          }}
          onClick={() => setActiveTab("url")}
        >
          URL Tool
        </button>
      </div>

      {activeTab === "files" && <FilesScreen />}
      {activeTab === "url" && <UrlToolScreen />}

      {activeTab === "comparator" && (
        <>
          <p style={styles.subheading}>
            Paste two resource JSONs to generate a diff containing only added or
            modified resources.
          </p>

          {/* Input textareas */}
          <div style={styles.grid}>
            <div>
              <label style={styles.label}>
                Old JSON{" "}
                <span
                  style={{
                    fontWeight: 400,
                    textTransform: "none",
                    color: "#999",
                  }}
                >
                  (optional)
                </span>
              </label>
              <textarea
                style={{
                  ...styles.textarea,
                  ...(oldError ? styles.textareaError : {}),
                }}
                placeholder={'{\n  "data": {\n    "resources": [...]\n  }\n}'}
                value={oldInput}
                onChange={handleOldChange}
                spellCheck={false}
              />
              {oldError && <p style={styles.errorText}>{oldError}</p>}
            </div>

            <div>
              <label style={styles.label}>New JSON</label>
              <textarea
                style={{
                  ...styles.textarea,
                  ...(newError ? styles.textareaError : {}),
                }}
                placeholder={'{\n  "data": {\n    "resources": [...]\n  }\n}'}
                value={newInput}
                onChange={handleNewChange}
                spellCheck={false}
              />
              {newError && <p style={styles.errorText}>{newError}</p>}
            </div>
          </div>

          {/* Actions */}
          <div style={styles.actions}>
            <select
              style={styles.select}
              value={company}
              onChange={(e) => setCompany(e.target.value)}
            >
              <option value="BCP">BCP</option>
              <option value="AB">AB</option>
            </select>

            <select
              style={styles.select}
              value={language}
              onChange={(e) => setLanguage(e.target.value)}
            >
              <option value="PT">PT</option>
              <option value="EN">EN</option>
            </select>

            <button
              style={{
                ...styles.button,
                ...(isCompareDisabled ? styles.buttonDisabled : {}),
              }}
              onClick={handleCompare}
              disabled={isCompareDisabled}
            >
              Compare JSONs
            </button>
          </div>

          {/* Global error */}
          {error && (
            <p
              style={{
                ...styles.errorText,
                fontSize: "14px",
                marginBottom: "16px",
              }}
            >
              {error}
            </p>
          )}

          {/* Excel Export — select a Supabase file and append rows */}
          {(changedTableRows.length > 0 || addedTableRows.length > 0) && (
            <div
              style={{
                marginTop: "32px",
                padding: "16px 20px",
                backgroundColor: "#f8fafc",
                border: "1px solid #e2e8f0",
                borderRadius: "8px",
              }}
            >
              <label style={{ ...styles.label, marginBottom: "12px" }}>
                Export to Supabase Excel
              </label>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "12px",
                  flexWrap: "wrap",
                }}
              >
                <select
                  style={{ ...styles.select, minWidth: "220px" }}
                  value={selectedFile}
                  onChange={(e) => {
                    setSelectedFile(e.target.value);
                    setUploadStatus(null);
                  }}
                >
                  <option value="">— Select a file —</option>
                  {EXCEL_FILES.map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>

                <button
                  style={{
                    ...styles.button,
                    ...(isExportDisabled ? styles.buttonDisabled : {}),
                  }}
                  onClick={handleExportToExcel}
                  disabled={isExportDisabled}
                >
                  {isUploading ? "Uploading…" : "Export to Excel"}
                </button>

                <button
                  style={{
                    ...styles.buttonSecondary,
                    ...(!selectedFile ? styles.buttonDisabled : {}),
                  }}
                  onClick={handleDownload}
                  disabled={!selectedFile}
                >
                  Download latest
                </button>
              </div>

              {selectedFile && (
                <p
                  style={{
                    fontSize: "12px",
                    color: "#888",
                    marginTop: "8px",
                    marginBottom: 0,
                  }}
                >
                  {language === "PT" ? (
                    <>
                      Rows will be appended to sheets <strong>Modified</strong>{" "}
                      and <strong>Added</strong> in{" "}
                      <strong>{selectedFile}.xlsx</strong>.
                    </>
                  ) : (
                    <>
                      Rows will be appended to sheets{" "}
                      <strong>Modified EN</strong> and <strong>Added EN</strong>{" "}
                      in <strong>{selectedFile}.xlsx</strong>.
                    </>
                  )}{" "}
                  Sheets will be created if they don't exist yet.
                </p>
              )}

              {uploadStatus && (
                <p
                  style={{
                    fontSize: "13px",
                    marginTop: "10px",
                    marginBottom: 0,
                    fontWeight: "600",
                    color: uploadStatus.ok ? "#16a34a" : "#dc2626",
                  }}
                >
                  {uploadStatus.message}
                </p>
              )}
            </div>
          )}

          {/* Excel Tables */}
          {(changedTableRows.length > 0 || addedTableRows.length > 0) && (
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                gap: "32px",
                marginTop: "32px",
              }}
            >
              {/* Modified table */}
              <div>
                <div style={{ ...styles.resultHeader, marginBottom: "12px" }}>
                  <label style={styles.label}>
                    Modified — Excel ({company})
                  </label>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "8px",
                    }}
                  >
                    <span style={{ fontSize: "12px", color: "#666" }}>
                      {changedTableRows.length}{" "}
                      {changedTableRows.length === 1 ? "row" : "rows"}
                    </span>
                    <button
                      style={styles.buttonSecondary}
                      onClick={() => handleCopyTable("table-changed")}
                    >
                      {copySuccess === "table-changed"
                        ? "Copied!"
                        : "Copy for Excel"}
                    </button>
                  </div>
                </div>
                <div
                  style={{
                    overflowX: "auto",
                    border: "1px solid #e2e8f0",
                    borderRadius: "6px",
                  }}
                >
                  <table style={styles.table}>
                    <thead>
                      <tr>
                        <th style={styles.th}>Company</th>
                        <th style={styles.th}>keyGroup</th>
                        <th style={styles.th}>key</th>
                        <th style={styles.th}>value</th>
                      </tr>
                    </thead>
                    <tbody>
                      {changedTableRows.map(
                        ([prefix, keyGroup, key, value], i) => (
                          <tr key={i} style={i % 2 === 1 ? styles.trEven : {}}>
                            <td style={styles.td}>{prefix}</td>
                            <td style={styles.td}>{keyGroup}</td>
                            <td style={styles.td}>{key}</td>
                            <td style={styles.td}>{value}</td>
                          </tr>
                        ),
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Added table */}
              <div>
                <div style={{ ...styles.resultHeader, marginBottom: "12px" }}>
                  <label style={styles.label}>Added — Excel ({company})</label>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "8px",
                    }}
                  >
                    <span style={{ fontSize: "12px", color: "#666" }}>
                      {addedTableRows.length}{" "}
                      {addedTableRows.length === 1 ? "row" : "rows"}
                    </span>
                    <button
                      style={styles.buttonSecondary}
                      onClick={() => handleCopyTable("table-added")}
                    >
                      {copySuccess === "table-added"
                        ? "Copied!"
                        : "Copy for Excel"}
                    </button>
                  </div>
                </div>
                <div
                  style={{
                    overflowX: "auto",
                    border: "1px solid #e2e8f0",
                    borderRadius: "6px",
                  }}
                >
                  <table style={styles.table}>
                    <thead>
                      <tr>
                        <th style={styles.th}>Company</th>
                        <th style={styles.th}>keyGroup</th>
                        <th style={styles.th}>key</th>
                        <th style={styles.th}>value</th>
                      </tr>
                    </thead>
                    <tbody>
                      {addedTableRows.map(
                        ([prefix, keyGroup, key, value], i) => (
                          <tr key={i} style={i % 2 === 1 ? styles.trEven : {}}>
                            <td style={styles.td}>{prefix}</td>
                            <td style={styles.td}>{keyGroup}</td>
                            <td style={styles.td}>{key}</td>
                            <td style={styles.td}>{value}</td>
                          </tr>
                        ),
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

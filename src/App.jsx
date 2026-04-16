import { useState } from "react";

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

// ─── UI Logic ────────────────────────────────────────────────────────────────

const styles = {
  container: {
    maxWidth: "960px",
    margin: "0 auto",
    padding: "32px 24px",
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
};

export default function App() {
  const [oldInput, setOldInput] = useState("");
  const [newInput, setNewInput] = useState("");
  const [changedResult, setChangedResult] = useState("");
  const [addedResult, setAddedResult] = useState("");
  const [error, setError] = useState("");
  const [oldError, setOldError] = useState("");
  const [newError, setNewError] = useState("");
  const [copySuccess, setCopySuccess] = useState(null); // 'changed' | 'added' | null
  const [changedCount, setChangedCount] = useState(null);
  const [addedCount, setAddedCount] = useState(null);

  const isCompareDisabled = oldInput.trim() === "" || newInput.trim() === "";

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
    setChangedResult("");
    setAddedResult("");
    setChangedCount(null);
    setAddedCount(null);
    setCopySuccess(null);

    const oldJson = validateJson(oldInput, setOldError);
    const newJson = validateJson(newInput, setNewError);

    if (!oldJson || !newJson) {
      setError("Invalid JSON input");
      return;
    }

    // Validate expected structure
    if (
      !Array.isArray(oldJson?.data?.resources) ||
      !Array.isArray(newJson?.data?.resources)
    ) {
      setError('Expected structure: { "data": { "resources": [...] } }');
      return;
    }

    const { changed, added } = compareResources(oldJson, newJson);
    setChangedCount(changed.data.resources.length);
    setAddedCount(added.data.resources.length);
    setChangedResult(JSON.stringify(changed, null, 2));
    setAddedResult(JSON.stringify(added, null, 2));
  }

  function handleCopy(type) {
    const text = type === "changed" ? changedResult : addedResult;
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      setCopySuccess(type);
      setTimeout(() => setCopySuccess(null), 2000);
    });
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
      <p style={styles.subheading}>
        Paste two resource JSONs to generate a diff containing only added or
        modified resources.
      </p>

      {/* Input textareas */}
      <div style={styles.grid}>
        <div>
          <label style={styles.label}>Old JSON</label>
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

      {/* Results */}
      {(changedResult || addedResult) && (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: "24px",
          }}
        >
          {/* Modified resources */}
          <div style={styles.resultSection}>
            <div style={styles.resultHeader}>
              <label style={styles.label}>Modified Resources</label>
              <div
                style={{ display: "flex", alignItems: "center", gap: "8px" }}
              >
                {changedCount !== null && (
                  <span style={styles.badge}>
                    {changedCount}{" "}
                    {changedCount === 1 ? "resource" : "resources"}
                  </span>
                )}
                <button
                  style={styles.buttonSecondary}
                  onClick={() => handleCopy("changed")}
                >
                  {copySuccess === "changed" ? "Copied!" : "Copy"}
                </button>
              </div>
            </div>
            <textarea
              style={{ ...styles.textarea, ...styles.textareaReadonly }}
              value={changedResult}
              readOnly
              spellCheck={false}
            />
          </div>

          {/* Added resources */}
          <div style={styles.resultSection}>
            <div style={styles.resultHeader}>
              <label style={styles.label}>Added Resources</label>
              <div
                style={{ display: "flex", alignItems: "center", gap: "8px" }}
              >
                {addedCount !== null && (
                  <span
                    style={{
                      ...styles.badge,
                      backgroundColor: "#dbeafe",
                      color: "#1e40af",
                    }}
                  >
                    {addedCount} {addedCount === 1 ? "resource" : "resources"}
                  </span>
                )}
                <button
                  style={styles.buttonSecondary}
                  onClick={() => handleCopy("added")}
                >
                  {copySuccess === "added" ? "Copied!" : "Copy"}
                </button>
              </div>
            </div>
            <textarea
              style={{ ...styles.textarea, ...styles.textareaReadonly }}
              value={addedResult}
              readOnly
              spellCheck={false}
            />
          </div>
        </div>
      )}
    </div>
  );
}

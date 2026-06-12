(function () {
  const baseUrl = window.location.origin || "https://localhost:3000";
  const cacheStorageKey = "localact.functionCache.v1";
  const logList = document.getElementById("logList");
  const serviceStatus = document.getElementById("serviceStatus");
  const footerStatus = document.getElementById("footerStatus");

  const uploadConfigs = [
    {
      buttonId: "updatePl",
      endpoint: "/api/import/pl",
      label: "P&L",
      files: [
        ["plData", "plData"],
        ["plMapping", "plMapping"]
      ]
    },
    {
      buttonId: "updatePlDepartment",
      endpoint: "/api/import/pl-department",
      label: "P&L by Departments",
      files: [
        ["plDepartmentData", "plDepartmentData"],
        ["departmentMapping", "departmentMapping"]
      ]
    },
    {
      buttonId: "updateBs",
      endpoint: "/api/import/bs",
      label: "Balance Sheet",
      files: [
        ["bsData", "bsData"],
        ["bsMapping", "bsMapping"]
      ]
    },
    {
      buttonId: "updateOps",
      endpoint: "/api/import/ops",
      label: "Operations",
      files: [["ops", "ops"]]
    }
  ];

  Office.onReady(() => {
    wireUploads();
    document.getElementById("convertFormulas").addEventListener("click", convertFormulasToValues);
    document.getElementById("clearLog").addEventListener("click", () => {
      logList.innerHTML = "";
    });
    checkHealth();
    refreshFunctionCache(false);
    setInterval(checkHealth, 15000);
  });

  function wireUploads() {
    for (const config of uploadConfigs) {
      document.getElementById(config.buttonId).addEventListener("click", () => runImport(config));
    }
  }

  async function runImport(config) {
    const formData = new FormData();
    for (const [fieldName, inputId] of config.files) {
      const input = document.getElementById(inputId);
      if (!input.files || input.files.length === 0) {
        addLog("error", `${config.label} update blocked`, `${inputId} file is required.`);
        return;
      }
      formData.append(fieldName, input.files[0]);
    }

    const button = document.getElementById(config.buttonId);
    button.disabled = true;
    try {
      const response = await fetch(`${baseUrl}${config.endpoint}`, {
        method: "POST",
        body: formData
      });
      const body = await readJsonResponse(response);
      if (!response.ok) {
        addLog("error", body.message || `${config.label} update blocked`, (body.errors || []).join("; "));
        return;
      }
      addLog("ok", body.message || `${config.label} updated successfully`, `${body.rowCount} rows imported`);
      await refreshFunctionCache(false);
    } catch (error) {
      addLog("error", `${config.label} update failed`, error instanceof Error ? error.message : String(error));
    } finally {
      button.disabled = false;
      checkHealth();
    }
  }

  async function readJsonResponse(response) {
    const text = await response.text();
    try {
      return text ? JSON.parse(text) : {};
    } catch {
      return {
        message: "LocalAct returned an unreadable response",
        errors: [text.slice(0, 300)]
      };
    }
  }

  async function convertFormulasToValues() {
    addLog("ok", "Formula conversion started", "Scanning the active worksheet.");
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["formulas", "values", "rowCount", "columnCount"]);
        await context.sync();

        if (usedRange.isNullObject) {
          addLog("ok", "No formulas converted", "The active worksheet is empty.");
          return;
        }

        let count = 0;
        let formulaCount = 0;
        const replacementValues = usedRange.formulas.map((rowValues, row) =>
          rowValues.map((formulaValue, column) => {
            const formula = String(formulaValue || "");
            if (formula) {
              formulaCount += 1;
            }
            if (isLocalActFormula(formula)) {
              count += 1;
              return usedRange.values[row][column];
            }
            return formulaValue;
          })
        );

        if (count > 0) {
          usedRange.formulas = replacementValues;
        }
        await context.sync();
        addLog(
          "ok",
          count > 0 ? "Formulas converted" : "No LocalAct formulas found",
          `${count} LocalAct formulas converted on the active worksheet. ${formulaCount} total formulas scanned.`
        );
      });
    } catch (error) {
      addLog("error", "Formula conversion failed", error instanceof Error ? error.message : String(error));
    }
  }

  function isLocalActFormula(formula) {
    return /(?:^|[=@])\s*(?:_xlfn\.)?LOCALACT\.(LOAD_PL|LOAD_BS|LOAD_PL_DEPT|LOAD_OPS)\s*\(/i.test(formula);
  }

  async function checkHealth() {
    try {
      const response = await fetch(`${baseUrl}/health`);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      setStatus("ok", "Local service is running");
    } catch {
      setStatus("bad", "Local service is not running");
    }
  }

  async function refreshFunctionCache(showLog) {
    try {
      const response = await fetch(`${baseUrl}/api/cache`);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      const cache = await response.json();
      const serialized = JSON.stringify(cache);
      if (window.OfficeRuntime && window.OfficeRuntime.storage) {
        await window.OfficeRuntime.storage.setItem(cacheStorageKey, serialized);
      }
      window.localStorage.setItem(cacheStorageKey, serialized);
      if (showLog) {
        addLog("ok", "Formula cache refreshed", summarizeCache(cache));
      }
    } catch (error) {
      if (showLog) {
        addLog("error", "Formula cache refresh failed", error instanceof Error ? error.message : String(error));
      }
    }
  }

  function summarizeCache(cache) {
    return `P&L ${Object.keys(cache.pl || {}).length}; BS ${Object.keys(cache.bs || {}).length}; Departments ${Object.keys(cache.plDepartment || {}).length}; Ops ${Object.keys(cache.ops || {}).length}`;
  }

  function setStatus(state, text) {
    const className = state === "ok" ? "dot-ok" : state === "bad" ? "dot-bad" : "dot-warn";
    serviceStatus.innerHTML = `<span class="dot ${className}"></span><span>${text}</span>`;
    footerStatus.textContent = state === "ok" ? "Connected" : "Disconnected";
  }

  function addLog(type, title, detail) {
    const item = document.createElement("div");
    const now = new Date();
    item.className = "log-item";
    item.innerHTML = `
      <span class="${type === "ok" ? "icon-ok" : "icon-error"}">${type === "ok" ? "OK" : "ERR"}</span>
      <time>${now.toLocaleString()}</time>
      <div><strong>${escapeHtml(title)}</strong>${detail ? `<div class="log-details">${escapeHtml(detail)}</div>` : ""}</div>
    `;
    logList.prepend(item);
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (char) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "\"": "&quot;",
      "'": "&#039;"
    })[char]);
  }
})();

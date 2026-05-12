(function () {
  const root = globalThis;
  const baseUrl = "https://localhost:3000";

  async function loadPl(summaryAccount, date) {
    return runFunction("LOAD_PL", () =>
      fetchValue("/api/value/pl", { summaryAccount, date: normalizeExactDate(date) }, "LOAD_PL")
    );
  }

  async function loadBs(summaryAccount, date) {
    return runFunction("LOAD_BS", () =>
      fetchValue("/api/value/bs", { summaryAccount, date: normalizeExactDate(date) }, "LOAD_BS")
    );
  }

  async function loadPlDept(summaryAccount, date, summaryDepartment) {
    return runFunction("LOAD_PL_DEPT", () =>
      fetchValue(
        "/api/value/pl-department",
        { summaryAccount, date: normalizeExactDate(date), summaryDepartment },
        "LOAD_PL_DEPT"
      )
    );
  }

  async function loadOps(metricLabel, date, channel) {
    return runFunction("LOAD_OPS", () =>
      fetchValue("/api/value/ops", { metricLabel, date: normalizeExactDate(date), channel }, "LOAD_OPS")
    );
  }

  function diag() {
    return 123;
  }

  async function diagBackend() {
    try {
      const response = await fetch(`${baseUrl}/health`);
      return Number(response.status || 0);
    } catch {
      return -1;
    }
  }

  async function diagBackendText() {
    try {
      const response = await fetch(`${baseUrl}/health`);
      const text = await response.text();
      return `status=${response.status}; body=${text.slice(0, 120)}`;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return `error=${message}`;
    }
  }

  function diagDate(date) {
    return normalizeExactDate(date);
  }

  async function runFunction(functionName, fn) {
    try {
      return await fn();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return customFunctionError(`${functionName} failed: ${message}`);
    }
  }

  async function fetchValue(path, params, functionName) {
    try {
      const search = new URLSearchParams();
      for (const [key, value] of Object.entries(params)) {
        search.set(key, String(value ?? "").trim());
      }
      const response = await fetch(`${baseUrl}${path}?${search.toString()}`);
      if (!response.ok) {
        return customFunctionError(`${functionName} failed: local service returned ${response.status}`);
      }
      const body = await response.json();
      return Number(body && body.value ? body.value : 0);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return customFunctionError(`${functionName} failed: ${message}`);
    }
  }

  function normalizeExactDate(value) {
    const date = parseInputDate(value);
    if (!date) {
      throw new Error("Date must be a valid Excel date.");
    }
    return date.toISOString().slice(0, 10);
  }

  function parseInputDate(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    }

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    }

    const text = String(value ?? "").trim();
    const slashDate = text.match(/^(\d{1,2})\/([A-Za-z]{3,})\/(\d{2}|\d{4})$/);
    if (slashDate) {
      const day = Number(slashDate[1]);
      const month = monthIndex(slashDate[2]);
      const year = normalizeYear(Number(slashDate[3]));
      if (month >= 0) {
        return new Date(Date.UTC(year, month, day));
      }
    }

    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
    }

    return null;
  }

  function monthIndex(value) {
    return ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(
      String(value || "").slice(0, 3).toLowerCase()
    );
  }

  function normalizeYear(year) {
    return year < 100 ? 2000 + year : year;
  }

  function customFunctionError(message) {
    if (root.CustomFunctions && root.CustomFunctions.Error) {
      return new root.CustomFunctions.Error(root.CustomFunctions.ErrorCode.invalidValue, message);
    }
    throw new Error(message);
  }

  if (root.CustomFunctions && root.CustomFunctions.associate) {
    root.CustomFunctions.associate("DIAG", diag);
    root.CustomFunctions.associate("DIAG_BACKEND", diagBackend);
    root.CustomFunctions.associate("DIAG_BACKEND_TEXT", diagBackendText);
    root.CustomFunctions.associate("DIAG_DATE", diagDate);
    root.CustomFunctions.associate("LOAD_PL", loadPl);
    root.CustomFunctions.associate("LOAD_BS", loadBs);
    root.CustomFunctions.associate("LOAD_PL_DEPT", loadPlDept);
    root.CustomFunctions.associate("LOAD_OPS", loadOps);
  }
})();

(function () {
  const root = globalThis;
  const baseUrl = "https://localhost:3000";

  async function loadPl(summaryAccount, date) {
    return fetchValue("/api/value/pl", { summaryAccount, date: normalizeExactDate(date) }, "LOAD_PL");
  }

  async function loadBs(summaryAccount, date) {
    return fetchValue("/api/value/bs", { summaryAccount, date: normalizeExactDate(date) }, "LOAD_BS");
  }

  async function loadPlDept(summaryAccount, date, summaryDepartment) {
    return fetchValue(
      "/api/value/pl-department",
      { summaryAccount, date: normalizeExactDate(date), summaryDepartment },
      "LOAD_PL_DEPT"
    );
  }

  async function loadOps(metricLabel, date, channel) {
    return fetchValue("/api/value/ops", { metricLabel, date: normalizeExactDate(date), channel }, "LOAD_OPS");
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

    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
    }

    return null;
  }

  function customFunctionError(message) {
    if (root.CustomFunctions && root.CustomFunctions.Error) {
      return new root.CustomFunctions.Error(root.CustomFunctions.ErrorCode.invalidValue, message);
    }
    throw new Error(message);
  }

  if (root.CustomFunctions && root.CustomFunctions.associate) {
    root.CustomFunctions.associate("LOAD_PL", loadPl);
    root.CustomFunctions.associate("LOAD_BS", loadBs);
    root.CustomFunctions.associate("LOAD_PL_DEPT", loadPlDept);
    root.CustomFunctions.associate("LOAD_OPS", loadOps);
  }
})();

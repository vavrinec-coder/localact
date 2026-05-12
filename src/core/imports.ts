import type {
  BsEnrichedRow,
  MappingRow,
  OpsRow,
  PlDataRow,
  PlDepartmentDataRow,
  PlDepartmentEnrichedRow,
  PlEnrichedRow,
  RawRow
} from "./types.js";
import { assertColumns, readAmount, readExcelDate, readText, trimText, uniqueSorted } from "./normalize.js";

type BuildResult<T> = {
  errors: string[];
  rows: T[];
};

type EnrichedBuildResult<T> = {
  errors: string[];
  enrichedRows: T[];
  mappings?: MappingRow[];
};

export function buildPlImport(dataRows: RawRow[], mappingRows: RawRow[]): EnrichedBuildResult<PlEnrichedRow> {
  const errors = [
    ...assertColumns(dataRows, ["Account label", "Date", "Amount"], "PL_Data"),
    ...assertColumns(mappingRows, ["Account label", "Summary Account"], "PL_Mapping")
  ];
  const parsedData = parsePlDataRows(dataRows, errors);
  const mappings = parseAccountMappings(mappingRows, errors, "PL_Mapping");
  const mappingByAccount = new Map(mappings.map((row) => [row.key, row.summary]));
  const missing = uniqueSorted(parsedData.filter((row) => !mappingByAccount.has(row.accountLabel)).map((row) => row.accountLabel));

  errors.push(...missing.map((account) => `Missing Summary Account mapping for Account label: ${account}`));
  if (errors.length > 0) {
    return { errors, enrichedRows: [], mappings };
  }

  return {
    errors: [],
    mappings,
    enrichedRows: parsedData.map((row) => ({
      ...row,
      summaryAccount: mappingByAccount.get(row.accountLabel) ?? ""
    }))
  };
}

export function buildPlDepartmentImport(
  dataRows: RawRow[],
  accountMappings: MappingRow[],
  departmentMappingRows: RawRow[]
): EnrichedBuildResult<PlDepartmentEnrichedRow> {
  const errors = [
    ...assertColumns(dataRows, ["Account label", "Date", "Amount", "Department"], "PL_Department_Data"),
    ...assertColumns(departmentMappingRows, ["Department", "Summary Department"], "Department_Mapping")
  ];
  const parsedData = parsePlDepartmentDataRows(dataRows, errors);
  const accountByLabel = new Map(accountMappings.map((row) => [row.key, row.summary]));
  const departmentMappings = parseDepartmentMappings(departmentMappingRows, errors);
  const departmentByName = new Map(departmentMappings.map((row) => [row.key, row.summary]));

  const missingAccounts = uniqueSorted(parsedData.filter((row) => !accountByLabel.has(row.accountLabel)).map((row) => row.accountLabel));
  const missingDepartments = uniqueSorted(parsedData.filter((row) => !departmentByName.has(row.department)).map((row) => row.department));

  errors.push(...missingAccounts.map((account) => `Missing Summary Account mapping for Account label: ${account}`));
  errors.push(...missingDepartments.map((department) => `Missing Summary Department mapping for Department: ${department}`));
  if (errors.length > 0) {
    return { errors, enrichedRows: [], mappings: departmentMappings };
  }

  return {
    errors: [],
    mappings: departmentMappings,
    enrichedRows: parsedData.map((row) => ({
      ...row,
      summaryAccount: accountByLabel.get(row.accountLabel) ?? "",
      summaryDepartment: departmentByName.get(row.department) ?? ""
    }))
  };
}

export function buildBsImport(dataRows: RawRow[], mappingRows: RawRow[]): EnrichedBuildResult<BsEnrichedRow> {
  const plResult = buildPlImport(dataRows, mappingRows);
  return {
    errors: plResult.errors.map((error) => error.replaceAll("PL_Data", "BS_Data").replaceAll("PL_Mapping", "BS_Mapping")),
    mappings: plResult.mappings,
    enrichedRows: plResult.enrichedRows
  };
}

export function buildOpsImport(dataRows: RawRow[]): BuildResult<OpsRow> {
  const errors = assertColumns(dataRows, ["Metric label", "Product", "Date", "Channel", "Amount"], "Ops");
  const rows = dataRows.map((row, index) => {
    const rowNumber = index + 2;
    return {
      metricLabel: readText(row, "Metric label", errors, rowNumber),
      product: readText(row, "Product", errors, rowNumber),
      date: readExcelDate(row, "Date", errors, rowNumber),
      channel: readText(row, "Channel", errors, rowNumber),
      amount: readAmount(row, "Amount", errors, rowNumber)
    };
  });

  return errors.length > 0 ? { errors, rows: [] } : { errors: [], rows };
}

export function parseAccountMappings(rows: RawRow[], errors: string[], label: string): MappingRow[] {
  return rows.map((row, index) => {
    const rowNumber = index + 2;
    return {
      key: readText(row, "Account label", errors, rowNumber),
      summary: readText(row, "Summary Account", errors, rowNumber)
    };
  }).filter((row) => row.key && row.summary);
}

function parseDepartmentMappings(rows: RawRow[], errors: string[]): MappingRow[] {
  return rows.map((row, index) => {
    const rowNumber = index + 2;
    return {
      key: readText(row, "Department", errors, rowNumber),
      summary: readText(row, "Summary Department", errors, rowNumber)
    };
  }).filter((row) => row.key && row.summary);
}

function parsePlDataRows(rows: RawRow[], errors: string[]): PlDataRow[] {
  return rows.map((row, index) => {
    const rowNumber = index + 2;
    return {
      accountLabel: readText(row, "Account label", errors, rowNumber),
      date: readExcelDate(row, "Date", errors, rowNumber),
      amount: readAmount(row, "Amount", errors, rowNumber)
    };
  });
}

function parsePlDepartmentDataRows(rows: RawRow[], errors: string[]): PlDepartmentDataRow[] {
  return rows.map((row, index) => {
    const rowNumber = index + 2;
    return {
      accountLabel: readText(row, "Account label", errors, rowNumber),
      date: readExcelDate(row, "Date", errors, rowNumber),
      amount: readAmount(row, "Amount", errors, rowNumber),
      department: readText(row, "Department", errors, rowNumber)
    };
  });
}

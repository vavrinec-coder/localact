import ExcelJS from "exceljs";
import type { RawRow } from "./core/types.js";

export async function workbookBufferToRows(buffer: Buffer): Promise<RawRow[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as never);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    return [];
  }

  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values as unknown[];
  const rows: RawRow[] = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }
    const output: RawRow = {};
    let hasValue = false;
    for (let column = 1; column < headers.length; column += 1) {
      const header = String(headers[column] ?? "").trim();
      if (!header) {
        continue;
      }
      const cellValue = normalizeCellValue(row.getCell(column).value);
      if (cellValue !== "") {
        hasValue = true;
      }
      output[header] = cellValue;
    }
    if (hasValue) {
      rows.push(output);
    }
  });

  return rows;
}

function normalizeCellValue(value: ExcelJS.CellValue): unknown {
  if (value instanceof Date) {
    return value;
  }
  if (value && typeof value === "object") {
    if ("result" in value) {
      return value.result ?? "";
    }
    if ("text" in value) {
      return value.text ?? "";
    }
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join("");
    }
  }
  return value ?? "";
}

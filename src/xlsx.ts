import ExcelJS from "exceljs";
import JSZip from "jszip";
import type { RawRow } from "./core/types.js";

export async function workbookBufferToRows(buffer: Buffer): Promise<RawRow[]> {
  const workbook = new ExcelJS.Workbook();
  const readableBuffer = await repairWorkbookContentTypes(buffer);
  await workbook.xlsx.load(readableBuffer as never);
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

async function repairWorkbookContentTypes(buffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);
  const contentTypesFile = zip.file("[Content_Types].xml");
  if (!contentTypesFile) {
    return buffer;
  }

  let contentTypes = await contentTypesFile.async("string");
  const hasWorkbookOverride = contentTypes.includes('PartName="/xl/workbook.xml"');
  const hasInvalidXmlDefault = contentTypes.includes(
    'Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"'
  );
  if (hasWorkbookOverride && !hasInvalidXmlDefault) {
    return buffer;
  }

  contentTypes = contentTypes.replace(
    'Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
    'Extension="xml" ContentType="application/xml"'
  );

  const workbookContentType = [
    '<Override PartName="/xl/workbook.xml"',
    ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />'
  ].join("");
  const repairedContentTypes = hasWorkbookOverride ? contentTypes : contentTypes.replace("</Types>", `${workbookContentType}</Types>`);
  zip.file("[Content_Types].xml", repairedContentTypes);
  await normalizeSpreadsheetNamespacePrefix(zip);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}

async function normalizeSpreadsheetNamespacePrefix(zip: JSZip): Promise<void> {
  const xmlFiles = Object.values(zip.files).filter((file) => !file.dir && file.name.endsWith(".xml"));
  await Promise.all(
    xmlFiles.map(async (file) => {
      let xml = await file.async("string");
      if (!xml.includes("<x:") && !xml.includes("</x:") && !xml.includes("<tableParts")) {
        return;
      }
      xml = xml.replace(/<x:tableParts[\s\S]*?<\/x:tableParts>/g, "");
      xml = xml.replace(/<tableParts[\s\S]*?<\/tableParts>/g, "");
      zip.file(file.name, xml.replace(/(<\/?)(x):/g, "$1"));
    })
  );
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

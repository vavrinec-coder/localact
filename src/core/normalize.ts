import type { RawRow } from "./types.js";

export function trimText(value: unknown): string {
  return String(value ?? "").trim();
}

export function readText(row: RawRow, column: string, errors: string[], rowNumber: number): string {
  const value = trimText(row[column]);
  if (!value) {
    errors.push(`Row ${rowNumber}: ${column} is required`);
  }
  return value;
}

export function readAmount(row: RawRow, column: string, errors: string[], rowNumber: number): number {
  const value = row[column];
  const numberValue = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(numberValue)) {
    errors.push(`Row ${rowNumber}: ${column} must be numeric`);
    return 0;
  }
  return numberValue;
}

export function readExcelDate(row: RawRow, column: string, errors: string[], rowNumber: number): string {
  const value = row[column];
  const date = parseExcelDate(value);
  if (!date) {
    errors.push(`Row ${rowNumber}: ${column} must be a valid Excel date`);
    return "";
  }

  const iso = date.toISOString().slice(0, 10);
  if (!isMonthEnd(date)) {
    errors.push(`Row ${rowNumber}: ${column} must be a month-end date`);
  }
  return iso;
}

export function parseExcelDate(value: unknown): Date | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    const millis = Date.UTC(1899, 11, 30) + Math.floor(value) * 86_400_000;
    return new Date(millis);
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }

  if (typeof value === "string" && value.trim()) {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
    }
  }

  return null;
}

export function isMonthEnd(date: Date): boolean {
  const nextDay = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate() + 1));
  return nextDay.getUTCDate() === 1;
}

export function assertColumns(rows: RawRow[], requiredColumns: string[], label: string): string[] {
  if (rows.length === 0) {
    return [`${label} has no rows`];
  }

  const available = new Set(Object.keys(rows[0] ?? {}));
  return requiredColumns
    .filter((column) => !available.has(column))
    .map((column) => `${label} is missing required column: ${column}`);
}

export function uniqueSorted(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

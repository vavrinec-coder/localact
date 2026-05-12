export type RawRow = Record<string, unknown>;

export type ImportResult = {
  ok: true;
  message: string;
  rowCount: number;
  warnings: string[];
};

export type ImportFailure = {
  ok: false;
  message: string;
  errors: string[];
};

export type SectionName = "pl" | "plDepartment" | "bs" | "ops";

export type PlDataRow = {
  accountLabel: string;
  date: string;
  amount: number;
};

export type MappingRow = {
  key: string;
  summary: string;
};

export type PlEnrichedRow = PlDataRow & {
  summaryAccount: string;
};

export type PlDepartmentDataRow = PlDataRow & {
  department: string;
};

export type PlDepartmentEnrichedRow = PlDepartmentDataRow & {
  summaryAccount: string;
  summaryDepartment: string;
};

export type BsEnrichedRow = PlEnrichedRow;

export type OpsRow = {
  metricLabel: string;
  date: string;
  channel: string;
  amount: number;
};

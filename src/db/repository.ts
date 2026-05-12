import { DatabaseSync } from "node:sqlite";
import type { BsEnrichedRow, MappingRow, OpsRow, PlDepartmentEnrichedRow, PlEnrichedRow } from "../core/types.js";
import { schemaSql } from "./schema.js";

export type FunctionCache = {
  pl: Record<string, number>;
  bs: Record<string, number>;
  plDepartment: Record<string, number>;
  ops: Record<string, number>;
};

export class LocalActRepository {
  private readonly db: DatabaseSync;

  constructor(path: string) {
    this.db = new DatabaseSync(path);
    this.db.exec("PRAGMA foreign_keys = ON;");
    this.db.exec(schemaSql);
  }

  close(): void {
    this.db.close();
  }

  getPlMappings(): MappingRow[] {
    return this.db
      .prepare("SELECT account_label AS key, summary_account AS summary FROM pl_mapping ORDER BY account_label")
      .all() as MappingRow[];
  }

  replacePl(mappings: MappingRow[], rows: PlEnrichedRow[]): void {
    this.transaction(() => {
      this.db.exec("DELETE FROM pl_enriched; DELETE FROM pl_mapping;");
      const insertMapping = this.db.prepare("INSERT INTO pl_mapping (account_label, summary_account) VALUES (?, ?)");
      for (const mapping of mappings) {
        insertMapping.run(mapping.key, mapping.summary);
      }
      const insertRow = this.db.prepare(
        "INSERT INTO pl_enriched (account_label, date, amount, summary_account) VALUES (?, ?, ?, ?)"
      );
      for (const row of rows) {
        insertRow.run(row.accountLabel, row.date, row.amount, row.summaryAccount);
      }
    });
  }

  replacePlDepartment(mappings: MappingRow[], rows: PlDepartmentEnrichedRow[]): void {
    this.transaction(() => {
      this.db.exec("DELETE FROM pl_department_enriched; DELETE FROM department_mapping;");
      const insertMapping = this.db.prepare("INSERT INTO department_mapping (department, summary_department) VALUES (?, ?)");
      for (const mapping of mappings) {
        insertMapping.run(mapping.key, mapping.summary);
      }
      const insertRow = this.db.prepare(
        `INSERT INTO pl_department_enriched
         (account_label, date, amount, department, summary_account, summary_department)
         VALUES (?, ?, ?, ?, ?, ?)`
      );
      for (const row of rows) {
        insertRow.run(row.accountLabel, row.date, row.amount, row.department, row.summaryAccount, row.summaryDepartment);
      }
    });
  }

  replaceBs(mappings: MappingRow[], rows: BsEnrichedRow[]): void {
    this.transaction(() => {
      this.db.exec("DELETE FROM bs_enriched; DELETE FROM bs_mapping;");
      const insertMapping = this.db.prepare("INSERT INTO bs_mapping (account_label, summary_account) VALUES (?, ?)");
      for (const mapping of mappings) {
        insertMapping.run(mapping.key, mapping.summary);
      }
      const insertRow = this.db.prepare(
        "INSERT INTO bs_enriched (account_label, date, amount, summary_account) VALUES (?, ?, ?, ?)"
      );
      for (const row of rows) {
        insertRow.run(row.accountLabel, row.date, row.amount, row.summaryAccount);
      }
    });
  }

  replaceOps(rows: OpsRow[]): void {
    this.transaction(() => {
      this.db.exec("DELETE FROM ops;");
      const insertRow = this.db.prepare("INSERT INTO ops (metric_label, date, channel, amount) VALUES (?, ?, ?, ?)");
      for (const row of rows) {
        insertRow.run(row.metricLabel, row.date, row.channel, row.amount);
      }
    });
  }

  sumPl(summaryAccount: string, date: string): number {
    return this.sum("SELECT COALESCE(SUM(amount), 0) AS value FROM pl_enriched WHERE summary_account = ? AND date = ?", [
      summaryAccount,
      date
    ]);
  }

  sumPlDepartment(summaryAccount: string, date: string, summaryDepartment: string): number {
    return this.sum(
      `SELECT COALESCE(SUM(amount), 0) AS value
       FROM pl_department_enriched
       WHERE summary_account = ? AND date = ? AND summary_department = ?`,
      [summaryAccount, date, summaryDepartment]
    );
  }

  sumBs(summaryAccount: string, date: string): number {
    return this.sum("SELECT COALESCE(SUM(amount), 0) AS value FROM bs_enriched WHERE summary_account = ? AND date = ?", [
      summaryAccount,
      date
    ]);
  }

  sumOps(metricLabel: string, date: string, channel: string): number {
    return this.sum("SELECT COALESCE(SUM(amount), 0) AS value FROM ops WHERE metric_label = ? AND date = ? AND channel = ?", [
      metricLabel,
      date,
      channel
    ]);
  }

  getFunctionCache(): FunctionCache {
    return {
      pl: this.aggregate("SELECT summary_account, date, SUM(amount) AS value FROM pl_enriched GROUP BY summary_account, date", [
        "summary_account",
        "date"
      ]),
      bs: this.aggregate("SELECT summary_account, date, SUM(amount) AS value FROM bs_enriched GROUP BY summary_account, date", [
        "summary_account",
        "date"
      ]),
      plDepartment: this.aggregate(
        `SELECT summary_account, date, summary_department, SUM(amount) AS value
         FROM pl_department_enriched
         GROUP BY summary_account, date, summary_department`,
        ["summary_account", "date", "summary_department"]
      ),
      ops: this.aggregate("SELECT metric_label, date, channel, SUM(amount) AS value FROM ops GROUP BY metric_label, date, channel", [
        "metric_label",
        "date",
        "channel"
      ])
    };
  }

  private sum(sql: string, params: (string | number | null)[]): number {
    const row = this.db.prepare(sql).get(...params) as { value?: number } | undefined;
    return Number(row?.value ?? 0);
  }

  private aggregate(sql: string, keyColumns: string[]): Record<string, number> {
    const output: Record<string, number> = {};
    const rows = this.db.prepare(sql).all() as Record<string, unknown>[];
    for (const row of rows) {
      const key = keyColumns.map((column) => normalizeCacheKeyPart(row[column])).join("|");
      output[key] = Number(row.value ?? 0);
    }
    return output;
  }

  private transaction(fn: () => void): void {
    this.db.exec("BEGIN IMMEDIATE;");
    try {
      fn();
      this.db.exec("COMMIT;");
    } catch (error) {
      this.db.exec("ROLLBACK;");
      throw error;
    }
  }
}

function normalizeCacheKeyPart(value: unknown): string {
  return String(value ?? "").trim().toLowerCase();
}

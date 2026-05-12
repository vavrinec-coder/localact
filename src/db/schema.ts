export const schemaSql = `
CREATE TABLE IF NOT EXISTS pl_mapping (
  account_label TEXT PRIMARY KEY,
  summary_account TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS pl_enriched (
  account_label TEXT NOT NULL,
  date TEXT NOT NULL,
  amount REAL NOT NULL,
  summary_account TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS department_mapping (
  department TEXT PRIMARY KEY,
  summary_department TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS pl_department_enriched (
  account_label TEXT NOT NULL,
  date TEXT NOT NULL,
  amount REAL NOT NULL,
  department TEXT NOT NULL,
  summary_account TEXT NOT NULL,
  summary_department TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS bs_mapping (
  account_label TEXT PRIMARY KEY,
  summary_account TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS bs_enriched (
  account_label TEXT NOT NULL,
  date TEXT NOT NULL,
  amount REAL NOT NULL,
  summary_account TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS ops (
  metric_label TEXT NOT NULL,
  date TEXT NOT NULL,
  channel TEXT NOT NULL,
  amount REAL NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_pl_lookup ON pl_enriched(summary_account, date);
CREATE INDEX IF NOT EXISTS idx_pl_dept_lookup ON pl_department_enriched(summary_account, date, summary_department);
CREATE INDEX IF NOT EXISTS idx_bs_lookup ON bs_enriched(summary_account, date);
CREATE INDEX IF NOT EXISTS idx_ops_lookup ON ops(metric_label, date, channel);
`;

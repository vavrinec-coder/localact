# LocalAct

LocalAct is a local-only Excel web add-in for importing standardized `.xlsx` actuals files into a local SQLite database and loading summarized values back into Excel.

## Requirements

- Microsoft 365 Excel for Windows
- Node.js 24 or newer

## Start LocalAct

```powershell
npm install
npm start
```

The service starts at `https://localhost:3000` and stores the SQLite database at `data/localact.sqlite`.

The local HTTPS certificate is self-signed. The first browser visit to `https://localhost:3000/health` may require accepting the local certificate warning before Excel can load the add-in cleanly.

For browser-only UI preview, start with `LOCALACT_HTTP_PORT=3001 npm start` and open `http://localhost:3001/taskpane.html`. Excel should still use the HTTPS manifest URL.

## Excel Manifest

Use `manifest.xml` for local sideloading. It points Excel to:

- `https://localhost:3000/taskpane.html`
- `https://localhost:3000/functions.js`
- `https://localhost:3000/functions.json`

## Imports

Each update replaces all existing data for that section. Text fields are trimmed. Missing required mappings block the update.

- P&L: `PL_Data` plus `PL_Mapping`
- P&L by Departments: `PL_Department_Data` plus `Department_Mapping`; reuses the current `PL_Mapping`
- Balance Sheet: `BS_Data` plus `BS_Mapping`
- Operations: `Ops`

## Custom Functions

Use these formulas in Excel:

```excel
=LocalAct.LOAD_PL(summary_account, date)
=LocalAct.LOAD_BS(summary_account, date)
=LocalAct.LOAD_PL_DEPT(summary_account, date, summary_department)
=LocalAct.LOAD_OPS(metric_label, date, channel)
```

Dates are exact Excel dates and should be month-end dates.

## Formulas to Values

The task pane button scans the active worksheet only. It converts LocalAct formulas to their current calculated values and leaves other formulas untouched.

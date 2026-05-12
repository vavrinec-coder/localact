import {
  buildBsImport,
  buildOpsImport,
  buildPlDepartmentImport,
  buildPlImport
} from "../src/core/imports.js";

describe("import builders", () => {
  it("trims text and enriches P&L rows with summary accounts", () => {
    const result = buildPlImport(
      [{ "Account label": " Revenue ", Date: new Date("2026-04-30T00:00:00Z"), Amount: 100 }],
      [{ "Account label": "Revenue", "Summary Account": " Net Revenue " }]
    );

    expect(result.errors).toEqual([]);
    expect(result.enrichedRows).toEqual([
      {
        accountLabel: "Revenue",
        date: "2026-04-30",
        amount: 100,
        summaryAccount: "Net Revenue"
      }
    ]);
  });

  it("blocks P&L imports when any data account has no mapping", () => {
    const result = buildPlImport(
      [{ "Account label": "COGS", Date: new Date("2026-04-30T00:00:00Z"), Amount: 42 }],
      [{ "Account label": "Revenue", "Summary Account": "Net Revenue" }]
    );

    expect(result.errors).toContain("Missing Summary Account mapping for Account label: COGS");
    expect(result.enrichedRows).toEqual([]);
  });

  it("blocks department P&L imports when account or department mappings are missing", () => {
    const result = buildPlDepartmentImport(
      [
        {
          "Account label": "Revenue",
          Date: new Date("2026-04-30T00:00:00Z"),
          Amount: 100,
          Department: "Sales"
        },
        {
          "Account label": "COGS",
          Date: new Date("2026-04-30T00:00:00Z"),
          Amount: 50,
          Department: "Ops"
        }
      ],
      [{ key: "Revenue", summary: "Net Revenue" }],
      [{ Department: "Sales", "Summary Department": "Commercial" }]
    );

    expect(result.errors).toEqual([
      "Missing Summary Account mapping for Account label: COGS",
      "Missing Summary Department mapping for Department: Ops"
    ]);
    expect(result.enrichedRows).toEqual([]);
  });

  it("builds balance sheet rows separately from P&L mappings", () => {
    const result = buildBsImport(
      [{ "Account label": "Cash", Date: 46142, Amount: 1000 }],
      [{ "Account label": "Cash", "Summary Account": "Cash and Equivalents" }]
    );

    expect(result.errors).toEqual([]);
    expect(result.enrichedRows[0]).toMatchObject({
      accountLabel: "Cash",
      date: "2026-04-30",
      summaryAccount: "Cash and Equivalents"
    });
  });

  it("builds operations rows with raw channel values", () => {
    const result = buildOpsImport([
      { "Metric label": "Orders", Product: " Device ", Date: new Date("2026-04-30T00:00:00Z"), Channel: " Amazon ", Amount: 8 }
    ]);

    expect(result.errors).toEqual([]);
    expect(result.rows).toEqual([
      { metricLabel: "Orders", product: "Device", date: "2026-04-30", channel: "Amazon", amount: 8 }
    ]);
  });
});

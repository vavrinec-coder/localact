import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";

type AssociateMap = Record<string, (...args: unknown[]) => unknown>;

function loadCustomFunctions(fetchImpl: typeof fetch = fetch): AssociateMap {
  const associates: AssociateMap = {};
  const source = fs.readFileSync(path.resolve("public/functions.js"), "utf8");
  const context = vm.createContext({
    fetch: fetchImpl,
    URLSearchParams,
    Date,
    Number,
    String,
    Error,
    CustomFunctions: {
      Error: class CustomFunctionError extends Error {
        constructor(_code: string, message: string) {
          super(message);
          this.name = "CustomFunctionError";
        }
      },
      ErrorCode: {
        invalidValue: "invalidValue"
      },
      associate(name: string, fn: (...args: unknown[]) => unknown) {
        associates[name] = fn;
      }
    }
  });

  vm.runInContext(source, context);
  return associates;
}

describe("custom functions", () => {
  it("registers diagnostics and loader functions", () => {
    const associates = loadCustomFunctions();

    expect(Object.keys(associates).sort()).toEqual([
      "DIAG",
      "DIAG_BACKEND",
      "DIAG_DATE",
      "LOAD_BS",
      "LOAD_OPS",
      "LOAD_PL",
      "LOAD_PL_DEPT"
    ]);
    expect(associates.DIAG()).toBe(123);
  });

  it("normalizes Excel serial and visible date strings to exact dates", () => {
    const associates = loadCustomFunctions();

    expect(associates.DIAG_DATE(46142)).toBe("2026-04-30");
    expect(associates.DIAG_DATE("30/Apr/26")).toBe("2026-04-30");
  });

  it("calls the P&L lookup endpoint with normalized parameters", async () => {
    let requestedUrl = "";
    const associates = loadCustomFunctions((async (url: string) => {
      requestedUrl = url;
      return new Response(JSON.stringify({ value: -57464.56 }), { status: 200 });
    }) as typeof fetch);

    await expect(associates.LOAD_PL("Discounts", "30/Apr/26")).resolves.toBe(-57464.56);
    expect(requestedUrl).toBe("https://localhost:3000/api/value/pl?summaryAccount=Discounts&date=2026-04-30");
  });
});

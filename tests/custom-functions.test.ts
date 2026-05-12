import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";

type AssociateMap = Record<string, (...args: unknown[]) => unknown>;

function loadCustomFunctions(fetchImpl: typeof fetch = fetch, storage: Record<string, string> = {}): AssociateMap {
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
    },
    OfficeRuntime: {
      storage: {
        async getItem(key: string) {
          return storage[key] || "";
        }
      }
    },
    localStorage: {
      getItem(key: string) {
        return storage[key] || "";
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
      "DIAG_BACKEND_TEXT",
      "DIAG_CACHE",
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

  it("loads P&L values from the shared formula cache", async () => {
    const associates = loadCustomFunctions(fetch, {
      "localact.functionCache.v1": JSON.stringify({
        pl: { "discounts|2026-04-30": -57464.56 },
        bs: {},
        plDepartment: {},
        ops: {}
      })
    });

    await expect(associates.DIAG_CACHE()).resolves.toBe("pl=1; bs=0; plDepartment=0; ops=0");
    await expect(associates.LOAD_PL("Discounts", "30/Apr/26")).resolves.toBe(-57464.56);
  });
});

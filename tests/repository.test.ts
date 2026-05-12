import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { LocalActRepository } from "../src/db/repository.js";

function makeRepo() {
  const dir = mkdtempSync(join(tmpdir(), "localact-"));
  const repo = new LocalActRepository(join(dir, "test.sqlite"));
  return {
    repo,
    cleanup: () => {
      repo.close();
      rmSync(dir, { recursive: true, force: true });
    }
  };
}

describe("LocalActRepository", () => {
  it("replaces P&L data and returns summed values", () => {
    const { repo, cleanup } = makeRepo();
    try {
      repo.replacePl(
        [{ key: "Revenue", summary: "Net Revenue" }],
        [
          { accountLabel: "Revenue", summaryAccount: "Net Revenue", date: "2026-04-30", amount: 10 },
          { accountLabel: "Revenue", summaryAccount: "Net Revenue", date: "2026-04-30", amount: 15 }
        ]
      );
      expect(repo.sumPl("Net Revenue", "2026-04-30")).toBe(25);

      repo.replacePl([{ key: "COGS", summary: "COGS" }], [{ accountLabel: "COGS", summaryAccount: "COGS", date: "2026-04-30", amount: 7 }]);
      expect(repo.sumPl("Net Revenue", "2026-04-30")).toBe(0);
      expect(repo.getPlMappings()).toEqual([{ key: "COGS", summary: "COGS" }]);
    } finally {
      cleanup();
    }
  });

  it("replaces department P&L and sums by summary department", () => {
    const { repo, cleanup } = makeRepo();
    try {
      repo.replacePlDepartment(
        [{ key: "Sales", summary: "Commercial" }],
        [
          {
            accountLabel: "Revenue",
            summaryAccount: "Net Revenue",
            date: "2026-04-30",
            amount: 10,
            department: "Sales",
            summaryDepartment: "Commercial"
          },
          {
            accountLabel: "Revenue",
            summaryAccount: "Net Revenue",
            date: "2026-04-30",
            amount: 5,
            department: "Sales",
            summaryDepartment: "Commercial"
          }
        ]
      );

      expect(repo.sumPlDepartment("Net Revenue", "2026-04-30", "Commercial")).toBe(15);
    } finally {
      cleanup();
    }
  });

  it("replaces BS and operations data and returns exact-date sums", () => {
    const { repo, cleanup } = makeRepo();
    try {
      repo.replaceBs(
        [{ key: "Cash", summary: "Cash" }],
        [{ accountLabel: "Cash", summaryAccount: "Cash", date: "2026-04-30", amount: 100 }]
      );
      repo.replaceOps([
        { metricLabel: "Orders", date: "2026-04-30", channel: "Amazon", amount: 3 },
        { metricLabel: "Orders", date: "2026-04-30", channel: "Amazon", amount: 4 }
      ]);

      expect(repo.sumBs("Cash", "2026-04-30")).toBe(100);
      expect(repo.sumOps("Orders", "2026-04-30", "Amazon")).toBe(7);
      expect(repo.sumOps("Orders", "2026-05-31", "Amazon")).toBe(0);
    } finally {
      cleanup();
    }
  });
});

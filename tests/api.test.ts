import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import request from "supertest";
import { createApp } from "../src/app.js";
import { LocalActRepository } from "../src/db/repository.js";

function makeApp() {
  const dir = mkdtempSync(join(tmpdir(), "localact-api-"));
  const repo = new LocalActRepository(join(dir, "test.sqlite"));
  return {
    app: createApp(repo),
    repo,
    cleanup: () => {
      repo.close();
      rmSync(dir, { recursive: true, force: true });
    }
  };
}

describe("LocalAct API", () => {
  it("reports health", async () => {
    const { app, cleanup } = makeApp();
    try {
      const response = await request(app).get("/health").expect(200);
      expect(response.body).toMatchObject({ ok: true, service: "LocalAct" });
    } finally {
      cleanup();
    }
  });

  it("imports P&L rows and returns lookup values", async () => {
    const { app, cleanup } = makeApp();
    try {
      await request(app)
        .post("/api/import/pl")
        .send({
          plData: [{ "Account label": " Revenue ", Date: "2026-04-30", Amount: 100 }],
          plMapping: [{ "Account label": "Revenue", "Summary Account": "Net Revenue" }]
        })
        .expect(200)
        .expect((response) => {
          expect(response.body).toMatchObject({ ok: true, rowCount: 1 });
        });

      const lookup = await request(app)
        .get("/api/value/pl")
        .query({ summaryAccount: "Net Revenue", date: "2026-04-30" })
        .expect(200);

      expect(lookup.body).toEqual({ value: 100 });
    } finally {
      cleanup();
    }
  });

  it("blocks missing P&L mappings", async () => {
    const { app, cleanup } = makeApp();
    try {
      const response = await request(app)
        .post("/api/import/pl")
        .send({
          plData: [{ "Account label": "COGS", Date: "2026-04-30", Amount: 10 }],
          plMapping: [{ "Account label": "Revenue", "Summary Account": "Net Revenue" }]
        })
        .expect(400);

      expect(response.body.errors).toContain("Missing Summary Account mapping for Account label: COGS");
    } finally {
      cleanup();
    }
  });

  it("imports department P&L using existing P&L mapping", async () => {
    const { app, cleanup } = makeApp();
    try {
      await request(app)
        .post("/api/import/pl")
        .send({
          plData: [{ "Account label": "Revenue", Date: "2026-04-30", Amount: 1 }],
          plMapping: [{ "Account label": "Revenue", "Summary Account": "Net Revenue" }]
        })
        .expect(200);

      await request(app)
        .post("/api/import/pl-department")
        .send({
          plDepartmentData: [{ "Account label": "Revenue", Date: "2026-04-30", Amount: 8, Department: "Sales" }],
          departmentMapping: [{ Department: "Sales", "Summary Department": "Commercial" }]
        })
        .expect(200);

      const lookup = await request(app)
        .get("/api/value/pl-department")
        .query({ summaryAccount: "Net Revenue", date: "2026-04-30", summaryDepartment: "Commercial" })
        .expect(200);

      expect(lookup.body).toEqual({ value: 8 });
    } finally {
      cleanup();
    }
  });
});

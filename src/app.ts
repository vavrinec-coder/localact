import path from "node:path";
import { fileURLToPath } from "node:url";
import cors from "cors";
import express, { type NextFunction, type Request, type Response } from "express";
import multer from "multer";
import { buildBsImport, buildOpsImport, buildPlDepartmentImport, buildPlImport } from "./core/imports.js";
import type { RawRow } from "./core/types.js";
import type { LocalActRepository } from "./db/repository.js";
import { workbookBufferToRows } from "./xlsx.js";

const upload = multer({ storage: multer.memoryStorage() });
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.resolve(__dirname, "../public");

export function createApp(repo: LocalActRepository): express.Express {
  const app = express();
  app.use(cors({ origin: true }));
  app.use(express.json({ limit: "50mb" }));
  app.use(express.static(publicDir));

  app.get("/health", (_req, res) => {
    res.json({ ok: true, service: "LocalAct" });
  });

  app.post("/api/import/pl", upload.any(), async (req, res, next) => {
    try {
    const plData = await rowsFromRequest(req, "plData", "plData");
    const plMapping = await rowsFromRequest(req, "plMapping", "plMapping");
    const result = buildPlImport(plData, plMapping);
    if (result.errors.length > 0) {
      return validationFailure(res, "P&L update blocked", result.errors);
    }
    repo.replacePl(result.mappings ?? [], result.enrichedRows);
    return res.json({ ok: true, message: "P&L updated successfully", rowCount: result.enrichedRows.length });
    } catch (error) {
      return next(error);
    }
  });

  app.post("/api/import/pl-department", upload.any(), async (req, res, next) => {
    try {
    const plDepartmentData = await rowsFromRequest(req, "plDepartmentData", "plDepartmentData");
    const departmentMapping = await rowsFromRequest(req, "departmentMapping", "departmentMapping");
    const result = buildPlDepartmentImport(plDepartmentData, repo.getPlMappings(), departmentMapping);
    if (result.errors.length > 0) {
      return validationFailure(res, "Department update blocked", result.errors);
    }
    repo.replacePlDepartment(result.mappings ?? [], result.enrichedRows);
    return res.json({
      ok: true,
      message: "P&L by Departments updated successfully",
      rowCount: result.enrichedRows.length
    });
    } catch (error) {
      return next(error);
    }
  });

  app.post("/api/import/bs", upload.any(), async (req, res, next) => {
    try {
    const bsData = await rowsFromRequest(req, "bsData", "bsData");
    const bsMapping = await rowsFromRequest(req, "bsMapping", "bsMapping");
    const result = buildBsImport(bsData, bsMapping);
    if (result.errors.length > 0) {
      return validationFailure(res, "Balance Sheet update blocked", result.errors);
    }
    repo.replaceBs(result.mappings ?? [], result.enrichedRows);
    return res.json({ ok: true, message: "Balance Sheet updated successfully", rowCount: result.enrichedRows.length });
    } catch (error) {
      return next(error);
    }
  });

  app.post("/api/import/ops", upload.any(), async (req, res, next) => {
    try {
    const ops = await rowsFromRequest(req, "ops", "ops");
    const result = buildOpsImport(ops);
    if (result.errors.length > 0) {
      return validationFailure(res, "Operations update blocked", result.errors);
    }
    repo.replaceOps(result.rows);
    return res.json({ ok: true, message: "Operations updated successfully", rowCount: result.rows.length });
    } catch (error) {
      return next(error);
    }
  });

  app.get("/api/value/pl", (req, res) => {
    res.json({ value: repo.sumPl(requiredQuery(req, "summaryAccount"), requiredQuery(req, "date")) });
  });

  app.get("/api/value/bs", (req, res) => {
    res.json({ value: repo.sumBs(requiredQuery(req, "summaryAccount"), requiredQuery(req, "date")) });
  });

  app.get("/api/value/pl-department", (req, res) => {
    res.json({
      value: repo.sumPlDepartment(
        requiredQuery(req, "summaryAccount"),
        requiredQuery(req, "date"),
        requiredQuery(req, "summaryDepartment")
      )
    });
  });

  app.get("/api/value/ops", (req, res) => {
    res.json({
      value: repo.sumOps(
        requiredQuery(req, "metricLabel"),
        requiredQuery(req, "product"),
        requiredQuery(req, "date"),
        requiredQuery(req, "channel")
      )
    });
  });

  app.get("/api/cache", (_req, res) => {
    res.json(repo.getFunctionCache());
  });

  app.use((error: unknown, _req: Request, res: Response, _next: NextFunction) => {
    const message = error instanceof Error ? error.message : String(error);
    return res.status(500).json({
      ok: false,
      message: "LocalAct import failed",
      errors: [message]
    });
  });

  return app;
}

function validationFailure(res: Response, message: string, errors: string[]): Response {
  return res.status(400).json({ ok: false, message, errors });
}

function requiredQuery(req: Request, key: string): string {
  return String(req.query[key] ?? "").trim();
}

async function rowsFromRequest(req: Request, bodyKey: string, fileField: string): Promise<RawRow[]> {
  const bodyValue = req.body?.[bodyKey];
  if (Array.isArray(bodyValue)) {
    return bodyValue as RawRow[];
  }
  if (typeof bodyValue === "string" && bodyValue.trim()) {
    return JSON.parse(bodyValue) as RawRow[];
  }

  const files = (req.files ?? []) as Express.Multer.File[];
  const file = files.find((candidate) => candidate.fieldname === fileField);
  return file ? workbookBufferToRows(file.buffer) : [];
}

import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import devCerts from "office-addin-dev-certs";
import { createApp } from "./app.js";
import { LocalActRepository } from "./db/repository.js";

const port = Number(process.env.LOCALACT_PORT || 3000);
const httpPreviewPort = process.env.LOCALACT_HTTP_PORT ? Number(process.env.LOCALACT_HTTP_PORT) : 0;
const dataDir = path.resolve(process.env.LOCALACT_DATA_DIR || "data");
const dbPath = path.join(dataDir, "localact.sqlite");

fs.mkdirSync(dataDir, { recursive: true });

const repo = new LocalActRepository(dbPath);
const app = createApp(repo);
const httpsOptions = await devCerts.getHttpsServerOptions();

const server = https.createServer(httpsOptions, app);
const httpPreviewServer = httpPreviewPort ? http.createServer(app) : undefined;

server.listen(port, () => {
  console.log(`LocalAct is running at https://localhost:${port}`);
  console.log(`SQLite database: ${dbPath}`);
});

httpPreviewServer?.listen(httpPreviewPort, () => {
  console.log(`LocalAct preview is running at http://localhost:${httpPreviewPort}`);
});

process.on("SIGINT", () => shutdown());
process.on("SIGTERM", () => shutdown());

function shutdown() {
  server.close(() => {
    httpPreviewServer?.close(() => {
      repo.close();
      process.exit(0);
    });
    if (!httpPreviewServer) {
      repo.close();
      process.exit(0);
    }
  });
}

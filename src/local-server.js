"use strict";

const http = require("http");
const fs = require("fs/promises");
const path = require("path");

const HOST = process.env.HOST || "127.0.0.1";
const PORT = Number(process.env.PORT || 5173);
const ROOT_DIR = __dirname;
const DATA_DIR = path.join(ROOT_DIR, "data");
const BODY_LIMIT_BYTES = 8 * 1024 * 1024;
const STORAGE_KEY_PATTERN = /^[a-zA-Z0-9_-]+$/;

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".txt": "text/plain; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".webp": "image/webp",
};

function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store",
  });
  res.end(body);
}

function sendText(res, statusCode, text) {
  res.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
    "Content-Length": Buffer.byteLength(text),
    "Cache-Control": "no-store",
  });
  res.end(text);
}

async function readBodyJson(req) {
  const chunks = [];
  let totalBytes = 0;
  for await (const chunk of req) {
    totalBytes += chunk.length;
    if (totalBytes > BODY_LIMIT_BYTES) {
      throw new Error("PAYLOAD_TOO_LARGE");
    }
    chunks.push(chunk);
  }
  if (!chunks.length) {
    return {};
  }
  const raw = Buffer.concat(chunks).toString("utf8");
  try {
    return JSON.parse(raw);
  } catch {
    throw new Error("INVALID_JSON");
  }
}

function resolveStorageFileByKey(key) {
  if (!STORAGE_KEY_PATTERN.test(key)) {
    return "";
  }
  return path.join(DATA_DIR, `${key}.json`);
}

async function ensureDataDir() {
  await fs.mkdir(DATA_DIR, { recursive: true });
}

async function handleApiStorage(req, res, key) {
  const filePath = resolveStorageFileByKey(key);
  if (!filePath) {
    sendJson(res, 400, { ok: false, error: "invalid_storage_key" });
    return;
  }

  if (req.method === "GET") {
    try {
      const raw = await fs.readFile(filePath, "utf8");
      const value = JSON.parse(raw);
      sendJson(res, 200, { ok: true, exists: true, value });
      return;
    } catch (error) {
      if (error && error.code === "ENOENT") {
        sendJson(res, 404, { ok: true, exists: false, value: null });
        return;
      }
      sendJson(res, 500, { ok: false, error: "read_storage_failed" });
      return;
    }
  }

  if (req.method === "POST") {
    let payload;
    try {
      payload = await readBodyJson(req);
    } catch (error) {
      if (error.message === "PAYLOAD_TOO_LARGE") {
        sendJson(res, 413, { ok: false, error: "payload_too_large" });
        return;
      }
      sendJson(res, 400, { ok: false, error: "invalid_json_body" });
      return;
    }

    const value = Object.prototype.hasOwnProperty.call(payload, "value") ? payload.value : null;
    try {
      await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
      sendJson(res, 200, { ok: true });
      return;
    } catch {
      sendJson(res, 500, { ok: false, error: "write_storage_failed" });
      return;
    }
  }

  sendJson(res, 405, { ok: false, error: "method_not_allowed" });
}

function resolveStaticFilePath(pathname) {
  let normalizedPathname = pathname;
  if (normalizedPathname === "/") {
    normalizedPathname = "/index.html";
  }

  const safePath = path.normalize(normalizedPathname).replace(/^(\.\.[/\\])+/, "");
  const absolutePath = path.resolve(ROOT_DIR, `.${safePath}`);
  if (!absolutePath.startsWith(ROOT_DIR)) {
    return "";
  }
  return absolutePath;
}

async function handleStaticFile(res, pathname) {
  const absolutePath = resolveStaticFilePath(pathname);
  if (!absolutePath) {
    sendText(res, 403, "Forbidden");
    return;
  }

  try {
    const stat = await fs.stat(absolutePath);
    if (stat.isDirectory()) {
      sendText(res, 404, "Not Found");
      return;
    }
    const file = await fs.readFile(absolutePath);
    const ext = path.extname(absolutePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": contentType,
      "Content-Length": file.length,
      "Cache-Control": "no-cache",
    });
    res.end(file);
  } catch {
    sendText(res, 404, "Not Found");
  }
}

async function handleRequest(req, res) {
  const requestUrl = new URL(req.url, `http://${req.headers.host || `${HOST}:${PORT}`}`);
  const pathname = requestUrl.pathname;

  if (pathname === "/api/health") {
    sendJson(res, 200, {
      ok: true,
      mode: "local-file-storage",
      dataDir: DATA_DIR,
    });
    return;
  }

  if (pathname.startsWith("/api/storage/")) {
    const key = decodeURIComponent(pathname.slice("/api/storage/".length)).trim();
    await handleApiStorage(req, res, key);
    return;
  }

  await handleStaticFile(res, pathname);
}

async function startServer() {
  await ensureDataDir();
  const server = http.createServer((req, res) => {
    handleRequest(req, res).catch(() => {
      sendJson(res, 500, { ok: false, error: "internal_server_error" });
    });
  });
  server.listen(PORT, HOST, () => {
    // eslint-disable-next-line no-console
    console.log(`物流台账本地服务已启动: http://${HOST}:${PORT}`);
    // eslint-disable-next-line no-console
    console.log(`数据目录: ${DATA_DIR}`);
  });
}

startServer().catch((error) => {
  // eslint-disable-next-line no-console
  console.error("启动失败:", error);
  process.exit(1);
});

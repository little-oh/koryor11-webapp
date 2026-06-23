import http from "node:http";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { randomUUID } from "node:crypto";

const PORT = Number(process.env.PORT || 8787);
const ROOT = path.dirname(fileURLToPath(import.meta.url));
const PUBLIC_DIR = path.join(ROOT, "public");
const RUNS_DIR = path.join(ROOT, "runs");
const DATA_DIR = path.join(ROOT, "data");
const CUSTOMER_DB_PATH = path.join(DATA_DIR, "customer-lists.json");
const SOURCE_CONVERTER = path.join(ROOT, "../ky11_converter.mjs");
const MAX_FILES = 10;
const MAX_CUSTOMER_SLOTS = 5;

await fs.mkdir(RUNS_DIR, { recursive: true });
await fs.mkdir(DATA_DIR, { recursive: true });

const converter = await import(pathToFileURL(SOURCE_CONVERTER).href);

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
  });
  res.end(JSON.stringify(data));
}

function sendText(res, statusCode, text, contentType = "text/plain; charset=utf-8") {
  res.writeHead(statusCode, { "content-type": contentType, "cache-control": "no-store" });
  res.end(text);
}

async function readBody(req, limitBytes = 50 * 1024 * 1024) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > limitBytes) throw new Error("Payload too large");
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString("utf8");
}

function safeName(name) {
  return String(name || "input.xlsx").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
}

function normalizeSlotId(slotId) {
  return String(slotId || "").trim().toLowerCase().replace(/[^a-z0-9_-]/g, "");
}

async function readCustomerDb() {
  try {
    const raw = await fs.readFile(CUSTOMER_DB_PATH, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed.slots) ? parsed.slots.filter((slot) => slot && slot.id) : [];
  } catch {
    return [];
  }
}

async function writeCustomerDb(slots) {
  await fs.writeFile(
    CUSTOMER_DB_PATH,
    JSON.stringify({ slots, updatedAt: new Date().toISOString() }, null, 2),
    "utf8",
  );
}

async function saveBase64ToFile(dir, fileName, base64) {
  const filePath = path.join(dir, safeName(fileName));
  await fs.writeFile(filePath, Buffer.from(base64, "base64"));
  return filePath;
}

async function loadCustomerNamesFromStoredSlots(slotIds) {
  const wanted = new Set((Array.isArray(slotIds) ? slotIds : []).map(normalizeSlotId));
  const slots = await readCustomerDb();
  const names = [];
  for (const slot of slots) {
    if (!wanted.has(normalizeSlotId(slot.id))) continue;
    if (Array.isArray(slot.names)) names.push(...slot.names);
  }
  return [...new Set(names.map((name) => String(name).trim()).filter(Boolean))];
}

async function makeOutputForInput(runDir, fileName, base64, customerNames) {
  const inputPath = path.join(runDir, `input-${safeName(fileName)}`);
  await fs.writeFile(inputPath, Buffer.from(base64, "base64"));
  const outName = `demo-customer-${safeName(fileName).replace(/\.(xlsx|xlsm|xls|csv)$/i, "")}.xlsx`;
  const outputPath = path.join(runDir, outName);
  const result = await converter.convert(inputPath, outputPath, {
    demoRandomNames: true,
    customerNames,
  });
  return {
    inputName: fileName,
    outputName: path.basename(result.outputPath),
    downloadUrl: `/runs/${path.basename(runDir)}/${encodeURIComponent(path.basename(result.outputPath))}`,
  };
}

async function handleCustomerListGet(res) {
  const slots = await readCustomerDb();
  sendJson(res, 200, {
    ok: true,
    maxSlots: MAX_CUSTOMER_SLOTS,
    slots: slots.slice(0, MAX_CUSTOMER_SLOTS),
  });
}

async function handleCustomerListPost(req, res) {
  const body = JSON.parse(await readBody(req));
  const slotId = normalizeSlotId(body.slotId);
  const file = body.file || null;
  if (!slotId) return sendJson(res, 400, { error: "Missing slotId" });
  if (!file?.base64 || !file?.name) return sendJson(res, 400, { error: "Missing file" });

  const tempPath = await saveBase64ToFile(RUNS_DIR, `customer-${slotId}-${safeName(file.name)}`, file.base64);
  const detailed = await converter.loadCustomerNamesDetailed(tempPath);
  const names = detailed.names;
  const slots = await readCustomerDb();
  const nextSlot = {
    id: slotId,
    label: body.label ? String(body.label).trim() : "",
    fileName: file.name,
    originalFileName: file.name,
    names,
    nameCount: names.length,
    usableCount: detailed.stats?.usableCount || names.length,
    detectedSheetName: detailed.stats?.sheetName || "",
    detectedColumnLabel: detailed.stats?.columnLabel || "A",
    sampleNames: Array.isArray(detailed.stats?.sampleNames) ? detailed.stats.sampleNames : [],
    updatedAt: new Date().toISOString(),
  };
  const nextSlots = slots.filter((slot) => normalizeSlotId(slot.id) !== slotId);
  nextSlots.push(nextSlot);
  nextSlots.sort((a, b) => String(a.id).localeCompare(String(b.id)));
  await writeCustomerDb(nextSlots.slice(0, MAX_CUSTOMER_SLOTS));
  sendJson(res, 200, { ok: true, slot: nextSlot });
}

async function handleCustomerListDelete(res, slotId) {
  const id = normalizeSlotId(slotId);
  const slots = await readCustomerDb();
  const nextSlots = slots.filter((slot) => normalizeSlotId(slot.id) !== id);
  await writeCustomerDb(nextSlots);
  sendJson(res, 200, { ok: true });
}

async function handleApiConvert(req, res) {
  const body = JSON.parse(await readBody(req));
  const files = Array.isArray(body.files) ? body.files : [];
  const customerSlotIds = Array.isArray(body.customerSlotIds) ? body.customerSlotIds : [];

  if (!files.length) return sendJson(res, 400, { error: "No files uploaded" });
  if (files.length > MAX_FILES) return sendJson(res, 400, { error: "Upload up to 10 files at a time" });

  const runId = `${Date.now()}-${randomUUID().slice(0, 8)}`;
  const runDir = path.join(RUNS_DIR, runId);
  await fs.mkdir(runDir, { recursive: true });
  const customerNames = await loadCustomerNamesFromStoredSlots(customerSlotIds);

  const results = [];
  for (const file of files) {
    results.push(await makeOutputForInput(runDir, file.name, file.base64, customerNames));
  }

  sendJson(res, 200, { ok: true, results, runId });
}

async function serveStatic(req, res, pathname) {
  const relative = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.normalize(path.join(PUBLIC_DIR, relative));
  if (!filePath.startsWith(PUBLIC_DIR)) return sendText(res, 403, "Forbidden");
  try {
    const data = await fs.readFile(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const types = {
      ".html": "text/html; charset=utf-8",
      ".js": "application/javascript; charset=utf-8",
      ".css": "text/css; charset=utf-8",
      ".json": "application/json; charset=utf-8",
      ".svg": "image/svg+xml",
    };
    sendText(res, 200, data.toString("utf8"), types[ext] || "application/octet-stream");
  } catch {
    sendText(res, 404, "Not found");
  }
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    if (req.method === "GET" && url.pathname === "/api/customer-lists") return handleCustomerListGet(res);
    if (req.method === "POST" && url.pathname === "/api/customer-lists") return handleCustomerListPost(req, res);
    if (req.method === "DELETE" && url.pathname.startsWith("/api/customer-lists/")) {
      return handleCustomerListDelete(res, url.pathname.slice("/api/customer-lists/".length));
    }
    if (req.method === "POST" && url.pathname === "/api/convert") return handleApiConvert(req, res);
    if (url.pathname.startsWith("/runs/")) {
      const rel = url.pathname
        .slice("/runs/".length)
        .split("/")
        .map((segment) => decodeURIComponent(segment))
        .join(path.sep);
      const filePath = path.normalize(path.join(RUNS_DIR, rel));
      if (!filePath.startsWith(RUNS_DIR)) return sendText(res, 403, "Forbidden");
      try {
        const data = await fs.readFile(filePath);
        res.writeHead(200, {
          "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "content-disposition": `attachment; filename="${path.basename(filePath)}"`,
          "cache-control": "no-store",
        });
        res.end(data);
      } catch {
        sendText(res, 404, "Not found");
      }
      return;
    }
    return serveStatic(req, res, url.pathname);
  } catch (error) {
    sendJson(res, 500, { error: error.message || String(error) });
  }
});

const HOST = process.env.HOST || "0.0.0.0";

server.listen(PORT, HOST, () => {
  console.log(`Webapp ready at http://${HOST}:${PORT}`);
});

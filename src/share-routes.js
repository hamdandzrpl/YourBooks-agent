import { readFile } from "node:fs/promises";
import { validateWorkspace, exportWorkspaceXlsx, importWorkspaceXlsx } from "./workspace-sheets.js";
import { computeWorkbook } from "./formula-engine.js";
import { withSummary } from "./workspace-summary.js";
import { coerceNumberString } from "./share-dataset.js";
import { renderWorkspacePage } from "./workspace-page.js";
import { sanitizeAnalysis } from "./analysis-spec.js";
import { generateWidgets } from "./analysis-ai.js";

const assetCache = new Map();

/** Workspace browser assets: request path → [file in /public, content type]. */
const assetNames = new Map([
  ["/app.js", ["workspace-app.js", "text/javascript; charset=utf-8"]],
  ["/app.css", ["workspace-app.css", "text/css; charset=utf-8"]],
  ["/analysis.js", ["workspace-analysis.js", "text/javascript; charset=utf-8"]],
  ["/analysis.css", ["workspace-analysis.css", "text/css; charset=utf-8"]],
]);

/** Serve a browser asset from /public, cached after the first read. */
async function readAsset(name) {
  if (assetCache.has(name)) return assetCache.get(name);
  const bytes = await readFile(new URL(`../public/${name}`, import.meta.url));
  assetCache.set(name, bytes);
  return bytes;
}

/**
 * Handle public workspace share routes: /w/<name>?k=<token> and its data API.
 * Returns true when the request was handled.
 */
export async function handleWorkspaceShare(request, response, { shareLinks, fileWorkspace, config }) {
  const url = new URL(request.url ?? "/", "http://localhost");
  const match = /^\/w\/([a-z0-9][a-z0-9-]{0,47})(\/.*)?$/.exec(url.pathname);
  if (!match) return false;

  const shareName = match[1];
  const token = url.searchParams.get("k") ?? "";
  const subPath = match[2] ?? "";
  const resolved = shareLinks?.lookup?.(shareName) ?? (token ? shareLinks?.resolve(shareName, token) : null);
  if (!resolved) {
    response.writeHead(404, { "content-type": "text/html; charset=utf-8" }).end("<!doctype html><meta charset=utf-8><title>Tidak ditemukan</title><p>Workspace tidak ditemukan. Minta agent untuk membuatnya.");
    return true;
  }

  const userId = config.panelOwnerNumber ?? "";

  // Browser assets for the workspace page. Cached so repeated loads stay cheap.
  if (request.method === "GET" && assetNames.has(subPath)) {
    const [file, type] = assetNames.get(subPath);
    try {
      const bytes = await readAsset(file);
      response.writeHead(200, { "content-type": type, "cache-control": "public, max-age=300" }).end(bytes);
    } catch {
      response.writeHead(404, { "content-type": "text/plain; charset=utf-8" }).end("Not found");
    }
    return true;
  }

  if (request.method === "GET" && !subPath) {
    const dataset = await readDataset(fileWorkspace, userId, resolved.folder);
    const html = await renderWorkspacePage({ name: shareName, title: resolved.title }, `/w/${shareName}`, dataset);
    response.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store", "x-robots-tag": "noindex" }).end(html);
    return true;
  }
  const sendJson = (status, body) => {
    response.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" }).end(JSON.stringify(body));
  };

  // Dataset backing the shared workspace: <folder>/dataset.json
  if (request.method === "GET" && subPath === "/data") {
    const dataset = await readDataset(fileWorkspace, userId, resolved.folder);
    sendJson(200, dataset);
    return true;
  }

  if (request.method === "POST" && subPath === "/workbook") {
    let body;
    try { body = await readJsonBody(request); } catch { sendJson(400, { error: "Body tidak valid" }); return true; }
    try {
      const workbook = validateWorkspace(body);
      const analysis = sanitizeAnalysis(body.analysis, workbook.sheets.length);
      await writeDataset(fileWorkspace, userId, resolved.folder, workbook, analysis);
      const recalculated = await readDataset(fileWorkspace, userId, resolved.folder);
      sendJson(200, { saved: true, sheets: recalculated.sheets, sheet_names: recalculated.sheets.map((sheet) => sheet.name), analysis: recalculated.analysis ?? null });
    } catch (error) { sendJson(400, { error: error.message }); }
    return true;
  }

  // Ask the AI (or the keyword fallback) for dashboard widgets.
  if (request.method === "POST" && subPath === "/analysis/widgets") {
    let body;
    try { body = await readJsonBody(request); } catch { sendJson(400, { error: "Body tidak valid" }); return true; }
    const prompt = typeof body?.prompt === "string" ? body.prompt.slice(0, 500) : "";
    if (!prompt.trim()) { sendJson(400, { error: "Tulis dulu mau lihat analisis apa." }); return true; }
    const dataset = await readDataset(fileWorkspace, userId, resolved.folder);
    try {
      const result = await generateWidgets({ prompt, dataset, config });
      sendJson(200, { widgets: result.widgets, source: result.source, note: result.note });
    } catch (error) {
      sendJson(500, { error: `Gagal menyusun widget: ${error.message}` });
    }
    return true;
  }

  // Persist the column binding chosen in the Sumber Data panel.
  if (request.method === "POST" && subPath === "/analysis/fields") {
    let body;
    try { body = await readJsonBody(request); } catch { sendJson(400, { error: "Body tidak valid" }); return true; }
    const dataset = await readDataset(fileWorkspace, userId, resolved.folder);
    const current = dataset.analysis ?? { widgets: [] };
    const next = sanitizeAnalysis({ fields: body?.fields, widgets: current.widgets }, dataset.sheets.length);
    if (!next) { sendJson(400, { error: "Kolom pemasukan atau pengeluaran harus dipilih." }); return true; }
    await writeDataset(fileWorkspace, userId, resolved.folder, dataset, next);
    sendJson(200, { saved: true, analysis: next });
    return true;
  }

  // Persist the widget list as it stands (after add/delete).
  if (request.method === "POST" && subPath === "/analysis/widgets/save") {
    let body;
    try { body = await readJsonBody(request); } catch { sendJson(400, { error: "Body tidak valid" }); return true; }
    const dataset = await readDataset(fileWorkspace, userId, resolved.folder);
    const current = dataset.analysis ?? {};
    const next = sanitizeAnalysis({ fields: body?.fields ?? current.fields, widgets: body?.widgets }, dataset.sheets.length);
    if (!next) { sendJson(400, { error: "Kolom pemasukan atau pengeluaran harus dipilih." }); return true; }
    await writeDataset(fileWorkspace, userId, resolved.folder, dataset, next);
    sendJson(200, { saved: true, analysis: next });
    return true;
  }

  if (request.method === "POST" && subPath === "/rows") {
    let body;
    try { body = await readJsonBody(request); } catch { sendJson(400, { error: "Body tidak valid" }); return true; }
    if (!body || !Array.isArray(body.rows) || body.rows.length > 5000) {
      sendJson(400, { error: "rows must be an array of at most 5000 rows" });
      return true;
    }
    const rows = body.rows
      .filter((row) => Array.isArray(row))
      .map((row) => row.slice(0, 64).map((cell) => String(cell ?? "").slice(0, 2000)));
    const dataset = await readDataset(fileWorkspace, userId, resolved.folder);
    const sheet = dataset.sheets[0];
    const columns = body.columns && Array.isArray(body.columns)
      ? body.columns.slice(0, 256).map((name) => String(name).slice(0, 120))
      : sheet.columns;
    await writeDataset(fileWorkspace, userId, resolved.folder, { title: resolved.title || resolved.name, sheets: [{ ...sheet, columns, rows }] });
    sendJson(200, { saved: true, row_count: rows.length });
    return true;
  }

  if (request.method === "POST" && subPath === "/import") {
    const ct = request.headers["content-type"] ?? "";
    if (!ct.includes("multipart/form-data")) { sendJson(400, { error: "Requires multipart/form-data" }); return true; }
    const boundaryMatch = ct.match(/boundary=(?:"([^"]+)"|([^\s;]+))/);
    if (!boundaryMatch) { sendJson(400, { error: "Missing boundary" }); return true; }
    const boundary = boundaryMatch[1] ?? boundaryMatch[2];
    try {
      const body = await readBody(request);
      const parts = parseMultipart(body, boundary);
      const filePart = parts.find((p) => p.name === "file" && p.filename);
      if (!filePart || !filePart.data || filePart.data.length === 0) {
        sendJson(400, { error: "No XLSX file uploaded" }); return true;
      }
      const dataset = await importWorkspaceXlsx(Buffer.from(filePart.data), resolved.title);
      await writeDataset(fileWorkspace, userId, resolved.folder, dataset);
      sendJson(200, { imported: dataset.sheets.map((s) => s.name), sheet_count: dataset.sheets.length, row_count: dataset.sheets.reduce((sum, s) => sum + s.rows.length, 0) });
    } catch (error) {
      sendJson(422, { error: `Import failed: ${error.message}` });
    }
    return true;
  }

  if (request.method === "GET" && subPath === "/export") {
    const dataset = await readDataset(fileWorkspace, userId, resolved.folder);
    const bytes = await exportWorkspaceXlsx(dataset);
    const safeName = `${shareName.replace(/[^a-z0-9-]/g, "_")}.xlsx`;
    response.writeHead(200, {
      "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "content-disposition": `attachment; filename="${safeName}"`,
      "cache-control": "no-store",
    }).end(bytes);
    return true;
  }

  sendJson(404, { error: "Not found" });
  return true;
}

const datasetPath = (folder) => `${folder}/dataset.json`.replace(/^\/+/, "");

function coerceCell(value) {
  if (typeof value !== "string") return value;
  const num = coerceNumberString(value);
  return num !== null && Number.isFinite(num) ? num : value;
}

async function readDataset(fileWorkspace, userId, folder) {
  const files = await fileWorkspace.listFiles(userId, folder);
  const existing = files.find((file) => file.path === datasetPath(folder));
  const fallback = { title: folder.split("/").pop(), sheets: [{ name: "Sheet1", columns: ["Kolom A", "Kolom B"], rows: [] }] };
  if (!existing) return fallback;
  try {
    const parsed = JSON.parse(await fileWorkspace.readTextFile(userId, existing.id));
    // Coerce legacy string numbers so the grid and formulas see real numbers.
    const rows = Array.isArray(parsed.rows) ? parsed.rows : [];
    for (const sheet of Array.isArray(parsed.sheets) ? parsed.sheets : []) {
      if (Array.isArray(sheet.rows)) {
        sheet.rows = sheet.rows.map((row) => Array.isArray(row) ? row.map(coerceCell) : row);
      }
    }
    if (rows.length) {
      parsed.rows = rows.map((row) => Array.isArray(row) ? row.map(coerceCell) : row);
    }
    const normalized = validateWorkspace(parsed);
    const analysis = sanitizeAnalysis(parsed.analysis, normalized.sheets.length);
    const withBlock = withSummary(normalized);
    return analysis ? { ...withBlock, analysis } : withBlock;
  } catch {
    return fallback;
  }
}

async function writeDataset(fileWorkspace, userId, folder, dataset, analysis) {
  const path = datasetPath(folder);
  dataset = validateWorkspace(dataset);
  const recalculated = computeWorkbook(dataset);
  const payload = { ...withSummary(recalculated), updated_at: new Date().toISOString() };
  // The analysis block lives beside the sheets, never inside them, so the
  // XLSX export stays a plain spreadsheet.
  if (analysis) payload.analysis = analysis;
  const bytes = Buffer.from(JSON.stringify(payload, null, 2), "utf8");
  const files = await fileWorkspace.listFiles(userId, folder);
  const existing = files.find((file) => file.path === path);
  if (existing) {
    await fileWorkspace.replaceGeneratedFile(userId, existing.id, bytes);
    return;
  }
  await fileWorkspace.saveGeneratedFile(userId, path, bytes, "application/json");
}

async function readBody(request, maxBytes = 16 * 1024 * 1024) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxBytes) throw new Error("Body too large");
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

function parseMultipart(body, boundary) {
  const parts = [];
  const delim = Buffer.from(`--${boundary}`);
  const end = Buffer.from(`--${boundary}--`);
  let offset = body.indexOf(delim) + delim.length;
  while (offset < body.length) {
    const next = body.indexOf(delim, offset);
    if (next < 0) break;
    const section = body.slice(offset, next);
    offset = next + delim.length;
    if (section.length < 4) continue;
    // Split headers from data
    const headerEnd = section.indexOf(Buffer.from("\r\n\r\n"));
    if (headerEnd < 0) continue;
    const headers = section.slice(0, headerEnd).toString("utf8");
    const dataStart = headerEnd + 4;
    let data = section.slice(dataStart);
    // Trim trailing \r\n
    if (data.length >= 2 && data[data.length - 1] === 0x0a) data = data.slice(0, -1);
    if (data.length >= 1 && data[data.length - 1] === 0x0d) data = data.slice(0, -1);
    const nameMatch = headers.match(/name="([^"]+)"/);
    const filenameMatch = headers.match(/filename="([^"]+)"/);
    parts.push({
      name: nameMatch ? nameMatch[1] : "",
      filename: filenameMatch ? filenameMatch[1] : "",
      data: [...data],
    });
  }
  return parts;
}

async function readJsonBody(request, maxBytes = 4 * 1024 * 1024) {
  const body = await readBody(request, maxBytes);
  return JSON.parse(body.toString("utf8"));
}

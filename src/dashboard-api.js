import { randomBytes, timingSafeEqual } from "node:crypto";
import { readFile } from "node:fs/promises";
import { adoptOriginFromRequest } from "./public-address.js";
import { loadSystemSkills, setSystemSkillEnabled } from "./skills.js";
import { providerHost } from "./provider.js";

const sessions = new Map();
export const dashboardSessions = sessions;
const loginFailures = new Map();
const sessionLifetimeMs = 12 * 60 * 60 * 1000;
const loginWindowMs = 15 * 60 * 1000;
const maxLoginFailures = 5;

function sendJson(response, status, body) {
  response.writeHead(status, {
    "cache-control": "no-store",
    "content-type": "application/json; charset=utf-8",
  });
  response.end(JSON.stringify(body));
}

function safeEqual(left, right) {
  if (typeof left !== "string" || typeof right !== "string" || !left || !right) return false;
  const supplied = Buffer.from(left);
  const expected = Buffer.from(right);
  return supplied.length === expected.length && timingSafeEqual(supplied, expected);
}

function requestToken(request) {
  const authorization = request.headers.authorization ?? "";
  if (authorization.startsWith("Bearer ")) return authorization.slice(7);
  const cookie = request.headers.cookie ?? "";
  return cookie.match(/(?:^|;\s*)dashboard_session=([^;]+)/)?.[1] ?? "";
}

function isAuthorized(request, config) {
  const token = requestToken(request);
  if (safeEqual(token, config.panelApiToken)) return true;
  const session = sessions.get(token);
  if (!session) return false;
  if (session.expiresAt <= Date.now()) {
    sessions.delete(token);
    return false;
  }
  return true;
}

/** Exported so the server can gate origin discovery on a real login: an
 *  unauthenticated caller must not be able to decide where workspace links
 *  point, because those links carry a bearer token. */
export function isDashboardAuthorized(request, config) {
  if (!config.panelPassword && !config.panelApiToken) return false;
  return isAuthorized(request, config);
}

function loginKey(request) {
  return request.socket?.remoteAddress ?? "unknown";
}

async function readJsonBody(request, maxBytes = 4096) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxBytes) throw new Error(`Request body exceeds ${maxBytes} bytes`);
    chunks.push(chunk);
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw new Error("Request body must be valid JSON");
  }
}

async function readBinaryBody(request, maxBytes = 128 * 1024 * 1024) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxBytes) throw new Error("Uploaded backup exceeds the 128 MB limit");
    chunks.push(chunk);
  }
  return Buffer.concat(chunks, size);
}

export async function handleDashboardApi(request, response, config, services = {}) {
  const url = new URL(request.url ?? "/", "http://localhost");
  const pathname = url.pathname;
  const skillMatch = /^\/api\/skills\/([a-z0-9]+(?:-[a-z0-9]+)*)$/.exec(pathname);
  const isSkillsEndpoint = pathname === "/api/skills" || skillMatch;
  const isRuntimeEndpoint = pathname === "/api/runtime";
  const isWorkflowEndpoint = request.method === "GET" && pathname === "/api/workflow/status";
  const isFileToolsEndpoint = pathname === "/api/mcp/files/tools";
  const isFileCallEndpoint = pathname === "/api/mcp/files/call";
  const isBackupEndpoint = request.method === "POST" && pathname === "/api/backup/export";
  const isBackupImportEndpoint = request.method === "POST" && pathname === "/api/backup/import";
  const isLoginEndpoint = request.method === "POST" && pathname === "/api/auth/login";
  const isSessionEndpoint = pathname === "/api/auth/session";
  const isLogoutEndpoint = request.method === "POST" && pathname === "/api/auth/logout";
  const isBooksEndpoint = pathname === "/api/books";
  const isFinanceEndpoint = pathname === "/api/finance";
  const isFinanceExportEndpoint = pathname === "/api/finance/export";
  const isFilesEndpoint = request.method === "GET" && pathname === "/api/files";
  const fileDownloadMatch = request.method === "GET" && /^\/api\/files\/([A-Za-z0-9-]+)\/download$/.exec(pathname);
  const isExcelToolsEndpoint = pathname === "/api/mcp/excel/tools";
  const isExcelCallEndpoint = pathname === "/api/mcp/excel/call";
  const shareLinkMatch = /^\/api\/share-links\/([a-z0-9][a-z0-9-]{0,47})$/.exec(pathname);
  const isShareLinksEndpoint = pathname === "/api/share-links" || shareLinkMatch;
  if (!isSkillsEndpoint && !isRuntimeEndpoint && !isWorkflowEndpoint && !isFileToolsEndpoint && !isFileCallEndpoint && !isBackupEndpoint && !isBackupImportEndpoint
    && !isLoginEndpoint && !isSessionEndpoint && !isLogoutEndpoint && !isBooksEndpoint && !isFinanceEndpoint
    && !isFinanceExportEndpoint && !isFilesEndpoint && !fileDownloadMatch
    && !isExcelToolsEndpoint && !isExcelCallEndpoint && !isShareLinksEndpoint) return false;

  if (isLoginEndpoint) {
    if (!config.panelPassword && !config.panelApiToken) {
      sendJson(response, 503, { error: "PANEL_PASSWORD is not configured" });
      return true;
    }
    const key = loginKey(request);
    const failureState = loginFailures.get(key);
    if (failureState && failureState.resetAt > Date.now() && failureState.count >= maxLoginFailures) {
      sendJson(response, 429, { error: "Too many failed login attempts. Try again in 15 minutes." });
      return true;
    }
    const body = await readJsonBody(request, 2048);
    const suppliedPassword = typeof body?.password === "string" ? body.password : "";
    if (!safeEqual(suppliedPassword, config.panelPassword) && !safeEqual(suppliedPassword, config.panelApiToken)) {
      const current = failureState?.resetAt > Date.now() ? failureState : { count: 0, resetAt: Date.now() + loginWindowMs };
      current.count += 1;
      loginFailures.set(key, current);
      sendJson(response, current.count >= maxLoginFailures ? 429 : 401, {
        error: current.count >= maxLoginFailures ? "Too many failed login attempts. Try again in 15 minutes." : "Password is incorrect",
      });
      return true;
    }
    loginFailures.delete(key);
    const token = randomBytes(32).toString("base64url");
    const expiresAt = Date.now() + sessionLifetimeMs;
    sessions.set(token, { expiresAt });
    // The owner just proved who they are, and the address they typed is the one
    // that reaches this server: use it for workspace links from now on.
    adoptOriginFromRequest(config, request);
    response.writeHead(200, {
      "cache-control": "no-store",
      "content-type": "application/json; charset=utf-8",
      "set-cookie": `dashboard_session=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${Math.floor(sessionLifetimeMs / 1000)}`,
    });
    response.end(JSON.stringify({ token, expiresAt }));
    return true;
  }

  if (!config.panelPassword && !config.panelApiToken) {
    sendJson(response, 503, { error: "Panel authentication is not configured" });
    return true;
  }
  if (!isAuthorized(request, config)) {
    sendJson(response, 401, { error: "Panel authentication required" });
    return true;
  }

  try {
    if (request.method === "GET" && isSessionEndpoint) {
      sendJson(response, 200, { authenticated: true });
      return true;
    }
    if (isLogoutEndpoint) {
      sessions.delete(requestToken(request));
      response.writeHead(200, {
        "cache-control": "no-store",
        "content-type": "application/json; charset=utf-8",
        "set-cookie": "dashboard_session=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0",
      });
      response.end(JSON.stringify({ loggedOut: true }));
      return true;
    }
    if (request.method === "GET" && pathname === "/api/skills") {
      sendJson(response, 200, { skills: loadSystemSkills(config.skillsDirectory).map((skill) => ({
        id: skill.id,
        name: skill.name,
        description: skill.description,
        version: skill.version,
        enabled: skill.enabled,
        instructions: skill.instructions,
      })) });
      return true;
    }
    if (request.method === "GET" && isRuntimeEndpoint) {
      sendJson(response, 200, {
        model: config.aiModel,
        reasoningEffort: config.aiReasoningEffort,
        providerHost: providerHost(config.chatCompletionsUrl),
        timeZone: config.timeZone,
        currency: config.defaultCurrency,
        port: config.port,
        nodeVersion: process.version,
        routerUrl: config.routerUrl || null,
      });
      return true;
    }
    if (isWorkflowEndpoint) {
      if (!config.panelOwnerNumber || !services.workflowLive) {
        sendJson(response, 503, { error: "Live workflow is not enabled" });
        return true;
      }
      sendJson(response, 200, services.workflowLive.snapshot());
      return true;
    }
    if (request.method === "GET" && isBooksEndpoint) {
      if (!config.panelOwnerNumber || !services.financeStore) {
        sendJson(response, 503, { error: "Finance books are not enabled" });
        return true;
      }
      const owner = config.panelOwnerNumber;
      const books = services.financeStore.listBooks(owner);
      sendJson(response, 200, { activeBook: services.financeStore.getActiveBook(owner), books });
      return true;
    }
    if (request.method === "POST" && isBooksEndpoint) {
      if (!config.panelOwnerNumber || !services.financeStore) {
        sendJson(response, 503, { error: "Finance books are not enabled" });
        return true;
      }
      const body = await readJsonBody(request, 4096);
      const owner = config.panelOwnerNumber;
      if (body?.action === "create" && typeof body.name === "string") {
        const book = services.financeStore.createBook(owner, body.name);
        sendJson(response, 200, { book });
        return true;
      }
      if (body?.action === "select" && typeof body.book_id === "string") {
        const book = services.financeStore.selectBook(owner, body.book_id);
        sendJson(response, 200, { book, activeBook: book });
        return true;
      }
      if (body?.action === "move_transaction" && typeof body.transaction_id === "string"
        && typeof body.target_book_id === "string") {
        const moved = services.financeStore.moveTransaction(owner, body.transaction_id, body.target_book_id);
        sendJson(response, 200, { transaction: moved });
        return true;
      }
      sendJson(response, 400, { error: "action must be create, select, or move_transaction with valid fields" });
      return true;
    }
    if (request.method === "GET" && isFinanceEndpoint) {
      if (!config.panelOwnerNumber || !services.financeStore) {
        sendJson(response, 503, { error: "Finance panel is not enabled" });
        return true;
      }
      const month = url.searchParams.get("month") || undefined;
      const requestedBookId = url.searchParams.get("book_id");
      const book = requestedBookId
        ? services.financeStore.getBook(config.panelOwnerNumber, requestedBookId)
        : services.financeStore.getActiveBook(config.panelOwnerNumber);
      if (!book) {
        sendJson(response, 404, { error: "Book not found in this user's workspace" });
        return true;
      }
      sendJson(response, 200, {
        summary: services.financeStore.getSummary(config.panelOwnerNumber, month, book.id),
        transactions: services.financeStore.listRecentTransactions(config.panelOwnerNumber, 25, book.id),
      });
      return true;
    }
    if (request.method === "GET" && isFinanceExportEndpoint) {
      if (!config.panelOwnerNumber || !services.financeStore) {
        sendJson(response, 503, { error: "Finance panel is not enabled" });
        return true;
      }
      const requestedBookId = url.searchParams.get("book_id");
      const book = requestedBookId
        ? services.financeStore.getBook(config.panelOwnerNumber, requestedBookId)
        : services.financeStore.getActiveBook(config.panelOwnerNumber);
      if (!book) {
        sendJson(response, 404, { error: "Book not found in this user's workspace" });
        return true;
      }
      const filename = `${book.name.replace(/[^a-zA-Z0-9_-]/g, "_")}-transactions.csv`;
      response.writeHead(200, {
        "cache-control": "no-store",
        "content-type": "text/csv; charset=utf-8",
        "content-disposition": `attachment; filename="${filename}"`,
      });
      response.end(services.financeStore.exportCsv(config.panelOwnerNumber, book.id));
      return true;
    }
    if (isFilesEndpoint) {
      if (!config.panelOwnerNumber || !services.fileWorkspace) {
        sendJson(response, 503, { error: "File workspace is not enabled" });
        return true;
      }
      const folder = url.searchParams.get("folder") ?? "";
      const allFiles = url.searchParams.get("view") === "all";
      const [entries, folders] = allFiles
        ? await Promise.all([
          services.fileWorkspace.listFiles(config.panelOwnerNumber).then((files) => files.map((file) => ({ ...file, type: "file" }))),
          services.fileWorkspace.listDirectory(config.panelOwnerNumber, "").then((items) => items.filter((item) => item.type === "folder")),
        ])
        : [await services.fileWorkspace.listDirectory(config.panelOwnerNumber, folder), []];
      sendJson(response, 200, { folder, entries, folders });
      return true;
    }
    if (isBackupEndpoint) {
      if (!config.panelOwnerNumber || !services.backupService) {
        sendJson(response, 503, { error: "Backup is not enabled" });
        return true;
      }
      const body = await readJsonBody(request, 4096);
      const backup = await services.backupService.createBackup(config.panelOwnerNumber, body?.components);
      response.writeHead(200, {
        "cache-control": "no-store",
        "content-type": "application/zip",
        "content-disposition": `attachment; filename="${backup.filename}"`,
        "content-length": backup.bytes.length,
        "x-content-type-options": "nosniff",
      });
      response.end(backup.bytes);
      return true;
    }
    if (isBackupImportEndpoint) {
      if (!config.panelOwnerNumber || !services.backupService) {
        sendJson(response, 503, { error: "Backup is not enabled" });
        return true;
      }
      if (!String(request.headers["content-type"] ?? "").toLowerCase().includes("application/zip")) {
        sendJson(response, 415, { error: "Upload a ZIP backup using application/zip content type" });
        return true;
      }
      const components = (url.searchParams.get("components") ?? "").split(",").filter(Boolean);
      const bytes = await readBinaryBody(request);
      const importTask = () => services.backupService.importBackup(config.panelOwnerNumber, bytes, components);
      const result = services.backupQueue ? await services.backupQueue.enqueue(importTask) : await importTask();
      sendJson(response, 200, result);
      return true;
    }
    if (fileDownloadMatch) {
      if (!config.panelOwnerNumber || !services.fileWorkspace) {
        sendJson(response, 503, { error: "File workspace is not enabled" });
        return true;
      }
      const file = services.fileWorkspace.getFile(config.panelOwnerNumber, fileDownloadMatch[1]);
      if (!file) {
        sendJson(response, 404, { error: "File not found" });
        return true;
      }
      const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
      const mimeType = String(file.mimeType || "application/octet-stream").replace(/[\r\n]/g, "");
      response.writeHead(200, {
        "cache-control": "no-store",
        "content-type": mimeType,
        "content-disposition": `attachment; filename="${safeName}"`,
        "x-content-type-options": "nosniff",
      });
      response.end(await readFile(await services.fileWorkspace.getFilePath(config.panelOwnerNumber, file.id)));
      return true;
    }
    if (request.method === "GET" && isFileToolsEndpoint) {
      if (!config.panelOwnerNumber || !services.filesystemMcp) {
        sendJson(response, 503, { error: "Filesystem MCP is not enabled" });
        return true;
      }
      sendJson(response, 200, { tools: services.filesystemMcp.definitions.map((tool) => {
        const definition = structuredClone(tool.function);
        if (definition.name === "mcp_filesystem_workspace_files") {
          definition.parameters.properties.action.enum = definition.parameters.properties.action.enum.filter((action) => action !== "send");
        }
        return definition;
      }) });
      return true;
    }
    if (request.method === "POST" && isFileCallEndpoint) {
      if (!config.panelOwnerNumber) {
        sendJson(response, 503, { error: "PANEL_OWNER_NUMBER is not configured" });
        return true;
      }
      if (!services.filesystemMcp) {
        sendJson(response, 503, { error: "Filesystem MCP is unavailable" });
        return true;
      }
      const body = await readJsonBody(request, 4 * 1024 * 1024);
      if (!body || typeof body.name !== "string" || !body.arguments || typeof body.arguments !== "object" || Array.isArray(body.arguments)) {
        sendJson(response, 400, { error: "name and arguments are required" });
        return true;
      }
      if (body.name === "mcp_filesystem_workspace_files" && body.arguments.action === "send") {
        sendJson(response, 400, { error: "Sending files from the panel is not supported" });
        return true;
      }
      const result = await services.filesystemMcp.call(body.name, body.arguments);
      sendJson(response, 200, { result });
      return true;
    }
    if (request.method === "GET" && isExcelToolsEndpoint) {
      if (!config.panelOwnerNumber || !services.excelMcp || !services.fileWorkspace) {
        sendJson(response, 503, { error: "Excel MCP is not enabled" });
        return true;
      }
      const source = await services.excelMcp.connect({ fileWorkspace: services.fileWorkspace, userId: config.panelOwnerNumber });
      sendJson(response, 200, { tools: source.definitions.map((tool) => tool.function) });
      return true;
    }
    if (request.method === "POST" && isExcelCallEndpoint) {
      if (!config.panelOwnerNumber || !services.excelMcp || !services.fileWorkspace) {
        sendJson(response, 503, { error: "Excel MCP is not enabled" });
        return true;
      }
      const body = await readJsonBody(request, 128 * 1024);
      if (!body || typeof body.name !== "string" || !body.arguments || typeof body.arguments !== "object" || Array.isArray(body.arguments)) {
        sendJson(response, 400, { error: "name and arguments are required" });
        return true;
      }
      const source = await services.excelMcp.connect({ fileWorkspace: services.fileWorkspace, userId: config.panelOwnerNumber });
      const result = await source.call(body.name, body.arguments);
      sendJson(response, 200, { result });
      return true;
    }
    if (request.method === "GET" && isShareLinksEndpoint) {
      if (!config.panelOwnerNumber || !services.shareLinks) {
        sendJson(response, 503, { error: "Share links are not enabled" });
        return true;
      }
      sendJson(response, 200, { links: services.shareLinks.list(config.panelOwnerNumber) });
      return true;
    }
    if (request.method === "POST" && shareLinkMatch) {
      if (!config.panelOwnerNumber || !services.shareLinks) {
        sendJson(response, 503, { error: "Share links are not enabled" });
        return true;
      }
      const rotated = services.shareLinks.rotateToken(config.panelOwnerNumber, shareLinkMatch[1]);
      sendJson(response, 200, rotated);
      return true;
    }
    if (request.method === "DELETE" && shareLinkMatch) {
      if (!config.panelOwnerNumber || !services.shareLinks) {
        sendJson(response, 503, { error: "Share links are not enabled" });
        return true;
      }
      sendJson(response, 200, services.shareLinks.revoke(config.panelOwnerNumber, shareLinkMatch[1]));
      return true;
    }
    if (request.method === "PATCH" && skillMatch) {
      const body = await readJsonBody(request);
      if (typeof body.enabled !== "boolean") {
        sendJson(response, 400, { error: "enabled must be a boolean" });
        return true;
      }
      const skill = setSystemSkillEnabled(skillMatch[1], body.enabled, config.skillsDirectory);
      sendJson(response, 200, { skill: {
        id: skill.id,
        name: skill.name,
        description: skill.description,
        version: skill.version,
        enabled: skill.enabled,
        instructions: skill.instructions,
      } });
      return true;
    }
    sendJson(response, 405, { error: "Method not allowed" });
    return true;
  } catch (error) {
    const status = error.message === "Skill not found" ? 404 : 400;
    sendJson(response, status, { error: error.message });
    return true;
  }
}
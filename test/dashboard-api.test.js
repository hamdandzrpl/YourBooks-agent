import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";
import test from "node:test";
import { handleDashboardApi } from "../src/dashboard-api.js";
import { FinanceStore } from "../src/finance/store.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { loadSystemSkills } from "../src/skills.js";
import { WorkflowLiveState } from "../src/workflow-live.js";

const token = "dashboard-test-token-that-is-at-least-32-chars";

function createRequest(method, url, body, authorization = `Bearer ${token}`) {
  const request = Readable.from(body === undefined ? [] : [Buffer.from(JSON.stringify(body))]);
  request.method = method;
  request.url = url;
  request.headers = { authorization };
  request.socket = { remoteAddress: "127.0.0.1" };
  return request;
}

async function callApi(request, config, services) {
  const response = {
    writeHead(status, headers) { this.status = status; this.headers = headers; },
    end(body) { this.body = body; },
  };
  assert.equal(await handleDashboardApi(request, response, config, services), true);
  return { ...response, json: JSON.parse(response.body) };
}

function createSkillDirectory() {
  const directory = mkdtempSync(join(tmpdir(), "dashboard-skills-"));
  writeFileSync(join(directory, "sample.json"), JSON.stringify({
    id: "sample-skill",
    name: "Sample Skill",
    description: "A test skill",
    version: "1.0.0",
    enabled: true,
    instructions: ["Follow the test instruction."],
  }));
  return directory;
}

test("dashboard API requires a configured token and authenticates requests", async () => {
  const unconfigured = await callApi(createRequest("GET", "/api/skills"), { panelApiToken: "" });
  assert.equal(unconfigured.status, 503);

  const unauthorized = await callApi(createRequest("GET", "/api/skills", undefined, "Bearer wrong-token"), {
    panelApiToken: token,
  });
  assert.equal(unauthorized.status, 401);
});

test("backup endpoint requires authentication and returns a selected ZIP archive", async () => {
  const owner = "6281234567890";
  const config = { panelApiToken: token, panelOwnerNumber: owner };
  const calls = [];
  const services = {
    backupService: {
      async createBackup(userId, components) {
        calls.push({ userId, components });
        return { filename: "agent-backup-test.zip", bytes: Buffer.from("zip-data") };
      },
    },
  };

  const denied = await callApi(createRequest("POST", "/api/backup/export", { components: ["files"] }, "Bearer invalid"), config, services);
  assert.equal(denied.status, 401);

  const response = {
    writeHead(status, headers) { this.status = status; this.headers = headers; },
    end(body) { this.body = body; },
  };
  await handleDashboardApi(createRequest("POST", "/api/backup/export", { components: ["ai", "files"] }), response, config, services);
  assert.equal(response.status, 200);
  assert.equal(response.headers["content-type"], "application/zip");
  assert.match(response.headers["content-disposition"], /agent-backup-test\.zip/);
  assert.deepEqual(response.body, Buffer.from("zip-data"));
  assert.deepEqual(calls, [{ userId: owner, components: ["ai", "files"] }]);
});

test("backup import endpoint accepts a raw ZIP and serializes restoration", async () => {
  const owner = "6281234567890";
  const calls = [];
  const config = { panelApiToken: token, panelOwnerNumber: owner };
  const services = {
    backupService: {
      async importBackup(userId, bytes, components) {
        calls.push({ userId, bytes: Buffer.from(bytes), components });
        return { restored: { files: { files: 1 } }, safetyBackup: "before-import.zip", restartRequired: false };
      },
    },
    backupQueue: {
      async enqueue(task) { calls.push({ queued: true }); return task(); },
    },
  };
  const request = Readable.from([Buffer.from("zip-payload")]);
  request.method = "POST";
  request.url = "/api/backup/import?components=files,ai";
  request.headers = { authorization: `Bearer ${token}`, "content-type": "application/zip" };
  request.socket = { remoteAddress: "127.0.0.1" };
  const response = {
    writeHead(status, headers) { this.status = status; this.headers = headers; },
    end(body) { this.body = body; },
  };

  await handleDashboardApi(request, response, config, services);
  assert.equal(response.status, 200);
  assert.deepEqual(JSON.parse(response.body), {
    restored: { files: { files: 1 } },
    safetyBackup: "before-import.zip",
    restartRequired: false,
  });
  assert.deepEqual(calls, [
    { queued: true },
    { userId: owner, bytes: Buffer.from("zip-payload"), components: ["files", "ai"] },
  ]);
});

test("dashboard API exposes safe runtime data and persists skill toggles", async (t) => {
  const skillsDirectory = createSkillDirectory();
  t.after(() => rmSync(skillsDirectory, { recursive: true, force: true }));
  const config = {
    panelApiToken: token,
    skillsDirectory,
    aiApiKey: "must-not-be-returned",
    aiModel: "vendor/model",
    aiReasoningEffort: "low",
    timeZone: "Asia/Jakarta",
    defaultCurrency: "IDR",
    port: 3000,
  };

  const runtime = await callApi(createRequest("GET", "/api/runtime"), config);
  assert.equal(runtime.status, 200);
  assert.deepEqual(runtime.json, {
    model: "vendor/model",
    reasoningEffort: "low",
    providerHost: "",
    timeZone: "Asia/Jakarta",
    currency: "IDR",
    port: 3000,
    nodeVersion: process.version,
    routerUrl: null,
  });
  assert.equal(JSON.stringify(runtime.json).includes(config.aiApiKey), false);

  const skills = await callApi(createRequest("GET", "/api/skills"), config);
  assert.equal(skills.json.skills[0].enabled, true);
  const updated = await callApi(createRequest("PATCH", "/api/skills/sample-skill", { enabled: false }), config);
  assert.equal(updated.json.skill.enabled, false);
  assert.equal(loadSystemSkills(skillsDirectory)[0].enabled, false);
  assert.equal(JSON.parse(readFileSync(join(skillsDirectory, "sample.json"), "utf8")).enabled, false);
});

test("dashboard skill API validates toggle payloads and IDs", async (t) => {
  const skillsDirectory = createSkillDirectory();
  t.after(() => rmSync(skillsDirectory, { recursive: true, force: true }));
  const config = { panelApiToken: token, skillsDirectory };
  const invalidBody = await callApi(createRequest("PATCH", "/api/skills/sample-skill", { enabled: "false" }), config);
  assert.equal(invalidBody.status, 400);
  const missingSkill = await callApi(createRequest("PATCH", "/api/skills/missing-skill", { enabled: false }), config);
  assert.equal(missingSkill.status, 404);
});

test("dashboard filesystem MCP routes scope file calls", async () => {
  const calls = [];
  const config = {
    panelApiToken: token,
    panelOwnerNumber: "6281234567890",
  };
  const services = {
    filesystemMcp: {
      definitions: [{ function: {
        name: "mcp_filesystem_workspace_files",
        parameters: { properties: { action: { enum: ["list", "send"] } } },
      } }],
      async call(name, args) {
        calls.push({ kind: "files", name, args });
        return { content: "[]", truncated: false, isError: false };
      },
    },
  };

  const fileToolsResponse = await callApi(createRequest("GET", "/api/mcp/files/tools"), config, services);
  assert.equal(fileToolsResponse.status, 200);
  assert.deepEqual(fileToolsResponse.json.tools[0].parameters.properties.action.enum, ["list"]);

  const fileCall = await callApi(createRequest("POST", "/api/mcp/files/call", {
    name: "mcp_filesystem_workspace_files",
    arguments: { action: "list", folder: "inbox" },
  }), config, services);
  assert.equal(fileCall.status, 200);
  assert.deepEqual(calls[0], {
    kind: "files",
    name: "mcp_filesystem_workspace_files",
    args: { action: "list", folder: "inbox" },
  });

  const blockedSend = await callApi(createRequest("POST", "/api/mcp/files/call", {
    name: "mcp_filesystem_workspace_files",
    arguments: { action: "send", file_id: "F-12345678" },
  }), config, services);
  assert.equal(blockedSend.status, 400);

});

test("dashboard password login issues an expiring session that can be revoked", async () => {
  const config = { panelPassword: "1122334455" };
  const invalid = await callApi(createRequest("POST", "/api/auth/login", { password: "wrong" }, ""), config);
  assert.equal(invalid.status, 401);

  const login = await callApi(createRequest("POST", "/api/auth/login", { password: "1122334455" }, ""), config);
  assert.equal(login.status, 200);
  assert.equal(typeof login.json.token, "string");
  assert.ok(login.json.expiresAt > Date.now());

  const authorization = `Bearer ${login.json.token}`;
  const session = await callApi(createRequest("GET", "/api/auth/session", undefined, authorization), config);
  assert.deepEqual(session.json, { authenticated: true });

  const logout = await callApi(createRequest("POST", "/api/auth/logout", {}, authorization), config);
  assert.equal(logout.status, 200);
  const revoked = await callApi(createRequest("GET", "/api/auth/session", undefined, authorization), config);
  assert.equal(revoked.status, 401);
});

test("dashboard password login rate-limits repeated failures", async () => {
  const config = { panelPassword: "1122334455" };
  let lastResponse;
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const request = createRequest("POST", "/api/auth/login", { password: "wrong" }, "");
    request.socket.remoteAddress = "192.0.2.44";
    lastResponse = await callApi(request, config);
  }
  assert.equal(lastResponse.status, 429);
  const blockedRequest = createRequest("POST", "/api/auth/login", { password: "1122334455" }, "");
  blockedRequest.socket.remoteAddress = "192.0.2.44";
  const blocked = await callApi(blockedRequest, config);
  assert.equal(blocked.status, 429);
});

test("finance panel routes expose only the configured owner's summary and export", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "dashboard-finance-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  t.after(() => {
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  const owner = "6281234567890";
  financeStore.recordTransaction(owner, {
    kind: "income",
    amount: 250000,
    currency: "IDR",
    date: "2026-10-18",
    category: "Jasa",
    description: "Pembayaran proyek",
  });
  const config = { panelApiToken: token, panelOwnerNumber: owner };
  const services = { financeStore };

  const summary = await callApi(createRequest("GET", "/api/finance?month=2026-10"), config, services);
  assert.equal(summary.status, 200);
  assert.equal(summary.json.summary.rows[0].total_minor, 250000);
  assert.equal(summary.json.transactions[0].description, "Pembayaran proyek");

  const request = createRequest("GET", "/api/finance/export");
  const response = {
    writeHead(status, headers) { this.status = status; this.headers = headers; },
    end(body) { this.body = body; },
  };
  await handleDashboardApi(request, response, config, services);
  assert.equal(response.status, 200);
  assert.match(response.headers["content-type"], /text\/csv/);
  assert.match(response.body, /Pembayaran proyek/);

  const unavailable = await callApi(createRequest("GET", "/api/finance?month=2026-10"), {
    panelApiToken: token,
    panelOwnerNumber: "6289999999999",
  }, services);
  assert.equal(unavailable.json.transactions.length, 0);
});

test("dashboard book routes create, select, scope, and move transactions with workbook sync", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "dashboard-books-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const owner = "6281234567890";
  const defaultBook = financeStore.getActiveBook(owner);
  const workbookSyncCalls = [];
  const financeWorkbookSync = {
    async createBook(userId, name) {
      const book = financeStore.createBook(userId, name);
      return { ...book, workbook_file_id: `F-${book.id.slice(2)}` };
    },
    async syncBook(userId, bookId) { workbookSyncCalls.push({ userId, bookId }); },
  };
  t.after(() => {
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  const config = { panelApiToken: token, panelOwnerNumber: owner };
  const services = { financeStore, financeWorkbookSync };

  const created = await callApi(createRequest("POST", "/api/books", { action: "create", name: "Bisnis" }), config, services);
  const businessBook = financeStore.getBook(owner, created.json.book.id);
  assert.equal(businessBook.name, "Bisnis");

  const transaction = financeStore.recordTransaction(owner, {
    kind: "income", amount: 120000, date: "2026-10-18", category: "Jasa", description: "Invoice",
  }, { bookId: defaultBook.id });
  const defaultSummary = await callApi(createRequest("GET", `/api/finance?month=2026-10&book_id=${defaultBook.id}`), config, services);
  assert.equal(defaultSummary.json.transactions[0].id, transaction.id);
  const emptyBusiness = await callApi(createRequest("GET", `/api/finance?month=2026-10&book_id=${businessBook.id}`), config, services);
  assert.equal(emptyBusiness.json.transactions.length, 0);

  const selected = await callApi(createRequest("POST", "/api/books", { action: "select", book_id: businessBook.id }), config, services);
  assert.equal(selected.json.activeBook.id, businessBook.id);
  assert.ok(workbookSyncCalls.some((call) => call.bookId === businessBook.id));

  const moved = await callApi(createRequest("POST", "/api/books", {
    action: "move_transaction",
    transaction_id: transaction.id,
    target_book_id: businessBook.id,
  }), config, services);
  assert.equal(moved.json.transaction.from_book_id, defaultBook.id);
  assert.equal(moved.json.transaction.book_id, businessBook.id);
  assert.equal(workbookSyncCalls.filter((call) => [defaultBook.id, businessBook.id].includes(call.bookId)).length >= 3, true);
});

test("dashboard file route lists and downloads only the configured owner's files", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "dashboard-files-"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  t.after(() => {
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  const owner = "6281234567890";
  const file = await fileWorkspace.createTextFile(owner, "inbox/catatan.txt", "private owner note");
  const exportFile = await fileWorkspace.createTextFile(owner, "exports/laporan.csv", "tanggal,jumlah");
  await fileWorkspace.createTextFile("6289999999999", "inbox/secret.txt", "not for panel");
  const config = { panelApiToken: token, panelOwnerNumber: owner };
  const services = { fileWorkspace };

  const listing = await callApi(createRequest("GET", "/api/files?folder=inbox"), config, services);
  assert.equal(listing.status, 200);
  assert.deepEqual(listing.json.entries.map((entry) => entry.name), ["catatan.txt"]);

  const allFiles = await callApi(createRequest("GET", "/api/files?view=all"), config, services);
  assert.deepEqual(allFiles.json.entries.map((entry) => entry.id).sort(), [file.id, exportFile.id].sort());
  assert.ok(allFiles.json.entries.every((entry) => entry.type === "file"));
  assert.deepEqual(allFiles.json.folders.map((entry) => entry.name), ["exports", "inbox"]);

  const request = createRequest("GET", `/api/files/${file.id}/download`);
  const response = {
    writeHead(status, headers) { this.status = status; this.headers = headers; },
    end(body) { this.body = body; },
  };
  await handleDashboardApi(request, response, config, services);
  assert.equal(response.status, 200);
  assert.equal(response.body.toString("utf8"), "private owner note");
  assert.equal(response.headers["x-content-type-options"], "nosniff");
});

test("dashboard proxies Excel MCP tools using the owner workspace", async () => {
  const calls = [];
  const config = { panelApiToken: token, panelOwnerNumber: "6281234567890" };
  const services = {
    fileWorkspace: {},
    excelMcp: {
      async connect(input) {
        calls.push({ kind: "connect", input });
        return {
          definitions: [{ function: { name: "mcp_excel_excel_describe_sheets", parameters: {} } }],
          async call(name, args) {
            calls.push({ kind: "call", name, args });
            return { content: "sheet list", isError: false };
          },
        };
      },
    },
  };

  const tools = await callApi(createRequest("GET", "/api/mcp/excel/tools"), config, services);
  assert.equal(tools.status, 200);
  assert.equal(tools.json.tools[0].name, "mcp_excel_excel_describe_sheets");
  const result = await callApi(createRequest("POST", "/api/mcp/excel/call", {
    name: "mcp_excel_excel_describe_sheets",
    arguments: { file_id: "F-12345678" },
  }), config, services);
  assert.equal(result.status, 200);
  assert.equal(result.json.result.content, "sheet list");
  assert.deepEqual(calls.at(-1), {
    kind: "call",
    name: "mcp_excel_excel_describe_sheets",
    args: { file_id: "F-12345678" },
  });
  assert.equal(calls[0].input.userId, config.panelOwnerNumber);
});

test("workflow status endpoint requires panel auth and returns only safe owner state", async () => {
  const workflowLive = new WorkflowLiveState({ ownerId: "6281234567890" });
  workflowLive.begin("6281234567890");
  workflowLive.update("6281234567890", { type: "model_start" });
  const config = { panelApiToken: token, panelOwnerNumber: "6281234567890" };
  const unauthorized = await callApi(createRequest("GET", "/api/workflow/status", undefined, "Bearer invalid"), config, { workflowLive });
  assert.equal(unauthorized.status, 401);

  const status = await callApi(createRequest("GET", "/api/workflow/status"), config, { workflowLive });
  assert.equal(status.status, 200);
  assert.equal(status.json.status, "running");
  assert.equal(status.json.activeNode, "model");
  assert.doesNotMatch(JSON.stringify(status.json), /sender|message text|api.?key|tool_call_id/i);
});
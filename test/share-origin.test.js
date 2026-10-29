import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";
import test from "node:test";
import { loadConfig } from "../src/config.js";
import { handleDashboardApi } from "../src/dashboard-api.js";
import { FinanceStore } from "../src/finance/store.js";
import { OfficeFiles } from "../src/files/office.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { connectFileWorkspaceMcp } from "../src/files/mcp-server.js";
import { ShareLinkStore } from "../src/share-links.js";
import { runTool } from "../src/tools.js";

const baseEnv = {
  AI_API_KEY: "test-key",
  TIME_ZONE: "Asia/Jakarta",
  DEFAULT_CURRENCY: "IDR",
  PANEL_OWNER_NUMBER: "6281234567890",
};

function services(t) {
  const dataDirectory = mkdtempSync(join(tmpdir(), "share-origin-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const shareLinks = new ShareLinkStore({ dataDirectory });
  const officeFiles = new OfficeFiles(fileWorkspace);
  t.after(() => {
    shareLinks.close();
    fileWorkspace.close();
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  return { dataDirectory, financeStore, fileWorkspace, shareLinks, officeFiles, userId: "6281234567890" };
}

test("an origin typed by the operator is marked explicit and never guessed over", () => {
  assert.equal(loadConfig({ ...baseEnv, PUBLIC_URL: "https://panel.example.test" }).publicOriginExplicit, true);
  assert.equal(loadConfig({ ...baseEnv, PUBLIC_IP: "203.0.113.10" }).publicOriginExplicit, true);
  assert.equal(loadConfig({ ...baseEnv, PUBLIC_HOST: "panel.example.test", SERVER_PORT: "25565" }).publicOriginExplicit, true);
  // Nothing pinned: Pterodactyl sets SERVER_PORT, so no origin is invented.
  const bare = loadConfig({ ...baseEnv, SERVER_PORT: "25565" });
  assert.equal(bare.publicOrigin, null);
  assert.equal(bare.publicOriginExplicit, false);
});

test("a bridge without share links still reports the real cause", async (t) => {
  const { fileWorkspace, officeFiles } = services(t);
  const bridge = await connectFileWorkspaceMcp({
    fileWorkspace, officeFiles, userId: "6281234567890",
    publicOrigin: "https://panel.example.test",
  });
  t.after(() => bridge.close());
  const result = await bridge.call("mcp_filesystem_workspace_files", {
    action: "share_link_create", folder: "documents/x", name: "x",
  });
  assert.equal(result.isError, true);
  assert.match(result.content, /Share links are not enabled on this server/);
});

test("a missing origin names the fix instead of failing silently", async (t) => {
  const { fileWorkspace, officeFiles, shareLinks } = services(t);
  const bridge = await connectFileWorkspaceMcp({
    fileWorkspace, officeFiles, userId: "6281234567890", shareLinks, publicOrigin: "",
  });
  t.after(() => bridge.close());
  const result = await bridge.call("mcp_filesystem_workspace_files", {
    action: "share_link_create", folder: "documents/y", name: "y",
  });
  assert.equal(result.isError, true);
  assert.match(result.content, /PUBLIC_URL/);
  assert.match(result.content, /open the dashboard once/);
});

test("an origin getter lets a long-lived bridge follow a learned address", async (t) => {
  const { fileWorkspace, officeFiles, shareLinks } = services(t);
  let origin = "";
  const bridge = await connectFileWorkspaceMcp({
    fileWorkspace, officeFiles, userId: "6281234567890", shareLinks,
    publicOrigin: () => origin,
  });
  t.after(() => bridge.close());

  const tooEarly = await bridge.call("mcp_filesystem_workspace_files", {
    action: "share_link_create", folder: "documents/z", name: "z",
  });
  assert.equal(tooEarly.isError, true);

  // The owner opens the dashboard at this address; the server adopts it.
  origin = "http://198.51.100.42:25565";
  const created = await bridge.call("mcp_filesystem_workspace_files", {
    action: "share_link_create", folder: "documents/z", name: "z",
  });
  assert.equal(created.isError, false, created.content);
  assert.match(created.content, /http:\/\/198\.51\.100\.42:25565\/w\/z\?k=/);
});

test("create_book returns a workspace link the owner can open", async (t) => {
  const { financeStore, fileWorkspace, shareLinks, officeFiles } = services(t);
  const context = {
    financeStore, fileWorkspace, shareLinks, officeFiles,
    userId: "6281234567890", publicOrigin: "https://panel.example.test", defaultCurrency: "IDR",
  };
  const book = await runTool("create_book", "Asia/Jakarta", { name: "Pertanian" }, context);
  assert.match(book.workspace_url, /^https:\/\/panel\.example\.test\/w\/finance-b-[a-f0-9]{8}\?k=/);
  // No "pending" note: the page really was created.
  assert.equal(book.workspace_note, undefined);
});

test("logging in from the owner's address teaches the server where links live", async (t) => {
  const { financeStore, fileWorkspace, shareLinks, officeFiles } = services(t);
  const config = {
    panelPassword: "rahasia12345", panelApiToken: "", panelOwnerNumber: "6281234567890",
    publicOrigin: null, publicOriginExplicit: false, dashboardUrl: null, healthUrl: null,
  };
  const request = Readable.from([Buffer.from(JSON.stringify({ password: "rahasia12345" }))]);
  request.method = "POST";
  request.url = "/api/auth/login";
  // The address the owner typed in the browser: the only one known to route.
  request.headers = { host: "203.0.113.9:25565" };
  request.socket = { remoteAddress: "127.0.0.1" };
  const response = {
    writeHead(status, headers) { this.status = status; this.headers = headers; },
    end(body) { this.body = body; },
  };
  await handleDashboardApi(request, response, config, {});
  assert.equal(response.status, 200);
  assert.equal(config.publicOrigin, "http://203.0.113.9:25565");

  // A link created afterwards points at that address, so the owner can click it.
  const bridge = await connectFileWorkspaceMcp({
    fileWorkspace, officeFiles, userId: "6281234567890", shareLinks,
    publicOrigin: () => config.publicOrigin,
  });
  t.after(() => bridge.close());
  const created = await bridge.call("mcp_filesystem_workspace_files", {
    action: "share_link_create", folder: "documents/kebun", name: "kebun",
  });
  assert.equal(created.isError, false, created.content);
  assert.match(created.content, /http:\/\/203\.0\.113\.9:25565\/w\/kebun\?k=/);
});

test("a failed login cannot teach the server a fake address", async (t) => {
  const config = {
    panelPassword: "rahasia12345", panelApiToken: "", panelOwnerNumber: "6281234567890",
    publicOrigin: null, publicOriginExplicit: false,
  };
  const request = Readable.from([Buffer.from(JSON.stringify({ password: "tebakan" }))]);
  request.method = "POST";
  request.url = "/api/auth/login";
  request.headers = { host: "attacker.example" };
  request.socket = { remoteAddress: "127.0.0.1" };
  const response = {
    writeHead(status) { this.status = status; },
    end() {},
  };
  await handleDashboardApi(request, response, config, {});
  assert.equal(response.status, 401);
  assert.equal(config.publicOrigin, null);
});

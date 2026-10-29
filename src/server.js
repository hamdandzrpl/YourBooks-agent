import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { createServer, request as httpRequest } from "node:http";
import { readFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { BaileysTransport } from "./baileys.js";
import { loadConfig } from "./config.js";
import { FinanceStore } from "./finance/store.js";
import { FileWorkspace } from "./files/workspace.js";
import { BackupService } from "./backup.js";
import { connectFileWorkspaceMcp } from "./files/mcp-server.js";
import { OfficeFiles } from "./files/office.js";
import { handleDashboardApi, dashboardSessions, isDashboardAuthorized } from "./dashboard-api.js";
import { ShareLinkStore } from "./share-links.js";
import { handleWorkspaceShare } from "./share-routes.js";
import { ExcelMcp } from "./mcp/excel.js";
import { processIncomingMessage } from "./message-handler.js";
import { adoptOwnerOrigin, resolveConsoleAddresses } from "./public-address.js";
import { ProgressTimeline } from "./progress.js";
import { WorkflowLiveState } from "./workflow-live.js";
import { createOutboundRateLimiter, createSerialQueue } from "./queue.js";

const config = loadConfig();
const consoleAddresses = await resolveConsoleAddresses(config);
// Egg environments often expose port but not public host. Use detected node IP;
// PUBLIC_URL remains required override behind domain/reverse-proxy deployments.
const publicOrigin = config.publicOrigin || (consoleAddresses.dashboardUrl ? new URL(consoleAddresses.dashboardUrl).origin : "");
config.publicOrigin = publicOrigin;
if (!config.dashboardUrl && publicOrigin) config.dashboardUrl = `${publicOrigin}/dashboard`;
if (!config.healthUrl && publicOrigin) config.healthUrl = `${publicOrigin}/healthz`;
const financeStore = new FinanceStore(join(config.dataDirectory, "finance.sqlite"));
const fileWorkspace = new FileWorkspace({ dataDirectory: config.dataDirectory, maxFileBytes: config.maxFileBytes });
const backupService = new BackupService({
  dataDirectory: config.dataDirectory,
  financeStore,
  fileWorkspace,
  config,
});
const officeFiles = new OfficeFiles(fileWorkspace, financeStore, config.defaultCurrency);
const excelMcp = new ExcelMcp({ financeStore });
const shareLinks = new ShareLinkStore({ dataDirectory: config.dataDirectory });
const filesystemMcp = config.panelOwnerNumber
  ? await connectFileWorkspaceMcp({
    fileWorkspace, officeFiles, userId: config.panelOwnerNumber, shareLinks,
    publicOrigin: () => config.publicOrigin,
  })
  : null;
const messageQueue = createSerialQueue();
const workflowLive = new WorkflowLiveState({ ownerId: config.panelOwnerNumber });
const maxHistoryMessages = 12;
const dashboardHtml = readFileSync(new URL("../public/panel.html", import.meta.url), "utf8");
let transport;

const sendOutbound = createOutboundRateLimiter({
  minIntervalMs: 5000,
  send: (kind, message) => kind === "text"
    ? transport.sendText(message)
    : transport.sendFile(message),
});
const sendText = (message) => sendOutbound("text", message);
const sendFile = (message) => sendOutbound("file", message);
const sendCsv = ({ to, filename, csv }) => sendFile({
  to,
  filename,
  mimeType: "text/csv",
  bytes: Buffer.from(csv, "utf8"),
});

async function handleMessage(message) {
  workflowLive.begin(message.sender);
  const onWorkflowEvent = (event) => workflowLive.update(message.sender, event);
  const progress = new ProgressTimeline({
    to: message.sender,
    sendText,
    editText: (update) => transport.editText(update),
  });
  try {
    const history = financeStore.getRecentConversation(message.sender, maxHistoryMessages);
    onWorkflowEvent({ type: "history_loaded" });
    await processIncomingMessage({
      message,
      config,
      financeStore,
      fileWorkspace,
      officeFiles,
      shareLinks,
      history,
      sendText,
      sendCsv,
      sendFile,
      onWorkflowEvent,
      progress,
    });
    onWorkflowEvent({ type: "message_complete" });
  } catch (error) {
    onWorkflowEvent({ type: "message_failed" });
    await progress.fail();
    console.error(`Could not prepare reply for WhatsApp message ${message.id}:`, error.message);
    await sendText({
      to: message.sender,
      text: "Maaf, sistem mengalami kendala saat menyiapkan balasan. Sebagian transaksi mungkin sudah tersimpan; minta saya tampilkan transaksi terbaru sebelum mengirim ulang agar tidak tercatat ganda.",
    });
  }
}

function enqueueMessage(message) {
  const sender = message.sender.replace(/\D/g, "");
  if (!financeStore.reserveMessage(message.id, sender)) {
    console.log(`Skipped duplicate WhatsApp message ${message.id}`);
    return;
  }
  const authorizedMessage = { ...message, sender };
  const stopTyping = transport.startTyping(sender);
  void messageQueue.enqueue(async () => {
    try {
      await handleMessage(authorizedMessage);
      financeStore.finishMessage(message.id, "completed");
    } catch (error) {
      financeStore.finishMessage(message.id, "failed");
      console.error(`Could not process WhatsApp message ${message.id}:`, error.message);
    } finally {
      await stopTyping();
    }
  }).catch((error) => console.error("WhatsApp message queue failed:", error.message));
}

async function handleIncomingMessage(message) {
  try {
    transport.rememberJid(message.sender, message.recipientJid);
    enqueueMessage(message);
  } catch (error) {
    console.error(`Could not handle incoming WhatsApp message ${message.id}:`, error.message);
    await sendText({
      to: message.sender,
      text: "Pesan diterima, tetapi agent mengalami kendala. Coba kirim ulang sebentar lagi.",
    });
  }
}

const server = createServer(async (request, response) => {
  try {
    await serveRequest(request, response);
  } catch (error) {
    if (!response.headersSent) {
      response.writeHead(400, { "content-type": "text/plain; charset=utf-8" }).end("Bad request");
    } else {
      response.destroy();
    }
    console.error("Could not serve request:", error.message);
  }
});

// Stream a request to the local 9router (127.0.0.1:routerPort) and pipe the
// response back, rewriting Location headers that point at router-local paths.
async function proxyToRouter(request, response, targetPath) {
  const upstream = httpRequest({
    host: "127.0.0.1",
    port: config.routerPort,
    path: targetPath || "/",
    method: request.method,
    headers: { ...request.headers, host: `127.0.0.1:${config.routerPort}`, cookie: request.headers.cookie ?? "" },
  });
  request.pipe(upstream);
  await new Promise((resolve) => {
    upstream.on("response", (up) => {
      const headers = { ...up.headers };
      if (headers.location && headers.location.startsWith("/")) headers.location = `/9router${headers.location}`;
      response.writeHead(up.statusCode ?? 502, headers);
      up.pipe(response);
      up.on("end", resolve);
    });
    upstream.on("error", (error) => {
      if (!response.headersSent) {
        response.writeHead(502, { "content-type": "text/plain; charset=utf-8" })
          .end(`9router is unreachable on port ${config.routerPort}: ${error.message}`);
      } else {
        response.destroy();
      }
      resolve();
    });
  });
}

async function serveRequest(request, response) {
  if (await handleWorkspaceShare(request, response, { shareLinks, fileWorkspace, config })) return;

  const url = new URL(request.url ?? "/", `http://${request.headers.host ?? "localhost"}`);
  // The owner's own address beats a guess: see adoptOwnerOrigin. Gated on a
  // real login, because the links carry a bearer token.
  adoptOwnerOrigin(config, request, url.pathname, isDashboardAuthorized);

  // Reverse proxy to the local 9router dashboard: exposes the router admin UI
  // through this server's public port because the router port is not reachable
  // from outside (Pterodactyl exposes one port; ISPs may filter plain HTTP).
  if (config.routerProxyEnabled && url.pathname === "/9router") {
    response.writeHead(302, { location: "/9router/" }).end();
    return;
  }
  if (config.routerProxyEnabled && url.pathname.startsWith("/9router/")) {
    await proxyToRouter(request, response, url.pathname.slice("/9router".length) + (url.search ?? ""));
    return;
  }
  if (await handleDashboardApi(request, response, config, { backupQueue: messageQueue, backupService, filesystemMcp, fileWorkspace, financeStore, excelMcp, workflowLive, shareLinks })) return;

  if (request.method === "GET" && url.pathname === "/healthz") {
    const isConnected = transport?.isConnected ?? false;
    response.writeHead(isConnected ? 200 : 503, { "content-type": "application/json" })
      .end(JSON.stringify({ status: isConnected ? "ok" : "disconnected", whatsapp: isConnected ? "connected" : "disconnected" }));
    return;
  }
  if (request.method === "GET" && (url.pathname === "/" || url.pathname === "")) {
    response.writeHead(302, { location: "/dashboard" }).end();
    return;
  }
  if (request.method === "GET" && (url.pathname === "/dashboard" || url.pathname === "/dashboard/")) {
    response.writeHead(200, { "content-type": "text/html; charset=utf-8" })
      .end(dashboardHtml);
    return;
  }
  // Public, non-sensitive info for the login screen (router admin link).
  if (request.method === "GET" && url.pathname === "/api/public-info") {
    response.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" })
      .end(JSON.stringify({ routerUrl: config.routerUrl || (config.routerProxyEnabled ? "/9router/" : null) }));
    return;
  }
  // QR viewer page: light page that renders the current WhatsApp QR.
  if (request.method === "GET" && url.pathname === "/qr") {
    const html = await readFile(new URL("../public/qr.html", import.meta.url), "utf8");
    response.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" }).end(html);
    return;
  }
  // WhatsApp QR for panel login flow: protected by the panel password/token so
  // only the owner can scan it; never exposed publicly without auth.
  if (request.method === "GET" && url.pathname === "/api/whatsapp/qr") {
    const token = new URL(url).searchParams.get("k") ?? "";
    const sessionOk = token && dashboardSessions.has(token);
    const password = new URL(url).searchParams.get("p") ?? "";
    const passwordOk = config.panelPassword && password === config.panelPassword;
    const tokenOk = config.panelApiToken && token === config.panelApiToken;
    if (!sessionOk && !passwordOk && !tokenOk) {
      response.writeHead(404).end("Not found");
      return;
    }
    if (!transport?.lastQr) {
      response.writeHead(404, { "content-type": "application/json; charset=utf-8" })
        .end(JSON.stringify({ connected: Boolean(transport?.connected), qr: null }));
      return;
    }
    response.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" })
      .end(JSON.stringify({ connected: false, qr: transport.lastQr, generated_at: transport.lastQrAt }));
    return;
  }
  response.writeHead(404).end("Not found");
}

server.on("error", (error) => {
  console.error(`Could not start dashboard server on port ${config.port}:`, error.message);
  if (error.code === "EADDRINUSE") {
    console.error(`Port ${config.port} already in use. Stop the other process or set SERVER_PORT.`);
  }
  process.exitCode = 1;
});

// Auto-start a local headless 9router when ROUTER_PROXY is on (default) so the
// /9router/ proxy works without touching the Pterodactyl startup command. The
// package ships via package.json (npm install already runs on the egg).
function startRouterProcess() {
  if (!config.routerProxyEnabled) return;
  // Run `node node_modules/9router/cli.js` directly instead of the .bin shim:
  // .cmd files cannot be spawned without a shell on Windows, and the Linux
  // symlink adds nothing here.
  const cli = join(process.cwd(), "node_modules", "9router", "cli.js");
  if (!existsSync(cli)) {
    console.warn("ROUTER_PROXY=1 but the 9router package is not installed. Add \"9router\" to package.json dependencies (the egg's npm install will fetch it). /9router/ will be unavailable until then.");
    return;
  }
  const child = spawn(process.execPath, [cli, "-n", "-H", "127.0.0.1", "-p", String(config.routerPort), "--skip-update"], {
    stdio: "ignore",
    detached: true,
  });
  child.unref();
  child.on("error", (error) => console.error("Could not start 9router:", error.message));
  console.log(`Starting local 9router on 127.0.0.1:${config.routerPort} (headless, proxied at /9router/).`);
}

startRouterProcess();

transport = new BaileysTransport({
  authDirectory: join(config.dataDirectory, "baileys-auth"),
  maxFileBytes: config.maxFileBytes,
  onMessage: handleIncomingMessage,
});

server.listen(config.port, "0.0.0.0", async () => {
  const cyan = "\u001b[36m";
  const green = "\u001b[32m";
  const dim = "\u001b[2m";
  const reset = "\u001b[0m";
  if (!config.dashboardUrl && !consoleAddresses.dashboardUrl) {
    console.log("Public URL unavailable. Set PUBLIC_URL in Pterodactyl variables for share links and dashboard access.");
  }
  const addresses = consoleAddresses;
  console.log([
    "",
    `  ${cyan}WHATSAPP AI AGENT${reset}`,
    `  ${dim}-----------------${reset}`,
    `  ${green}Dashboard${reset}  ${addresses.dashboardUrl ?? `/dashboard (port ${config.port})`}`,
    `  ${green}Health${reset}     ${addresses.healthUrl ?? `/healthz (port ${config.port})`}`,
    ...(config.routerUrl ? [`  ${green}Router${reset}     ${config.routerUrl}  ${dim}(9router admin)`] : []),
    ...(config.routerProxyEnabled && !config.routerUrl ? [`  ${green}Router${reset}     ${addresses.dashboardUrl ?? `(port ${config.port})`}/9router/  ${dim}(9router via proxy)`] : []),
    `  ${green}QR Code${reset}    ${addresses.dashboardUrl ? `${addresses.dashboardUrl}/..` : `(port ${config.port})`}/api/whatsapp/qr?p=<PANEL_PASSWORD>  ${dim}(scan dari HP bila QR console tertimpa)`,
    ...(addresses.error ? [`  ${dim}Address lookup failed: ${addresses.error}${reset}`] : []),
    `  ${dim}Runtime    Listening on port ${config.port}${reset}`,
    "",
  ].join("\n"));
});

void transport.start().catch((error) => {
  console.error("Could not start Baileys WhatsApp connection:", error.message);
  server.close();
  process.exitCode = 1;
});

async function shutdown() {
  await transport.stop();
  await filesystemMcp?.close();
  await excelMcp.close();
  shareLinks.close();
  server.close();
  financeStore.close();
  fileWorkspace.close();
}

process.once("SIGINT", () => { void shutdown(); });
process.once("SIGTERM", () => { void shutdown(); });

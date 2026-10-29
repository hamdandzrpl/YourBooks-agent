import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { EventEmitter } from "node:events";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { BaileysTransport } from "../src/baileys.js";
import { FinanceStore } from "../src/finance/store.js";
import { OfficeFiles } from "../src/files/office.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { ShareLinkStore } from "../src/share-links.js";
import { processIncomingMessage } from "../src/message-handler.js";

test("a plain greeting still gets a normal reply when the Excel MCP cannot start", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "whatsapp-greeting-fallback-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  let agentInput;
  let reply;
  await processIncomingMessage({
    message: { id: "wamid.plain-greeting", sender: "6281234567890", text: "halo" },
    config: { xkiroApiKey: "test-key", xkiroModel: "test-model", timeZone: "Asia/Jakarta", defaultCurrency: "IDR" },
    financeStore,
    fileWorkspace,
    officeFiles: null,
    excelMcp: { async connect() { throw new Error("Excel binary unavailable"); } },
    runAgentImpl: async (input) => { agentInput = input; return "Halo! Ada yang bisa saya bantu?"; },
    sendText: async (input) => { reply = input; },
    sendCsv: async () => {},
    sendFile: async () => {},
  });

  assert.equal(agentInput.text, "halo");
  assert.deepEqual(agentInput.mcpSources, []);
  assert.equal(reply.text, "Halo! Ada yang bisa saya bantu?");
});

test("incoming WhatsApp file is saved for its sender before XKIRO processes the request", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "whatsapp-message-test-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const bytes = Buffer.from("template workbook bytes");
  const sha256 = createHash("sha256").update(bytes).digest("hex");
  let agentInput;
  let reply;
  const result = await processIncomingMessage({
    message: {
      id: "wamid.upload.template",
      sender: "6281234567890",
      text: "Pakai template ini untuk mencatat keuangan nanti",
      attachment: {
        mediaId: "media-1",
        filename: "template.xlsx",
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        sha256,
        kind: "document",
        load: async () => bytes,
      },
    },
    config: {
      maxFileBytes: 1024,
      xkiroApiKey: "test-key",
      xkiroModel: "stealth/space-bunny-alpha:free",
      timeZone: "Asia/Jakarta",
      defaultCurrency: "IDR",
    },
    financeStore,
    fileWorkspace,
    officeFiles: {},
    runAgentImpl: async (input) => {
      agentInput = input;
      return "Template sudah disimpan untuk dipakai nanti.";
    },
    sendText: async (input) => { reply = input; },
    sendCsv: async () => {},
    sendFile: async () => {},
  });

  assert.equal(result.savedAttachment.name, "template.xlsx");
  assert.equal(agentInput.userId, "6281234567890");
  assert.match(agentInput.text, new RegExp(result.savedAttachment.id));
  assert.equal(agentInput.fileWorkspace, fileWorkspace);
  assert.deepEqual(agentInput.officeFiles, {});
  assert.equal(result.agentText, agentInput.text);
  assert.equal(reply.to, "6281234567890");
  assert.match(reply.text, /disimpan/);
  assert.equal((await fileWorkspace.listFiles("6281234567890")).length, 1);
  assert.equal((await fileWorkspace.listFiles("6289999999999")).length, 0);
});

test("natural-language file requests reach the AI with scoped workspace tools and no keyword handler", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "whatsapp-file-tools-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  await fileWorkspace.createTextFile("6281234567890", "notes.txt", "private");

  let agentInput;
  await processIncomingMessage({
    message: { id: "wamid.no-file-tools", sender: "6281234567890", text: "Tolong tampilkan catatan saya" },
    config: { xkiroApiKey: "test-key", xkiroModel: "test-model", timeZone: "Asia/Jakarta", defaultCurrency: "IDR" },
    financeStore,
    fileWorkspace,
    officeFiles: {},
    runAgentImpl: async (input) => {
      agentInput = input;
      return "Akses file tersedia melalui panel web.";
    },
    sendText: async () => {},
    sendCsv: async () => {},
    sendFile: async () => {},
  });

  assert.equal(agentInput.text, "Tolong tampilkan catatan saya");
  assert.equal(agentInput.fileWorkspace, fileWorkspace);
  assert.deepEqual(agentInput.officeFiles, {});
});

test("legacy-looking finance text is sent to the AI without local keyword dispatch", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "whatsapp-ai-routing-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  let agentInput;
  await processIncomingMessage({
    message: { id: "wamid.ai-routing", sender: "6281234567890", text: "LAPORAN 2026-10" },
    config: { xkiroApiKey: "test-key", xkiroModel: "test-model", timeZone: "Asia/Jakarta", defaultCurrency: "IDR" },
    financeStore,
    fileWorkspace,
    officeFiles: null,
    runAgentImpl: async (input) => {
      agentInput = input;
      return "Aku akan menyiapkan laporan lewat data tersimpan.";
    },
    sendText: async () => {},
    sendCsv: async () => {},
    sendFile: async () => {},
  });

  assert.equal(agentInput.text, "LAPORAN 2026-10");
});

test("a WhatsApp turn uses stored memory/history and persists the next turn", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "whatsapp-memory-turn-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  financeStore.saveConversationTurn("6281234567890", "wamid-old", "Saya suka laporan mingguan", "Akan saya ingat.");
  financeStore.updateMemory("6281234567890", "user", [
    { action: "add", content: "User prefers concise Indonesian replies." },
  ]);

  let modelInput;
  await processIncomingMessage({
    message: { id: "wamid-new", sender: "6281234567890", text: "Kirim ringkasan minggu ini" },
    config: {
      xkiroApiKey: "test-key",
      xkiroModel: "stealth/space-bunny-alpha:free",
      timeZone: "Asia/Jakarta",
      defaultCurrency: "IDR",
    },
    financeStore,
    fileWorkspace,
    history: financeStore.getRecentConversation("6281234567890", 12),
    runAgentImpl: async (input) => {
      modelInput = input;
      return "Berikut ringkasan mingguan.";
    },
    sendText: async () => {},
    sendCsv: async () => {},
    sendFile: async () => {},
  });

  assert.equal(modelInput.history[0].content, "Saya suka laporan mingguan");
  assert.match(modelInput.memoryState.user.entries[0].content, /concise Indonesian/);
  assert.deepEqual(financeStore.getRecentConversation("6281234567890", 2).map((turn) => turn.role), ["user", "assistant"]);
});

test("an incoming Baileys message reaches the AI handler and sends its reply", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "whatsapp-reply-flow-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const events = new EventEmitter();
  const sentMessages = [];
  let resolveSent;
  const sent = new Promise((resolve) => { resolveSent = resolve; });
  let transport;

  t.after(async () => {
    await transport.stop();
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const socket = {
    ev: events,
    sendMessage: async (jid, content) => {
      sentMessages.push({ jid, content });
      resolveSent();
    },
    sendPresenceUpdate: async () => {},
    end: () => {},
  };
  transport = new BaileysTransport({
    authDirectory: join(dataDirectory, "auth"),
    maxFileBytes: 1024,
    socketFactory: () => socket,
    authLoader: async () => ({ state: { creds: {}, keys: {} }, saveCreds: async () => {} }),
    qrRenderer: () => {},
    onMessage: (message) => processIncomingMessage({
      message,
      config: {
        xkiroApiKey: "test-key",
        xkiroModel: "test-model",
        timeZone: "Asia/Jakarta",
        defaultCurrency: "IDR",
      },
      financeStore,
      fileWorkspace,
      history: financeStore.getRecentConversation(message.sender),
      runAgentImpl: async () => "Halo, ini jawaban agent.",
      sendText: (reply) => transport.sendText(reply),
      sendCsv: async () => {},
      sendFile: async () => {},
    }),
  });

  await transport.start();
  events.emit("connection.update", { connection: "open" });
  events.emit("messages.upsert", {
    type: "notify",
    messages: [{
      key: { id: "wamid.reply-flow", remoteJid: "6281234567890@s.whatsapp.net" },
      message: { conversation: "Halo" },
    }],
  });
  await sent;

  assert.deepEqual(sentMessages, [{
    jid: "6281234567890@s.whatsapp.net",
    content: { text: "Halo, ini jawaban agent." },
  }]);
  assert.deepEqual(financeStore.getRecentConversation("6281234567890", 2).map(({ role }) => role), ["user", "assistant"]);
});

test("the agent's workspace tools can create a share link", async (t) => {
  // The per-message MCP bridge used to be built without shareLinks, so every
  // share_link_* action answered "Share links are not enabled on this server"
  // and the agent told the user the web page could not be created.
  const dataDirectory = mkdtempSync(join(tmpdir(), "whatsapp-share-bridge-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const shareLinks = new ShareLinkStore({ dataDirectory });
  t.after(() => {
    shareLinks.close();
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  let agentInput;
  let bridge;
  let created;
  let written;
  await processIncomingMessage({
    message: { id: "wamid.share-bridge", sender: "6281234567890", text: "buatkan workspace kebun" },
    config: {
      xkiroApiKey: "test-key", xkiroModel: "test-model", timeZone: "Asia/Jakarta",
      defaultCurrency: "IDR", publicOrigin: "https://panel.example.test",
    },
    financeStore,
    fileWorkspace,
    officeFiles: new OfficeFiles(fileWorkspace),
    shareLinks,
    runAgentImpl: async (input) => {
      agentInput = input;
      // The bridge is closed when the turn ends, so exercise it while the turn
      // is still running, exactly as the agent would.
      bridge = input.mcpSources[0];
      created = await bridge.call("mcp_filesystem_workspace_files", {
        action: "share_link_create",
        folder: "documents/kebun",
        name: "kebun",
        title: "Kebun",
      });
      written = await bridge.call("mcp_filesystem_workspace_files", {
        action: "share_link_set_rows",
        folder: "documents/kebun",
        sheets: [
          { name: "Dashboard", columns: ["Komponen", "Nilai"], rows: [["Total Pemasukan", 100], ["Total Pengeluaran", 40]] },
          { name: "Data", columns: ["Nilai"], rows: [[100]] },
        ],
      });
      return "Siap.";
    },
    sendText: async () => {},
    sendCsv: async () => {},
    sendFile: async () => {},
  });

  assert.equal(agentInput.shareLinks, shareLinks);
  // A getter, so an address learned later in the turn is still used.
  assert.equal(agentInput.publicOrigin(), "https://panel.example.test");
  assert.ok(bridge, "the agent must receive the file workspace tools");
  assert.equal(created.isError, false, created.content);
  assert.match(created.content, /https:\/\/panel\.example\.test\/w\/kebun\?k=/);
  assert.equal(written.isError, false, written.content);
  const payload = JSON.parse(written.content);
  assert.equal(payload.saved, true);
  assert.equal(payload.summary.kpis.length, 2);
});
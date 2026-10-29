import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { EventEmitter } from "node:events";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";
import test from "node:test";
import { BaileysTransport, parseBaileysMessage, redactLibsignalSessionLogs } from "../src/baileys.js";

test("libsignal session logs redact key-bearing session objects", () => {
  const output = [];
  const consoleTarget = {
    info(...args) { output.push(["info", ...args]); },
    warn(...args) { output.push(["warn", ...args]); },
  };
  const restore = redactLibsignalSessionLogs(consoleTarget);
  try {
    consoleTarget.info("Closing session:", { currentRatchet: { rootKey: "private-key-material" } });
    consoleTarget.warn("Session already closed", { currentRatchet: { rootKey: "private-key-material" } });
    consoleTarget.info("ordinary diagnostic", { status: "ok" });
  } finally {
    restore();
  }

  assert.deepEqual(output, [
    ["info", "Closing session (details redacted)"],
    ["warn", "Session already closed (details redacted)"],
    ["info", "ordinary diagnostic", { status: "ok" }],
  ]);
});

test("Baileys parser accepts any direct sender and ignores groups and self messages", () => {
  const parse = (message) => parseBaileysMessage(message, { maxFileBytes: 1024, socket: {} });

  assert.deepEqual(parse({
    key: { id: "msg-1", remoteJid: "6281234567890@s.whatsapp.net" },
    message: { conversation: "Halo" },
  }), {
    id: "msg-1",
    sender: "6281234567890",
    recipientJid: "6281234567890@s.whatsapp.net",
    text: "Halo",
  });
  assert.equal(parse({ key: { id: "group-1", remoteJid: "123@g.us" }, message: { conversation: "Halo" } }), null);
  assert.equal(parse({ key: { id: "self-1", remoteJid: "6281234567890@s.whatsapp.net", fromMe: true }, message: { conversation: "Halo" } }), null);
  assert.deepEqual(parse({
    key: { id: "other-1", remoteJid: "6289999999999@s.whatsapp.net" },
    message: { conversation: "Halo" },
  }), {
    id: "other-1",
    sender: "6289999999999",
    recipientJid: "6289999999999@s.whatsapp.net",
    text: "Halo",
  });
  const firstRunMessage = parseBaileysMessage({
    key: { id: "pairing-1", remoteJid: "6289999999999@s.whatsapp.net" },
    message: { conversation: "AKTIFKAN ABCDEF012345" },
  }, { maxFileBytes: 1024, socket: {} });
  assert.equal(firstRunMessage.text, "AKTIFKAN ABCDEF012345");
});

test("Baileys media is loaded lazily with a size bound and checksum verification", async () => {
  const bytes = Buffer.from("baileys media bytes");
  const sha256 = createHash("sha256").update(bytes).digest();
  let downloaded = false;
  const message = {
    key: { id: "media-1", remoteJid: "6281234567890@s.whatsapp.net" },
    message: {
      documentMessage: {
        fileName: "template.xlsx",
        mimetype: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileLength: String(bytes.length),
        fileSha256: sha256,
        caption: "Template kas",
      },
    },
  };
  const parsed = parseBaileysMessage(message, {
    maxFileBytes: 1024,
    socket: { updateMediaMessage: async () => {} },
    mediaDownloader: async () => {
      downloaded = true;
      return Readable.from([bytes]);
    },
  });

  assert.equal(downloaded, false);
  assert.equal(parsed.text, "Template kas");
  assert.equal(parsed.attachment.filename, "template.xlsx");
  assert.deepEqual(await parsed.attachment.load(), bytes);
  assert.equal(downloaded, true);

  const tooLarge = parseBaileysMessage({
    key: { id: "big-1", remoteJid: "6281234567890@s.whatsapp.net" },
    message: { documentMessage: { fileName: "big.xlsx", fileLength: "5000" } },
  }, { maxFileBytes: 1024, socket: {} });
  assert.match(tooLarge.attachmentError, /melebihi batas/);
});

test("Baileys transport persists auth, prints QR, ignores backfill, and replies to direct JID", async (t) => {
  const directory = mkdtempSync(join(tmpdir(), "baileys-transport-test-"));
  t.after(() => rmSync(directory, { recursive: true, force: true }));
  const events = new EventEmitter();
  const sent = [];
  const incoming = [];
  const presenceUpdates = [];
  let savedCredentials = 0;
  const renderedOutput = [];
  t.mock.method(console, "log", (...values) => renderedOutput.push(values.join(" ")));
  const socket = {
    ev: events,
    sendMessage: async (jid, content) => {
      sent.push({ jid, content });
      return { key: { remoteJid: jid, fromMe: true, id: `sent-${sent.length}` } };
    },
    sendPresenceUpdate: async (presence, jid) => { presenceUpdates.push({ presence, jid }); },
    end: () => {},
    updateMediaMessage: async () => {},
  };
  const transport = new BaileysTransport({
    authDirectory: join(directory, "auth"),
    maxFileBytes: 1024,
    onMessage: (message) => incoming.push(message),
    socketFactory: () => socket,
    authLoader: async () => ({ state: { creds: {}, keys: {} }, saveCreds: async () => { savedCredentials += 1; } }),
  });

  await transport.start();
  events.emit("connection.update", { qr: "qr-secret-payload" });
  events.emit("creds.update");
  events.emit("messages.upsert", {
    type: "append",
    messages: [{ key: { id: "history", remoteJid: "6281234567890@s.whatsapp.net" }, message: { conversation: "old" } }],
  });
  events.emit("messages.upsert", {
    type: "notify",
    messages: [{ key: { id: "new", remoteJid: "6281234567890@s.whatsapp.net" }, message: { conversation: "hi" } }],
  });
  await new Promise((resolve) => setImmediate(resolve));

  assert.match(renderedOutput.join("\n"), /\u2580|\u2584|\u2588|\u001b\[47m/);
  assert.equal(savedCredentials, 1);
  assert.equal(incoming.length, 1);
  assert.equal(incoming[0].text, "hi");
  const stopTyping = transport.startTyping("6281234567890");
  await stopTyping();
  assert.deepEqual(presenceUpdates.map(({ presence }) => presence), ["composing", "paused"]);
  assert.ok(presenceUpdates.every(({ jid }) => jid === "6281234567890@s.whatsapp.net"));
  const sentMessage = await transport.sendText({ to: "6281234567890", text: "reply" });
  assert.deepEqual(sent[0], {
    jid: "6281234567890@s.whatsapp.net",
    content: { text: "reply" },
  });
  await transport.editText({ to: "6281234567890", messageKey: sentMessage.key, text: "Progress\n> Selesai" });
  assert.deepEqual(sent[1], {
    jid: "6281234567890@s.whatsapp.net",
    content: { text: "Progress\n> Selesai", edit: sentMessage.key },
  });
  await assert.rejects(() => transport.editText({
    to: "6281234567890",
    messageKey: { remoteJid: "6281234567890@s.whatsapp.net", fromMe: false, id: "incoming" },
    text: "unsafe edit",
  }), /valid sent WhatsApp message key/);
  await transport.stop();
});

test("Baileys send and typing timeouts do not block message processing", async (t) => {
  const directory = mkdtempSync(join(tmpdir(), "baileys-timeout-test-"));
  t.after(() => rmSync(directory, { recursive: true, force: true }));
  const socket = {
    ev: new EventEmitter(),
    sendMessage: () => new Promise(() => {}),
    sendPresenceUpdate: () => new Promise(() => {}),
    end: () => {},
  };
  const transport = new BaileysTransport({
    authDirectory: join(directory, "auth"),
    maxFileBytes: 1024,
    onMessage: () => {},
    socketFactory: () => socket,
    authLoader: async () => ({ state: { creds: {}, keys: {} }, saveCreds: async () => {} }),
    qrRenderer: () => {},
    sendTimeoutMs: 10,
    presenceTimeoutMs: 10,
  });

  await transport.start();
  await assert.rejects(() => transport.sendText({ to: "6281234567890", text: "reply" }), /text send timed out/);
  const stopTyping = transport.startTyping("6281234567890");
  await stopTyping();
  await transport.stop();
});
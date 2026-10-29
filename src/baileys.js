import { createHash } from "node:crypto";
import { mkdir } from "node:fs/promises";
import makeWASocket, {
  DisconnectReason,
  downloadMediaMessage,
  getContentType,
  normalizeMessageContent,
  useMultiFileAuthState,
} from "@whiskeysockets/baileys";
import { withTimeout } from "./async.js";
import pino from "pino";
import qrcode from "qrcode-terminal";

const logger = pino({ level: "silent" });
const sensitiveSessionLogPrefixes = [
  "Session already closed",
  "Closing session:",
  "Opening session:",
  "Removing old closed session:",
];
const mediaTypes = {
  imageMessage: { kind: "image", mimeType: "image/jpeg", extension: "jpg" },
  videoMessage: { kind: "video", mimeType: "video/mp4", extension: "mp4" },
  audioMessage: { kind: "audio", mimeType: "audio/ogg", extension: "ogg" },
  documentMessage: { kind: "document", mimeType: "application/octet-stream", extension: "bin" },
  stickerMessage: { kind: "sticker", mimeType: "image/webp", extension: "webp" },
};

export function redactLibsignalSessionLogs(consoleTarget = console) {
  const originals = new Map();
  for (const level of ["info", "warn"]) {
    const original = consoleTarget[level];
    if (typeof original !== "function") continue;
    originals.set(level, original);
    consoleTarget[level] = function (message, ...details) {
      if (typeof message === "string" && sensitiveSessionLogPrefixes.some((prefix) => message.startsWith(prefix))) {
        const safeMessage = message.endsWith(":")
          ? `${message.slice(0, -1)} (details redacted)`
          : `${message} (details redacted)`;
        return Reflect.apply(original, this, [safeMessage]);
      }
      return Reflect.apply(original, this, [message, ...details]);
    };
  }
  return () => {
    for (const [level, original] of originals) consoleTarget[level] = original;
  };
}

redactLibsignalSessionLogs();

function getDirectJid(message) {
  const candidates = [message.key?.remoteJidAlt, message.key?.remoteJid];
  return candidates.find((jid) => typeof jid === "string" && jid.endsWith("@s.whatsapp.net")) ?? null;
}

function phoneFromJid(jid) {
  const number = jid.slice(0, jid.indexOf("@")).split(":")[0].replace(/\D/g, "");
  return /^\d{8,15}$/.test(number) ? number : null;
}

async function collectMediaStream(stream, maxBytes) {
  const chunks = [];
  let size = 0;
  for await (const chunk of stream) {
    size += chunk.length;
    if (size > maxBytes) {
      stream.destroy?.();
      throw new Error("Incoming WhatsApp file exceeds the configured size limit");
    }
    chunks.push(Buffer.from(chunk));
  }
  if (size === 0) throw new Error("WhatsApp media download returned an empty file");
  return Buffer.concat(chunks, size);
}

export function parseBaileysMessage(message, { maxFileBytes, socket, mediaDownloader = downloadMediaMessage }) {
  if (!message?.key?.id || message.key.fromMe) return null;
  const jid = getDirectJid(message);
  if (!jid) return null;
  const sender = phoneFromJid(jid);
  if (!sender) return null;

  const content = normalizeMessageContent(message.message);
  if (!content) return null;
  const type = getContentType(content);
  if (type === "conversation" || type === "extendedTextMessage") {
    const text = type === "conversation" ? content.conversation : content.extendedTextMessage?.text;
    if (!text) return null;
    return { id: message.key.id, sender, recipientJid: jid, text };
  }

  const mediaConfig = mediaTypes[type];
  const media = mediaConfig ? content[type] : null;
  if (!media) return null;
  const filename = media.fileName ?? `${message.key.id}.${mediaConfig.extension}`;
  const mimeType = media.mimetype ?? mediaConfig.mimeType;
  const declaredSize = Number(media.fileLength ?? 0);
  if (declaredSize > maxFileBytes) {
    return {
      id: message.key.id,
      sender,
      recipientJid: jid,
      text: media.caption ?? "",
      attachmentError: "File melebihi batas ukuran yang diizinkan.",
    };
  }

  return {
    id: message.key.id,
    sender,
    recipientJid: jid,
    text: media.caption ?? "",
    attachment: {
      filename,
      mimeType,
      kind: mediaConfig.kind,
      sha256: Buffer.isBuffer(media.fileSha256) ? media.fileSha256.toString("hex") : null,
      load: async () => {
        const stream = await mediaDownloader(message, "stream", {}, {
          logger,
          reuploadRequest: socket.updateMediaMessage,
        });
        const bytes = await collectMediaStream(stream, maxFileBytes);
        if (declaredSize && bytes.length !== declaredSize) {
          throw new Error("Downloaded file size does not match WhatsApp metadata");
        }
        const expectedHash = Buffer.isBuffer(media.fileSha256) ? media.fileSha256.toString("hex") : null;
        const actualHash = createHash("sha256").update(bytes).digest("hex");
        if (expectedHash && actualHash !== expectedHash) {
          throw new Error("Downloaded file checksum does not match WhatsApp metadata");
        }
        return bytes;
      },
    },
  };
}

export class BaileysTransport {
  constructor({ authDirectory, maxFileBytes, onMessage, socketFactory = makeWASocket, authLoader = useMultiFileAuthState, qrRenderer = qrcode.generate.bind(qrcode), sendTimeoutMs = 30_000, presenceTimeoutMs = 5_000 }) {
    this.authDirectory = authDirectory;
    this.maxFileBytes = maxFileBytes;
    this.onMessage = onMessage;
    this.socketFactory = socketFactory;
    this.authLoader = authLoader;
    this.qrRenderer = qrRenderer;
    this.sendTimeoutMs = sendTimeoutMs;
    this.presenceTimeoutMs = presenceTimeoutMs;
    this.socket = null;
    this.connected = false;
    this.authState = null;
    this.saveCredentials = null;
    this.reconnectTimer = null;
    this.stopped = false;
    this.lastQr = "";
    this.lastQrAt = "";
    this.jidsByNumber = new Map();
  }

  async start() {
    await mkdir(this.authDirectory, { recursive: true, mode: 0o700 });
    const auth = await this.authLoader(this.authDirectory);
    this.authState = auth.state;
    this.saveCredentials = auth.saveCreds;
    this.stopped = false;
    this.connect();
  }

  connect() {
    if (this.stopped) return;
    const socket = this.socketFactory({
      auth: this.authState,
      logger,
      markOnlineOnConnect: false,
    });
    this.socket = socket;
    socket.ev.on("creds.update", this.saveCredentials);
    socket.ev.on("connection.update", ({ connection, lastDisconnect, qr }) => {
      if (qr) {
        this.lastQr = qr;
        this.lastQrAt = new Date().toISOString();
        console.log("WhatsApp belum terhubung. Scan QR melalui Perangkat tertaut:");
        this.qrRenderer(qr, { small: true });
      }
      if (connection === "open") {
        this.lastQr = "";
        this.connected = true;
        console.log("WhatsApp terhubung.");
      }
      if (connection === "close" && this.socket === socket) {
        this.connected = false;
        this.socket = null;
        const statusCode = lastDisconnect?.error?.output?.statusCode;
        console.warn(`Baileys connection closed (${statusCode ?? "unknown"}): ${lastDisconnect?.error?.message ?? "unknown reason"}`);
        if (statusCode === DisconnectReason.loggedOut) {
          console.error("Baileys session logged out; remove the persisted auth folder and scan a new QR code.");
          return;
        }
        if (!this.stopped && !this.reconnectTimer) {
          this.reconnectTimer = setTimeout(() => {
            this.reconnectTimer = null;
            this.connect();
          }, 5000);
          this.reconnectTimer.unref?.();
        }
      }
    });
    socket.ev.on("messages.upsert", ({ messages, type }) => {
      if (type !== "notify") return;
      for (const rawMessage of messages) {
        const message = parseBaileysMessage(rawMessage, {
          maxFileBytes: this.maxFileBytes,
          socket,
        });
        if (!message) {
          continue;
        }
        this.rememberJid(message.sender, message.recipientJid);
        Promise.resolve(this.onMessage(message)).catch((error) => {
          console.error(`Could not enqueue WhatsApp message ${message.id}:`, error.message);
        });
      }
    });
  }

  getSocket() {
    if (!this.socket) throw new Error("Baileys is not connected to WhatsApp yet");
    return this.socket;
  }

  get isConnected() {
    return this.connected;
  }

  rememberJid(phoneNumber, jid) {
    if (typeof jid === "string"
      && jid.endsWith("@s.whatsapp.net")) {
      this.jidsByNumber.set(phoneNumber, jid);
    }
  }

  startTyping(to) {
    const jid = this.jidsByNumber.get(to) ?? `${to}@s.whatsapp.net`;
    let warningLogged = false;
    let presenceTask = Promise.resolve();
    const updatePresence = (presence) => {
      presenceTask = presenceTask
        .catch(() => {})
        .then(() => withTimeout(
          () => this.getSocket().sendPresenceUpdate(presence, jid),
          this.presenceTimeoutMs,
          "WhatsApp typing presence timed out",
        ))
      .catch((error) => {
        if (!warningLogged) {
          warningLogged = true;
          console.warn(`Could not update WhatsApp typing presence: ${error.message}`);
        }
      });
      return presenceTask;
    };

    void updatePresence("composing");
    const timer = setInterval(() => {
      void updatePresence("composing");
    }, 10_000);
    timer.unref?.();

    return async () => {
      clearInterval(timer);
      await updatePresence("paused");
    };
  }

  async sendText({ to, text }) {
    const jid = this.jidsByNumber.get(to) ?? `${to}@s.whatsapp.net`;
    return withTimeout(
      () => this.getSocket().sendMessage(jid, { text: String(text).slice(0, 4000) }),
      this.sendTimeoutMs,
      "WhatsApp text send timed out",
    );
  }

  async editText({ to, messageKey, text }) {
    const jid = this.jidsByNumber.get(to) ?? `${to}@s.whatsapp.net`;
    if (!messageKey?.id || messageKey.fromMe !== true || messageKey.remoteJid !== jid) {
      throw new Error("A valid sent WhatsApp message key is required to edit progress");
    }
    return withTimeout(
      () => this.getSocket().sendMessage(jid, {
        text: String(text).slice(0, 4000),
        edit: messageKey,
      }),
      this.sendTimeoutMs,
      "WhatsApp progress edit timed out",
    );
  }

  async sendFile({ to, filename, mimeType, bytes }) {
    if (!Buffer.isBuffer(bytes) || bytes.length === 0 || bytes.length > this.maxFileBytes) {
      throw new Error("Outgoing file is empty or exceeds the configured size limit");
    }
    const jid = this.jidsByNumber.get(to) ?? `${to}@s.whatsapp.net`;
    return withTimeout(() => this.getSocket().sendMessage(jid, {
        document: bytes,
        mimetype: mimeType || "application/octet-stream",
        fileName: filename,
        caption: "File dari agent.",
      }),
      this.sendTimeoutMs,
      "WhatsApp file send timed out",
    );
  }

  async stop() {
    this.stopped = true;
    this.connected = false;
    if (this.reconnectTimer) clearTimeout(this.reconnectTimer);
    this.reconnectTimer = null;
    this.socket?.end(undefined);
    this.socket = null;
  }
}
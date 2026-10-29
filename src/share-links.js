import { createHash, randomBytes, timingSafeEqual, scryptSync } from "node:crypto";
import { DatabaseSync } from "node:sqlite";
import { mkdirSync } from "node:fs";
import { join } from "node:path";

const namePattern = /^[a-z0-9][a-z0-9-]{0,47}$/;

function normalizeName(value) {
  const name = String(value ?? "").trim().toLowerCase()
    .replaceAll("_", "-")
    .replace(/\s+/g, "-")
    .replace(/[^a-z0-9-]/g, "")
    .replace(/-{2,}/g, "-")
    .replace(/^-+|-+$/g, "");
  if (!namePattern.test(name)) {
    throw new Error("Workspace name must contain at least one letter or number (max 48: a-z, 0-9, dashes)");
  }
  return name;
}

function hashToken(token) {
  return createHash("sha256").update(token).digest("hex");
}

/**
 * Public share links for workspace folders. Each shared folder ("workspace")
 * maps a URL-safe name to a folder path inside the owner's file workspace and
 * carries a rotating bearer token. Links look like /w/<name>?k=<token>.
 */
export class ShareLinkStore {
  constructor({ dataDirectory, now = () => new Date() }) {
    const databasePath = join(dataDirectory, "share-links.sqlite");
    mkdirSync(dataDirectory, { recursive: true });
    this.database = new DatabaseSync(databasePath);
    this.database.exec("PRAGMA journal_mode = WAL; PRAGMA busy_timeout = 5000;");
    this.database.exec(`
      CREATE TABLE IF NOT EXISTS share_links (
        name TEXT PRIMARY KEY,
        user_key TEXT NOT NULL,
        folder TEXT NOT NULL,
        token_hash TEXT NOT NULL,
        title TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        pin_salt TEXT,
        pin_hash TEXT,
        failed_attempts INTEGER NOT NULL DEFAULT 0,
        locked_until INTEGER NOT NULL DEFAULT 0
      );
    `);
    const columns = new Set(this.database.prepare("PRAGMA table_info(share_links)").all().map((column) => column.name));
    if (!columns.has("pin_salt")) this.database.exec("ALTER TABLE share_links ADD COLUMN pin_salt TEXT");
    if (!columns.has("pin_hash")) this.database.exec("ALTER TABLE share_links ADD COLUMN pin_hash TEXT");
    if (!columns.has("failed_attempts")) this.database.exec("ALTER TABLE share_links ADD COLUMN failed_attempts INTEGER NOT NULL DEFAULT 0");
    if (!columns.has("locked_until")) this.database.exec("ALTER TABLE share_links ADD COLUMN locked_until INTEGER NOT NULL DEFAULT 0");
    this.now = now;
  }

  close() {
    this.database.close();
  }

  /** Create (or re-point) a share for a workspace folder. Returns the raw
   *  token exactly once; only its hash is stored. */
  issueToken(userId, folderPath, { name, title } = {}) {
    const shareName = normalizeName(name ?? folderPath.split("/").pop());
    const token = randomBytes(24).toString("base64url");
    const now = this.now().toISOString();
    this.database.prepare(`
      INSERT INTO share_links (name, user_key, folder, token_hash, title, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(name) DO UPDATE SET
        user_key = excluded.user_key,
        folder = excluded.folder,
        token_hash = excluded.token_hash,
        title = excluded.title,
        updated_at = excluded.updated_at
    `).run(shareName, createHash("sha256").update(userId).digest("hex"), folderPath, hashToken(token), String(title ?? "").slice(0, 120), now, now);
    return { name: shareName, token, folder: folderPath, title: String(title ?? "").slice(0, 120) };
  }

  /** Resolve a link visitor to (userId, folder). Returns null when unknown. */
  resolve(name, token) {
    const shareName = normalizeName(name);
    if (typeof token !== "string" || !token) return null;
    const row = this.database.prepare(
      "SELECT user_key, folder, token_hash, title FROM share_links WHERE name = ?"
    ).get(shareName);
    if (!row) return null;
    const supplied = Buffer.from(hashToken(token));
    const expected = Buffer.from(row.token_hash);
    if (supplied.length !== expected.length || !timingSafeEqual(supplied, expected)) return null;
    return { userKey: row.user_key, folder: row.folder, name: shareName, title: row.title };
  }

  lookup(name) {
    const shareName = normalizeName(name);
    const row = this.database.prepare("SELECT user_key, folder, title FROM share_links WHERE name = ?").get(shareName);
    return row ? { userKey: row.user_key, folder: row.folder, title: row.title, name: shareName } : null;
  }

  /** Set a 6-digit PIN once, or replace it only when owner explicitly resets. */
  setPin(userId, name, pin, { replace = false } = {}) {
    const shareName = normalizeName(name);
    const pinText = String(pin ?? "").trim();
    if (!/^\d{6}$/.test(pinText)) throw new Error("PIN harus 6 digit angka");
    const salt = randomBytes(16).toString("hex");
    const hash = scryptSync(pinText, salt, 32).toString("hex");
    const sql = replace
      ? "UPDATE share_links SET pin_salt = ?, pin_hash = ?, failed_attempts = 0, locked_until = 0, updated_at = ? WHERE name = ? AND user_key = ?"
      : "UPDATE share_links SET pin_salt = ?, pin_hash = ?, failed_attempts = 0, locked_until = 0, updated_at = ? WHERE name = ? AND user_key = ? AND pin_hash IS NULL";
    const result = this.database.prepare(sql).run(salt, hash, this.now().toISOString(), shareName, createHash("sha256").update(userId).digest("hex"));
    if (!result.changes) throw new Error(replace ? "Workspace not found" : "PIN workspace already configured");
    return { pinSet: true, name: shareName };
  }

  /** True when this workspace is protected by a PIN. */
  hasPin(userId, name) {
    const shareName = normalizeName(name);
    const row = this.database.prepare("SELECT pin_salt, pin_hash FROM share_links WHERE name = ? AND user_key = ?")
      .get(shareName, createHash("sha256").update(userId).digest("hex"));
    return Boolean(row?.pin_salt && row?.pin_hash);
  }

  /** Verify a PIN, throttling repeated failures. Returns ok/locked info. */
  verifyPin(userId, name, pin, { maxAttempts = 5, lockMs = 5 * 60 * 1000 } = {}) {
    const shareName = normalizeName(name);
    const userKey = createHash("sha256").update(userId).digest("hex");
    const row = this.database.prepare("SELECT pin_salt, pin_hash, failed_attempts, locked_until FROM share_links WHERE name = ? AND user_key = ?")
      .get(shareName, userKey);
    if (!row?.pin_hash) return { ok: false, reason: "no_pin" };
    if (row.locked_until > Date.now()) {
      return { ok: false, reason: "locked", retryAfterMs: row.locked_until - Date.now() };
    }
    const pinText = String(pin ?? "").trim();
    if (!/^\d{6}$/.test(pinText)) return { ok: false, reason: "invalid" };
    const supplied = scryptSync(pinText, row.pin_salt, 32);
    const expected = Buffer.from(row.pin_hash, "hex");
    const ok = supplied.length === expected.length && timingSafeEqual(supplied, expected);
    if (ok) {
      this.database.prepare("UPDATE share_links SET failed_attempts = 0 WHERE name = ? AND user_key = ?").run(shareName, userKey);
      return { ok: true, name: shareName };
    }
    let failures = Number(row.failed_attempts ?? 0) + 1;
    const lockedUntil = failures >= maxAttempts ? Date.now() + lockMs : 0;
    if (lockedUntil > 0) failures = 0;
    this.database.prepare("UPDATE share_links SET failed_attempts = ?, locked_until = ? WHERE name = ? AND user_key = ?")
      .run(failures, lockedUntil, shareName, userKey);
    return { ok: false, reason: "invalid", locked: lockedUntil > 0, retryAfterMs: lockedUntil > 0 ? lockMs : null };
  }

  /** Clear the PIN so the workspace opens directly again. */
  clearPin(userId, name) {
    const shareName = normalizeName(name);
    this.database.prepare("UPDATE share_links SET pin_salt = NULL, pin_hash = NULL, failed_attempts = 0, locked_until = 0 WHERE name = ? AND user_key = ?")
      .run(shareName, createHash("sha256").update(userId).digest("hex"));
    return { cleared: shareName };
  }

  /** Rotate the token; existing links die immediately. Returns the new token. */
  rotateToken(userId, name) {
    const shareName = normalizeName(name);
    const userKeyHash = createHash("sha256").update(userId).digest("hex");
    const row = this.database.prepare("SELECT name FROM share_links WHERE name = ? AND user_key = ?").get(shareName, userKeyHash);
    if (!row) throw new Error("Share link not found");
    const token = randomBytes(24).toString("base64url");
    this.database.prepare("UPDATE share_links SET token_hash = ?, updated_at = ? WHERE name = ? AND user_key = ?")
      .run(hashToken(token), this.now().toISOString(), shareName, userKeyHash);
    return { name: shareName, token };
  }

  revoke(userId, name) {
    const shareName = normalizeName(name);
    const userKeyHash = createHash("sha256").update(userId).digest("hex");
    const result = this.database.prepare("DELETE FROM share_links WHERE name = ? AND user_key = ?").run(shareName, userKeyHash);
    if (result.changes === 0) throw new Error("Share link not found");
    return { revoked: shareName };
  }

  list(userId) {
    const userKeyHash = createHash("sha256").update(userId).digest("hex");
    return this.database.prepare(
      "SELECT name, folder, title, created_at, updated_at FROM share_links WHERE user_key = ? ORDER BY name"
    ).all(userKeyHash);
  }
}


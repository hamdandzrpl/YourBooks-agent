import { createHash, randomUUID } from "node:crypto";
import { chmodSync, mkdirSync } from "node:fs";
import { lstat, mkdir, mkdtemp, readdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep } from "node:path";

const readableExtensions = new Set([".txt", ".md", ".csv", ".json", ".log"]);
const writableExtensions = new Set([".txt", ".md", ".csv", ".json"]);
const excludedNames = new Set([".versions", ".trash"]);
const maxTextBytes = 2_000_000;

function userKey(userId) {
  if (typeof userId !== "string" || !/^\d{8,15}$/.test(userId)) {
    throw new Error("A valid authorized user ID is required");
  }
  return createHash("sha256").update(userId).digest("hex");
}

function normalizeRelativePath(value, { allowEmpty = false } = {}) {
  if (typeof value !== "string") throw new Error("File path must be text");
  const input = value.trim().replaceAll("\\", "/");
  if (allowEmpty && !input) return "";
  if (!input || input.length > 240 || input.startsWith("/") || isAbsolute(input)) {
    throw new Error("Use a non-empty path relative to your workspace");
  }
  const parts = input.split("/");
  if (parts.some((part) => !part || part === "." || part === ".." || part.startsWith("."))) {
    throw new Error("Hidden, empty, and parent-directory path segments are not allowed");
  }
  if (parts.some((part) => /[<>:"|?*\u0000-\u001f]/.test(part))) {
    throw new Error("Path contains unsupported characters");
  }
  return parts.join("/");
}

function safeDisplayName(filename) {
  const leaf = basename(String(filename ?? "").replaceAll("\\", "/"))
    .replace(/[\u0000-\u001f]/g, "")
    .trim();
  const safe = leaf.replace(/[<>:"|?*]/g, "_").replace(/^\.+/, "").slice(0, 120);
  return safe || `attachment-${randomUUID().slice(0, 8)}`;
}

function metadata(record) {
  return {
    id: record.id,
    name: basename(record.relative_path),
    path: record.relative_path,
    mimeType: record.mime_type,
    sizeBytes: record.size_bytes,
    sha256: record.sha256,
    sourceMessageId: record.source_message_id,
    createdAt: record.created_at,
    updatedAt: record.updated_at,
  };
}

export class FileWorkspace {
  constructor({ dataDirectory, maxFileBytes = 20 * 1024 * 1024, now = () => new Date() }) {
    this.rootDirectory = resolve(dataDirectory);
    this.workspacesDirectory = join(this.rootDirectory, "workspaces");
    mkdirSync(this.workspacesDirectory, { recursive: true, mode: 0o700 });
    this.maxFileBytes = maxFileBytes;
    this.now = now;
    const databasePath = join(this.rootDirectory, "workspace.sqlite");
    this.database = new DatabaseSync(databasePath);
    this.database.exec("PRAGMA foreign_keys = ON; PRAGMA journal_mode = WAL; PRAGMA busy_timeout = 5000;");
    this.database.exec(`
      CREATE TABLE IF NOT EXISTS files (
        id TEXT PRIMARY KEY,
        user_key TEXT NOT NULL,
        relative_path TEXT NOT NULL,
        mime_type TEXT NOT NULL,
        size_bytes INTEGER NOT NULL,
        sha256 TEXT NOT NULL,
        source_message_id TEXT,
        status TEXT NOT NULL CHECK (status IN ('active', 'deleted')),
        trash_path TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        UNIQUE(user_key, relative_path)
      );
      CREATE UNIQUE INDEX IF NOT EXISTS files_source_message_unique
        ON files(user_key, source_message_id) WHERE source_message_id IS NOT NULL;
      CREATE TABLE IF NOT EXISTS delete_requests (
        id TEXT PRIMARY KEY,
        user_key TEXT NOT NULL,
        file_id TEXT NOT NULL REFERENCES files(id),
        expires_at TEXT NOT NULL,
        status TEXT NOT NULL CHECK (status IN ('pending', 'completed', 'expired'))
      );
      CREATE TABLE IF NOT EXISTS file_events (
        id INTEGER PRIMARY KEY,
        user_key TEXT NOT NULL,
        action TEXT NOT NULL,
        file_id TEXT NOT NULL,
        created_at TEXT NOT NULL
      );
    `);
    try { chmodSync(databasePath, 0o600); } catch {}
  }

  close() {
    this.database.close();
  }

  async exportBackupSnapshot(userId, maxBytes = 128 * 1024 * 1024, maxEntries = 20_000) {
    const key = userKey(userId);
    const root = await this.ensureWorkspace(userId);
    const entries = [];
    let totalBytes = 0;
    const walk = async (directory) => {
      for (const entry of await readdir(directory, { withFileTypes: true })) {
        const absolutePath = join(directory, entry.name);
        if (entry.isSymbolicLink()) throw new Error("Workspace backup refused a symbolic link");
        if (entry.isDirectory()) {
          await walk(absolutePath);
          continue;
        }
        if (!entry.isFile()) throw new Error("Workspace backup found an unsupported filesystem entry");
        const bytes = await readFile(absolutePath);
        totalBytes += bytes.length;
        if (totalBytes > maxBytes) throw new Error("Workspace backup exceeds the 128 MB archive limit; deselect File Explorer or remove large files first");
        entries.push({ path: relative(root, absolutePath).split(sep).join("/"), bytes });
        if (entries.length > maxEntries) throw new Error("Workspace backup contains too many files; remove some files and try again");
      }
    };
    await walk(root);
    return {
      entries,
      totalBytes,
      index: {
        files: this.database.prepare("SELECT * FROM files WHERE user_key = ? ORDER BY relative_path").all(key),
        delete_requests: this.database.prepare("SELECT * FROM delete_requests WHERE user_key = ? ORDER BY id").all(key),
        file_events: this.database.prepare("SELECT * FROM file_events WHERE user_key = ? ORDER BY id").all(key),
      },
    };
  }

  async restoreBackupSnapshot(userId, { entries, index }) {
    const key = userKey(userId);
    if (!Array.isArray(entries) || entries.length > 20_000 || !index
      || !Array.isArray(index.files) || !Array.isArray(index.delete_requests) || !Array.isArray(index.file_events)) {
      throw new Error("File Explorer backup index is invalid");
    }
    const safePath = (value) => {
      if (typeof value !== "string" || !value || value.length > 240 || value.startsWith("/") || value.includes("\\")) {
        throw new Error("Backup contains an unsafe workspace path");
      }
      const parts = value.split("/");
      if (parts.some((part) => !part || part === "." || part === ".." || /[<>:"|?*\u0000-\u001f]/.test(part))) {
        throw new Error("Backup contains an unsafe workspace path");
      }
      const internalRoot = parts[0] === ".trash" || parts[0] === ".versions";
      if (parts.some((part, position) => part.startsWith(".") && !(position === 0 && internalRoot))) {
        throw new Error("Backup contains an unsupported hidden workspace path");
      }
      return parts.join("/");
    };
    const fileRows = index.files;
    const activePaths = new Set();
    const trashPaths = new Set();
    const fileIds = new Set();
    const entryByPath = new Map();
    let totalBytes = 0;
    for (const entry of entries) {
      if (!entry || !Buffer.isBuffer(entry.bytes)) throw new Error("File Explorer backup entry is invalid");
      const path = safePath(entry.path);
      if (entryByPath.has(path)) throw new Error("Backup contains duplicate workspace paths");
      totalBytes += entry.bytes.length;
      if (totalBytes > 128 * 1024 * 1024) throw new Error("File Explorer backup exceeds the 128 MB restore limit");
      entryByPath.set(path, entry.bytes);
    }
    for (const row of fileRows) {
      if (!row || row.user_key !== key || !/^F-[A-F0-9]{8}$/.test(row.id)
        || !["active", "deleted"].includes(row.status) || fileIds.has(row.id)) {
        throw new Error("File Explorer backup contains invalid or foreign file records");
      }
      fileIds.add(row.id);
      const expectedPath = row.status === "active" ? safePath(row.relative_path) : safePath(row.trash_path);
      if (row.status === "active" && expectedPath.startsWith(".")) {
        throw new Error("Active file uses a reserved workspace path");
      }
      if (row.status === "deleted" && !expectedPath.startsWith(".trash/")) {
        throw new Error("Deleted file is not stored in the workspace Trash");
      }
      (row.status === "active" ? activePaths : trashPaths).add(expectedPath);
      const bytes = entryByPath.get(expectedPath);
      if (!bytes || bytes.length !== row.size_bytes || createHash("sha256").update(bytes).digest("hex") !== row.sha256) {
        throw new Error("File Explorer backup content does not match its file index");
      }
      if (this.database.prepare("SELECT 1 FROM files WHERE id = ? AND user_key <> ?").get(row.id, key)) {
        throw new Error("Backup contains a file ID owned by another user");
      }
    }
    for (const path of entryByPath.keys()) {
      const parts = path.split("/");
      if (parts[0] === ".versions") {
        if (parts.length < 3 || !fileIds.has(parts[1])) throw new Error("Backup revision is not linked to a known workspace file");
      } else if (parts[0] === ".trash") {
        if (!trashPaths.has(path)) throw new Error("Backup contains an unindexed Trash file");
      } else if (!activePaths.has(path)) {
        throw new Error("Backup contains an unindexed workspace file");
      }
    }
    for (const row of [...index.delete_requests, ...index.file_events]) {
      if (!row || row.user_key !== key) throw new Error("File Explorer backup index contains data for another user");
    }
    for (const request of index.delete_requests) {
      if (!fileIds.has(request.file_id)) throw new Error("Backup contains a delete request for an unknown file");
      const existing = this.database.prepare("SELECT user_key FROM delete_requests WHERE id = ?").get(request.id);
      if (existing && existing.user_key !== key) throw new Error("Backup contains a delete request owned by another user");
    }

    const root = await this.ensureWorkspace(userId);
    const stagingRoot = await mkdtemp(join(this.workspacesDirectory, ".restore-"));
    const oldRoot = `${root}.previous-${randomUUID()}`;
    let movedOldRoot = false;
    let movedStagingRoot = false;
    let committed = false;
    try {
      for (const [path, bytes] of entryByPath) {
        const target = resolve(stagingRoot, ...path.split("/"));
        if (!target.startsWith(`${stagingRoot}${sep}`)) throw new Error("Backup path escapes its staging directory");
        await mkdir(dirname(target), { recursive: true, mode: 0o700 });
        await writeFile(target, bytes, { flag: "wx", mode: 0o600 });
      }
      await rename(root, oldRoot);
      movedOldRoot = true;
      await rename(stagingRoot, root);
      movedStagingRoot = true;
      this.database.exec("BEGIN IMMEDIATE");
      try {
        this.database.prepare("DELETE FROM delete_requests WHERE user_key = ?").run(key);
        this.database.prepare("DELETE FROM file_events WHERE user_key = ?").run(key);
        this.database.prepare("DELETE FROM files WHERE user_key = ?").run(key);
        for (const row of fileRows) {
          this.database.prepare(`INSERT INTO files
            (id, user_key, relative_path, mime_type, size_bytes, sha256, source_message_id, status, trash_path, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(row.id, row.user_key, row.relative_path, row.mime_type, row.size_bytes, row.sha256, row.source_message_id, row.status, row.trash_path, row.created_at, row.updated_at);
        }
        for (const row of index.delete_requests) {
          this.database.prepare("INSERT INTO delete_requests (id, user_key, file_id, expires_at, status) VALUES (?, ?, ?, ?, ?)")
            .run(row.id, row.user_key, row.file_id, row.expires_at, row.status);
        }
        for (const row of index.file_events) {
          this.database.prepare("INSERT INTO file_events (user_key, action, file_id, created_at) VALUES (?, ?, ?, ?)")
            .run(row.user_key, row.action, row.file_id, row.created_at);
        }
        this.database.exec("COMMIT");
        committed = true;
      } catch (error) {
        if (this.database.isTransaction) this.database.exec("ROLLBACK");
        throw error;
      }
      await rm(oldRoot, { recursive: true, force: true }).catch(() => {});
      return { files: fileRows.length, bytes: totalBytes };
    } catch (error) {
      if (committed) throw error;
      if (movedStagingRoot) await rename(root, stagingRoot).catch(() => {});
      if (movedOldRoot) await rename(oldRoot, root).catch(() => {});
      await rm(stagingRoot, { recursive: true, force: true });
      throw error;
    }
  }

  async ensureWorkspace(userId) {
    const root = join(this.workspacesDirectory, userKey(userId));
    await mkdir(root, { recursive: true, mode: 0o700 });
    return root;
  }

  async resolveSafePath(userId, relativePath, { allowEmpty = false } = {}) {
    const normalized = normalizeRelativePath(relativePath, { allowEmpty });
    const root = await this.ensureWorkspace(userId);
    const target = normalized ? resolve(root, ...normalized.split("/")) : root;
    if (target !== root && !target.startsWith(`${root}${sep}`)) {
      throw new Error("File path escapes the user workspace");
    }

    let current = root;
    for (const part of normalized.split("/").filter(Boolean)) {
      current = join(current, part);
      try {
        const info = await lstat(current);
        if (info.isSymbolicLink()) throw new Error("Symbolic links are not allowed in workspaces");
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
      }
    }
    return { root, target, relativePath: normalized };
  }

  async createFolder(userId, relativePath) {
    const { target, relativePath: normalized } = await this.resolveSafePath(userId, relativePath);
    await mkdir(target, { recursive: true, mode: 0o700 });
    return { path: normalized, created: true };
  }

  async listFiles(userId, folder = "") {
    const key = userKey(userId);
    const normalizedFolder = normalizeRelativePath(folder, { allowEmpty: true });
    const prefix = normalizedFolder ? `${normalizedFolder}/` : "";
    const rows = this.database.prepare(`
      SELECT id, relative_path, mime_type, size_bytes, sha256, source_message_id, created_at, updated_at
      FROM files WHERE user_key = ? AND status = 'active' AND relative_path LIKE ? ORDER BY relative_path LIMIT 500
    `).all(key, `${prefix}%`);
    return rows.map(metadata);
  }

  async listDirectory(userId, folder = "") {
    const { target, relativePath } = await this.resolveSafePath(userId, folder, { allowEmpty: true });
    const filesByPath = new Map((await this.listFiles(userId, relativePath)).map((file) => [file.path, file]));
    const entries = await readdir(target, { withFileTypes: true });
    return entries.flatMap((entry) => {
      if (entry.name.startsWith(".") || entry.isSymbolicLink()) return [];
      const path = relativePath ? `${relativePath}/${entry.name}` : entry.name;
      if (entry.isDirectory()) return [{ type: "folder", name: entry.name, path }];
      const file = filesByPath.get(path);
      return entry.isFile() && file ? [{ ...file, type: "file" }] : [];
    }).sort((left, right) => Number(right.type === "folder") - Number(left.type === "folder")
      || left.name.localeCompare(right.name, "id"));
  }

  async searchFiles(userId, query) {
    if (typeof query !== "string" || !query.trim()) throw new Error("Search query cannot be empty");
    const needle = query.trim().toLocaleLowerCase();
    const files = await this.listFiles(userId);
    const results = [];
    for (const file of files) {
      if (file.path.toLocaleLowerCase().includes(needle)) {
        results.push(file);
      } else if (readableExtensions.has(extname(file.name).toLowerCase()) && file.sizeBytes <= maxTextBytes) {
        const content = await this.readTextFile(userId, file.id).catch(() => "");
        if (content.toLocaleLowerCase().includes(needle)) results.push(file);
      }
      if (results.length >= 50) break;
    }
    return results;
  }

  async saveAttachment(userId, { filename, bytes, mimeType = "application/octet-stream", sha256 = null, sourceMessageId }) {
    const key = userKey(userId);
    if (!Buffer.isBuffer(bytes) || bytes.length === 0 || bytes.length > this.maxFileBytes) {
      throw new Error("Incoming file is empty or exceeds the configured size limit");
    }
    if (!sourceMessageId) throw new Error("Incoming file must have a WhatsApp message ID");

    const existing = this.database.prepare(`
      SELECT id, relative_path, mime_type, size_bytes, sha256, source_message_id, created_at, updated_at
      FROM files WHERE user_key = ? AND source_message_id = ?
    `).get(key, sourceMessageId);
    if (existing) return metadata(existing);

    const safeName = safeDisplayName(filename);
    const root = await this.ensureWorkspace(userId);
    const inbox = join(root, "inbox");
    await mkdir(inbox, { recursive: true, mode: 0o700 });
    const extension = extname(safeName);
    const stem = safeName.slice(0, safeName.length - extension.length) || "attachment";
    let relativePath = `inbox/${safeName}`;
    let target = join(root, ...relativePath.split("/"));
    for (let suffix = 2; await this.pathExists(target); suffix += 1) {
      relativePath = `inbox/${stem} (${suffix})${extension}`;
      target = join(root, ...relativePath.split("/"));
    }

    const id = `F-${randomUUID().slice(0, 8).toUpperCase()}`;
    const actualHash = createHash("sha256").update(bytes).digest("hex");
    if (sha256 && /^[a-f0-9]{64}$/i.test(sha256) && actualHash.toLowerCase() !== sha256.toLowerCase()) {
      throw new Error("Incoming file checksum does not match WhatsApp metadata");
    }
    const now = this.now().toISOString();
    await writeFile(target, bytes, { flag: "wx", mode: 0o600 });
    try {
      this.database.prepare(`
        INSERT INTO files
          (id, user_key, relative_path, mime_type, size_bytes, sha256, source_message_id, status, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, 'active', ?, ?)
      `).run(id, key, relativePath, mimeType, bytes.length, actualHash, sourceMessageId, now, now);
      this.addEvent(key, "attachment_saved", id);
    } catch (error) {
      await this.removePath(target);
      throw error;
    }
    return this.getFile(userId, id);
  }

  async createTextFile(userId, relativePath, content, { mimeType } = {}) {
    const { root, target, relativePath: normalized } = await this.resolveSafePath(userId, relativePath);
    const extension = extname(normalized).toLowerCase();
    if (!writableExtensions.has(extension)) {
      throw new Error("Text creation supports only .txt, .md, .csv, and .json; use a document tool for Office files");
    }
    const bytes = Buffer.from(String(content), "utf8");
    if (bytes.length > maxTextBytes) throw new Error("Text file exceeds the 2 MB limit");
    await mkdir(dirname(target), { recursive: true, mode: 0o700 });
    const id = `F-${randomUUID().slice(0, 8).toUpperCase()}`;
    const now = this.now().toISOString();
    await writeFile(target, bytes, { flag: "wx", mode: 0o600 });
    try {
      this.insertFile({
        id, userKey: userKey(userId), relativePath: normalized,
        mimeType: mimeType ?? this.mimeForExtension(extension),
        bytes, sourceMessageId: null, now,
      });
      this.addEvent(userKey(userId), "file_created", id);
    } catch (error) {
      await this.removePath(target);
      throw error;
    }
    return this.getFile(userId, id);
  }

  async saveGeneratedFile(userId, relativePath, bytes, mimeType) {
    if (!Buffer.isBuffer(bytes) || bytes.length === 0 || bytes.length > this.maxFileBytes) {
      throw new Error("Generated file is empty or exceeds the configured size limit");
    }
    const { root, target, relativePath: normalized } = await this.resolveSafePath(userId, relativePath);
    if (await this.pathExists(target)) throw new Error("A file already exists at the requested path");
    await mkdir(dirname(target), { recursive: true, mode: 0o700 });
    const id = `F-${randomUUID().slice(0, 8).toUpperCase()}`;
    const now = this.now().toISOString();
    await writeFile(target, bytes, { flag: "wx", mode: 0o600 });
    try {
      this.insertFile({
        id, userKey: userKey(userId), relativePath: normalized,
        mimeType: mimeType || "application/octet-stream",
        bytes, sourceMessageId: null, now,
      });
      this.addEvent(userKey(userId), "file_created", id);
    } catch (error) {
      await this.removePath(target);
      throw error;
    }
    return this.getFile(userId, id);
  }

  async replaceGeneratedFile(userId, fileId, bytes) {
    if (!Buffer.isBuffer(bytes) || bytes.length === 0 || bytes.length > this.maxFileBytes) {
      throw new Error("Replacement file is empty or exceeds the configured size limit");
    }
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    const key = userKey(userId);
    const { root, target } = await this.resolveSafePath(userId, file.path);
    const current = await readFile(target);
    const versionsDirectory = join(root, ".versions", file.id);
    await mkdir(versionsDirectory, { recursive: true, mode: 0o700 });
    const backupPath = join(versionsDirectory, `${this.now().getTime()}-${randomUUID()}.bak`);
    const tempPath = join(dirname(target), `.${basename(target)}-${randomUUID()}.tmp`);
    const displacedPath = join(dirname(target), `.${basename(target)}-${randomUUID()}.old`);
    await writeFile(backupPath, current, { flag: "wx", mode: 0o600 });
    await writeFile(tempPath, bytes, { flag: "wx", mode: 0o600 });
    await rename(target, displacedPath);
    try {
      await rename(tempPath, target);
      const now = this.now().toISOString();
      this.database.prepare("UPDATE files SET size_bytes = ?, sha256 = ?, updated_at = ? WHERE id = ? AND user_key = ?")
        .run(bytes.length, createHash("sha256").update(bytes).digest("hex"), now, file.id, key);
      await this.removePath(displacedPath);
      this.addEvent(key, "file_replaced", file.id);
      return this.getFile(userId, file.id);
    } catch (error) {
      await this.removePath(target);
      await rename(displacedPath, target).catch(() => {});
      await this.removePath(tempPath);
      throw error;
    }
  }

  async saveGeneratedFiles(userId, files) {
    if (!Array.isArray(files) || files.length < 1 || files.length > 1000) {
      throw new Error("Provide 1-1000 files to save");
    }
    const key = userKey(userId);
    const paths = new Set();
    const staged = [];
    const committedPaths = [];

    try {
      for (const file of files) {
        if (!file || !Buffer.isBuffer(file.bytes) || file.bytes.length > this.maxFileBytes) {
          throw new Error("Generated file is invalid or exceeds the configured size limit");
        }
        const { root, target, relativePath } = await this.resolveSafePath(userId, file.relativePath);
        if (paths.has(relativePath) || await this.pathExists(target)) {
          throw new Error(`A file already exists at ${relativePath}`);
        }
        paths.add(relativePath);
        await mkdir(dirname(target), { recursive: true, mode: 0o700 });
        const tempPath = join(dirname(target), `.${basename(target)}-${randomUUID()}.tmp`);
        await writeFile(tempPath, file.bytes, { flag: "wx", mode: 0o600 });
        staged.push({ ...file, root, target, tempPath, relativePath });
      }

      for (const file of staged) {
        if (await this.pathExists(file.target)) {
          throw new Error(`A file already exists at ${file.relativePath}`);
        }
        await rename(file.tempPath, file.target);
        committedPaths.push(file.target);
      }

      const now = this.now().toISOString();
      const created = staged.map((file) => ({ ...file, id: `F-${randomUUID().slice(0, 8).toUpperCase()}` }));
      this.database.exec("BEGIN IMMEDIATE");
      try {
        for (const file of created) {
          this.insertFile({
            id: file.id,
            userKey: key,
            relativePath: file.relativePath,
            mimeType: file.mimeType || this.mimeForExtension(extname(file.relativePath).toLowerCase()),
            bytes: file.bytes,
            sourceMessageId: null,
            now,
          });
          this.addEvent(key, "file_created", file.id);
        }
        this.database.exec("COMMIT");
      } catch (error) {
        if (this.database.isTransaction) this.database.exec("ROLLBACK");
        throw error;
      }
      return created.map((file) => this.getFile(userId, file.id));
    } catch (error) {
      for (const file of staged) await this.removePath(file.tempPath);
      for (const path of committedPaths) await this.removePath(path);
      throw error;
    }
  }

  async readTextFile(userId, fileId) {
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    if (!readableExtensions.has(extname(file.name).toLowerCase()) || file.sizeBytes > maxTextBytes) {
      throw new Error("This file is not a supported, size-limited text file");
    }
    const { target } = await this.resolveSafePath(userId, file.path);
    const bytes = await readFile(target);
    if (bytes.includes(0)) throw new Error("Binary files cannot be read as text");
    return bytes.toString("utf8");
  }

  async editTextFile(userId, fileId, content) {
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    if (!writableExtensions.has(extname(file.name).toLowerCase())) {
      throw new Error("Only .txt, .md, .csv, and .json files can be edited as text");
    }
    const bytes = Buffer.from(String(content), "utf8");
    if (bytes.length > maxTextBytes) throw new Error("Text file exceeds the 2 MB limit");
    const { root, target } = await this.resolveSafePath(userId, file.path);
    const current = await readFile(target);
    const versionsDirectory = join(root, ".versions", file.id);
    await mkdir(versionsDirectory, { recursive: true, mode: 0o700 });
    await writeFile(join(versionsDirectory, `${this.now().getTime()}-${randomUUID()}.bak`), current, { flag: "wx", mode: 0o600 });
    const tempPath = join(dirname(target), `.${basename(target)}-${randomUUID()}.tmp`);
    await writeFile(tempPath, bytes, { flag: "wx", mode: 0o600 });
    await rename(tempPath, target);
    const now = this.now().toISOString();
    this.database.prepare("UPDATE files SET size_bytes = ?, sha256 = ?, updated_at = ? WHERE id = ? AND user_key = ?")
      .run(bytes.length, createHash("sha256").update(bytes).digest("hex"), now, file.id, userKey(userId));
    this.addEvent(userKey(userId), "file_edited", file.id);
    return this.getFile(userId, file.id);
  }

  async renameFile(userId, fileId, newRelativePath) {
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    const { root, target: newTarget, relativePath: normalized } = await this.resolveSafePath(userId, newRelativePath);
    const oldTarget = (await this.resolveSafePath(userId, file.path)).target;
    if (oldTarget === newTarget) return file;
    if (await this.pathExists(newTarget)) throw new Error("A file or folder already exists at the new path");
    await mkdir(dirname(newTarget), { recursive: true, mode: 0o700 });
    await rename(oldTarget, newTarget);
    try {
      this.database.prepare("UPDATE files SET relative_path = ?, updated_at = ? WHERE id = ? AND user_key = ?")
        .run(normalized, this.now().toISOString(), file.id, userKey(userId));
    } catch (error) {
      await rename(newTarget, oldTarget).catch(() => {});
      throw error;
    }
    this.addEvent(userKey(userId), "file_renamed", file.id);
    return this.getFile(userId, file.id);
  }

  async copyFile(userId, fileId, newRelativePath) {
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    const bytes = await readFile(await this.getFilePath(userId, file.id));
    const [copy] = await this.saveGeneratedFiles(userId, [{
      relativePath: newRelativePath,
      bytes,
      mimeType: file.mimeType,
    }]);
    this.addEvent(userKey(userId), "file_copied", copy.id);
    return copy;
  }

  requestDelete(userId, fileId) {
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    const id = `X-${randomUUID().slice(0, 8).toUpperCase()}`;
    const now = this.now();
    const expiresAt = new Date(now.getTime() + 10 * 60 * 1000).toISOString();
    this.database.prepare(`
      INSERT INTO delete_requests (id, user_key, file_id, expires_at, status) VALUES (?, ?, ?, ?, 'pending')
    `).run(id, userKey(userId), file.id, expiresAt);
    return { id, file, expiresAt };
  }

  async confirmDelete(userId, requestId) {
    if (!requestId) throw new Error("request_id is required");
    const key = userKey(userId);
    const request = this.database.prepare(`
      SELECT id, file_id, expires_at, status FROM delete_requests WHERE id = ? AND user_key = ?
    `).get(requestId.toUpperCase(), key);
    if (!request || request.status !== "pending") throw new Error("Delete request not found or already used");
    if (request.expires_at <= this.now().toISOString()) {
      this.database.prepare("UPDATE delete_requests SET status = 'expired' WHERE id = ?").run(request.id);
      throw new Error("Delete request expired; request it again");
    }
    const file = this.getFile(userId, request.file_id);
    if (!file) throw new Error("File not found in this user's workspace");
    const { root, target } = await this.resolveSafePath(userId, file.path);
    const trashDirectory = join(root, ".trash");
    await mkdir(trashDirectory, { recursive: true, mode: 0o700 });
    const trashPath = join(trashDirectory, `${file.id}-${basename(file.path)}`);
    await rename(target, trashPath);
    this.database.exec("BEGIN IMMEDIATE");
    try {
      this.database.prepare("UPDATE files SET status = 'deleted', trash_path = ?, updated_at = ? WHERE id = ? AND user_key = ?")
        .run(relative(root, trashPath), this.now().toISOString(), file.id, key);
      this.database.prepare("UPDATE delete_requests SET status = 'completed' WHERE id = ?").run(request.id);
      this.addEvent(key, "file_moved_to_trash", file.id);
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      await rename(trashPath, target).catch(() => {});
      throw error;
    }
    return file;
  }

  async softDeleteFile(userId, fileId) {
    const key = userKey(userId);
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    return this.trashFile(userId, file);
  }

  /** Find the single active file whose name matches, so the agent can delete
   *  by name without having seen a file id. Fails loudly when the name is
   *  ambiguous across folders. */
  findActiveByName(userId, filename) {
    const key = userKey(userId);
    const wanted = basename(String(filename).replaceAll("\\", "/"));
    const matches = this.database.prepare(`
      SELECT id FROM files WHERE user_key = ? AND status = 'active'
    `).all(key).map((row) => this.getFile(userId, row.id))
      .filter((file) => basename(file.path) === wanted);
    if (matches.length === 0) throw new Error(`File not found with name: ${filename}`);
    if (matches.length > 1) {
      throw new Error(`Multiple files named ${filename}; pass file_id instead`);
    }
    return matches[0];
  }

  async softDeleteFileByName(userId, filename) {
    if (!filename) throw new Error("filename is required");
    return this.trashFile(userId, this.findActiveByName(userId, filename));
  }

  async trashFile(userId, file) {
    const key = userKey(userId);
    const { root, target } = await this.resolveSafePath(userId, file.path);
    const trashDirectory = join(root, ".trash");
    await mkdir(trashDirectory, { recursive: true, mode: 0o700 });
    const trashPath = join(trashDirectory, `${file.id}-${basename(file.path)}`);
    await rename(target, trashPath);
    this.database.exec("BEGIN IMMEDIATE");
    try {
      this.database.prepare("UPDATE files SET status = 'deleted', trash_path = ?, updated_at = ? WHERE id = ? AND user_key = ?")
        .run(relative(root, trashPath), this.now().toISOString(), file.id, key);
      this.addEvent(key, "file_moved_to_trash", file.id);
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      await rename(trashPath, target).catch(() => {});
      throw error;
    }
    return file;
  }

  /** Find the most recently trashed file whose original name matches, so the
   *  agent can restore by name without ever having seen a file id. */
  findDeletedByName(userId, filename) {
    const key = userKey(userId);
    const rows = this.database.prepare(`
      SELECT id, relative_path, trash_path FROM files
      WHERE user_key = ? AND status = 'deleted'
      ORDER BY updated_at DESC
    `).all(key);
    const wanted = basename(String(filename).replaceAll("\\", "/"));
    return rows.find((row) => basename(row.relative_path) === wanted) ?? null;
  }

  async restoreFile(userId, fileId) {
    if (!fileId) throw new Error("file_id is required");
    const key = userKey(userId);
    const row = this.database.prepare(`
      SELECT id, relative_path, trash_path FROM files WHERE id = ? AND user_key = ? AND status = 'deleted'
    `).get(fileId.toUpperCase(), key);
    if (!row) throw new Error("Deleted file not found");
    return this.restoreDeletedRow(userId, row);
  }

  async restoreFileByName(userId, filename) {
    if (!filename) throw new Error("filename is required");
    const row = this.findDeletedByName(userId, filename);
    if (!row) throw new Error(`Deleted file not found: ${filename}`);
    return this.restoreDeletedRow(userId, row);
  }

  async restoreDeletedRow(userId, row) {
    const key = userKey(userId);
    const root = await this.ensureWorkspace(userId);
    const target = (await this.resolveSafePath(userId, row.relative_path)).target;
    const trashPath = resolve(root, row.trash_path);
    if (!trashPath.startsWith(`${join(root, ".trash")}${sep}`)) throw new Error("Invalid trash record");
    if (await this.pathExists(target)) throw new Error("The original file path is occupied; rename or move that file first");
    await mkdir(dirname(target), { recursive: true, mode: 0o700 });
    await rename(trashPath, target);
    this.database.prepare("UPDATE files SET status = 'active', trash_path = NULL, updated_at = ? WHERE id = ? AND user_key = ?")
      .run(this.now().toISOString(), row.id, key);
    this.addEvent(key, "file_restored", row.id);
    return this.getFile(userId, row.id);
  }

  getFile(userId, fileId) {
    if (!fileId) throw new Error("file_id is required");
    const row = this.database.prepare(`
      SELECT id, relative_path, mime_type, size_bytes, sha256, source_message_id, created_at, updated_at
      FROM files WHERE id = ? AND user_key = ? AND status = 'active'
    `).get(fileId.toUpperCase(), userKey(userId));
    return row ? metadata(row) : null;
  }

  async getFilePath(userId, fileId) {
    const file = this.getFile(userId, fileId);
    if (!file) throw new Error("File not found in this user's workspace");
    return (await this.resolveSafePath(userId, file.path)).target;
  }

  insertFile({ id, userKey: key, relativePath, mimeType, bytes, sourceMessageId, now }) {
    this.database.prepare(`
      INSERT INTO files
        (id, user_key, relative_path, mime_type, size_bytes, sha256, source_message_id, status, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, 'active', ?, ?)
    `).run(id, key, relativePath, mimeType, bytes.length, createHash("sha256").update(bytes).digest("hex"), sourceMessageId, now, now);
  }

  addEvent(key, action, fileId) {
    this.database.prepare("INSERT INTO file_events (user_key, action, file_id, created_at) VALUES (?, ?, ?, ?)")
      .run(key, action, fileId, this.now().toISOString());
  }

  async pathExists(path) {
    try { await lstat(path); return true; } catch (error) { if (error.code === "ENOENT") return false; throw error; }
  }

  async removePath(path) {
    const { unlink } = await import("node:fs/promises");
    await unlink(path).catch(() => {});
  }

  mimeForExtension(extension) {
    return ({ ".txt": "text/plain", ".md": "text/markdown", ".csv": "text/csv", ".json": "application/json" })[extension]
      ?? "application/octet-stream";
  }
}
import { randomUUID } from "node:crypto";
import { chmod, lstat, mkdir, readdir, readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import { dirname, join, relative, resolve, sep } from "node:path";
import PizZip from "pizzip";
import { loadSystemSkills, setSystemSkillEnabled } from "./skills.js";

const supportedComponents = new Set(["ai", "finance", "files", "system", "whatsapp"]);
const maxArchiveBytes = 128 * 1024 * 1024;
const maxArchiveEntries = 20_000;
const retainedArchives = 5;

function jsonBytes(value) {
  return Buffer.from(`${JSON.stringify(value, null, 2)}\n`, "utf8");
}

export class BackupService {
  constructor({ dataDirectory, financeStore, fileWorkspace, config = {} }) {
    this.dataDirectory = resolve(dataDirectory);
    this.backupDirectory = join(this.dataDirectory, "backups");
    this.financeStore = financeStore;
    this.fileWorkspace = fileWorkspace;
    this.config = config;
  }

  async createBackup(userId, components) {
    if (!Array.isArray(components) || !components.length
      || components.some((component) => typeof component !== "string" || !supportedComponents.has(component))) {
      throw new Error("Choose at least one supported backup component");
    }
    const selected = [...new Set(components)].sort();
    const archive = new PizZip();
    let inputBytes = 0;
    let entryCount = 0;
    const add = (path, bytes) => {
      inputBytes += bytes.length;
      entryCount += 1;
      if (inputBytes > maxArchiveBytes) throw new Error("Backup exceeds the 128 MB uncompressed limit; deselect large components or files");
      if (entryCount > maxArchiveEntries) throw new Error("Backup contains too many files");
      archive.file(path, bytes);
    };

    if (selected.includes("ai") || selected.includes("finance")) {
      const snapshot = this.financeStore.exportBackupSnapshot(userId, selected);
      if (snapshot.ai) add("owner-data/ai-memory-and-conversations.json", jsonBytes(snapshot.ai));
      if (snapshot.finance) add("owner-data/finance-ledger-and-audit.json", jsonBytes(snapshot.finance));
    }

    if (selected.includes("files")) {
      const workspace = await this.fileWorkspace.exportBackupSnapshot(userId, maxArchiveBytes - inputBytes, maxArchiveEntries);
      for (const entry of workspace.entries) add(`file-explorer/${entry.path}`, entry.bytes);
      add("file-explorer/workspace-index.json", jsonBytes(workspace.index));
    }

    const componentStatus = {};
    if (selected.includes("system")) {
      const skills = loadSystemSkills(this.config.skillsDirectory);
      add("system/skills.json", jsonBytes(skills));
      add("system/runtime-settings.json", jsonBytes({
        model: this.config.aiModel ?? null,
        reasoningEffort: this.config.aiReasoningEffort ?? null,
        timeZone: this.config.timeZone ?? null,
        currency: this.config.defaultCurrency ?? null,
        maxFileBytes: this.config.maxFileBytes ?? null,
      }));
    }

    if (selected.includes("whatsapp")) {
      const authDirectory = join(this.dataDirectory, "baileys-auth");
      let authInfo;
      try {
        authInfo = await lstat(authDirectory);
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
        componentStatus.whatsapp = "not present on this server";
      }
      if (authInfo) {
        if (authInfo.isSymbolicLink() || !authInfo.isDirectory()) {
          throw new Error("WhatsApp credentials must be a regular directory inside DATA_DIR");
        }
        await this.addDirectory(add, authDirectory, "sensitive/whatsapp-auth");
        componentStatus.whatsapp = "included; contains reusable WhatsApp session credentials";
      }
    }

    add("BACKUP-INFO.txt", Buffer.from([
      "WhatsApp Finance Agent owner-scoped backup",
      `Owner WhatsApp number: ${userId}`,
      `Created: ${new Date().toISOString()}`,
      `Components: ${selected.join(", ")}`,
      "API keys, panel passwords, .env files, and other owners' data are not included.",
      selected.includes("whatsapp") ? "Sensitive: this archive may contain reusable WhatsApp session credentials. Store it privately." : "WhatsApp session credentials were not selected.",
      "Finance and AI database content is exported as owner-scoped JSON. File Explorer includes workspace files, Trash, revisions, and an owner-scoped index.",
      "This archive is an export; it is not encrypted and does not include application source code.",
      "",
    ].join("\n"), "utf8"));

    const bytes = archive.generate({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    await mkdir(this.backupDirectory, { recursive: true, mode: 0o700 });
    const filename = `agent-backup-${new Date().toISOString().replaceAll(/[:.]/g, "-")}-${randomUUID().slice(0, 8)}.zip`;
    const target = join(this.backupDirectory, filename);
    const temporary = `${target}.tmp`;
    try {
      await writeFile(temporary, bytes, { flag: "wx", mode: 0o600 });
      await rename(temporary, target);
      await chmod(target, 0o600).catch(() => {});
      await this.pruneArchives(filename);
    } catch (error) {
      await rm(temporary, { force: true });
      throw error;
    }
    return { filename, bytes, components: selected, componentStatus };
  }

  async importBackup(userId, bytes, components) {
    const selected = this.validateComponents(components);
    const parsed = this.parseBackupArchive(userId, bytes, selected);
    const safetyBackup = await this.createBackup(userId, selected);
    try {
      const restored = await this.applyBackup(userId, parsed, selected);
      return { restored, safetyBackup: safetyBackup.filename, restartRequired: selected.includes("whatsapp") };
    } catch (error) {
      try {
        const safetySnapshot = this.parseBackupArchive(userId, safetyBackup.bytes, selected);
        await this.applyBackup(userId, safetySnapshot, selected);
      } catch (rollbackError) {
        throw new Error(`Backup import failed and automatic rollback also failed. Safety archive: ${safetyBackup.filename}. Import: ${error.message}. Rollback: ${rollbackError.message}`);
      }
      throw new Error(`Backup import failed; previous data was restored. Safety archive: ${safetyBackup.filename}. ${error.message}`);
    }
  }

  validateComponents(components) {
    if (!Array.isArray(components) || !components.length
      || components.some((component) => typeof component !== "string" || !supportedComponents.has(component))) {
      throw new Error("Choose at least one supported backup component");
    }
    return [...new Set(components)].sort();
  }

  parseBackupArchive(userId, bytes, selected) {
    if (!Buffer.isBuffer(bytes) || bytes.length === 0 || bytes.length > maxArchiveBytes) {
      throw new Error("ZIP file is empty or exceeds the 128 MB compressed limit");
    }
    let archive;
    try {
      archive = new PizZip(bytes);
    } catch {
      throw new Error("Uploaded file is not a readable ZIP backup");
    }
    const entries = Object.entries(archive.files);
    if (entries.length > maxArchiveEntries) throw new Error("Backup contains too many entries");
    const files = new Map();
    let expandedBytes = 0;
    for (const [name, entry] of entries) {
      const normalized = name.replaceAll("\\", "/");
      const parts = normalized.replace(/\/$/, "").split("/");
      if (normalized.startsWith("/") || name.includes("\\")
        || parts.some((part) => !part || part === "." || part === ".." || /[\u0000-\u001f]/.test(part))) {
        throw new Error("Backup contains an unsafe ZIP path");
      }
      if (entry.dir) continue;
      const compressedBytes = Number(entry._data?.compressedSize ?? 0);
      const uncompressedBytes = Number(entry._data?.uncompressedSize ?? 0);
      expandedBytes += uncompressedBytes;
      if (!Number.isFinite(uncompressedBytes) || expandedBytes > maxArchiveBytes) {
        throw new Error("Backup expands beyond the 128 MB restore limit");
      }
      if (uncompressedBytes > 0 && (compressedBytes === 0 || uncompressedBytes / compressedBytes > 250)) {
        throw new Error("Backup compression ratio exceeds the safe limit");
      }
      const allowed = normalized === "BACKUP-INFO.txt"
        || normalized === "owner-data/ai-memory-and-conversations.json"
        || normalized === "owner-data/finance-ledger-and-audit.json"
        || normalized === "file-explorer/workspace-index.json"
        || normalized === "system/skills.json"
        || normalized === "system/runtime-settings.json"
        || normalized.startsWith("file-explorer/")
        || normalized.startsWith("sensitive/whatsapp-auth/");
      if (!allowed) throw new Error("Backup contains an unsupported archive entry");
      files.set(normalized, entry.asNodeBuffer());
    }
    const info = files.get("BACKUP-INFO.txt")?.toString("utf8");
    if (!info || !info.split(/\r?\n/).includes(`Owner WhatsApp number: ${userId}`)) {
      throw new Error("Backup belongs to a different owner or has no valid backup manifest");
    }
    const componentLine = info.match(/^Components: (.+)$/m)?.[1] ?? "";
    const archivedComponents = new Set(componentLine.split(", ").filter(Boolean));
    if (selected.some((component) => !archivedComponents.has(component))) {
      throw new Error("Selected restore components are not present in this backup");
    }
    for (const component of selected) {
      const requiredPath = {
        ai: "owner-data/ai-memory-and-conversations.json",
        finance: "owner-data/finance-ledger-and-audit.json",
        files: "file-explorer/workspace-index.json",
        system: "system/skills.json",
      }[component];
      if (requiredPath && !files.has(requiredPath)) throw new Error(`Backup is missing required ${component} data`);
    }
    return { files };
  }

  async applyBackup(userId, parsed, selected) {
    const { files } = parsed;
    const readJson = (path) => {
      try { return JSON.parse(files.get(path).toString("utf8")); }
      catch { throw new Error(`Backup JSON is invalid: ${path}`); }
    };
    const snapshot = {};
    if (selected.includes("ai")) snapshot.ai = readJson("owner-data/ai-memory-and-conversations.json");
    if (selected.includes("finance")) snapshot.finance = readJson("owner-data/finance-ledger-and-audit.json");
    const restored = {};
    if (selected.includes("ai") || selected.includes("finance")) {
      Object.assign(restored, this.financeStore.restoreBackupSnapshot(userId, selected, snapshot));
    }
    if (selected.includes("files")) {
      const index = readJson("file-explorer/workspace-index.json");
      const entries = [...files.entries()]
        .filter(([path]) => path.startsWith("file-explorer/") && path !== "file-explorer/workspace-index.json")
        .map(([path, fileBytes]) => ({ path: path.slice("file-explorer/".length), bytes: fileBytes }));
      restored.files = await this.fileWorkspace.restoreBackupSnapshot(userId, { entries, index });
    }
    if (selected.includes("system")) {
      const skills = readJson("system/skills.json");
      const installed = loadSystemSkills(this.config.skillsDirectory);
      if (!Array.isArray(skills) || skills.some((skill) => !installed.some((item) => item.id === skill.id) || typeof skill.enabled !== "boolean")) {
        throw new Error("Backup skills do not match the skills installed on this server");
      }
      for (const skill of skills) {
        const current = installed.find((item) => item.id === skill.id);
        if (current.enabled !== skill.enabled) setSystemSkillEnabled(skill.id, skill.enabled, this.config.skillsDirectory);
      }
      restored.skills = skills.length;
    }
    if (selected.includes("whatsapp")) {
      const prefix = "sensitive/whatsapp-auth/";
      const authEntries = [...files.entries()].filter(([path]) => path.startsWith(prefix));
      await this.restoreWhatsappAuth(authEntries.map(([path, fileBytes]) => ({ path: path.slice(prefix.length), bytes: fileBytes })));
      restored.whatsapp = authEntries.length;
    }
    return restored;
  }

  async restoreWhatsappAuth(entries) {
    const authDirectory = join(this.dataDirectory, "baileys-auth");
    const stagingDirectory = join(this.dataDirectory, `.baileys-auth-restore-${randomUUID()}`);
    const previousDirectory = join(this.dataDirectory, `.baileys-auth-previous-${randomUUID()}`);
    let movedPrevious = false;
    let movedStaging = false;
    await mkdir(stagingDirectory, { recursive: true, mode: 0o700 });
    try {
      for (const entry of entries) {
        const parts = entry.path.split("/");
        if (parts.some((part) => !part || part === "." || part === ".." || part.startsWith(".") || /[\\:\u0000-\u001f]/.test(part))) {
          throw new Error("WhatsApp backup contains an unsafe credential path");
        }
        const target = join(stagingDirectory, ...parts);
        await mkdir(dirname(target), { recursive: true, mode: 0o700 });
        await writeFile(target, entry.bytes, { flag: "wx", mode: 0o600 });
      }
      try {
        const current = await lstat(authDirectory);
        if (current.isSymbolicLink() || !current.isDirectory()) throw new Error("Existing WhatsApp credentials are not a regular directory");
        await rename(authDirectory, previousDirectory);
        movedPrevious = true;
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
      }
      await rename(stagingDirectory, authDirectory);
      movedStaging = true;
      await rm(previousDirectory, { recursive: true, force: true }).catch(() => {});
    } catch (error) {
      if (movedStaging) await rename(authDirectory, stagingDirectory).catch(() => {});
      if (movedPrevious) await rename(previousDirectory, authDirectory).catch(() => {});
      await rm(stagingDirectory, { recursive: true, force: true });
      throw error;
    }
  }

  async addDirectory(add, sourceDirectory, archivePrefix) {
    const root = resolve(sourceDirectory);
    const walk = async (directory) => {
      for (const entry of await readdir(directory, { withFileTypes: true })) {
        const sourcePath = join(directory, entry.name);
        if (entry.isSymbolicLink()) throw new Error("Backup refused a symbolic link");
        if (entry.isDirectory()) {
          await walk(sourcePath);
        } else if (entry.isFile()) {
          const archivePath = relative(root, sourcePath).split(sep).join("/");
          add(`${archivePrefix}/${archivePath}`, await readFile(sourcePath));
        } else {
          throw new Error("Backup found an unsupported filesystem entry");
        }
      }
    };
    await walk(root);
  }

  async pruneArchives(currentFilename) {
    const entries = await readdir(this.backupDirectory, { withFileTypes: true });
    const archives = await Promise.all(entries
      .filter((entry) => entry.isFile() && /^agent-backup-.*\.zip$/.test(entry.name))
      .map(async (entry) => ({ name: entry.name, modified: (await stat(join(this.backupDirectory, entry.name))).mtimeMs })));
    archives.sort((left, right) => right.modified - left.modified);
    const retained = new Set([currentFilename, ...archives.slice(0, retainedArchives - 1).map((entry) => entry.name)]);
    await Promise.all(archives.filter((entry) => !retained.has(entry.name))
      .map((entry) => rm(join(this.backupDirectory, entry.name), { force: true })));
  }
}

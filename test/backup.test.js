import assert from "node:assert/strict";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import PizZip from "pizzip";
import { BackupService } from "../src/backup.js";
import { FinanceStore } from "../src/finance/store.js";
import { FileWorkspace } from "../src/files/workspace.js";

test("ZIP backup contains selected owner data, workspace history, and optional WhatsApp credentials", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "agent-backup-test-"));
  const skillsDirectory = join(dataDirectory, "skills");
  mkdirSync(skillsDirectory);
  writeFileSync(join(skillsDirectory, "sample.json"), JSON.stringify({
    id: "sample-skill",
    name: "Sample",
    description: "Test skill",
    version: "1.0.0",
    enabled: true,
    instructions: ["Test instruction"],
  }));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const owner = "6281234567890";
  const otherOwner = "6289999999999";
  const backupService = new BackupService({
    dataDirectory,
    financeStore,
    fileWorkspace,
    config: { skillsDirectory, xkiroModel: "vendor/model", timeZone: "Asia/Jakarta" },
  });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  financeStore.recordTransaction(owner, {
    kind: "income", amount: 100, currency: "IDR", date: "2026-10-18", category: "Jasa", description: "Owner transaction",
  });
  financeStore.recordTransaction(otherOwner, {
    kind: "income", amount: 200, currency: "IDR", date: "2026-10-18", category: "Jasa", description: "Other owner secret",
  });
  financeStore.database.prepare(`
    INSERT INTO memory_entries (user_id, target, content, created_at, updated_at)
    VALUES (?, 'memory', 'Owner memory', '2026-10-18T00:00:00.000Z', '2026-10-18T00:00:00.000Z')
  `).run(owner);
  financeStore.database.prepare(`
    INSERT INTO conversation_messages (user_id, source_message_id, role, content, created_at)
    VALUES (?, 'message-1', 'user', 'Owner conversation', '2026-10-18T00:00:00.000Z')
  `).run(owner);
  const workspaceFile = await fileWorkspace.createTextFile(owner, "notes/private.txt", "Workspace note");
  const trashedFile = await fileWorkspace.createTextFile(owner, "old.txt", "Trash content");
  await fileWorkspace.softDeleteFile(owner, trashedFile.id);
  const authDirectory = join(dataDirectory, "baileys-auth");
  mkdirSync(authDirectory);
  writeFileSync(join(authDirectory, "creds.json"), "sensitive-session-data");

  const ordinaryBackup = await backupService.createBackup(owner, ["ai", "finance", "files", "system"]);
  const ordinaryZip = new PizZip(ordinaryBackup.bytes);
  const finance = JSON.parse(ordinaryZip.file("owner-data/finance-ledger-and-audit.json").asText());
  const ai = JSON.parse(ordinaryZip.file("owner-data/ai-memory-and-conversations.json").asText());
  assert.equal(finance.transactions.length, 1);
  assert.equal(finance.transactions[0].description, "Owner transaction");
  assert.equal(JSON.stringify(finance).includes("Other owner secret"), false);
  assert.equal(ai.memory_entries[0].content, "Owner memory");
  assert.equal(ai.conversation_messages[0].content, "Owner conversation");
  assert.equal(ordinaryZip.file(`file-explorer/${workspaceFile.path}`).asText(), "Workspace note");
  assert.ok(Object.keys(ordinaryZip.files).some((path) => path.startsWith("file-explorer/.trash/")));
  assert.equal(ordinaryZip.file("sensitive/whatsapp-auth/creds.json"), null);
  assert.ok(ordinaryZip.file("system/skills.json"));
  assert.ok(ordinaryZip.file("BACKUP-INFO.txt"));
  assert.equal(readFileSync(join(dataDirectory, "backups", ordinaryBackup.filename)).equals(ordinaryBackup.bytes), true);

  const sensitiveBackup = await backupService.createBackup(owner, ["whatsapp"]);
  const sensitiveZip = new PizZip(sensitiveBackup.bytes);
  assert.equal(sensitiveZip.file("sensitive/whatsapp-auth/creds.json").asText(), "sensitive-session-data");
  assert.equal(sensitiveBackup.componentStatus.whatsapp.includes("credentials"), true);
});

test("ZIP import restores selected components and creates a pre-import safety backup", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "agent-backup-restore-test-"));
  const skillsDirectory = join(dataDirectory, "skills");
  mkdirSync(skillsDirectory);
  writeFileSync(join(skillsDirectory, "sample.json"), JSON.stringify({
    id: "sample-skill",
    name: "Sample",
    description: "Test skill",
    version: "1.0.0",
    enabled: true,
    instructions: ["Test instruction"],
  }));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const owner = "6281234567890";
  const backupService = new BackupService({
    dataDirectory,
    financeStore,
    fileWorkspace,
    config: { skillsDirectory, xkiroModel: "vendor/model", timeZone: "Asia/Jakarta" },
  });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const original = financeStore.recordTransaction(owner, {
    kind: "income", amount: 100, currency: "IDR", date: "2026-10-18", category: "Jasa", description: "Before backup",
  });
  financeStore.database.prepare(`
    INSERT INTO memory_entries (user_id, target, content, created_at, updated_at)
    VALUES (?, 'memory', 'Saved memory', '2026-10-18T00:00:00.000Z', '2026-10-18T00:00:00.000Z')
  `).run(owner);
  const file = await fileWorkspace.createTextFile(owner, "notes.txt", "before backup");
  const archive = await backupService.createBackup(owner, ["ai", "finance", "files", "system"]);

  financeStore.recordTransaction(owner, {
    kind: "expense", amount: 25, currency: "IDR", date: "2026-10-18", category: "Makan", description: "After backup",
  });
  financeStore.database.prepare("UPDATE memory_entries SET content = 'Changed memory' WHERE user_id = ?").run(owner);
  await fileWorkspace.editTextFile(owner, file.id, "changed after backup");

  const result = await backupService.importBackup(owner, archive.bytes, ["ai", "finance", "files", "system"]);
  assert.equal(result.restored.finance, 1);
  assert.equal(result.restored.files.files, 1);
  assert.equal(financeStore.listRecentTransactions(owner, 10)[0].id, original.id);
  assert.equal(financeStore.database.prepare("SELECT content FROM memory_entries WHERE user_id = ?").get(owner).content, "Saved memory");
  assert.equal(await fileWorkspace.readTextFile(owner, file.id), "before backup");
  assert.equal(readFileSync(join(dataDirectory, "backups", result.safetyBackup)).length > 0, true);
});

test("ZIP import replaces WhatsApp auth directory safely", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "agent-backup-whatsapp-restore-test-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const owner = "6281234567890";
  const backupService = new BackupService({ dataDirectory, financeStore, fileWorkspace });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  const authDirectory = join(dataDirectory, "baileys-auth");
  mkdirSync(authDirectory);
  writeFileSync(join(authDirectory, "old.json"), "old-session");
  const archive = await backupService.createBackup(owner, ["whatsapp"]);
  writeFileSync(join(authDirectory, "changed.json"), "must-disappear");
  const result = await backupService.importBackup(owner, archive.bytes, ["whatsapp"]);
  assert.equal(result.restored.whatsapp, 1);
  assert.equal(readFileSync(join(authDirectory, "old.json"), "utf8"), "old-session");
  assert.equal(existsSync(join(authDirectory, "changed.json")), false);
});

test("ZIP import rejects an archive belonging to another WhatsApp owner", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "agent-backup-owner-test-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const backupService = new BackupService({ dataDirectory, financeStore, fileWorkspace });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const archive = new PizZip();
  archive.file("BACKUP-INFO.txt", "Owner WhatsApp number: 6289999999999\nComponents: files\n");
  archive.file("file-explorer/workspace-index.json", JSON.stringify({ files: [], delete_requests: [], file_events: [] }));
  await assert.rejects(
    backupService.importBackup("6281234567890", archive.generate({ type: "nodebuffer" }), ["files"]),
    /different owner/,
  );
});
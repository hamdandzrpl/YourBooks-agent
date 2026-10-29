import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { FileWorkspace } from "../src/files/workspace.js";

function withWorkspace(t, options = {}) {
  const directory = mkdtempSync(join(tmpdir(), "whatsapp-file-test-"));
  const workspace = new FileWorkspace({ dataDirectory: directory, ...options });
  t.after(() => {
    workspace.close();
    rmSync(directory, { recursive: true, force: true });
  });
  return { workspace, directory };
}

test("uploaded attachments persist under safe names and deduplicate by message", async (t) => {
  const { workspace } = withWorkspace(t);
  const input = {
    filename: "../../cashbook.xlsx",
    bytes: Buffer.from("template-bytes"),
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    sourceMessageId: "wamid.upload.1",
  };
  const saved = await workspace.saveAttachment("6281234567890", input);
  const duplicate = await workspace.saveAttachment("6281234567890", input);

  assert.equal(saved.id, duplicate.id);
  assert.equal(saved.name, "cashbook.xlsx");
  assert.equal(saved.path, "inbox/cashbook.xlsx");
  assert.deepEqual(await workspace.listFiles("6281234567890"), [saved]);
  assert.deepEqual(await workspace.listFiles("6289999999999"), []);
});

test("workspace creates, searches, reads, edits, renames, trashes, and restores files", async (t) => {
  const { workspace } = withWorkspace(t);
  await workspace.createFolder("6281234567890", "reports/2026");
  assert.deepEqual(await workspace.listDirectory("6281234567890"), [
    { type: "folder", name: "reports", path: "reports" },
  ]);
  assert.deepEqual(await workspace.listDirectory("6281234567890", "reports"), [
    { type: "folder", name: "2026", path: "reports/2026" },
  ]);
  const file = await workspace.createTextFile("6281234567890", "reports/2026/notes.md", "pemasukan awal");
  assert.equal((await workspace.listDirectory("6281234567890", "reports/2026"))[0].type, "file");
  assert.equal(await workspace.readTextFile("6281234567890", file.id), "pemasukan awal");
  assert.equal((await workspace.searchFiles("6281234567890", "pemasukan")).length, 1);

  const edited = await workspace.editTextFile("6281234567890", file.id, "pemasukan direvisi");
  assert.equal(edited.sizeBytes, Buffer.byteLength("pemasukan direvisi"));
  const renamed = await workspace.renameFile("6281234567890", file.id, "reports/2026/kas.md");
  assert.equal(renamed.path, "reports/2026/kas.md");

  const deletion = workspace.requestDelete("6281234567890", file.id);
  assert.equal(workspace.getFile("6281234567890", file.id).id, file.id);
  await workspace.confirmDelete("6281234567890", deletion.id);
  assert.equal(workspace.getFile("6281234567890", file.id), null);
  const restored = await workspace.restoreFile("6281234567890", file.id);
  assert.equal(restored.path, "reports/2026/kas.md");
  assert.equal(await workspace.readTextFile("6281234567890", file.id), "pemasukan direvisi");
});

test("file copy keeps the source intact and stays in the owner's workspace", async (t) => {
  const { workspace } = withWorkspace(t);
  const original = await workspace.createTextFile("6281234567890", "reports/source.txt", "original text");
  const copy = await workspace.copyFile("6281234567890", original.id, "reports/copy.txt");

  assert.notEqual(copy.id, original.id);
  assert.equal(copy.path, "reports/copy.txt");
  assert.equal(await workspace.readTextFile("6281234567890", original.id), "original text");
  assert.equal(await workspace.readTextFile("6281234567890", copy.id), "original text");
  await assert.rejects(workspace.copyFile("6289999999999", original.id, "copy.txt"), /not found/);
});

test("workspace refuses traversal paths, unsupported text file types, and oversized uploads", async (t) => {
  const { workspace } = withWorkspace(t, { maxFileBytes: 8 });
  await assert.rejects(
    workspace.createTextFile("6281234567890", "../outside.txt", "bad"),
    /relative|Hidden|parent/,
  );
  await assert.rejects(
    workspace.createTextFile("6281234567890", "script.js", "console.log(1)"),
    /only \.txt/,
  );
  await assert.rejects(workspace.saveAttachment("6281234567890", {
    filename: "too-large.pdf",
    bytes: Buffer.from("more than eight bytes"),
    sourceMessageId: "wamid.large",
  }), /size limit/);
});

test("delete requests are bound to the owning WhatsApp user", async (t) => {
  const { workspace } = withWorkspace(t);
  const file = await workspace.createTextFile("6281234567890", "notes.txt", "private");
  const request = workspace.requestDelete("6281234567890", file.id);
  await assert.rejects(workspace.confirmDelete("6289999999999", request.id), /not found/);
  assert.ok(workspace.getFile("6281234567890", file.id));
});

test("function-call deletion moves a file to Trash and can be restored without a keyword", async (t) => {
  const { workspace } = withWorkspace(t);
  const file = await workspace.createTextFile("6281234567890", "notes.md", "private notes");
  const deleted = await workspace.softDeleteFile("6281234567890", file.id);
  assert.equal(deleted.id, file.id);
  assert.equal(workspace.getFile("6281234567890", file.id), null);
  const restored = await workspace.restoreFile("6281234567890", file.id);
  assert.equal(restored.id, file.id);
  assert.equal(await workspace.readTextFile("6281234567890", file.id), "private notes");
});

test("restore by filename finds the matching trashed file and rejects when missing", async (t) => {
  const { workspace } = withWorkspace(t);
  const owner = "6281234567890";
  const file = await workspace.createTextFile(owner, "Catatan/backup.txt", "backup content");
  await workspace.softDeleteFile(owner, file.id);

  const restored = await workspace.restoreFileByName(owner, "backup.txt");
  assert.equal(restored.id, file.id);
  assert.equal(await workspace.readTextFile(owner, restored.id), "backup content");

  await assert.rejects(workspace.restoreFileByName(owner, "missing.txt"), /not found/);
  await assert.rejects(workspace.restoreFileByName(owner, ""), /required/);
});

test("delete by filename works when unique and fails when missing or ambiguous", async (t) => {
  const { workspace } = withWorkspace(t);
  const owner = "6281234567890";
  await workspace.createTextFile(owner, "Catatan/single.txt", "one");
  await workspace.createTextFile(owner, "Catatan/dup.txt", "two");
  await workspace.createTextFile(owner, "Lain/dup.txt", "three");

  const deleted = await workspace.softDeleteFileByName(owner, "single.txt");
  assert.equal(deleted.path, "Catatan/single.txt");
  assert.equal(workspace.getFile(owner, deleted.id), null);

  await assert.rejects(workspace.softDeleteFileByName(owner, "dup.txt"), /Multiple files/);
  await assert.rejects(workspace.softDeleteFileByName(owner, "missing.txt"), /not found/);
  await assert.rejects(workspace.softDeleteFileByName(owner, ""), /required/);
});

test("generated binary files use the same workspace confinement and metadata index", async (t) => {
  const { workspace } = withWorkspace(t);
  const file = await workspace.saveGeneratedFile(
    "6281234567890",
    "exports/report.xlsx",
    Buffer.from("binary-office-file"),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  assert.equal(file.path, "exports/report.xlsx");
  assert.equal(file.sizeBytes, Buffer.byteLength("binary-office-file"));
  assert.equal((await workspace.listFiles("6281234567890", "exports")).length, 1);
  await assert.rejects(
    workspace.saveGeneratedFile("6281234567890", "../escape.docx", Buffer.from("bad"), "application/octet-stream"),
    /relative|Hidden|parent/,
  );
});

test("workspace can replace a managed workbook in place and preserves its previous revision", async (t) => {
  const { workspace } = withWorkspace(t);
  const owner = "6281234567890";
  const original = await workspace.saveGeneratedFile(owner, "books/main.xlsx", Buffer.from("old workbook"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const updated = await workspace.replaceGeneratedFile(owner, original.id, Buffer.from("new workbook"));

  assert.equal(updated.id, original.id);
  assert.equal(updated.path, original.path);
  assert.equal(updated.sizeBytes, Buffer.byteLength("new workbook"));
  assert.equal((await workspace.getFilePath(owner, original.id)).endsWith("main.xlsx"), true);
  const revisions = workspace.database.prepare("SELECT COUNT(*) AS count FROM file_events WHERE action = 'file_replaced'").get();
  assert.equal(revisions.count, 1);
  await assert.rejects(workspace.replaceGeneratedFile("6289999999999", original.id, Buffer.from("bad")), /not found/);
});

test("workspace backup snapshots include owner files, Trash, revisions, and private index only", async (t) => {
  const { workspace } = withWorkspace(t);
  const owner = "6281234567890";
  const otherOwner = "6289999999999";
  const current = await workspace.saveGeneratedFile(owner, "books/main.xlsx", Buffer.from("old version"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  await workspace.replaceGeneratedFile(owner, current.id, Buffer.from("current version"));
  const deleted = await workspace.createTextFile(owner, "old.txt", "trash content");
  await workspace.softDeleteFile(owner, deleted.id);
  await workspace.createTextFile(otherOwner, "private.txt", "other owner's data");

  const snapshot = await workspace.exportBackupSnapshot(owner);
  const archivedPaths = snapshot.entries.map((entry) => entry.path);
  assert.ok(archivedPaths.includes("books/main.xlsx"));
  assert.ok(archivedPaths.some((path) => path.startsWith(`.versions/${current.id}/`)));
  assert.ok(archivedPaths.some((path) => path.startsWith(".trash/")));
  assert.equal(snapshot.entries.some((entry) => entry.bytes.toString("utf8").includes("other owner's data")), false);
  assert.equal(snapshot.index.files.length, 2);
  assert.equal(snapshot.index.file_events.every((event) => event.user_key !== workspace.database.prepare(
    "SELECT user_key FROM files WHERE relative_path = 'private.txt'",
  ).get()?.user_key), true);
});

test("workspace backup restore reinstates a selected owner's files and leaves other owners intact", async (t) => {
  const { workspace } = withWorkspace(t);
  const owner = "6281234567890";
  const otherOwner = "6289999999999";
  const original = await workspace.createTextFile(owner, "notes/original.txt", "before backup");
  const snapshot = await workspace.exportBackupSnapshot(owner);
  await workspace.editTextFile(owner, original.id, "changed after backup");
  await workspace.createTextFile(owner, "notes/later.txt", "remove on restore");
  const other = await workspace.createTextFile(otherOwner, "private.txt", "keep other owner");

  const restored = await workspace.restoreBackupSnapshot(owner, snapshot);
  assert.equal(restored.files, 1);
  assert.equal(await workspace.readTextFile(owner, original.id), "before backup");
  assert.equal(workspace.getFile(owner, other.id), null);
  assert.equal(await workspace.readTextFile(otherOwner, other.id), "keep other owner");
  assert.deepEqual((await workspace.listFiles(owner)).map((file) => file.path), ["notes/original.txt"]);
});

test("workspace saves generated file sets atomically and cleans staging files on collision", async (t) => {
  const { workspace } = withWorkspace(t);
  const existing = await workspace.createTextFile("6281234567890", "unpacked/existing.txt", "keep");

  await assert.rejects(workspace.saveGeneratedFiles("6281234567890", [
    { relativePath: "unpacked/new.txt", bytes: Buffer.from("new"), mimeType: "text/plain" },
    { relativePath: existing.path, bytes: Buffer.from("overwrite"), mimeType: "text/plain" },
  ]), /already exists/);

  assert.equal(workspace.getFile("6281234567890", existing.id).id, existing.id);
  assert.equal(await workspace.readTextFile("6281234567890", existing.id), "keep");
  assert.deepEqual((await workspace.listFiles("6281234567890")).map((file) => file.path), [existing.path]);

  const created = await workspace.saveGeneratedFiles("6281234567890", [
    { relativePath: "unpacked/one.txt", bytes: Buffer.from("one"), mimeType: "text/plain" },
    { relativePath: "unpacked/empty.bin", bytes: Buffer.alloc(0), mimeType: "application/octet-stream" },
  ]);
  assert.equal(created.length, 2);
  assert.equal(created[1].sizeBytes, 0);
});
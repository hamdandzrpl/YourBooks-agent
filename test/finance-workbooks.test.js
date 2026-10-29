import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import { FinanceStore } from "../src/finance/store.js";
import { FinanceWorkbookSync } from "../src/finance/workbooks.js";
import { OfficeFiles } from "../src/files/office.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { ExcelMcp } from "../src/mcp/excel.js";

test("finance books create registered workbooks and sync ledger rows in place", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-workbooks-"));
  const owner = "6281234567890";
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const officeFiles = new OfficeFiles(fileWorkspace);
  const excelMcp = new ExcelMcp({ financeStore });
  const workbookSync = new FinanceWorkbookSync({ financeStore, fileWorkspace, excelMcp });
  t.after(async () => {
    await excelMcp.close();
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const book = financeStore.getActiveBook(owner);
  const initialFile = await workbookSync.ensureBookWorkbook(owner, book.id);
  // The managed workbook is named after the book, not its internal id, so the
  // user and the agent never see a file called `B-88DD86E4.xlsx`.
  assert.equal(initialFile.path, "books/buku-utama.xlsx");
  assert.equal(initialFile.name, "buku-utama.xlsx");
  assert.equal(financeStore.getBook(owner, book.id).workbook_file_id, initialFile.id);

  const initialWorkbook = new ExcelJS.Workbook();
  await initialWorkbook.xlsx.load(readFileSync(await fileWorkspace.getFilePath(owner, initialFile.id)));
  initialWorkbook.addWorksheet("Catatan").getCell("A1").value = "Data milik pengguna";
  await fileWorkspace.replaceGeneratedFile(owner, initialFile.id, Buffer.from(await initialWorkbook.xlsx.writeBuffer()));

  const excelTools = await excelMcp.connect({ fileWorkspace, userId: owner });
  await assert.rejects(() => excelTools.call("mcp_excel_excel_write_to_sheet", {
    file_id: initialFile.id,
    sheetName: "transaksi",
    newSheet: false,
    range: "A1:A1",
    values: [["tamper"]],
  }), /managed by the ledger/);

  const transaction = financeStore.recordTransaction(owner, {
    kind: "income",
    amount: 100000,
    currency: "IDR",
    date: "2026-10-18",
    category: "Penjualan",
    description: "Invoice A",
  }, { bookId: book.id });
  const syncedFile = await workbookSync.syncBook(owner, book.id);
  assert.equal(syncedFile.id, initialFile.id);
  assert.notEqual(syncedFile.sha256, initialFile.sha256);

  const preview = await officeFiles.inspectSpreadsheet(owner, syncedFile.id);
  const sheetNames = preview.worksheets.map((sheet) => sheet.name);
  for (const name of ["Ringkasan", "Arus Kas", "Kategori", "Transaksi", "Anggaran", "Anomali", "Petunjuk"]) {
    assert.ok(sheetNames.includes(name), `missing sheet ${name}`);
  }
  const ledgerPreview = preview.worksheets.find((sheet) => sheet.name === "Transaksi");
  assert.equal(ledgerPreview.preview[1][1], transaction.id);
  assert.equal(ledgerPreview.preview[1][5], "Invoice A");
  assert.equal(ledgerPreview.preview[1][6], 100000);
  const syncedWorkbook = new ExcelJS.Workbook();
  await syncedWorkbook.xlsx.load(readFileSync(await fileWorkspace.getFilePath(owner, syncedFile.id)));
  assert.equal(syncedWorkbook.getWorksheet("Catatan").getCell("A1").value, "Data milik pengguna");
  assert.equal((await fileWorkspace.listFiles(owner, "books")).length, 1);
  const replacementsBeforeRefresh = fileWorkspace.database.prepare(
    "SELECT COUNT(*) AS count FROM file_events WHERE action = 'file_replaced'",
  ).get().count;
  const restartedSync = new FinanceWorkbookSync({ financeStore, fileWorkspace });
  await restartedSync.syncBook(owner, book.id);
  assert.equal(fileWorkspace.database.prepare(
    "SELECT COUNT(*) AS count FROM file_events WHERE action = 'file_replaced'",
  ).get().count, replacementsBeforeRefresh);

  const createdBook = await workbookSync.createBook(owner, "Bisnis");
  assert.ok(createdBook.workbook_file_id);
  assert.equal(fileWorkspace.getFile(owner, createdBook.workbook_file_id).name, "bisnis.xlsx");
});

test("an existing book-id workbook is renamed to a readable name on the next sync", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-workbooks-legacy-"));
  const owner = "6281234567890";
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const workbookSync = new FinanceWorkbookSync({ financeStore, fileWorkspace });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const book = financeStore.getActiveBook(owner);
  // Simulate a workspace written by the old scheme: `books/<book id>.xlsx`.
  const legacy = await fileWorkspace.saveGeneratedFile(
    owner,
    `books/${book.id}.xlsx`,
    Buffer.from(await workbookSync.renderBytes(owner, book, [])),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  financeStore.updateBookWorkbook(owner, book.id, legacy.id);

  const migrated = await workbookSync.ensureBookWorkbook(owner, book.id);
  assert.equal(migrated.id, legacy.id, "the same file must be renamed, not duplicated");
  assert.equal(migrated.path, "books/buku-utama.xlsx");
  assert.equal((await fileWorkspace.listFiles(owner, "books")).length, 1);

  // A name the user picked themselves must survive the next sync untouched.
  await fileWorkspace.renameFile(owner, legacy.id, "books/laporan-pribadi.xlsx");
  const kept = await workbookSync.ensureBookWorkbook(owner, book.id);
  assert.equal(kept.path, "books/laporan-pribadi.xlsx");
});

test("two books whose names collapse to the same slug get distinct files", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-workbooks-slug-"));
  const owner = "6281234567890";
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const workbookSync = new FinanceWorkbookSync({ financeStore, fileWorkspace });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const first = await workbookSync.createBook(owner, "Toko A");
  const second = await workbookSync.createBook(owner, "Toko-A");
  const firstName = fileWorkspace.getFile(owner, first.workbook_file_id).name;
  const secondName = fileWorkspace.getFile(owner, second.workbook_file_id).name;
  assert.notEqual(firstName, secondName);
  assert.equal(firstName, "toko-a.xlsx");
  assert.equal(secondName, "toko-a-2.xlsx");
});

test("asking for the same report twice saves instead of failing", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-workbooks-repeat-"));
  const owner = "6281234567890";
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const workbookSync = new FinanceWorkbookSync({ financeStore, fileWorkspace });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  financeStore.recordTransaction(owner, {
    kind: "income", amount: 1_000_000, category: "Penjualan", description: "jual", date: "2026-10-21",
  });
  // The second save used to throw "A file already exists at the requested path",
  // which the user saw as "the agent cannot save to the workspace".
  const first = await workbookSync.renderReportFile(owner, {});
  const second = await workbookSync.renderReportFile(owner, {});
  const third = await workbookSync.renderReportFile(owner, { filename: "laporan-oktober.xlsx" });
  const fourth = await workbookSync.renderReportFile(owner, { filename: "laporan-oktober.xlsx" });
  assert.equal(first.file.path, "exports/laporan-buku-utama-semua-periode.xlsx");
  assert.equal(second.file.path, "exports/laporan-buku-utama-semua-periode-2.xlsx");
  assert.equal(third.file.path, "exports/laporan-oktober.xlsx");
  assert.equal(fourth.file.path, "exports/laporan-oktober-2.xlsx");
  assert.equal((await fileWorkspace.listFiles(owner, "exports")).length, 4);
});
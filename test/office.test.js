import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import PizZip from "pizzip";
import ExcelJS from "exceljs";
import { FinanceStore } from "../src/finance/store.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { OfficeFiles } from "../src/files/office.js";
import { fileToolNames, runFileTool } from "../src/files/tools.js";

function withOfficeFiles(t) {
  const directory = mkdtempSync(join(tmpdir(), "whatsapp-office-test-"));
  const workspace = new FileWorkspace({ dataDirectory: directory });
  const office = new OfficeFiles(workspace);
  t.after(() => {
    workspace.close();
    rmSync(directory, { recursive: true, force: true });
  });
  return { workspace, office, directory };
}

test("XLSX templates can be inspected and appended to a new copy", async (t) => {
  const { workspace, office } = withOfficeFiles(t);
  const template = await office.createSpreadsheet("6281234567890", {
    filename: "template.xlsx",
    sheetName: "Kas Harian",
    columns: ["Tanggal", "Kategori", "Jumlah"],
    rows: [["2026-10-14", "Makan", 25000]],
  });
  const inspection = await office.inspectSpreadsheet("6281234567890", template.id);
  assert.equal(inspection.worksheets[0].name, "Kas Harian");
  assert.equal(inspection.worksheets[0].preview[0][1], "Tanggal");
  assert.equal(inspection.worksheets[0].rowCount, 2);

  const result = await office.appendSpreadsheetRows("6281234567890", {
    fileId: template.id,
    sheetName: "Kas Harian",
    rows: [["2026-10-15", "Transport", 10000]],
    saveMode: "copy",
  });
  assert.equal(result.name, "template-updated.xlsx");
  assert.equal((await office.inspectSpreadsheet("6281234567890", result.id)).worksheets[0].rowCount, 3);
  assert.equal((await office.inspectSpreadsheet("6281234567890", template.id)).worksheets[0].rowCount, 2);

  const edited = await office.updateSpreadsheetCells("6281234567890", {
    fileId: template.id,
    sheetName: "Kas Harian",
    updates: [{ cell: "B2", value: "Belanja" }],
    outputFilename: "template-corrected.xlsx",
    saveMode: "copy",
  });
  assert.equal((await office.inspectSpreadsheet("6281234567890", edited.id)).worksheets[0].preview[1][2], "Belanja");
  assert.equal((await office.inspectSpreadsheet("6281234567890", template.id)).worksheets[0].preview[1][2], "Makan");

  const revisionCount = workspace.database.prepare(
    "SELECT COUNT(*) AS count FROM file_events WHERE action = 'file_replaced' AND file_id = ?",
  ).get(result.id).count;
  const updatedInPlace = await office.appendSpreadsheetRows("6281234567890", {
    fileId: result.id,
    sheetName: "Kas Harian",
    rows: [["2026-10-16", "Makan", 12000]],
    saveMode: "in_place",
  });
  assert.equal(updatedInPlace.id, result.id);
  assert.equal(updatedInPlace.path, result.path);
  assert.equal((await office.inspectSpreadsheet("6281234567890", result.id)).worksheets[0].rowCount, 4);
  assert.equal(workspace.database.prepare(
    "SELECT COUNT(*) AS count FROM file_events WHERE action = 'file_replaced' AND file_id = ?",
  ).get(result.id).count, revisionCount + 1);
});

test("finance workbook templates include styled summaries, formulas, and scoped ledger data", async (t) => {
  const directory = mkdtempSync(join(tmpdir(), "whatsapp-finance-workbook-test-"));
  const workspace = new FileWorkspace({ dataDirectory: directory });
  const owner = "6281234567890";
  const financeStore = new FinanceStore(join(directory, "finance.sqlite"));
  const office = new OfficeFiles(workspace, financeStore);
  t.after(() => {
    financeStore.close();
    workspace.close();
    rmSync(directory, { recursive: true, force: true });
  });
  const book = financeStore.getActiveBook(owner);
  financeStore.recordTransaction(owner, {
    kind: "income", amount: 1250000, currency: "IDR", date: "2026-10-03", category: "Penjualan", description: "Invoice 01",
  }, { bookId: book.id });
  financeStore.recordTransaction(owner, {
    kind: "expense", amount: 175000, currency: "IDR", date: "2026-10-08", category: "Operasional", description: "Perlengkapan",
  }, { bookId: book.id });
  financeStore.recordTransaction(owner, {
    kind: "income", amount: 100, currency: "USD", date: "2026-10-10", category: "Service", description: "USD invoice",
  }, { bookId: book.id });

  await assert.rejects(office.createFinanceWorkbook(owner, { period: "2026-10" }), /multiple currencies/);

  const report = await runFileTool("spreadsheet", {
    action: "create_finance_workbook",
    template: "cashbook",
    period: "2026-10",
    currency: "IDR",
    filename: "laporan-oktober.xlsx",
  }, { fileWorkspace: workspace, officeFiles: office, userId: owner });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await readFile(await workspace.getFilePath(owner, report.id)));
  const summary = workbook.getWorksheet("Ringkasan");
  const transactions = workbook.getWorksheet("Transaksi");
  assert.equal(report.transaction_count, 2);
  assert.equal(summary.getCell("A5").value.result, 1250000);
  assert.match(summary.getCell("A5").value.formula, /SUMIFS/);
  assert.equal(summary.getCell("C5").value.result, 175000);
  assert.equal(summary.getCell("E5").value.result, 1075000);
  assert.equal(transactions.getCell("A2").value, financeStore.listBookTransactions(owner, book.id)[0].id);
  assert.equal(transactions.getCell("F3").value, 175000);
  assert.equal(summary.getCell("A1").fill.fgColor.argb, "FF1D2A32");
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Ringkasan", "Transaksi", "Petunjuk"]);

  const budgetFile = await runFileTool("spreadsheet", {
    action: "create_finance_workbook", template: "budget", period: "2026-10", currency: "IDR",
  }, { fileWorkspace: workspace, officeFiles: office, userId: owner });
  const budgetWorkbook = new ExcelJS.Workbook();
  await budgetWorkbook.xlsx.load(await readFile(await workspace.getFilePath(owner, budgetFile.id)));
  const budgetSheet = budgetWorkbook.getWorksheet("Anggaran");
  assert.equal(budgetSheet.getCell("A2").value, "Operasional");
  assert.equal(budgetSheet.getCell("B2").value, null);
  assert.equal(budgetSheet.getCell("C2").value.result, 175000);
  assert.match(budgetSheet.getCell("C2").value.formula, /SUMIFS/);

  assert.equal((await office.createFinanceWorkbook(owner, { currency: "USD" })).transaction_count, 1);
});

test("DOCX documents can be created, read, and filled without overwriting the template", async (t) => {
  const { office } = withOfficeFiles(t);
  const template = await office.createWordDocument("6281234567890", {
    filename: "receipt-template.docx",
    title: "Kwitansi",
    paragraphs: ["Nama: {nama}", "Jumlah: {jumlah}"],
  });
  const read = await office.readWordDocument("6281234567890", template.id);
  assert.match(read.text, /Nama: \{nama\}/);

  const result = await office.fillWordTemplate("6281234567890", {
    fileId: template.id,
    data: { nama: "Sari", jumlah: "Rp25.000" },
  });
  assert.equal(result.name, "receipt-template-filled.docx");
  const filled = await office.readWordDocument("6281234567890", result.id);
  assert.match(filled.text, /Nama: Sari/);
  assert.match(filled.text, /Jumlah: Rp25\.000/);
  assert.match((await office.readWordDocument("6281234567890", template.id)).text, /\{nama\}/);
});

test("Office operations reject files from another user's workspace", async (t) => {
  const { office } = withOfficeFiles(t);
  const workbook = await office.createSpreadsheet("6281234567890", {
    filename: "private.xlsx",
    columns: ["Amount"],
  });
  await assert.rejects(office.inspectSpreadsheet("6289999999999", workbook.id), /not found/);
});

test("workspace tools create and safely extract ZIP archives for the owning user", async (t) => {
  const { workspace, office } = withOfficeFiles(t);
  const first = await workspace.createTextFile("6281234567890", "reports/summary.txt", "summary data");
  const second = await workspace.saveGeneratedFile(
    "6281234567890",
    "reports/data.bin",
    Buffer.from([0, 1, 2, 3]),
    "application/octet-stream",
  );
  const context = { userId: "6281234567890", fileWorkspace: workspace, officeFiles: office };
  const archive = await runFileTool("workspace_files", {
    action: "create_zip",
    file_ids: [first.id, second.id],
    filename: "reports.zip",
  }, context);
  assert.equal(archive.name, "reports.zip");

  const extracted = await runFileTool("workspace_files", {
    action: "extract_zip",
    file_id: archive.id,
    destination_folder: "exports/recovered",
  }, context);
  assert.deepEqual(extracted.map((file) => file.path).sort(), [
    "exports/recovered/reports/data.bin",
    "exports/recovered/reports/summary.txt",
  ]);
  assert.equal(await workspace.readTextFile("6281234567890", extracted.find((file) => file.name === "summary.txt").id), "summary data");
  const extractedBinary = await readFile(await workspace.getFilePath(
    "6281234567890",
    extracted.find((file) => file.name === "data.bin").id,
  ));
  assert.deepEqual(extractedBinary, Buffer.from([0, 1, 2, 3]));
  await assert.rejects(office.createZipArchive("6289999999999", { fileIds: [first.id] }), /not found/);
  await assert.rejects(office.extractZipArchive("6289999999999", { fileId: archive.id }), /not found/);
});

test("ZIP extraction rejects path traversal and refuses to overwrite existing files", async (t) => {
  const { workspace, office } = withOfficeFiles(t);
  const unsafeZip = new PizZip();
  unsafeZip.file("../escape.txt", "outside");
  const unsafe = await workspace.saveGeneratedFile(
    "6281234567890",
    "inbox/unsafe.zip",
    Buffer.from(unsafeZip.generate({ type: "nodebuffer", compression: "DEFLATE" })),
    "application/zip",
  );
  await assert.rejects(office.extractZipArchive("6281234567890", { fileId: unsafe.id }), /unsafe path/);

  const validZip = new PizZip();
  validZip.file("report.txt", "from archive");
  const valid = await workspace.saveGeneratedFile(
    "6281234567890",
    "inbox/valid.zip",
    Buffer.from(validZip.generate({ type: "nodebuffer", compression: "DEFLATE" })),
    "application/zip",
  );
  await workspace.createTextFile("6281234567890", "exports/valid/report.txt", "keep this");
  await assert.rejects(office.extractZipArchive("6281234567890", {
    fileId: valid.id,
    destinationFolder: "exports/valid",
  }), /overwrite existing/);
  assert.equal((await workspace.listFiles("6281234567890", "exports/valid")).length, 1);
});

test("agent file tools create, list, and send a user's Office file", async (t) => {
  const { workspace, office } = withOfficeFiles(t);
  const context = { userId: "6281234567890", fileWorkspace: workspace, officeFiles: office };
  assert.equal(fileToolNames.has("spreadsheet"), true);
  const workbook = await office.createSpreadsheet("6281234567890", {
    filename: "from-agent.xlsx",
    sheetName: "Kas",
    columns: ["Tanggal", "Jumlah"],
    rows: [["2026-10-14", 20000]],
  });
  const listing = await runFileTool("workspace_files", { action: "list" }, context);
  assert.equal(listing.some((file) => file.id === workbook.id), true);

  let sent;
  const result = await runFileTool("workspace_files", { action: "send", file_id: workbook.id }, {
    ...context,
    sendFile: async (file) => { sent = file; },
  });
  assert.equal(result.sent, true);
  assert.equal(sent.filename, "from-agent.xlsx");
  assert.ok(sent.bytes.length > 0);
});

test("Office reader rejects high-ratio ZIP entries before expanding them", async (t) => {
  const { workspace, office } = withOfficeFiles(t);
  const zip = new PizZip();
  zip.file("word/document.xml", "x".repeat(20_000));
  const compressed = zip.generate({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  const file = await workspace.saveGeneratedFile(
    "6281234567890",
    "inbox/compressed.docx",
    compressed,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  );
  await assert.rejects(office.readWordDocument("6281234567890", file.id), /compression ratio/);
});
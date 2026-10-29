import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import { analyzeBook, primaryCurrency } from "../src/finance/analytics.js";
import { renderReport } from "../src/finance/report.js";
import { readZip } from "../src/finance/xlsx/zip.js";
import { FinanceStore } from "../src/finance/store.js";
import { FinanceWorkbookSync } from "../src/finance/workbooks.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { OfficeFiles } from "../src/files/office.js";
import { runTool } from "../src/tools.js";

const owner = "6281234567890";

function ledgerRow(index, { kind, amount, category, description, date, currency = "IDR" }) {
  return {
    id: `T-${String(index).padStart(8, "0")}`,
    book_id: "B-1",
    kind,
    amount_minor: Math.round(amount),
    currency,
    occurred_on: date,
    category,
    description,
    approved_at: `${date}T09:00:00.000Z`,
    status: "active",
  };
}

function sampleRows() {
  const rows = [];
  let index = 0;
  const push = (options) => {
    index += 1;
    rows.push(ledgerRow(index, options));
  };
  for (let month = 1; month <= 6; month += 1) {
    const mm = String(month).padStart(2, "0");
    push({ kind: "income", amount: 12_000_000, category: "Gaji", description: "Gaji", date: `2026-${mm}-01` });
    push({ kind: "income", amount: 2_000_000 + month * 100_000, category: "Proyek", description: "Proyek", date: `2026-${mm}-12` });
    push({ kind: "expense", amount: 4_500_000, category: "Sewa", description: "Sewa", date: `2026-${mm}-02` });
    push({ kind: "expense", amount: 1_800_000 + month * 50_000, category: "Makan", description: "Belanja", date: `2026-${mm}-05` });
    push({ kind: "expense", amount: 600_000, category: "Transport", description: "Bensin", date: `2026-${mm}-07` });
    push({ kind: "expense", amount: 300_000, category: "Hiburan", description: "Jajan", date: `2026-${mm}-15` });
  }
  push({ kind: "expense", amount: 8_500_000, category: "Elektronik", description: "Laptop", date: "2026-03-20" });
  return rows;
}

test("analyzeBook derives totals, monthly series, categories and insights", () => {
  const rows = sampleRows();
  const analysis = analyzeBook({ transactions: rows, currency: "IDR", bookName: "Buku Utama", today: "2026-06-30" });

  assert.equal(analysis.totals.count, rows.length);
  assert.equal(analysis.months.length, 6);
  assert.equal(analysis.months[0].key, "2026-01");
  assert.ok(analysis.totals.income > analysis.totals.expense);
  assert.equal(
    analysis.months.reduce((total, month) => total + month.income, 0),
    analysis.totals.income,
  );
  assert.equal(analysis.expenseCategories[0].name, "Sewa");
  assert.equal(analysis.expenseCategories[0].kind, "expense");
  assert.ok(analysis.expenseCategories[0].share > 0);
  assert.ok(analysis.insights.length >= 4);
  assert.ok(analysis.insights.every((insight) => insight.title && insight.detail));
  assert.equal(analysis.anomalies[0].description, "Laptop");
  assert.equal(analysis.otherCurrencies.length, 0);
});

test("analyzeBook flags foreign currency instead of mixing totals", () => {
  const rows = [
    ledgerRow(1, { kind: "expense", amount: 100_000, category: "Makan", description: "Lokal", date: "2026-01-02" }),
    ledgerRow(2, { kind: "expense", amount: 50_00, category: "Makan", description: "USD", date: "2026-01-03", currency: "USD" }),
  ];
  const analysis = analyzeBook({ transactions: rows, currency: "IDR", today: "2026-01-31" });
  assert.equal(analysis.totals.count, 1);
  assert.deepEqual(analysis.otherCurrencies, ["USD"]);
  assert.equal(primaryCurrency(rows, "IDR"), "IDR");
});

test("renderReport builds a formatted multi-sheet workbook with native charts", async () => {
  const rows = sampleRows();
  const analysis = analyzeBook({ transactions: rows, currency: "IDR", bookName: "Buku Utama", today: "2026-06-30" });
  const { bytes, charts } = await renderReport({ analysis, transactions: rows });

  assert.equal(charts, 3);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  assert.deepEqual(
    workbook.worksheets.map((sheet) => sheet.name),
    ["Ringkasan", "Arus Kas", "Kategori", "Transaksi", "Anggaran", "Anomali", "Petunjuk"],
  );

  const summary = workbook.getWorksheet("Ringkasan");
  assert.equal(summary.getCell("A1").value, "LAPORAN KEUANGAN");
  assert.equal(summary.views[0].showGridLines, false);
  assert.equal(summary.getCell("A5").value.formula, "'Arus Kas'!B11");
  assert.equal(summary.getCell("A5").value.result, analysis.totals.income);
  assert.equal(summary.getCell("A4").fill.fgColor.argb, "FFF6F8F7");

  const cash = workbook.getWorksheet("Arus Kas");
  assert.equal(cash.getCell("D5").value.formula, "B5-C5");
  assert.equal(cash.getCell("F5").value.formula, "SUM($D$5:D5)");
  assert.ok(cash.getCell("A11").value, "total row label");
  assert.equal(cash.getCell("B11").value.result, analysis.totals.income);

  const ledger = workbook.getWorksheet("Transaksi");
  assert.ok(ledger.getTable("BukuTransaksi"), "ledger keeps a real Excel table");
  assert.equal(ledger.getCell("C1").value, "Jenis");

  const budget = workbook.getWorksheet("Anggaran");
  assert.match(budget.getCell("C5").value.formula, /SUMIFS\('Transaksi'!\$G/);
  assert.match(budget.getCell("F5").value.formula, /Lewat anggaran/);
  assert.equal(budget.getCell("B5").fill.fgColor.argb, "FFFFF8E1");

  const parts = readZip(bytes);
  assert.ok(parts.has("xl/charts/chart1.xml"));
  assert.ok(parts.has("xl/drawings/drawing1.xml"));
  assert.match(parts.get("xl/charts/chart1.xml").toString("utf8"), /<c:barChart>/);
  assert.match(parts.get("xl/charts/chart2.xml").toString("utf8"), /<c:doughnutChart>/);
  const summarySheet = [...parts.keys()].find((name) => name === "xl/worksheets/sheet1.xml");
  const summaryXml = parts.get(summarySheet).toString("utf8");
  assert.match(summaryXml, /<drawing r:id="rId/);
  // CT_Worksheet order is strict. Excel treats a <drawing> after <tableParts> or
  // <extLst> as corruption and "recovers" the workbook as empty.
  const drawingAt = summaryXml.indexOf("<drawing ");
  // The trailing <extLst> is the sheet-level one; earlier ones live inside
  // conditional-formatting extensions and are not siblings of <drawing>.
  for (const tag of ["<tableParts", "<extLst"]) {
    const at = summaryXml.lastIndexOf(tag);
    if (at > -1) assert.ok(drawingAt < at, `<drawing> must come before ${tag}`);
  }
  const ledgerSheet = [...parts.keys()].find((name) => name === "xl/worksheets/sheet4.xml");
  const ledgerXml = parts.get(ledgerSheet).toString("utf8");
  assert.ok(ledgerXml.indexOf("<drawing ") < ledgerXml.indexOf("<tableParts"), "drawing before tableParts");
  const contentTypes = parts.get("[Content_Types].xml").toString("utf8");
  assert.match(contentTypes, /drawingml\.chart\+xml/);
});

test("renderReport survives an empty book without charts", async () => {
  const analysis = analyzeBook({ transactions: [], currency: "IDR", bookName: "Kosong", today: "2026-06-30" });
  const { bytes, charts } = await renderReport({ analysis, transactions: [] });
  assert.equal(charts, 0);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  assert.equal(workbook.worksheets.length, 7);
  assert.match(String(workbook.getWorksheet("Ringkasan").getCell("B9").value), /Belum ada transaksi/);
});

test("create_finance_report renders and sends an xlsx for the active book", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-report-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const workbookSync = new FinanceWorkbookSync({ financeStore, fileWorkspace, defaultCurrency: "IDR" });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  for (const row of sampleRows()) {
    financeStore.recordTransaction(owner, {
      kind: row.kind,
      amount: row.amount_minor,
      currency: row.currency,
      date: row.occurred_on,
      category: row.category,
      description: row.description,
    });
  }

  let sent;
  const result = await runTool("create_finance_report", "Asia/Jakarta", { period: "2026-03" }, {
    financeStore,
    financeWorkbookSync: workbookSync,
    fileWorkspace,
    userId: owner,
    defaultCurrency: "IDR",
    sendFile: async (file) => { sent = file; },
  });

  assert.equal(result.sent, true);
  assert.equal(result.charts >= 0, true);
  assert.match(sent.filename, /^laporan-buku-utama-2026-03\.xlsx$/);
  assert.equal(sent.mimeType, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  assert.ok(result.insights.length >= 3);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(readFileSync(await fileWorkspace.getFilePath(owner, result.file_id)));
  const ledger = workbook.getWorksheet("Transaksi");
  const marchRows = sampleRows().filter((row) => row.occurred_on.startsWith("2026-03")).length;
  assert.equal(ledger.rowCount, marchRows + 1, "header + March transactions only");
});

test("renderReport can drop sections the user does not want", async () => {
  const rows = sampleRows();
  const analysis = analyzeBook({ transactions: rows, currency: "IDR", bookName: "Buku Utama", today: "2026-06-30" });
  const { bytes, sheets, charts } = await renderReport({ analysis, transactions: rows, sheets: ["Ringkasan", "Kategori"] });

  assert.deepEqual(sheets, ["Ringkasan", "Kategori"]);
  assert.equal(charts, 1, "only the doughnut chart survives");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Ringkasan", "Kategori"]);
  // KPI falls back to static values because the Arus Kas sheet is absent.
  const kpi = workbook.getWorksheet("Ringkasan").getCell("A5").value;
  assert.equal(typeof kpi, "number");
  assert.equal(kpi, analysis.totals.income);
});

test("renderReport keeps the ledger when the budget sheet is requested", async () => {
  const rows = sampleRows();
  const analysis = analyzeBook({ transactions: rows, currency: "IDR", bookName: "Buku Utama", today: "2026-06-30" });
  const { sheets } = await renderReport({ analysis, transactions: rows, sheets: ["Anggaran"] });
  assert.deepEqual(sheets, ["Transaksi", "Anggaran"]);
});

test("renderReport rejects unknown sections instead of silently dropping them", async () => {
  const analysis = analyzeBook({ transactions: [], currency: "IDR", bookName: "X", today: "2026-06-30" });
  await assert.rejects(
    () => renderReport({ analysis, transactions: [], sheets: ["Neraca"] }),
    /Unknown report section/,
  );
});

test("create_finance_report honours a custom title and section list", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-report-custom-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const workbookSync = new FinanceWorkbookSync({ financeStore, fileWorkspace, defaultCurrency: "IDR" });
  t.after(() => {
    financeStore.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  financeStore.recordTransaction(owner, {
    kind: "expense", amount: 175_000, date: "2026-05-09", category: "Makan", description: "Bakso",
  });

  const result = await runTool("create_finance_report", "Asia/Jakarta", {
    sheets: ["Ringkasan", "Transaksi"],
    title: "Laporan Toko Saya",
    send: false,
  }, {
    financeStore, financeWorkbookSync: workbookSync, fileWorkspace, userId: owner, defaultCurrency: "IDR",
  });

  assert.equal(result.sent, false);
  assert.deepEqual(result.sheets, ["Ringkasan", "Transaksi"]);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(readFileSync(await fileWorkspace.getFilePath(owner, result.file_id)));
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Ringkasan", "Transaksi"]);
  assert.equal(workbook.getWorksheet("Ringkasan").getCell("A1").value, "Laporan Toko Saya");
});

test("get_finance_summary returns analysis the assistant can narrate", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "finance-summary-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  t.after(() => {
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  financeStore.recordTransaction(owner, {
    kind: "expense", amount: 250_000, date: "2026-05-04", category: "Makan", description: "Makan siang",
  });

  const result = await runTool("get_finance_summary", "Asia/Jakarta", { month: "2026-05" }, {
    financeStore, userId: owner, defaultCurrency: "IDR",
  });
  assert.equal(result.totals.expense, 250_000);
  assert.equal(result.expense_categories[0].category, "Makan");
  assert.ok(result.insights.length >= 2);
});

test("editing a charted workbook keeps its native charts", async () => {
  const dir = mkdtempSync(join(tmpdir(), "chart-preserve-"));
  const owner = "6281234567890";
  const store = new FinanceStore(join(dir, "finance.sqlite"));
  const workspace = new FileWorkspace({ dataDirectory: dir });
  try {
    const office = new OfficeFiles(workspace, store);
    const sync = new FinanceWorkbookSync({ financeStore: store, fileWorkspace: workspace, defaultCurrency: "IDR" });
    for (const [kind, amount, category] of [
      ["income", 5_000_000, "Gaji"],
      ["expense", 1_500_000, "Sewa"],
      ["income", 7_000_000, "Gaji"],
      ["expense", 2_000_000, "Makan"],
    ]) store.recordTransaction(owner, { kind, amount, category, description: category });

    const report = await sync.renderReportFile(owner, {});
    const before = readZip(await readFile(await workspace.getFilePath(owner, report.file.id)));
    const chartsBefore = [...before.keys()].filter((name) => /^xl\/charts\//.test(name)).length;
    assert.ok(chartsBefore > 0, "fixture must contain a native chart");

    await office.updateSpreadsheetCells(owner, {
      fileId: report.file.id,
      sheetName: "Ringkasan",
      updates: [{ cell: "A1", value: "LAPORAN KEUANGAN 2026" }],
      saveMode: "in_place",
    });

    const after = readZip(await readFile(await workspace.getFilePath(owner, report.file.id)));
    assert.equal([...after.keys()].filter((name) => /^xl\/charts\//.test(name)).length, chartsBefore);
    const sheetXml = after.get("xl/worksheets/sheet1.xml").toString("utf8");
    assert.match(sheetXml, /<drawing r:id="rId/);
    // The drawing must be a sibling of the sheet-level elements, never inside a
    // conditional-formatting rule, or Excel silently drops the chart.
    const drawingAt = sheetXml.indexOf("<drawing ");
    assert.ok(drawingAt > sheetXml.indexOf("</conditionalFormatting>") || !sheetXml.includes("</conditionalFormatting>"));
    assert.ok(drawingAt > sheetXml.indexOf("</sheetData>"));
  } finally {
    store.close();
    workspace.close();
    rmSync(dir, { recursive: true, force: true });
  }
});

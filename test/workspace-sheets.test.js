import assert from "node:assert/strict";
import test from "node:test";
import ExcelJS from "exceljs";
import { normalizeWorkspace, validateWorkspace, exportWorkspaceXlsx, importWorkspaceXlsx } from "../src/workspace-sheets.js";
import { convertWorkbookToDataset } from "../src/share-convert.js";

test("workspace supports sheets, unlimited practical columns, and formulas", async () => {
  const data = validateWorkspace({ title: "Pertanian", sheets: [
    { name: "Transaksi", columns: ["Komoditas", "Pemasukan", "Pengeluaran", "Saldo"], rows: [["Kopi", 1000000, 250000, { formula: "B2-C2" }]] },
    { name: "Ringkasan", columns: ["Total"], rows: [[{ formula: "SUM(Transaksi!B2:B2)" }]] },
  ] });
  assert.equal(data.sheets.length, 2);
  // The named columns became row 1, so the data row is row 2 — which is exactly
  // what the stored formula "B2-C2" points at.
  assert.deepEqual(data.sheets[0].rows[0], ["Komoditas", "Pemasukan", "Pengeluaran", "Saldo"]);
  assert.deepEqual(data.sheets[0].rows[1][3], { formula: "B2-C2" });
  // All formulas accepted — stored as-is, never evaluated
  const unsafe = await exportWorkspaceXlsx({ sheets: [{ name: "X", columns: ["A"], rows: [[{ formula: "VLOOKUP(A1, B1:C10, 2, FALSE)" }]] }] });
  assert.deepEqual([...unsafe].slice(0, 2), [0x50, 0x4b]);
  const bytes = await exportWorkspaceXlsx(data);
  assert.deepEqual([...bytes].slice(0, 2), [0x50, 0x4b]);
});

test("import keeps every row at its original Excel position", async () => {
  const buf = await exportWorkspaceXlsx({ title: "Test", sheets: [
    { name: "Data", columns: ["Nama", "Nilai"], rows: [["A", 100], ["B", 200]] },
  ] });
  const imported = await importWorkspaceXlsx(buf, "Reimport");
  assert.equal(imported.sheets.length, 1);
  // Letter columns are grid coordinates, so no header row is exported and the
  // imported sheet keeps the same A1 coordinates the user had in Excel.
  assert.deepEqual(imported.sheets[0].columns, ["A", "B"]);
  assert.deepEqual(imported.sheets[0].rows[0], ["Nama", "Nilai"]);
  assert.equal(imported.sheets[0].rows.length, 3);
});

test("imported formulas carry computed values, not raw text", async () => {
  const buf = await exportWorkspaceXlsx({ title: "Calc", sheets: [
    { name: "Data", columns: ["A", "B", "C"], rows: [[10, 20, { formula: "A1+B1" }]] },
  ] });
  const imported = await importWorkspaceXlsx(buf, "Calc");
  const cell = imported.sheets[0].rows[0][2];
  assert.equal(cell.formula, "A1+B1");
  assert.equal(cell.value, 30);
});

test("named columns still export a header row", async () => {
  const buf = await exportWorkspaceXlsx({ title: "Named", sheets: [
    { name: "Data", columns: ["Komoditas", "Nilai"], rows: [["Kopi", 100]] },
  ] });
  const imported = await importWorkspaceXlsx(buf, "Named");
  assert.deepEqual(imported.sheets[0].rows[0], ["Komoditas", "Nilai"]);
  assert.deepEqual(imported.sheets[0].rows[1], ["Kopi", 100]);
});

test("legacy one-sheet dataset normalizes without data loss", () => {
  const data = normalizeWorkspace({ title: "Legacy", columns: ["A", "B"], rows: [[1, 2]] });
  assert.deepEqual(data.sheets[0].columns, ["A", "B"]);
  assert.deepEqual(data.sheets[0].rows, [[1, 2]]);
});

test("named columns become a real header row, not invisible metadata", () => {
  // share_link_set_rows writes named columns plus data rows only. The web grid
  // renders rows, never `columns`, so a header kept in `columns` used to be
  // invisible on the web while the XLSX export re-created it.
  const data = validateWorkspace({
    title: "Kas",
    sheets: [{ name: "Kas", columns: ["Tanggal", "Keterangan", "Pemasukan"], rows: [["2026-01-03", "Penjualan", 1500000]] }],
  });
  const sheet = data.sheets[0];
  assert.deepEqual(sheet.columns, ["A", "B", "C"], "columns are grid letters now");
  assert.deepEqual(sheet.rows[0], ["Tanggal", "Keterangan", "Pemasukan"], "header is row 1");
  assert.deepEqual(sheet.rows[1], ["2026-01-03", "Penjualan", 1500000]);
  assert.equal(sheet.headerRow, true);

  // Normalizing twice must not duplicate the header.
  const twice = validateWorkspace(data);
  assert.deepEqual(twice.sheets[0].rows, sheet.rows);
});

test("recorded styles move down with the materialized header", () => {
  const data = validateWorkspace({
    title: "Kas",
    sheets: [{ name: "Kas", columns: ["Tanggal", "Nilai"], rows: [["2026-01-03", 100]], styles: { "0,0": { bold: true } } }],
  });
  assert.deepEqual(data.sheets[0].styles, { "1,0": { bold: true } }, "the bold cell followed its row");
});

test("export writes the header once and keeps it bold and frozen", async () => {
  const bytes = await exportWorkspaceXlsx({
    title: "Kas",
    sheets: [{ name: "Kas", columns: ["Tanggal", "Keterangan"], rows: [["2026-01-03", "Penjualan"]] }],
  });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  const sheet = workbook.worksheets[0];
  assert.deepEqual(sheet.getRow(1).values.slice(1, 3), ["Tanggal", "Keterangan"]);
  assert.deepEqual(sheet.getRow(2).values.slice(1, 3), ["2026-01-03", "Penjualan"]);
  assert.equal(sheet.getRow(1).font.bold, true);
  assert.equal(sheet.views[0].state, "frozen");
});

test("converting an .xlsx keeps its title block, blank rows and row numbers", async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Pembukuan");
  worksheet.addRow(["PEMBUKUAN WARUNG"]);
  worksheet.mergeCells("A1:D1");
  worksheet.addRow(["Total Pemasukan", 15000000]);
  worksheet.addRow([]);
  worksheet.addRow(["Tanggal", "Keterangan", "Pemasukan", "Pengeluaran"]);
  worksheet.addRow(["2026-01-03", "Penjualan", 1500000, 0]);
  const dataset = await convertWorkbookToDataset(Buffer.from(await workbook.xlsx.writeBuffer()), {});

  assert.deepEqual(dataset.columns, ["A", "B", "C", "D"], "letters, so nothing is lost as metadata");
  assert.equal(dataset.rows[0][0], "PEMBUKUAN WARUNG", "title row survives");
  assert.equal(dataset.rows[0][1], "", "a merged title is stored once, not repeated");
  assert.deepEqual(dataset.rows[1], ["Total Pemasukan", "15000000", "", ""]);
  assert.equal(dataset.rows[3][0], "Tanggal", "the blank row keeps the header on row 4");
  assert.equal(dataset.rows[4][0], "2026-01-03");
});

import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { FileWorkspace } from "../src/files/workspace.js";
import { OfficeFiles } from "../src/files/office.js";
import { ShareLinkStore } from "../src/share-links.js";
import { runFileTool } from "../src/files/tools.js";
import { appendRows, setRow, deleteRow, describeForAgent, auditFormulas, shiftFormulaRows, crossSheetRefs } from "../src/share-edit.js";
import { computeWorkbook } from "../src/formula-engine.js";
import { buildShareDataset } from "../src/share-dataset.js";
import { extractSummary } from "../src/workspace-summary.js";

function context(t) {
  const directory = mkdtempSync(join(tmpdir(), "share-edit-test-"));
  const fileWorkspace = new FileWorkspace({ dataDirectory: directory });
  const shareLinks = new ShareLinkStore({ dataDirectory: directory });
  const officeFiles = new OfficeFiles(fileWorkspace);
  t.after(() => {
    shareLinks.close();
    fileWorkspace.close();
    rmSync(directory, { recursive: true, force: true });
  });
  return { fileWorkspace, officeFiles, shareLinks, userId: "6281234567890", publicOrigin: "https://example.test" };
}

/** A farm workbook in the shape the skill prescribes. */
const farm = () => ({
  sheets: [
    { name: "Dashboard", columns: ["Komponen", "Nilai"], headerRow: true, rows: [
      ["Total Pemasukan", { formula: "SUM(Pemasukan!G:G)" }],
      ["Total Pengeluaran", { formula: "SUM(Pengeluaran!G:G)" }],
      ["Laba Bersih", { formula: "B2-B3" }],
      ["Margin", { formula: "IF(B2>0,B4/B2,0)" }],
      [],
      ["Komoditas", "Volume", "Total Pemasukan", "Total Biaya", "Laba Bersih", "Margin"],
      ["Kopi", { formula: "SUMIFS(Pemasukan!D:D,Pemasukan!C:C,A8)" }, { formula: "SUMIFS(Pemasukan!G:G,Pemasukan!C:C,A8)" }, { formula: "SUMIFS(Pengeluaran!G:G,Pengeluaran!E:E,A8)" }, { formula: "C8-D8" }, { formula: "IF(C8>0,E8/C8,0)" }],
    ] },
    { name: "Pemasukan", columns: ["Tanggal", "Bulan", "Komoditas", "Kuantitas", "Satuan", "Harga Satuan", "Total"], headerRow: true, rows: [
      ["2024-01-15", "2024-01", "Kopi", 370, "kg", 80000, { formula: "D2*F2" }],
    ] },
    { name: "Pengeluaran", columns: ["Tanggal", "Bulan", "Kategori Biaya", "Uraian", "Komoditas", "Kuantitas", "Total"], headerRow: true, rows: [
      ["2024-01-05", "2024-01", "Pupuk", "NPK", "Kopi", 1, 150000],
    ] },
    { name: "Laba Rugi Bulanan", columns: ["Periode", "Pemasukan", "Pengeluaran", "Laba Bersih", "Kumulatif"], headerRow: true, rows: [
      ["2024-01", { formula: "SUMIFS(Pemasukan!G:G,Pemasukan!B:B,A2)" }, { formula: "SUMIFS(Pengeluaran!G:G,Pengeluaran!B:B,A2)" }, { formula: "B2-C2" }, { formula: "D2" }],
    ] },
  ],
});

test("appending one row keeps the other sheets and inherits the row formula", () => {
  let dataset = computeWorkbook(buildShareDataset(farm()));
  const before = extractSummary(dataset);
  assert.equal(before.kpis[0].value, 29600000);

  dataset = appendRows(dataset, {
    sheet: "Pemasukan",
    // Only the values the user gave: the Total cell is left null on purpose.
    rows: [["2024-02-03", "2024-02", "Cabai Rawit", 50, "kg", 60000, null]],
  });
  dataset = computeWorkbook(dataset);

  // Every sheet survives, not just the one that was edited.
  assert.deepEqual(dataset.sheets.map((sheet) => sheet.name),
    ["Dashboard", "Pemasukan", "Pengeluaran", "Laba Rugi Bulanan"]);
  const rows = dataset.sheets[1].rows;
  assert.equal(rows.length, 3);
  assert.deepEqual(rows[2].slice(0, 6), ["2024-02-03", "2024-02", "Cabai Rawit", 50, "kg", 60000]);
  // The Total formula was copied down and shifted, not left blank.
  assert.deepEqual(rows[2][6], { formula: "D3*F3", value: 3000000 });

  // The dashboard totals now include the new row, so the sheets really are wired.
  const after = extractSummary(dataset);
  assert.equal(after.kpis[0].value, 32600000);
});

test("a null cell copies the style of the row above", () => {
  const base = buildShareDataset({
    sheets: [{ name: "Data", columns: ["A", "B"], rows: [["x", 1]], styles: { "1,1": { align: "right", bold: true } } }],
  });
  const next = appendRows(base, { sheet: "Data", rows: [["y", null]] });
  assert.deepEqual(next.sheets[0].styles["2,1"], { align: "right", bold: true });
});

test("appending refuses a sheet that does not exist and names the real ones", () => {
  const dataset = buildShareDataset(farm());
  assert.throws(() => appendRows(dataset, { sheet: "Penjualan", rows: [["x"]] }), /No sheet named "Penjualan".*Pemasukan/s);
});

test("set_row corrects a cell in place and delete_row removes one row", () => {
  let dataset = computeWorkbook(buildShareDataset(farm()));
  dataset = setRow(dataset, { sheet: "Pemasukan", row: 2, values: ["2024-01-15", "2024-01", "Kopi Arabika"] });
  assert.equal(dataset.sheets[1].rows[1][2], "Kopi Arabika");
  dataset = deleteRow(dataset, { sheet: "Pemasukan", row: 2 });
  assert.equal(dataset.sheets[1].rows.length, 1, "only the header remains");
});

test("deleting the header row is refused", () => {
  const dataset = buildShareDataset(farm());
  assert.throws(() => deleteRow(dataset, { sheet: "Pemasukan", row: 1 }), /Refusing to delete the header row/);
});

test("an off-by-one column reference is reported with the columns that do hold data", () => {
  const dataset = computeWorkbook(buildShareDataset({
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], headerRow: true, rows: [
        // H instead of G: the classic slip, and it silently totals nothing.
        ["Total Pemasukan", { formula: "SUM(Pemasukan!H:H)" }],
      ] },
      { name: "Pemasukan", columns: ["Tanggal", "Bulan", "Komoditas", "Kuantitas", "Satuan", "Harga Satuan", "Total"], headerRow: true, rows: [
        ["2024-01-15", "2024-01", "Kopi", 370, "kg", 80000, { formula: "D2*F2" }],
      ] },
    ],
  }));
  const issues = auditFormulas(dataset);
  assert.equal(issues.length, 1);
  assert.equal(issues[0].cell, "B2");
  assert.match(issues[0].problem, /column H of "Pemasukan" has no data/);
  assert.match(issues[0].hint, /G=Total/);
});

test("a formula pointing at a sheet that does not exist is reported", () => {
  const dataset = computeWorkbook(buildShareDataset({
    sheets: [{ name: "Dashboard", columns: ["Komponen", "Nilai"], headerRow: true, rows: [
      ["Total", { formula: "SUM(Penjualan!G:G)" }],
    ] }],
  }));
  const issues = auditFormulas(dataset);
  assert.equal(issues.length, 1);
  assert.match(issues[0].problem, /references sheet "Penjualan" which does not exist/);
  assert.match(issues[0].hint, /Dashboard/);
});

test("a healthy workbook reports no issues", () => {
  const dataset = computeWorkbook(buildShareDataset(farm()));
  assert.deepEqual(auditFormulas(dataset), []);
});

test("describe gives the agent column letters without reading the whole file", () => {
  const dataset = computeWorkbook(buildShareDataset(farm()));
  const brief = describeForAgent(dataset, { sheetName: "Pemasukan" });
  assert.equal(brief.sheet_count, 4);
  assert.equal(brief.sheets.length, 1);
  assert.equal(brief.sheets[0].name, "Pemasukan");
  assert.equal(brief.sheets[0].data_rows, 1);
  assert.equal(brief.sheets[0].columns, "A=Tanggal, B=Bulan, C=Komoditas, D=Kuantitas, E=Satuan, F=Harga Satuan, G=Total");
  assert.deepEqual(brief.sheets[0].last_rows[0].slice(0, 3), ["2024-01-15", "2024-01", "Kopi"]);
});

test("formula row shifting leaves absolute rows and literals alone", () => {
  assert.equal(shiftFormulaRows("D2*F2", 2), "D4*F4");
  assert.equal(shiftFormulaRows("SUM($D$2:D2)", 1), "SUM($D$2:D3)");
  assert.equal(shiftFormulaRows('IF(A2="B2",B2,0)', 1), 'IF(A3="B2",B3,0)');
});

test("cross-sheet references are read out of quoted names and literals", () => {
  const refs = crossSheetRefs(`SUMIFS('Laba Rugi'!G:G, Pemasukan!C:C, "Kopi!A1")`);
  assert.deepEqual(refs.map((ref) => ref.sheet), ["Laba Rugi", "Pemasukan"]);
  assert.equal(refs[0].whole_column, true);
});

test("the tool appends to a stored document instead of rewriting it", async (t) => {
  const ctx = context(t);
  await runFileTool("workspace_files", {
    action: "share_link_create", folder: "documents/kebun", name: "kebun", title: "Kebun",
  }, ctx);
  await runFileTool("workspace_files", {
    action: "share_link_set_rows", folder: "documents/kebun", ...farm(),
  }, ctx);

  // The follow-up the user actually asks for.
  const appended = await runFileTool("workspace_files", {
    action: "share_link_append_rows",
    folder: "documents/kebun",
    sheet_name: "Pemasukan",
    rows: [["2024-02-03", "2024-02", "Cabai Rawit", 50, "kg", 60000, null]],
  }, ctx);
  assert.equal(appended.saved, true);
  assert.equal(appended.sheet_count, 4, "every sheet must survive an append");
  assert.equal(appended.summary.kpis[0].value, 32600000);
  assert.equal(appended.formula_issues, undefined, "a clean append reports no issues");

  // A second append keeps working, which is what used to time out.
  const again = await runFileTool("workspace_files", {
    action: "share_link_append_rows",
    folder: "documents/kebun",
    sheet_name: "Pemasukan",
    rows: [["2024-02-10", "2024-02", "Pepaya", 20, "kg", 15000, null]],
  }, ctx);
  // 29.600.000 + 3.000.000 (cabai) + 300.000 (pepaya)
  assert.equal(again.summary.kpis[0].value, 32900000);
});

test("the tool describes and edits an existing document", async (t) => {
  const ctx = context(t);
  await runFileTool("workspace_files", {
    action: "share_link_create", folder: "documents/kebun2", name: "kebun2",
  }, ctx);
  await runFileTool("workspace_files", {
    action: "share_link_set_rows", folder: "documents/kebun2", ...farm(),
  }, ctx);

  const described = await runFileTool("workspace_files", {
    action: "share_link_describe", folder: "documents/kebun2",
  }, ctx);
  assert.equal(described.sheet_count, 4);
  assert.match(described.sheets.find((sheet) => sheet.name === "Pemasukan").columns, /G=Total/);

  await runFileTool("workspace_files", {
    action: "share_link_set_row", folder: "documents/kebun2", sheet_name: "Pemasukan", row: 2,
    values: ["2024-01-15", "2024-01", "Kopi Arabika", 370, "kg", 80000, null],
  }, ctx);
  const afterEdit = await runFileTool("workspace_files", {
    action: "share_link_describe", folder: "documents/kebun2", sheet_name: "Pemasukan",
  }, ctx);
  assert.equal(afterEdit.sheets[0].last_rows[0][2], "Kopi Arabika");

  await runFileTool("workspace_files", {
    action: "share_link_delete_row", folder: "documents/kebun2", sheet_name: "Pemasukan", row: 2,
  }, ctx);
  const afterDelete = await runFileTool("workspace_files", {
    action: "share_link_describe", folder: "documents/kebun2", sheet_name: "Pemasukan",
  }, ctx);
  assert.equal(afterDelete.sheets[0].data_rows, 0);
});

test("editing a folder with no document says so instead of failing deep inside", async (t) => {
  const ctx = context(t);
  await assert.rejects(
    runFileTool("workspace_files", {
      action: "share_link_append_rows", folder: "documents/kosong", sheet_name: "Data", rows: [["x"]],
    }, ctx),
    /No web document in "documents\/kosong" yet/,
  );
});

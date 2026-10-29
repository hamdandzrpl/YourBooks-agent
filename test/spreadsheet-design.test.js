import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import { FinanceStore } from "../src/finance/store.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { OfficeFiles } from "../src/files/office.js";
import { FinanceWorkbookSync } from "../src/finance/workbooks.js";
import { runFileTool } from "../src/files/tools.js";
import { readZip } from "../src/finance/xlsx/zip.js";

const owner = "6281234567890";

function withOffice(t) {
  const directory = mkdtempSync(join(tmpdir(), "spreadsheet-design-"));
  const workspace = new FileWorkspace({ dataDirectory: directory });
  const store = new FinanceStore(join(directory, "finance.sqlite"));
  const office = new OfficeFiles(workspace, store, "IDR");
  t.after(() => {
    store.close();
    workspace.close();
    rmSync(directory, { recursive: true, force: true });
  });
  return { workspace, store, office, directory };
}

async function load(workspace, userId, fileId) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await readFile(await workspace.getFilePath(userId, fileId)));
  return workbook;
}

async function seed(office, { rows = [] } = {}) {
  return office.createSpreadsheet(owner, {
    filename: "kas.xlsx",
    sheetName: "Kas",
    columns: ["Bulan", "Masuk", "Keluar"],
    rows,
  });
}

test("update_cells writes values, formulas, look, merges, widths, and heights in one call", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 10_000_000, 2_500_000], ["Feb", 12_000_000, 3_000_000]] });

  const updated = await office.updateSpreadsheetCells(owner, {
    fileId: file.id,
    sheetName: "Kas",
    saveMode: "in_place",
    updates: [
      {
        cell: "A7",
        value: "UANG KOTOR",
        format: { bold: true, size: 12, color: "FFFFFF", fill: "0B3B33", align: "center", border: true },
      },
      { range: "A7:C7", merge: true },
      { cell: "B5", formula: "SUM(B2:B3)", format: { numFmt: "currency", bold: true, align: "right" } },
      { column: "A", width: 30 },
      { row: 7, height: 26 },
    ],
  });

  const workbook = await load(workspace, owner, updated.id);
  const sheet = workbook.getWorksheet("Kas");
  assert.equal(sheet.getCell("A7").value, "UANG KOTOR");
  assert.deepEqual(sheet.getCell("B5").value, { formula: "SUM(B2:B3)" });
  assert.equal(sheet.getCell("A7").font.bold, true);
  assert.equal(sheet.getCell("A7").fill.fgColor.argb, "FF0B3B33");
  assert.equal(sheet.getCell("B5").numFmt, "#,##0;[Red](#,##0);\"-\"");
  assert.equal(sheet.getColumn(1).width, 30);
  assert.equal(sheet.getRow(7).height, 26);
  assert.ok(sheet.model.merges.includes("A7:C7"));

  // The tool result is proof, not a promise: every changed cell is read back.
  const byAddress = new Map(updated.verified.map((entry) => [entry.cell, entry]));
  assert.equal(byAddress.get("A7").value, "UANG KOTOR");
  assert.equal(byAddress.get("B5").formula, "SUM(B2:B3)");
  assert.equal(byAddress.get("B5").bold, true);
});

test("merging over a filled range reports what it cleared", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 10_000_000, 2_500_000]] });
  const updated = await office.updateSpreadsheetCells(owner, {
    fileId: file.id,
    sheetName: "Kas",
    updates: [
      { cell: "A2", value: "Judul" },
      { range: "A2:C2", merge: true },
    ],
  });
  assert.equal(updated.verified.find((entry) => entry.cell === "A2").value, "Judul");
  assert.ok(updated.warnings.some((warning) => /keeps only the top-left cell/.test(warning)));
});

test("merging across an existing merge explains how to fix it", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 1, 2]] });
  await office.updateSpreadsheetCells(owner, {
    fileId: file.id, sheetName: "Kas", saveMode: "in_place",
    updates: [{ range: "A1:C1", merge: true }],
  });
  await assert.rejects(
    office.updateSpreadsheetCells(owner, {
      fileId: file.id, sheetName: "Kas",
      updates: [{ range: "A1:B2", merge: true }],
    }),
    /overlaps the existing merge A1:C1[\s\S]*Unmerge/,
  );
});

test("writing into a merged slave cell lands on the master and says so", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 1, 2]] });
  const updated = await office.updateSpreadsheetCells(owner, {
    fileId: file.id, sheetName: "Kas", saveMode: "in_place",
    updates: [
      { range: "A1:C1", merge: true },
      { cell: "B1", value: "Judul Baru" },
    ],
  });
  assert.equal(updated.verified.find((entry) => entry.cell === "A1").value, "Judul Baru");
  assert.ok(updated.warnings.some((warning) => /written to A1/.test(warning)));
});

test("update_cells rejects an unknown format key instead of silently ignoring it", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office);
  await assert.rejects(
    office.updateSpreadsheetCells(owner, {
      fileId: file.id,
      sheetName: "Kas",
      updates: [{ cell: "A2", value: "x", format: { boldd: true } }],
    }),
    /Unknown format key "boldd"/,
  );
});

test("update_cells can copy the look of an existing cell", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office);
  const styled = await office.updateSpreadsheetCells(owner, {
    fileId: file.id,
    sheetName: "Kas",
    saveMode: "in_place",
    updates: [{ cell: "A1", format: { bold: true, fill: "0F766E", color: "FFFFFF", size: 14 } }],
  });
  await office.updateSpreadsheetCells(owner, {
    fileId: styled.id,
    sheetName: "Kas",
    saveMode: "in_place",
    updates: [{ cell: "A2", value: "Januari", style_from: "A1" }],
  });
  const sheet = (await load(workspace, owner, styled.id)).getWorksheet("Kas");
  assert.equal(sheet.getCell("A2").font.bold, true);
  assert.equal(sheet.getCell("A2").fill.fgColor.argb, "FF0F766E");
  assert.equal(sheet.getCell("A2").font.size, 14);
});

test("rows and columns can be inserted and deleted without losing neighbours", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 1, 2], ["Feb", 3, 4], ["Mar", 5, 6]] });

  await office.updateSpreadsheetCells(owner, {
    fileId: file.id,
    sheetName: "Kas",
    saveMode: "in_place",
    updates: [{ cell: "A2", format: { fill: "0F766E", bold: true } }],
  });
  const withRow = await office.insertSpreadsheetRows(owner, {
    fileId: file.id, sheetName: "Kas", at: 3, count: 2, saveMode: "in_place",
  });
  let sheet = (await load(workspace, owner, withRow.id)).getWorksheet("Kas");
  assert.equal(sheet.getCell("A3").value, null);
  assert.equal(sheet.getCell("A5").value, "Feb");
  // The inserted rows inherit the styling of the row above them.
  assert.equal(sheet.getCell("A3").fill.fgColor.argb, "FF0F766E");

  const removed = await office.deleteSpreadsheetRows(owner, {
    fileId: file.id, sheetName: "Kas", at: 3, count: 2, saveMode: "in_place",
  });
  sheet = (await load(workspace, owner, removed.id)).getWorksheet("Kas");
  assert.equal(sheet.getCell("A3").value, "Feb");
  assert.equal(sheet.getCell("A4").value, "Mar");

  const withColumn = await office.insertSpreadsheetColumns(owner, {
    fileId: file.id, sheetName: "Kas", at: 2, count: 1, saveMode: "in_place",
  });
  sheet = (await load(workspace, owner, withColumn.id)).getWorksheet("Kas");
  assert.equal(sheet.getCell("A1").value, "Bulan");
  assert.equal(sheet.getCell("B1").value, null);
  assert.equal(sheet.getCell("C1").value, "Masuk");

  const dropped = await office.deleteSpreadsheetColumns(owner, {
    fileId: file.id, sheetName: "Kas", at: 2, count: 1, saveMode: "in_place",
  });
  sheet = (await load(workspace, owner, dropped.id)).getWorksheet("Kas");
  assert.equal(sheet.getCell("B1").value, "Masuk");
});

test("worksheets can be added, renamed, duplicated, and deleted", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 1, 2]] });
  await office.updateSpreadsheetCells(owner, {
    fileId: file.id,
    sheetName: "Kas",
    saveMode: "in_place",
    updates: [
      { cell: "A1", format: { bold: true, fill: "0B3B33" } },
      { column: "A", width: 28 },
      { range: "A1:C1", merge: false },
    ],
  });

  const added = await office.manageSpreadsheetSheet(owner, {
    fileId: file.id, action: "add", newName: "Catatan",
  });
  assert.deepEqual(added.sheets, ["Kas", "Catatan"]);

  const duplicated = await office.manageSpreadsheetSheet(owner, {
    fileId: file.id, action: "duplicate", sheetName: "Kas", newName: "Kas Salinan",
  });
  assert.ok(duplicated.sheets.includes("Kas Salinan"));
  let workbook = await load(workspace, owner, file.id);
  const copy = workbook.getWorksheet("Kas Salinan");
  assert.equal(copy.getCell("A1").value, "Bulan");
  assert.equal(copy.getCell("A1").font.bold, true);
  assert.equal(copy.getColumn(1).width, 28);
  assert.equal(copy.getCell("A2").value, "Jan");

  const renamed = await office.manageSpreadsheetSheet(owner, {
    fileId: file.id, action: "rename", sheetName: "Catatan", newName: "Memo",
  });
  assert.ok(renamed.sheets.includes("Memo"));
  assert.ok(!renamed.sheets.includes("Catatan"));

  const deleted = await office.manageSpreadsheetSheet(owner, {
    fileId: file.id, action: "delete", sheetName: "Memo",
  });
  assert.ok(!deleted.sheets.includes("Memo"));
  workbook = await load(workspace, owner, file.id);
  assert.ok(workbook.getWorksheet("Kas"));
});

test("the last worksheet cannot be deleted", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office);
  await assert.rejects(
    office.manageSpreadsheetSheet(owner, { fileId: file.id, action: "delete", sheetName: "Kas" }),
    /at least one worksheet/,
  );
});

test("design adds coloured panels and native charts, and keeps existing charts", async (t) => {
  const { office, workspace, store } = withOffice(t);
  const sync = new FinanceWorkbookSync({ financeStore: store, fileWorkspace: workspace, defaultCurrency: "IDR" });
  for (const [kind, amount, category] of [
    ["income", 5_000_000, "Gaji"],
    ["expense", 1_500_000, "Sewa"],
    ["income", 7_000_000, "Proyek"],
  ]) store.recordTransaction(owner, { kind, amount, category, description: category });
  const report = await sync.renderReportFile(owner, {});
  const chartsBefore = [...readZip(await readFile(await workspace.getFilePath(owner, report.file.id))).keys()]
    .filter((name) => /^xl\/charts\//.test(name)).length;
  assert.ok(chartsBefore > 0);

  const designed = await office.designSpreadsheet(owner, {
    fileId: report.file.id,
    sheetName: "Ringkasan",
    saveMode: "in_place",
    shapes: [{
      from: "J44", to: "R47", text: "UANG KOTOR\nRp 12.000.000",
      fill: "0F766E", border: "0B3B33", text_color: "FFFFFF", font_size: 14, shadow: true,
    }],
    charts: [{
      type: "bar",
      title: "Kas Bulanan",
      from: "J49", to: "R64",
      categories: "Arus Kas!A5:A5",
      series: [{ name: "Pemasukan", range: "Arus Kas!B5:B5" }],
    }],
  });
  assert.equal(designed.added_shapes, 1);
  assert.equal(designed.added_charts, 1);
  assert.equal(designed.design_error, undefined);

  const bytes = await readFile(await workspace.getFilePath(owner, designed.id));
  const zip = readZip(bytes);
  const drawing = zip.get("xl/drawings/drawing1.xml").toString("utf8");
  assert.match(drawing, /<xdr:sp\b/);
  assert.match(drawing, /UANG KOTOR/);
  assert.match(drawing, /<a:srgbClr val="0F766E"\/>/);
  // Existing report charts survive, and the new chart part is a sibling.
  assert.equal([...zip.keys()].filter((name) => /^xl\/charts\//.test(name)).length, chartsBefore + 1);
  assert.equal([...zip.keys()].filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name)).length, 1);

  const workbook = await load(workspace, owner, designed.id);
  assert.ok(workbook.getWorksheet("Ringkasan"));
});

test("design reads chart values from live cells so the chart stays correct", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 10, 4], ["Feb", 20, 8]] });
  const designed = await office.designSpreadsheet(owner, {
    fileId: file.id,
    sheetName: "Kas",
    saveMode: "in_place",
    charts: [{
      type: "line",
      title: "Tren",
      from: "E2", to: "K18",
      categories: "A2:A3",
      series: [{ name: "Masuk", range: "B2:B3", color: "0F766E" }],
    }],
  });
  const zip = readZip(await readFile(await workspace.getFilePath(owner, designed.id)));
  const chartXml = [...zip.entries()]
    .filter(([name]) => /^xl\/charts\/chart\d+\.xml$/.test(name))
    .map(([, data]) => data.toString("utf8"))
    .join("");
  assert.match(chartXml, /&apos;Kas&apos;!B2:B3/);
  assert.match(chartXml, /<c:pt idx="0"><c:v>10<\/c:v><\/c:pt>/);
  assert.match(chartXml, /<c:pt idx="1"><c:v>20<\/c:v><\/c:pt>/);
});

test("a mismatched chart series range is refused instead of drawing a wrong chart", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 10, 4], ["Feb", 20, 8]] });
  await assert.rejects(
    office.designSpreadsheet(owner, {
      fileId: file.id,
      sheetName: "Kas",
      charts: [{
        type: "bar", from: "E2", to: "K18", categories: "A2:A3",
        series: [{ name: "Masuk", range: "B2:B2" }],
      }],
    }),
    /categories has 2/,
  );
});

test("repeated copy-mode edits allocate unique filenames instead of failing", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office);
  const first = await office.updateSpreadsheetCells(owner, {
    fileId: file.id, sheetName: "Kas", saveMode: "copy", updates: [{ cell: "A2", value: "satu" }],
  });
  const second = await office.updateSpreadsheetCells(owner, {
    fileId: file.id, sheetName: "Kas", saveMode: "copy", updates: [{ cell: "A3", value: "dua" }],
  });
  assert.equal(first.name, "kas-updated.xlsx");
  assert.equal(second.name, "kas-updated-2.xlsx");
  assert.notEqual(first.id, second.id);
});

test("editing defaults to in_place so the original workspace file is revised", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office);
  const edited = await office.updateSpreadsheetCells(owner, {
    fileId: file.id, sheetName: "Kas", updates: [{ cell: "A3", value: "tanpa save_mode" }],
  });
  assert.equal(edited.id, file.id, "no save_mode must revise the same file, not create a copy");
  assert.equal((await load(workspace, owner, file.id)).getWorksheet("Kas").getCell("A3").value, "tanpa save_mode");
  const copies = (await workspace.listFiles(owner, "exports")).filter((entry) => /updated/.test(entry.name));
  assert.equal(copies.length, 0, "no duplicate file may appear in the explorer for a default edit");
});

test("the spreadsheet tool exposes design and sheet actions with verified results", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office);
  const context = { fileWorkspace: workspace, officeFiles: office, userId: owner };

  const updated = await runFileTool("spreadsheet", {
    action: "update_cells",
    file_id: file.id,
    sheet_name: "Kas",
    save_mode: "in_place",
    updates: [{ cell: "A2", value: "Uang kotor", format: { bold: true, fill: "B08D57" } }],
  }, context);
  assert.equal(updated.verified[0].value, "Uang kotor");
  assert.equal(updated.verified[0].fill, "FFB08D57");
  assert.equal(updated.bytes, undefined, "workbook bytes must never reach the tool payload");

  const sheet = await runFileTool("spreadsheet", {
    action: "sheet", file_id: file.id, sheet_action: "add", new_name: "Ringkasan",
  }, context);
  assert.ok(sheet.sheets.includes("Ringkasan"));

  const designed = await runFileTool("spreadsheet", {
    action: "design",
    file_id: file.id,
    sheet_name: "Ringkasan",
    shapes: [{ from: "A1", to: "D4", text: "UANG KOTOR", fill: "0B3B33" }],
  }, context);
  assert.equal(designed.added_shapes, 1);
  assert.equal(designed.bytes, undefined);
});

test("inspect reports formulas, merges, widths, and existing drawings", async (t) => {
  const { office } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 10, 4]] });
  await office.updateSpreadsheetCells(owner, {
    fileId: file.id,
    sheetName: "Kas",
    saveMode: "in_place",
    updates: [
      { cell: "A4", formula: "SUM(B2:B3)" },
      { range: "A5:B5", merge: true },
      { column: "C", width: 22 },
    ],
  });
  const designed = await office.designSpreadsheet(owner, {
    fileId: file.id, sheetName: "Kas", saveMode: "in_place",
    shapes: [{ from: "E2", to: "H5", text: "Panel" }],
  });
  const inspection = await office.inspectSpreadsheet(owner, designed.id, { sheetName: "Kas", range: "A4:B5" });
  const sheet = inspection.worksheets[0];
  assert.deepEqual(inspection.sheets, ["Kas"]);
  assert.ok(sheet.formulas.some((entry) => entry.includes("A4: SUM(B2:B3)")));
  assert.ok(sheet.merges.includes("A5:B5"));
  assert.ok(sheet.widths.includes("C:22"));
  assert.equal(sheet.design.shapes, 1);
  const byAddress = new Map(sheet.detail.map((entry) => [entry.cell, entry]));
  assert.equal(byAddress.get("A4").formula, "SUM(B2:B3)");
  assert.equal(byAddress.get("A4").cell, "A4");
});

test("editing a multi-month report keeps every chart part declared", async (t) => {
  const { office, workspace, store } = withOffice(t);
  const sync = new FinanceWorkbookSync({ financeStore: store, fileWorkspace: workspace, defaultCurrency: "IDR" });
  // Three months -> three native charts across two drawings. ExcelJS drops all of
  // them on rewrite, and restoring only the first one per kind left chart2..3
  // without a content type, which made Excel report the file as corrupt.
  for (const month of ["07", "08", "09"]) {
    store.recordTransaction(owner, { kind: "income", amount: 9_000_000, category: "Gaji", description: "Gaji", date: `2026-${month}-01` });
    store.recordTransaction(owner, { kind: "expense", amount: 3_000_000, category: "Sewa", description: "Sewa", date: `2026-${month}-02` });
  }
  const report = await sync.renderReportFile(owner, {});
  const before = readZip(await readFile(await workspace.getFilePath(owner, report.file.id)));
  const chartPartsBefore = [...before.keys()].filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name));
  assert.ok(chartPartsBefore.length >= 3, "fixture must contain several native charts");

  await office.updateSpreadsheetCells(owner, {
    fileId: report.file.id,
    sheetName: "Ringkasan",
    saveMode: "in_place",
    updates: [{ cell: "A26", value: "UANG KOTOR", format: { bold: true, fill: "0B3B33" } }],
  });

  const after = readZip(await readFile(await workspace.getFilePath(owner, report.file.id)));
  const contentTypes = after.get("[Content_Types].xml").toString("utf8");
  const declared = new Set([...contentTypes.matchAll(/PartName="\/([^"]+)"/g)].map((match) => match[1]));
  const drawingParts = [...after.keys()].filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name));
  assert.equal(drawingParts.length, 2);
  for (const part of [...chartPartsBefore, ...drawingParts]) {
    assert.ok(declared.has(part), `${part} has no content type override`);
  }
});

test("a designed workbook still reloads after a later cell edit", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 10, 4]] });
  const designed = await office.designSpreadsheet(owner, {
    fileId: file.id, sheetName: "Kas", saveMode: "in_place",
    shapes: [{ from: "E2", to: "H5", text: "Panel", fill: "0F766E" }],
    charts: [{
      type: "bar", from: "E7", to: "K20", categories: "A2:A2",
      series: [{ name: "Masuk", range: "B2:B2" }],
    }],
  });
  await office.updateSpreadsheetCells(owner, {
    fileId: designed.id, sheetName: "Kas", saveMode: "in_place",
    updates: [{ cell: "B2", value: 99, format: { numFmt: "currency" } }],
  });
  const workbook = await load(workspace, owner, designed.id);
  assert.equal(workbook.getWorksheet("Kas").getCell("B2").value, 99);
  const zip = readZip(await readFile(await workspace.getFilePath(owner, designed.id)));
  assert.match(zip.get("xl/drawings/drawing1.xml").toString("utf8"), /<xdr:sp\b/);
  assert.equal([...zip.keys()].filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name)).length, 1);
});

test("editing actions fall back to the first sheet when sheet_name is omitted", async (t) => {
  const { office, workspace } = withOffice(t);
  const file = await seed(office, { rows: [["Jan", 10, 4], ["Feb", 7, 2]] });

  // update_cells without sheet_name must not throw "Cannot read properties of
  // undefined" — it resolves to the workbook's first sheet ("Kas").
  const updated = await office.updateSpreadsheetCells(owner, {
    fileId: file.id, saveMode: "in_place",
    updates: [{ cell: "B3", value: 8, format: { bold: true } }],
  });
  const workbook = await load(workspace, owner, updated.id);
  assert.equal(workbook.getWorksheet("Kas").getCell("B3").value, 8);
  assert.ok(updated.verified.some((entry) => entry.cell === "B3"));

  const appended = await office.appendSpreadsheetRows(owner, {
    fileId: file.id, saveMode: "in_place", rows: [["Mar", 9, 3]],
  });
  const afterAppend = await load(workspace, owner, appended.id);
  assert.equal(afterAppend.getWorksheet("Kas").getCell("A4").value, "Mar");

  await office.insertSpreadsheetRows(owner, { fileId: file.id, at: 2, count: 1, saveMode: "in_place" });
  await office.deleteSpreadsheetColumns(owner, { fileId: file.id, at: 2, count: 1, saveMode: "in_place" });
  const final = await load(workspace, owner, file.id);
  assert.equal(final.getWorksheet("Kas").getCell("B1").value, "Keluar");

  await office.designSpreadsheet(owner, {
    fileId: file.id, saveMode: "in_place",
    shapes: [{ from: "E2", to: "H4", text: "Panel", fill: "0F766E" }],
  });
  const designed = await load(workspace, owner, file.id);
  assert.ok(designed.getWorksheet("Kas"));
  const zip = readZip(await readFile(await workspace.getFilePath(owner, file.id)));
  assert.match(zip.get("xl/drawings/drawing1.xml").toString("utf8"), /<xdr:sp\b/);
});

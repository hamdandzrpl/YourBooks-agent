import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import { FileWorkspace } from "../src/files/workspace.js";
import { OfficeFiles } from "../src/files/office.js";
import { ShareLinkStore } from "../src/share-links.js";
import { runFileTool } from "../src/files/tools.js";
import { buildShareDataset } from "../src/share-dataset.js";

function context(t) {
  const directory = mkdtempSync(join(tmpdir(), "share-tool-test-"));
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

async function readDataset(fileWorkspace, userId, folder) {
  const files = await fileWorkspace.listFiles(userId, folder);
  const dataset = files.find((file) => file.path === `${folder}/dataset.json`);
  assert.ok(dataset, "dataset.json must be written");
  return JSON.parse(await fileWorkspace.readTextFile(userId, dataset.id));
}

test("share_link_set_rows writes a multi-sheet dataset with a readable summary", async (t) => {
  const ctx = context(t);
  const created = await runFileTool("workspace_files", {
    action: "share_link_create", folder: "documents/kebun", name: "kebun", title: "Kebun",
  }, ctx);
  assert.match(created.url, /\/w\/kebun\?k=/);

  const result = await runFileTool("workspace_files", {
    action: "share_link_set_rows",
    folder: "documents/kebun",
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], headerRow: true, rows: [
        ["Total Pemasukan", { formula: "SUM(Pemasukan!B:B)" }],
        ["Total Pengeluaran", { formula: "SUM(Pengeluaran!B:B)" }],
        ["Laba Bersih", { formula: "B2-B3" }],
      ] },
      { name: "Pemasukan", columns: ["Tanggal", "Total"], rows: [["2024-01-15", 7500000]] },
      { name: "Pengeluaran", columns: ["Tanggal", "Total"], rows: [["2024-01-05", 2500000]] },
    ],
    formats: [{ sheet: "Dashboard", range: "B2:B4", align: "right" }],
  }, ctx);

  assert.equal(result.saved, true);
  assert.equal(result.sheet_count, 3);
  assert.deepEqual(result.sheets.map((sheet) => sheet.name), ["Dashboard", "Pemasukan", "Pengeluaran"]);
  // The agent gets the summary back so it can verify the sheet it just wrote.
  assert.ok(result.summary, "tool result must include the summary block");
  assert.equal(result.summary.sheet, "Dashboard");
  // KPI labels and value types, but not the exact last decimal: the formula
  // engine and JavaScript compute 5000000/7500000 with a different epsilon.
  assert.equal(result.summary.kpis.length, 4);
  const byLabel = Object.fromEntries(result.summary.kpis.map((kpi) => [kpi.label, kpi]));
  assert.deepEqual(byLabel["Total Pemasukan"].value, 7500000);
  assert.deepEqual(byLabel["Total Pengeluaran"].value, 2500000);
  assert.deepEqual(byLabel["Laba Bersih"].value, 5000000);
  assert.ok(Math.abs(byLabel["Margin"].value - 5000000 / 7500000) < 0.0001);
  assert.deepEqual(result.summary.missing.filter((item) => item.includes("KPI")), []);

  const stored = await readDataset(ctx.fileWorkspace, ctx.userId, "documents/kebun");
  assert.equal(stored.sheets.length, 3);
  assert.ok(stored.summary, "the stored dataset carries the summary block");
  // Formula results are cached, so the page shows numbers on first paint.
  assert.equal(stored.sheets[0].rows[1][1].value, 7500000);
  assert.equal(stored.sheets[0].styles["1,1"].align, "right");
});

test("the tool reports a summary that cannot be read instead of hiding it", async (t) => {
  const ctx = context(t);
  await runFileTool("workspace_files", {
    action: "share_link_create", folder: "documents/rapi", name: "rapi", title: "Rapi",
  }, ctx);
  const result = await runFileTool("workspace_files", {
    action: "share_link_set_rows",
    folder: "documents/rapi",
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], rows: [["Total Pemasukan", "belum diisi"]] },
      { name: "Pemasukan", columns: ["Tanggal"], rows: [["2024-01-01"]] },
    ],
  }, ctx);
  assert.ok(result.summary, "a summary block is still reported");
  assert.ok(result.summary.missing.length > 0, "the agent must be told what is missing");
});

test("share_link_set_rows still accepts the single-sheet form", async (t) => {
  const ctx = context(t);
  await runFileTool("workspace_files", {
    action: "share_link_create", folder: "documents/kas", name: "kas", title: "Kas",
  }, ctx);
  const result = await runFileTool("workspace_files", {
    action: "share_link_set_rows",
    folder: "documents/kas",
    columns: ["Tanggal", "Nominal"],
    rows: [["2024-01-01", 5000]],
  }, ctx);
  assert.equal(result.sheet_count, 1);
  assert.deepEqual(result.sheets[0].name, "Sheet1");
  const stored = await readDataset(ctx.fileWorkspace, ctx.userId, "documents/kas");
  assert.deepEqual(stored.sheets[0].rows[0], ["Tanggal", "Nominal"]);
});

test("share_link_from_file imports every sheet of a workbook and derives its summary", async (t) => {
  const ctx = context(t);
  // Build a small workbook that mirrors the mandated layout.
  const workbook = new ExcelJS.Workbook();
  const dashboard = workbook.addWorksheet("Dashboard");
  dashboard.addRow(["Komponen", "Nilai"]);
  dashboard.addRow(["Total Pemasukan", { formula: "SUM(Pemasukan!B:B)", result: 7500000 }]);
  dashboard.addRow(["Total Pengeluaran", { formula: "SUM(Pengeluaran!B:B)", result: 2500000 }]);
  dashboard.addRow(["Laba Bersih", { formula: "B2-B3", result: 5000000 }]);
  const income = workbook.addWorksheet("Pemasukan");
  income.addRow(["Tanggal", "Total"]);
  income.addRow(["2024-01-15", 7500000]);
  const spending = workbook.addWorksheet("Pengeluaran");
  spending.addRow(["Tanggal", "Total"]);
  spending.addRow(["2024-01-05", 2500000]);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());

  const saved = await ctx.fileWorkspace.saveAttachment(ctx.userId, {
    filename: "kebun.xlsx",
    bytes,
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    sourceMessageId: "wamid.share.fromfile",
  });

  const result = await runFileTool("workspace_files", {
    action: "share_link_from_file",
    folder: "documents/kebun-xlsx",
    file_id: saved.id,
    name: "kebun-xlsx",
    title: "Kebun",
  }, ctx);
  assert.match(result.url, /\/w\/kebun-xlsx\?k=/);

  const stored = await readDataset(ctx.fileWorkspace, ctx.userId, "documents/kebun-xlsx");
  // All three sheets travel, not just the first one.
  assert.deepEqual(stored.sheets.map((sheet) => sheet.name), ["Dashboard", "Pemasukan", "Pengeluaran"]);
  assert.ok(stored.summary, "the imported workbook carries a summary block");
  assert.equal(stored.summary.sheet, "Dashboard");
  assert.deepEqual(stored.summary.kpis.map((kpi) => [kpi.label, kpi.value]), [
    ["Total Pemasukan", 7500000],
    ["Total Pengeluaran", 2500000],
    ["Laba Bersih", 5000000],
  ]);
});

test("an explicit sheet_name keeps the single-sheet import", async (t) => {
  const ctx = context(t);
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet("Dashboard").addRow(["Total Pemasukan", 100]);
  const data = workbook.addWorksheet("Pemasukan");
  data.addRow(["Tanggal", "Total"]);
  data.addRow(["2024-01-15", 100]);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const saved = await ctx.fileWorkspace.saveAttachment(ctx.userId, {
    filename: "dua.xlsx", bytes,
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    sourceMessageId: "wamid.share.sheetname",
  });
  await runFileTool("workspace_files", {
    action: "share_link_from_file",
    folder: "documents/satu-sheet",
    file_id: saved.id,
    sheet_name: "Pemasukan",
    name: "satu-sheet",
  }, ctx);
  const stored = await readDataset(ctx.fileWorkspace, ctx.userId, "documents/satu-sheet");
  assert.equal(stored.sheets.length, 1);
  assert.deepEqual(stored.sheets[0].rows[0], ["Tanggal", "Total"]);
});

test("a rewrite that breaks a previously readable summary is reported", async (t) => {
  const ctx = context(t);
  await runFileTool("workspace_files", {
    action: "share_link_create", folder: "documents/jaga", name: "jaga", title: "Jaga",
  }, ctx);
  const good = await runFileTool("workspace_files", {
    action: "share_link_set_rows",
    folder: "documents/jaga",
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], rows: [["Total Pemasukan", 100], ["Total Pengeluaran", 40]] },
      { name: "Data", columns: ["Nilai"], rows: [[100]] },
    ],
  }, ctx);
  assert.equal(good.warning, undefined);
  assert.deepEqual(good.summary.missing.filter((item) => item.includes("KPI")), []);

  // Same document, but the KPI values are gone: the Workspace view would show
  // nothing, so the tool must say so instead of returning a clean save.
  const broken = await runFileTool("workspace_files", {
    action: "share_link_set_rows",
    folder: "documents/jaga",
    replace: true,
    sheets: [{ name: "Data", columns: ["Nilai"], rows: [[100]] }],
  }, ctx);
  assert.equal(broken.saved, true);
  assert.match(String(broken.warning), /tadinya terbaca/);
  assert.match(String(broken.warning), /tidak ada lagi|hilang/);
  // With no summary sheet left there is no block to report.
  assert.equal(broken.summary, null);
});

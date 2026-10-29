import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import test from "node:test";
import { importWorkspaceXlsx, exportWorkspaceXlsx } from "../src/workspace-sheets.js";
import { computeWorkbook } from "../src/formula-engine.js";
import { extractSummary, detectSummarySheet, toNumber, withSummary, summaryBrief } from "../src/workspace-summary.js";
import { buildShareDataset } from "../src/share-dataset.js";
import { loadSystemSkills } from "../src/skills.js";

// The real workbook lives outside the repo. On a deploy machine it is absent,
// so the tests that need it skip instead of reporting a false failure — the
// extraction contract itself is still covered by the synthetic cases below.
const examplePath = process.env.WORKSPACE_EXAMPLE_XLSX
  ?? "C:/Users/ACER/Downloads/Pembukuan Usaha Pertanian - Multi Komoditas.xlsx";
const hasExample = existsSync(examplePath);
const example = () => importWorkspaceXlsx(readFileSync(examplePath), "Pertanian");
const needsExample = hasExample ? {} : { skip: `example workbook not found at ${examplePath}` };

test("the real example workbook yields the four KPIs and the commodity table", needsExample, async () => {
  const summary = extractSummary(await example());
  assert.equal(summary.sheet, "Dashboard & Ringkasan Hasil");
  assert.match(summary.title, /DASHBOARD KINERJA PERTANIAN/i);
  assert.deepEqual(summary.kpis.map((kpi) => [kpi.label, kpi.value, kpi.cell]), [
    ["Total Pemasukan", 88990000, "B5"],
    ["Total Pengeluaran", 11855000, "D5"],
    ["Laba Bersih", 77135000, "F5"],
    ["Margin Keuntungan", 0.8667827845825373, "H5"],
  ]);
  assert.equal(summary.kpis[3].kind, "percent");
  assert.equal(summary.table.header.length, 6);
  assert.deepEqual(summary.table.header, ["Komoditas", "Volume Penjualan", "Total Pemasukan", "Total Biaya Terkait", "Laba Bersih", "Profit Margin"]);
  assert.equal(summary.table.rows.length, 6);
  assert.deepEqual(summary.table.rows[0], ["Kopi", 370, 29850000, 1475000, 28375000, 0.95058626466]);
  assert.deepEqual(summary.missing, []);
});

test("the KPI strip is not mistaken for the summary table", needsExample, async () => {
  const summary = extractSummary(await example());
  // The strip's labels sit on one row and its values on the next; the table is
  // the six-row commodity block further down.
  assert.equal(summary.table.start_row, 10);
  assert.equal(summary.table.end_row, 16);
  assert.ok(!summary.table.header.includes("Volume Penjualan") || summary.table.rows.length > 1);
});

test("an uploaded workbook keeps the numbers Excel already computed", needsExample, async () => {
  // The example uses SUMIFS with date criteria, which the web engine cannot
  // evaluate. Its stored results must survive, or the whole dashboard reads
  // #VALUE!.
  const dataset = await example();
  const errors = [];
  for (const sheet of dataset.sheets) {
    sheet.rows.forEach((row, r) => row.forEach((cell, c) => {
      const value = cell && typeof cell === "object" ? cell.value : cell;
      if (typeof value === "string" && value.startsWith("#")) errors.push(`${sheet.name}!${r + 1},${c + 1}`);
    }));
  }
  assert.deepEqual(errors, []);
  const dashboard = dataset.sheets[0];
  assert.equal(dashboard.rows[4][1].value, 88990000);
  assert.equal(dashboard.rows[4][3].value, 11855000);
});

test("a sheet the engine cannot evaluate keeps its cached value instead of #VALUE!", () => {
  const dataset = computeWorkbook({
    sheets: [{
      name: "Ringkasan",
      columns: ["A"],
      rows: [
        ["Total Pemasukan", 88990000],
        ["Nilai", { formula: "SUMIFS(X!A:A, X!B:B, \">=\"&A2)", value: 42 }],
      ],
    }],
  });
  assert.equal(dataset.sheets[0].rows[1][1].value, 42, "cached result wins over the engine's error");
});

test("the skill's own layout is readable", () => {
  const dataset = withSummary(buildShareDataset({
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], headerRow: true, rows: [
        ["Total Pemasukan", 14200000],
        ["Total Pengeluaran", 6250000],
        ["Laba Bersih", 7950000],
        ["Margin", 0.5598],
      ] },
      { name: "Pemasukan", columns: ["Tanggal", "Total"], rows: [["2024-01-15", 7500000]] },
    ],
  }));
  const summary = dataset.summary;
  assert.equal(summary.sheet, "Dashboard");
  assert.deepEqual(summary.kpis.map((kpi) => [kpi.label, kpi.value]), [
    ["Total Pemasukan", 14200000],
    ["Total Pengeluaran", 6250000],
    ["Laba Bersih", 7950000],
    ["Margin", 0.5598],
  ]);
  assert.equal(summary.missing.filter((item) => item.includes("KPI")).length, 0);
});

test("a missing KPI pair is reported instead of silently dropped", () => {
  const summary = extractSummary({ sheets: [{ name: "Dashboard", rows: [["Total Pemasukan", "belum diisi"]] }] });
  assert.deepEqual(summary.kpis, []);
  assert.match(summary.missing.join(" "), /no KPI label paired with a numeric value/);
  // No input data anywhere, so the table nag stays quiet (see the empty-workbook test).
  assert.ok(!summary.missing.some((item) => item.startsWith("no summary table")));
});

test("an expense total and a table are required, and named when absent", () => {
  const summary = extractSummary({
    sheets: [{ name: "Ringkasan", rows: [
      ["Total Pemasukan", 1000],
      ["", ""],
      ["Komponen", "Volume", "Total"],
      ["Kopi", 10, 500],
      ["Pisang", 20, 500],
    ] }],
  });
  assert.equal(summary.kpis.length, 1);
  assert.equal(summary.table.rows.length, 2);
  assert.match(summary.missing.join(" "), /no expense total/);
});

test("an empty workbook is a valid empty state: KPIs read 0, no table nagging", async () => {
  // The shape the skill mandates for a fresh document: header-only input
  // sheets and a KPI block of cross-sheet formulas that sum empty columns.
  const dataset = withSummary(computeWorkbook(buildShareDataset({
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], headerRow: true, rows: [
        ["Total Pemasukan", { formula: "SUM(Pemasukan!G:G)" }],
        ["Total Pengeluaran", { formula: "SUM(Pengeluaran!G:G)" }],
        ["Laba Bersih", { formula: "B2-B3" }],
        ["Margin", { formula: "IF(B2>0,B4/B2,0)" }],
      ] },
      { name: "Pemasukan", columns: ["Tanggal", "Bulan", "Komoditas", "Kuantitas", "Satuan", "Harga Satuan", "Total"], headerRow: true, rows: [] },
      { name: "Pengeluaran", columns: ["Tanggal", "Bulan", "Kategori Biaya", "Uraian", "Komoditas", "Kuantitas", "Total"], headerRow: true, rows: [] },
    ],
  })));
  assert.deepEqual(dataset.summary.kpis.map((kpi) => kpi.value), [0, 0, 0, 0]);
  assert.equal(dataset.summary.table, null);
  assert.ok(!dataset.summary.missing.some((item) => item.startsWith("no summary table")),
    "an empty workbook must not be nagged about the missing table");
  assert.equal(dataset.summary.missing.length, 0);
});

test("once data exists, a missing summary table is reported again", () => {
  const summary = extractSummary({
    sheets: [{ name: "Dashboard", rows: [
      ["Total Pemasukan", 1000],
      ["Total Pengeluaran", 200],
      ["Laba Bersih", 800],
      ["Margin", 0.8],
      ["", ""],
      ["Komponen", "Volume", "Total Pemasukan", "Total Biaya", "Laba Bersih", "Margin"],
      ["Kopi", 10, 700, 100, 600, 0.857],
      ["Pisang", 20, 300, 100, 200, 0.667],
    ] }],
  });
  assert.equal(summary.table.rows.length, 2);
});

test("sheet detection prefers a summary tab name over sheet order", () => {
  const sheets = [
    { name: "Pemasukan", rows: [["Tanggal", "Total"], ["2024-01-01", 5]] },
    { name: "Ringkasan", rows: [["Total Pemasukan", 5]] },
  ];
  assert.equal(detectSummarySheet(sheets).name, "Ringkasan");
  // With no named summary, a sheet carrying KPI labels is used instead.
  assert.equal(detectSummarySheet([
    { name: "Data", rows: [["Tanggal", "Total"]] },
    { name: "Laporan", rows: [["Total Pemasukan", 5], ["Total Pengeluaran", 2]] },
  ]).name, "Laporan");
});

test("numbers typed by a human are parsed", () => {
  assert.equal(toNumber(1000), 1000);
  assert.equal(toNumber("Rp 88.990.000"), 88990000);
  assert.equal(toNumber("1.234,56"), 1234.56);
  assert.equal(toNumber("86,7%"), 86.7);
  assert.equal(toNumber("belum ada"), null);
  assert.equal(toNumber(null), null);
});

test("the summary block never leaks into the exported spreadsheet", async () => {
  const dataset = withSummary(buildShareDataset({
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], rows: [["Total Pemasukan", 100]] },
      { name: "Pemasukan", columns: ["Tanggal"], rows: [["2024-01-01"]] },
    ],
  }));
  assert.ok(dataset.summary, "the dataset carries a summary");
  const bytes = await exportWorkspaceXlsx(dataset);
  const reopened = await importWorkspaceXlsx(bytes, "Reimport");
  assert.deepEqual(reopened.sheets.map((sheet) => sheet.name), ["Dashboard", "Pemasukan"]);
  // Two sheets only: no extra tab was created for the summary.
  assert.equal(reopened.sheets.length, 2);
});

test("summaryBrief gives the agent a compact verification payload", needsExample, async () => {
  const brief = summaryBrief(extractSummary(await example()));
  assert.equal(brief.sheet, "Dashboard & Ringkasan Hasil");
  assert.equal(brief.kpis.length, 4);
  assert.deepEqual(brief.table.columns.slice(0, 2), ["Komoditas", "Volume Penjualan"]);
  assert.equal(brief.table.rows, 6);
  assert.deepEqual(brief.missing, []);
});

test("the skill teaches server-owned dashboard and no-dummy-data rule", () => {
  const skill = loadSystemSkills().find((entry) => entry.id === "dashboard-workbooks");
  const text = skill.instructions.join(" ");
  assert.match(text, /DASHBOARD DIBANGUN SERVER/i);
  assert.match(text, /TIDAK BOLEH ISI DATA DUMMY/i);
  assert.match(text, /share_link_append_rows/i);
  assert.match(text, /summary\.missing/);
});

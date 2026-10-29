import assert from "node:assert/strict";
import test from "node:test";
import { loadSystemSkills, formatSystemSkills } from "../src/skills.js";
import { buildShareDataset } from "../src/share-dataset.js";
import { computeWorkbook } from "../src/formula-engine.js";
import { exportWorkspaceXlsx, importWorkspaceXlsx } from "../src/workspace-sheets.js";

/** The structure the skill mandates: input sheets separate from the summary. */
function requestedWorkbook() {
  return {
    title: "Pembukuan Usaha Pertanian",
    sheets: [
      {
        name: "Dashboard",
        columns: ["Komponen", "Nilai"],
        headerRow: true,
        rows: [
          ["Total Pemasukan", { formula: "SUM(Pemasukan!G:G)" }],
          ["Total Pengeluaran", { formula: "SUM(Pengeluaran!G:G)" }],
          ["Laba Bersih", { formula: "B2-B3" }],
          ["Margin", { formula: "IF(B2>0, B4/B2, 0)" }],
        ],
      },
      {
        name: "Pemasukan",
        columns: ["Tanggal", "Bulan", "Komoditas", "Kuantitas", "Satuan", "Harga Satuan", "Total"],
        rows: [
          ["2024-01-15", "2024-01", "Kopi", 100, "kg", 75000, { formula: "D2*F2" }],
          ["2024-02-10", "2024-02", "Pisang", 10, "tandan", 400000, { formula: "D3*F3" }],
          ["2024-01-25", "2024-01", "Cabai", 60, "kg", 45000, { formula: "D4*F4" }],
        ],
      },
      {
        name: "Pengeluaran",
        columns: ["Tanggal", "Bulan", "Kategori Biaya", "Kuantitas", "Satuan", "Biaya Satuan", "Total"],
        rows: [
          ["2024-01-05", "2024-01", "Pupuk", 5, "karung", 850000, { formula: "D2*F2" }],
          ["2024-02-05", "2024-02", "Tenaga Kerja", 1, "bulan", 2000000, { formula: "D3*F3" }],
        ],
      },
      {
        name: "Laba Rugi Bulanan",
        columns: ["Periode", "Pemasukan", "Pengeluaran", "Laba Bersih", "Kumulatif"],
        rows: [
          ["2024-01",
            { formula: "SUMIFS(Pemasukan!G:G, Pemasukan!B:B, A2)" },
            { formula: "SUMIFS(Pengeluaran!G:G, Pengeluaran!B:B, A2)" },
            { formula: "B2-C2" }, { formula: "D2" }],
          ["2024-02",
            { formula: "SUMIFS(Pemasukan!G:G, Pemasukan!B:B, A3)" },
            { formula: "SUMIFS(Pengeluaran!G:G, Pengeluaran!B:B, A3)" },
            { formula: "B3-C3" }, { formula: "E2+D3" }],
        ],
      },
    ],
    formats: [
      { sheet: "Dashboard", range: "A1:B1", align: "center", bold: true, fill: "0F766E", color: "FFFFFF" },
      { sheet: "Dashboard", range: "B2:B5", align: "right", format: "#,##0" },
      { sheet: "Pemasukan", range: "D:D", align: "right" },
      { sheet: "Pemasukan", range: "F:G", align: "right" },
    ],
  };
}

const value = (cell) => (cell && typeof cell === "object" ? cell.value : cell);

test("the structure skill is installed, enabled, and teaches server-owned dashboards", () => {
  const skill = loadSystemSkills().find((entry) => entry.id === "dashboard-workbooks");
  assert.ok(skill, "dashboard-workbooks skill must exist");
  assert.equal(skill.enabled, true);
  const text = skill.instructions.join(" ");
  assert.match(text, /DASHBOARD DIBANGUN SERVER/i, "The skill must explain the server builds dashboards");
  assert.match(text, /STRUKTUR INPUT.*wajib/i, "The skill must cover mandatory input structure");
  assert.match(text, /JANGAN BUAT WORKSPACE BARU/i, "The skill must forbid creating a new workspace for recording");
  assert.match(text, /share_link_append_rows/i, "The skill must direct to append_rows for transactions");
  // It must reach the model's prompt.
  assert.match(formatSystemSkills(loadSystemSkills()), /Standar Tugas Buku Keuangan Web/);
});

test("the mandated structure builds four sheets in the documented order", () => {
  const dataset = buildShareDataset(requestedWorkbook());
  assert.deepEqual(
    dataset.sheets.map((sheet) => sheet.name),
    ["Dashboard", "Pemasukan", "Pengeluaran", "Laba Rugi Bulanan"],
  );
  // Input sheets carry a real header row, so the user can type under it.
  assert.deepEqual(dataset.sheets[1].rows[0], ["Tanggal", "Bulan", "Komoditas", "Kuantitas", "Satuan", "Harga Satuan", "Total"]);
  assert.equal(dataset.sheets[1].headerRow, true);
});

test("input totals and cross-sheet KPI values compute without errors", () => {
  const computed = computeWorkbook(buildShareDataset(requestedWorkbook()));
  const dashboard = computed.sheets[0].rows;
  assert.equal(value(dashboard[1][1]), 14200000, "total pemasukan = 7,5jt + 4jt + 2,7jt");
  assert.equal(value(dashboard[2][1]), 6250000, "total pengeluaran = 4,25jt + 2jt");
  assert.equal(value(dashboard[3][1]), 7950000, "laba bersih");
  assert.ok(Math.abs(value(dashboard[4][1]) - 7950000 / 14200000) < 1e-9, "margin");

  const inputTotals = computed.sheets[1].rows.slice(1).map((row) => value(row[6]));
  assert.deepEqual(inputTotals, [7500000, 4000000, 2700000], "each row multiplies kuantitas x harga");

  const monthly = computed.sheets[3].rows.slice(1).map((row) => row.map(value));
  assert.deepEqual(monthly, [
    ["2024-01", 10200000, 4250000, 5950000, 5950000],
    ["2024-02", 4000000, 2000000, 2000000, 7950000],
  ], "monthly SUMIFS groups by the Bulan helper column");
});

test("no cell in the built workbook resolves to an error value", () => {
  const computed = computeWorkbook(buildShareDataset(requestedWorkbook()));
  const bad = [];
  for (const sheet of computed.sheets) {
    sheet.rows.forEach((row, r) => row.forEach((cell, c) => {
      const result = value(cell);
      if (typeof result === "string" && result.startsWith("#")) bad.push(`${sheet.name}!${r + 1},${c + 1}=${result}`);
    }));
  }
  assert.deepEqual(bad, [], "a formula the web engine cannot evaluate would show #VALUE! to the user");
});

test("formats reach the header and the numeric columns", () => {
  const dataset = buildShareDataset(requestedWorkbook());
  const dashboard = dataset.sheets[0].styles;
  assert.equal(dashboard["0,0"].align, "center");
  assert.equal(dashboard["0,0"].bold, true);
  assert.equal(dashboard["0,0"].fill, "#0F766E");
  assert.equal(dashboard["1,1"].align, "right");
  assert.equal(dashboard["1,1"].format, "#,##0");
  // Whole-column format on the input sheet stops at the last written row.
  const pemasukan = dataset.sheets[1].styles;
  assert.equal(pemasukan["1,3"].align, "right");
  assert.equal(pemasukan["3,3"].align, "right");
  assert.equal(pemasukan["4,3"], undefined);
});

test("the exported workbook keeps every sheet, formula, and cached value", async () => {
  const computed = computeWorkbook(buildShareDataset(requestedWorkbook()));
  const bytes = await exportWorkspaceXlsx(computed);
  const reopened = await importWorkspaceXlsx(bytes, "Reimport");
  assert.deepEqual(
    reopened.sheets.map((sheet) => sheet.name),
    ["Dashboard", "Pemasukan", "Pengeluaran", "Laba Rugi Bulanan"],
  );
  const dashboard = reopened.sheets[0].rows;
  assert.equal(value(dashboard[1][1]), 14200000);
  assert.equal(value(dashboard[3][1]), 7950000);
  // Formulas survive the round trip, so Excel recalculates from source.
  assert.match(String(dashboard[1][1].formula), /SUM\(Pemasukan!G:G\)/);
});

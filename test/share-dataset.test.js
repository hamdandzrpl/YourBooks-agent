import assert from "node:assert/strict";
import test from "node:test";
import { buildShareDataset, describeDataset, parseCellRange, columnIndex } from "../src/share-dataset.js";

test("multi-sheet request keeps every sheet, its order, and its header", () => {
  const dataset = buildShareDataset({
    title: "Pembukuan Pertanian",
    sheets: [
      { name: "Dashboard", columns: ["Total Pemasukan", "Total Pengeluaran"], rows: [[88990000, 11855000]], headerRow: true },
      { name: "Pemasukan", columns: ["Tanggal", "Komoditas", "Total"], rows: [["2024-01-15", "Kopi", 7500000]] },
    ],
  });
  assert.equal(dataset.title, "Pembukuan Pertanian");
  assert.deepEqual(dataset.sheets.map((sheet) => sheet.name), ["Dashboard", "Pemasukan"]);
  // Named columns are materialized as a real first row, so the grid and the
  // XLSX export show the same header.
  assert.deepEqual(dataset.sheets[1].rows[0], ["Tanggal", "Komoditas", "Total"]);
  assert.deepEqual(dataset.sheets[1].rows[1], ["2024-01-15", "Kopi", 7500000]);
  assert.equal(dataset.sheets[0].headerRow, true);
});

test("single-sheet request still works with columns and rows", () => {
  const dataset = buildShareDataset({ folder: "documents/kas", columns: ["Tanggal", "Nominal"], rows: [["2024-01-01", 5000]] });
  assert.equal(dataset.sheets.length, 1);
  assert.equal(dataset.sheets[0].name, "Sheet1");
  assert.deepEqual(dataset.sheets[0].rows[0], ["Tanggal", "Nominal"]);
});

test("numbers, booleans and formulas survive instead of being stringified", () => {
  const dataset = buildShareDataset({
    sheets: [{
      name: "Data",
      columns: ["Nama", "Nilai", "Pajak", "Aktif"],
      rows: [["Kopi", 1000000, { formula: "B2*0.11" }, true]],
    }],
  });
  const row = dataset.sheets[0].rows[1];
  assert.equal(row[1], 1000000);
  assert.equal(row[3], true);
  assert.deepEqual(row[2], { formula: "B2*0.11" });
});

test("a summary sheet may compute from a data sheet", () => {
  const dataset = buildShareDataset({
    sheets: [
      { name: "Ringkasan", columns: ["Total"], rows: [[{ formula: "SUM(Data!B2:B3)" }]] },
      { name: "Data", columns: ["Nilai"], rows: [[100], [250]] },
    ],
  });
  assert.deepEqual(dataset.sheets[0].rows[1][0], { formula: "SUM(Data!B2:B3)" });
});

test("formats set alignment on the requested range only", () => {
  const dataset = buildShareDataset({
    sheets: [{ name: "Ringkasan", columns: ["Label", "Nilai"], rows: [["Pemasukan", 100], ["Pengeluaran", 40]] }],
    formats: [
      { range: "B1:B3", align: "right" },
      { range: "A1:B1", align: "center", bold: true, fill: "0F766E", color: "FFFFFF" },
    ],
  });
  const styles = dataset.sheets[0].styles;
  // Header row: center + bold + fill from the second entry.
  assert.equal(styles["0,0"].align, "center");
  assert.equal(styles["0,0"].bold, true);
  assert.equal(styles["0,0"].fill, "#0F766E");
  // Body of column B: right aligned, and the second entry overwrote B1.
  assert.equal(styles["1,1"].align, "right");
  assert.equal(styles["2,1"].align, "right");
  assert.equal(styles["0,1"].align, "center");
  // Column A below the header is untouched.
  assert.equal(styles["1,0"], undefined);
});

test("formats can target a named sheet", () => {
  const dataset = buildShareDataset({
    sheets: [
      { name: "Dashboard", columns: ["Total"], rows: [[10]] },
      { name: "Data", columns: ["Nilai"], rows: [[10]] },
    ],
    formats: [{ sheet: "Data", range: "A1", align: "center" }],
  });
  assert.equal(dataset.sheets[0].styles, undefined);
  assert.equal(dataset.sheets[1].styles["0,0"].align, "center");
});

test("whole-column format stops at the last written row", () => {
  const dataset = buildShareDataset({
    sheets: [{ name: "Data", columns: ["Nilai"], rows: [[1], [2], [3]] }],
    formats: [{ range: "A:A", align: "right" }],
  });
  const keys = Object.keys(dataset.sheets[0].styles).sort();
  assert.deepEqual(keys, ["0,0", "1,0", "2,0", "3,0"]);
});

test("border shorthand becomes four drawn edges", () => {
  const dataset = buildShareDataset({
    sheets: [{ name: "Data", columns: ["A"], rows: [[1]] }],
    formats: [{ range: "A1:A2", border: "thin", border_color: "D4D4D8" }],
  });
  const style = dataset.sheets[0].styles["1,0"];
  assert.deepEqual(Object.keys(style.border).sort(), ["bottom", "left", "right", "top"]);
  assert.equal(style.border.top.style, "thin");
  assert.equal(style.border.top.color, "#D4D4D8");
});

test("unknown format keys are dropped, never persisted", () => {
  const dataset = buildShareDataset({
    sheets: [{ name: "Data", columns: ["A"], rows: [[1]] }],
    formats: [{ range: "A1", align: "right", script: "evil()", fill: "not-a-colour" }],
  });
  const style = dataset.sheets[0].styles["0,0"];
  assert.deepEqual(Object.keys(style), ["align"]);
});

test("an empty request is rejected instead of writing a blank document", () => {
  assert.throws(() => buildShareDataset({ folder: "documents/kosong" }), /Provide sheets/);
});

test("a formula keeps its bare form and its cached value", () => {
  const dataset = buildShareDataset({
    sheets: [{ name: "X", columns: ["Nilai"], rows: [[{ formula: "=B2*2", value: 42 }]] }],
  });
  assert.deepEqual(dataset.sheets[0].rows[1][0], { formula: "B2*2", value: 42 });
});

test("duplicate sheet names are rejected", () => {
  assert.throws(
    () => buildShareDataset({ sheets: [{ name: "Data", rows: [[1]] }, { name: "data", rows: [[2]] }] }),
    /unique/,
  );
});

test("describeDataset reports what was written", () => {
  const dataset = buildShareDataset({
    sheets: [{ name: "Dashboard", columns: ["Total"], rows: [[1]], headerRow: true }],
    formats: [{ range: "A1", align: "center" }],
  });
  assert.deepEqual(describeDataset(dataset), {
    title: "Workspace",
    sheet_count: 1,
    sheets: [{ name: "Dashboard", rows: 2, columns: 1, header_row: true, formatted_cells: 1 }],
  });
});

test("range parsing accepts a cell, a rectangle, and a whole column", () => {
  assert.equal(columnIndex("A"), 0);
  assert.equal(columnIndex("Z"), 25);
  assert.equal(columnIndex("AB"), 27);
  assert.equal(columnIndex("1"), -1);
  assert.deepEqual(parseCellRange("B7"), { start: { col: 1, row: 6 }, end: { col: 1, row: 6 } });
  assert.deepEqual(parseCellRange("C2:A1"), { start: { col: 0, row: 0 }, end: { col: 2, row: 1 } });
  assert.deepEqual(parseCellRange("B:B"), { start: { col: 1, row: 0 }, end: { col: 1, row: Infinity } });
  assert.equal(parseCellRange("bukan-range"), null);
});

import assert from "node:assert/strict";
import test from "node:test";
import { readFileSync } from "node:fs";
import vm from "node:vm";

/** Load the browser module into a sandbox and hand back its public API. */
function loadAnalysis() {
  const source = readFileSync(new URL("../public/workspace-analysis.js", import.meta.url), "utf8");
  const sandbox = { window: {}, Intl, Date, Math, Number, String, Object, Array, JSON, RegExp, console };
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  return sandbox.window.WorkspaceAnalysis;
}

const ledger = {
  title: "Pembukuan Warung",
  sheets: [{
    name: "Transaksi",
    columns: ["A", "B", "C", "D", "E"],
    rows: [
      ["Pembukuan Warung Bu Sri", "", "", "", ""],
      ["", "", "", "", ""],
      ["Tanggal", "Keterangan", "Kategori", "Pemasukan", "Pengeluaran"],
      ["2026-01-03", "Penjualan harian", "Penjualan", 1500000, 0],
      ["2026-01-05", "Beli bahan baku", "Bahan", 0, 620000],
      ["2026-02-02", "Penjualan harian", "Penjualan", 1750000, 0],
      ["2026-02-08", "Bayar listrik", "Operasional", 0, 340000],
      ["2026-03-01", "Penjualan katering", "Penjualan", 2400000, 0],
      ["2026-03-09", "Gaji karyawan", "Operasional", 0, 900000],
    ],
  }],
};

test("columns are detected from the header row, not the first row", () => {
  const api = loadAnalysis();
  const fields = api.detectFields(ledger);
  assert.equal(fields.sheet, 0);
  assert.equal(fields.headerRow, 2, "header sits on the third row");
  assert.equal(fields.date, 0);
  assert.equal(fields.note, 1);
  assert.equal(fields.category, 2);
  assert.equal(fields.income, 3);
  assert.equal(fields.expense, 4);
});

test("the ledger reads every data row and skips the title", () => {
  const api = loadAnalysis();
  const built = api.buildLedger(ledger, api.detectFields(ledger));
  assert.equal(built.records.length, 6);
  assert.equal(built.sheetName, "Transaksi");
  assert.equal(built.records[0].income, 1500000);
  assert.equal(built.records[0].net, 1500000);
  assert.equal(built.records[1].expense, 620000);
  assert.equal(built.records[1].net, -620000);
});

test("KPI totals match a hand calculation", () => {
  const api = loadAnalysis();
  const built = api.buildLedger(ledger, api.detectFields(ledger));
  const income = api.computeWidget({ type: "kpi", metric: "income", agg: "sum" }, built);
  const expense = api.computeWidget({ type: "kpi", metric: "expense", agg: "sum" }, built);
  const net = api.computeWidget({ type: "kpi", metric: "net", agg: "sum" }, built);

  assert.equal(income.value, 5650000);
  assert.equal(expense.value, 1860000);
  assert.equal(net.value, 3790000);
  assert.match(net.sub, /Margin 67,1%/);
});

test("category and month groupings are correct", () => {
  const api = loadAnalysis();
  const built = api.buildLedger(ledger, api.detectFields(ledger));

  const byCategory = api.computeWidget({ type: "bar", metric: "expense", agg: "sum", groupBy: "category", limit: 8 }, built);
  const map = {};
  for (const item of byCategory.items) map[item.label] = item.value;
  assert.equal(map.Operasional, 1240000);
  assert.equal(map.Bahan, 620000);
  assert.equal(byCategory.items[0].label, "Operasional", "sorted biggest first");

  const byMonth = api.computeWidget({ type: "line", metric: "net", agg: "sum", groupBy: "month" }, built);
  assert.deepEqual([...byMonth.points.map((p) => p.label)], ["Jan 2026", "Feb 2026", "Mar 2026"]);
  assert.deepEqual([...byMonth.points.map((p) => p.value)], [880000, 1410000, 1500000]);
});

test("the table widget ranks by absolute size", () => {
  const api = loadAnalysis();
  const built = api.buildLedger(ledger, api.detectFields(ledger));
  const table = api.computeWidget({ type: "table", limit: 3 }, built);
  assert.equal(table.rows.length, 3);
  assert.equal(table.rows[0].value, 2400000);
  assert.equal(table.rows[0].note, "Penjualan katering");
});

test("progress reports its ratio against the target", () => {
  const api = loadAnalysis();
  const built = api.buildLedger(ledger, api.detectFields(ledger));
  const progress = api.computeWidget({ type: "progress", metric: "income", agg: "sum", target: 5650000 }, built);
  assert.equal(progress.value, 5650000);
  assert.equal(progress.ratio, 1);
});

test("Indonesian numbers are read and written correctly", () => {
  const api = loadAnalysis();
  assert.equal(api.toNumber("1.500.000"), 1500000);
  assert.equal(api.toNumber("1.500.000,50"), 1500000.5);
  assert.equal(api.toNumber("Rp 2.400.000"), 2400000);
  assert.equal(api.toNumber(""), 0);
  assert.equal(api.toNumber("abc"), 0);
  assert.equal(api.toNumber(-250), -250);
  assert.equal(api.toNumber("1500"), 1500, "a plain number is not mistaken for thousands");
  assert.equal(api.formatCompact(5650000), "Rp 5,7 jt");
  assert.equal(api.formatCompact(-1500000), "-Rp 1,5 jt");
  assert.equal(api.formatNumber(3790000, "currency"), "Rp 3.790.000");
});

test("dates are read in the shapes bookkeeping files use", () => {
  const api = loadAnalysis();
  assert.equal(api.toDate("2026-01-03").toISOString().slice(0, 10), "2026-01-03");
  assert.equal(api.toDate("03/01/2026").toISOString().slice(0, 10), "2026-01-03");
  assert.equal(api.toDate("3 Januari 2026").toISOString().slice(0, 10), "2026-01-03");
  assert.equal(api.toDate("Jan 2026").toISOString().slice(0, 7), "2026-01");
  assert.equal(api.toDate(""), null);
  assert.equal(api.toDate("bukan tanggal"), null);
  assert.equal(api.toDate(46025).toISOString().slice(0, 10), "2026-01-03", "Excel serial date");
});

test("a single signed amount column is understood", () => {
  const api = loadAnalysis();
  const dataset = {
    sheets: [{
      name: "Kas",
      columns: ["A", "B"],
      rows: [
        ["Keterangan", "Nominal"],
        ["Penjualan", 500000],
        ["Beli gas", -120000],
        ["Penjualan", 300000],
      ],
    }],
  };
  const built = api.buildLedger(dataset, api.detectFields(dataset));
  const net = api.computeWidget({ type: "kpi", metric: "net", agg: "sum" }, built);
  assert.equal(net.value, 680000);
});

test("an empty workbook does not crash and says so", () => {
  const api = loadAnalysis();
  const empty = { sheets: [{ name: "Sheet1", columns: ["A", "B"], rows: [] }] };
  assert.equal(api.detectFields(empty), null);
  const built = api.buildLedger(empty, null);
  assert.equal(built.records.length, 0);
  assert.deepEqual(JSON.parse(JSON.stringify(api.computeWidget({ type: "kpi", metric: "net" }, built))), { empty: true });
});

test("switching sheet re-detects that sheet's own columns", () => {
  const api = loadAnalysis();
  const dataset = {
    sheets: [
      ledger.sheets[0],
      {
        name: "Ringkasan",
        columns: ["A", "B", "C"],
        rows: [
          ["Bulan", "Masuk", "Keluar"],
          ["Jan", 100, 40],
          ["Feb", 200, 60],
        ],
      },
    ],
  };
  const second = api.detectForSheet(dataset, 1);
  assert.equal(second.sheet, 1);
  assert.equal(second.headerRow, 0);
  assert.equal(second.income, 1);
  assert.equal(second.expense, 2);
});

test("editing a cell changes the dashboard numbers", () => {
  const api = loadAnalysis();
  const before = api.computeWidget({ type: "kpi", metric: "income", agg: "sum" }, api.buildLedger(ledger, api.detectFields(ledger)));
  const edited = JSON.parse(JSON.stringify(ledger));
  edited.sheets[0].rows[3][3] = 2000000;
  const after = api.computeWidget({ type: "kpi", metric: "income", agg: "sum" }, api.buildLedger(edited, api.detectFields(edited)));
  assert.equal(after.value - before.value, 500000);
});

test("widget rendering produces safe HTML for hostile cell text", () => {
  const api = loadAnalysis();
  const hostile = {
    sheets: [{
      name: "X",
      columns: ["A", "B", "C", "D", "E"],
      rows: [
        ["Tanggal", "Keterangan", "Kategori", "Pemasukan", "Pengeluaran"],
        ["2026-01-01", "<img src=x onerror=alert(1)>", "<script>bad()</script>", 1000, 0],
        ["2026-01-02", "normal", "Aman", 2000, 0],
        ["2026-01-03", "\"><svg onload=alert(2)>", "Aman", 0, 500],
      ],
    }],
  };
  const fields = api.detectFields(hostile);
  assert.equal(fields.confidence, "header");
  const built = api.buildLedger(hostile, fields);
  const grid = { innerHTML: "" };
  api.renderWidgets(grid, api.autoWidgets(), built);

  assert.ok(grid.innerHTML.includes("&lt;script&gt;"), "script tags must be escaped");
  assert.ok(!grid.innerHTML.includes("<script>"), "no raw script tag may reach the DOM");
  // The default widgets now echo cell text (the recent-transactions table), so
  // "onerror=alert" can legitimately appear as escaped text. What must never
  // appear is a real tag or an event-handler attribute built from that text.
  assert.ok(!/<img\b/i.test(grid.innerHTML), "no img tag may reach the DOM");
  assert.ok(!/<svg[^>]*\son[a-z]+\s*=/i.test(grid.innerHTML), "no event handler attribute may reach the DOM");
});

test("a workbook with no ledger is flagged as a guess, not presented as fact", () => {
  const api = loadAnalysis();
  const showcase = {
    sheets: [{
      name: "Rumus",
      columns: ["A", "B", "C"],
      rows: [
        ["JUDUL DOKUMEN", "", ""],
        ["Nama Fungsi", "Hasil Uji", "Status"],
        ["SUM", 10, "PASS"],
        ["AVERAGE", 30, "PASS"],
        ["MAX", 50, "PASS"],
      ],
    }],
  };
  const fields = api.detectFields(showcase);
  assert.equal(fields.confidence, "guess", "no money column exists here");
  const built = api.buildLedger(showcase, fields);
  assert.ok(built.records.length >= 1);

  const meta = { textContent: "", className: "" };
  const root = { querySelector: (selector) => (selector === "#aw-meta" ? meta : null) };
  api.render(root, showcase, { fields, widgets: api.autoWidgets() });
  assert.match(meta.textContent, /ditebak/);
  assert.equal(meta.className, "aw-meta warn");
});

test("the keyword helper mirrors the server fallback", () => {
  const api = loadAnalysis();
  const widgets = api.keywordWidgets("untung bersih per bulan");
  assert.ok(widgets.some((w) => w.metric === "net"));
  assert.ok(widgets.some((w) => w.type === "line"));
  assert.ok(api.keywordWidgets("apa saja").length >= 1);
});

test("a Jenis column never hijacks the category binding", () => {
  const api = loadAnalysis();
  // Shape of the managed Transaksi sheet: ID | Tanggal | Jenis | Kategori | ...
  const book = {
    sheets: [{
      name: "Transaksi",
      columns: ["A", "B", "C", "D", "E", "F", "G"],
      rows: [
        ["ID", "Tanggal", "Jenis", "Kategori", "Keterangan", "Pemasukan", "Pengeluaran"],
        ["T-1", "2026-05-01", "Pemasukan", "Gaji", "Gaji bulanan", 14500000, ""],
        ["T-2", "2026-05-06", "Pemasukan", "Proyek", "Desain logo", 3200000, ""],
        ["T-3", "2026-05-02", "Pengeluaran", "Sewa", "Sewa apartemen", "", 4800000],
        ["T-4", "2026-05-05", "Pengeluaran", "Makan", "Belanja bulanan", "", 2150000],
        ["T-5", "2026-05-09", "Pengeluaran", "Transport", "Bensin dan tol", "", 680000],
      ],
    }],
  };
  const fields = api.detectFields(book);
  assert.equal(fields.date, 1);
  assert.equal(fields.category, 3, "Kategori, not Jenis");
  assert.equal(fields.note, 4);
  assert.equal(fields.income, 5);
  assert.equal(fields.expense, 6);

  const built = api.buildLedger(book, fields);
  const categories = new Set(built.records.map((record) => record.category));
  assert.ok(categories.has("Sewa"));
  assert.ok(!categories.has("Pengeluaran"), "the type flag must not become a category");
});

test("the default table shows the latest rows, newest first", () => {
  const api = loadAnalysis();
  const fields = api.detectFields(ledger);
  const built = api.buildLedger(ledger, fields);
  const recent = api.autoWidgets().find((widget) => widget.type === "table");
  assert.ok(recent && recent.recent, "the default table is a recent-activity list");
  const data = api.computeWidget(recent, built);
  const dates = data.rows.map((row) => row.date.toISOString().slice(0, 10));
  // join() keeps the comparison inside one realm; deepEqual would trip over the
  // sandbox's own Array prototype.
  assert.equal(dates.join(","), [...dates].sort().reverse().join(","), "newest first");
  assert.equal(dates[0], "2026-03-09");
});

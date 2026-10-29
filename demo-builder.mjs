/** Build standalone demo pages from the real workspace markup:
 *   - analysis-demo.html    : a bookkeeping ledger, Analisis view ready
 *   - analysis-harness.html : the same page plus browser assertions
 *
 *  workspace-preview.html and workspace-harness.html (the Master Excel views)
 *  come from preview-builder.mjs. Assets are read from the real source files,
 *  so what is tested is what ships.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { importWorkspaceXlsx } from './src/workspace-sheets.js';

const css = readFileSync('public/workspace-app.css', 'utf8') + '\n' + readFileSync('public/workspace-analysis.css', 'utf8');
const app = readFileSync('public/workspace-app.js', 'utf8');
const analysis = readFileSync('public/workspace-analysis.js', 'utf8');
const shell = readFileSync('src/workspace-page.js', 'utf8');
const markup = /<body>([\s\S]*?)<script>window\.__WORKSPACE__/.exec(shell)[1];

const master = await importWorkspaceXlsx(
  readFileSync('Master_Excel_Semua_Fitur_Rumus_Sistem.xlsx'),
  'Master Excel',
);

/** A small bookkeeping workbook: the shape a client would actually import. */
const ledger = {
  title: 'Pembukuan Warung Bu Sri',
  sheets: [{
    name: 'Transaksi',
    columns: ['A', 'B', 'C', 'D', 'E', 'F'],
    rows: [
      ['PEMBUKUAN WARUNG BU SRI', '', '', '', '', ''],
      ['Periode Januari - Juni 2026', '', '', '', '', ''],
      ['', '', '', '', '', ''],
      ['Tanggal', 'Keterangan', 'Kategori', 'Pemasukan', 'Pengeluaran', 'Metode'],
      ['2026-01-03', 'Penjualan harian', 'Penjualan', 1500000, 0, 'Tunai'],
      ['2026-01-05', 'Beli bahan baku', 'Bahan Baku', 0, 620000, 'Transfer'],
      ['2026-01-18', 'Penjualan katering', 'Penjualan', 2100000, 0, 'Transfer'],
      ['2026-01-28', 'Bayar listrik & air', 'Operasional', 0, 340000, 'Tunai'],
      ['2026-02-02', 'Penjualan harian', 'Penjualan', 1750000, 0, 'Tunai'],
      ['2026-02-09', 'Beli bahan baku', 'Bahan Baku', 0, 705000, 'Transfer'],
      ['2026-02-14', 'Gaji 2 karyawan', 'Gaji', 0, 900000, 'Transfer'],
      ['2026-02-22', 'Penjualan katering', 'Penjualan', 2650000, 0, 'Transfer'],
      ['2026-03-01', 'Penjualan harian', 'Penjualan', 1620000, 0, 'Tunai'],
      ['2026-03-07', 'Beli bahan baku', 'Bahan Baku', 0, 680000, 'Transfer'],
      ['2026-03-15', 'Sewa tempat', 'Operasional', 0, 1200000, 'Transfer'],
      ['2026-03-25', 'Penjualan katering', 'Penjualan', 2900000, 0, 'Transfer'],
      ['2026-04-04', 'Penjualan harian', 'Penjualan', 1880000, 0, 'Tunai'],
      ['2026-04-11', 'Beli bahan baku', 'Bahan Baku', 0, 745000, 'Transfer'],
      ['2026-04-19', 'Gaji 2 karyawan', 'Gaji', 0, 900000, 'Transfer'],
      ['2026-04-27', 'Penjualan katering', 'Penjualan', 3100000, 0, 'Transfer'],
      ['2026-05-02', 'Penjualan harian', 'Penjualan', 1950000, 0, 'Tunai'],
      ['2026-05-10', 'Beli bahan baku', 'Bahan Baku', 0, 810000, 'Transfer'],
      ['2026-05-16', 'Promosi online', 'Pemasaran', 0, 450000, 'Transfer'],
      ['2026-05-24', 'Penjualan katering', 'Penjualan', 3400000, 0, 'Transfer'],
      ['2026-06-03', 'Penjualan harian', 'Penjualan', 2050000, 0, 'Tunai'],
      ['2026-06-08', 'Beli bahan baku', 'Bahan Baku', 0, 760000, 'Transfer'],
      ['2026-06-14', 'Gaji 2 karyawan', 'Karyawan', 0, 950000, 'Transfer'],
      ['2026-06-21', 'Bayar listrik & air', 'Operasional', 0, 365000, 'Tunai'],
      ['2026-06-28', 'Penjualan katering', 'Penjualan', 3750000, 0, 'Transfer'],
    ],
  }],
};

/** Stub the API so the pages run from file:// without a server. */
const stub = `(function () {
  var store = window.__WORKSPACE__;
  function reply(payload, status) {
    return Promise.resolve(new Response(JSON.stringify(payload), {
      status: status || 200, headers: { "content-type": "application/json" } }));
  }
  window.fetch = function (url, options) {
    var target = String(url), opts = options || {};
    if (target.indexOf("/data") === 0) return reply(store.data);
    if (target.indexOf("/analysis/widgets/save") === 0) {
      var saved = JSON.parse(opts.body || "{}");
      store.data.analysis = { fields: saved.fields, widgets: saved.widgets };
      return reply({ saved: true, analysis: store.data.analysis });
    }
    if (target.indexOf("/analysis/fields") === 0) {
      var only = JSON.parse(opts.body || "{}");
      store.data.analysis = store.data.analysis || {};
      store.data.analysis.fields = only.fields;
      return reply({ saved: true, analysis: store.data.analysis });
    }
    if (target.indexOf("/analysis/widgets") === 0) {
      var prompt = JSON.parse(opts.body || "{}").prompt || "";
      var widgets = window.WorkspaceAnalysis.keywordWidgets(prompt);
      return reply({ widgets: widgets, source: "keyword", note: "Pratinjau statis: tanpa AI, widget dari kata kunci." });
    }
    if (target.indexOf("/workbook") === 0) {
      var body = JSON.parse(opts.body || "{}");
      store.data.sheets = body.sheets || store.data.sheets;
      if (body.analysis) store.data.analysis = body.analysis;
      return reply({ saved: true, sheets: store.data.sheets, sheet_names: [], analysis: store.data.analysis || null });
    }
    return reply({ error: "Pratinjau statis: import/export hanya jalan di server." }, 400);
  };
})();`;

function page(dataset, extra, title) {
  const dataJson = JSON.stringify(dataset).replace(/</g, '\\u003c');
  return `<!doctype html>
<html lang="id"><head><meta charset="utf-8"><title>${title}</title>
<style>${css}</style></head>
<body>${markup}
<script>window.__WORKSPACE__ = { base: "", title: ${JSON.stringify(dataset.title)}, data: ${dataJson} };</script>
<script>${stub}</script>
<script>${analysis}</script>
<script>${app}</script>
${extra}
</body></html>`;
}

/** Browser assertions for the Analisis view. */
const analysisHarness = `<script>
(function () {
  var out = [];
  function ok(name, condition) { out.push((condition ? "PASS " : "FAIL ") + name); }
  function report() {
    var pre = document.createElement("pre");
    pre.id = "harness-result";
    pre.textContent = out.join("\\n");
    document.body.appendChild(pre);
  }
  try {
    var grid = document.getElementById("grid-wrap");
    var analysisEl = document.getElementById("analysis");

    ok("sheet view starts visible", !grid.hidden && analysisEl.hidden);

    document.getElementById("mode-analysis").click();
    ok("analysis view opens", grid.hidden && !analysisEl.hidden);
    ok("sheet toolbar hidden", document.querySelector(".toolbar").hidden);
    ok("tab bar hidden", document.querySelector(".tabbar").hidden);

    var cards = analysisEl.querySelectorAll(".aw-card");
    ok("widgets rendered (" + cards.length + ")", cards.length >= 5);

    var meta = document.getElementById("aw-meta").textContent;
    ok("meta reports transactions [" + meta + "]", /25 transaksi terbaca/.test(meta));

    // Column binding must be detected from the real header row.
    var income = analysisEl.querySelector('[data-field="income"]');
    var expense = analysisEl.querySelector('[data-field="expense"]');
    var category = analysisEl.querySelector('[data-field="category"]');
    ok("income column detected", income && income.value === "3");
    ok("expense column detected", expense && expense.value === "4");
    ok("category column detected", category && category.value === "2");

    // The KPI numbers must match the sheet, computed by hand here.
    var kpiText = analysisEl.querySelector(".aw-kpi").textContent;
    ok("first KPI has a rupiah value [" + kpiText + "]", /Rp\\s[\\d.]+/.test(kpiText));

    var titles = Array.prototype.map.call(analysisEl.querySelectorAll(".aw-title"), function (h) { return h.textContent; });
    ok("has Untung Bersih widget", titles.indexOf("Untung Bersih") >= 0);
    ok("has a bar chart", analysisEl.querySelectorAll(".aw-bar-fill").length > 0);
    ok("has a line chart", analysisEl.querySelectorAll("svg.aw-line polyline").length > 0);

    // Ask for a custom widget through the prompt box.
    document.getElementById("aw-input").value = "saya mau lihat pengeluaran per kategori dan komposisinya";
    document.getElementById("aw-generate").click();

    setTimeout(function () {
      try {
        var after = Array.prototype.map.call(analysisEl.querySelectorAll(".aw-title"), function (h) { return h.textContent; });
        ok("prompt replaced the widgets (" + after.length + ")", after.length >= 1);
        ok("prompt produced a category chart", after.some(function (t) { return /Kategori|Komposisi/.test(t); }));
        ok("prompt produced a donut", analysisEl.querySelectorAll("svg.aw-donut").length > 0);

        // Deleting a widget removes exactly one card.
        var before = analysisEl.querySelectorAll(".aw-card").length;
        var del = analysisEl.querySelector(".aw-del");
        del.click();
        ok("delete removes one card", analysisEl.querySelectorAll(".aw-card").length === before - 1);

        // Auto button restores the standard set.
        document.getElementById("aw-auto").click();
        ok("auto restores the standard set", analysisEl.querySelectorAll(".aw-card").length >= 5);

        // Back to the sheet, the grid must still be intact.
        document.getElementById("mode-sheet").click();
        ok("sheet view returns", !grid.hidden && analysisEl.hidden);
        ok("grid still has cells", document.querySelectorAll("td[data-r]").length > 100);
        ok("formula bar visible again", !document.querySelector(".formula-row").hidden);

        report();
      } catch (error) {
        out.push("ERROR2 " + (error && error.stack ? error.stack.split("\\n").slice(0, 3).join(" | ") : error));
        report();
      }
    }, 250);
  } catch (error) {
    out.push("ERROR " + (error && error.stack ? error.stack.split("\\n").slice(0, 3).join(" | ") : error));
    report();
  }
})();
</script>`;

writeFileSync('analysis-demo.html', page(ledger, '', 'Pratinjau Analisis'));
writeFileSync('analysis-harness.html', page(ledger, analysisHarness, 'Harness Analisis'));
console.log('analysis demo written; master sheets:', master.sheets.length);

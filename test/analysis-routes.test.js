import assert from "node:assert/strict";
import { createServer } from "node:http";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { FinanceStore } from "../src/finance/store.js";
import { FileWorkspace } from "../src/files/workspace.js";
import { ShareLinkStore } from "../src/share-links.js";
import { runTool } from "../src/tools.js";
import { handleWorkspaceShare } from "../src/share-routes.js";

const LEDGER = [
  ["Pembukuan Warung", "", "", "", ""],
  ["", "", "", "", ""],
  ["Tanggal", "Keterangan", "Kategori", "Pemasukan", "Pengeluaran"],
  ["2026-01-03", "Penjualan harian", "Penjualan", 1500000, 0],
  ["2026-01-05", "Beli bahan", "Bahan", 0, 620000],
  ["2026-02-02", "Penjualan harian", "Penjualan", 1750000, 0],
  ["2026-02-08", "Bayar listrik", "Operasional", 0, 340000],
];

async function startWorkspace(t) {
  const dataDirectory = mkdtempSync(join(tmpdir(), "analysis-route-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const shareLinks = new ShareLinkStore({ dataDirectory });
  const userId = "6281234567890";
  const share = await runTool("create_book", "Asia/Jakarta", { name: "Warung" }, {
    financeStore, fileWorkspace, shareLinks, userId, publicOrigin: "https://example.test",
  });
  const config = { panelOwnerNumber: userId };
  const server = createServer((request, response) => {
    void handleWorkspaceShare(request, response, { shareLinks, fileWorkspace, config })
      .then((handled) => { if (!handled && !response.writableEnded) response.writeHead(404).end(); })
      .catch((error) => {
        if (!response.headersSent) response.writeHead(500);
        response.end(error.message);
      });
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const link = shareLinks.list(userId)[0];
  t.after(async () => {
    await new Promise((resolve) => server.close(resolve));
    shareLinks.close();
    fileWorkspace.close();
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });
  return { origin, name: link.name, base: `${origin}/w/${link.name}` };
}

test("analysis assets are served and the page carries the mode switch", async (t) => {
  const { base } = await startWorkspace(t);

  const html = await (await fetch(base)).text();
  assert.match(html, /id="mode-sheet"/);
  assert.match(html, /id="mode-analysis"/);
  assert.match(html, /id="aw-grid"/);
  assert.match(html, /id="aw-source"/);
  assert.match(html, /\/analysis\.js/);
  assert.match(html, /\/analysis\.css/);

  const script = await fetch(`${base}/analysis.js`);
  assert.equal(script.status, 200);
  assert.match(script.headers.get("content-type"), /javascript/);
  assert.ok((await script.text()).length > 5000);

  const style = await fetch(`${base}/analysis.css`);
  assert.equal(style.status, 200);
  assert.match(style.headers.get("content-type"), /css/);
});

test("the AI endpoint returns usable widgets without an API key", async (t) => {
  const { base } = await startWorkspace(t);
  await fetch(`${base}/workbook`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ title: "Warung", sheets: [{ name: "Transaksi", columns: ["A", "B", "C", "D", "E"], rows: LEDGER }] }),
  });

  const response = await fetch(`${base}/analysis/widgets`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ prompt: "saya mau lihat untung bersih dan pengeluaran per kategori" }),
  });
  assert.equal(response.status, 200);
  const result = await response.json();
  assert.equal(result.source, "keyword", "no key configured on this server");
  assert.ok(result.widgets.length >= 2);
  assert.ok(result.widgets.every((w) => typeof w.type === "string" && typeof w.title === "string"));
  assert.ok(result.widgets.some((w) => w.metric === "net"));

  const empty = await fetch(`${base}/analysis/widgets`, {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ prompt: "  " }),
  });
  assert.equal(empty.status, 400);
});

test("column bindings and widget lists are persisted", async (t) => {
  const { base } = await startWorkspace(t);
  await fetch(`${base}/workbook`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ title: "Warung", sheets: [{ name: "Transaksi", columns: ["A", "B", "C", "D", "E"], rows: LEDGER }] }),
  });

  const fields = { sheet: 0, headerRow: 2, date: 0, note: 1, category: 2, income: 3, expense: 4, amount: null };
  const savedFields = await fetch(`${base}/analysis/fields`, {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ fields }),
  });
  assert.equal(savedFields.status, 200);
  assert.equal((await savedFields.json()).analysis.fields.income, 3);

  const widgets = [{ id: "w1", type: "kpi", title: "Untung Bersih", metric: "net", agg: "sum", groupBy: "none", format: "currency", limit: 8 }];
  const savedWidgets = await fetch(`${base}/analysis/widgets/save`, {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ fields, widgets }),
  });
  assert.equal(savedWidgets.status, 200);

  // The binding survives a reload and rides along with the dataset.
  const data = await (await fetch(`${base}/data`)).json();
  assert.equal(data.analysis.fields.expense, 4);
  assert.equal(data.analysis.widgets.length, 1);
  assert.equal(data.analysis.widgets[0].title, "Untung Bersih");

  // The analysis block must never leak into the spreadsheet itself.
  assert.equal(data.sheets[0].analysis, undefined);
  const exported = await fetch(`${base}/export`);
  assert.equal(exported.status, 200);
  assert.deepEqual([...new Uint8Array(await exported.arrayBuffer())].slice(0, 2), [0x50, 0x4b]);
});

test("a binding without a money column is rejected", async (t) => {
  const { base } = await startWorkspace(t);
  const response = await fetch(`${base}/analysis/fields`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ fields: { sheet: 0, headerRow: 0, date: 0, note: 1, category: 2, income: null, expense: null, amount: null } }),
  });
  assert.equal(response.status, 400);
});

test("saving the workbook keeps the analysis block", async (t) => {
  const { base } = await startWorkspace(t);
  const fields = { sheet: 0, headerRow: 2, date: 0, note: 1, category: 2, income: 3, expense: 4, amount: null };
  const widgets = [{ id: "w1", type: "kpi", title: "Pemasukan", metric: "income", agg: "sum", groupBy: "none", format: "currency", limit: 8 }];

  const saved = await fetch(`${base}/workbook`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      title: "Warung",
      sheets: [{ name: "Transaksi", columns: ["A", "B", "C", "D", "E"], rows: LEDGER }],
      analysis: { fields, widgets },
    }),
  });
  assert.equal(saved.status, 200);
  const body = await saved.json();
  assert.equal(body.analysis.widgets[0].metric, "income");

  // A later save that omits the block must not silently wipe it.
  const second = await fetch(`${base}/data`);
  assert.equal((await second.json()).analysis.fields.income, 3);
});

test("unknown widget types are stripped instead of stored", async (t) => {
  const { base } = await startWorkspace(t);
  const fields = { sheet: 0, headerRow: 2, date: 0, note: 1, category: 2, income: 3, expense: 4, amount: null };
  const response = await fetch(`${base}/analysis/widgets/save`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      fields,
      widgets: [
        { id: "ok", type: "kpi", title: "Aman", metric: "net" },
        { id: "bad", type: "iframe", title: "Jahat" },
        { id: "bad2", type: "kpi", title: "Metric aneh", metric: "__proto__" },
      ],
    }),
  });
  assert.equal(response.status, 200);
  const stored = (await response.json()).analysis.widgets;
  assert.equal(stored.length, 2);
  assert.ok(stored.every((w) => w.type !== "iframe"));
  assert.equal(stored.find((w) => w.id === "bad2").metric, "net");
});

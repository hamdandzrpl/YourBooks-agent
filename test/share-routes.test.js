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

test("workspace page loads, saves rows, and exports without tokens", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "share-route-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const shareLinks = new ShareLinkStore({ dataDirectory });
  const userId = "6281234567890";
  const share = await runTool("create_book", "Asia/Jakarta", { name: "Pertanian" }, {
    financeStore, fileWorkspace, shareLinks, userId, publicOrigin: "https://example.test",
  });
  const server = createServer((request, response) => {
    void handleWorkspaceShare(request, response, {
      shareLinks, fileWorkspace, config: { panelOwnerNumber: userId },
    }).then((handled) => {
      if (!handled && !response.writableEnded) response.writeHead(404).end();
    }).catch((error) => {
      if (!response.headersSent) response.writeHead(500);
      response.end(error.message);
    });
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  t.after(async () => {
    await new Promise((resolve) => server.close(resolve));
    shareLinks.close();
    fileWorkspace.close();
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const page = await fetch(origin + share.workspace_path);
  assert.equal(page.status, 200);
  const html = await page.text();
  const transaction = await runTool("record_transaction", "Asia/Jakarta", {
    kind: "income", amount: 1_000_000, category: "Kopi", description: "Penjualan kopi",
  }, { financeStore, fileWorkspace, shareLinks, userId, publicOrigin: "https://example.test", sourceMessageId: "share-route-regression" });
  assert.ok(transaction.id);
  // The shell carries the API base and loads its client from /app.js.
  assert.match(html, /window\.__WORKSPACE__ = \{/);
  assert.match(html, new RegExp(`"base":"/w/finance-b-[a-f0-9]{8}"`));
  assert.match(html, /\/w\/finance-b-[a-f0-9]{8}\/app\.js/);
  assert.match(html, /id="name-box"/);
  assert.match(html, /id="formula-input"/);
  assert.doesNotMatch(html, /share-name/);

  // Assets are served from /public and must parse as real JS.
  const assetBase = `/w/${shareLinks.list(userId)[0].name}`;
  const script = await fetch(origin + assetBase + "/app.js");
  assert.equal(script.status, 200);
  assert.match(script.headers.get("content-type"), /javascript/);
  assert.ok((await script.text()).length > 5000);
  const style = await fetch(origin + assetBase + "/app.css");
  assert.equal(style.status, 200);
  assert.match(style.headers.get("content-type"), /css/);

  const reopened = await fetch(origin + share.workspace_path);
  assert.equal(reopened.status, 200);
  const link = shareLinks.list(userId)[0];
  const dataUrl = `${origin}/w/${link.name}/data`;
  const initial = await fetch(dataUrl);
  assert.equal(initial.status, 200);
  const initialData = await initial.json();
  // The header travels as a real first row so the web grid can draw it; the
  // XLSX export then writes it exactly once instead of re-creating it.
  assert.deepEqual(initialData.sheets[0].rows[0], ["Tanggal", "Deskripsi", "Kategori", "Jenis", "Nominal"]);
  assert.equal(initialData.sheets[0].rows.length, 2);

  const saved = await fetch(`${origin}/w/${link.name}/rows`, {
    method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify({ rows: [["2026-10-27", "Panen kopi", "Kopi", "Pemasukan", "1000000"]] }),
  });
  assert.equal(saved.status, 200);
  assert.equal((await saved.json()).row_count, 1);

  const exportResponse = await fetch(`${origin}/w/${link.name}/export`);
  assert.equal(exportResponse.status, 200);
  assert.match(exportResponse.headers.get("content-type"), /spreadsheetml/);
  assert.match(exportResponse.headers.get("content-disposition"), /filename="finance-b-[a-f0-9]+\.xlsx"/);
  assert.deepEqual([...new Uint8Array(await exportResponse.arrayBuffer())].slice(0, 2), [0x50, 0x4b]);
});

test("a dashboard workbook is readable as a summary block on /data", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "share-summary-"));
  const financeStore = new FinanceStore(join(dataDirectory, "finance.sqlite"));
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const shareLinks = new ShareLinkStore({ dataDirectory });
  const userId = "6281234567891";
  const share = await runTool("create_book", "Asia/Jakarta", { name: "Kebun" }, {
    financeStore, fileWorkspace, shareLinks, userId, publicOrigin: "https://example.test",
  });
  const server = createServer((request, response) => {
    void handleWorkspaceShare(request, response, {
      shareLinks, fileWorkspace, config: { panelOwnerNumber: userId },
    }).then((handled) => {
      if (!handled && !response.writableEnded) response.writeHead(404).end();
    }).catch((error) => {
      if (!response.headersSent) response.writeHead(500);
      response.end(error.message);
    });
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  t.after(async () => {
    await new Promise((resolve) => server.close(resolve));
    shareLinks.close();
    fileWorkspace.close();
    financeStore.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const link = shareLinks.list(userId)[0];
  const workbook = {
    title: "Kebun",
    sheets: [
      { name: "Dashboard", columns: ["Komponen", "Nilai"], headerRow: true, rows: [
        ["Total Pemasukan", { formula: "SUM(Pemasukan!B:B)" }],
        ["Total Pengeluaran", { formula: "SUM(Pengeluaran!B:B)" }],
        ["Laba Bersih", { formula: "B2-B3" }],
      ] },
      { name: "Pemasukan", columns: ["Tanggal", "Total"], rows: [["2024-01-15", 7500000]] },
      { name: "Pengeluaran", columns: ["Tanggal", "Total"], rows: [["2024-01-05", 2500000]] },
    ],
  };
  const saved = await fetch(`${origin}/w/${link.name}/workbook`, {
    method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify(workbook),
  });
  assert.equal(saved.status, 200);
  assert.equal((await saved.json()).saved, true);

  const data = await (await fetch(`${origin}/w/${link.name}/data`)).json();
  assert.equal(data.sheets.length, 3);
  assert.ok(data.summary, "the dataset must carry a summary block");
  assert.equal(data.summary.sheet, "Dashboard");
  assert.deepEqual(data.summary.kpis.map((kpi) => [kpi.label, kpi.value]), [
    ["Total Pemasukan", 7500000],
    ["Total Pengeluaran", 2500000],
    ["Laba Bersih", 5000000],
  ]);
  assert.deepEqual(data.summary.missing.filter((item) => item.includes("KPI")), []);
  // The summary must not become a fourth sheet in the grid or the export.
  assert.deepEqual(data.sheets.map((sheet) => sheet.name), ["Dashboard", "Pemasukan", "Pengeluaran"]);
});

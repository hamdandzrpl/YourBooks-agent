import assert from "node:assert/strict";
import test from "node:test";
import {
  generateWidgets,
  widgetsFromKeywords,
  sanitizeWidgets,
  sanitizeWidget,
  extractJson,
  buildDigest,
} from "../src/analysis-ai.js";

/** A realistic small ledger, in the shape the importer produces. */
function ledgerDataset() {
  return {
    title: "Pembukuan",
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
}

test("the model is asked for a specification only, never data", async () => {
  let captured = null;
  const fakeFetch = async (url, options) => {
    captured = { url, body: JSON.parse(options.body), headers: options.headers };
    return {
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '{"widgets":[{"type":"kpi","title":"Untung Bersih","metric":"net","format":"currency"}]}' } }],
      }),
    };
  };

  const result = await generateWidgets({
    prompt: "saya mau lihat untung bersih",
    dataset: ledgerDataset(),
    config: { aiApiKey: "test-key", chatCompletionsUrl: "https://api.example.com/v1/chat/completions", aiModel: "test/model" },
    fetchImpl: fakeFetch,
  });

  assert.equal(result.source, "ai");
  assert.equal(result.widgets.length, 1);
  assert.equal(result.widgets[0].metric, "net");
  assert.equal(result.widgets[0].type, "kpi");
  assert.equal(captured.body.model, "test/model");
  assert.match(captured.headers.authorization, /^Bearer /);
  // The digest must not ship the whole workbook.
  const userMessage = captured.body.messages[1].content;
  assert.ok(userMessage.includes("Permintaan: saya mau lihat untung bersih"));
  assert.ok(!userMessage.includes("Gaji karyawan"), "digest should sample, not dump every row");
});

test("a broken provider falls back to keywords instead of failing", async () => {
  const result = await generateWidgets({
    prompt: "pengeluaran per kategori",
    dataset: ledgerDataset(),
    config: { aiApiKey: "k", chatCompletionsUrl: "https://api.example.com/v1/chat/completions", aiModel: "m" },
    fetchImpl: async () => { throw new Error("network down"); },
  });
  assert.equal(result.source, "keyword");
  assert.ok(result.widgets.length >= 1);
  assert.ok(result.widgets.some((w) => w.type === "bar" && w.groupBy === "category"));
  assert.match(result.note, /kata kunci/);
});

test("an HTTP error from the provider also falls back", async () => {
  const result = await generateWidgets({
    prompt: "untung bersih",
    dataset: ledgerDataset(),
    config: { aiApiKey: "k", chatCompletionsUrl: "https://api.example.com/v1/chat/completions", aiModel: "m" },
    fetchImpl: async () => ({ ok: false, status: 429 }),
  });
  assert.equal(result.source, "keyword");
  assert.match(result.note, /429/);
});

test("a provider reply with no usable widgets falls back", async () => {
  const result = await generateWidgets({
    prompt: "untung bersih",
    dataset: ledgerDataset(),
    config: { aiApiKey: "k", chatCompletionsUrl: "https://api.example.com/v1/chat/completions", aiModel: "m" },
    fetchImpl: async () => ({ ok: true, json: async () => ({ choices: [{ message: { content: "maaf saya tidak bisa" } }] }) }),
  });
  assert.equal(result.source, "keyword");
  assert.match(result.note, /tidak mengembalikan/);
});

test("without a key the keyword parser answers directly", async () => {
  const result = await generateWidgets({ prompt: "untung bersih per bulan", dataset: ledgerDataset(), config: {} });
  assert.equal(result.source, "keyword");
  assert.ok(result.widgets.some((w) => w.metric === "net"));
  assert.ok(result.widgets.some((w) => w.type === "line" && w.groupBy === "month"));
});

test("an empty prompt asks for one instead of guessing", async () => {
  const result = await generateWidgets({ prompt: "   ", dataset: ledgerDataset(), config: {} });
  assert.equal(result.source, "none");
  assert.equal(result.widgets.length, 0);
});

test("widget specs are whitelisted, never passed through", () => {
  assert.equal(sanitizeWidget({ type: "eval" }), null);
  assert.equal(sanitizeWidget({ type: "kpi", metric: "steal" }).metric, "net");
  assert.equal(sanitizeWidget({ type: "kpi", limit: 9999 }).limit, 20);
  assert.equal(sanitizeWidget({ type: "kpi", limit: -5 }).limit, 1);
  assert.equal(sanitizeWidget({ type: "kpi", groupBy: "drop table" }).groupBy, "none");
  assert.equal(sanitizeWidget({ type: "bar", groupBy: "none" }).groupBy, "category");
  assert.equal(sanitizeWidget({ type: "line", groupBy: "category" }).groupBy, "month");
  assert.equal(sanitizeWidget({ type: "table", metric: "income" }).metric, "net");
  assert.equal(sanitizeWidget({ type: "kpi", title: "x".repeat(200) }).title.length, 80);

  const progress = sanitizeWidget({ type: "progress", target: -50 });
  assert.equal(progress.target, undefined);
  assert.equal(progress.format, "percent");

  assert.equal(sanitizeWidgets([{ type: "nope" }, { type: "kpi" }]).length, 1);
  assert.equal(sanitizeWidgets("not an array").length, 0);
  assert.equal(sanitizeWidgets(new Array(50).fill({ type: "kpi" })).length, 24);

  // Generated ids must stay unique so delete buttons never collide.
  const many = sanitizeWidgets(new Array(12).fill({ type: "kpi", id: "same" }));
  assert.equal(new Set(many.map((w) => w.id)).size, many.length);
});

test("JSON is extracted from fenced and chatty replies", () => {
  assert.deepEqual(extractJson('```json\n{"widgets":[]}\n```'), { widgets: [] });
  assert.deepEqual(extractJson('Tentu! Ini hasilnya:\n{"a":1}\nSemoga membantu.'), { a: 1 });
  assert.deepEqual(extractJson('[{"type":"kpi"}]'), [{ type: "kpi" }]);
  assert.equal(extractJson("tidak ada json"), null);
  assert.equal(extractJson('{"broken": '), null);
  // Braces inside strings must not end the object early.
  assert.deepEqual(extractJson('{"title":"a } b"}'), { title: "a } b" });
});

test("the digest stays small even on a big workbook", () => {
  const rows = [];
  for (let i = 0; i < 5000; i++) rows.push([`row ${i}`, i, "x".repeat(200)]);
  const digest = buildDigest({ sheets: [{ name: "Big", columns: ["A", "B", "C"], rows }] });
  assert.equal(digest[0].rows, 5000);
  assert.ok(digest[0].sample.length <= 8);
  assert.ok(digest[0].sample[0].every((cell) => String(cell).length <= 40));
});

test("keyword parsing covers the phrasings a client actually types", () => {
  const cases = [
    ["saya mau lihat untung bersih", "net"],
    ["berapa pemasukan bulan ini", "income"],
    ["total pengeluaran", "expense"],
  ];
  for (const [prompt, metric] of cases) {
    const widgets = widgetsFromKeywords(prompt);
    assert.ok(widgets.some((w) => w.metric === metric), `${prompt} should mention ${metric}`);
  }
  assert.ok(widgetsFromKeywords("komposisi pengeluaran").some((w) => w.type === "donut"));
  assert.ok(widgetsFromKeywords("transaksi terbesar").some((w) => w.type === "table"));
  assert.ok(widgetsFromKeywords("tren per bulan").some((w) => w.type === "line"));
  // An unparseable request still returns something usable.
  assert.ok(widgetsFromKeywords("halo").length >= 1);
});

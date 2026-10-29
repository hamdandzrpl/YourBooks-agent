/** Turn a plain-language dashboard request into widget specs.
 *
 *  The model only ever returns a *specification* (which metric, which grouping,
 *  which chart). Every number is computed from the spreadsheet by the browser,
 *  so the AI cannot invent a figure and cannot see the whole dataset.
 *
 *  When no API key is configured the keyword parser takes over, so the feature
 *  still works on a deployment without credentials.
 */

const SYSTEM_PROMPT = `You design dashboard widgets for a spreadsheet.
Reply with JSON only, no prose, no code fences. Shape:
{"widgets":[{"type":"kpi","title":"Untung Bersih","metric":"net","agg":"sum","groupBy":"none","format":"currency","limit":8}]}

Allowed values:
- type: kpi, bar, line, donut, table, progress
- metric: income, expense, net, balance, count
- agg: sum, avg, count, min, max
- groupBy: none, category, month
- format: currency, number, percent
- limit: integer 1..20 (used by bar, donut, table)
- table widgets may add "recent": true to list the latest transactions instead of the largest
- progress widgets also take "target": a positive number

Rules:
- 1 to 5 widgets, chosen to answer the request.
- Titles in the user's language, short (max 6 words).
- Never return data or numbers from the sheet; only the specification.`;

import { sanitizeWidgets } from "./analysis-spec.js";

// Re-exported so callers have one obvious import for widget validation.
export { sanitizeWidget, sanitizeWidgets } from "./analysis-spec.js";

/** Compact description of the workbook for the prompt. */
export function buildDigest(dataset, { maxSheets = 6, maxRows = 8, maxCols = 10 } = {}) {
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  return sheets.slice(0, maxSheets).map((sheet, index) => {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    const sample = rows.slice(0, maxRows).map((row) =>
      (Array.isArray(row) ? row.slice(0, maxCols) : []).map((cell) => {
        const value = cell && typeof cell === "object" ? cell.value : cell;
        if (value === undefined || value === null) return "";
        return typeof value === "number" ? value : String(value).slice(0, 40);
      }));
    return { index, name: sheet.name, rows: rows.length, sample };
  });
}

/** First JSON object/array found in a reply, tolerating fences and prose. */
export function extractJson(text) {
  if (typeof text !== "string") return null;
  const start = text.search(/[[{]/);
  if (start < 0) return null;
  const open = text[start];
  const close = open === "[" ? "]" : "}";
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === open) depth++;
    else if (ch === close) {
      depth--;
      if (depth === 0) {
        try { return JSON.parse(text.slice(start, i + 1)); } catch { return null; }
      }
    }
  }
  return null;
}

/** Offline fallback: keyword rules. Enough to be genuinely useful. */
export function widgetsFromKeywords(prompt) {
  const text = String(prompt).toLowerCase();
  const found = [];
  const has = (pattern) => pattern.test(text);

  const wantsIncome = has(/pemasukan|pendapatan|income|masuk|omset|penjualan|revenue/);
  const wantsExpense = has(/pengeluaran|belanja|biaya|keluar|expense|cost|spending/);
  const wantsNet = has(/untung|laba|profit|margin|net|bersih|selisih/);
  const wantsCategory = has(/kategori|category|jenis|pos|akun|per bagian/);
  const wantsMonth = has(/bulan|tren|trend|waktu|periode|monthly|per bulan/);
  const wantsShare = has(/komposisi|proporsi|porsi|persentase|share|donat|pie/);
  const wantsTop = has(/terbesar|tertinggi|top|terbanyak|ranking/);
  const wantsTarget = has(/target|progres|progress|capai/);

  const push = (widget) => { if (found.length < 5) found.push(widget); };

  if (wantsNet) push({ type: "kpi", title: "Untung Bersih", metric: "net", format: "currency" });
  if (wantsIncome) push({ type: "kpi", title: "Pemasukan", metric: "income", format: "currency" });
  if (wantsExpense) push({ type: "kpi", title: "Pengeluaran", metric: "expense", format: "currency" });
  if (wantsShare) {
    push({ type: "donut", title: "Komposisi Pengeluaran", metric: wantsExpense ? "expense" : "income", groupBy: "category", format: "currency" });
  } else if (wantsCategory) {
    push({ type: "bar", title: "Per Kategori", metric: wantsExpense ? "expense" : "income", groupBy: "category", format: "currency" });
  }
  if (wantsMonth) push({ type: "line", title: "Tren per Bulan", metric: wantsNet ? "net" : "income", groupBy: "month", format: "currency" });
  if (wantsTop) push({ type: "table", title: "Transaksi Terbesar", format: "currency" });
  if (wantsTarget) push({ type: "progress", title: "Target", metric: "income", target: 1000000, format: "percent" });

  if (!found.length) {
    push({ type: "kpi", title: "Untung Bersih", metric: "net", format: "currency" });
    push({ type: "bar", title: "Pengeluaran per Kategori", metric: "expense", groupBy: "category", format: "currency" });
  }
  return sanitizeWidgets(found);
}

/**
 * Ask the model for widgets. Never throws: on any failure it falls back to the
 * keyword parser and says so in `note`, so the UI can explain what happened.
 */
export async function generateWidgets({ prompt, dataset, config, fetchImpl = fetch, timeoutMs = 25_000 }) {
  const request = String(prompt ?? "").trim().slice(0, 500);
  if (!request) return { widgets: [], source: "none", note: "Tulis dulu mau lihat analisis apa." };

  if (!config?.aiApiKey || !config?.chatCompletionsUrl) {
    return {
      widgets: widgetsFromKeywords(request),
      source: "keyword",
      note: "AI belum dikonfigurasi, widget disusun dari kata kunci.",
    };
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const { buildHeaders } = await import("./provider.js");
    const response = await fetchImpl(config.chatCompletionsUrl, {
      method: "POST",
      headers: buildHeaders(config.chatCompletionsUrl, config.aiApiKey),
      body: JSON.stringify({
        model: config.aiModel,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          {
            role: "user",
            content: `Data (contoh baris, bukan seluruh isi):\n${JSON.stringify(buildDigest(dataset))}\n\nPermintaan: ${request}`,
          },
        ],
        temperature: 0.2,
      }),
      signal: controller.signal,
    });

    if (!response.ok) {
      return {
        widgets: widgetsFromKeywords(request),
        source: "keyword",
        note: `AI menolak permintaan (${response.status}), widget disusun dari kata kunci.`,
      };
    }

    const payload = await response.json();
    const content = payload?.choices?.[0]?.message?.content ?? "";
    const parsed = extractJson(content);
    const widgets = sanitizeWidgets(Array.isArray(parsed) ? parsed : parsed?.widgets);
    if (!widgets.length) {
      return {
        widgets: widgetsFromKeywords(request),
        source: "keyword",
        note: "AI tidak mengembalikan widget yang valid, dipakai kata kunci.",
      };
    }
    return { widgets, source: "ai", note: "" };
  } catch {
    return {
      widgets: widgetsFromKeywords(request),
      source: "keyword",
      note: "AI tidak bisa dihubungi, widget disusun dari kata kunci.",
    };
  } finally {
    clearTimeout(timer);
  }
}

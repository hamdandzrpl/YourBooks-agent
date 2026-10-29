/** Analysis layer for a workspace workbook.
 *
 *  The spreadsheet stays the source of truth. An `analysis` block records which
 *  column means what (income, expense, date, category) and which dashboard
 *  widgets to draw. Nothing here changes the sheets, and the XLSX export
 *  ignores it entirely, so a workbook with analysis still exports as pure Excel.
 *
 *  This module is also the single place where a widget spec is validated, so
 *  the browser, the API and the AI reply all go through the same whitelist.
 */

export const WIDGET_TYPES = new Set(["kpi", "bar", "line", "donut", "table", "progress"]);
export const METRICS = new Set(["income", "expense", "net", "balance", "count"]);
export const AGGS = new Set(["sum", "avg", "count", "min", "max"]);
export const GROUP_BYS = new Set(["none", "category", "month"]);
export const FORMATS = new Set(["currency", "number", "percent"]);

export const MAX_WIDGETS = 24;
const MAX_TITLE = 80;
const MAX_COL = 255;
const MAX_LIMIT = 20;

function intOrNull(value, max) {
  // `Number(null)` is 0, which would silently bind column A; reject it first.
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  if (!Number.isInteger(n) || n < 0 || n > max) return null;
  return n;
}

function textOrNull(value, max) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim().slice(0, max);
  return trimmed || null;
}

/** Sanitize the column binding that tells the dashboard where the money is. */
export function sanitizeFields(raw, sheetCount) {
  if (!raw || typeof raw !== "object") return null;
  const sheet = intOrNull(raw.sheet, Math.max(0, sheetCount - 1));
  if (sheet === null) return null;
  const fields = {
    sheet,
    headerRow: intOrNull(raw.headerRow, 5000) ?? 0,
    date: intOrNull(raw.date, MAX_COL),
    income: intOrNull(raw.income, MAX_COL),
    expense: intOrNull(raw.expense, MAX_COL),
    category: intOrNull(raw.category, MAX_COL),
    note: intOrNull(raw.note, MAX_COL),
    amount: intOrNull(raw.amount, MAX_COL),
  };
  // A binding with no money column at all cannot produce a dashboard.
  if (fields.income === null && fields.expense === null && fields.amount === null) return null;
  // Only keep a guess marker that the browser set; never trust other values.
  if (raw.confidence === "header" || raw.confidence === "guess") fields.confidence = raw.confidence;
  return fields;
}

export function defaultTitle(type, metric) {
  const names = { income: "Pemasukan", expense: "Pengeluaran", net: "Untung Bersih", balance: "Saldo", count: "Jumlah Transaksi" };
  const base = names[metric] ?? "Nilai";
  if (type === "kpi") return base;
  if (type === "bar") return `${base} per Kategori`;
  if (type === "line") return `${base} per Bulan`;
  if (type === "donut") return `Komposisi ${base}`;
  if (type === "table") return "Transaksi Terbesar";
  return `Target ${base}`;
}

/** Whitelist one widget coming from the browser or from an AI reply.
 *  Anything unknown is dropped or replaced with a safe default, never passed
 *  through, so a hostile payload cannot reach the DOM. */
export function sanitizeWidget(raw, index = 0) {
  if (!raw || typeof raw !== "object") return null;
  const type = String(raw.type ?? "").trim().toLowerCase();
  if (!WIDGET_TYPES.has(type)) return null;

  const widget = {
    id: /^[\w-]{1,24}$/.test(String(raw.id ?? "")) ? String(raw.id) : `w${index + 1}`,
    type,
    title: textOrNull(raw.title, MAX_TITLE) ?? defaultTitle(type, raw.metric),
    metric: METRICS.has(raw.metric) ? raw.metric : "net",
    agg: AGGS.has(raw.agg) ? raw.agg : "sum",
    groupBy: GROUP_BYS.has(raw.groupBy) ? raw.groupBy : "none",
    format: FORMATS.has(raw.format) ? raw.format : "currency",
    limit: Math.min(MAX_LIMIT, Math.max(1, Math.round(Number(raw.limit)) || 8)),
  };

  if (type === "progress") {
    const target = Number(raw.target);
    if (Number.isFinite(target) && target > 0) widget.target = target;
    widget.format = "percent";
  }
  // A table always shows the biggest transactions; a line always runs by month.
  if (type === "table") {
    widget.metric = "net";
    widget.groupBy = "none";
    widget.format = "currency";
    // `recent` flips the table from "largest" to "latest", which is what a
    // bookkeeping view wants on its first screen.
    if (raw.recent === true) {
      widget.recent = true;
      if (!textOrNull(raw.title, MAX_TITLE)) widget.title = "Transaksi Terakhir";
    }
  }
  if (type === "line") widget.groupBy = "month";
  if ((type === "bar" || type === "donut") && widget.groupBy === "none") widget.groupBy = "category";
  return widget;
}

export function sanitizeWidgets(raw) {
  const list = Array.isArray(raw) ? raw : Array.isArray(raw?.widgets) ? raw.widgets : [];
  const out = [];
  const seen = new Set();
  for (const item of list.slice(0, 50)) {
    if (out.length >= MAX_WIDGETS) break;
    const widget = sanitizeWidget(item, out.length);
    if (!widget) continue;
    // Duplicate ids would break the delete buttons.
    while (seen.has(widget.id)) widget.id = `${widget.id}_`;
    seen.add(widget.id);
    out.push(widget);
  }
  return out;
}

export function sanitizeAnalysis(raw, sheetCount) {
  if (!raw || typeof raw !== "object") return null;
  const fields = sanitizeFields(raw.fields, sheetCount);
  if (!fields) return null;
  return { fields, widgets: sanitizeWidgets(raw.widgets) };
}

/** Read the summary sheet of a workspace dataset into a structured block.
 *
 *  The Workspace view must show the numbers the workbook's own formulas
 *  produced, not a second calculation. That only works if the summary sheet
 *  has a shape the system can read, so this module defines that shape and
 *  reports exactly what is missing when a sheet does not match it. The rules it
 *  enforces are the ones written in `skills/dashboard-workbooks.json`.
 */

const SUMMARY_NAME = /^(dashboard|ringkasan|summary|ikhtisar|dasbor)\b/i;
const KPI_WORDS = /(total|jumlah|laba|margin|untung|rugi|pemasukan|pendapatan|pengeluaran|biaya|belanja|saldo|kas|omset|penjualan|income|expense|profit|revenue)/i;
const PERCENT_WORDS = /(margin|persen|persentase|%|ratio|rasio)/i;
const MAX_SCAN_ROWS = 400;
const MAX_KPI = 12;
const MAX_TABLE_ROWS = 200;

const cellValue = (cell) => (cell && typeof cell === "object" ? cell.value : cell);

/** Numbers may arrive as numbers, or as text a human typed ("Rp 88.990.000").
 *  The separator is ambiguous without a locale: "88.990" is eighty-eight
 *  thousand in Indonesia and eighty-eight point nine nine in English. A dotted
 *  group of exactly three digits is read as a thousands separator, anything
 *  else as a decimal mark. */
export function toNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string") return null;
  const text = value.trim();
  if (!text || !/\d/.test(text)) return null;
  const bare = text.replace(/[^\d.,-]/g, "");
  if (!bare || !/\d/.test(bare)) return null;

  const hasDot = bare.includes(".");
  const hasComma = bare.includes(",");
  let normalised;
  if (hasDot && hasComma) {
    // Whichever comes last is the decimal mark.
    const decimal = bare.lastIndexOf(".") > bare.lastIndexOf(",") ? "." : ",";
    const thousands = decimal === "." ? "," : ".";
    normalised = bare.split(thousands).join("").replace(decimal, ".");
  } else if (hasDot || hasComma) {
    const mark = hasDot ? "." : ",";
    const grouped = new RegExp(`^\\d{1,3}(\\${mark}\\d{3})+$`).test(bare);
    normalised = grouped ? bare.split(mark).join("") : bare.replace(mark, ".");
  } else {
    normalised = bare;
  }

  const number = Number(normalised);
  return Number.isFinite(number) ? number : null;
}

const isText = (value) => typeof value === "string" && value.trim() !== "";

/** The first sheet that looks like a summary: a matching tab name wins, then
 *  whichever sheet carries at least two KPI labels. */
export function detectSummarySheet(sheets) {
  const list = Array.isArray(sheets) ? sheets : [];
  const byName = list.find((sheet) => SUMMARY_NAME.test(String(sheet?.name ?? "").trim()));
  if (byName) return byName;
  for (const sheet of list) {
    const labels = kpiLabels(sheet);
    if (labels.length >= 2) return sheet;
  }
  return null;
}

/** Every cell text in a sheet that reads like a KPI label. */
function kpiLabels(sheet) {
  const found = [];
  (sheet?.rows ?? []).forEach((row, r) => (row ?? []).forEach((cell, c) => {
    const text = cellValue(cell);
    if (isText(text) && KPI_WORDS.test(text) && text.length <= 40) found.push({ r, c, text: text.trim() });
  }));
  return found;
}

function formatOf(sheet, r, c) {
  const styles = sheet?.styles;
  if (!styles) return null;
  return styles[`${r},${c}`]?.format ?? null;
}

function kindOf(label, format) {
  if (PERCENT_WORDS.test(label) || (typeof format === "string" && format.includes("%"))) return "percent";
  if (typeof format === "string" && /[$€£]|Rp|#,##0|0\.00/.test(format)) return "number";
  return "number";
}

/** The value that belongs to a label: the nearest number to its right on the
 *  same row, otherwise the nearest number directly below it. Both layouts are
 *  common — a KPI strip puts the value under its label, a KPI list puts it
 *  beside it. */
function valueFor(sheet, r, c) {
  const rows = sheet?.rows ?? [];
  const row = rows[r] ?? [];
  for (let step = 1; step <= 3; step++) {
    const number = toNumber(cellValue(row[c + step]));
    if (number !== null) return { row: r, col: c + step, value: number };
    const other = cellValue(row[c + step]);
    if (isText(other) && KPI_WORDS.test(other)) break; // a neighbouring label, not a value
  }
  for (let step = 1; step <= 2; step++) {
    const below = rows[r + step] ?? [];
    const number = toNumber(cellValue(below[c]));
    if (number !== null) return { row: r + step, col: c, value: number };
  }
  return null;
}

function rowIsBlank(row) {
  return !(row ?? []).some((cell) => cellValue(cell) !== "" && cellValue(cell) !== null && cellValue(cell) !== undefined);
}

/** A summary table starts at a header row of at least three text cells whose
 *  next row carries at least two numbers, and whose body is a list of
 *  dimensions (text) with numbers beside them.
 *
 *  That last rule is what separates a table from the KPI strip: a strip of
 *  "Total Pemasukan | Total Pengeluaran | ..." over one row of numbers has no
 *  dimension column, so it is read as KPIs instead of as a one-row table. */
function findTable(sheet) {
  const rows = sheet?.rows ?? [];
  const limit = Math.min(rows.length, MAX_SCAN_ROWS);
  let best = null;
  for (let r = 0; r < limit; r++) {
    const header = (rows[r] ?? []).map(cellValue);
    const filled = header.filter((value) => value !== "" && value !== null && value !== undefined);
    if (filled.length < 3) continue;
    if (!filled.every(isText)) continue;
    const next = (rows[r + 1] ?? []).map(cellValue);
    if (next.filter((value) => toNumber(value) !== null).length < 2) continue;

    const body = [];
    for (let r2 = r + 1; r2 < Math.min(rows.length, r + 1 + MAX_TABLE_ROWS); r2++) {
      const row = rows[r2] ?? [];
      if (rowIsBlank(row)) break;
      if (!row.some((value) => toNumber(cellValue(value)) !== null)) break;
      body.push(row);
    }
    if (!body.length) continue;

    const firstHeader = header.findIndex((value) => value !== "" && value !== null && value !== undefined);
    const firstBody = (body[0] ?? []).map(cellValue);
    const dimensionIsText = firstHeader >= 0 && isText(header[firstHeader])
      && isText(firstBody[firstHeader]);
    if (!dimensionIsText && body.length < 2) continue;

    const candidate = { r, header, body };
    if (!best || candidate.body.length > best.body.length) best = candidate;
  }
  if (!best) return null;

  // Trim the empty columns a KPI strip leaves behind (a block layout starts at
  // column B), so the table the UI receives has no blank leading column.
  const used = [];
  for (const row of [best.header, ...best.body]) {
    (row ?? []).forEach((value, index) => {
      const cell = cellValue(value);
      if (cell !== "" && cell !== null && cell !== undefined) used.push(index);
    });
  }
  const from = Math.min(...used);
  const to = Math.max(...used);
  const width = to - from + 1;
  return {
    title: null,
    header: Array.from({ length: width }, (_, i) => String(best.header[from + i] ?? "").trim()),
    rows: best.body.map((row) => Array.from({ length: width }, (_, i) => {
      const value = cellValue(row[from + i]);
      const number = toNumber(value);
      return number !== null ? number : (value === null || value === undefined ? "" : String(value));
    })),
    start_row: best.r + 1,
    end_row: best.r + 1 + best.body.length,
  };
}

function titleOf(sheet, firstKpiRow) {
  const rows = sheet?.rows ?? [];
  for (let r = 0; r < Math.min(firstKpiRow, rows.length); r++) {
    const text = (rows[r] ?? []).map(cellValue).find((value) => isText(value) && value.trim().length >= 4);
    if (text) return text.trim().slice(0, 120);
  }
  return null;
}

/**
 * Extract the summary block from a dataset or a single sheet.
 *
 * Returns `{ sheet, title, kpis, table, missing }`. `missing` lists what a
 * reader could not find, so the agent gets a concrete instruction instead of a
 * silently empty dashboard.
 */
export function extractSummary(source) {
  const sheet = Array.isArray(source?.sheets) ? detectSummarySheet(source.sheets) : source;
  if (!sheet) {
    return { sheet: null, title: null, kpis: [], table: null, missing: ["no sheet named Dashboard/Ringkasan and no sheet with KPI labels"] };
  }

  const rows = sheet.rows ?? [];
  const table = findTable(sheet);
  // KPI labels live above the summary table. Scanning the whole sheet would
  // also pick up the table's own header ("Total Pemasukan" is a column there),
  // which reports the first data row as if it were the grand total.
  const kpiLimit = table ? table.start_row - 1 : Math.min(rows.length, MAX_SCAN_ROWS);
  const labels = kpiLabels(sheet).filter((label) => label.r < kpiLimit).slice(0, MAX_KPI);
  const kpis = [];

  for (const label of labels) {
    const hit = valueFor(sheet, label.r, label.c);
    if (!hit) continue;
    if (kpis.some((kpi) => kpi.label === label.text)) continue;
    const format = formatOf(sheet, hit.row, hit.col);
    kpis.push({
      label: label.text,
      value: hit.value,
      kind: kindOf(label.text, format),
      format: format ?? null,
      cell: `${columnName(hit.col)}${hit.row + 1}`,
    });
  }

  const missing = [];
  if (!kpis.length) missing.push("no KPI label paired with a numeric value above the summary table");
  // An empty workbook is the correct empty state: no data exists to summarize,
  // so a missing table is not a defect the agent must fix. Only data on the
  // INPUT sheets counts — the Dashboard's own KPI block is not "data".
  const hasDataRows = Array.isArray(source?.sheets)
    ? source.sheets.some((entry) => entry !== sheet
      && (((entry.headerRow === true ? entry.rows.slice(1) : entry.rows) ?? []).length > 0))
    : ((sheet.headerRow === true ? sheet.rows.slice(1) : sheet.rows) ?? []).length > 0;
  if (!table && hasDataRows) missing.push("no summary table (a header row of 3+ labels followed by numeric rows)");
  if (kpis.length && !kpis.some((kpi) => /total|pemasukan|pendapatan|income/i.test(kpi.label))) {
    missing.push("no income total among the KPIs");
  }
  if (kpis.length && !kpis.some((kpi) => /pengeluaran|biaya|belanja|expense/i.test(kpi.label))) {
    missing.push("no expense total among the KPIs");
  }

  return {
    sheet: String(sheet.name ?? ""),
    title: titleOf(sheet, Math.min(...labels.map((label) => label.r), rows.length)),
    kpis,
    table,
    missing,
  };
}

function columnName(index) {
  let name = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

/** Store the summary beside the sheets, never inside them, so the XLSX export
 *  stays a plain spreadsheet and the grid is unaffected. */
export function withSummary(dataset) {
  const summary = extractSummary(dataset);
  if (!summary.sheet) return dataset;
  return { ...dataset, summary };
}

/** A short version for tool results: enough for the agent to notice a sheet it
 *  just wrote cannot be read, without repeating every number back. */
export function summaryBrief(summary) {
  if (!summary) return null;
  return {
    sheet: summary.sheet,
    title: summary.title,
    kpis: summary.kpis.map((kpi) => ({ label: kpi.label, value: kpi.value, kind: kpi.kind, cell: kpi.cell })),
    table: summary.table
      ? { columns: summary.table.header, rows: summary.table.rows.length, start_row: summary.table.start_row, end_row: summary.table.end_row }
      : null,
    missing: summary.missing,
  };
}

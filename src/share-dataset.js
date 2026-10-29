/** Build the dataset behind a shared web workspace from agent tool arguments.
 *
 *  The web grid already renders several sheets and a per-cell style map, but
 *  the only writer the agent had (`share_link_set_rows`) collapsed everything
 *  into one sheet of plain strings: formulas were stringified and alignment
 *  could not be set at all. This module turns a compact, token-cheap request
 *  (sheets + a few range formats) into the same shape the browser and the XLSX
 *  export already understand.
 */

import { columnLetter } from "./formula-engine.js";
import { sanitizeStyle, BORDER_STYLES } from "./cell-styles.js";
import { validateWorkspace } from "./workspace-sheets.js";

const MAX_SHEETS = 20;
const MAX_ROWS = 5000;
const MAX_COLUMNS = 64;
const MAX_FORMATS = 400;
const MAX_STYLE_CELLS = 20000;
const MAX_CELL_TEXT = 2000;

/** "AB" -> 27 (zero based), matching the grid's column indexes. */
export function columnIndex(letters) {
  const name = String(letters ?? "").trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(name)) return -1;
  let index = 0;
  for (const char of name) index = index * 26 + (char.charCodeAt(0) - 64);
  return index - 1;
}

/** Accepts "B7", "A1:E1" and whole-column "B:B". Returns zero-based bounds
 *  or null when the reference is not usable. */
export function parseCellRange(reference) {
  const text = String(reference ?? "").trim().toUpperCase().replace(/\$/g, "");
  let match = /^([A-Z]{1,3})(\d{1,5})(?::([A-Z]{1,3})(\d{1,5}))?$/.exec(text);
  if (match) {
    const first = { col: columnIndex(match[1]), row: Number(match[2]) - 1 };
    const last = match[3] ? { col: columnIndex(match[3]), row: Number(match[4]) - 1 } : first;
    if (first.col < 0 || last.col < 0) return null;
    return {
      start: { col: Math.min(first.col, last.col), row: Math.min(first.row, last.row) },
      end: { col: Math.max(first.col, last.col), row: Math.max(first.row, last.row) },
    };
  }
  match = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(text);
  if (match) {
    const first = columnIndex(match[1]);
    const last = columnIndex(match[2]);
    if (first < 0 || last < 0) return null;
    return {
      start: { col: Math.min(first, last), row: 0 },
      end: { col: Math.max(first, last), row: Infinity },
    };
  }
  return null;
}

/** Try to read a string as a number the way an Indonesian bookkeeper types it.
 *
 *  The agent sometimes pastes a number with the thousands dot but no comma
 *  (".20000000", "1.234.567", "Rp 88.990.000,-"). Left as text, a SUM over the
 *  column silently treats it as 0, so the dashboard reads wrong totals. When
 *  the string parses unambiguously it becomes a real number cell; otherwise
 *  null keeps the original text.
 */
export function coerceNumberString(value) {
  let text = String(value ?? "").trim();
  if (!text) return null;
  text = text.replace(/^(Rp|IDR|USD|EUR)\s*/i, "").replace(/[-\s]+$/g, "").replace(/,$/, "");
  if (!text) return null;
  if (/^-?\d+$/.test(text)) return Number(text);
  // 1.234.567 — thousands dots, no decimal.
  if (/^\d{1,3}(\.\d{3})+$/.test(text)) return Number(text.replace(/\./g, ""));
  // .20000000 — a mangled paste; treat the digits as the integer.
  if (/^\.\d+$/.test(text) && text.length > 2) return Number(text.slice(1));
  // 1.234,56 or 1234,56 — Indonesian decimal comma.
  if (/^\d{1,3}(\.\d{3})+,\d+$/.test(text)) return Number(text.replace(/\./g, "").replace(",", "."));
  if (/^\d+,\d+$/.test(text)) return Number(text.replace(",", "."));
  // 1234.56 — a plain decimal point.
  if (/^-?\d+\.\d+$/.test(text)) return Number(text);
  return null;
}

/** A cell keeps its number, boolean, or formula object; everything else
 *  becomes text. Stringifying blindly is what used to erase formulas. */
function normalizeCellValue(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : "";
  if (typeof value === "boolean") return value;
  if (value && typeof value === "object" && typeof value.formula === "string") {
    const cell = { formula: value.formula.replace(/^=/, "").slice(0, 1000) };
    if (typeof value.value === "number" || typeof value.value === "boolean" || typeof value.value === "string") {
      cell.value = value.value;
    }
    return cell;
  }
  if (typeof value === "string") {
    const number = coerceNumberString(value);
    if (number !== null && Number.isFinite(number)) return number;
  }
  return String(value ?? "").slice(0, MAX_CELL_TEXT);
}

function normalizeRows(rows, width) {
  const out = [];
  for (const row of (Array.isArray(rows) ? rows : []).slice(0, MAX_ROWS)) {
    if (!Array.isArray(row)) continue;
    out.push(row.slice(0, width).map(normalizeCellValue));
  }
  return out;
}

function sheetName(raw, index) {
  return String(raw?.name ?? `Sheet${index + 1}`).trim().slice(0, 31) || `Sheet${index + 1}`;
}

/** Turn one requested sheet into `{name, columns, rows, headerRow?, styles?}`.
 *  Named columns become a real first row here, before any format is applied,
 *  so a range like A1:C1 means the header the agent sees in the grid. */
function buildSheet(raw, index, formats) {
  const name = sheetName(raw, index);
  const labels = (Array.isArray(raw?.columns) ? raw.columns : [])
    .slice(0, MAX_COLUMNS)
    .map((value) => String(value ?? "").trim().slice(0, 120));
  const dataRows = Array.isArray(raw?.rows) ? raw.rows : [];
  const width = Math.max(
    labels.length,
    ...dataRows.slice(0, MAX_ROWS).map((row) => (Array.isArray(row) ? Math.min(row.length, MAX_COLUMNS) : 0)),
    1,
  );
  const rows = normalizeRows(dataRows, width);
  const sheet = { name, columns: Array.from({ length: width }, (_, i) => columnLetter(i)) };
  if (labels.length) {
    sheet.rows = [Array.from({ length: width }, (_, i) => labels[i] ?? ""), ...rows];
    sheet.headerRow = true;
  } else {
    sheet.rows = rows;
    if (raw?.headerRow === true) sheet.headerRow = true;
  }

  // Direct style maps are accepted so a caller can copy formatting verbatim;
  // `formats` below is the token-cheap path for everything else.
  const styles = {};
  if (raw?.styles && typeof raw.styles === "object") {
    for (const key of Object.keys(raw.styles).slice(0, MAX_FORMATS)) {
      const style = sanitizeStyle(raw.styles[key]);
      if (style) styles[key] = style;
    }
  }
  if (formats) applyFormatRange(styles, sheet, formats);
  if (Object.keys(styles).length) sheet.styles = styles;
  return sheet;
}

/** Write each requested format entry onto a sheet's style map. Whole-column
 *  ranges are bounded by the sheet height so they cannot explode the map. */
function applyFormatRange(styles, sheet, entries) {
  for (const entry of entries) {
    const bounds = parseCellRange(entry.range);
    if (!bounds) continue;
    const style = sanitizeStyle(formatStyle(entry));
    if (!style) continue;
    const lastRow = bounds.end.row === Infinity
      ? Math.max(sheet.rows.length - 1, 0)
      : Math.min(bounds.end.row, MAX_ROWS - 1);
    const lastCol = Math.min(bounds.end.col, sheet.columns.length - 1);
    for (let row = bounds.start.row; row <= lastRow; row++) {
      for (let col = bounds.start.col; col <= lastCol; col++) {
        if (col < 0) continue;
        if (Object.keys(styles).length >= MAX_STYLE_CELLS) return;
        styles[`${row},${col}`] = style;
      }
    }
  }
}

/** Callers write hex with or without the leading "#"; the style map only
 *  accepts the canonical form, so normalise here instead of dropping the
 *  colour and silently falling back to black. */
function normalizeHex(value) {
  if (typeof value !== "string") return undefined;
  const text = value.trim().replace(/^#/, "");
  return /^[0-9a-fA-F]{6}$/.test(text) ? `#${text.toUpperCase()}` : undefined;
}

/** A single border name applies to all four edges, which is what callers mean
 *  by "beri garis". An object is passed through to the style sanitizer. */
function formatStyle(entry) {
  const style = { ...entry };
  delete style.sheet;
  delete style.range;
  if (typeof style.border === "string") {
    const kind = BORDER_STYLES.has(style.border) ? style.border : "thin";
    const color = normalizeHex(style.border_color) ?? "#D4D4D8";
    style.border = { top: { style: kind, color }, right: { style: kind, color }, bottom: { style: kind, color }, left: { style: kind, color } };
  }
  delete style.border_color;
  if (style.color !== undefined) style.color = normalizeHex(style.color);
  if (style.fill !== undefined) style.fill = normalizeHex(style.fill);
  return style;
}

/** Group requested formats by sheet name; unnamed entries target the first
 *  sheet, which is what a single-sheet request expects. */
function formatsBySheet(formats, names) {
  const grouped = new Map();
  for (const entry of (Array.isArray(formats) ? formats : []).slice(0, MAX_FORMATS)) {
    if (!entry || typeof entry !== "object") continue;
    const wanted = typeof entry.sheet === "string" ? entry.sheet.trim().toLowerCase() : "";
    const target = wanted ? names.find((name) => name.toLowerCase() === wanted) : names[0];
    if (!target) continue;
    if (!grouped.has(target)) grouped.set(target, []);
    grouped.get(target).push(entry);
  }
  return grouped;
}

/**
 * Build a validated workspace dataset from tool arguments.
 *
 * Accepts either the multi-sheet form (`sheets: [...]`) or the original
 * single-sheet form (`columns` + `rows`), plus an optional `formats` list of
 * range formatting such as `{range: "A1:E1", align: "center", bold: true}`.
 */
export function buildShareDataset(args, fallbackTitle = "Workspace") {
  const title = String(args?.title ?? fallbackTitle ?? "Workspace").slice(0, 120) || "Workspace";
  const requested = Array.isArray(args?.sheets) && args.sheets.length
    ? args.sheets.slice(0, MAX_SHEETS)
    : [{ name: "Sheet1", columns: args?.columns, rows: args?.rows, headerRow: args?.headerRow }];
  if (!requested.length) throw new Error("At least one sheet is required");
  if (!Array.isArray(args?.sheets) && !Array.isArray(args?.columns) && !Array.isArray(args?.rows)) {
    throw new Error("Provide sheets[] for a multi-sheet document, or columns/rows for a single sheet");
  }

  const names = requested.map((sheet, index) => sheetName(sheet, index));
  const formats = formatsBySheet(args?.formats, names);
  const sheets = requested.map((sheet, index) => buildSheet(sheet, index, formats.get(names[index]) ?? null));

  // validateWorkspace rejects duplicate names, unsafe formulas, and oversized
  // input, so the tool can never persist a dataset the browser cannot render.
  return validateWorkspace({ title, sheets });
}

/** Human-readable summary for the tool result: the agent verifies what it
 *  wrote from this instead of re-reading the whole dataset. */
export function describeDataset(dataset) {
  return {
    title: dataset.title,
    sheet_count: dataset.sheets.length,
    sheets: dataset.sheets.map((sheet) => ({
      name: sheet.name,
      rows: sheet.rows.length,
      columns: sheet.columns.length,
      header_row: sheet.headerRow === true,
      formatted_cells: sheet.styles ? Object.keys(sheet.styles).length : 0,
    })),
  };
}

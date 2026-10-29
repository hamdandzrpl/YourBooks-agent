import ExcelJS from "exceljs";
import { computeWorkbook, columnLetter } from "./formula-engine.js";
import { readExcelStyle, toExcelStyle, sanitizeStyles } from "./cell-styles.js";

const maxSheets = 50;
const maxRows = 5000;
const maxColumns = 256;
const safeFormula = /^[\s\S]{1,1000}$/;

export function normalizeWorkspace(data) {
  if (Array.isArray(data.sheets) && data.sheets.length) {
    const names = new Set();
    const sheets = data.sheets.slice(0, maxSheets).map((sheet, index) => {
      const normalized = normalizeSheet(sheet, index);
      if (names.has(normalized.name.toLowerCase())) throw new Error("Sheet names must be unique");
      names.add(normalized.name.toLowerCase());
      return normalized;
    });
    return { title: String(data.title ?? "Workspace").slice(0, 120), sheets };
  }
  return { title: String(data.title ?? "Workspace").slice(0, 120), sheets: [normalizeSheet({ name: "Sheet1", columns: data.columns ?? [], rows: data.rows ?? [] }, 0)] };
}

/** True when every column label is just an Excel letter (A, B, ... AA).
 *  Those are grid coordinates, not a header row. */
export function columnsAreExcelLetters(columns) {
  return Array.isArray(columns) && columns.length > 0 &&
    columns.every((name, index) => String(name ?? "").trim().toUpperCase() === columnLetter(index));
}

/** Move every recorded style down by `delta` rows. Prepending a header row
 *  shifts all data, so the style map has to follow or formatting lands on the
 *  wrong cells. */
function shiftStyleRows(styles, delta) {
  const out = {};
  for (const key of Object.keys(styles)) {
    const [row, column] = key.split(",");
    const next = Number(row) + delta;
    if (!Number.isFinite(next) || next < 0) continue;
    out[`${next},${column}`] = styles[key];
  }
  return out;
}

function normalizeSheet(sheet, index) {
  let columns = (Array.isArray(sheet.columns) ? sheet.columns : []).slice(0, maxColumns).map((value) => String(value ?? "").slice(0, 120));
  let rows = (Array.isArray(sheet.rows) ? sheet.rows : []).slice(0, maxRows).map((row) => Array.isArray(row) ? row.slice(0, maxColumns).map(normalizeCell) : []);
  let styles = sanitizeStyles(sheet.styles);
  let headerRow = sheet.headerRow === true;
  // One convention for every dataset: the header is an ordinary first row and
  // `columns` only carries Excel letters. Documents written before this rule
  // stored the header in `columns` and left it out of `rows`, so the web grid
  // drew no header while the XLSX export re-created one — the header looked
  // like it vanished on the web and came back in Excel.
  if (columns.length && !columnsAreExcelLetters(columns)) {
    const first = rows[0] ?? [];
    const alreadyFirst = first.length >= columns.length &&
      columns.every((name, i) => String(first[i] ?? "").trim() === name.trim());
    if (!alreadyFirst) {
      rows = [columns.slice(), ...rows].slice(0, maxRows);
      styles = shiftStyleRows(styles, 1);
    }
    columns = [];
    headerRow = true;
  }
  const width = Math.max(columns.length, ...rows.map((row) => row.length), 1);
  columns = Array.from({ length: width }, (_, i) => columns[i] || columnLetter(i));
  const out = { name: String(sheet.name ?? `Sheet${index + 1}`).slice(0, 31) || `Sheet${index + 1}`, columns, rows };
  // Only carry the map when it holds something, so plain sheets stay lean.
  if (headerRow) out.headerRow = true;
  if (Object.keys(styles).length) out.styles = styles;
  return out;
}

function normalizeCell(value) {
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (value && typeof value === "object" && typeof value.formula === "string") {
    const cell = { formula: value.formula.replace(/^=/, "").slice(0, 1000) };
    if (value.value !== undefined) cell.value = value.value;
    return cell;
  }
  return String(value ?? "").slice(0, 2000);
}

/** A formula result Excel stored in the file, in the shape a dataset cell uses.
 *  Errors are kept as strings so a reader can tell them apart from a value. */
function readableCached(result) {
  if (result === undefined || result === null) return undefined;
  if (typeof result === "number") return Number.isFinite(result) ? result : undefined;
  if (typeof result === "boolean") return result;
  if (result instanceof Date) return result.toISOString().slice(0, 10);
  if (typeof result === "string") return result.slice(0, 2000);
  if (typeof result === "object" && "error" in result) return String(result.error ?? "").slice(0, 40);
  return undefined;
}

export function validateWorkspace(data) {
  const normalized = normalizeWorkspace(data);
  for (const sheet of normalized.sheets) for (const row of sheet.rows) for (const cell of row) {
    if (cell && typeof cell === "object" && cell.formula && !safeFormula.test(cell.formula)) {
      throw new Error("Formula cannot be calculated safely. Supported: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, IFERROR, ROUND, ABS, PRODUCT, CONCAT, SUMIF, COUNTIF.");
    }
  }
  return normalized;
}

export async function importWorkspaceXlsx(buffer, title) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheets = [];
  for (const ws of wb.worksheets) {
    // Keep every row at its original Excel position so formulas like F18 stay valid.
    const rows = [];
    const styles = {};
    let colCount = 0;
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const rowData = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        while (rowData.length < colNumber - 1) rowData.push("");
        const v = cell.value;
        if (v && typeof v === "object" && v.formula) {
          const entry = { formula: String(v.formula).replace(/^=/, "").slice(0, 1000) };
          // Keep the value Excel already computed. Without it a workbook whose
          // formulas the web engine cannot evaluate (SUMIFS with date criteria,
          // for instance) would show #VALUE! everywhere instead of the numbers
          // the file already carries.
          const cached = readableCached(v.result);
          if (cached !== undefined) entry.value = cached;
          rowData.push(entry);
        } else if (typeof v === "number" || typeof v === "boolean") {
          rowData.push(v);
        } else if (v instanceof Date) {
          rowData.push(v.toISOString().slice(0, 10));
        } else if (v && typeof v === "object" && v.sharedFormula) {
          // A shared formula carries only a result (sometimes none at all); the
          // expression itself lives on the anchor cell, so store the value and
          // leave the cell blank when Excel never computed one.
          rowData.push(readableCached(v.result) ?? "");
        } else if (v && typeof v === "object" && v.richText) {
          rowData.push(String(v.richText.map((t) => t.text).join("")).slice(0, 2000));
        } else if (v && typeof v === "object" && "result" in v) {
          rowData.push(String(v.result ?? "").slice(0, 2000));
        } else {
          rowData.push(String(v ?? "").slice(0, 2000));
        }
        // Formatting is captured whether or not the cell holds a value.
        const style = readExcelStyle(cell);
        if (style) styles[`${rowNumber - 1},${colNumber - 1}`] = style;
      });
      if (rowData.length > colCount) colCount = rowData.length;
      rows.push(rowData);
    });
    // Pad every row to the sheet width so column indexes match Excel columns.
    for (const row of rows) while (row.length < colCount) row.push("");
    const width = Math.max(colCount, 1);
    const columns = Array.from({ length: width }, (_, i) => columnLetter(i));
    const sheet = { name: ws.name, columns, rows: rows.slice(0, maxRows) };
    if (Object.keys(styles).length) sheet.styles = styles;
    sheets.push(normalizeSheet(sheet, sheets.length));
  }
  if (!sheets.length) {
    sheets.push({ name: "Sheet1", columns: ["A", "B"], rows: [] });
  }
  return computeWorkbook({ title: String(title ?? wb.creator ?? "Imported").slice(0, 120), sheets });
}

/** True when every column label is just an Excel letter (A, B, ... AA).
 *  Those are grid coordinates, not a header row, so they must not be exported
 *  as data. Normalized datasets always land here, so the header is written by
 *  the rows themselves and never duplicated. */
function columnsAreLetters(columns) {
  return columnsAreExcelLetters(columns);
}

export async function exportWorkspaceXlsx(data) {
  const data2 = validateWorkspace(data);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Workspace";
  for (const sheet of data2.sheets) {
    const worksheet = workbook.addWorksheet(sheet.name);
    const hasHeaderRow = !columnsAreLetters(sheet.columns);
    if (hasHeaderRow) {
      worksheet.addRow(sheet.columns);
    }
    for (const row of sheet.rows) worksheet.addRow(row.map((cell) => {
      if (cell && typeof cell === "object" && typeof cell.formula === "string") {
        // Store the cached result so Excel shows the number before recalculating.
        const value = cell.value;
        if (typeof value === "number" || typeof value === "boolean" || typeof value === "string") {
          return { formula: cell.formula, result: value };
        }
        return { formula: cell.formula };
      }
      return cell;
    }));

    // The header is a normal row now; keep the Excel niceties that used to come
    // from writing it separately (bold, frozen) via the recorded flag. This must
    // run after the rows exist — touching getRow(1) on an empty sheet creates a
    // phantom first row and pushes every real row down by one.
    if (sheet.headerRow || hasHeaderRow) {
      worksheet.getRow(1).font = { bold: true };
      worksheet.views = [{ state: "frozen", ySplit: 1 }];
    } else {
      worksheet.views = [];
    }

    // Apply formatting only where the dataset recorded it.
    const styles = sheet.styles ?? {};
    const rowOffset = hasHeaderRow ? 1 : 0;
    for (const key of Object.keys(styles)) {
      const [r, c] = key.split(",").map(Number);
      const excelStyle = toExcelStyle(styles[key]);
      if (!Object.keys(excelStyle).length) continue;
      worksheet.getCell(r + 1 + rowOffset, c + 1).style = excelStyle;
    }
  }
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

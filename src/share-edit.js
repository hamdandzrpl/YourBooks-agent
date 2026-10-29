/** Incremental editing and self-checking for a shared web workspace.
 *
 *  Two gaps made the agent fail at simple follow-ups:
 *
 *  1. `share_link_set_rows` replaces the whole dataset. Adding one row meant
 *     re-emitting every sheet, so a request like "tambah cabai rawit" burned
 *     the whole tool budget and died. `appendRows` edits in place instead.
 *
 *  2. Nothing told the agent which column letter holds which header, and
 *     nothing checked that a cross-sheet formula pointed at data at all. A
 *     dashboard whose SUMIFS reads an empty column renders as zeros, which
 *     looks like the sheets are not connected. `describeForAgent` exposes the
 *     column letters and `auditFormulas` names the broken references.
 */

import { columnLetter } from "./formula-engine.js";
import { columnIndex, coerceNumberString } from "./share-dataset.js";

const MAX_CELL_TEXT = 2000;

/** Read a sheet's header row as `[{letter, label}]`, so a caller never has to
 *  count columns by hand. */
export function headerColumns(sheet) {
  const width = sheet?.columns?.length ?? 0;
  const header = sheet?.headerRow === true ? (sheet.rows?.[0] ?? []) : null;
  const columns = [];
  for (let index = 0; index < width; index++) {
    const label = header ? String(header[index] ?? "").trim() : "";
    columns.push({ letter: columnLetter(index), index, label });
  }
  return columns;
}

/** True when a column holds any value below the header row. */
function columnHasData(sheet, colIndex) {
  const start = sheet?.headerRow === true ? 1 : 0;
  for (let row = start; row < (sheet.rows?.length ?? 0); row++) {
    const cell = sheet.rows[row]?.[colIndex];
    const value = cell && typeof cell === "object" ? cell.value : cell;
    if (value === null || value === undefined || value === "") continue;
    if (typeof value === "string" && !value.trim()) continue;
    return true;
  }
  return false;
}

/** Compact structural view for the agent: names, column letters, row counts,
 *  and the last data row. Far cheaper than reading dataset.json back. */
export function describeForAgent(dataset, { sheetName, tail = 3 } = {}) {
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  const wanted = typeof sheetName === "string" && sheetName.trim()
    ? sheets.filter((sheet) => sheet.name.toLowerCase() === sheetName.trim().toLowerCase())
    : sheets;
  return {
    title: dataset?.title ?? null,
    sheet_count: sheets.length,
    sheets: wanted.map((sheet) => {
      const header = sheet.headerRow === true ? sheet.rows?.[0] ?? [] : null;
      const dataRows = sheet.headerRow === true ? sheet.rows.slice(1) : sheet.rows;
      return {
        name: sheet.name,
        row_count: sheet.rows?.length ?? 0,
        data_rows: dataRows.length,
        // "A=Tanggal, B=Bulan, ..." so formulas can use the right letter.
        columns: headerColumns(sheet)
          .filter((column) => column.label)
          .map((column) => `${column.letter}=${column.label}`)
          .join(", "),
        header: header ? header.map((cell) => String(cell ?? "")) : null,
        last_rows: dataRows.slice(-Math.max(1, tail)).map((row) =>
          (row ?? []).map((cell) => {
            if (cell && typeof cell === "object") return cell.value ?? cell.formula;
            return cell;
          })),
      };
    }),
  };
}

/** Split a formula into its string literals and everything else, so reference
 *  scanning never rewrites or misreads a quoted label. */
function splitLiterals(formula) {
  const parts = [];
  let index = 0;
  let buffer = "";
  while (index < formula.length) {
    const char = formula[index];
    if (char === '"') {
      const end = formula.indexOf('"', index + 1);
      if (end < 0) { buffer += formula.slice(index); break; }
      parts.push({ literal: false, text: buffer });
      parts.push({ literal: true, text: formula.slice(index, end + 1) });
      buffer = "";
      index = end + 1;
      continue;
    }
    buffer += char;
    index++;
  }
  parts.push({ literal: false, text: buffer });
  return parts;
}

/** Every `Sheet!A:A` / `Sheet!A1` reference in a formula, with the sheet name
 *  unquoted and the column resolved to a zero-based index. */
export function crossSheetRefs(formula) {
  const refs = [];
  for (const part of splitLiterals(String(formula ?? ""))) {
    if (part.literal) continue;
    const pattern = /(?:'([^']+)'|([A-Za-z0-9_\u00C0-\u00FF]+))!\$?([A-Z]{1,3})(?:\$?(\d{1,5}))?(?::\$?([A-Z]{1,3})\$?(\d{1,5}))?/g;
    let match;
    while ((match = pattern.exec(part.text))) {
      const name = match[1] ?? match[2];
      const first = columnIndex(match[3]);
      const last = match[5] ? columnIndex(match[5]) : first;
      if (!name || first < 0) continue;
      refs.push({
        sheet: name,
        from_col: Math.min(first, last < 0 ? first : last),
        to_col: Math.max(first, last < 0 ? first : last),
        whole_column: match[4] === undefined && match[6] === undefined,
        text: match[0],
      });
    }
  }
  return refs;
}

/**
 * Find formula cells that cannot be right, without guessing.
 *
 * Three failures are deterministic and cover the real breakage:
 *  - the formula evaluates to an Excel error;
 *  - it references a sheet that does not exist;
 *  - it references a column of that sheet that holds no data, which is the
 *    classic off-by-one column letter and reads as 0 in the dashboard.
 *
 * Each issue carries a hint naming the column that does have the data, so the
 * agent can repair it in one step instead of re-deriving the layout.
 */
export function auditFormulas(dataset) {
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  const byName = new Map(sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]));
  const issues = [];
  for (const sheet of sheets) {
    (sheet.rows ?? []).forEach((row, rowIndex) => {
      (row ?? []).forEach((cell, colIndex) => {
        if (!cell || typeof cell !== "object" || typeof cell.formula !== "string") return;
        const where = { sheet: sheet.name, cell: `${columnLetter(colIndex)}${rowIndex + 1}`, formula: cell.formula };
        // Check the references before the error value: an unknown sheet
        // evaluates to #REF!, and naming the sheet is far more useful than
        // repeating the code the engine already produced.
        let namedMissingSheet = false;
        for (const ref of crossSheetRefs(cell.formula)) {
          const target = byName.get(ref.sheet.toLowerCase());
          if (!target) {
            namedMissingSheet = true;
            issues.push({
              ...where,
              problem: `references sheet "${ref.sheet}" which does not exist`,
              hint: `Available sheets: ${sheets.map((entry) => entry.name).join(", ")}.`,
            });
            continue;
          }
          if (target === sheet && ref.sheet.toLowerCase() === sheet.name.toLowerCase()) {
            // A same-sheet reference is the sheet's own business.
            continue;
          }
          // A brand-new input sheet has no rows yet, so every whole-column
          // reference would point at an empty column. That is not a bug: the
          // dashboard is supposed to read 0 until the user fills it in. Only
          // when the sheet already holds data is an empty referenced column a
          // real off-by-one to report.
          const dataRowCount = (target.headerRow === true ? target.rows.slice(1) : target.rows).length;
          if (!dataRowCount) continue;
          const empty = [];
          for (let col = ref.from_col; col <= ref.to_col; col++) {
            if (!columnHasData(target, col)) empty.push(columnLetter(col));
          }
          if (!empty.length) continue;
          const withData = headerColumns(target)
            .filter((column) => column.label && columnHasData(target, column.index))
            .map((column) => `${column.letter}=${column.label}`);
          issues.push({
            ...where,
            problem: `reads ${ref.text} but column ${empty.join(",")} of "${target.name}" has no data`,
            hint: withData.length
              ? `Columns of "${target.name}" that do hold data: ${withData.join(", ")}.`
              : `"${target.name}" is still empty, so totals stay 0 until rows are added.`,
          });
        }
        if (namedMissingSheet) return;
        const value = cell.value;
        if (typeof value === "string" && value.startsWith("#")) {
          issues.push({ ...where, problem: `evaluates to ${value}`, hint: "Simplify the formula or use a helper text column (see dashboard-workbooks)." });
        }
      });
    });
  }
  return issues;
}

/** Shift the relative row numbers in a formula by `delta` rows, leaving
 *  absolute (`$5`) references and string literals untouched. */
export function shiftFormulaRows(formula, delta) {
  if (!delta) return formula;
  return splitLiterals(String(formula ?? "")).map((part) => {
    if (part.literal) return part.text;
    return part.text.replace(/(\$?)([A-Z]{1,3})(\$?)(\d{1,5})/g, (match, colAbs, letters, rowAbs, digits) => {
      if (rowAbs === "$") return match;
      const shifted = Number(digits) + delta;
      if (!Number.isFinite(shifted) || shifted < 1) return match;
      return `${colAbs}${letters}${rowAbs}${shifted}`;
    });
  }).join("");
}

function normalizeCell(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : "";
  if (typeof value === "boolean") return value;
  if (typeof value === "object" && typeof value.formula === "string") {
    const cell = { formula: value.formula.replace(/^=/, "").slice(0, 1000) };
    if (["number", "boolean", "string"].includes(typeof value.value)) cell.value = value.value;
    return cell;
  }
  if (typeof value === "string") {
    const number = coerceNumberString(value);
    if (number !== null && Number.isFinite(number)) return number;
  }
  return String(value).slice(0, MAX_CELL_TEXT);
}

/**
 * Append rows to one sheet, in place.
 *
 * A cell given as `null` (or omitted) inherits the formula of the row above,
 * shifted to the new row. That is what makes "add one commodity" cheap: the
 * caller sends only the values it knows and the Total column keeps working.
 * Styles from the row above are copied down for the same reason.
 *
 * Returns a new dataset; the input is not mutated.
 */
export function appendRows(dataset, { sheet: sheetName, rows } = {}) {
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  if (!sheets.length) throw new Error("This workspace has no sheets yet");
  const wanted = String(sheetName ?? "").trim().toLowerCase();
  const index = wanted
    ? sheets.findIndex((sheet) => sheet.name.toLowerCase() === wanted)
    : 0;
  if (index < 0) {
    throw new Error(`No sheet named "${sheetName}". Sheets: ${sheets.map((sheet) => sheet.name).join(", ")}`);
  }
  const target = sheets[index];
  const incoming = (Array.isArray(rows) ? rows : []).filter((row) => Array.isArray(row));
  if (!incoming.length) throw new Error("rows must be a non-empty array of arrays");

  const width = Math.max(target.columns?.length ?? 0, 1);
  const nextRows = (target.rows ?? []).map((row) => [...(row ?? [])]);
  const styles = { ...(target.styles ?? {}) };

  for (const row of incoming) {
    const rowIndex = nextRows.length;
    const above = rowIndex - 1;
    const out = [];
    for (let col = 0; col < width; col++) {
      const given = normalizeCell(row[col]);
      if (given !== null) { out.push(given); continue; }
      // Inherit the formula above, shifted, and the style above as-is.
      const aboveCell = nextRows[above]?.[col];
      if (aboveCell && typeof aboveCell === "object" && typeof aboveCell.formula === "string") {
        out.push({ formula: shiftFormulaRows(aboveCell.formula, rowIndex - above) });
      } else {
        out.push("");
      }
      const aboveStyle = styles[`${above},${col}`];
      if (aboveStyle && !styles[`${rowIndex},${col}`]) styles[`${rowIndex},${col}`] = aboveStyle;
    }
    // Trailing cells beyond the sheet width are dropped rather than silently
    // widening the grid, which would break every whole-column formula.
    nextRows.push(out);
  }

  const updated = { ...target, rows: nextRows };
  if (Object.keys(styles).length) updated.styles = styles;
  return {
    ...dataset,
    sheets: sheets.map((sheet, position) => (position === index ? updated : sheet)),
  };
}

/** Replace every cell of one row, keeping the row in place. Used to correct a
 *  typo without rewriting the workbook. */
export function setRow(dataset, { sheet: sheetName, row, values } = {}) {
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  const wanted = String(sheetName ?? "").trim().toLowerCase();
  const index = wanted ? sheets.findIndex((entry) => entry.name.toLowerCase() === wanted) : 0;
  if (index < 0) throw new Error(`No sheet named "${sheetName}"`);
  const target = sheets[index];
  const rowIndex = Number(row) - 1;
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= (target.rows?.length ?? 0)) {
    throw new Error(`row must be between 1 and ${target.rows?.length ?? 0}`);
  }
  if (!Array.isArray(values)) throw new Error("values must be an array");
  const width = Math.max(target.columns?.length ?? 0, 1);
  const nextRows = target.rows.map((entry) => [...(entry ?? [])]);
  for (let col = 0; col < width; col++) {
    if (col >= values.length) break;
    nextRows[rowIndex][col] = normalizeCell(values[col]) ?? "";
  }
  const updated = { ...target, rows: nextRows };
  return { ...dataset, sheets: sheets.map((entry, position) => (position === index ? updated : entry)) };
}

/** Remove one data row by its 1-based number. */
export function deleteRow(dataset, { sheet: sheetName, row } = {}) {
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  const wanted = String(sheetName ?? "").trim().toLowerCase();
  const index = wanted ? sheets.findIndex((entry) => entry.name.toLowerCase() === wanted) : 0;
  if (index < 0) throw new Error(`No sheet named "${sheetName}"`);
  const target = sheets[index];
  const rowIndex = Number(row) - 1;
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= (target.rows?.length ?? 0)) {
    throw new Error(`row must be between 1 and ${target.rows?.length ?? 0}`);
  }
  const firstDataRow = target.headerRow === true ? 1 : 0;
  if (rowIndex < firstDataRow) throw new Error("Refusing to delete the header row");
  const nextRows = target.rows.filter((_, position) => position !== rowIndex);
  const updated = { ...target, rows: nextRows };
  return { ...dataset, sheets: sheets.map((entry, position) => (position === index ? updated : entry)) };
}

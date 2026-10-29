import { HyperFormula } from "hyperformula";
import { ExcelExtrasPlugin, excelExtrasTranslations } from "./excel-functions.js";

let pluginRegistered = false;
function ensurePlugins() {
  if (pluginRegistered) return;
  try {
    HyperFormula.registerFunctionPlugin(ExcelExtrasPlugin, excelExtrasTranslations);
  } catch {
    // Already registered by another module instance.
  }
  pluginRegistered = true;
}

/** Column letters: 0 -> A, 25 -> Z, 26 -> AA */
export function columnLetter(index) {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

/** Sheet names that Excel leaves unquoted but formula parsers reject
 *  (leading digit, spaces, punctuation) must be quoted for HyperFormula. */
function needsQuote(name) {
  return !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name);
}

/** Quote sheet references inside a formula without touching string literals. */
function quoteSheetRefs(formula, sheetNames) {
  let out = "";
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];
    if (char === '"') {
      const end = formula.indexOf('"', index + 1);
      if (end < 0) { out += formula.slice(index); break; }
      out += formula.slice(index, end + 1);
      index = end + 1;
      continue;
    }
    const match = /^([A-Za-z0-9_\u00C0-\u00FF]+)!/.exec(formula.slice(index));
    if (match && sheetNames.includes(match[1]) && needsQuote(match[1])) {
      out += `'${match[1]}'!`;
      index += match[0].length;
      continue;
    }
    out += char;
    index++;
  }
  return out;
}

/** Excel stores newer functions with an `_xlfn.` / `_xlws.` prefix.
 *  HyperFormula and our plugin registry know the bare names. */
function stripFuturePrefixes(formula) {
  return formula
    .replace(/_xlfn\./gi, "")
    .replace(/_xlws\./gi, "");
}

/** Build a plain 2D grid from a workspace sheet. Rows keep their original
 *  Excel coordinates so A1/F18 style references stay valid. */
function sheetGrid(sheet, sheetNames) {
  const width = Math.max(sheet.columns?.length ?? 0, ...sheet.rows.map((row) => row?.length ?? 0), 1);
  return sheet.rows.map((row) => {
    const out = [];
    for (let col = 0; col < width; col++) {
      const cell = row?.[col];
      if (cell && typeof cell === "object" && typeof cell.formula === "string") {
        const prepared = stripFuturePrefixes(cell.formula.replace(/^=/, ""));
        out.push("=" + quoteSheetRefs(prepared, sheetNames));
      } else if (cell === null || cell === undefined) {
        out.push("");
      } else {
        out.push(cell);
      }
    }
    return out;
  });
}

function readableValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "value" in value) return String(value.value ?? "");
  if (typeof value === "number") return Number.isFinite(value) ? value : "";
  if (typeof value === "boolean") return value;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value);
}

/** Recalculate every formula cell in a workspace dataset.
 *  Returns a new dataset where each formula cell also carries `value`.
 *  Never throws: on engine failure the dataset is returned unchanged. */
export function computeWorkbook(dataset) {
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  if (!sheets.length) return dataset;
  const hasFormula = sheets.some((sheet) =>
    (sheet.rows ?? []).some((row) => (row ?? []).some((cell) => cell && typeof cell === "object" && cell.formula)));
  if (!hasFormula) return dataset;

  const names = [];
  for (let index = 0; index < sheets.length; index++) {
    let name = String(sheets[index].name ?? `Sheet${index + 1}`).slice(0, 31) || `Sheet${index + 1}`;
    while (names.includes(name)) name = `${name}_${index + 1}`.slice(0, 31);
    names.push(name);
  }
  const grids = {};
  for (let index = 0; index < sheets.length; index++) {
    grids[names[index]] = sheetGrid(sheets[index], names);
  }

  let engine;
  try {
    ensurePlugins();
    engine = HyperFormula.buildFromSheets(grids, { licenseKey: "gpl-v3" });
  } catch {
    return dataset;
  }

  const computed = sheets.map((sheet, index) => {
    const sheetId = engine.getSheetId(names[index]);
    return {
      ...sheet,
      rows: (sheet.rows ?? []).map((row, rowIndex) =>
        (row ?? []).map((cell, colIndex) => {
          if (!cell || typeof cell !== "object" || !cell.formula) return cell;
          let value = "";
          try {
            value = readableValue(engine.getCellValue({ sheet: sheetId, col: colIndex, row: rowIndex }));
          } catch {
            value = "";
          }
          // The engine supports a subset of Excel. When it fails but the file
          // already carried a good cached result, the file wins: showing
          // #VALUE! over a number Excel computed would be a regression, and
          // formulas like SUMIFS with date criteria hit exactly this.
          const engineFailed = typeof value === "string" && value.startsWith("#");
          const cached = cell.value;
          const cachedGood = cached !== undefined && cached !== null && cached !== ""
            && !(typeof cached === "string" && cached.startsWith("#"));
          if (engineFailed && cachedGood) value = cached;
          return { formula: cell.formula, value };
        })),
    };
  });
  try { engine.destroy(); } catch { /* engine already gone */ }
  return { ...dataset, sheets: computed, calculated_at: new Date().toISOString() };
}

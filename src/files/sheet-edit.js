// Rich worksheet editing primitives.
//
// The agent used to be able to write plain values into A1 addresses only, so any
// request that changed the *shape* or the *look* of a workbook (add a section,
// colour a total, add a column, merge a heading, widen a column) either failed
// or silently produced an unstyled result. Everything here works on an already
// loaded ExcelJS worksheet and stays inside the same size limits as the rest of
// the file tooling.
//
// The contract is intentionally forgiving: an update describes *intent*
// (label, value, formula, look) and the applier maps it onto ExcelJS. Unknown
// format keys are rejected so a typo never becomes a silent no-op.

const a1Pattern = /^([A-Z]{1,3})([1-9]\d{0,5})$/;
const rangePattern = /^([A-Z]{1,3})([1-9]\d{0,5}):([A-Z]{1,3})([1-9]\d{0,5})$/;

const numFmtPresets = new Map([
  ["currency", "currency"],
  ["uang", "currency"],
  ["money", "currency"],
  ["percent", "0.0%"],
  ["persen", "0.0%"],
  ["date", "dd/mm/yyyy"],
  ["tanggal", "dd/mm/yyyy"],
  ["integer", "#,##0"],
  ["angka", "#,##0"],
  ["decimal", "#,##0.00"],
  ["desimal", "#,##0.00"],
  ["text", "@"],
  ["general", "General"],
]);

const formatKeys = new Set([
  "bold", "italic", "underline", "strike", "size", "font", "color", "fill",
  "align", "valign", "wrap", "numFmt", "border", "border_color", "indent",
]);

export function columnNumber(letters) {
  let value = 0;
  for (const character of String(letters).toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value;
}

export function columnLetter(index) {
  let remaining = Number(index);
  let letters = "";
  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + modulo) + letters;
    remaining = Math.floor((remaining - modulo) / 26);
  }
  return letters;
}

/** Normalise #rgb / rrggbb / aarrggbb into an ExcelJS ARGB string. */
export function normalizeColor(value) {
  if (value === null || value === undefined || value === "") return null;
  const raw = String(value).trim().replace(/^#/, "").toUpperCase();
  if (raw === "NONE" || raw === "TRANSPARENT" || raw === "NOFILL") return null;
  if (/^[0-9A-F]{6}$/.test(raw)) return `FF${raw}`;
  if (/^[0-9A-F]{8}$/.test(raw)) return raw;
  throw new Error(`Invalid colour "${value}": use 6-digit hex (0F766E) or 8-digit ARGB (FF0F766E)`);
}

export function parseA1(reference) {
  const match = a1Pattern.exec(String(reference ?? "").trim().toUpperCase());
  if (!match) throw new Error(`Cell address must use A1 notation (for example B7), got: ${reference}`);
  return { col: columnNumber(match[1]), row: Number(match[2]) };
}

export function parseRange(reference) {
  const match = rangePattern.exec(String(reference ?? "").trim().toUpperCase());
  if (!match) throw new Error(`Range must use A1:B2 notation, got: ${reference}`);
  const first = { col: columnNumber(match[1]), row: Number(match[2]) };
  const second = { col: columnNumber(match[3]), row: Number(match[4]) };
  return {
    start: { col: Math.min(first.col, second.col), row: Math.min(first.row, second.row) },
    end: { col: Math.max(first.col, second.col), row: Math.max(first.row, second.row) },
  };
}

export function rangeAddress({ start, end }) {
  return `${columnLetter(start.col)}${start.row}:${columnLetter(end.col)}${end.row}`;
}

export function resolveNumberFormat(value, currency = "IDR") {
  const key = String(value ?? "").trim();
  const preset = numFmtPresets.get(key.toLowerCase());
  if (preset === "currency") {
    const digits = new Intl.NumberFormat("en", { style: "currency", currency })
      .resolvedOptions().maximumFractionDigits;
    const decimals = digits ? `.${"0".repeat(digits)}` : "";
    return `#,##0${decimals};[Red](#,##0${decimals});"-"`;
  }
  if (preset) return preset;
  if (!key) throw new Error("numFmt cannot be empty");
  return key;
}

function resolveBorder(value, colorValue) {
  if (value === null || value === undefined) return null;
  if (value === false || String(value).toLowerCase() === "none") return "none";
  const spec = typeof value === "string" ? { style: value, color: colorValue } : { ...value, color: value.color ?? colorValue };
  const style = String(spec.style ?? "thin").toLowerCase();
  const allowed = ["thin", "medium", "thick", "dotted", "dashed", "double", "hair", "mediumdashed", "dashdot"];
  if (!allowed.includes(style)) throw new Error(`Unsupported border style: ${style}`);
  const color = normalizeColor(spec.color) ?? "FFCBD5E1";
  return { style, color };
}

function applyBorder(cell, value) {
  const edge = { style: value.style, color: { argb: value.color } };
  cell.border = { top: edge, left: edge, bottom: edge, right: edge };
}

function applyFont(cell, format) {
  const current = { ...(cell.font ?? {}) };
  const next = { ...current };
  if (format.font !== undefined) next.name = String(format.font);
  if (format.size !== undefined) {
    const size = Number(format.size);
    if (!Number.isFinite(size) || size < 6 || size > 96) throw new Error("Font size must be between 6 and 96");
    next.size = size;
  }
  if (format.bold !== undefined) next.bold = Boolean(format.bold);
  if (format.italic !== undefined) next.italic = Boolean(format.italic);
  if (format.strike !== undefined) next.strike = Boolean(format.strike);
  if (format.underline !== undefined) next.underline = format.underline ? "single" : false;
  if (format.color !== undefined) {
    const color = normalizeColor(format.color);
    next.color = color ? { argb: color } : undefined;
  }
  if (JSON.stringify(next) !== JSON.stringify(current)) cell.font = next;
}

function applyAlignment(cell, format) {
  const current = { ...(cell.alignment ?? {}) };
  const next = { ...current };
  if (format.align !== undefined) {
    const value = String(format.align).toLowerCase();
    const map = { left: "left", center: "center", centre: "center", right: "right", justify: "justify" };
    if (!map[value]) throw new Error("align must be left, center, right, or justify");
    next.horizontal = map[value];
  }
  if (format.valign !== undefined) {
    const value = String(format.valign).toLowerCase();
    const map = { top: "top", middle: "middle", center: "middle", centre: "middle", bottom: "bottom" };
    if (!map[value]) throw new Error("valign must be top, middle, or bottom");
    next.vertical = map[value];
  }
  if (format.wrap !== undefined) next.wrapText = Boolean(format.wrap);
  if (format.indent !== undefined) {
    const indent = Number(format.indent);
    if (!Number.isFinite(indent) || indent < 0 || indent > 20) throw new Error("indent must be 0-20");
    next.indent = indent;
  }
  if (JSON.stringify(next) !== JSON.stringify(current)) cell.alignment = next;
}

function applyFormat(cell, format, context) {
  if (!format || typeof format !== "object") throw new Error("format must be an object");
  for (const key of Object.keys(format)) {
    if (!formatKeys.has(key)) {
      throw new Error(`Unknown format key "${key}". Allowed: ${[...formatKeys].join(", ")}`);
    }
  }
  applyFont(cell, format);
  applyAlignment(cell, format);
  if (format.fill !== undefined) {
    const fill = normalizeColor(format.fill);
    cell.fill = fill
      ? { type: "pattern", pattern: "solid", fgColor: { argb: fill } }
      : { type: "pattern", pattern: "none" };
  }
  if (format.numFmt !== undefined) cell.numFmt = resolveNumberFormat(format.numFmt, context.currency);
  if (format.border !== undefined) {
    const resolved = resolveBorder(format.border, format.border_color);
    if (resolved === "none") cell.border = {};
    else if (resolved) applyBorder(cell, resolved);
  }
}

function validateValue(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new Error("Cell number must be finite");
    return value;
  }
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    if (value.length > 20_000) throw new Error("Cell text exceeds the 20,000 character limit");
    return value;
  }
  throw new Error("Cell values must be strings, numbers, booleans, or null. Use `formula` for formulas.");
}

function validateFormula(formula) {
  const text = String(formula ?? "").trim().replace(/^=/, "");
  if (!text) throw new Error("formula cannot be empty");
  if (text.length > 2000) throw new Error("Formula exceeds the 2000 character limit");
  return text;
}

/**
 * Apply one update item to a worksheet.
 *
 * @returns {Array<{cell: string, kind: string}>} touched addresses, for verification
 */
export function applyUpdate(worksheet, update, context = {}) {
  const { currency = "IDR", maxRows = 5000, maxColumns = 50 } = context;
  if (!update || typeof update !== "object") throw new Error("Each update must be an object");
  const touched = [];

  if (typeof update.column === "string" && update.cell === undefined && update.range === undefined) {
    const index = columnNumber(update.column.trim().toUpperCase());
    if (index > maxColumns) throw new Error(`Column ${update.column} is outside the safe editing width`);
    const column = worksheet.getColumn(index);
    if (update.width !== undefined) {
      const width = Number(update.width);
      if (!Number.isFinite(width) || width < 0 || width > 255) throw new Error("Column width must be 0-255");
      column.width = width;
    }
    if (update.hidden !== undefined) column.hidden = Boolean(update.hidden);
    if (update.format !== undefined) {
      for (let row = 1; row <= Math.max(worksheet.rowCount, 1); row += 1) {
        applyFormat(worksheet.getCell(row, index), update.format, { currency });
      }
    }
    return [{ cell: `column ${columnLetter(index)}`, kind: "column" }];
  }

  if (update.row !== undefined && update.cell === undefined && update.range === undefined) {
    const index = Number(update.row);
    if (!Number.isInteger(index) || index < 1 || index > maxRows) throw new Error("Row must be a positive integer");
    const row = worksheet.getRow(index);
    if (update.height !== undefined) {
      const height = Number(update.height);
      if (!Number.isFinite(height) || height < 0 || height > 409) throw new Error("Row height must be 0-409");
      row.height = height;
    }
    if (update.format !== undefined) {
      for (let column = 1; column <= Math.max(worksheet.columnCount, 1); column += 1) {
        applyFormat(row.getCell(column), update.format, { currency });
      }
    }
    return [{ cell: `row ${index}`, kind: "row" }];
  }

  const hasRange = typeof update.range === "string";
  const hasCell = typeof update.cell === "string";
  if (!hasRange && !hasCell) {
    throw new Error("Each update needs `cell`, `range`, `column`, or `row`");
  }
  const bounds = hasRange ? parseRange(update.range) : { start: parseA1(update.cell), end: parseA1(update.cell) };
  if (bounds.end.row > maxRows || bounds.end.col > maxColumns) {
    throw new Error("Update is outside the safe editing dimensions (5000 rows x 50 columns)");
  }

  if (update.merge !== undefined) {
    if (!hasRange) throw new Error("merge needs a `range`, for example A1:H1");
    const address = rangeAddress(bounds);
    if (update.merge) {
      // ExcelJS keeps only the top-left value of a merged block. Move a single
      // stray value into the master cell and warn, so a merge never loses data
      // without the agent (and the user) being told.
      const master = worksheet.getCell(bounds.start.row, bounds.start.col);
      const strays = [];
      for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
        for (let col = bounds.start.col; col <= bounds.end.col; col += 1) {
          if (row === bounds.start.row && col === bounds.start.col) continue;
          const cell = worksheet.getCell(row, col);
          if (cell.value !== null && cell.value !== undefined) strays.push(cell);
        }
      }
      const emptyMaster = master.value === null || master.value === undefined;
      if (emptyMaster && strays.length === 1) master.value = strays[0].value;
      let mergeNote = null;
      try {
        worksheet.mergeCells(address);
      } catch (error) {
        // Only on failure do we pay for the merge list; a plain merge stays cheap.
        const overlapping = (worksheet.model?.merges ?? []).find((merge) => rangesOverlap(merge, address));
        if (overlapping !== address) {
          throw new Error(
            `${address} cannot be merged: it overlaps the existing merge ${overlapping ?? "(unknown)"}. `
            + `Unmerge that range first (merge: false on it) or pick a free row.`,
          );
        }
        mergeNote = `${address} is already merged; nothing changed.`;
      }
      if (!mergeNote && strays.length) {
        mergeNote = `Merging ${address} keeps only the top-left cell (${master.address}); `
          + `${strays.length} other cell(s) in the range were cleared.`;
      }
      touched.push(mergeNote ? { cell: address, kind: "merge", warning: mergeNote } : { cell: address, kind: "merge" });
    } else {
      worksheet.unMergeCells(address);
      touched.push({ cell: address, kind: "merge" });
    }
  }

  if (update.style_from !== undefined) {
    const source = worksheet.getCell(String(update.style_from));
    const style = structuredClone(source.style ?? {});
    for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
      for (let col = bounds.start.col; col <= bounds.end.col; col += 1) {
        worksheet.getCell(row, col).style = structuredClone(style);
      }
    }
    touched.push({ cell: rangeAddress(bounds), kind: "style" });
  }

  if (update.format !== undefined) {
    for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
      for (let col = bounds.start.col; col <= bounds.end.col; col += 1) {
        applyFormat(worksheet.getCell(row, col), update.format, { currency });
      }
    }
    touched.push({ cell: rangeAddress(bounds), kind: "format" });
  }

  if (update.height !== undefined && hasRange) {
    for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
      worksheet.getRow(row).height = Number(update.height);
    }
    touched.push({ cell: rangeAddress(bounds), kind: "height" });
  }

  if (update.clear === true) {
    for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
      for (let col = bounds.start.col; col <= bounds.end.col; col += 1) {
        worksheet.getCell(row, col).value = null;
      }
    }
    touched.push({ cell: rangeAddress(bounds), kind: "clear" });
  }

  const target = worksheet.getCell(bounds.start.row, bounds.start.col);
  // A merged block only shows its top-left cell, so a write to a slave cell would
  // silently vanish. Redirect it to the master and say so.
  const writeTarget = target.isMerged && target.master && target.master.address !== target.address
    ? target.master
    : target;
  if (writeTarget !== target) {
    touched.push({
      cell: target.address,
      kind: "merge",
      warning: `${target.address} is inside the merge ${writeTarget.address}; the value was written to ${writeTarget.address}.`,
    });
  }
  if (update.formula !== undefined) {
    writeTarget.value = { formula: validateFormula(update.formula) };
    touched.push({ cell: writeTarget.address, kind: "formula" });
  } else if (update.value !== undefined) {
    writeTarget.value = validateValue(update.value);
    touched.push({ cell: writeTarget.address, kind: "value" });
  }

  return touched;
}

function rangesOverlap(first, second) {
  try {
    const a = parseRange(first);
    const b = parseRange(second);
    return a.start.row <= b.end.row && b.start.row <= a.end.row
      && a.start.col <= b.end.col && b.start.col <= a.end.col;
  } catch {
    return false;
  }
}

/** Insert blank rows at `at`, inheriting the look of the row above. */
export function insertRows(worksheet, { at, count = 1, styleFromRow = null }) {
  const start = Number(at);
  const amount = Number(count);
  if (!Number.isInteger(start) || start < 1) throw new Error("`at` must be a positive row number");
  if (!Number.isInteger(amount) || amount < 1 || amount > 200) throw new Error("count must be 1-200");
  const sourceRow = styleFromRow === null || styleFromRow === undefined
    ? Math.max(1, start - 1)
    : Number(styleFromRow);
  worksheet.spliceRows(start, 0, ...Array.from({ length: amount }, () => []));
  const source = worksheet.getRow(sourceRow);
  for (let offset = 0; offset < amount; offset += 1) {
    const row = worksheet.getRow(start + offset);
    if (source.height !== undefined) row.height = source.height;
    source.eachCell({ includeEmpty: true }, (cell, column) => {
      row.getCell(column).style = structuredClone(cell.style ?? {});
    });
  }
  return { inserted: amount, at: start, from_row: sourceRow };
}

/** Delete rows starting at `at`. */
export function deleteRows(worksheet, { at, count = 1 }) {
  const start = Number(at);
  const amount = Number(count);
  if (!Number.isInteger(start) || start < 1) throw new Error("`at` must be a positive row number");
  if (!Number.isInteger(amount) || amount < 1 || amount > 200) throw new Error("count must be 1-200");
  if (start + amount - 1 > worksheet.rowCount) {
    throw new Error(`Cannot delete rows ${start}-${start + amount - 1}; the sheet only has ${worksheet.rowCount} rows`);
  }
  worksheet.spliceRows(start, amount);
  return { deleted: amount, at: start };
}

/** Insert blank columns at `at`, inheriting the look of the column to the left. */
export function insertColumns(worksheet, { at, count = 1, styleFromColumn = null }) {
  const start = Number(at);
  const amount = Number(count);
  if (!Number.isInteger(start) || start < 1) throw new Error("`at` must be a positive column number");
  if (!Number.isInteger(amount) || amount < 1 || amount > 20) throw new Error("count must be 1-20");
  const sourceIndex = styleFromColumn === null || styleFromColumn === undefined
    ? Math.max(1, start - 1)
    : Number(styleFromColumn);
  worksheet.spliceColumns(start, 0, ...Array.from({ length: amount }, () => []));
  const source = worksheet.getColumn(sourceIndex);
  for (let offset = 0; offset < amount; offset += 1) {
    const column = worksheet.getColumn(start + offset);
    column.width = source.width;
    if (source.hidden !== undefined) column.hidden = source.hidden;
    source.eachCell({ includeEmpty: true }, (cell, row) => {
      worksheet.getCell(row, start + offset).style = structuredClone(cell.style ?? {});
    });
  }
  return { inserted: amount, at: start, from_column: sourceIndex };
}

/** Delete columns starting at `at`. */
export function deleteColumns(worksheet, { at, count = 1 }) {
  const start = Number(at);
  const amount = Number(count);
  if (!Number.isInteger(start) || start < 1) throw new Error("`at` must be a positive column number");
  if (!Number.isInteger(amount) || amount < 1 || amount > 20) throw new Error("count must be 1-20");
  if (start + amount - 1 > worksheet.columnCount) {
    throw new Error(`Cannot delete columns ${start}-${start + amount - 1}; the sheet only has ${worksheet.columnCount} columns`);
  }
  worksheet.spliceColumns(start, amount);
  return { deleted: amount, at: start };
}

/** Read back a cell as a compact, JSON-safe description for verification. */
export function describeCell(cell) {
  const value = cell.value;
  const described = { cell: cell.address };
  if (value && typeof value === "object" && !Array.isArray(value)) {
    if (Object.hasOwn(value, "formula") || Object.hasOwn(value, "sharedFormula")) {
      described.formula = value.formula ?? value.sharedFormula;
      if (value.result !== undefined && value.result !== null) described.result = value.result;
    } else if (Array.isArray(value.richText)) {
      described.value = value.richText.map((run) => run.text).join("");
    } else {
      described.value = null;
    }
  } else {
    described.value = value ?? null;
  }
  const style = cell.style ?? {};
  if (style.numFmt) described.numFmt = style.numFmt;
  if (style.font?.bold) described.bold = true;
  if (style.font?.color?.argb) described.color = style.font.color.argb;
  if (style.fill?.fgColor?.argb) described.fill = style.fill.fgColor.argb;
  if (style.alignment?.horizontal) described.align = style.alignment.horizontal;
  return described;
}

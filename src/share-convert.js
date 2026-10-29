import ExcelJS from "exceljs";

const maxRows = 5000;
const maxColumns = 64;

/** Extract the first (or named) sheet of an .xlsx into a workspace dataset:
 *  { title, columns, rows }.
 *
 *  The sheet keeps its own coordinates: row 1 of Excel is row 1 of the dataset
 *  and `columns` only holds Excel letters. A title block or a header row is
 *  therefore ordinary data — the web grid draws it, the XLSX export writes it
 *  back unchanged. Pulling the first row out into `columns` (the old behaviour)
 *  hid that header on the web while the export re-created it. */
export async function convertWorkbookToDataset(bytes, { sheetName } = {}) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(bytes);
  } catch (error) {
    throw new Error(`File cannot be read as an Excel workbook: ${error.message}`);
  }
  const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  if (!worksheet) throw new Error(`Worksheet "${sheetName}" not found in the workbook`);

  const lastRow = Math.min(worksheet.rowCount || 0, maxRows);
  if (!lastRow) throw new Error("The worksheet is empty");

  const rows = [];
  let width = 0;
  for (let r = 1; r <= lastRow; r += 1) {
    const row = worksheet.getRow(r);
    const values = [];
    let lastUsed = 0;
    for (let c = 1; c <= maxColumns; c += 1) {
      const text = cellText(row.getCell(c));
      values.push(text);
      if (text !== "") lastUsed = c;
    }
    if (lastUsed > width) width = lastUsed;
    rows.push(values.slice(0, lastUsed));
  }

  // Drop trailing all-empty rows but never shift the rows above them: a total
  // block at the bottom must stay on the row its formulas point at.
  while (rows.length && !rows[rows.length - 1].some((value) => value !== "")) rows.pop();
  if (!rows.length) throw new Error("The worksheet is empty");

  width = Math.max(width, 1);
  const columns = Array.from({ length: width }, (_, index) => columnLetter(index));
  for (const row of rows) while (row.length < width) row.push("");

  return { title: worksheet.name || "Dokumen", columns, rows };
}

function columnLetter(index) {
  let name = "";
  let n = index + 1;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function cellText(cell) {
  const value = cell.value;
  if (value === null || value === undefined) return "";
  // ExcelJS reports the merged value in every covered cell; only the master
  // cell owns it, so a merged title is stored once instead of five times.
  if (cell.isMerged && cell.master && cell.master.address !== cell.address) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    // ExcelJS rich text / formula results.
    if (typeof value.text === "string") return value.text;
    if (value.richText) return value.richText.map((part) => part.text).join("");
    if (value.result !== undefined && value.result !== null) return String(value.result);
    if (value.formula) return ""; // uncached formula: leave blank, never leak the expression
    return "";
  }
  return String(value);
}

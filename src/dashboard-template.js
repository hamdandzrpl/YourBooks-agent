/** Server-owned summary sheets for web workspaces.
 *
 *  The user's reported failure — "minta transaksi, malah semua terhapus" — was
 *  the agent rewriting the whole workbook with `share_link_set_rows` and, along
 *  the way, inventing the Dashboard layout (or dropping it). The reliable
 *  division of labour is the opposite:
 *
 *    - the agent writes INPUT sheets only (Pemasukan, Pengeluaran, ...);
 *    - THIS module builds the summary sheets (Dashboard, Laba Rugi Bulanan)
 *      deterministically from the input sheets' own columns;
 *    - the agent reads the sheets' column letters with `share_link_describe`
 *      and never has to guess a formula target again.
 *
 *  Rules used here mirror the `dashboard-workbooks` skill contract, so a
 *  generated Dashboard is by construction readable by `extractSummary`:
 *  KPI labels "Total Pemasukan" / "Total Pengeluaran" / "Laba Bersih" /
 *  "Margin", a component table with a text dimension column, and only
 *  cross-sheet formulas — never typed numbers.
 */

import { columnLetter } from "./formula-engine.js";

const INCOME = /pemasukan|pendapatan|penerimaan|income|terima/i;
const EXPENSE = /pengeluaran|biaya|belanja|expense|keluar|bayar/i;
const SUMMARY = /^(dashboard|ringkasan|summary|ikhtisar|dasbor)\b/i;
const MONTHLY = /laba rugi|rugi laba|bulanan|monthly/i;
const TOTAL = /total|jumlah|nominal|amount|nilai/i;
const DATEISH = /tanggal|date|bulan|periode|waktu|jam|tgl|tahun/i;
const DIMENSION = /komoditas|komponen|kategori|jenis|uraian|deskripsi|nama|pembeli|pemasok|klien|barang|produk/i;

const MAX_DIMENSIONS = 12;
const MAX_MONTHS = 24;

/** Wrap a sheet name in single quotes when Excel would otherwise reject it. */
function sheetRef(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name}'`;
}

function hasHeader(sheet) {
  return sheet?.headerRow === true || Array.isArray(sheet?.columns);
}

function headerLabels(sheet) {
  if (!hasHeader(sheet)) return [];
  return (sheet.rows?.[0] ?? []).map((cell) =>
    String(cell && typeof cell === "object" ? cell.value ?? "" : cell ?? "").trim());
}

function dataRowsOf(sheet) {
  return hasHeader(sheet) ? (sheet.rows ?? []).slice(1) : (sheet.rows ?? []);
}

function cellValue(cell) {
  return cell && typeof cell === "object" && "value" in cell ? cell.value : cell;
}

/** First column whose header reads like the money column. */
function totalColumn(sheet) {
  const headers = headerLabels(sheet);
  if (!headers.length) return -1;
  const index = headers.findIndex((label) => TOTAL.test(label) && !DATEISH.test(label));
  return index >= 0 ? index : -1;
}

/** The dimension column: the first non-date, non-total text column whose header
 *  names the thing being tracked. When `prefer` is given and that header
 *  exists in this sheet, it wins — that is how the expense side discovers it
 *  should also use "Komoditas" rather than "Kategori Biaya" to match the
 *  income side. */
function dimensionColumn(sheet, exclude = -1, prefer = "") {
  const headers = headerLabels(sheet);
  if (!headers.length) return -1;
  if (prefer) {
    const match = headers.findIndex((label) => label.toLowerCase() === prefer.toLowerCase());
    if (match >= 0 && match !== exclude) return match;
  }
  let fallback = -1;
  for (let index = 0; index < headers.length; index++) {
    const label = headers[index];
    if (!label || index === exclude) continue;
    if (DATEISH.test(label) || TOTAL.test(label)) continue;
    const hasText = (dataRowsOf(sheet) ?? []).some((row) => {
      const value = cellValue(row?.[index]);
      return typeof value === "string" && value.trim() && !value.startsWith("#");
    });
    if (!hasText) continue;
    if (DIMENSION.test(label)) return index;
    if (fallback < 0) fallback = index;
  }
  return fallback;
}

function distinctValues(sheet, column, cap) {
  const seen = [];
  const map = new Map();
  for (const row of dataRowsOf(sheet)) {
    const value = cellValue(row?.[column]);
    const text = typeof value === "string" ? value.trim() : "";
    if (!text || map.has(text)) continue;
    map.set(text, true);
    seen.push(text);
    if (seen.length >= cap) break;
  }
  return seen;
}

/** SUM over one whole column of a sheet, or 0 when there is no such sheet/column. */
function colSum(sheet, column) {
  if (!sheet || column < 0) return "0";
  return `SUM(${sheetRef(sheet.name)}!${columnLetter(column)}:${columnLetter(column)})`;
}

/** SUMIFS over one column keyed on exact text equality with the given cell. */
function colSumIf(sheet, totalColumn, dimColumn, anchor) {
  if (!sheet || dimColumn < 0) return "0";
  if (totalColumn < 0) return "0";
  return `SUMIFS(${sheetRef(sheet.name)}!${columnLetter(totalColumn)}:${columnLetter(totalColumn)},`
    + `${sheetRef(sheet.name)}!${columnLetter(dimColumn)}:${columnLetter(dimColumn)},${anchor})`;
}

function styleMap(rows) {
  const styles = {};
  // Header: bold, dark fill, white text, centred.
  const header = rows[0] ?? [];
  for (let col = 0; col < header.length; col++) {
    styles[`0,${col}`] = { fill: "0F766E", color: "FFFFFF", bold: true, align: "center" };
  }
  // The value column reads right-aligned.
  for (let r = 1; r < rows.length; r++) styles[`${r},1`] = { align: "right" };
  return styles;
}

function dashboardSheet(income, expense, warnings) {
  const incomeTotal = income ? totalColumn(income) : -1;
  const expenseTotal = expense ? totalColumn(expense) : -1;
  const incomeDim = income ? dimensionColumn(income, incomeTotal) : -1;
  const expenseDim = expense
    ? dimensionColumn(expense, expenseTotal, incomeDim >= 0 ? headerLabels(income)[incomeDim] : "")
    : -1;
  if (income && incomeTotal < 0) warnings.push(`Sheet "${income.name}" tidak punya kolom Total/Nominal/Jumlah; tambahkan kolom itu agar pemasukan terhitung.`);
  if (expense && expenseTotal < 0) warnings.push(`Sheet "${expense.name}" tidak punya kolom Total/Nominal/Jumlah; tambahkan kolom itu agar pengeluaran terhitung.`);

  const rows = [["Komponen", "Nilai"]];
  const incomeRef = `B${rows.length + 1}`;
  rows.push(["Total Pemasukan", { formula: colSum(income, incomeTotal) }]);
  const expenseRef = `B${rows.length + 1}`;
  rows.push(["Total Pengeluaran", { formula: colSum(expense, expenseTotal) }]);
  const profitRef = `B${rows.length + 1}`;
  rows.push(["Laba Bersih", { formula: `${incomeRef}-${expenseRef}` }]);
  rows.push(["Margin", { formula: `IF(${incomeRef}>0,${profitRef}/${incomeRef},0)` }]);
  rows.push([]);

  // Component table: every distinct dimension, with formulas only.
  const dims = [];
  const seen = new Set();
  for (const sheet of [income, expense]) {
    const dim = sheet === income ? incomeDim : expenseDim;
    for (const value of sheet && dim >= 0 ? distinctValues(sheet, dim, MAX_DIMENSIONS) : []) {
      if (seen.has(value)) continue;
      seen.add(value);
      dims.push({ value, incomeDim: sheet === income ? dim : -1, expenseDim: sheet === expense ? dim : -1 });
      if (dims.length >= MAX_DIMENSIONS) break;
    }
  }
  rows.push(["Komponen", "Total Pemasukan", "Total Biaya", "Laba Bersih", "Margin"]);
  for (const dim of dims) {
    const r = rows.length + 1;
    const anchor = `A${r}`;
    rows.push([
      dim.value,
      { formula: colSumIf(income, incomeTotal, dim.incomeDim, anchor) },
      { formula: colSumIf(expense, expenseTotal, dim.expenseDim, anchor) },
      { formula: `B${r}-C${r}` },
      { formula: `IF(B${r}>0,D${r}/B${r},0)` },
    ]);
  }
  if (!dims.length) {
    warnings.push(`Belum ada kolom dimensi (Komoditas/Kategori/Uraian) yang terisi di sheet input, jadi tabel komponen di Dashboard masih kosong.`);
  }

  return {
    name: "Dashboard",
    columns: ["A", "B", "C", "D", "E"],
    headerRow: true,
    rows,
    styles: styleMap(rows),
  };
}

/** Monthly statement — only when both input sheets carry a "Bulan" column. */
function monthlySheet(income, expense) {
  const bulanOf = (sheet) => (sheet
    ? headerLabels(sheet).findIndex((label) => label.toLowerCase() === "bulan")
    : -1);
  const incomeBulan = bulanOf(income);
  const expenseBulan = bulanOf(expense);
  if (incomeBulan < 0 || expenseBulan < 0) return null;
  const incomeTotal = income ? totalColumn(income) : -1;
  const expenseTotal = expense ? totalColumn(expense) : -1;
  if (incomeTotal < 0 || expenseTotal < 0) return null;

  const months = [];
  const seen = new Set();
  for (const sheet of [income, expense]) {
    const column = sheet === income ? incomeBulan : expenseBulan;
    for (const month of distinctValues(sheet, column, MAX_MONTHS)) {
      if (seen.has(month)) continue;
      seen.add(month);
      months.push(month);
    }
  }
  months.sort();
  if (!months.length) return null;

  const rows = [["Periode", "Pemasukan", "Pengeluaran", "Laba Bersih", "Kumulatif"]];
  for (const month of months) {
    const r = rows.length + 1;
    const anchor = `A${r}`;
    const incomeFormula = colSumIf(income, incomeTotal, incomeBulan, anchor);
    const expenseFormula = colSumIf(expense, expenseTotal, expenseBulan, anchor);
    const cumulative = r === 2 ? `D${r}` : `D${r}+E${r - 1}`;
    rows.push([
      month,
      { formula: incomeFormula },
      { formula: expenseFormula },
      { formula: `B${r}-C${r}` },
      { formula: cumulative },
    ]);
  }
  return {
    name: "Laba Rugi Bulanan",
    columns: ["A", "B", "C", "D", "E"],
    headerRow: true,
    rows,
    styles: styleMap(rows),
  };
}

/**
 * Given the workbook's sheets, return the sheet list with the summary sheets
 * rebuilt from the input sheets, plus any structural warnings.
 *
 * Summary sheets present in the input (from an agent or an import) are
 * replaced: their content is derived state, and letting both the server and
 * the agent own it is how a write once came to delete everything.
 */
export function applySummarySheets(sheets) {
  const list = Array.isArray(sheets) ? sheets : [];
  const income = list.find((sheet) => INCOME.test(String(sheet.name ?? "").trim()));
  const expense = list.find((sheet) => EXPENSE.test(String(sheet.name ?? "").trim()));
  // No recognised roles at all: the document is a legacy single-sheet or a
  // one-off list. Keeping it as-is preserves backward compatibility — the
  // agent can still use share_link_set_rows on those without a dashboard.
  if (!income && !expense) return { sheets: list, warnings: [], generated: false };

  const input = list.filter((sheet) => !SUMMARY.test(String(sheet.name ?? "").trim()) && !MONTHLY.test(String(sheet.name ?? "").trim()));
  const warnings = [];
  const built = [dashboardSheet(income, expense, warnings)];
  for (const sheet of input) built.push(sheet);
  const monthly = monthlySheet(income, expense);
  if (monthly) built.push(monthly);

  return { sheets: built, warnings, generated: true };
}
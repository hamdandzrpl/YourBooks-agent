// Professional finance report renderer.
//
// Replaces the old "write a bare values table through Excel MCP" path. Everything
// is authored natively with ExcelJS so the output has real structure: KPI cards,
// live formulas, Excel tables with filters, conditional formatting, data bars,
// data validation, print setup and native charts.
//
// The ledger (SQLite) stays the single source of truth; this module only renders.
import ExcelJS from "exceljs";
import { injectCharts, sheetRef } from "./xlsx/charts.js";
import { MONTH_SHORT, formatAmount, formatPercent, monthLabel, monthLabelLong } from "./analytics.js";

export const MANAGED_SHEETS = ["Ringkasan", "Arus Kas", "Kategori", "Transaksi", "Anggaran", "Anomali", "Petunjuk"];
export const LEDGER_SHEET = "Transaksi";

const SHEET_ALIASES = {
  ringkasan: "Ringkasan",
  "arus-kas": "Arus Kas",
  "arus kas": "Arus Kas",
  kategori: "Kategori",
  transaksi: "Transaksi",
  anggaran: "Anggaran",
  anomali: "Anomali",
  petunjuk: "Petunjuk",
};

/**
 * Normalise a requested sheet list into canonical names.
 * Unknown names are rejected so a typo never silently drops a section.
 * @param {string[]|undefined} sheets
 * @returns {string[]} canonical sheet names, in canonical order
 */
export function normalizeSheets(sheets) {
  if (sheets === undefined || sheets === null) return [...MANAGED_SHEETS];
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw new Error("sheets must be a non-empty array of section names");
  }
  const wanted = new Set();
  for (const entry of sheets) {
    const key = String(entry).trim().toLocaleLowerCase("id-ID");
    const canonical = SHEET_ALIASES[key];
    if (!canonical) {
      throw new Error(`Unknown report section "${entry}". Valid: ${MANAGED_SHEETS.join(", ")}`);
    }
    wanted.add(canonical);
  }
  // The budget sheet is driven by SUMIFS over the ledger sheet, so it cannot
  // exist on its own.
  if (wanted.has("Anggaran")) wanted.add("Transaksi");
  return MANAGED_SHEETS.filter((name) => wanted.has(name));
}

const CREATOR = "WhatsApp Finance Agent";

const THEMES = {
  emerald: {
    ink: "FF10231F", deep: "FF0B3B33", accent: "FF0F766E", accentSoft: "FFD7EAE6",
    paper: "FFF6F8F7", line: "FFDCE6E3", muted: "FF64736F", white: "FFFFFFFF",
    positive: "FF127C56", negative: "FFB3341F", warn: "FFB7791F", gold: "FFB08D57",
    palette: ["0F766E", "B08D57", "3B6E8F", "8A5A96", "B3341F", "4E8C6A", "7A7F8C", "C77B3C"],
  },
};

const DISPLAY_FONT = "Calibri Light";
const BODY_FONT = "Calibri";

function amountFormat(currency) {
  const digits = new Intl.NumberFormat("en", { style: "currency", currency })
    .resolvedOptions().maximumFractionDigits;
  const decimals = digits ? `.${"0".repeat(digits)}` : "";
  return `#,##0${decimals};[Red](#,##0${decimals});"-"`;
}

function argb(hex) {
  return hex.startsWith("FF") && hex.length === 8 ? hex : `FF${hex}`;
}

function solid(theme, key) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: theme[key] } };
}

function solidHex(hex) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: argb(hex) } };
}

function border(theme, { bottom, top, left, right } = {}) {
  const thin = { style: "thin", color: { argb: theme.line } };
  const medium = { style: "medium", color: { argb: theme.accent } };
  return {
    top: top === "medium" ? medium : top ? thin : undefined,
    bottom: bottom === "medium" ? medium : bottom ? thin : undefined,
    left: left ? thin : undefined,
    right: right ? thin : undefined,
  };
}

function pageSetup(sheet, { landscape = true } = {}) {
  sheet.pageSetup = {
    paperSize: 9,
    orientation: landscape ? "landscape" : "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
  };
  sheet.headerFooter = {
    oddFooter: `&L${CREATOR}&R Halaman &P / &N`,
  };
}

function titleBand(sheet, theme, { title, subtitle, width = 8 }) {
  const lastColumn = String.fromCharCode(64 + width);
  sheet.mergeCells(`A1:${lastColumn}1`);
  sheet.mergeCells(`A2:${lastColumn}2`);
  const titleCell = sheet.getCell("A1");
  titleCell.value = title;
  titleCell.font = { name: DISPLAY_FONT, size: 22, bold: true, color: { argb: theme.white } };
  titleCell.fill = solid(theme, "deep");
  titleCell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  sheet.getRow(1).height = 44;

  const subtitleCell = sheet.getCell("A2");
  subtitleCell.value = subtitle;
  subtitleCell.font = { name: BODY_FONT, size: 10, color: { argb: theme.white } };
  subtitleCell.fill = solid(theme, "accent");
  subtitleCell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  sheet.getRow(2).height = 22;
}

function sectionHeader(sheet, theme, row, label, { from = "A", to = "H" } = {}) {
  sheet.mergeCells(`${from}${row}:${to}${row}`);
  const cell = sheet.getCell(`${from}${row}`);
  cell.value = label;
  cell.font = { name: BODY_FONT, size: 11, bold: true, color: { argb: theme.ink } };
  cell.alignment = { vertical: "middle", indent: 0 };
  cell.border = { bottom: { style: "medium", color: { argb: theme.accent } } };
  sheet.getRow(row).height = 24;
}

function tableHeader(sheet, theme, row, labels, { from = 1 } = {}) {
  labels.forEach((label, index) => {
    const cell = sheet.getCell(row, from + index);
    cell.value = label;
    cell.font = { name: BODY_FONT, size: 9, bold: true, color: { argb: theme.white } };
    cell.fill = solid(theme, "accent");
    cell.alignment = { vertical: "middle", horizontal: index === 0 ? "left" : "right", wrapText: true };
    cell.border = { bottom: { style: "thin", color: { argb: theme.deep } } };
  });
  sheet.getRow(row).height = 30;
}

function noteCell(cell, theme, text) {
  cell.value = text;
  cell.font = { name: BODY_FONT, size: 9, italic: true, color: { argb: theme.muted } };
  cell.alignment = { vertical: "top", wrapText: true };
}

const levelStyles = {
  good: { label: "BAIK", key: "positive" },
  warn: { label: "PERHATIAN", key: "warn" },
  danger: { label: "RISIKO", key: "negative" },
  info: { label: "INFO", key: "muted" },
};

/* ------------------------------------------------------------------ Ringkasan */

function buildSummarySheet(workbook, analysis, theme, { hasCashFlow = true, title = null } = {}) {
  const sheet = workbook.addWorksheet("Ringkasan", { properties: { tabColor: theme.deep } });
  sheet.views = [{ state: "normal", showGridLines: false }];
  sheet.columns = Array.from({ length: 8 }, (_, index) => ({ width: index === 0 ? 30 : 17 }));
  pageSetup(sheet, { landscape: true });

  const periodText = analysis.period.from
    ? `${monthLabelLong(analysis.period.from.slice(0, 7))} – ${monthLabelLong(analysis.period.to.slice(0, 7))}`
    : "Belum ada transaksi";
  titleBand(sheet, theme, {
    title: title || "LAPORAN KEUANGAN",
    subtitle: `${analysis.bookName || "Buku"}   •   ${periodText}   •   ${analysis.currency}`,
  });

  const money = (value) => formatAmount(value, { currency: analysis.currency, digits: analysis.digits });

  // KPI cards ---------------------------------------------------------------
  // Values are live formulas over the Arus Kas sheet when a period exists, so a
  // user editing the budget sheet still sees a consistent dashboard.
  const monthCount = analysis.months.length;
  const totalRow = monthCount ? 5 + monthCount : 0;
  const cashRef = (column) => (hasCashFlow && monthCount
    ? `'Arus Kas'!${column}${totalRow}`
    : null);
  const cards = [
    { label: "TOTAL PEMASUKAN", value: analysis.totals.income, color: theme.positive, format: "money", formula: cashRef("B") },
    { label: "TOTAL PENGELUARAN", value: analysis.totals.expense, color: theme.negative, format: "money", formula: cashRef("C") },
    { label: "ARUS KAS BERSIH", value: analysis.totals.net, color: analysis.totals.net >= 0 ? theme.accent : theme.negative, format: "money", formula: cashRef("D") },
    { label: "RASIO TABUNGAN", value: (analysis.totals.savingsRate ?? 0) / 100, color: theme.gold, format: "percent", formula: cashRef("E") },
  ];
  const format = amountFormat(analysis.currency);
  cards.forEach((card, index) => {
    const startColumn = index * 2 + 1;
    const labelCell = sheet.getCell(4, startColumn);
    const valueCell = sheet.getCell(5, startColumn);
    sheet.mergeCells(4, startColumn, 4, startColumn + 1);
    sheet.mergeCells(5, startColumn, 6, startColumn + 1);

    labelCell.value = card.label;
    labelCell.font = { name: BODY_FONT, size: 8, bold: true, color: { argb: theme.muted } };
    labelCell.fill = solid(theme, "paper");
    labelCell.alignment = { vertical: "middle", indent: 1 };

    valueCell.value = card.formula
      ? { formula: card.formula, result: card.value }
      : card.value;
    valueCell.font = { name: DISPLAY_FONT, size: 16, bold: true, color: { argb: card.color } };
    valueCell.fill = solid(theme, "paper");
    valueCell.numFmt = card.format === "percent" ? '0.0%;[Red](0.0%);"-"' : format;
    valueCell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };

    for (let row = 4; row <= 6; row += 1) {
      for (let column = startColumn; column <= startColumn + 1; column += 1) {
        sheet.getCell(row, column).border = {
          top: row === 4 ? { style: "thin", color: { argb: theme.accent } } : undefined,
          bottom: row === 6 ? { style: "thin", color: { argb: theme.line } } : undefined,
          left: column === startColumn ? { style: "thin", color: { argb: theme.line } } : undefined,
          right: column === startColumn + 1 ? { style: "thin", color: { argb: theme.line } } : undefined,
        };
      }
    }
  });
  sheet.getRow(4).height = 20;
  sheet.getRow(5).height = 26;
  sheet.getRow(6).height = 18;

  // Narrative insights ------------------------------------------------------
  let row = 8;
  sectionHeader(sheet, theme, row, "SOROTAN OTOMATIS");
  row += 1;
  for (const insight of analysis.insights.slice(0, 9)) {
    const style = levelStyles[insight.level] ?? levelStyles.info;
    const badge = sheet.getCell(row, 1);
    badge.value = style.label;
    badge.font = { name: BODY_FONT, size: 8, bold: true, color: { argb: theme.white } };
    badge.fill = solid(theme, style.key);
    badge.alignment = { vertical: "middle", horizontal: "center" };

    sheet.mergeCells(row, 2, row, 8);
    const text = sheet.getCell(row, 2);
    text.value = `${insight.title} — ${insight.detail}`;
    text.font = { name: BODY_FONT, size: 9, color: { argb: theme.ink } };
    text.alignment = { vertical: "middle", wrapText: true, indent: 1 };
    sheet.getRow(row).height = 30;
    row += 1;
  }

  // Category overview -------------------------------------------------------
  row += 1;
  sectionHeader(sheet, theme, row, "KATEGORI PENGELUARAN TERATAS");
  row += 1;
  const categoryStart = row;
  tableHeader(sheet, theme, row, ["Kategori", "Total", "Porsi", "Transaksi", "Rata-rata", "Visual"], { from: 1 });
  sheet.mergeCells(row, 6, row, 8);
  row += 1;
  const topCategories = analysis.expenseCategories.slice(0, 8);
  const maxCategory = topCategories[0]?.total ?? 1;
  for (const category of topCategories) {
    sheet.getCell(row, 1).value = category.name;
    sheet.getCell(row, 1).font = { name: BODY_FONT, size: 9, color: { argb: theme.ink } };
    sheet.getCell(row, 2).value = category.total;
    sheet.getCell(row, 2).numFmt = format;
    sheet.getCell(row, 3).value = (category.share ?? 0) / 100;
    sheet.getCell(row, 3).numFmt = "0.0%";
    sheet.getCell(row, 4).value = category.count;
    sheet.getCell(row, 4).numFmt = "#,##0";
    sheet.getCell(row, 5).value = category.average;
    sheet.getCell(row, 5).numFmt = format;
    sheet.mergeCells(row, 6, row, 8);
    sheet.getCell(row, 6).value = "█".repeat(Math.max(1, Math.round((category.total / maxCategory) * 24)));
    sheet.getCell(row, 6).font = { name: BODY_FONT, size: 9, color: { argb: theme.accent } };
    for (let column = 1; column <= 8; column += 1) {
      sheet.getCell(row, column).border = border(theme, { bottom: true });
      sheet.getCell(row, column).alignment = {
        vertical: "middle",
        horizontal: column === 1 ? "left" : column === 6 ? "left" : "right",
      };
    }
    row += 1;
  }
  if (!topCategories.length) {
    noteCell(sheet.getCell(row, 1), theme, "Belum ada pengeluaran pada periode ini.");
    sheet.mergeCells(row, 1, row, 8);
    row += 1;
  }
  const categoryEnd = row - 1;
  if (topCategories.length) {
    sheet.addConditionalFormatting({
      ref: `C${categoryStart + 1}:C${categoryEnd}`,
      rules: [{
        type: "dataBar",
        cfvo: [{ type: "num", value: 0 }, { type: "num", value: 1 }],
        color: { argb: theme.accent },
        showValue: true,
      }],
    });
  }

  // Footer note -------------------------------------------------------------
  row += 1;
  sheet.mergeCells(row, 1, row, 8);
  noteCell(sheet.getCell(row, 1), theme, "Angka dihitung dari ledger transaksi aktif. Transaksi batal (void) tidak diikutkan. Grafik di sebelah kanan mengikuti periode yang sama.");
  sheet.getRow(row).height = 28;

  // Chart columns (J..R) keep the dashboard readable without crowding text.
  for (let column = 9; column <= 18; column += 1) sheet.getColumn(column).width = 9;
  sheet.getColumn(9).width = 3;

  return sheet;
}

/* ------------------------------------------------------------------ Arus Kas */

function buildCashFlowSheet(workbook, analysis, theme) {
  const sheet = workbook.addWorksheet("Arus Kas", { properties: { tabColor: theme.accent } });
  sheet.views = [{ state: "frozen", ySplit: 4, showGridLines: false }];
  sheet.columns = [{ width: 26 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 15 }, { width: 18 }, { width: 3 }];
  pageSetup(sheet);

  titleBand(sheet, theme, {
    title: "ARUS KAS BULANAN",
    subtitle: `${analysis.bookName || "Buku"}   •   ${analysis.currency}   •   ${analysis.months.length} bulan tercatat`,
  });

  const format = amountFormat(analysis.currency);
  tableHeader(sheet, theme, 4, ["Bulan", "Pemasukan", "Pengeluaran", "Arus Bersih", "Rasio Tabungan", "Kumulatif"]);

  let row = 5;
  let cumulative = 0;
  for (const month of analysis.months) {
    cumulative += month.net;
    sheet.getCell(row, 1).value = monthLabelLong(month.key);
    sheet.getCell(row, 2).value = month.income;
    sheet.getCell(row, 3).value = month.expense;
    sheet.getCell(row, 4).value = { formula: `B${row}-C${row}`, result: month.net };
    sheet.getCell(row, 5).value = { formula: `IFERROR(D${row}/B${row},0)`, result: (month.savingsRate ?? 0) / 100 };
    sheet.getCell(row, 6).value = { formula: `SUM($D$5:D${row})`, result: cumulative };

    for (const column of [2, 3, 4, 6]) sheet.getCell(row, column).numFmt = format;
    sheet.getCell(row, 5).numFmt = '0.0%;[Red](0.0%);"-"';
    for (let column = 1; column <= 6; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.font = { name: BODY_FONT, size: 10, color: { argb: theme.ink } };
      cell.border = border(theme, { bottom: true });
      cell.alignment = { vertical: "middle", horizontal: column === 1 ? "left" : "right", indent: column === 1 ? 1 : 0 };
    }
    sheet.getCell(row, 1).font = { name: BODY_FONT, size: 10, bold: true, color: { argb: theme.ink } };
    sheet.getRow(row).height = 22;
    row += 1;
  }

  const firstDataRow = 5;
  const lastDataRow = Math.max(firstDataRow, row - 1);
  if (analysis.months.length) {
    // Totals row
    sheet.getCell(row, 1).value = "TOTAL";
    sheet.getCell(row, 2).value = { formula: `SUM(B${firstDataRow}:B${lastDataRow})`, result: analysis.totals.income };
    sheet.getCell(row, 3).value = { formula: `SUM(C${firstDataRow}:C${lastDataRow})`, result: analysis.totals.expense };
    sheet.getCell(row, 4).value = { formula: `B${row}-C${row}`, result: analysis.totals.net };
    sheet.getCell(row, 5).value = { formula: `IFERROR(D${row}/B${row},0)`, result: (analysis.totals.savingsRate ?? 0) / 100 };
    for (const column of [2, 3, 4]) sheet.getCell(row, column).numFmt = format;
    sheet.getCell(row, 5).numFmt = '0.0%;[Red](0.0%);"-"';
    for (let column = 1; column <= 6; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.font = { name: BODY_FONT, size: 10, bold: true, color: { argb: theme.white } };
      cell.fill = solid(theme, "deep");
      cell.border = { top: { style: "thin", color: { argb: theme.deep } } };
      cell.alignment = { vertical: "middle", horizontal: column === 1 ? "left" : "right", indent: column === 1 ? 1 : 0 };
    }
    sheet.getRow(row).height = 24;

    sheet.addConditionalFormatting({
      ref: `D${firstDataRow}:D${lastDataRow}`,
      rules: [{
        type: "colorScale",
        cfvo: [{ type: "min" }, { type: "num", value: 0 }, { type: "max" }],
        color: [{ argb: theme.negative }, { argb: "FFFFFFFF" }, { argb: theme.positive }],
      }],
    });
    sheet.addConditionalFormatting({
      ref: `B${firstDataRow}:B${lastDataRow}`,
      rules: [{
        type: "dataBar",
        cfvo: [{ type: "num", value: 0 }, { type: "max" }],
        color: { argb: theme.positive },
        showValue: true,
      }],
    });
    sheet.addConditionalFormatting({
      ref: `C${firstDataRow}:C${lastDataRow}`,
      rules: [{
        type: "dataBar",
        cfvo: [{ type: "num", value: 0 }, { type: "max" }],
        color: { argb: theme.negative },
        showValue: true,
      }],
    });
    row += 2;
  } else {
    noteCell(sheet.getCell(5, 1), theme, "Belum ada transaksi. Catat transaksi pertama lalu buka laporan ini lagi.");
    sheet.mergeCells(5, 1, 5, 6);
    row = 7;
  }

  sectionHeader(sheet, theme, row, "RINGKASAN PERIODE", { to: "F" });
  row += 1;
  const summaryRows = [
    ["Jumlah transaksi", analysis.totals.count, "#,##0"],
    ["Hari aktif", analysis.totals.activeDays, "#,##0"],
    ["Rata-rata per transaksi", analysis.totals.avgTicket, format],
    ["Pengeluaran rata-rata per hari aktif", analysis.totals.burnRate, format],
    ["Transaksi terbesar", analysis.totals.largestExpense, format],
    ["Median pengeluaran", analysis.totals.medianExpense, format],
  ];
  for (const [label, value, numFmt] of summaryRows) {
    sheet.getCell(row, 1).value = label;
    sheet.getCell(row, 1).font = { name: BODY_FONT, size: 9, color: { argb: theme.muted } };
    sheet.getCell(row, 2).value = value;
    sheet.getCell(row, 2).numFmt = numFmt;
    sheet.getCell(row, 2).font = { name: BODY_FONT, size: 10, bold: true, color: { argb: theme.ink } };
    sheet.getCell(row, 2).alignment = { horizontal: "right" };
    sheet.getRow(row).height = 18;
    row += 1;
  }

  return sheet;
}

/* ------------------------------------------------------------------- Kategori */

function buildCategorySheet(workbook, analysis, theme) {
  const sheet = workbook.addWorksheet("Kategori", { properties: { tabColor: theme.gold } });
  sheet.views = [{ state: "frozen", ySplit: 4, showGridLines: false }];
  sheet.columns = [{ width: 28 }, { width: 14 }, { width: 18 }, { width: 14 }, { width: 18 }, { width: 12 }, { width: 14 }];
  pageSetup(sheet);

  titleBand(sheet, theme, {
    title: "ANALISIS KATEGORI",
    subtitle: `${analysis.bookName || "Buku"}   •   ${analysis.currency}   •   diurutkan dari nominal terbesar`,
  });

  const format = amountFormat(analysis.currency);
  tableHeader(sheet, theme, 4, ["Kategori", "Jenis", "Total", "Transaksi", "Rata-rata", "Porsi", "Bulan Aktif"]);

  const rows = [
    ...analysis.expenseCategories.map((entry) => ({ ...entry, kindLabel: "Pengeluaran" })),
    ...analysis.incomeCategories.map((entry) => ({ ...entry, kindLabel: "Pemasukan" })),
  ];
  let row = 5;
  for (const entry of rows) {
    sheet.getCell(row, 1).value = entry.name;
    sheet.getCell(row, 2).value = entry.kindLabel;
    sheet.getCell(row, 3).value = entry.total;
    sheet.getCell(row, 4).value = entry.count;
    sheet.getCell(row, 5).value = entry.average;
    sheet.getCell(row, 6).value = (entry.share ?? 0) / 100;
    sheet.getCell(row, 7).value = entry.monthsActive;

    sheet.getCell(row, 3).numFmt = format;
    sheet.getCell(row, 5).numFmt = format;
    sheet.getCell(row, 4).numFmt = "#,##0";
    sheet.getCell(row, 6).numFmt = "0.0%";
    sheet.getCell(row, 7).numFmt = "#,##0";
    for (let column = 1; column <= 7; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.font = { name: BODY_FONT, size: 10, color: { argb: theme.ink } };
      cell.border = border(theme, { bottom: true });
      cell.alignment = { vertical: "middle", horizontal: column === 1 || column === 2 ? "left" : "right", indent: column === 1 ? 1 : 0 };
    }
    sheet.getCell(row, 2).font = {
      name: BODY_FONT, size: 9,
      color: { argb: entry.kind === "income" ? theme.positive : theme.negative },
    };
    sheet.getRow(row).height = 20;
    row += 1;
  }
  if (!rows.length) {
    noteCell(sheet.getCell(5, 1), theme, "Belum ada kategori pada periode ini.");
    sheet.mergeCells(5, 1, 5, 7);
    row = 6;
  }
  const lastDataRow = Math.max(5, row - 1);
  if (rows.length) {
    sheet.getCell(row, 1).value = "TOTAL";
    sheet.getCell(row, 3).value = { formula: `SUM(C5:C${lastDataRow})`, result: analysis.totals.income + analysis.totals.expense };
    sheet.getCell(row, 4).value = { formula: `SUM(D5:D${lastDataRow})`, result: analysis.totals.count };
    sheet.getCell(row, 3).numFmt = format;
    for (let column = 1; column <= 7; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.font = { name: BODY_FONT, size: 10, bold: true, color: { argb: theme.white } };
      cell.fill = solid(theme, "deep");
      cell.alignment = { vertical: "middle", horizontal: column === 1 ? "left" : "right", indent: column === 1 ? 1 : 0 };
    }
    sheet.getRow(row).height = 22;
    sheet.addConditionalFormatting({
      ref: `C5:C${lastDataRow}`,
      rules: [{
        type: "dataBar",
        cfvo: [{ type: "num", value: 0 }, { type: "max" }],
        color: { argb: theme.accent },
        showValue: true,
      }],
    });
  }
  sheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: lastDataRow, column: 7 } };

  return sheet;
}

/* ------------------------------------------------------------------ Transaksi */

function buildLedgerSheet(workbook, analysis, theme, transactions) {
  const sheet = workbook.addWorksheet(LEDGER_SHEET, { properties: { tabColor: theme.ink } });
  const format = amountFormat(analysis.currency);
  const headers = ["ID", "Tanggal", "Jenis", "Kategori", "Keterangan", "Pemasukan", "Pengeluaran", "Mata Uang"];
  const rows = transactions.map((transaction) => [
    transaction.id,
    new Date(`${transaction.occurred_on}T00:00:00.000Z`),
    transaction.kind === "income" ? "Pemasukan" : "Pengeluaran",
    transaction.category,
    transaction.description,
    transaction.kind === "income" ? Number(transaction.amount_minor) / (10 ** analysis.digits) : null,
    transaction.kind === "expense" ? Number(transaction.amount_minor) / (10 ** analysis.digits) : null,
    transaction.currency,
  ]);

  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.columns = [
    { width: 16 }, { width: 13 }, { width: 13 }, { width: 24 }, { width: 46 }, { width: 17 }, { width: 17 }, { width: 11 },
  ];
  pageSetup(sheet);

  sheet.addTable({
    name: "BukuTransaksi",
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    style: { theme: "TableStyleMedium2", showRowStripes: true },
    columns: headers.map((name) => ({ name, filterButton: true })),
    rows: rows.length ? rows : [[null, null, null, null, null, null, null, null]],
  });

  const dataRowCount = Math.max(1, rows.length);
  const lastRow = dataRowCount + 1;
  sheet.getColumn(2).numFmt = "dd mmm yyyy";
  sheet.getColumn(6).numFmt = format;
  sheet.getColumn(7).numFmt = format;
  sheet.getColumn(5).alignment = { wrapText: true, vertical: "top" };

  sheet.dataValidations.add(`C2:C${lastRow + 200}`, {
    type: "list", allowBlank: true, formulae: ['"Pemasukan,Pengeluaran"'],
  });
  sheet.dataValidations.add(`F2:G${lastRow + 200}`, {
    type: "decimal", operator: "greaterThanOrEqual", allowBlank: true, formulae: [0],
  });
  // The Excel table already provides header filters; a second sheet-level
  // autoFilter makes Excel ask to repair the file.
  sheet.getColumn(1).font = { name: BODY_FONT, size: 9, color: { argb: theme.muted } };

  return { sheet, lastRow };
}

/* ------------------------------------------------------------------- Anggaran */

function buildBudgetSheet(workbook, analysis, theme, ledgerLastRow) {
  const sheet = workbook.addWorksheet("Anggaran", { properties: { tabColor: "FFB08D57" } });
  sheet.views = [{ state: "frozen", ySplit: 4, showGridLines: false }];
  sheet.columns = [{ width: 30 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 14 }, { width: 16 }];
  pageSetup(sheet);

  titleBand(sheet, theme, {
    title: "ANGGARAN VS REALISASI",
    subtitle: "Isi kolom Anggaran (kuning) — realisasi, sisa, dan status menghitung otomatis",
    width: 6,
  });

  const format = amountFormat(analysis.currency);
  tableHeader(sheet, theme, 4, ["Kategori", "Anggaran", "Realisasi", "Sisa", "% Terpakai", "Status"]);

  const limit = Math.max(ledgerLastRow + 200, 500);
  const categories = analysis.expenseCategories.map((entry) => entry.name);
  const lineCount = Math.max(15, categories.length);
  let row = 5;
  for (let index = 0; index < lineCount; index += 1) {
    const category = categories[index] ?? null;
    const realization = analysis.expenseCategories.find((entry) => entry.name === category)?.total ?? 0;
    sheet.getCell(row, 1).value = category;
    sheet.getCell(row, 2).value = null;
    sheet.getCell(row, 3).value = {
      formula: category
        ? `SUMIFS(${sheetRef(LEDGER_SHEET, `$G$2:$G$${limit}`)},${sheetRef(LEDGER_SHEET, `$D$2:$D$${limit}`)},$A${row})`
        : "",
      result: category ? realization : "",
    };
    sheet.getCell(row, 4).value = { formula: `IF($B${row}="","",$B${row}-$C${row})`, result: "" };
    sheet.getCell(row, 5).value = { formula: `IFERROR($C${row}/$B${row},"")`, result: "" };
    sheet.getCell(row, 6).value = {
      formula: `IF($B${row}="","",IF($C${row}>$B${row},"Lewat anggaran",IF($C${row}>$B${row}*0.8,"Hampir habis","Aman")))`,
      result: "",
    };

    for (const column of [2, 3, 4]) sheet.getCell(row, column).numFmt = format;
    sheet.getCell(row, 5).numFmt = "0.0%";
    for (let column = 1; column <= 6; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.font = { name: BODY_FONT, size: 10, color: { argb: theme.ink } };
      cell.border = border(theme, { bottom: true });
      cell.alignment = { vertical: "middle", horizontal: column === 1 || column === 6 ? "left" : "right", indent: column === 1 ? 1 : 0 };
    }
    sheet.getCell(row, 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF8E1" } };
    sheet.getCell(row, 2).font = { name: BODY_FONT, size: 10, bold: true, color: { argb: "FF8A6D1F" } };
    sheet.getRow(row).height = 20;
    row += 1;
  }
  const firstDataRow = 5;
  const lastDataRow = row - 1;

  sheet.getCell(row, 1).value = "TOTAL";
  sheet.getCell(row, 2).value = { formula: `SUM(B${firstDataRow}:B${lastDataRow})`, result: 0 };
  sheet.getCell(row, 3).value = { formula: `SUM(C${firstDataRow}:C${lastDataRow})`, result: analysis.totals.expense };
  sheet.getCell(row, 4).value = { formula: `B${row}-C${row}`, result: 0 };
  sheet.getCell(row, 5).value = { formula: `IFERROR(C${row}/B${row},"")`, result: "" };
  for (const column of [2, 3, 4]) sheet.getCell(row, column).numFmt = format;
  sheet.getCell(row, 5).numFmt = "0.0%";
  for (let column = 1; column <= 6; column += 1) {
    const cell = sheet.getCell(row, column);
    cell.font = { name: BODY_FONT, size: 10, bold: true, color: { argb: theme.white } };
    cell.fill = solid(theme, "deep");
  }
  sheet.getRow(row).height = 22;

  sheet.addConditionalFormatting({
    ref: `E${firstDataRow}:E${lastDataRow}`,
    rules: [
      { type: "cellIs", operator: "greaterThan", formulae: [1], style: { font: { color: { argb: theme.negative }, bold: true } } },
      { type: "cellIs", operator: "between", formulae: [0.8, 1], style: { font: { color: { argb: theme.warn }, bold: true } } },
    ],
  });
  sheet.addConditionalFormatting({
    ref: `D${firstDataRow}:D${lastDataRow}`,
    rules: [{ type: "cellIs", operator: "lessThan", formulae: [0], style: { font: { color: { argb: theme.negative } } } }],
  });
  sheet.dataValidations.add(`B${firstDataRow}:B${lastDataRow}`, {
    type: "decimal", operator: "greaterThanOrEqual", allowBlank: true, formulae: [0],
    showErrorMessage: true, errorTitle: "Anggaran tidak valid", error: "Masukkan angka 0 atau lebih.",
  });

  return sheet;
}

/* -------------------------------------------------------------------- Anomali */

function buildAnomalySheet(workbook, analysis, theme) {
  const sheet = workbook.addWorksheet("Anomali", { properties: { tabColor: theme.negative } });
  sheet.views = [{ state: "frozen", ySplit: 4, showGridLines: false }];
  sheet.columns = [{ width: 13 }, { width: 46 }, { width: 22 }, { width: 18 }, { width: 14 }, { width: 16 }];
  pageSetup(sheet);

  titleBand(sheet, theme, {
    title: "PEMERIKSAAN ANOMALI",
    subtitle: `Transaksi dengan nominal ≥ 3× median pengeluaran (median ${formatAmount(analysis.totals.medianExpense, { currency: analysis.currency, digits: analysis.digits })})`,
    width: 6,
  });

  const format = amountFormat(analysis.currency);
  tableHeader(sheet, theme, 4, ["Tanggal", "Keterangan", "Kategori", "Nominal", "× Median", "Tindak lanjut"]);

  let row = 5;
  for (const anomaly of analysis.anomalies) {
    sheet.getCell(row, 1).value = new Date(`${anomaly.date}T00:00:00.000Z`);
    sheet.getCell(row, 1).numFmt = "dd mmm yyyy";
    sheet.getCell(row, 2).value = anomaly.description || "(tanpa keterangan)";
    sheet.getCell(row, 3).value = anomaly.category;
    sheet.getCell(row, 4).value = anomaly.amount;
    sheet.getCell(row, 4).numFmt = format;
    sheet.getCell(row, 5).value = anomaly.ratio;
    sheet.getCell(row, 5).numFmt = '0.0"×"';
    sheet.getCell(row, 6).value = "Periksa bukti transaksi";
    for (let column = 1; column <= 6; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.font = { name: BODY_FONT, size: 10, color: { argb: theme.ink } };
      cell.border = border(theme, { bottom: true });
      cell.alignment = { vertical: "middle", horizontal: column >= 4 && column <= 5 ? "right" : "left", indent: column === 2 || column === 1 ? 1 : 0 };
    }
    sheet.getRow(row).height = 20;
    row += 1;
  }

  if (!analysis.anomalies.length) {
    sheet.mergeCells(5, 1, 5, 6);
    noteCell(sheet.getCell(5, 1), theme, "Tidak ada transaksi yang menonjol. Semua nominal pengeluaran berada dalam rentang wajar.");
    sheet.getRow(5).height = 24;
    row = 6;
  }

  row += 1;
  sectionHeader(sheet, theme, row, "SEBARAN HARI DALAM SEMINGGU", { to: "F" });
  row += 1;
  tableHeader(sheet, theme, row, ["Hari", "Total", "Transaksi", "Rata-rata", "", ""]);
  row += 1;
  const maxWeekday = Math.max(...analysis.weekdays.map((entry) => entry.total), 1);
  for (const weekday of analysis.weekdays) {
    sheet.getCell(row, 1).value = weekday.name;
    sheet.getCell(row, 2).value = weekday.total;
    sheet.getCell(row, 2).numFmt = format;
    sheet.getCell(row, 3).value = weekday.count;
    sheet.getCell(row, 4).value = weekday.count ? weekday.total / weekday.count : 0;
    sheet.getCell(row, 4).numFmt = format;
    sheet.mergeCells(row, 5, row, 6);
    sheet.getCell(row, 5).value = "█".repeat(Math.max(0, Math.round((weekday.total / maxWeekday) * 18)));
    sheet.getCell(row, 5).font = { name: BODY_FONT, size: 9, color: { argb: theme.gold } };
    for (let column = 1; column <= 6; column += 1) {
      sheet.getCell(row, column).border = border(theme, { bottom: true });
      sheet.getCell(row, column).alignment = { vertical: "middle", horizontal: column <= 4 ? "right" : "left" };
    }
    sheet.getCell(row, 1).alignment = { horizontal: "left", indent: 1 };
    row += 1;
  }

  return sheet;
}

/* ------------------------------------------------------------------- Petunjuk */

function buildGuideSheet(workbook, analysis, theme) {
  const sheet = workbook.addWorksheet("Petunjuk", { properties: { tabColor: theme.muted } });
  sheet.views = [{ state: "normal", showGridLines: false }];
  sheet.columns = [{ width: 3 }, { width: 30 }, { width: 88 }];
  pageSetup(sheet, { landscape: false });

  sheet.mergeCells("B2:C2");
  const title = sheet.getCell("B2");
  title.value = "CARA MEMBACA LAPORAN INI";
  title.font = { name: DISPLAY_FONT, size: 16, bold: true, color: { argb: theme.ink } };

  const sections = [
    ["Ringkasan", "Kartu KPI (pemasukan, pengeluaran, arus bersih, rasio tabungan), sorotan otomatis, kategori teratas, dan grafik periode."],
    ["Arus Kas", "Tren bulanan. Kolom Kumulatif menunjukkan akumulasi arus kas sejak transaksi pertama. Warna merah menandakan bulan defisit."],
    ["Kategori", "Semua kategori pemasukan dan pengeluaran, lengkap dengan porsi terhadap jenisnya dan rata-rata per transaksi."],
    ["Transaksi", "Ledger mentah. Sheet ini dikelola otomatis dari database: tambah, ubah, atau batalkan lewat chat, bukan dengan mengedit sel."],
    ["Anggaran", "Isi kolom kuning dengan plafon belanja per kategori. Realisasi, sisa, dan status dihitung otomatis dari ledger."],
    ["Anomali", "Transaksi yang nominalnya jauh di atas kebiasaan, plus sebaran belanja per hari dalam seminggu."],
  ];
  let row = 4;
  for (const [name, description] of sections) {
    sheet.getCell(row, 2).value = name;
    sheet.getCell(row, 2).font = { name: BODY_FONT, size: 10, bold: true, color: { argb: theme.accent } };
    sheet.getCell(row, 2).alignment = { vertical: "top", wrapText: true };
    sheet.getCell(row, 3).value = description;
    sheet.getCell(row, 3).font = { name: BODY_FONT, size: 10, color: { argb: theme.ink } };
    sheet.getCell(row, 3).alignment = { vertical: "top", wrapText: true };
    sheet.getRow(row).height = 34;
    row += 1;
  }

  row += 1;
  sheet.mergeCells(row, 2, row, 3);
  const method = sheet.getCell(row, 2);
  method.value = "METODE PERHITUNGAN";
  method.font = { name: BODY_FONT, size: 11, bold: true, color: { argb: theme.ink } };
  method.border = { bottom: { style: "medium", color: { argb: theme.accent } } };
  row += 1;

  const notes = [
    ["Sumber angka", "Seluruh angka berasal dari ledger transaksi aktif pada database. Transaksi yang dibatalkan (void) tetap tersimpan untuk jejak audit tetapi tidak dihitung."],
    ["Rasio tabungan", "(Pemasukan − Pengeluaran) ÷ Pemasukan. Di bawah 10% dianggap tipis, di bawah 0% berarti defisit."],
    ["Arus bersih", "Pemasukan − Pengeluaran pada periode yang sama."],
    ["Median", "Nilai tengah pengeluaran, dipakai sebagai pembanding wajar karena tidak terpengaruh nominal ekstrem."],
    ["Anomali", "Pengeluaran yang nilainya minimal 3× median dan minimal 100.000 pada mata uang utama."],
    ["Mata uang", `Laporan ini hanya menjumlahkan transaksi dalam ${analysis.currency}.${analysis.otherCurrencies.length ? ` Transaksi dalam ${analysis.otherCurrencies.join(", ")} tidak digabung karena kurs tidak diasumsikan.` : ""}`],
    ["Perbarui data", "Catat, ubah, atau batalkan transaksi lewat chat WhatsApp. Workbook ini tersinkron otomatis setelah setiap perubahan."],
    ["Batasan", "Anggaran bersifat perencanaan manual; laporan ini bukan laporan keuangan yang diaudit."],
  ];
  for (const [label, text] of notes) {
    sheet.getCell(row, 2).value = label;
    sheet.getCell(row, 2).font = { name: BODY_FONT, size: 10, bold: true, color: { argb: theme.muted } };
    sheet.getCell(row, 2).alignment = { vertical: "top" };
    sheet.getCell(row, 3).value = text;
    sheet.getCell(row, 3).font = { name: BODY_FONT, size: 10, color: { argb: theme.ink } };
    sheet.getCell(row, 3).alignment = { vertical: "top", wrapText: true };
    sheet.getRow(row).height = 30;
    row += 1;
  }

  row += 1;
  sheet.mergeCells(row, 2, row, 3);
  noteCell(sheet.getCell(row, 2), theme, `Dibuat otomatis oleh ${CREATOR} pada ${new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" })}.`);

  return sheet;
}

/* --------------------------------------------------------------------- charts */

function chartSpecs(analysis, ledgerLastRow, enabled = new Set(MANAGED_SHEETS)) {
  if (!analysis.months.length) return [];
  const specs = [];
  const monthLabels = analysis.months.map((entry) => MONTH_SHORT[Number(entry.key.slice(5, 7)) - 1] + " " + entry.key.slice(2, 4));
  const hasCash = enabled.has("Arus Kas");
  const hasSummary = enabled.has("Ringkasan");

  if (analysis.months.length >= 2 && hasCash && hasSummary) {
    // Charts read the Arus Kas sheet so they stay live when months are added.
    const firstRow = 5;
    const lastRow = firstRow + analysis.months.length - 1;
    specs.push({
      sheet: "Ringkasan",
      type: "bar",
      title: "Pemasukan vs Pengeluaran per Bulan",
      from: { col: 9, row: 3 },
      to: { col: 17, row: 21 },
      categories: { ref: sheetRef("Arus Kas", `$A$${firstRow}:$A$${lastRow}`), values: monthLabels },
      series: [
        {
          name: "Pemasukan",
          nameRef: sheetRef("Arus Kas", "$B$4"),
          ref: sheetRef("Arus Kas", `$B$${firstRow}:$B$${lastRow}`),
          values: analysis.months.map((entry) => entry.income),
          color: "127C56",
        },
        {
          name: "Pengeluaran",
          nameRef: sheetRef("Arus Kas", "$C$4"),
          ref: sheetRef("Arus Kas", `$C$${firstRow}:$C$${lastRow}`),
          values: analysis.months.map((entry) => entry.expense),
          color: "B3341F",
        },
      ],
    });
  }

  if (analysis.expenseCategories.length && enabled.has("Kategori") && hasSummary) {
    const top = analysis.expenseCategories.slice(0, 7);
    specs.push({
      sheet: "Ringkasan",
      type: "doughnut",
      title: "Komposisi Pengeluaran",
      from: { col: 9, row: 23 },
      to: { col: 17, row: 41 },
      palette: THEMES.emerald.palette,
      categories: {
        ref: sheetRef("Kategori", `$A$5:$A$${4 + top.length}`),
        values: top.map((entry) => entry.name),
      },
      series: [{
        name: "Pengeluaran",
        nameRef: sheetRef("Kategori", "$C$4"),
        ref: sheetRef("Kategori", `$C$5:$C$${4 + top.length}`),
        values: top.map((entry) => entry.total),
      }],
    });
  }

  if (analysis.months.length >= 2 && hasCash) {
    const firstRow = 5;
    const lastRow = firstRow + analysis.months.length - 1;
    specs.push({
      sheet: "Arus Kas",
      type: "line",
      title: "Akumulasi Arus Kas",
      from: { col: 8, row: 3 },
      to: { col: 16, row: 24 },
      categories: { ref: sheetRef("Arus Kas", `$A$${firstRow}:$A$${lastRow}`), values: monthLabels },
      series: [{
        name: "Kumulatif",
        nameRef: sheetRef("Arus Kas", "$F$4"),
        ref: sheetRef("Arus Kas", `$F$${firstRow}:$F$${lastRow}`),
        values: analysis.months.reduce((acc, entry) => {
          acc.push((acc.at(-1) ?? 0) + entry.net);
          return acc;
        }, []),
        color: "0F766E",
      }],
    });
  }

  return specs;
}

/* ------------------------------------------------------------------ public API */

/**
 * Render a complete report workbook.
 *
 * @param {object} options
 * @param {object} options.analysis result of analyzeBook()
 * @param {Array} options.transactions raw ledger rows
 * @param {ExcelJS.Workbook} [options.workbook] existing workbook to update in place
 * @param {boolean} [options.charts] include native charts (default true)
 * @param {string} [options.theme]
 * @returns {Promise<{workbook: ExcelJS.Workbook, bytes: Buffer, charts: number}>}
 */
export async function renderReport({ analysis, transactions, workbook, charts = true, theme: themeName = "emerald", verify = true, sheets, title }) {
  const theme = THEMES[themeName] ?? THEMES.emerald;
  const target = workbook ?? new ExcelJS.Workbook();
  const enabled = new Set(normalizeSheets(sheets));

  for (const name of MANAGED_SHEETS) {
    const existing = target.getWorksheet(name);
    if (existing) target.removeWorksheet(existing.id);
  }

  target.creator = CREATOR;
  target.lastModifiedBy = CREATOR;
  target.title = title || `Laporan Keuangan — ${analysis.bookName || "Buku"}`;
  target.subject = "Laporan keuangan otomatis";
  target.company = analysis.bookName || "Finance Agent";
  target.created = target.created ?? new Date();
  target.modified = new Date();

  const ledgerLastRow = transactions.length + 1;
  if (enabled.has("Ringkasan")) buildSummarySheet(target, analysis, theme, { hasCashFlow: enabled.has("Arus Kas"), title });
  if (enabled.has("Arus Kas")) buildCashFlowSheet(target, analysis, theme);
  if (enabled.has("Kategori")) buildCategorySheet(target, analysis, theme);
  if (enabled.has("Transaksi")) buildLedgerSheet(target, analysis, theme, transactions);
  if (enabled.has("Anggaran")) buildBudgetSheet(target, analysis, theme, ledgerLastRow);
  if (enabled.has("Anomali")) buildAnomalySheet(target, analysis, theme);
  if (enabled.has("Petunjuk")) buildGuideSheet(target, analysis, theme);

  const buffer = Buffer.from(await target.xlsx.writeBuffer());
  if (!charts) return { workbook: target, bytes: buffer, charts: 0, sheets: [...enabled] };

  const specs = chartSpecs(analysis, ledgerLastRow, enabled);
  if (!specs.length) return { workbook: target, bytes: buffer, charts: 0, sheets: [...enabled] };

  try {
    const withCharts = injectCharts(buffer, specs);
    if (!verify) return { workbook: target, bytes: withCharts, charts: specs.length, sheets: [...enabled] };
    // Prove the archive still parses before handing it back.
    const check = new ExcelJS.Workbook();
    await check.xlsx.load(withCharts);
    return { workbook: target, bytes: withCharts, charts: specs.length, sheets: [...enabled] };
  } catch (error) {
    console.warn(`Native charts skipped: ${error.message}`);
    return { workbook: target, bytes: buffer, charts: 0, sheets: [...enabled] };
  }
}

export { amountFormat, THEMES };

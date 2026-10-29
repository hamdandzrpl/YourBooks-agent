import { DOMParser } from "@xmldom/xmldom";
import Docxtemplater from "docxtemplater";
import ExcelJS from "exceljs";
import { Document, HeadingLevel, Packer, Paragraph } from "docx";
import PizZip from "pizzip";
import { readFile } from "node:fs/promises";
import { basename, extname, join } from "node:path";
import { preserveCharts } from "./spreadsheet-charts.js";
import { sheetRef } from "../finance/xlsx/charts.js";
import {
  applyUpdate, columnLetter, deleteColumns, deleteRows, describeCell, insertColumns,
  insertRows, parseRange, parseA1,
} from "./sheet-edit.js";
import { describeDesign, injectDesign } from "./xlsx-design.js";

const xlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const docxMimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const maxOfficeBytes = 20 * 1024 * 1024;
const maxOfficeExpandedBytes = 100 * 1024 * 1024;
const maxOfficeArchiveEntries = 1000;
const maxRows = 5000;
const maxColumns = 50;

function safeCellValue(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  if (typeof value !== "object") return value;
  if (Array.isArray(value.richText)) return value.richText.map((run) => run.text).join("");
  if (value.hyperlink) return value.text ?? value.hyperlink;
  if (Object.hasOwn(value, "result")) return safeCellValue(value.result);
  return null;
}

function validateRows(rows, columns) {
  if (!Array.isArray(rows) || rows.length > maxRows) {
    throw new Error(`Spreadsheet rows must be an array of at most ${maxRows} rows`);
  }
  if (columns !== undefined && (!Array.isArray(columns) || columns.length === 0 || columns.length > maxColumns)) {
    throw new Error(`Spreadsheet must have 1-${maxColumns} columns`);
  }
  for (const row of rows) {
    if (!Array.isArray(row) || row.length > maxColumns) {
      throw new Error("Each spreadsheet row must be an array with no more than 50 values");
    }
    if (row.some((value) => value !== null && !["string", "number", "boolean"].includes(typeof value))) {
      throw new Error("Spreadsheet values must be strings, numbers, booleans, or null");
    }
  }
}

function requireExtension(file, extension) {
  if (extname(file.name).toLowerCase() !== extension) {
    throw new Error(`File ${file.name} is not a supported ${extension} document`);
  }
  if (file.sizeBytes > maxOfficeBytes) throw new Error("Office document exceeds the 20 MB processing limit");
}

function outputName(value, fallback) {
  const name = String(value || fallback).trim();
  if (!name || name.length > 160 || name.includes("/") || name.includes("\\") || name === "." || name === "..") {
    throw new Error("Output filename must be a simple name without a path");
  }
  return name;
}

function archiveEntryPath(value) {
  const normalized = String(value).replaceAll("\\", "/");
  const trimmed = normalized.endsWith("/") ? normalized.slice(0, -1) : normalized;
  const parts = trimmed.split("/");
  if (!trimmed || normalized.startsWith("/") || /^[a-z]:/i.test(normalized)
    || parts.some((part) => !part || part === "." || part === ".." || part.startsWith(".")
      || /[<>:"|?*\u0000-\u001f]/.test(part))) {
    throw new Error("ZIP archive contains an unsafe or unsupported path");
  }
  return parts.join("/");
}

function isSymbolicLink(entry) {
  const permissions = entry.unixPermissions;
  const mode = typeof permissions === "string" ? Number.parseInt(permissions, 8) : permissions;
  return Number.isInteger(mode) && (mode & 0o170000) === 0o120000;
}

export class OfficeFiles {
  constructor(workspace, financeStore = null, defaultCurrency = "IDR") {
    this.workspace = workspace;
    this.financeStore = financeStore;
    this.defaultCurrency = defaultCurrency;
  }

  assertSpreadsheetSheetEditable(userId, fileId, sheetName) {
    if (!sheetName || sheetName.toLocaleLowerCase("en-US") !== "transaksi") return;
    const isManagedWorkbook = this.financeStore?.listBooks(userId)
      .some((book) => book.workbook_file_id === fileId);
    if (isManagedWorkbook) {
      throw new Error("The Transaksi sheet in a finance book is managed by the ledger and cannot be edited through general spreadsheet tools");
    }
  }

  /** Resolve an optional sheet name to the workbook's first sheet. */
  resolveSheetName(workbook, sheetName) {
    const name = sheetName ?? workbook.worksheets[0]?.name;
    if (!name) throw new Error("Workbook has no worksheets");
    return name;
  }

  async saveEditedSpreadsheet(userId, file, bytes, { saveMode = "in_place", outputFilename, suffix }) {
    if (!new Set(["copy", "in_place"]).has(saveMode)) {
      throw new Error("save_mode must be copy or in_place");
    }
    if (saveMode === "in_place") {
      if (outputFilename) throw new Error("output_filename can only be used when save_mode is copy");
      return this.workspace.replaceGeneratedFile(userId, file.id, bytes);
    }
    const originalBase = basename(file.name, ".xlsx");
    const filename = outputName(outputFilename, `${originalBase}-${suffix}.xlsx`);
    if (extname(filename).toLowerCase() !== ".xlsx") throw new Error("Output filename must end in .xlsx");
    // Repeated copies used to fail with "A file already exists at the requested
    // path", which the agent surfaced as a failed edit. Allocate a free name.
    const base = basename(filename, ".xlsx");
    for (let attempt = 0; attempt < 50; attempt += 1) {
      const candidate = attempt === 0 ? filename : `${base}-${attempt + 1}.xlsx`;
      try {
        return await this.workspace.saveGeneratedFile(userId, `exports/${candidate}`, bytes, xlsxMimeType);
      } catch (error) {
        if (!String(error.message).includes("already exists")) throw error;
      }
    }
    throw new Error("Could not allocate a unique export filename");
  }

  /** Load a workbook for editing, with the same archive safety checks as before. */
  async loadWorkbookForEdit(userId, fileId) {
    const file = this.workspace.getFile(userId, fileId);
    if (!file) throw new Error("Spreadsheet not found in this user's workspace");
    requireExtension(file, ".xlsx");
    const originalBytes = await readFile(await this.workspace.getFilePath(userId, fileId));
    readSafeOfficeArchive(originalBytes);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(originalBytes);
    return { file, originalBytes, workbook };
  }

  /**
   * Serialise, keep native charts alive, optionally inject drawings, then save.
   * Returns the workspace file plus what was actually verified on disk.
   */
  async commitWorkbook(userId, {
    file, originalBytes, workbook, design = null, saveMode, outputFilename, suffix, changes = [], sheetName = null,
  }) {
    if (workbook.worksheets.length === 0) throw new Error("A workbook needs at least one worksheet");
    if (workbook.worksheets.some((worksheet) => worksheet.columnCount > maxColumns)) {
      throw new Error(`Worksheet exceeds the safe editing width (${maxColumns} columns)`);
    }
    const updatedBytes = Buffer.from(await workbook.xlsx.writeBuffer());
    // ExcelJS drops charts on write; carry them over from the source file.
    let bytes = preserveCharts(originalBytes, updatedBytes).bytes;
    let injected = { shapes: 0, charts: 0 };
    if (design && (design.shapes?.length || design.charts?.length)) {
      try {
        injected = injectDesign(bytes, design);
        const check = new ExcelJS.Workbook();
        await check.xlsx.load(injected.bytes);
        bytes = injected.bytes;
      } catch (error) {
        injected = { shapes: 0, charts: 0, error: error.message };
      }
    }

    const saved = await this.saveEditedSpreadsheet(userId, file, bytes, { saveMode, outputFilename, suffix });
    const result = { ...saved };
    // `bytes` is kept for the caller's own read-back, then dropped by the tool layer.
    result.bytes = bytes;
    if (changes.length) result.verified = await this.verifyCells(bytes, sheetName, changes);
    const warnings = changes.map((entry) => entry.warning).filter(Boolean);
    if (warnings.length) result.warnings = warnings;
    if (design) {
      result.added_shapes = injected.shapes ?? 0;
      result.added_charts = injected.charts ?? 0;
      if (injected.error) result.design_error = injected.error;
      // What the sheet actually carries now, so "charts: 0 added" is never
      // misread as "this workbook has no charts".
      const summary = describeDesign(bytes);
      if (sheetName && summary.has(sheetName)) result.design = summary.get(sheetName);
    }
    return result;
  }

  /** Read the edited cells back out of the saved bytes â€” proof, not a promise. */
  async verifyCells(bytes, sheetName, changes) {
    const addresses = [...new Set(changes.filter((entry) => entry.cell && /^[A-Z]+\d+$/.test(entry.cell))
      .map((entry) => entry.cell))].slice(0, 80);
    if (!addresses.length) return [];
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(bytes);
      const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
      if (!worksheet) return [];
      return addresses.map((address) => describeCell(worksheet.getCell(address)));
    } catch {
      return [];
    }
  }

  async createSpreadsheet(userId, { filename, sheetName = "Sheet1", columns, rows = [] }) {
    validateRows(rows, columns);
    if (!/^[^\[\]*?:/\\]{1,31}$/.test(sheetName)) throw new Error("Worksheet name must be 1-31 valid characters");
    const name = outputName(filename, "spreadsheet.xlsx");
    if (extname(name).toLowerCase() !== ".xlsx") throw new Error("Spreadsheet filename must end in .xlsx");
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "WhatsApp Finance Agent";
    workbook.created = new Date();
    const worksheet = workbook.addWorksheet(sheetName);
    worksheet.addRow(columns);
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    for (const row of rows) worksheet.addRow(row);
    const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
    return this.workspace.saveGeneratedFile(userId, `exports/${name}`, bytes, xlsxMimeType);
  }

  async createFinanceWorkbook(userId, { filename, template = "cashbook", period, currency } = {}) {
    if (!this.financeStore) throw new Error("Finance workbook templates are not available");
    if (!["cashbook", "budget"].includes(template)) throw new Error("Template must be cashbook or budget");
    if (period !== undefined && !/^\d{4}-(0[1-9]|1[0-2])$/.test(period)) {
      throw new Error("Period must use YYYY-MM format");
    }
    if (currency !== undefined && !/^[A-Za-z]{3}$/.test(currency)) {
      throw new Error("Currency must be a three-letter code");
    }

    const book = this.financeStore.getActiveBook(userId);
    const periodTransactions = this.financeStore.listBookTransactions(userId, book.id)
      .filter((transaction) => !period || transaction.occurred_on.startsWith(period));
    const currencies = [...new Set(periodTransactions.map((transaction) => transaction.currency))];
    const selectedCurrency = (currency ?? currencies[0] ?? this.defaultCurrency).toUpperCase();
    const currencyDigits = new Intl.NumberFormat("en", { style: "currency", currency: selectedCurrency })
      .resolvedOptions().maximumFractionDigits;
    if (!currency && currencies.length > 1) {
      throw new Error(`Selected period contains multiple currencies (${currencies.join(", ")}); choose one currency`);
    }
    const transactions = periodTransactions.filter((transaction) => transaction.currency === selectedCurrency);
    const scale = 10 ** currencyDigits;
    const incomeTotal = transactions.filter((transaction) => transaction.kind === "income")
      .reduce((total, transaction) => total + Number(transaction.amount_minor) / scale, 0);
    const expenseTotal = transactions.filter((transaction) => transaction.kind === "expense")
      .reduce((total, transaction) => total + Number(transaction.amount_minor) / scale, 0);
    const expenseByCategory = new Map();
    const incomeByCategory = new Map();
    for (const transaction of transactions) {
      const totals = transaction.kind === "income" ? incomeByCategory : expenseByCategory;
      totals.set(transaction.category, (totals.get(transaction.category) ?? 0) + Number(transaction.amount_minor) / scale);
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "WhatsApp Finance Agent";
    workbook.lastModifiedBy = "WhatsApp Finance Agent";
    workbook.title = template === "cashbook" ? `Buku Kas - ${book.name}` : `Anggaran - ${book.name}`;
    workbook.subject = `Template keuangan ${book.name}`;
    workbook.created = new Date();
    workbook.modified = new Date();

    const colors = {
      ink: "FF1D2A32",
      teal: "FF087E78",
      pale: "FFEAF4F2",
      paper: "FFF5F7F7",
      white: "FFFFFFFF",
      muted: "FF68777D",
      green: "FF16734B",
      red: "FFB44747",
      line: "FFD6E0E0",
    };
    const amountFormat = currencyDigits === 0
      ? "#,##0;[Red](#,##0);-"
      : `#,##0.${"0".repeat(currencyDigits)};[Red](#,##0.${"0".repeat(currencyDigits)});-`;
    const formulaLimit = 10002;
    const incomeFormula = `SUMIFS('Transaksi'!$F$2:$F$${formulaLimit},'Transaksi'!$C$2:$C$${formulaLimit},"Pemasukan")`;
    const expenseFormula = `SUMIFS('Transaksi'!$F$2:$F$${formulaLimit},'Transaksi'!$C$2:$C$${formulaLimit},"Pengeluaran")`;
    const categoryNames = [...new Set(transactions
      .filter((transaction) => template !== "budget" || transaction.kind === "expense")
      .map((transaction) => transaction.category))].sort((a, b) => a.localeCompare(b));
    const transactionRows = transactions.map((transaction) => [
      transaction.id,
      new Date(`${transaction.occurred_on}T00:00:00.000Z`),
      transaction.kind === "income" ? "Pemasukan" : "Pengeluaran",
      transaction.category,
      transaction.description,
      Number(transaction.amount_minor) / scale,
      selectedCurrency,
    ]);

    const summary = workbook.addWorksheet("Ringkasan", { properties: { tabColor: colors.teal } });
    summary.mergeCells("A1:H1");
    summary.mergeCells("A2:H2");
    summary.getCell("A1").value = template === "cashbook" ? "BUKU KAS" : "ANGGARAN BULANAN";
    summary.getCell("A2").value = `${book.name}  |  ${period ?? "Semua periode"}  |  ${selectedCurrency}`;
    summary.getCell("A1").font = { name: "Aptos Display", size: 20, bold: true, color: { argb: colors.white } };
    summary.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.ink } };
    summary.getCell("A1").alignment = { vertical: "middle" };
    summary.getRow(1).height = 38;
    summary.getCell("A2").font = { name: "Aptos", size: 10, color: { argb: colors.muted } };
    summary.getRow(2).height = 24;
    summary.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
    summary.getColumn(1).width = 24;
    summary.getColumn(2).width = 20;
    summary.views = [{ state: "normal", showGridLines: false }];

    const kpis = [
      { label: "TOTAL PEMASUKAN", start: "A", end: "B", formula: incomeFormula, result: incomeTotal, color: colors.green },
      { label: "TOTAL PENGELUARAN", start: "C", end: "D", formula: expenseFormula, result: expenseTotal, color: colors.red },
      { label: "ARUS BERSIH", start: "E", end: "F", formula: "A5-C5", result: incomeTotal - expenseTotal, color: colors.teal },
      { label: "JUMLAH TRANSAKSI", start: "G", end: "H", formula: `COUNTA('Transaksi'!$A$2:$A$${formulaLimit})`, result: transactions.length, color: colors.ink },
    ];
    for (const [index, kpi] of kpis.entries()) {
      summary.mergeCells(`${kpi.start}4:${kpi.end}4`);
      summary.mergeCells(`${kpi.start}5:${kpi.end}6`);
      const labelCell = summary.getCell(`${kpi.start}4`);
      const valueCell = summary.getCell(`${kpi.start}5`);
      labelCell.value = kpi.label;
      labelCell.font = { name: "Aptos", size: 9, bold: true, color: { argb: colors.muted } };
      labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.paper } };
      labelCell.alignment = { vertical: "middle" };
      valueCell.value = { formula: kpi.formula, result: kpi.result };
      valueCell.font = { name: "Aptos Display", size: index === 3 ? 18 : 16, bold: true, color: { argb: kpi.color } };
      valueCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.pale } };
      valueCell.alignment = { vertical: "middle" };
      valueCell.numFmt = index === 3 ? "#,##0" : amountFormat;
      for (let row = 4; row <= 6; row += 1) {
        for (let column = index * 2 + 1; column <= index * 2 + 2; column += 1) {
          summary.getCell(row, column).border = { bottom: { style: "thin", color: { argb: colors.line } } };
        }
      }
    }
    summary.getRow(4).height = 22;
    summary.getRow(5).height = 28;
    summary.getRow(6).height = 18;

    const transactionSheet = workbook.addWorksheet("Transaksi", { properties: { tabColor: colors.ink } });
    const transactionHeaders = ["ID", "Tanggal", "Jenis", "Kategori", "Keterangan", "Jumlah", "Mata Uang"];
    transactionSheet.addTable({
      name: "DataTransaksi",
      ref: "A1:G1",
      headerRow: true,
      style: { theme: "TableStyleMedium2", showRowStripes: true },
      columns: transactionHeaders.map((name) => ({ name, filterButton: true })),
      rows: transactionRows.length ? transactionRows : [[null, null, null, null, null, null, null]],
    });
    transactionSheet.columns = [
      { width: 16 }, { width: 15 }, { width: 16 }, { width: 22 }, { width: 38 }, { width: 18 }, { width: 14 },
    ];
    transactionSheet.views = [{ state: "frozen", ySplit: 1 }];
    transactionSheet.getColumn(2).numFmt = "dd mmm yyyy";
    transactionSheet.getColumn(6).numFmt = amountFormat;
    transactionSheet.dataValidations.add(`C2:C${formulaLimit}`, {
      type: "list", allowBlank: true, formulae: ['"Pemasukan,Pengeluaran"'],
    });
    transactionSheet.dataValidations.add(`F2:F${formulaLimit}`, {
      type: "decimal", operator: "greaterThan", allowBlank: true, formulae: [0],
    });

    if (template === "cashbook") {
      summary.mergeCells("A9:D9");
      summary.getCell("A9").value = "RINGKASAN PER KATEGORI";
      summary.getCell("A9").font = { name: "Aptos", size: 11, bold: true, color: { argb: colors.ink } };
      summary.getCell("A9").border = { bottom: { style: "medium", color: { argb: colors.teal } } };
      const headers = ["Kategori", "Pemasukan", "Pengeluaran", "Bersih"];
      summary.addRow([]);
      summary.addRow(headers);
      const headerRow = summary.getRow(11);
      headerRow.font = { bold: true, color: { argb: colors.white } };
      headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.teal } };
      for (const [index, category] of categoryNames.entries()) {
        const rowNumber = index + 12;
        summary.getCell(`A${rowNumber}`).value = category;
        summary.getCell(`B${rowNumber}`).value = {
          formula: `SUMIFS('Transaksi'!$F$2:$F$${formulaLimit},'Transaksi'!$D$2:$D$${formulaLimit},A${rowNumber},'Transaksi'!$C$2:$C$${formulaLimit},"Pemasukan")`,
          result: incomeByCategory.get(category) ?? 0,
        };
        summary.getCell(`C${rowNumber}`).value = {
          formula: `SUMIFS('Transaksi'!$F$2:$F$${formulaLimit},'Transaksi'!$D$2:$D$${formulaLimit},A${rowNumber},'Transaksi'!$C$2:$C$${formulaLimit},"Pengeluaran")`,
          result: expenseByCategory.get(category) ?? 0,
        };
        summary.getCell(`D${rowNumber}`).value = { formula: `B${rowNumber}-C${rowNumber}`, result: (incomeByCategory.get(category) ?? 0) - (expenseByCategory.get(category) ?? 0) };
        for (const column of ["B", "C", "D"]) summary.getCell(`${column}${rowNumber}`).numFmt = amountFormat;
      }
      summary.autoFilter = categoryNames.length ? `A11:D${categoryNames.length + 11}` : undefined;
    } else {
      const budget = workbook.addWorksheet("Anggaran", { properties: { tabColor: "FFB88A45" } });
      const headers = ["Kategori", "Anggaran", "Realisasi", "Sisa Anggaran", "% Terpakai"];
      budget.addRow(headers);
      const budgetLineCount = Math.max(20, categoryNames.length);
      for (let index = 0; index < budgetLineCount; index += 1) {
        const rowNumber = index + 2;
        const category = categoryNames[index] ?? null;
        budget.addRow([category, null, {
          formula: `IF(A${rowNumber}=\"\",\"\",SUMIFS('Transaksi'!$F$2:$F$${formulaLimit},'Transaksi'!$D$2:$D$${formulaLimit},A${rowNumber},'Transaksi'!$C$2:$C$${formulaLimit},\"Pengeluaran\"))`,
          result: category ? expenseByCategory.get(category) ?? 0 : "",
        }, {
          formula: `IF(B${rowNumber}=\"\",\"\",B${rowNumber}-C${rowNumber})`,
          result: "",
        }, {
          formula: `IFERROR(C${rowNumber}/B${rowNumber},0)`,
          result: 0,
        }]);
      }
      budget.columns = [{ width: 28 }, { width: 20 }, { width: 20 }, { width: 20 }, { width: 16 }];
      budget.views = [{ state: "frozen", ySplit: 1 }];
      budget.autoFilter = `A1:E${budgetLineCount + 1}`;
      budget.getRow(1).font = { bold: true, color: { argb: colors.white } };
      budget.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.teal } };
      budget.getColumn(2).numFmt = amountFormat;
      budget.getColumn(3).numFmt = amountFormat;
      budget.getColumn(4).numFmt = amountFormat;
      budget.getColumn(5).numFmt = "0.0%;[Red](0.0%);-";
      budget.dataValidations.add(`B2:B${budgetLineCount + 1}`, {
        type: "decimal", operator: "greaterThanOrEqual", allowBlank: true, formulae: [0],
      });
      budget.addConditionalFormatting({
        ref: `D2:D${budgetLineCount + 1}`,
        rules: [{ type: "cellIs", operator: "lessThan", formulae: [0], style: { font: { color: { argb: colors.red } } } }],
      });
      summary.mergeCells("A9:D9");
      summary.getCell("A9").value = "ANGGARAN VS REALISASI";
      summary.getCell("A9").font = { bold: true, color: { argb: colors.ink } };
      summary.getCell("A9").border = { bottom: { style: "medium", color: { argb: colors.teal } } };
      summary.getCell("A10").value = "Anggaran";
      summary.getCell("B10").value = { formula: `SUM('Anggaran'!$B$2:$B$${budgetLineCount + 1})`, result: 0 };
      summary.getCell("A11").value = "Realisasi";
      summary.getCell("B11").value = { formula: `SUM('Anggaran'!$C$2:$C$${budgetLineCount + 1})`, result: expenseTotal };
      summary.getCell("A12").value = "Sisa";
      summary.getCell("B12").value = { formula: "B10-B11", result: -expenseTotal };
      for (const rowNumber of [10, 11, 12]) summary.getCell(`B${rowNumber}`).numFmt = amountFormat;
      summary.getColumn(1).width = 26;
    }

    const guide = workbook.addWorksheet("Petunjuk", { properties: { tabColor: colors.muted } });
    guide.columns = [{ width: 24 }, { width: 100 }];
    guide.addRows([
      ["CATATAN WORKBOOK", ""],
      ["Buku sumber", book.name],
      ["Periode", period ?? "Semua periode"],
      ["Mata uang", selectedCurrency],
      ["Sumber angka", "Transaksi aktif yang tersimpan saat workbook dibuat; transaksi void tidak disertakan."],
      ["Saldo kas", "Arus bersih bukan saldo akhir. Saldo awal tidak tersedia pada sumber data."],
      ["Pembaruan", "Perubahan pada workbook ini tidak mengubah ledger aplikasi."],
      ...(template === "budget" ? [["Anggaran", "Isi kolom Anggaran pada sheet Anggaran; data realisasi diambil dari pengeluaran tersimpan."]] : []),
    ]);
    guide.getRow(1).font = { bold: true, size: 14, color: { argb: colors.white } };
    guide.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.ink } };
    guide.getColumn(1).font = { bold: true, color: { argb: colors.ink } };
    guide.views = [{ state: "normal", showGridLines: false }];

    const defaultFilename = `${template === "cashbook" ? "Buku-Kas" : "Anggaran"}-${book.id}-${Date.now()}.xlsx`;
    const name = outputName(filename, defaultFilename);
    if (extname(name).toLowerCase() !== ".xlsx") throw new Error("Workbook filename must end in .xlsx");
    const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
    const file = await this.workspace.saveGeneratedFile(userId, `exports/${name}`, bytes, xlsxMimeType);
    return {
      ...file,
      template,
      book_id: book.id,
      book_name: book.name,
      period: period ?? "all",
      currency: selectedCurrency,
      transaction_count: transactions.length,
    };
  }

  async inspectSpreadsheet(userId, fileId, { sheetName = null, range = null } = {}) {
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    const design = describeDesign(originalBytes);
    const worksheets = sheetName ? workbook.worksheets.filter((sheet) => sheet.name === sheetName) : workbook.worksheets;
    if (sheetName && worksheets.length === 0) throw new Error(`Worksheet not found: ${sheetName}`);
    return {
      id: file.id,
      filename: file.name,
      sheets: workbook.worksheets.map((sheet) => sheet.name),
      worksheets: worksheets.slice(0, 30).map((worksheet) => {
        const preview = [];
        const formulas = [];
        const previewRows = 40;
        const lastRow = worksheet.rowCount;
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          // Show the top of the sheet and, for long sheets, the tail as well, so a
          // footnote or total block at the bottom is visible before an edit.
          const inHead = rowNumber <= previewRows;
          const inTail = lastRow > previewRows && rowNumber > lastRow - 10;
          if (!inHead && !inTail) return;
          // A merged block repeats its master value in every column; blank the
          // slaves so the preview stays readable and the payload stays small.
          preview.push([rowNumber, ...row.values.slice(1, maxColumns + 1).map((value, index) => {
            const cell = row.getCell(index + 1);
            if (cell.isMerged && cell.master && cell.master.address !== cell.address) return null;
            return safeCellValue(value);
          })]);
          row.eachCell({ includeEmpty: false }, (cell, column) => {
            const value = cell.value;
            if (value && typeof value === "object" && (value.formula || value.sharedFormula)) {
              formulas.push(`${columnLetter(column)}${rowNumber}: ${value.formula ?? value.sharedFormula}`);
            }
          });
        });
        const entry = {
          name: worksheet.name,
          rowCount: worksheet.rowCount,
          columnCount: Math.min(worksheet.columnCount, maxColumns),
          // Each preview row starts with its real row number, then the cell values.
          preview,
          formulas,
          merges: (worksheet.model?.merges ?? []).slice(0, 40),
          widths: worksheet.columns
            .map((column, index) => (column?.width ? `${columnLetter(index + 1)}:${Math.round(column.width)}` : null))
            .filter(Boolean),
          freeze: worksheet.views?.[0]?.state === "frozen"
            ? `${columnLetter(worksheet.views[0].xSplit ?? 0)}${(worksheet.views[0].ySplit ?? 0) + 1}`
            : null,
          gridlines: worksheet.views?.[0]?.showGridLines !== false,
          design: design.get(worksheet.name) ?? { charts: 0, shapes: 0, images: 0 },
        };
        if (range) {
          const bounds = parseRange(range);
          entry.detail = [];
          for (let row = bounds.start.row; row <= bounds.end.row && entry.detail.length < 250; row += 1) {
            for (let col = bounds.start.col; col <= bounds.end.col && entry.detail.length < 250; col += 1) {
              entry.detail.push(describeCell(worksheet.getCell(row, col)));
            }
          }
        }
        return entry;
      }),
    };
  }

  async appendSpreadsheetRows(userId, { fileId, sheetName, rows, outputFilename, saveMode = "in_place" }) {
    validateRows(rows);
    const file = this.workspace.getFile(userId, fileId);
    if (!file) throw new Error("Spreadsheet not found in this user's workspace");
    requireExtension(file, ".xlsx");
    const workbook = new ExcelJS.Workbook();
    const bytes = await readFile(await this.workspace.getFilePath(userId, fileId));
    readSafeOfficeArchive(bytes);
    await workbook.xlsx.load(bytes);
    const targetSheet = this.resolveSheetName(workbook, sheetName);
    this.assertSpreadsheetSheetEditable(userId, fileId, targetSheet);
    const worksheet = workbook.getWorksheet(targetSheet);
    if (!worksheet) throw new Error(`Worksheet not found: ${targetSheet}`);
    if (worksheet.columnCount > maxColumns) throw new Error("Worksheet has too many columns to edit safely");

    for (const values of rows) {
      const newRow = worksheet.addRow(values);
      const styleSource = worksheet.getRow(newRow.number - 1);
      styleSource.eachCell({ includeEmpty: true }, (cell, column) => {
        newRow.getCell(column).style = { ...cell.style };
        if (cell.numFmt) newRow.getCell(column).numFmt = cell.numFmt;
      });
    }

    const updatedBytes = Buffer.from(await workbook.xlsx.writeBuffer());
    // ExcelJS drops charts on write; carry them over from the source file.
    return this.saveEditedSpreadsheet(userId, file, preserveCharts(bytes, updatedBytes).bytes, {
      saveMode,
      outputFilename,
      suffix: "updated",
    });
  }

  async updateSpreadsheetCells(userId, { fileId, sheetName, updates, outputFilename, saveMode = "in_place" }) {
    if (!Array.isArray(updates) || updates.length === 0 || updates.length > 1000) {
      throw new Error("Provide 1-1000 cell updates");
    }
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    const targetSheet = this.resolveSheetName(workbook, sheetName);
    this.assertSpreadsheetSheetEditable(userId, fileId, targetSheet);
    const worksheet = workbook.getWorksheet(targetSheet);
    if (!worksheet) throw new Error(`Worksheet not found: ${targetSheet}`);
    if (worksheet.columnCount > maxColumns || worksheet.rowCount > maxRows) {
      throw new Error("Worksheet exceeds the safe editing dimensions");
    }

    const changes = [];
    for (const update of updates) {
      changes.push(...applyUpdate(worksheet, update, {
        currency: this.defaultCurrency,
        maxRows,
        maxColumns,
      }));
    }

    return this.commitWorkbook(userId, {
      file,
      originalBytes,
      workbook,
      saveMode,
      outputFilename,
      suffix: "updated",
      changes,
      sheetName: targetSheet,
    });
  }

  async insertSpreadsheetRows(userId, { fileId, sheetName, at, count = 1, styleFromRow = null, outputFilename, saveMode = "in_place" }) {
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    const targetSheet = this.resolveSheetName(workbook, sheetName);
    this.assertSpreadsheetSheetEditable(userId, fileId, targetSheet);
    const worksheet = workbook.getWorksheet(targetSheet);
    if (!worksheet) throw new Error(`Worksheet not found: ${targetSheet}`);
    if (worksheet.rowCount + Number(count) > maxRows) throw new Error("Worksheet would exceed the safe row limit");
    const result = insertRows(worksheet, { at, count, styleFromRow });
    return this.commitWorkbook(userId, {
      file,
      originalBytes,
      workbook,
      saveMode,
      outputFilename,
      suffix: "updated",
      changes: [{ cell: `${columnLetter(1)}${result.at}`, kind: "row" }],
      sheetName: targetSheet,
    });
  }

  async deleteSpreadsheetRows(userId, { fileId, sheetName, at, count = 1, outputFilename, saveMode = "in_place" }) {
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    const targetSheet = this.resolveSheetName(workbook, sheetName);
    this.assertSpreadsheetSheetEditable(userId, fileId, targetSheet);
    const worksheet = workbook.getWorksheet(targetSheet);
    if (!worksheet) throw new Error(`Worksheet not found: ${targetSheet}`);
    const result = deleteRows(worksheet, { at, count });
    return this.commitWorkbook(userId, {
      file,
      originalBytes,
      workbook,
      saveMode,
      outputFilename,
      suffix: "updated",
      changes: [],
      sheetName: targetSheet,
      ...{ deleted_rows: result.deleted },
    });
  }

  async insertSpreadsheetColumns(userId, { fileId, sheetName, at, count = 1, styleFromColumn = null, outputFilename, saveMode = "in_place" }) {
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    const targetSheet = this.resolveSheetName(workbook, sheetName);
    this.assertSpreadsheetSheetEditable(userId, fileId, targetSheet);
    const worksheet = workbook.getWorksheet(targetSheet);
    if (!worksheet) throw new Error(`Worksheet not found: ${targetSheet}`);
    if (worksheet.columnCount + Number(count) > maxColumns) throw new Error("Worksheet would exceed the safe column limit");
    const result = insertColumns(worksheet, { at, count, styleFromColumn });
    return this.commitWorkbook(userId, {
      file,
      originalBytes,
      workbook,
      saveMode,
      outputFilename,
      suffix: "updated",
      changes: [],
      sheetName: targetSheet,
      ...{ inserted_columns: result.inserted },
    });
  }

  async deleteSpreadsheetColumns(userId, { fileId, sheetName, at, count = 1, outputFilename, saveMode = "in_place" }) {
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    const targetSheet = this.resolveSheetName(workbook, sheetName);
    this.assertSpreadsheetSheetEditable(userId, fileId, targetSheet);
    const worksheet = workbook.getWorksheet(targetSheet);
    if (!worksheet) throw new Error(`Worksheet not found: ${targetSheet}`);
    const result = deleteColumns(worksheet, { at, count });
    return this.commitWorkbook(userId, {
      file,
      originalBytes,
      workbook,
      saveMode,
      outputFilename,
      suffix: "updated",
      changes: [],
      sheetName: targetSheet,
      ...{ deleted_columns: result.deleted },
    });
  }

  /**
   * Add, rename, duplicate, or delete a worksheet.
   * Duplicating copies values, formulas, styles, widths, merges, and freeze panes
   * so an agent can build a new section from a layout that already looks right.
   */
  async manageSpreadsheetSheet(userId, {
    fileId, action, sheetName = null, newName = null, outputFilename, saveMode = "in_place",
  }) {
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    let result = {};
    if (action === "add") {
      const name = validateSheetName(newName ?? sheetName ?? "Sheet");
      if (workbook.getWorksheet(name)) throw new Error(`Worksheet already exists: ${name}`);
      const worksheet = workbook.addWorksheet(name);
      result = { sheet: worksheet.name, sheets: workbook.worksheets.map((sheet) => sheet.name) };
    } else if (action === "rename") {
      if (!sheetName) throw new Error("rename needs sheet_name");
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) throw new Error(`Worksheet not found: ${sheetName}`);
      const name = validateSheetName(newName);
      if (workbook.getWorksheet(name)) throw new Error(`Worksheet already exists: ${name}`);
      worksheet.name = name;
      result = { from: sheetName, to: name, sheets: workbook.worksheets.map((sheet) => sheet.name) };
    } else if (action === "duplicate") {
      if (!sheetName) throw new Error("duplicate needs sheet_name");
      const source = workbook.getWorksheet(sheetName);
      if (!source) throw new Error(`Worksheet not found: ${sheetName}`);
      const name = validateSheetName(newName ?? `${sheetName} (2)`);
      if (workbook.getWorksheet(name)) throw new Error(`Worksheet already exists: ${name}`);
      const target = workbook.addWorksheet(name);
      copyWorksheetContents(source, target);
      result = { from: sheetName, to: name, sheets: workbook.worksheets.map((sheet) => sheet.name) };
    } else if (action === "delete") {
      if (!sheetName) throw new Error("delete needs sheet_name");
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) throw new Error(`Worksheet not found: ${sheetName}`);
      if (workbook.worksheets.length <= 1) throw new Error("A workbook must keep at least one worksheet");
      workbook.removeWorksheet(worksheet.id);
      result = { deleted: sheetName, sheets: workbook.worksheets.map((sheet) => sheet.name) };
    } else {
      throw new Error(`Unsupported sheet action: ${action}. Use add, rename, duplicate, or delete`);
    }

    const saved = await this.commitWorkbook(userId, {
      file,
      originalBytes,
      workbook,
      saveMode,
      outputFilename,
      suffix: "updated",
      changes: [],
    });
    return { ...saved, ...result };
  }

  /**
   * Add coloured panels (shapes) and native charts to a worksheet.
   * Charts read their values from live cell ranges, so they stay correct when the
   * agent later edits the numbers.
   */
  async designSpreadsheet(userId, {
    fileId, sheetName, shapes = [], charts = [], outputFilename, saveMode = "in_place",
  }) {
    const { file, originalBytes, workbook } = await this.loadWorkbookForEdit(userId, fileId);
    const targetSheet = this.resolveSheetName(workbook, sheetName);
    this.assertSpreadsheetSheetEditable(userId, fileId, targetSheet);
    const worksheet = workbook.getWorksheet(targetSheet);
    if (!worksheet) throw new Error(`Worksheet not found: ${targetSheet}`);

    const normalizedShapes = shapes.map((shape, index) => {
      const anchor = shape.anchor ?? {};
      return {
        sheet: targetSheet,
        text: shape.text ?? "",
        from: resolveAnchorPoint(anchor.from, shape.from, `shapes[${index}].from`),
        to: resolveAnchorPoint(anchor.to, shape.to, `shapes[${index}].to`),
        fill: shape.fill ?? null,
        border: shape.border ?? null,
        textColor: shape.text_color ?? shape.textColor ?? "FFFFFF",
        fontSize: shape.font_size ?? shape.fontSize ?? 11,
        bold: shape.bold ?? true,
        align: shape.align ?? "left",
        valign: shape.valign ?? "middle",
        rounded: shape.rounded ?? true,
        shadow: shape.shadow ?? false,
      };
    });

    const normalizedCharts = charts.map((chart, index) => buildChartSpec(workbook, worksheet, chart, index));
    if (!normalizedShapes.length && !normalizedCharts.length) {
      throw new Error("Provide at least one shape or chart to add");
    }

    return this.commitWorkbook(userId, {
      file,
      originalBytes,
      workbook,
      design: { shapes: normalizedShapes, charts: normalizedCharts },
      saveMode,
      outputFilename,
      suffix: "designed",
      changes: [],
      sheetName: targetSheet,
    });
  }

  async createWordDocument(userId, { filename, title, paragraphs }) {
    if (typeof title !== "string" || !title.trim() || title.length > 200) {
      throw new Error("Word document title must contain 1-200 characters");
    }
    if (!Array.isArray(paragraphs) || paragraphs.length > 500 || paragraphs.some((item) => typeof item !== "string" || item.length > 5000)) {
      throw new Error("Word document paragraphs must be up to 500 strings of at most 5000 characters");
    }
    const name = outputName(filename, "document.docx");
    if (extname(name).toLowerCase() !== ".docx") throw new Error("Word document filename must end in .docx");
    const document = new Document({
      sections: [{
        children: [
          new Paragraph({ text: title.trim(), heading: HeadingLevel.TITLE }),
          ...paragraphs.map((text) => new Paragraph({ text })),
        ],
      }],
    });
    const bytes = await Packer.toBuffer(document);
    return this.workspace.saveGeneratedFile(userId, `exports/${name}`, bytes, docxMimeType);
  }

  async readWordDocument(userId, fileId) {
    const file = this.workspace.getFile(userId, fileId);
    if (!file) throw new Error("Word document not found in this user's workspace");
    requireExtension(file, ".docx");
    const bytes = await readFile(await this.workspace.getFilePath(userId, fileId));
    const zip = readSafeOfficeArchive(bytes);
    const documentXml = zip.file("word/document.xml")?.asText();
    if (!documentXml) throw new Error("Word document has no main document body");
    const parsed = new DOMParser().parseFromString(documentXml, "application/xml");
    const paragraphs = Array.from(parsed.getElementsByTagName("w:p"))
      .map((paragraph) => Array.from(paragraph.getElementsByTagName("w:t"))
        .map((text) => text.textContent ?? "").join(""))
      .filter(Boolean);
    return { id: file.id, filename: file.name, text: paragraphs.join("\n").slice(0, 20_000) };
  }

  async fillWordTemplate(userId, { fileId, data, outputFilename }) {
    const file = this.workspace.getFile(userId, fileId);
    if (!file) throw new Error("Word template not found in this user's workspace");
    requireExtension(file, ".docx");
    if (!data || typeof data !== "object" || Array.isArray(data) || Object.keys(data).length > 100) {
      throw new Error("Template data must be an object with no more than 100 fields");
    }
    const filePath = await this.workspace.getFilePath(userId, fileId);
    const templateBytes = await readFile(filePath);
    const zip = readSafeOfficeArchive(templateBytes);
    const template = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    template.render(data);
    const bytes = template.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
    const filename = outputName(outputFilename, `${basename(file.name, ".docx")}-filled.docx`);
    if (extname(filename).toLowerCase() !== ".docx") throw new Error("Output filename must end in .docx");
    return this.workspace.saveGeneratedFile(userId, `exports/${filename}`, bytes, docxMimeType);
  }

  async createZipArchive(userId, { fileIds, filename }) {
    if (!Array.isArray(fileIds) || fileIds.length < 1 || fileIds.length > 100
      || fileIds.some((fileId) => typeof fileId !== "string")
      || new Set(fileIds).size !== fileIds.length) {
      throw new Error("Provide 1-100 unique file IDs to archive");
    }
    const archiveName = outputName(filename, "workspace.zip");
    if (extname(archiveName).toLowerCase() !== ".zip") throw new Error("Archive filename must end in .zip");
    const zip = new PizZip();
    let totalBytes = 0;
    for (const fileId of fileIds) {
      const file = this.workspace.getFile(userId, fileId);
      if (!file) throw new Error(`File ${fileId} was not found in this user's workspace`);
      totalBytes += file.sizeBytes;
      if (totalBytes > maxOfficeExpandedBytes) throw new Error("Selected files exceed the 100 MB archive limit");
      zip.file(file.path, await readFile(await this.workspace.getFilePath(userId, file.id)));
    }
    const bytes = Buffer.from(zip.generate({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    }));
    return this.workspace.saveGeneratedFile(userId, `exports/${archiveName}`, bytes, "application/zip");
  }

  async extractZipArchive(userId, { fileId, destinationFolder }) {
    const file = this.workspace.getFile(userId, fileId);
    if (!file) throw new Error("ZIP archive not found in this user's workspace");
    requireExtension(file, ".zip");
    const zip = readSafeOfficeArchive(await readFile(await this.workspace.getFilePath(userId, fileId)));
    const destination = await this.workspace.resolveSafePath(
      userId,
      destinationFolder || `exports/${basename(file.name, extname(file.name))}`,
    );
    const extracted = [];
    const outputPaths = new Set();
    let totalBytes = 0;

    for (const [entryName, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      if (isSymbolicLink(entry)) throw new Error("ZIP archive contains a symbolic link");
      const entryPath = archiveEntryPath(entryName);
      const relativePath = `${destination.relativePath}/${entryPath}`;
      if (outputPaths.has(relativePath)) throw new Error("ZIP archive contains duplicate output paths");
      outputPaths.add(relativePath);
      const target = await this.workspace.resolveSafePath(userId, relativePath);
      if (await this.workspace.pathExists(target.target)) {
        throw new Error(`Extraction would overwrite existing file ${relativePath}`);
      }
      const bytes = entry.asNodeBuffer();
      totalBytes += bytes.length;
      if (bytes.length > this.workspace.maxFileBytes || totalBytes > maxOfficeExpandedBytes) {
        throw new Error("Extracted files exceed the configured size limits");
      }
      extracted.push({
        relativePath,
        bytes,
        mimeType: this.workspace.mimeForExtension(extname(relativePath).toLowerCase()),
      });
    }
    if (extracted.length === 0) throw new Error("ZIP archive contains no files to extract");
    return this.workspace.saveGeneratedFiles(userId, extracted);
  }
}

function validateSheetName(value) {
  const name = String(value ?? "").trim();
  if (!/^[^\[\]*?:/\\]{1,31}$/.test(name)) {
    throw new Error("Worksheet name must be 1-31 characters without [ ] * ? : / \\");
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new Error("Worksheet name cannot start or end with an apostrophe");
  }
  return name;
}

/** Copy values, formulas, styles, widths, merges, and freeze panes to a new sheet. */
function copyWorksheetContents(source, target) {
  if (source.rowCount > 2000) throw new Error("Worksheet is too large to duplicate safely (2000 row limit)");
  source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const targetRow = target.getRow(rowNumber);
    if (row.height !== undefined) targetRow.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, column) => {
      const targetCell = targetRow.getCell(column);
      if (cell.value !== null && cell.value !== undefined) targetCell.value = structuredClone(cell.value);
      if (cell.style) targetCell.style = structuredClone(cell.style);
    });
  });
  source.columns.forEach((column, index) => {
    if (!column) return;
    const targetColumn = target.getColumn(index + 1);
    if (column.width !== undefined) targetColumn.width = column.width;
    if (column.hidden) targetColumn.hidden = true;
  });
  for (const merge of source.model?.merges ?? []) {
    try {
      target.mergeCells(merge);
    } catch {
      // A merge that cannot be recreated is cosmetic; never fail the copy for it.
    }
  }
  if (source.views?.length) target.views = structuredClone(source.views);
  if (source.properties?.tabColor) {
    target.properties = { ...target.properties, tabColor: structuredClone(source.properties.tabColor) };
  }
}

function cellScalar(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return value;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) return value.richText.map((run) => run.text).join("");
    if (Object.hasOwn(value, "result")) return cellScalar(value.result);
    if (typeof value.text === "string") return value.text;
    return null;
  }
  return value;
}

/** Accept "A1:B2" or "Sheet!A1:B2" and return the sheet name plus the range. */
function splitSheetReference(reference, fallbackSheet, label) {
  const text = String(reference ?? "").trim();
  if (!text) throw new Error(`${label} is required`);
  const bang = text.lastIndexOf("!");
  if (bang === -1) return { sheetName: fallbackSheet, range: text };
  const sheetName = text.slice(0, bang).replace(/^'|'$/g, "").replaceAll("''", "'");
  return { sheetName, range: text.slice(bang + 1) };
}

function readRangeScalars(worksheet, address, label) {
  let bounds;
  try {
    bounds = parseRange(address);
  } catch {
    throw new Error(`${label} must be a range like B5:B12, got: ${address}`);
  }
  const values = [];
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    for (let col = bounds.start.col; col <= bounds.end.col; col += 1) {
      values.push(cellScalar(worksheet.getCell(row, col).value));
    }
  }
  return values;
}

/** Accept an A1 address ("J3") or a zero-based {col,row} anchor point. */
function resolveAnchorPoint(anchor, fallbackA1, label) {
  const source = anchor ?? fallbackA1;
  if (source === undefined || source === null || source === "") throw new Error(`${label} is required`);
  if (typeof source === "string") {
    const point = parseA1(source);
    return { col: point.col - 1, row: point.row - 1 };
  }
  if (typeof source === "object") {
    const col = Number(source.col ?? source.column);
    const row = Number(source.row);
    if (!Number.isInteger(col) || !Number.isInteger(row) || col < 0 || row < 0) {
      throw new Error(`${label} must be an A1 address or {col,row} zero-based integers`);
    }
    return { col, row };
  }
  throw new Error(`${label} must be an A1 address (J3) or {col,row}`);
}

const chartPalette = ["0F766E", "B08D57", "3B6E8F", "8A5A96", "B3341F", "4E8C6A", "7A7F8C", "C77B3C"];

/**
 * Turn the agent's friendly chart request into the spec the zip injector needs.
 * Values are read from the live cells, so the chart is a real Excel chart that
 * recalculates when the numbers change.
 */
function buildChartSpec(workbook, worksheet, chart, index) {
  const label = `charts[${index}]`;
  const type = String(chart.type ?? "bar").toLowerCase();
  if (!["bar", "line", "doughnut"].includes(type)) {
    throw new Error(`${label}.type must be bar, line, or doughnut`);
  }
  const from = resolveAnchorPoint(chart.anchor?.from, chart.from, `${label}.from`);
  const to = resolveAnchorPoint(chart.anchor?.to, chart.to, `${label}.to`);
  if (to.col <= from.col || to.row <= from.row) {
    throw new Error(`${label} anchor must end below and to the right of its start`);
  }
  const sourceSheet = chart.source_sheet ? workbook.getWorksheet(chart.source_sheet) : worksheet;
  if (!sourceSheet) throw new Error(`${label}.source_sheet not found: ${chart.source_sheet}`);
  if (!chart.categories) throw new Error(`${label}.categories range is required, e.g. A5:A10`);
  const categoryRef = splitSheetReference(chart.categories, worksheet.name, `${label}.categories`);
  const categorySheet = workbook.getWorksheet(categoryRef.sheetName);
  if (!categorySheet) throw new Error(`${label}.categories sheet not found: ${categoryRef.sheetName}`);
  const categories = readRangeScalars(categorySheet, categoryRef.range, `${label}.categories`)
    .map((value) => (value === null ? "" : String(value)));
  if (!categories.length) throw new Error(`${label}.categories range is empty`);

  const requestedSeries = Array.isArray(chart.series) && chart.series.length
    ? chart.series
    : [{ name: chart.title ?? "Nilai", range: chart.values }];
  const series = requestedSeries.map((entry, seriesIndex) => {
    const seriesLabel = `${label}.series[${seriesIndex}]`;
    if (!entry?.range) throw new Error(`${seriesLabel}.range is required, e.g. B5:B10`);
    const valueRef = splitSheetReference(entry.range, worksheet.name, `${seriesLabel}.range`);
    const target = entry.sheet_name ? workbook.getWorksheet(entry.sheet_name) : workbook.getWorksheet(valueRef.sheetName);
    if (!target) throw new Error(`${seriesLabel} sheet not found: ${entry.sheet_name ?? valueRef.sheetName}`);
    const values = readRangeScalars(target, valueRef.range, `${seriesLabel}.range`)
      .map((value) => (Number.isFinite(Number(value)) ? Number(value) : 0));
    if (values.length !== categories.length) {
      throw new Error(`${seriesLabel}.range has ${values.length} values but categories has ${categories.length}`);
    }
    const nameCell = entry.name_cell ? parseA1(entry.name_cell) : null;
    return {
      name: entry.name ?? chart.title ?? "Nilai",
      nameRef: nameCell ? sheetRef(target.name, `$${columnLetter(nameCell.col)}$${nameCell.row}`) : null,
      ref: sheetRef(target.name, valueRef.range.toUpperCase()),
      values,
      color: String(entry.color ?? chartPalette[seriesIndex % chartPalette.length]).replace(/^#/, "").toUpperCase(),
    };
  });

  return {
    sheet: worksheet.name,
    type,
    title: chart.title ?? null,
    from,
    to,
    categories: { ref: sheetRef(categorySheet.name, categoryRef.range.toUpperCase()), values: categories },
    series,
    palette: Array.isArray(chart.palette) ? chart.palette : chartPalette,
    formatCode: chart.format ?? "#,##0",
  };
}

function readSafeOfficeArchive(bytes) {  const zip = new PizZip(bytes);
  const entries = Object.entries(zip.files);
  if (entries.length > maxOfficeArchiveEntries) {
    throw new Error("Office archive contains too many entries to process safely");
  }
  let expandedBytes = 0;
  for (const [name, entry] of entries) {
    const normalizedName = name.replaceAll("\\", "/");
    if (normalizedName.startsWith("/") || normalizedName.split("/").includes("..")) {
      throw new Error("Office archive contains an unsafe path");
    }
    const compressedBytes = Number(entry._data?.compressedSize ?? 0);
    const uncompressedBytes = Number(entry._data?.uncompressedSize ?? 0);
    expandedBytes += uncompressedBytes;
    if (!Number.isFinite(uncompressedBytes) || expandedBytes > maxOfficeExpandedBytes) {
      throw new Error("Office archive expands beyond the 100 MB processing limit");
    }
    if (uncompressedBytes > 0 && (compressedBytes === 0 || uncompressedBytes / compressedBytes > 250)) {
      throw new Error("Office archive compression ratio exceeds the safe limit");
    }
  }
  return zip;
}
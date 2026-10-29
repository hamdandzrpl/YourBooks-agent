import ExcelJS from "exceljs";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { analyzeBook, primaryCurrency } from "./analytics.js";
import { renderReport } from "./report.js";

const workbookMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const periodPattern = /^\d{4}-(0[1-9]|1[0-2])$/;

function ledgerSignature(transactions) {
  return createHash("sha256").update(JSON.stringify(transactions)).digest("hex");
}

/**
 * Readable filename for a book's managed workbook. The old scheme used the raw
 * book id (`books/B-88DD86E4.xlsx`), which the user and the agent both saw in the
 * explorer and in WhatsApp as an unreadable name for their own report.
 */
export function bookWorkbookSlug(name) {
  const slug = String(name ?? "")
    .toLocaleLowerCase("id-ID")
    .replace(/[^\p{L}\p{N}]+/gu, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);
  return slug || "buku";
}

/**
 * Build the analysis + rendered report for a set of ledger rows.
 * Kept pure-ish so both the managed book workbook and one-off exports share it.
 */
export async function buildReport({ transactions, bookName, currency, existingBytes, today, charts = true, verify = true, sheets, title }) {
  const workbook = new ExcelJS.Workbook();
  if (existingBytes) {
    try {
      await workbook.xlsx.load(existingBytes);
    } catch {
      // A damaged or exotic workbook should never block a fresh report.
    }
  }
  const analysis = analyzeBook({ transactions, currency, bookName, today });
  const rendered = await renderReport({ analysis, transactions, workbook, charts, verify, sheets, title });
  return { analysis, ...rendered };
}

export class FinanceWorkbookSync {
  constructor({ financeStore, fileWorkspace, excelMcp, defaultCurrency = "IDR", now = () => new Date() }) {
    this.financeStore = financeStore;
    this.fileWorkspace = fileWorkspace;
    this.excelMcp = excelMcp;
    this.defaultCurrency = defaultCurrency;
    this.now = now;
    this.syncQueues = new Map();
    this.lastSyncSignatures = new Map();
  }

  /** `books/<nama-buku>.xlsx`, with a numeric suffix when the name is taken. */
  async uniqueBookWorkbookPath(userId, book, { excludeFileId = null } = {}) {
    const base = `books/${bookWorkbookSlug(book.name)}`;
    const taken = async (path) => {
      const existing = (await this.fileWorkspace.listFiles(userId, "books"))
        .find((entry) => entry.path === path && entry.id !== excludeFileId);
      return Boolean(existing);
    };
    let candidate = `${base}.xlsx`;
    for (let attempt = 2; attempt <= 50 && await taken(candidate); attempt += 1) {
      candidate = `${base}-${attempt}.xlsx`;
    }
    return candidate;
  }

  async ensureBookWorkbook(userId, bookId) {
    const book = this.financeStore.getBook(userId, bookId);
    if (!book) throw new Error("Book not found in this user's workspace");
    if (book.workbook_file_id) {
      const existing = this.fileWorkspace.getFile(userId, book.workbook_file_id);
      if (existing) return this.relabelBookWorkbook(userId, book, existing);
      try {
        const restored = await this.fileWorkspace.restoreFile(userId, book.workbook_file_id);
        if (restored) return this.relabelBookWorkbook(userId, book, restored);
      } catch {}
    }

    const bytes = await this.renderBytes(userId, book, []);
    let file;
    try {
      const path = await this.uniqueBookWorkbookPath(userId, book);
      file = await this.fileWorkspace.saveGeneratedFile(userId, path, bytes, workbookMimeType);
    } catch (error) {
      if (!String(error.message).includes("already exists")) throw error;
      file = await this.fileWorkspace.saveGeneratedFile(
        userId,
        `books/${bookWorkbookSlug(book.name)}-${Date.now()}.xlsx`,
        bytes,
        workbookMimeType,
      );
    }
    this.financeStore.updateBookWorkbook(userId, book.id, file.id);
    const emptySignature = ledgerSignature([]);
    this.financeStore.markBookWorkbookSynced(userId, book.id, emptySignature);
    this.lastSyncSignatures.set(`${userId}:${book.id}`, emptySignature);
    return file;
  }

  /**
   * Migrate a legacy managed workbook (`books/<book-id>.xlsx`) to its readable
   * name. The same file id is kept, so `workbook_file_id` stays valid. A name the
   * user chose themselves is never touched: only the id-shaped legacy path is
   * rewritten.
   */
  async relabelBookWorkbook(userId, book, file) {
    const legacyNames = new Set([
      `books/${book.id}.xlsx`,
      `books/${book.id.toLocaleLowerCase("en-US")}.xlsx`,
    ]);
    if (!legacyNames.has(file.path)) return file;
    const desired = await this.uniqueBookWorkbookPath(userId, book, { excludeFileId: file.id });
    try {
      return await this.fileWorkspace.renameFile(userId, file.id, desired);
    } catch {
      // A rename is cosmetic; never let it break a sync or a report.
      return file;
    }
  }

  async renderBytes(userId, book, transactions, { existingBytes = null } = {}) {
    const currency = primaryCurrency(transactions, this.defaultCurrency);
    const { bytes } = await buildReport({
      transactions,
      bookName: book.name,
      currency,
      existingBytes,
      today: this.now().toISOString().slice(0, 10),
      // Managed sync runs after every recorded transaction, so skip the extra
      // verification pass; exports keep it because they leave the workspace.
      verify: false,
    });
    return bytes;
  }

  /** Re-render the managed workbook in place, preserving any sheets the user added. */
  async writeBookWorkbook(userId, file, transactions, book) {
    const registeredPath = await this.fileWorkspace.getFilePath(userId, file.id);
    const existingBytes = await readFile(registeredPath).catch(() => null);
    const bytes = await this.renderBytes(userId, book, transactions, { existingBytes });
    return this.fileWorkspace.replaceGeneratedFile(userId, file.id, bytes);
  }

  syncBook(userId, bookId) {
    const queueKey = `${userId}:${bookId}`;
    const previous = this.syncQueues.get(queueKey) ?? Promise.resolve();
    const current = previous.catch(() => {}).then(() => this.performSync(userId, bookId));
    this.syncQueues.set(queueKey, current);
    return current.finally(() => {
      if (this.syncQueues.get(queueKey) === current) this.syncQueues.delete(queueKey);
    });
  }

  async performSync(userId, bookId) {
    let book = this.financeStore.getBook(userId, bookId);
    if (!book) throw new Error("Book not found in this user's workspace");
    const transactions = this.financeStore.listBookTransactions(userId, book.id);
    const file = await this.ensureBookWorkbook(userId, book.id);
    const queueKey = `${userId}:${book.id}`;
    const signature = ledgerSignature(transactions);
    book = this.financeStore.getBook(userId, book.id);
    if (book.workbook_sync_hash === signature || this.lastSyncSignatures.get(queueKey) === signature) {
      this.lastSyncSignatures.set(queueKey, signature);
      return file;
    }
    const updated = await this.writeBookWorkbook(userId, file, transactions, book);
    this.financeStore.markBookWorkbookSynced(userId, book.id, signature);
    this.lastSyncSignatures.set(queueKey, signature);
    return updated;
  }

  async createBook(userId, name) {
    const book = this.financeStore.createBook(userId, name);
    try {
      const file = await this.ensureBookWorkbook(userId, book.id);
      this.lastSyncSignatures.set(`${userId}:${book.id}`, ledgerSignature([]));
      return { ...this.financeStore.getBook(userId, book.id), workbook: file };
    } catch (error) {
      return { ...book, workbook_sync_error: error.message };
    }
  }

  async syncTransactionChange(userId, transaction) {
    if (!transaction?.book_id) return null;
    return this.syncBook(userId, transaction.book_id);
  }

  /**
   * Render a standalone report (export) without touching the managed workbook.
   * @returns {Promise<{file: object, analysis: object, charts: number}>}
   */
  async renderReportFile(userId, { bookId = null, period = null, currency = null, filename = null, sheets = undefined, title = null } = {}) {
    const book = bookId ? this.financeStore.getBook(userId, bookId) : this.financeStore.getActiveBook(userId);
    if (!book) throw new Error("Book not found in this user's workspace");
    if (period && !periodPattern.test(period)) throw new Error("Period must use YYYY-MM format");

    const all = this.financeStore.listBookTransactions(userId, book.id);
    const transactions = period ? all.filter((row) => row.occurred_on.startsWith(period)) : all;
    if (!transactions.length) throw new Error("Tidak ada transaksi pada periode tersebut; laporan tidak dibuat");

    const selectedCurrency = currency
      ? String(currency).toUpperCase()
      : primaryCurrency(transactions, this.defaultCurrency);
    if (!/^[A-Z]{3}$/.test(selectedCurrency)) throw new Error("Currency must be a three-letter code");

    const { bytes, analysis, charts, sheets: renderedSheets } = await buildReport({
      transactions,
      bookName: book.name,
      currency: selectedCurrency,
      today: this.now().toISOString().slice(0, 10),
      sheets,
      title,
    });

    const slug = book.name.toLocaleLowerCase("id-ID").replace(/[^\p{L}\p{N}]+/gu, "-").replace(/^-|-$/g, "") || "buku";
    const suffix = period ? period : "semua-periode";
    const name = filename && /^[\w .-]+\.xlsx$/u.test(filename)
      ? filename
      : `laporan-${slug}-${suffix}.xlsx`;
    const file = await this.fileWorkspace.saveGeneratedFile(
      userId, await this.uniqueExportPath(userId, name), bytes, workbookMimeType,
    );
    return { file, analysis, charts, sheets: renderedSheets };
  }

  /**
   * Free path for an exported report. Asking for the same period twice used to
   * fail with "A file already exists at the requested path", which the agent
   * surfaced to the user as "laporan tidak bisa disimpan".
   */
  async uniqueExportPath(userId, name) {
    const taken = new Set((await this.fileWorkspace.listFiles(userId, "exports")).map((entry) => entry.name));
    if (!taken.has(name)) return `exports/${name}`;
    const base = name.replace(/\.xlsx$/i, "");
    for (let attempt = 2; attempt <= 100; attempt += 1) {
      const candidate = `${base}-${attempt}.xlsx`;
      if (!taken.has(candidate)) return `exports/${candidate}`;
    }
    return `exports/${base}-${Date.now()}.xlsx`;
  }
}

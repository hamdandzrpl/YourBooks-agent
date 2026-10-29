import { randomUUID } from "node:crypto";
import { chmodSync, mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { DatabaseSync } from "node:sqlite";

if (process.platform !== "win32") {
  process.umask(0o077);
}

const memoryCharacterLimits = { memory: 2200, user: 1375 };
const maxMemoryEntryCharacters = 500;
const defaultBookName = "Buku Utama";

function getCurrencyDigits(currency) {
  if (!/^[A-Z]{3}$/.test(currency)) {
    throw new Error("Currency must be a three-letter uppercase code");
  }
  return new Intl.NumberFormat("en", { style: "currency", currency })
    .resolvedOptions().maximumFractionDigits;
}

function toMinorUnits(amount, currency) {
  if (typeof amount !== "number" || !Number.isFinite(amount) || amount <= 0) {
    throw new Error("Amount must be a positive finite number");
  }
  const scaledAmount = amount * (10 ** getCurrencyDigits(currency));
  const minor = Math.round(scaledAmount);
  const roundingTolerance = Number.EPSILON * Math.max(1, Math.abs(scaledAmount)) * 4;
  if (Math.abs(scaledAmount - minor) > roundingTolerance) {
    throw new Error("Amount has more decimal places than the currency supports");
  }
  if (!Number.isSafeInteger(minor) || minor <= 0) {
    throw new Error("Amount is outside the supported range");
  }
  return minor;
}

function fromMinorUnits(amountMinor, currency) {
  const digits = getCurrencyDigits(currency);
  if (digits === 0) return String(amountMinor);
  const sign = amountMinor < 0 ? "-" : "";
  const absolute = String(Math.abs(amountMinor)).padStart(digits + 1, "0");
  return `${sign}${absolute.slice(0, -digits)}.${absolute.slice(-digits)}`;
}

function isValidDate(date) {
  return typeof date === "string"
    && /^\d{4}-\d{2}-\d{2}$/.test(date)
    && new Date(`${date}T00:00:00.000Z`).toISOString().slice(0, 10) === date;
}

function normalizeTransactionInput(input, defaultCurrency, defaultDate) {
  const kind = input.kind;
  if (!["income", "expense"].includes(kind)) {
    throw new Error("Transaction kind must be income or expense");
  }
  const currency = (input.currency || defaultCurrency).toUpperCase();
  const amountMinor = toMinorUnits(input.amount, currency);
  const occurredOn = input.date || defaultDate;
  if (!isValidDate(occurredOn)) {
    throw new Error("Transaction date must be a real YYYY-MM-DD date");
  }
  const category = typeof input.category === "string" && input.category.trim()
    ? input.category.trim()
    : "Lain-lain";
  if (!/^[\p{L}\p{N} _.-]{1,60}$/u.test(category)) {
    throw new Error("Category must be 1-60 letters, numbers, spaces, dots, underscores, or hyphens");
  }
  const description = typeof input.description === "string" && input.description.trim()
    ? input.description.trim()
    : kind === "income" ? "Pemasukan" : "Pengeluaran";
  if (!description || description.length > 500 || /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(description)) {
    throw new Error("Description must contain 1-500 printable characters");
  }
  return { kind, amountMinor, currency, occurredOn, category, description };
}

function csvCell(value) {
  let text = String(value ?? "");
  if (/^[\s\u0000-\u0020]*[=+\-@]/.test(text)) {
    text = `'${text}`;
  }
  return `"${text.replaceAll('"', '""')}"`;
}

function validateMemoryEntry(content) {
  if (typeof content !== "string") throw new Error("Memory content must be text");
  const normalized = content.trim();
  const characterCount = [...normalized].length;
  if (characterCount < 1 || characterCount > maxMemoryEntryCharacters) {
    throw new Error(`Memory entries must contain 1-${maxMemoryEntryCharacters} characters`);
  }
  if (/[\u0000-\u001F\u007F-\u009F\u200B-\u200F\u202A-\u202E\u2060-\u2069\uFEFF]/u.test(normalized)) {
    throw new Error("Memory cannot contain control or invisible Unicode characters");
  }
  if (/ignore\s+(?:all\s+|previous\s+|prior\s+)?instructions|reveal\s+(?:the\s+)?system prompt|exfiltrat|bypass\s+(?:all\s+)?safety|send\s+.*(?:secret|credentials)/i.test(normalized)) {
    throw new Error("Memory cannot contain prompt-injection or exfiltration instructions");
  }
  if (/(?:\b(?:api[_ -]?key|password|passwd|secret|bearer token|otp|private key)\s*[:=]|\b(?:rp\.?|idr|usd|eur)\s*[\d][\d.,]*|\b\d{12,19}\b)/i.test(normalized)) {
    throw new Error("Do not store credentials, account identifiers, or exact financial amounts in long-term memory");
  }
  return normalized;
}

export class FinanceStore {
  constructor(databasePath, { now = () => new Date() } = {}) {
    const resolvedPath = databasePath === ":memory:" ? databasePath : resolve(databasePath);
    if (resolvedPath !== ":memory:") {
      mkdirSync(dirname(resolvedPath), { recursive: true, mode: 0o700 });
    }
    this.now = now;
    this.database = new DatabaseSync(resolvedPath);
    this.database.exec("PRAGMA foreign_keys = ON; PRAGMA journal_mode = WAL; PRAGMA busy_timeout = 5000;");
    this.database.exec(`
      CREATE TABLE IF NOT EXISTS books (
        id TEXT PRIMARY KEY,
        user_id TEXT NOT NULL,
        name TEXT NOT NULL,
        workbook_file_id TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        UNIQUE(user_id, name COLLATE NOCASE)
      );
      CREATE TABLE IF NOT EXISTS user_book_state (
        user_id TEXT PRIMARY KEY,
        active_book_id TEXT NOT NULL REFERENCES books(id),
        updated_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS drafts (
        id TEXT PRIMARY KEY,
        user_id TEXT NOT NULL,
        book_id TEXT REFERENCES books(id),
        source_message_id TEXT,
        kind TEXT NOT NULL CHECK (kind IN ('income', 'expense')),
        amount_minor INTEGER NOT NULL CHECK (amount_minor > 0),
        currency TEXT NOT NULL,
        occurred_on TEXT NOT NULL,
        category TEXT NOT NULL,
        description TEXT NOT NULL,
        status TEXT NOT NULL CHECK (status IN ('pending', 'approved', 'cancelled', 'expired')),
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL
      );
      CREATE UNIQUE INDEX IF NOT EXISTS drafts_source_message_unique
        ON drafts(user_id, source_message_id) WHERE source_message_id IS NOT NULL;
      CREATE TABLE IF NOT EXISTS transactions (
        id TEXT PRIMARY KEY,
        draft_id TEXT NOT NULL UNIQUE REFERENCES drafts(id),
        user_id TEXT NOT NULL,
        book_id TEXT REFERENCES books(id),
        source_message_id TEXT UNIQUE,
        kind TEXT NOT NULL CHECK (kind IN ('income', 'expense')),
        amount_minor INTEGER NOT NULL CHECK (amount_minor > 0),
        currency TEXT NOT NULL,
        occurred_on TEXT NOT NULL,
        category TEXT NOT NULL,
        description TEXT NOT NULL,
        approved_at TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'active' CHECK (status IN ('active', 'voided'))
      );
      CREATE TABLE IF NOT EXISTS postings (
        id INTEGER PRIMARY KEY,
        transaction_id TEXT NOT NULL REFERENCES transactions(id),
        account TEXT NOT NULL,
        amount_minor INTEGER NOT NULL,
        currency TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS message_receipts (
        message_id TEXT PRIMARY KEY,
        user_id TEXT NOT NULL,
        status TEXT NOT NULL CHECK (status IN ('queued', 'completed', 'failed')),
        received_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS audit_events (
        id INTEGER PRIMARY KEY,
        user_id TEXT NOT NULL,
        action TEXT NOT NULL,
        reference_id TEXT NOT NULL,
        created_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS transaction_revisions (
        id INTEGER PRIMARY KEY,
        user_id TEXT NOT NULL,
        transaction_id TEXT NOT NULL REFERENCES transactions(id),
        before_json TEXT NOT NULL,
        after_json TEXT NOT NULL,
        created_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS memory_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id TEXT NOT NULL,
        target TEXT NOT NULL CHECK (target IN ('memory', 'user')),
        content TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        UNIQUE(user_id, target, content)
      );
      CREATE INDEX IF NOT EXISTS memory_entries_user_target
        ON memory_entries(user_id, target, id);
      CREATE TABLE IF NOT EXISTS conversation_messages (
        id INTEGER PRIMARY KEY,
        user_id TEXT NOT NULL,
        source_message_id TEXT NOT NULL,
        role TEXT NOT NULL CHECK (role IN ('user', 'assistant')),
        content TEXT NOT NULL,
        created_at TEXT NOT NULL,
        UNIQUE(user_id, source_message_id, role)
      );
      CREATE INDEX IF NOT EXISTS conversation_messages_recent
        ON conversation_messages(user_id, created_at DESC, id DESC);
    `);
    const bookColumns = this.database.prepare("PRAGMA table_info(books)").all();
    if (!bookColumns.some((column) => column.name === "workbook_sync_hash")) {
      this.database.exec("ALTER TABLE books ADD COLUMN workbook_sync_hash TEXT");
    }
    const transactionColumns = this.database.prepare("PRAGMA table_info(transactions)").all();
    if (!transactionColumns.some((column) => column.name === "status")) {
      this.database.exec("ALTER TABLE transactions ADD COLUMN status TEXT NOT NULL DEFAULT 'active' CHECK (status IN ('active', 'voided'))");
    }
    const draftColumns = this.database.prepare("PRAGMA table_info(drafts)").all();
    if (!draftColumns.some((column) => column.name === "book_id")) {
      this.database.exec("ALTER TABLE drafts ADD COLUMN book_id TEXT REFERENCES books(id)");
    }
    if (!transactionColumns.some((column) => column.name === "book_id")) {
      this.database.exec("ALTER TABLE transactions ADD COLUMN book_id TEXT REFERENCES books(id)");
    }
    this.migrateDefaultBooks();
    this.database.exec(`
      CREATE INDEX IF NOT EXISTS transactions_user_book_active_date
        ON transactions(user_id, book_id, status, occurred_on DESC);
      CREATE INDEX IF NOT EXISTS drafts_user_book_status
        ON drafts(user_id, book_id, status);
    `);
    this.hasFts5 = false;

    if (resolvedPath !== ":memory:") {
      try {
        chmodSync(resolvedPath, 0o600);
      } catch {
        // File permissions are best-effort on platforms that do not support POSIX modes.
      }
    }
  }

  close() {
    this.database.close();
  }

  exportBackupSnapshot(userId, components) {
    if (typeof userId !== "string" || !/^\d{8,15}$/.test(userId)) {
      throw new Error("A valid authorized user ID is required");
    }
    const selected = new Set(components);
    const snapshot = { formatVersion: 1, exportedAt: this.now().toISOString() };
    if (selected.has("ai")) {
      snapshot.ai = {
        memory_entries: this.database.prepare("SELECT * FROM memory_entries WHERE user_id = ? ORDER BY id").all(userId),
        conversation_messages: this.database.prepare("SELECT * FROM conversation_messages WHERE user_id = ? ORDER BY id").all(userId),
        message_receipts: this.database.prepare("SELECT * FROM message_receipts WHERE user_id = ? ORDER BY received_at, message_id").all(userId),
      };
    }
    if (selected.has("finance")) {
      snapshot.finance = {
        books: this.database.prepare("SELECT * FROM books WHERE user_id = ? ORDER BY created_at, id").all(userId),
        user_book_state: this.database.prepare("SELECT * FROM user_book_state WHERE user_id = ?").all(userId),
        drafts: this.database.prepare("SELECT * FROM drafts WHERE user_id = ? ORDER BY created_at, id").all(userId),
        transactions: this.database.prepare("SELECT * FROM transactions WHERE user_id = ? ORDER BY approved_at, id").all(userId),
        postings: this.database.prepare(`
          SELECT * FROM postings WHERE transaction_id IN
            (SELECT id FROM transactions WHERE user_id = ?) ORDER BY id
        `).all(userId),
        audit_events: this.database.prepare("SELECT * FROM audit_events WHERE user_id = ? ORDER BY id").all(userId),
        transaction_revisions: this.database.prepare("SELECT * FROM transaction_revisions WHERE user_id = ? ORDER BY id").all(userId),
      };
    }
    return snapshot;
  }

  restoreBackupSnapshot(userId, components, snapshot) {
    const selected = new Set(components);
    const requireRows = (rows, name) => {
      if (!Array.isArray(rows) || rows.length > 100_000
        || rows.some((row) => !row || typeof row !== "object" || row.user_id !== userId)) {
        throw new Error(`Backup ${name} rows are invalid or belong to another owner`);
      }
      return rows;
    };
    if (selected.has("ai") && (!snapshot?.ai || typeof snapshot.ai !== "object")) {
      throw new Error("Backup AI snapshot is missing or invalid");
    }
    if (selected.has("finance") && (!snapshot?.finance || typeof snapshot.finance !== "object")) {
      throw new Error("Backup finance snapshot is missing or invalid");
    }
    const ai = selected.has("ai") ? snapshot.ai : null;
    const finance = selected.has("finance") ? snapshot.finance : null;
    if (ai) {
      requireRows(ai.memory_entries, "memory");
      requireRows(ai.conversation_messages, "conversation");
      requireRows(ai.message_receipts, "message receipt");
    }
    if (finance) {
      const books = requireRows(finance.books, "book");
      const drafts = requireRows(finance.drafts, "draft");
      const transactions = requireRows(finance.transactions, "transaction");
      requireRows(finance.audit_events, "audit");
      requireRows(finance.transaction_revisions, "revision");
      if (!Array.isArray(finance.postings) || finance.postings.length > 200_000
        || !Array.isArray(finance.user_book_state) || finance.user_book_state.length > 1) {
        throw new Error("Backup finance references are invalid");
      }
      const bookIds = new Set(books.map((book) => book.id));
      const draftIds = new Set(drafts.map((draft) => draft.id));
      const transactionIds = new Set(transactions.map((transaction) => transaction.id));
      if (books.some((book) => !bookIds.has(book.id))
        || drafts.some((draft) => draft.book_id && !bookIds.has(draft.book_id))
        || transactions.some((transaction) => !draftIds.has(transaction.draft_id)
          || (transaction.book_id && !bookIds.has(transaction.book_id)))
        || finance.postings.some((posting) => !transactionIds.has(posting.transaction_id))
        || finance.transaction_revisions.some((revision) => !transactionIds.has(revision.transaction_id))
        || finance.user_book_state.some((state) => state.user_id !== userId || !bookIds.has(state.active_book_id))) {
        throw new Error("Backup finance references do not match its owner-scoped records");
      }
      for (const book of books) {
        if (this.database.prepare("SELECT 1 FROM books WHERE id = ? AND user_id <> ?").get(book.id, userId)) {
          throw new Error("Backup contains a book ID owned by another user");
        }
      }
      for (const transaction of transactions) {
        if (this.database.prepare("SELECT 1 FROM transactions WHERE id = ? AND user_id <> ?").get(transaction.id, userId)) {
          throw new Error("Backup contains a transaction ID owned by another user");
        }
        if (transaction.source_message_id && this.database.prepare(
          "SELECT 1 FROM transactions WHERE source_message_id = ? AND user_id <> ?",
        ).get(transaction.source_message_id, userId)) {
          throw new Error("Backup contains a message reference owned by another user");
        }
      }
      for (const draft of drafts) {
        if (this.database.prepare("SELECT 1 FROM drafts WHERE id = ? AND user_id <> ?").get(draft.id, userId)) {
          throw new Error("Backup contains a draft ID owned by another user");
        }
      }
    }

    this.database.exec("BEGIN IMMEDIATE");
    try {
      if (ai) {
        for (const table of ["message_receipts", "conversation_messages", "memory_entries"]) {
          this.database.prepare(`DELETE FROM ${table} WHERE user_id = ?`).run(userId);
        }
        for (const row of ai.memory_entries) {
          this.database.prepare(`INSERT INTO memory_entries (user_id, target, content, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)`).run(row.user_id, row.target, row.content, row.created_at, row.updated_at);
        }
        for (const row of ai.conversation_messages) {
          this.database.prepare(`INSERT INTO conversation_messages (user_id, source_message_id, role, content, created_at)
            VALUES (?, ?, ?, ?, ?)`).run(row.user_id, row.source_message_id, row.role, row.content, row.created_at);
        }
        for (const row of ai.message_receipts) {
          const conflict = this.database.prepare("SELECT user_id FROM message_receipts WHERE message_id = ?").get(row.message_id);
          if (conflict && conflict.user_id !== userId) throw new Error("Backup contains a message receipt owned by another user");
          this.database.prepare(`INSERT INTO message_receipts (message_id, user_id, status, received_at, updated_at)
            VALUES (?, ?, ?, ?, ?)`).run(row.message_id, row.user_id, row.status, row.received_at, row.updated_at);
        }
      }
      if (finance) {
        this.database.prepare("DELETE FROM postings WHERE transaction_id IN (SELECT id FROM transactions WHERE user_id = ?)").run(userId);
        this.database.prepare("DELETE FROM transaction_revisions WHERE user_id = ?").run(userId);
        this.database.prepare("DELETE FROM audit_events WHERE user_id = ?").run(userId);
        this.database.prepare("DELETE FROM transactions WHERE user_id = ?").run(userId);
        this.database.prepare("DELETE FROM drafts WHERE user_id = ?").run(userId);
        this.database.prepare("DELETE FROM user_book_state WHERE user_id = ?").run(userId);
        this.database.prepare("DELETE FROM books WHERE user_id = ?").run(userId);
        for (const row of finance.books) {
          this.database.prepare(`INSERT INTO books
            (id, user_id, name, workbook_file_id, created_at, updated_at, workbook_sync_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?)`).run(row.id, row.user_id, row.name, row.workbook_file_id, row.created_at, row.updated_at, row.workbook_sync_hash);
        }
        for (const row of finance.drafts) {
          this.database.prepare(`INSERT INTO drafts
            (id, user_id, book_id, source_message_id, kind, amount_minor, currency, occurred_on, category, description, status, created_at, expires_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(row.id, row.user_id, row.book_id, row.source_message_id, row.kind, row.amount_minor, row.currency, row.occurred_on, row.category, row.description, row.status, row.created_at, row.expires_at);
        }
        for (const row of finance.transactions) {
          this.database.prepare(`INSERT INTO transactions
            (id, draft_id, user_id, book_id, source_message_id, kind, amount_minor, currency, occurred_on, category, description, approved_at, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(row.id, row.draft_id, row.user_id, row.book_id, row.source_message_id, row.kind, row.amount_minor, row.currency, row.occurred_on, row.category, row.description, row.approved_at, row.status);
        }
        for (const row of finance.postings) {
          this.database.prepare("INSERT INTO postings (transaction_id, account, amount_minor, currency) VALUES (?, ?, ?, ?)")
            .run(row.transaction_id, row.account, row.amount_minor, row.currency);
        }
        for (const row of finance.audit_events) {
          this.database.prepare("INSERT INTO audit_events (user_id, action, reference_id, created_at) VALUES (?, ?, ?, ?)")
            .run(row.user_id, row.action, row.reference_id, row.created_at);
        }
        for (const row of finance.transaction_revisions) {
          this.database.prepare(`INSERT INTO transaction_revisions (user_id, transaction_id, before_json, after_json, created_at)
            VALUES (?, ?, ?, ?, ?)`).run(row.user_id, row.transaction_id, row.before_json, row.after_json, row.created_at);
        }
        for (const row of finance.user_book_state) {
          this.database.prepare("INSERT INTO user_book_state (user_id, active_book_id, updated_at) VALUES (?, ?, ?)")
            .run(row.user_id, row.active_book_id, row.updated_at);
        }
      }
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
    return {
      ai: ai ? ai.memory_entries.length + ai.conversation_messages.length : 0,
      finance: finance ? finance.transactions.length : 0,
    };
  }

  migrateDefaultBooks() {
    const userIds = this.database.prepare(`
      SELECT user_id FROM drafts
      UNION SELECT user_id FROM transactions
      UNION SELECT user_id FROM conversation_messages
      UNION SELECT user_id FROM memory_entries
    `).all().map(({ user_id }) => user_id);
    const timestamp = this.now().toISOString();
    this.database.exec("BEGIN IMMEDIATE");
    try {
      for (const userId of userIds) {
        let book = this.database.prepare(`
          SELECT id FROM books WHERE user_id = ? ORDER BY created_at, id LIMIT 1
        `).get(userId);
        if (!book) {
          book = { id: `B-${randomUUID().slice(0, 8).toUpperCase()}` };
          this.database.prepare(`
            INSERT INTO books (id, user_id, name, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
          `).run(book.id, userId, defaultBookName, timestamp, timestamp);
        }
        this.database.prepare(`
          INSERT INTO user_book_state (user_id, active_book_id, updated_at)
          VALUES (?, ?, ?)
          ON CONFLICT(user_id) DO NOTHING
        `).run(userId, book.id, timestamp);
        this.database.prepare("UPDATE drafts SET book_id = ? WHERE user_id = ? AND book_id IS NULL")
          .run(book.id, userId);
        this.database.prepare("UPDATE transactions SET book_id = ? WHERE user_id = ? AND book_id IS NULL")
          .run(book.id, userId);
      }
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
  }

  getBook(userId, bookId) {
    return this.database.prepare(`
      SELECT id, user_id, name, workbook_file_id, workbook_sync_hash, created_at, updated_at
      FROM books WHERE id = ? AND user_id = ?
    `).get(bookId, userId) ?? null;
  }

  getActiveBook(userId) {
    const active = this.database.prepare(`
      SELECT books.id, books.user_id, books.name, books.workbook_file_id, books.workbook_sync_hash, books.created_at, books.updated_at
      FROM user_book_state JOIN books ON books.id = user_book_state.active_book_id AND books.user_id = user_book_state.user_id
      WHERE user_book_state.user_id = ?
    `).get(userId);
    if (active) return active;
    const existing = this.database.prepare("SELECT id FROM books WHERE user_id = ? ORDER BY created_at, id LIMIT 1").get(userId);
    if (!existing) return null;
    this.database.prepare("INSERT INTO user_book_state (user_id, active_book_id, updated_at) VALUES (?, ?, ?)").run(userId, existing.id, this.now().toISOString());
    return this.getBook(userId, existing.id);
  }

  updateBookWorkbook(userId, bookId, fileId) {
    const book = this.getBook(userId, bookId);
    if (!book) throw new Error("Book not found in this user's workspace");
    if (typeof fileId !== "string" || !/^F-[A-F0-9]{8}$/.test(fileId)) {
      throw new Error("A valid workspace workbook file ID is required");
    }
    this.database.prepare("UPDATE books SET workbook_file_id = ?, workbook_sync_hash = NULL, updated_at = ? WHERE id = ? AND user_id = ?")
      .run(fileId, this.now().toISOString(), bookId, userId);
    this.database.prepare(`
      INSERT INTO audit_events (user_id, action, reference_id, created_at)
      VALUES (?, 'book_workbook_linked', ?, ?)
    `).run(userId, bookId, this.now().toISOString());
    return this.getBook(userId, bookId);
  }

  markBookWorkbookSynced(userId, bookId, syncHash) {
    if (typeof syncHash !== "string" || !/^[a-f0-9]{64}$/.test(syncHash)) {
      throw new Error("Workbook sync hash must be a SHA-256 digest");
    }
    const result = this.database.prepare(`
      UPDATE books SET workbook_sync_hash = ?, updated_at = ? WHERE id = ? AND user_id = ?
    `).run(syncHash, this.now().toISOString(), bookId, userId);
    if (Number(result.changes) !== 1) throw new Error("Book not found in this user's workspace");
    return this.getBook(userId, bookId);
  }

  moveTransaction(userId, transactionId, targetBookId) {
    const targetBook = this.getBook(userId, targetBookId);
    if (!targetBook) throw new Error("Target book not found in this user's workspace");
    const normalizedId = transactionId.toUpperCase();
    const current = this.database.prepare(`
      SELECT id, draft_id, book_id, kind, amount_minor, currency, occurred_on, category, description, status
      FROM transactions WHERE id = ? AND user_id = ?
    `).get(normalizedId, userId);
    if (!current) throw new Error("Transaction not found in this user's records");
    if (current.book_id === targetBook.id) {
      return { id: current.id, from_book_id: current.book_id, book_id: targetBook.id, book_name: targetBook.name, unchanged: true };
    }

    const sourceBook = this.getBook(userId, current.book_id);
    const timestamp = this.now().toISOString();
    const before = JSON.stringify({ ...current, book_name: sourceBook?.name ?? defaultBookName });
    const after = JSON.stringify({ ...current, book_id: targetBook.id, book_name: targetBook.name });
    this.database.exec("BEGIN IMMEDIATE");
    try {
      this.database.prepare("UPDATE transactions SET book_id = ? WHERE id = ? AND user_id = ?")
        .run(targetBook.id, normalizedId, userId);
      this.database.prepare("UPDATE drafts SET book_id = ? WHERE id = ? AND user_id = ?")
        .run(targetBook.id, current.draft_id, userId);
      this.database.prepare(`
        INSERT INTO transaction_revisions (user_id, transaction_id, before_json, after_json, created_at)
        VALUES (?, ?, ?, ?, ?)
      `).run(userId, normalizedId, before, after, timestamp);
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'transaction_moved', ?, ?)
      `).run(userId, normalizedId, timestamp);
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
    return {
      id: normalizedId,
      from_book_id: current.book_id,
      book_id: targetBook.id,
      book_name: targetBook.name,
      status: current.status,
    };
  }

  listBookTransactions(userId, bookId, limit = 10000) {
    const book = this.getBook(userId, bookId);
    if (!book) throw new Error("Book not found in this user's workspace");
    if (!Number.isInteger(limit) || limit < 1 || limit > 10000) {
      throw new Error("Book transaction limit must be between 1 and 10000");
    }
    return this.database.prepare(`
      SELECT id, book_id, kind, amount_minor, currency, occurred_on, category, description, approved_at, status
      FROM transactions WHERE user_id = ? AND book_id = ? AND status = 'active'
      ORDER BY occurred_on, approved_at LIMIT ?
    `).all(userId, book.id, limit);
  }

  listBooks(userId) {
    return this.database.prepare(`
      SELECT books.id, books.name, books.created_at, books.updated_at,
        CASE WHEN user_book_state.active_book_id = books.id THEN 1 ELSE 0 END AS is_active,
        (SELECT COUNT(*) FROM transactions WHERE transactions.user_id = books.user_id
          AND transactions.book_id = books.id AND transactions.status = 'active') AS transaction_count
      FROM books LEFT JOIN user_book_state ON user_book_state.user_id = books.user_id
      WHERE books.user_id = ? ORDER BY books.created_at, books.id
    `).all(userId);
  }

  createBook(userId, name) {
    if (typeof name !== "string") throw new Error("Book name must be text");
    const normalizedName = name.trim();
    if (!normalizedName || normalizedName.length > 80 || /[\u0000-\u001F\u007F]/.test(normalizedName)) {
      throw new Error("Book name must contain 1-80 printable characters");
    }
    const timestamp = this.now().toISOString();
    const id = `B-${randomUUID().slice(0, 8).toUpperCase()}`;
    try {
      this.database.prepare(`
        INSERT INTO books (id, user_id, name, created_at, updated_at) VALUES (?, ?, ?, ?, ?)
      `).run(id, userId, normalizedName, timestamp, timestamp);
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'book_created', ?, ?)
      `).run(userId, id, timestamp);
    } catch (error) {
      if (String(error.message).includes("UNIQUE constraint failed")) throw new Error("A book with this name already exists");
      throw error;
    }
    return this.getBook(userId, id);
  }

  selectBook(userId, bookId) {
    const book = this.getBook(userId, bookId);
    if (!book) throw new Error("Book not found in this user's workspace");
    const timestamp = this.now().toISOString();
    this.database.prepare(`
      INSERT INTO user_book_state (user_id, active_book_id, updated_at) VALUES (?, ?, ?)
      ON CONFLICT(user_id) DO UPDATE SET active_book_id = excluded.active_book_id, updated_at = excluded.updated_at
    `).run(userId, book.id, timestamp);
    this.database.prepare(`
      INSERT INTO audit_events (user_id, action, reference_id, created_at)
      VALUES (?, 'book_selected', ?, ?)
    `).run(userId, book.id, timestamp);
    return book;
  }

  reserveMessage(messageId, userId) {
    const now = this.now().toISOString();
    const result = this.database.prepare(`
      INSERT INTO message_receipts (message_id, user_id, status, received_at, updated_at)
      VALUES (?, ?, 'queued', ?, ?)
      ON CONFLICT(message_id) DO UPDATE SET status = 'queued', updated_at = excluded.updated_at
      WHERE message_receipts.status = 'failed'
    `).run(messageId, userId, now, now);
    return Number(result.changes) === 1;
  }

  finishMessage(messageId, status) {
    if (!['completed', 'failed'].includes(status)) {
      throw new Error("Invalid message receipt status");
    }
    this.database.prepare("UPDATE message_receipts SET status = ?, updated_at = ? WHERE message_id = ?")
      .run(status, this.now().toISOString(), messageId);
  }

  saveConversationTurn(userId, sourceMessageId, userContent, assistantContent) {
    if (typeof sourceMessageId !== "string" || !sourceMessageId) {
      throw new Error("Conversation turns require a source message ID");
    }
    const createdAt = this.now().toISOString();
    const insertMessage = this.database.prepare(`
      INSERT INTO conversation_messages (user_id, source_message_id, role, content, created_at)
      VALUES (?, ?, ?, ?, ?)
      ON CONFLICT(user_id, source_message_id, role)
      DO UPDATE SET content = excluded.content, created_at = excluded.created_at
    `);
    this.database.exec("BEGIN IMMEDIATE");
    try {
      insertMessage.run(userId, sourceMessageId, "user", String(userContent ?? "").slice(0, 12_000), createdAt);
      insertMessage.run(userId, sourceMessageId, "assistant", String(assistantContent ?? "").slice(0, 12_000), createdAt);
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
  }

  getRecentConversation(userId, limit = 12) {
    if (!Number.isInteger(limit) || limit < 1 || limit > 40) {
      throw new Error("Conversation history limit must be between 1 and 40 messages");
    }
    return this.database.prepare(`
      SELECT role, content FROM conversation_messages
      WHERE user_id = ? ORDER BY created_at DESC, id DESC LIMIT ?
    `).all(userId, limit).reverse();
  }

  searchConversation(userId, query, limit = 6) {
    if (typeof query !== "string" || !query.trim()) throw new Error("History search query cannot be empty");
    if (!Number.isInteger(limit) || limit < 1 || limit > 10) throw new Error("History search limit must be between 1 and 10");
    const terms = query.normalize("NFKC").match(/[\p{L}\p{N}]{2,}/gu)?.slice(0, 8) ?? [];
    if (terms.length === 0) return [];
    const rows = this.database.prepare(`
      SELECT role, content, created_at FROM conversation_messages
      WHERE user_id = ? AND (${terms.map(() => "content LIKE ? COLLATE NOCASE").join(" OR ")})
      ORDER BY (${terms.map(() => "CASE WHEN content LIKE ? COLLATE NOCASE THEN 1 ELSE 0 END").join(" + ")}) DESC,
        created_at DESC, id DESC LIMIT ?
    `).all(
      userId,
      ...terms.map((term) => `%${term}%`),
      ...terms.map((term) => `%${term}%`),
      limit,
    );
    return rows.map((row) => ({
      role: row.role,
      content: row.content.slice(0, 1200),
      createdAt: row.created_at,
    }));
  }

  getMemoryState(userId) {
    const entries = this.database.prepare(`
      SELECT id, target, content, created_at, updated_at
      FROM memory_entries WHERE user_id = ? ORDER BY target, id
    `).all(userId);
    const state = {};
    for (const target of ["memory", "user"]) {
      const targetEntries = entries.filter((entry) => entry.target === target);
      const usedCharacters = targetEntries.reduce((total, entry) => total + [...entry.content].length, 0);
      state[target] = {
        entries: targetEntries,
        usedCharacters,
        maxCharacters: memoryCharacterLimits[target],
      };
    }
    return state;
  }

  updateMemory(userId, target, operations) {
    if (!Object.hasOwn(memoryCharacterLimits, target)) {
      throw new Error("Memory target must be memory or user");
    }
    if (!Array.isArray(operations) || operations.length < 1 || operations.length > 10) {
      throw new Error("Provide 1-10 memory operations");
    }

    const timestamp = this.now().toISOString();
    const entries = this.database.prepare(`
      SELECT content, created_at FROM memory_entries
      WHERE user_id = ? AND target = ? ORDER BY id
    `).all(userId, target).map((entry) => ({ ...entry, updated_at: timestamp }));
    const changes = [];

    for (const operation of operations) {
      if (operation.action === "add") {
        const content = validateMemoryEntry(operation.content);
        if (entries.some((entry) => entry.content.toLocaleLowerCase() === content.toLocaleLowerCase())) {
          changes.push({ action: "add", status: "duplicate" });
          continue;
        }
        entries.push({ content, created_at: timestamp, updated_at: timestamp });
        changes.push({ action: "add", status: "added" });
        continue;
      }

      if (!["patch", "remove"].includes(operation.action)
        || typeof operation.old_text !== "string" || operation.old_text.trim().length < 3) {
        throw new Error("Memory operations must be add, patch, or remove with a unique old_text");
      }
      const matches = entries
        .map((entry, index) => ({ entry, index }))
        .filter(({ entry }) => entry.content.includes(operation.old_text));
      if (matches.length !== 1) {
        throw new Error(matches.length === 0
          ? "Memory substring not found"
          : "Memory substring matches multiple entries; provide a more specific phrase");
      }
      const { entry, index } = matches[0];
      const occurrences = entry.content.split(operation.old_text).length - 1;
      if (occurrences !== 1) throw new Error("Memory substring appears more than once; provide a unique phrase");

      if (operation.action === "remove") {
        entries.splice(index, 1);
        changes.push({ action: "remove", status: "removed" });
      } else {
        const replacement = validateMemoryEntry(operation.content);
        const updated = validateMemoryEntry(entry.content.replace(operation.old_text, replacement));
        if (entries.some((candidate, candidateIndex) => candidateIndex !== index
          && candidate.content.toLocaleLowerCase() === updated.toLocaleLowerCase())) {
          throw new Error("Memory update would create a duplicate entry");
        }
        entries[index] = { ...entry, content: updated, updated_at: timestamp };
        changes.push({ action: "patch", status: "updated" });
      }
    }

    const usedCharacters = entries.reduce((total, entry) => total + [...entry.content].length, 0);
    const maxCharacters = memoryCharacterLimits[target];
    if (usedCharacters > maxCharacters) {
      throw new Error(`Memory ${target} limit is ${maxCharacters} characters; current batch would use ${usedCharacters}`);
    }

    this.database.exec("BEGIN IMMEDIATE");
    try {
      this.database.prepare("DELETE FROM memory_entries WHERE user_id = ? AND target = ?").run(userId, target);
      const insert = this.database.prepare(`
        INSERT INTO memory_entries (user_id, target, content, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?)
      `);
      for (const entry of entries) {
        insert.run(userId, target, entry.content, entry.created_at, entry.updated_at);
      }
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'memory_updated', ?, ?)
      `).run(userId, target, timestamp);
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
    return { target, changes, usedCharacters, maxCharacters, entries };
  }

  createDraft(userId, input, { sourceMessageId = null, defaultCurrency = "IDR", defaultDate = null, bookId = null } = {}) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const normalized = normalizeTransactionInput(
      input,
      defaultCurrency,
      defaultDate || this.now().toISOString().slice(0, 10),
    );
    const { kind, amountMinor, currency, occurredOn, category, description } = normalized;

    if (sourceMessageId) {
      const existing = this.database.prepare(`
        SELECT id, user_id, kind, amount_minor, currency, occurred_on, category, description, status
        FROM drafts WHERE user_id = ? AND source_message_id = ?
      `).get(userId, sourceMessageId);
      if (existing) return existing;
    }

    const id = `D-${randomUUID().slice(0, 8).toUpperCase()}`;
    const createdAt = this.now();
    const expiresAt = new Date(createdAt.getTime() + 24 * 60 * 60 * 1000);
    this.database.exec("BEGIN IMMEDIATE");
    try {
      this.database.prepare(`
        INSERT INTO drafts
          (id, user_id, book_id, source_message_id, kind, amount_minor, currency, occurred_on, category,
           description, status, created_at, expires_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
      `).run(
        id, userId, selectedBook.id, sourceMessageId, kind, amountMinor, currency, occurredOn, category,
        description, createdAt.toISOString(), expiresAt.toISOString(),
      );
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'draft_created', ?, ?)
      `).run(userId, id, createdAt.toISOString());
      this.database.exec("COMMIT");
    } catch (error) {
      this.database.exec("ROLLBACK");
      throw error;
    }
    return this.getDraft(userId, id);
  }

  recordTransaction(userId, input, { sourceMessageId = null, defaultCurrency = "IDR", defaultDate = null, bookId = null } = {}) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const normalized = normalizeTransactionInput(
      input,
      defaultCurrency,
      defaultDate || this.now().toISOString().slice(0, 10),
    );
    const { kind, amountMinor, currency, occurredOn, category, description } = normalized;

    if (sourceMessageId) {
      const existing = this.database.prepare(`
        SELECT * FROM transactions WHERE user_id = ? AND source_message_id = ?
      `).get(userId, sourceMessageId);
      if (existing) return {
        ...existing,
        book_name: this.getBook(userId, existing.book_id)?.name ?? defaultBookName,
        already_recorded: true,
      };
      const oldDraft = this.database.prepare(`
        SELECT id, book_id, status FROM drafts WHERE user_id = ? AND source_message_id = ?
      `).get(userId, sourceMessageId);
      if (oldDraft?.status === "pending") return this.approveDraft(userId, oldDraft.id, oldDraft.book_id);
    }

    const transactionId = `T-${randomUUID().slice(0, 8).toUpperCase()}`;
    const draftId = `D-${randomUUID().slice(0, 8).toUpperCase()}`;
    const recordedAt = this.now();
    const timestamp = recordedAt.toISOString();
    const expiresAt = new Date(recordedAt.getTime() + 24 * 60 * 60 * 1000).toISOString();
    this.database.exec("BEGIN IMMEDIATE");
    try {
      this.database.prepare(`
        INSERT INTO drafts
          (id, user_id, book_id, source_message_id, kind, amount_minor, currency, occurred_on, category,
           description, status, created_at, expires_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'approved', ?, ?)
      `).run(
        draftId, userId, selectedBook.id, sourceMessageId, kind, amountMinor, currency, occurredOn,
        category, description, timestamp, expiresAt,
      );
      this.database.prepare(`
        INSERT INTO transactions
          (id, draft_id, user_id, book_id, source_message_id, kind, amount_minor, currency, occurred_on,
           category, description, approved_at, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active')
      `).run(
        transactionId, draftId, userId, selectedBook.id, sourceMessageId, kind, amountMinor,
        currency, occurredOn, category, description, timestamp,
      );
      this.insertBalancedPostings(transactionId, normalized);
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'transaction_recorded', ?, ?)
      `).run(userId, transactionId, timestamp);
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
    return {
      ...this.database.prepare("SELECT * FROM transactions WHERE id = ? AND user_id = ?").get(transactionId, userId),
      book_name: selectedBook.name,
      already_recorded: false,
    };
  }

  insertBalancedPostings(transactionId, transaction) {
    const amountMinor = transaction.amountMinor ?? transaction.amount_minor;
    const assetAmount = transaction.kind === "expense" ? -amountMinor : amountMinor;
    const categoryAmount = -assetAmount;
    const categoryAccount = `${transaction.kind === "expense" ? "Expenses" : "Income"}:${transaction.category}`;
    const insertPosting = this.database.prepare(`
      INSERT INTO postings (transaction_id, account, amount_minor, currency) VALUES (?, ?, ?, ?)
    `);
    insertPosting.run(transactionId, "Assets:Cash", assetAmount, transaction.currency);
    insertPosting.run(transactionId, categoryAccount, categoryAmount, transaction.currency);
    const postingTotal = this.database.prepare(
      "SELECT SUM(amount_minor) AS total FROM postings WHERE transaction_id = ?",
    ).get(transactionId).total;
    if (postingTotal !== 0) throw new Error("Ledger postings are not balanced");
  }

  listRecentTransactions(userId, limit = 10, bookId = null) {
    if (!Number.isInteger(limit) || limit < 1 || limit > 25) {
      throw new Error("Transaction list limit must be between 1 and 25");
    }
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    return this.database.prepare(`
      SELECT id, book_id, kind, amount_minor, currency, occurred_on, category, description, approved_at
      FROM transactions WHERE user_id = ? AND book_id = ? AND status = 'active'
      ORDER BY approved_at DESC LIMIT ?
    `).all(userId, selectedBook.id, limit);
  }

  updateTransaction(userId, transactionId, updates, bookId = null) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const normalizedId = transactionId.toUpperCase();
    const current = this.database.prepare(`
      SELECT * FROM transactions WHERE id = ? AND user_id = ? AND status = 'active'
        AND book_id = ?
    `).get(normalizedId, userId, selectedBook.id);
    if (!current) throw new Error("Transaction not found in this user's active records");
    if (!updates || typeof updates !== "object" || Array.isArray(updates) || Object.keys(updates).length === 0) {
      throw new Error("Provide at least one transaction field to update");
    }
    const digits = getCurrencyDigits(current.currency);
    const normalized = normalizeTransactionInput({
      kind: updates.kind ?? current.kind,
      amount: updates.amount ?? current.amount_minor / (10 ** digits),
      currency: updates.currency ?? current.currency,
      date: updates.date ?? current.occurred_on,
      category: updates.category ?? current.category,
      description: updates.description ?? current.description,
    }, current.currency, current.occurred_on);
    const updatedAt = this.now().toISOString();
    const before = JSON.stringify({
      kind: current.kind,
      amount_minor: current.amount_minor,
      currency: current.currency,
      occurred_on: current.occurred_on,
      category: current.category,
      description: current.description,
    });
    const after = JSON.stringify({
      kind: normalized.kind,
      amount_minor: normalized.amountMinor,
      currency: normalized.currency,
      occurred_on: normalized.occurredOn,
      category: normalized.category,
      description: normalized.description,
    });

    this.database.exec("BEGIN IMMEDIATE");
    try {
      this.database.prepare(`
        UPDATE transactions
        SET kind = ?, amount_minor = ?, currency = ?, occurred_on = ?, category = ?, description = ?
        WHERE id = ? AND user_id = ? AND status = 'active'
      `).run(
        normalized.kind, normalized.amountMinor, normalized.currency, normalized.occurredOn,
        normalized.category, normalized.description, normalizedId, userId,
      );
      this.database.prepare(`
        UPDATE drafts
        SET kind = ?, amount_minor = ?, currency = ?, occurred_on = ?, category = ?, description = ?
        WHERE id = ? AND user_id = ?
      `).run(
        normalized.kind, normalized.amountMinor, normalized.currency, normalized.occurredOn,
        normalized.category, normalized.description, current.draft_id, userId,
      );
      this.database.prepare("DELETE FROM postings WHERE transaction_id = ?").run(normalizedId);
      this.insertBalancedPostings(normalizedId, normalized);
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'transaction_updated', ?, ?)
      `).run(userId, normalizedId, updatedAt);
      this.database.prepare(`
        INSERT INTO transaction_revisions (user_id, transaction_id, before_json, after_json, created_at)
        VALUES (?, ?, ?, ?, ?)
      `).run(userId, normalizedId, before, after, updatedAt);
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
    return this.database.prepare("SELECT * FROM transactions WHERE id = ? AND user_id = ?").get(normalizedId, userId);
  }

  voidTransaction(userId, transactionId, bookId = null) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const normalizedId = transactionId.toUpperCase();
    this.database.exec("BEGIN IMMEDIATE");
    try {
      const result = this.database.prepare(`
        UPDATE transactions SET status = 'voided'
        WHERE id = ? AND user_id = ? AND book_id = ? AND status = 'active'
      `).run(normalizedId, userId, selectedBook.id);
      if (Number(result.changes) !== 1) {
        throw new Error("Transaction not found in this user's active records");
      }
      this.database.prepare(`
        INSERT INTO postings (transaction_id, account, amount_minor, currency)
        SELECT transaction_id, account, -amount_minor, currency
        FROM postings WHERE transaction_id = ?
      `).run(normalizedId);
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'transaction_voided', ?, ?)
      `).run(userId, normalizedId, this.now().toISOString());
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
    return { id: normalizedId, status: "voided" };
  }

  getDraft(userId, draftId, bookId = null) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) return null;
    return this.database.prepare(`
      SELECT id, user_id, book_id, source_message_id, kind, amount_minor, currency, occurred_on, category,
             description, status, created_at, expires_at
      FROM drafts WHERE id = ? AND user_id = ? AND book_id = ?
    `).get(draftId.toUpperCase(), userId, selectedBook.id) ?? null;
  }

  listPendingDrafts(userId, bookId = null) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    return this.database.prepare(`
      SELECT id, book_id, kind, amount_minor, currency, occurred_on, category, description, created_at, expires_at
      FROM drafts WHERE user_id = ? AND book_id = ? AND status = 'pending' AND expires_at > ?
      ORDER BY created_at ASC
    `).all(userId, selectedBook.id, this.now().toISOString());
  }

  approveDraft(userId, draftId, bookId = null) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const normalizedId = draftId.toUpperCase();
    const approvedAt = this.now().toISOString();
    this.database.exec("BEGIN IMMEDIATE");
    try {
      const draft = this.getDraft(userId, normalizedId, selectedBook.id);
      if (draft?.status === "approved") {
        const existing = this.database.prepare("SELECT * FROM transactions WHERE draft_id = ?").get(normalizedId);
        if (existing?.status === "voided") throw new Error("Transaction has been voided");
        this.database.exec("COMMIT");
        return { ...existing, already_approved: true };
      }
      if (!draft || draft.status !== "pending") {
        throw new Error("Draft not found or no longer pending");
      }
      if (draft.expires_at <= approvedAt) {
        this.database.prepare("UPDATE drafts SET status = 'expired' WHERE id = ?").run(normalizedId);
        this.database.exec("COMMIT");
        throw new Error("Draft has expired; create it again");
      }

      const transactionId = `T-${randomUUID().slice(0, 8).toUpperCase()}`;
      this.database.prepare(`
        INSERT INTO transactions
          (id, draft_id, user_id, book_id, source_message_id, kind, amount_minor, currency, occurred_on,
           category, description, approved_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(
        transactionId, normalizedId, userId, selectedBook.id, draft.source_message_id, draft.kind,
        draft.amount_minor, draft.currency, draft.occurred_on, draft.category,
        draft.description, approvedAt,
      );

      this.insertBalancedPostings(transactionId, draft);

      this.database.prepare("UPDATE drafts SET status = 'approved' WHERE id = ?").run(normalizedId);
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'transaction_approved', ?, ?)
      `).run(userId, transactionId, approvedAt);
      this.database.exec("COMMIT");
      return this.database.prepare("SELECT * FROM transactions WHERE id = ? AND user_id = ?")
        .get(transactionId, userId);
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
  }

  cancelDraft(userId, draftId, bookId = null) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const normalizedId = draftId.toUpperCase();
    this.database.exec("BEGIN IMMEDIATE");
    try {
      const result = this.database.prepare(`
        UPDATE drafts SET status = 'cancelled'
        WHERE id = ? AND user_id = ? AND book_id = ? AND status = 'pending'
      `).run(normalizedId, userId, selectedBook.id);
      if (Number(result.changes) !== 1) {
        throw new Error("Draft not found or no longer pending");
      }
      this.database.prepare(`
        INSERT INTO audit_events (user_id, action, reference_id, created_at)
        VALUES (?, 'draft_cancelled', ?, ?)
      `).run(userId, normalizedId, this.now().toISOString());
      this.database.exec("COMMIT");
    } catch (error) {
      if (this.database.isTransaction) this.database.exec("ROLLBACK");
      throw error;
    }
    return this.getDraft(userId, normalizedId, selectedBook.id);
  }

  getSummary(userId, month = this.now().toISOString().slice(0, 7), bookId = null) {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
      throw new Error("Month must use YYYY-MM format");
    }
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const rows = this.database.prepare(`
      SELECT kind, currency, category, SUM(amount_minor) AS total_minor, COUNT(*) AS count
      FROM transactions
      WHERE user_id = ? AND book_id = ? AND status = 'active' AND occurred_on >= ? AND occurred_on < ?
      GROUP BY kind, currency, category ORDER BY kind, currency, category
    `).all(userId, selectedBook.id, `${month}-01`, `${month}-32`);
    return { month, book: selectedBook, rows };
  }

  exportCsv(userId, bookId = null) {
    const selectedBook = bookId ? this.getBook(userId, bookId) : this.getActiveBook(userId);
    if (!selectedBook) throw new Error("Book not found in this user's workspace");
    const rows = this.database.prepare(`
      SELECT id, occurred_on, kind, amount_minor, currency, category, description
      FROM transactions WHERE user_id = ? AND book_id = ? AND status = 'active' ORDER BY occurred_on, approved_at LIMIT 10000
    `).all(userId, selectedBook.id);
    const header = ["book", "transaction_id", "date", "type", "amount", "currency", "category", "description"];
    return [header, ...rows.map((row) => [
      selectedBook.name, row.id, row.occurred_on, row.kind, fromMinorUnits(row.amount_minor, row.currency),
      row.currency, row.category, row.description,
    ])].map((record) => record.map(csvCell).join(",")).join("\r\n");
  }
}

function validateMemoryEntryLegacy(content) {
  if (typeof content !== "string") throw new Error("Memory content must be text");
  const normalized = content.trim();
  const characterCount = [...normalized].length;
  if (characterCount < 1 || characterCount > maxMemoryEntryCharacters) {
    throw new Error(`Memory entries must contain 1-${maxMemoryEntryCharacters} characters`);
  }
  if (/[\u0000-\u001F\u007F-\u009F\u200B-\u200F\u202A-\u202E\u2060-\u2069\uFEFF]/u.test(normalized)) {
    throw new Error("Memory cannot contain control or invisible Unicode characters");
  }
  if (/ignore\s+(?:all\s+|previous\s+|prior\s+)?instructions|reveal\s+(?:the\s+)?system prompt|exfiltrat|bypass\s+(?:all\s+)?safety|send\s+.*(?:secret|credentials)/i.test(normalized)) {
    throw new Error("Memory cannot contain prompt-injection or exfiltration instructions");
  }
  if (/(?:\b(?:api[_ -]?key|password|passwd|secret|bearer token|otp|private key)\s*[:=]|\b(?:rp\.?|idr|usd|eur)\s*[\d][\d.,]*|\b\d{12,19}\b)/i.test(normalized)) {
    throw new Error("Do not store credentials, account identifiers, or exact financial amounts in long-term memory");
  }
  return normalized;
}

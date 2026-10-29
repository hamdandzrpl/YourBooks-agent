// Pure analytics for finance books. No dependencies, no I/O: everything here is
// derived from the ledger rows returned by FinanceStore.listBookTransactions.
//
// The workbook builder and the agent tools both consume `analyzeBook` so the
// numbers the assistant talks about always match the numbers in the spreadsheet.

export const MONTH_NAMES = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember",
];
export const MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
export const WEEKDAY_NAMES = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

export function currencyDigits(currency) {
  return new Intl.NumberFormat("en", { style: "currency", currency })
    .resolvedOptions().maximumFractionDigits;
}

/** Minor units (integer) -> major units (float) for the given currency. */
export function toMajor(amountMinor, currency) {
  return Number(amountMinor) / (10 ** currencyDigits(currency));
}

export function monthKey(date) {
  return String(date).slice(0, 7);
}

export function monthLabel(key) {
  const [year, month] = key.split("-");
  return `${MONTH_SHORT[Number(month) - 1] ?? month} ${year}`;
}

export function monthLabelLong(key) {
  const [year, month] = key.split("-");
  return `${MONTH_NAMES[Number(month) - 1] ?? month} ${year}`;
}

export function formatAmount(value, { currency, digits = currencyDigits(currency), compact = false } = {}) {
  const sign = value < 0 ? "-" : "";
  const absolute = Math.abs(Number(value) || 0);
  if (compact && absolute >= 1_000_000_000) return `${sign}${(absolute / 1_000_000_000).toFixed(1)} M`;
  if (compact && absolute >= 1_000_000) return `${sign}${(absolute / 1_000_000).toFixed(1)} jt`;
  if (compact && absolute >= 10_000) return `${sign}${(absolute / 1000).toFixed(0)} rb`;
  const text = absolute.toLocaleString("id-ID", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
  return `${sign}${text}`;
}

/** Pick the currency that carries the most activity; ties go to the configured default. */
export function primaryCurrency(transactions, fallback = "IDR") {
  const totals = new Map();
  for (const transaction of transactions) {
    const current = totals.get(transaction.currency) ?? { count: 0, amount: 0 };
    current.count += 1;
    current.amount += toMajor(transaction.amount_minor, transaction.currency);
    totals.set(transaction.currency, current);
  }
  if (!totals.size) return fallback.toUpperCase();
  return [...totals.entries()]
    .sort((a, b) => b[1].count - a[1].count || b[1].amount - a[1].amount)[0][0];
}

export function formatPercent(value, digits = 1) {
  if (!Number.isFinite(value)) return "—";
  return `${value >= 0 ? "" : "-"}${Math.abs(value).toFixed(digits).replace(".", ",")}%`;
}

function round(value, digits = 2) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function safePercent(part, whole) {
  if (!whole) return null;
  return round((part / whole) * 100);
}

function sumBy(rows, predicate) {
  return rows.reduce((total, row) => (predicate(row) ? total + row.amount : total), 0);
}

/** Sparkline-ish bar used in text cells (works offline, no chart engine). */
export function bar(value, max, width = 12) {
  if (!max || max <= 0 || value <= 0) return "";
  const filled = Math.max(1, Math.round((value / max) * width));
  return "█".repeat(Math.min(width, filled));
}

/**
 * Turn raw ledger rows into a full analytical picture of a book.
 *
 * @param {object} options
 * @param {Array} options.transactions raw rows (amount_minor, currency, kind, ...)
 * @param {string} options.currency currency to report in
 * @param {string} [options.bookName]
 * @param {string} [options.today] ISO date (YYYY-MM-DD) used for "current month"
 * @returns {object} analysis
 */
export function analyzeBook({ transactions = [], currency, bookName = "", today }) {
  const rows = transactions
    .filter((row) => row && row.status !== "voided")
    .map((row) => ({
      id: row.id,
      date: String(row.occurred_on),
      month: monthKey(row.occurred_on),
      kind: row.kind,
      category: row.category || "Lain-lain",
      description: row.description || "",
      currency: row.currency,
      amount: toMajor(row.amount_minor, row.currency),
      approvedAt: row.approved_at ?? null,
    }))
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  const digits = currencyDigits(currency);
  const otherCurrencies = [...new Set(rows.map((row) => row.currency).filter((code) => code !== currency))];
  const scoped = rows.filter((row) => row.currency === currency);

  const income = sumBy(scoped, (row) => row.kind === "income");
  const expense = sumBy(scoped, (row) => row.kind === "expense");
  const net = income - expense;
  const count = scoped.length;
  const dates = [...new Set(scoped.map((row) => row.date))];
  const activeDays = dates.length;
  const avgTicket = count ? round((income + expense) / count) : 0;

  // ---- monthly series (dense: empty months inside the range are kept) ----
  const monthMap = new Map();
  for (const row of scoped) {
    const entry = monthMap.get(row.month) ?? { key: row.month, income: 0, expense: 0, count: 0 };
    if (row.kind === "income") entry.income += row.amount;
    else entry.expense += row.amount;
    entry.count += 1;
    monthMap.set(row.month, entry);
  }
  let months = [...monthMap.values()].sort((a, b) => (a.key < b.key ? -1 : 1));
  if (months.length > 1) {
    months = fillMonthGaps(months);
  }
  const monthRows = months.map((entry) => ({
    ...entry,
    income: round(entry.income),
    expense: round(entry.expense),
    net: round(entry.income - entry.expense),
    savingsRate: safePercent(entry.income - entry.expense, entry.income),
  }));

  // ---- categories ----
  const categoryMap = new Map();
  for (const row of scoped) {
    const key = `${row.kind}:${row.category}`;
    const entry = categoryMap.get(key) ?? {
      name: row.category, kind: row.kind, total: 0, count: 0, months: new Set(),
    };
    entry.total += row.amount;
    entry.count += 1;
    entry.months.add(row.month);
    categoryMap.set(key, entry);
  }
  const categories = [...categoryMap.values()]
    .map((entry) => ({
      name: entry.name,
      kind: entry.kind,
      total: round(entry.total),
      count: entry.count,
      monthsActive: entry.months.size,
      average: round(entry.total / entry.count),
      share: safePercent(entry.total, entry.kind === "income" ? income : expense),
    }))
    .sort((a, b) => b.total - a.total);

  const expenseCategories = categories.filter((entry) => entry.kind === "expense");
  const incomeCategories = categories.filter((entry) => entry.kind === "income");

  // ---- weekday distribution ----
  const weekdayMap = new Map();
  for (const row of scoped) {
    const day = new Date(`${row.date}T00:00:00.000Z`).getUTCDay();
    const entry = weekdayMap.get(day) ?? { day, total: 0, count: 0 };
    entry.total += row.amount;
    entry.count += 1;
    weekdayMap.set(day, entry);
  }
  const weekdays = WEEKDAY_NAMES.map((name, day) => ({
    day,
    name,
    total: round(weekdayMap.get(day)?.total ?? 0),
    count: weekdayMap.get(day)?.count ?? 0,
  }));

  // ---- top transactions & outliers ----
  const expenses = scoped.filter((row) => row.kind === "expense").sort((a, b) => b.amount - a.amount);
  const topExpenses = expenses.slice(0, 8).map((row) => ({
    id: row.id, date: row.date, category: row.category, description: row.description, amount: round(row.amount),
  }));
  const expenseAmounts = expenses.map((row) => row.amount).sort((a, b) => a - b);
  const medianExpense = expenseAmounts.length
    ? expenseAmounts[Math.floor(expenseAmounts.length / 2)]
    : 0;
  const anomalies = medianExpense > 0
    ? expenses
      .filter((row) => row.amount >= medianExpense * 3 && row.amount >= 100_000)
      .slice(0, 6)
      .map((row) => ({
        id: row.id, date: row.date, category: row.category, description: row.description,
        amount: round(row.amount), ratio: round(row.amount / medianExpense, 1),
      }))
    : [];

  // ---- growth vs previous month ----
  const currentMonth = today ? monthKey(today) : (monthRows.at(-1)?.key ?? null);
  const currentIndex = monthRows.findIndex((entry) => entry.key === currentMonth);
  const current = currentIndex >= 0 ? monthRows[currentIndex] : null;
  const previous = currentIndex > 0 ? monthRows[currentIndex - 1] : null;
  const growth = (now, before) => (before ? safePercent(now - before, Math.abs(before)) : null);
  const deltas = {
    income: previous ? growth(current.income, previous.income) : null,
    expense: previous ? growth(current.expense, previous.expense) : null,
    net: previous ? growth(current.net, previous.net) : null,
  };

  const savingsRate = safePercent(net, income);
  const burnRate = activeDays ? round(expense / activeDays) : 0;
  const runwayMonths = expense > 0 && net > 0 ? round(net / (expense / Math.max(1, monthRows.length)), 1) : null;

  const bestMonth = monthRows.reduce((best, entry) => (!best || entry.net > best.net ? entry : best), null);
  const worstMonth = monthRows.reduce((worst, entry) => (!worst || entry.net < worst.net ? entry : worst), null);
  const peakMonth = monthRows.reduce((peak, entry) => (!peak || entry.expense > peak.expense ? entry : peak), null);

  const insights = buildInsights({
    count, income, expense, net, savingsRate, deltas, current, previous,
    expenseCategories, incomeCategories, anomalies, monthRows, bestMonth, worstMonth,
    runwayMonths, burnRate, avgTicket, currency, digits, otherCurrencies,
  });

  return {
    bookName,
    currency,
    digits,
    generatedAt: new Date().toISOString(),
    otherCurrencies,
    period: {
      from: scoped[0]?.date ?? null,
      to: scoped.at(-1)?.date ?? null,
      months: monthRows.length,
    },
    totals: {
      income: round(income),
      expense: round(expense),
      net: round(net),
      count,
      activeDays,
      avgTicket,
      medianExpense: round(medianExpense),
      savingsRate,
      burnRate,
      runwayMonths,
      largestExpense: topExpenses[0]?.amount ?? 0,
    },
    months: monthRows,
    categories,
    expenseCategories,
    incomeCategories,
    topExpenses,
    weekdays,
    anomalies,
    deltas,
    bestMonth,
    worstMonth,
    peakMonth,
    currentMonth,
    insights,
  };
}

function fillMonthGaps(months) {
  const byKey = new Map(months.map((entry) => [entry.key, entry]));
  const first = months[0].key;
  const last = months.at(-1).key;
  const [startYear, startMonth] = first.split("-").map(Number);
  const [endYear, endMonth] = last.split("-").map(Number);
  const filled = [];
  let year = startYear;
  let month = startMonth;
  while (year < endYear || (year === endYear && month <= endMonth)) {
    const key = `${year}-${String(month).padStart(2, "0")}`;
    filled.push(byKey.get(key) ?? { key, income: 0, expense: 0, count: 0 });
    month += 1;
    if (month > 12) {
      month = 1;
      year += 1;
    }
  }
  return filled;
}

function buildInsights(context) {
  const {
    count, income, expense, net, savingsRate, deltas, current, previous,
    expenseCategories, anomalies, monthRows, bestMonth, worstMonth,
    runwayMonths, burnRate, avgTicket, currency, digits, otherCurrencies,
  } = context;
  const money = (value) => formatAmount(value, { currency, digits });
  const insights = [];

  if (count === 0) {
    insights.push({
      level: "info",
      title: "Belum ada transaksi",
      detail: "Buku ini masih kosong. Catat pemasukan atau pengeluaran pertama lewat chat, lalu minta laporan lagi.",
    });
    return insights;
  }

  insights.push({
    level: net >= 0 ? "good" : "danger",
    title: net >= 0 ? "Arus kas positif" : "Arus kas negatif",
    detail: net >= 0
      ? `Pemasukan ${money(income)} melampaui pengeluaran ${money(expense)}; selisih bersih ${money(net)}.`
      : `Pengeluaran ${money(expense)} lebih besar ${money(Math.abs(net))} dari pemasukan ${money(income)}. Perlu pemangkasan atau tambahan pemasukan.`,
  });

  if (savingsRate !== null) {
    const level = savingsRate >= 20 ? "good" : savingsRate >= 0 ? "warn" : "danger";
    const verdict = savingsRate >= 20
      ? "sehat, di atas ambang 20% yang umum dipakai perencana keuangan"
      : savingsRate >= 10
        ? "tipis; target aman biasanya 20%"
        : savingsRate >= 0
          ? "sangat tipis, hampir seluruh pemasukan habis"
          : "defisit: pengeluaran melebihi pemasukan";
    insights.push({
      level,
      title: `Rasio tabungan ${formatPercent(savingsRate)}`,
      detail: `Dari setiap pemasukan, bagian yang tersisa adalah ${formatPercent(savingsRate)} — ${verdict}.`,
    });
  }

  if (previous && current) {
    if (deltas.expense !== null) {
      const up = deltas.expense > 0;
      insights.push({
        level: Math.abs(deltas.expense) < 5 ? "info" : up ? "warn" : "good",
        title: `Pengeluaran ${up ? "naik" : "turun"} ${formatPercent(Math.abs(deltas.expense))} bulan ini`,
        detail: `${monthLabel(current.key)}: ${money(current.expense)} vs ${monthLabel(previous.key)}: ${money(previous.expense)}.`,
      });
    }
    if (deltas.income !== null && Math.abs(deltas.income) >= 5) {
      insights.push({
        level: deltas.income >= 0 ? "good" : "warn",
        title: `Pemasukan ${deltas.income >= 0 ? "naik" : "turun"} ${formatPercent(Math.abs(deltas.income))}`,
        detail: `${money(previous.income)} → ${money(current.income)} dibanding bulan sebelumnya.`,
      });
    }
  }

  const top = expenseCategories[0];
  if (top && top.share !== null) {
    insights.push({
      level: top.share >= 40 ? "warn" : "info",
      title: `Kategori terbesar: ${top.name}`,
      detail: `${money(top.total)} (${formatPercent(top.share)} dari total pengeluaran, ${top.count} transaksi, rata-rata ${money(top.average)}).${top.share >= 40 ? " Konsentrasi tinggi — satu kategori menyerap hampir separuh belanja." : ""}`,
    });
  }

  if (expenseCategories.length >= 3) {
    const topThree = expenseCategories.slice(0, 3);
    const share = topThree.reduce((total, entry) => total + (entry.share ?? 0), 0);
    insights.push({
      level: "info",
      title: "Tiga kategori mendominasi",
      detail: `${topThree.map((entry) => entry.name).join(", ")} menyumbang ${formatPercent(share)} pengeluaran. Fokus pemangkasan di sini memberi dampak tercepat.`,
    });
  }

  if (anomalies.length) {
    insights.push({
      level: "warn",
      title: `${anomalies.length} transaksi mencurigakan`,
      detail: `Nominal jauh di atas median pengeluaran (${money(context.medianExpense ?? 0)}). Periksa: ${anomalies.slice(0, 3).map((entry) => `${entry.description || entry.category} (${money(entry.amount)})`).join(", ")}.`,
    });
  }

  if (monthRows.length >= 2 && bestMonth && worstMonth) {
    insights.push({
      level: "info",
      title: `Bulan terbaik ${monthLabel(bestMonth.key)}`,
      detail: `Arus kas ${money(bestMonth.net)}. Bulan terlemah ${monthLabel(worstMonth.key)} dengan ${money(worstMonth.net)}.`,
    });
  }

  if (runwayMonths !== null && runwayMonths < 6) {
    insights.push({
      level: "warn",
      title: `Daya tahan kas ${runwayMonths} bulan`,
      detail: `Dengan pola pengeluaran sekarang, surplus setahun hanya menutup ${runwayMonths} bulan biaya. Target dana darurat umumnya 3–6 bulan.`,
    });
  }

  if (burnRate > 0) {
    insights.push({
      level: "info",
      title: `Rata-rata belanja ${money(burnRate)}/hari aktif`,
      detail: `${count} transaksi dalam ${context.count} catatan, rata-rata nominal ${money(avgTicket)} per transaksi.`,
    });
  }

  if (count < 10) {
    insights.push({
      level: "info",
      title: "Data masih tipis",
      detail: `Baru ${count} transaksi. Analisis tren dan perbandingan bulan akan jauh lebih tajam setelah 1–2 bulan pencatatan rutin.`,
    });
  }

  if (otherCurrencies.length) {
    insights.push({
      level: "info",
      title: "Ada mata uang lain",
      detail: `Laporan ini hanya menghitung ${currency}. Transaksi dalam ${otherCurrencies.join(", ")} tidak digabung agar kurs tidak dikarang.`,
    });
  }

  return insights;
}

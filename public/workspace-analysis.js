/* Analysis view for a workspace workbook.
 *
 * The spreadsheet stays the source of truth: this module only reads the sheets
 * and draws a dashboard. Nothing here mutates cell data.
 *
 * No chart library is used. Bars are CSS, the line and donut charts are inline
 * SVG, so the page keeps loading with zero external requests.
 */
(function (global) {
  'use strict';

  var MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  var MONTH_NAMES = {
    jan: 1, januari: 1, january: 1,
    feb: 2, februari: 2, february: 2,
    mar: 3, maret: 3, march: 3,
    apr: 4, april: 4,
    mei: 5, may: 5,
    jun: 6, juni: 6, june: 6,
    jul: 7, juli: 7, july: 7,
    agu: 8, agt: 8, agustus: 8, aug: 8, august: 8,
    sep: 9, sept: 9, september: 9,
    okt: 10, oktober: 10, oct: 10, october: 10,
    nov: 11, november: 11,
    des: 12, desember: 12, dec: 12, december: 12
  };

  var ROLE_WORDS = {
    income: ['pemasukan', 'pendapatan', 'income', 'revenue', 'penjualan', 'omset', 'masuk', 'debit', 'penerimaan', 'cashin'],
    expense: ['pengeluaran', 'belanja', 'biaya', 'keluar', 'expense', 'cost', 'spending', 'beban', 'kredit', 'pembayaran', 'cashout'],
    date: ['tanggal', 'tgl', 'date', 'waktu', 'periode', 'bulan', 'time'],
    category: ['kategori', 'category', 'jenis', 'tipe', 'type', 'akun', 'pos', 'klasifikasi', 'grup', 'group'],
    note: ['keterangan', 'deskripsi', 'catatan', 'note', 'memo', 'uraian', 'detail'],
    amount: ['jumlah', 'nominal', 'nilai', 'amount', 'total', 'value']
  };

  var PALETTE = ['#1F4E79', '#2E75B6', '#5B9BD5', '#A9C7E8', '#F2A93B', '#E06C5A', '#7FB069', '#B4A7D6'];

  /* ── values ──────────────────────────────────────────── */

  function rawValue(cell) {
    if (cell === undefined || cell === null) return '';
    if (typeof cell === 'object') return cell.value === undefined ? '' : cell.value;
    return cell;
  }

  function toText(value) {
    if (value === undefined || value === null) return '';
    return typeof value === 'string' ? value.trim() : String(value);
  }

  /* Accepts 1.500.000,50 as well as 1500000.50. */
  function toNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    if (typeof value === 'boolean') return value ? 1 : 0;
    var text = toText(value);
    if (!text) return 0;
    var cleaned = text.replace(/[^\d.,-]/g, '');
    if (!cleaned || cleaned === '-' || cleaned === '.' || cleaned === ',') return 0;
    var hasDot = cleaned.indexOf('.') >= 0;
    var hasComma = cleaned.indexOf(',') >= 0;
    if (hasDot && hasComma) {
      cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    } else if (hasComma) {
      cleaned = cleaned.replace(',', '.');
    } else if (hasDot) {
      var parts = cleaned.split('.');
      if (parts.length > 2 || parts[parts.length - 1].length === 3) cleaned = cleaned.replace(/\./g, '');
    }
    var n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }

  function toDate(value) {
    if (typeof value === 'number') {
      // Excel serial dates live in this range for realistic bookkeeping data.
      if (value > 20000 && value < 80000) {
        return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
      }
      return null;
    }
    var text = toText(value);
    if (!text) return null;
    var m = /^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?/.exec(text);
    if (m) return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3] || 1)));
    m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/.exec(text);
    if (m) {
      var year = Number(m[3]);
      if (year < 100) year += year < 70 ? 2000 : 1900;
      return new Date(Date.UTC(year, Number(m[2]) - 1, Number(m[1])));
    }
    m = /^(\d{1,2})[\s-]+([A-Za-z]+)[\s-]+(\d{2,4})/.exec(text);
    if (m && MONTH_NAMES[m[2].toLowerCase()]) {
      var y2 = Number(m[3]);
      if (y2 < 100) y2 += y2 < 70 ? 2000 : 1900;
      return new Date(Date.UTC(y2, MONTH_NAMES[m[2].toLowerCase()] - 1, Number(m[1])));
    }
    m = /^([A-Za-z]+)[\s-]+(\d{4})$/.exec(text);
    if (m && MONTH_NAMES[m[1].toLowerCase()]) {
      return new Date(Date.UTC(Number(m[2]), MONTH_NAMES[m[1].toLowerCase()] - 1, 1));
    }
    return null;
  }

  function monthKey(date) {
    return date ? date.getUTCFullYear() + '-' + String(date.getUTCMonth() + 1).padStart(2, '0') : '';
  }

  function monthLabel(key) {
    if (!key) return '';
    var parts = key.split('-');
    return MONTH_SHORT[Number(parts[1]) - 1] + ' ' + parts[0];
  }

  function columnLabel(index) {
    var name = '';
    var n = index;
    do {
      name = String.fromCharCode(65 + (n % 26)) + name;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return name;
  }

  function escapeHtml(text) {
    return String(text === undefined || text === null ? '' : text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /* ── column detection ────────────────────────────────── */

  function normalizeLabel(text) {
    return toText(text).toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  function roleOfLabel(label) {
    var key = normalizeLabel(label);
    if (!key) return null;
    for (var role in ROLE_WORDS) {
      if (!Object.prototype.hasOwnProperty.call(ROLE_WORDS, role)) continue;
      var words = ROLE_WORDS[role];
      for (var i = 0; i < words.length; i++) if (key === words[i]) return { role: role, score: 3 };
      for (var j = 0; j < words.length; j++) if (key.indexOf(words[j]) >= 0) return { role: role, score: 2 };
    }
    return null;
  }

  /* Score a candidate header row by how many known role words it contains. */
  function inspectRow(rows, headerRow, width) {
    var roles = {};
    var options = {};
    var score = 0;
    var row = rows[headerRow] || [];
    for (var c = 0; c < width; c++) {
      var found = roleOfLabel(rawValue(row[c]));
      if (!found) continue;
      score += found.score;
      if (roles[found.role] === undefined) roles[found.role] = c;
      if (!options[found.role]) options[found.role] = [];
      options[found.role].push(c);
    }
    return { headerRow: headerRow, roles: roles, options: options, score: score };
  }

  /* How many different values a column holds below the header (capped). A real
     category column has many; a "Jenis" column that only says Pemasukan or
     Pengeluaran has two. */
  function distinctCount(rows, col, start) {
    var seen = {};
    var count = 0;
    var limit = Math.min(rows.length, start + 200);
    for (var r = start; r < limit; r++) {
      var text = toText(rawValue((rows[r] || [])[col]));
      if (!text) continue;
      var key = text.toLowerCase();
      if (seen[key]) continue;
      seen[key] = 1;
      if (++count > 24) break;
    }
    return count;
  }

  var TYPE_VALUES = /^(pemasukan|pengeluaran|masuk|keluar|income|expense|debit|kredit|credit|in|out)$/i;

  /* A column that only repeats "Pemasukan"/"Pengeluaran" is a type flag, not a
     category. Keeping it out of `category` is what stops a dashboard from
     saying "biggest expense: Pengeluaran". */
  function looksLikeTypeColumn(rows, col, start) {
    var total = 0;
    var typed = 0;
    var limit = Math.min(rows.length, start + 200);
    for (var r = start; r < limit; r++) {
      var text = toText(rawValue((rows[r] || [])[col]));
      if (!text) continue;
      total++;
      if (TYPE_VALUES.test(text)) typed++;
    }
    return total > 0 && typed / total >= 0.8;
  }

  /* When two columns claim the same role, keep the one that actually varies.
     A book with ID | Tanggal | Jenis | Kategori | Keterangan must bind
     `category` to "Kategori", not to "Jenis". */
  function resolveRole(rows, headerRow, options, role) {
    var list = options ? options[role] : null;
    if (!list || !list.length) return undefined;
    if (list.length === 1) return list[0];
    var bestCol = list[0];
    var bestScore = -Infinity;
    for (var i = 0; i < list.length; i++) {
      var score = distinctCount(rows, list[i], headerRow + 1);
      if (role === 'category' && looksLikeTypeColumn(rows, list[i], headerRow + 1)) score -= 100;
      if (score > bestScore) { bestScore = score; bestCol = list[i]; }
    }
    return bestCol;
  }

  /* How many rows under the header actually hold a usable number in the bound
     money columns. A header that says "Pemasukan" over empty cells is not a
     ledger, so this stops a title sheet from being read as one. */
  function countNumericRows(sheet, fields) {
    var rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    var cols = [fields.income, fields.expense, fields.amount].filter(function (c) {
      return c !== null && c !== undefined;
    });
    if (!cols.length) return 0;
    var start = (fields.headerRow || 0) + 1;
    var hits = 0;
    for (var r = start; r < rows.length; r++) {
      var row = rows[r] || [];
      for (var i = 0; i < cols.length; i++) {
        var value = rawValue(row[cols[i]]);
        if (typeof value === 'number' && Number.isFinite(value) && value !== 0) { hits++; break; }
        if (typeof value === 'string' && toNumber(value) !== 0) { hits++; break; }
      }
    }
    return hits;
  }

  /* Find the sheet and header row that look most like a transaction table. */
  function detectFields(dataset) {
    var sheets = (dataset && Array.isArray(dataset.sheets)) ? dataset.sheets : [];
    var best = null;

    for (var s = 0; s < sheets.length; s++) {
      var rows = Array.isArray(sheets[s].rows) ? sheets[s].rows : [];
      if (!rows.length) continue;
      var width = Math.max(
        Array.isArray(sheets[s].columns) ? sheets[s].columns.length : 0,
        rows[0] ? rows[0].length : 0,
        1
      );
      var limit = Math.min(rows.length - 2, 25);
      for (var h = 0; h <= limit; h++) {
        var found = inspectRow(rows, h, width);
        if (!found.roles.income && !found.roles.expense && !found.roles.amount) continue;
        var dataRows = Math.min(rows.length - h - 1, 500);
        if (dataRows < 2) continue;
        var total = found.score + Math.min(4, dataRows / 20);
        if (!best || total > best.total) {
          best = { sheet: s, headerRow: h, roles: found.roles, options: found.options, total: total };
        }
      }
    }

    if (!best) return detectByNumbers(sheets);

    var bestRows = Array.isArray(sheets[best.sheet].rows) ? sheets[best.sheet].rows : [];
    var fields = {
      sheet: best.sheet,
      headerRow: best.headerRow,
      date: resolveRole(bestRows, best.headerRow, best.options, 'date'),
      income: resolveRole(bestRows, best.headerRow, best.options, 'income'),
      expense: resolveRole(bestRows, best.headerRow, best.options, 'expense'),
      category: resolveRole(bestRows, best.headerRow, best.options, 'category'),
      note: resolveRole(bestRows, best.headerRow, best.options, 'note'),
      amount: null
    };
    for (var key in fields) {
      if (key !== 'sheet' && key !== 'headerRow' && fields[key] === undefined) fields[key] = null;
    }
    // Only fall back to a single signed column when neither side was labelled.
    if (fields.income === null && fields.expense === null) {
      fields.amount = best.roles.amount === undefined ? null : best.roles.amount;
    }
    // A labelled header with no numbers beneath it is not a ledger.
    if (countNumericRows(sheets[best.sheet], fields) < 2) {
      return detectByNumbers(sheets);
    }
    return Object.assign(fields, { confidence: 'header' });
  }

  /* Detect the header row and column roles for one specific sheet. Used when
     the user switches sheet in the Sumber Data panel. */
  function detectForSheet(dataset, sheetIndex) {
    var sheets = (dataset && Array.isArray(dataset.sheets)) ? dataset.sheets : [];
    var sheet = sheets[sheetIndex];
    if (!sheet) return null;
    var rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (!rows.length) return null;
    var width = Math.max(
      Array.isArray(sheet.columns) ? sheet.columns.length : 0,
      rows[0] ? rows[0].length : 0,
      1
    );
    var best = null;
    var limit = Math.min(rows.length - 2, 25);
    for (var h = 0; h <= limit; h++) {
      var found = inspectRow(rows, h, width);
      if (!found.roles.income && !found.roles.expense && !found.roles.amount) continue;
      if (!best || found.score > best.score) best = found;
    }
    var roles = best ? best.roles : {};
    var options = best ? best.options : {};
    var headerRow = best ? best.headerRow : 0;
    var fields = {
      sheet: sheetIndex,
      headerRow: headerRow,
      date: resolveRole(rows, headerRow, options, 'date'),
      income: resolveRole(rows, headerRow, options, 'income'),
      expense: resolveRole(rows, headerRow, options, 'expense'),
      category: resolveRole(rows, headerRow, options, 'category'),
      note: resolveRole(rows, headerRow, options, 'note'),
      amount: null
    };
    for (var key in fields) {
      if (key !== 'sheet' && key !== 'headerRow' && fields[key] === undefined) fields[key] = null;
    }
    if (fields.income === null && fields.expense === null) {
      fields.amount = roles.amount === undefined ? null : roles.amount;
    }
    return fields;
  }

  /* No recognisable headers: guess from the shape of the numbers. */
  function detectByNumbers(sheets) {
    for (var s = 0; s < sheets.length; s++) {
      var rows = Array.isArray(sheets[s].rows) ? sheets[s].rows : [];
      if (rows.length < 2) continue;
      var width = rows[0] ? rows[0].length : 0;
      var numeric = [];
      for (var c = 0; c < width; c++) {
        var count = 0;
        var negative = 0;
        for (var r = 0; r < Math.min(rows.length, 40); r++) {
          var value = rawValue(rows[r] ? rows[r][c] : undefined);
          if (typeof value === 'number') {
            count++;
            if (value < 0) negative++;
          }
        }
        if (count >= Math.min(3, rows.length - 1)) numeric.push({ col: c, negative: negative });
      }
      if (numeric.length >= 2) {
        return { sheet: s, headerRow: 0, date: null, income: numeric[0].col, expense: numeric[1].col, category: null, note: null, amount: null, confidence: 'guess' };
      }
      if (numeric.length === 1) {
        var signed = numeric[0].negative > 0;
        return {
          sheet: s, headerRow: 0, date: null,
          income: signed ? null : numeric[0].col,
          expense: null, category: null, note: null,
          amount: signed ? numeric[0].col : null,
          confidence: 'guess'
        };
      }
    }
    return null;
  }

  /* ── ledger ──────────────────────────────────────────── */

  function buildLedger(dataset, fields) {
    var empty = { records: [], sheetName: '', rowCount: 0 };
    if (!fields || !dataset || !Array.isArray(dataset.sheets)) return empty;
    var sheet = dataset.sheets[fields.sheet];
    if (!sheet) return empty;
    var rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    var records = [];
    var start = Math.min(rows.length, (fields.headerRow || 0) + 1);

    for (var r = start; r < rows.length; r++) {
      var row = rows[r] || [];
      var income = fields.income === null ? 0 : toNumber(rawValue(row[fields.income]));
      var expense = fields.expense === null ? 0 : toNumber(rawValue(row[fields.expense]));
      if (fields.amount !== null) {
        var signed = toNumber(rawValue(row[fields.amount]));
        if (signed < 0) expense += -signed;
        else income += signed;
      }
      var category = fields.category === null ? '' : toText(rawValue(row[fields.category]));
      var note = fields.note === null ? '' : toText(rawValue(row[fields.note]));
      var date = fields.date === null ? null : toDate(rawValue(row[fields.date]));
      if (!income && !expense && !category && !note) continue;
      records.push({
        row: r,
        income: income,
        expense: expense,
        net: income - expense,
        category: category || 'Lainnya',
        note: note,
        date: date,
        month: monthKey(date)
      });
    }
    return { records: records, sheetName: sheet.name, rowCount: rows.length };
  }

  /* ── aggregation ─────────────────────────────────────── */

  function autoWidgets() {
    return [
      { id: 'w1', type: 'kpi', title: 'Pemasukan', metric: 'income', agg: 'sum', groupBy: 'none', format: 'currency', limit: 8 },
      { id: 'w2', type: 'kpi', title: 'Pengeluaran', metric: 'expense', agg: 'sum', groupBy: 'none', format: 'currency', limit: 8 },
      { id: 'w3', type: 'kpi', title: 'Untung Bersih', metric: 'net', agg: 'sum', groupBy: 'none', format: 'currency', limit: 8 },
      { id: 'w4', type: 'line', title: 'Arus Kas per Bulan', metric: 'net', agg: 'sum', groupBy: 'month', format: 'currency', limit: 24 },
      { id: 'w5', type: 'bar', title: 'Pengeluaran per Kategori', metric: 'expense', agg: 'sum', groupBy: 'category', format: 'currency', limit: 8 },
      { id: 'w6', type: 'donut', title: 'Komposisi Pengeluaran', metric: 'expense', agg: 'sum', groupBy: 'category', format: 'currency', limit: 6 },
      { id: 'w7', type: 'table', title: 'Transaksi Terakhir', metric: 'net', agg: 'sum', groupBy: 'none', format: 'currency', limit: 8, recent: true }
    ];
  }

  function metricValue(record, metric) {
    if (metric === 'income') return record.income;
    if (metric === 'expense') return record.expense;
    if (metric === 'count') return 1;
    return record.net;
  }

  function aggregate(values, agg) {
    if (!values.length) return 0;
    if (agg === 'count') return values.length;
    var sum = 0;
    for (var i = 0; i < values.length; i++) sum += values[i];
    if (agg === 'avg') return sum / values.length;
    if (agg === 'min') return Math.min.apply(null, values);
    if (agg === 'max') return Math.max.apply(null, values);
    return sum;
  }

  function groupValues(records, metric, keyOf, agg) {
    var buckets = {};
    var order = [];
    for (var i = 0; i < records.length; i++) {
      var key = keyOf(records[i]);
      if (!key) continue;
      if (!buckets[key]) { buckets[key] = []; order.push(key); }
      buckets[key].push(metricValue(records[i], metric));
    }
    return order.map(function (key) { return { label: key, value: aggregate(buckets[key], agg) }; });
  }

  function computeWidget(widget, ledger) {
    var records = ledger.records;
    if (!records.length) return { empty: true };

    if (widget.type === 'kpi') {
      var total = aggregate(records.map(function (record) { return metricValue(record, widget.metric); }), widget.agg);
      var sub = '';
      if (widget.metric === 'net') {
        var income = aggregate(records.map(function (r) { return r.income; }), 'sum');
        sub = income ? 'Margin ' + ((total / income) * 100).toFixed(1).replace('.', ',') + '%' : records.length + ' transaksi';
      } else {
        sub = records.length.toLocaleString('id-ID') + ' transaksi';
      }
      return { value: total, sub: sub };
    }

    if (widget.type === 'bar' || widget.type === 'donut') {
      var grouped = groupValues(records, widget.metric, function (record) {
        return widget.groupBy === 'month' ? record.month : record.category;
      }, widget.agg);
      grouped.sort(function (a, b) { return Math.abs(b.value) - Math.abs(a.value); });
      var items = grouped.slice(0, widget.limit);
      var max = 0;
      for (var i = 0; i < items.length; i++) max = Math.max(max, Math.abs(items[i].value));
      return { items: items, max: max || 1 };
    }

    if (widget.type === 'line') {
      var byMonth = groupValues(records, widget.metric, function (record) { return record.month; }, widget.agg);
      byMonth.sort(function (a, b) { return a.label < b.label ? -1 : a.label > b.label ? 1 : 0; });
      var points = byMonth.slice(-24).map(function (item) {
        return { label: monthLabel(item.label), value: item.value };
      });
      if (widget.metric === 'balance') {
        var running = 0;
        points = points.map(function (point) { running += point.value; return { label: point.label, value: running }; });
      }
      var min = 0;
      var high = 0;
      for (var p = 0; p < points.length; p++) {
        min = Math.min(min, points[p].value);
        high = Math.max(high, points[p].value);
      }
      return { points: points, min: min, max: high };
    }

    if (widget.type === 'table') {
      var sorted;
      if (widget.recent) {
        // Newest first; rows without a date keep their sheet order reversed,
        // so an undated ledger still shows the latest entries it has.
        var dated = records.some(function (record) { return record.date; });
        sorted = records.slice().sort(function (a, b) {
          if (!dated) return 0;
          var ta = a.date ? a.date.getTime() : 0;
          var tb = b.date ? b.date.getTime() : 0;
          if (tb !== ta) return tb - ta;
          return b.row - a.row;
        });
        if (!dated) sorted.reverse();
      } else {
        sorted = records.slice().sort(function (a, b) { return Math.abs(b.net) - Math.abs(a.net); });
      }
      return {
        rows: sorted.slice(0, widget.limit).map(function (record) {
          return { date: record.date, category: record.category, note: record.note, value: record.net };
        })
      };
    }

    if (widget.type === 'progress') {
      var done = aggregate(records.map(function (record) { return metricValue(record, widget.metric); }), widget.agg);
      var target = widget.target || 1;
      return { value: done, target: target, ratio: target ? done / target : 0 };
    }

    return { empty: true };
  }

  /* ── formatting ──────────────────────────────────────── */

  function formatNumber(value, format) {
    if (!Number.isFinite(value)) return '—';
    if (format === 'percent') return (value * 100).toFixed(1).replace('.', ',') + '%';
    var rounded = Math.round(value * 100) / 100;
    var text = new Intl.NumberFormat('id-ID', { maximumFractionDigits: Math.abs(rounded) < 100 ? 2 : 0 }).format(rounded);
    return format === 'currency' ? 'Rp ' + text : text;
  }

  function formatCompact(value) {
    if (!Number.isFinite(value)) return '—';
    var abs = Math.abs(value);
    var sign = value < 0 ? '-' : '';
    if (abs >= 1e12) return sign + 'Rp ' + (abs / 1e12).toFixed(1).replace('.', ',') + ' T';
    if (abs >= 1e9) return sign + 'Rp ' + (abs / 1e9).toFixed(1).replace('.', ',') + ' M';
    if (abs >= 1e6) return sign + 'Rp ' + (abs / 1e6).toFixed(1).replace('.', ',') + ' jt';
    if (abs >= 1e3) return sign + 'Rp ' + Math.round(abs / 1e3) + ' rb';
    return formatNumber(value, 'currency');
  }

  /* ── widget rendering ────────────────────────────────── */

  function renderKpi(widget, data) {
    if (data.empty) return '<div class="aw-empty">Belum ada data</div>';
    var tone = widget.metric === 'expense' ? 'out' : widget.metric === 'net' ? (data.value < 0 ? 'neg' : 'pos') : 'in';
    return '<div class="aw-kpi ' + tone + '">' + escapeHtml(formatNumber(data.value, widget.format)) + '</div>' +
      (data.sub ? '<div class="aw-sub">' + escapeHtml(data.sub) + '</div>' : '');
  }

  function renderBar(widget, data) {
    if (data.empty) return '<div class="aw-empty">Belum ada data</div>';
    if (!data.items.length) return '<div class="aw-empty">Tidak ada nilai</div>';
    var html = '<div class="aw-bars">';
    for (var i = 0; i < data.items.length; i++) {
      var item = data.items[i];
      var pct = Math.max(2, (Math.abs(item.value) / data.max) * 100);
      html += '<div class="aw-bar-row">' +
        '<span class="aw-bar-label" title="' + escapeHtml(item.label) + '">' + escapeHtml(item.label) + '</span>' +
        '<span class="aw-bar-track"><span class="aw-bar-fill' + (item.value < 0 ? ' neg' : '') + '" style="width:' + pct.toFixed(1) + '%"></span></span>' +
        '<span class="aw-bar-value">' + escapeHtml(formatCompact(item.value)) + '</span></div>';
    }
    return html + '</div>';
  }

  function renderDonut(widget, data) {
    if (data.empty) return '<div class="aw-empty">Belum ada data</div>';
    if (!data.items.length) return '<div class="aw-empty">Tidak ada nilai</div>';
    var total = data.items.reduce(function (n, item) { return n + Math.abs(item.value); }, 0) || 1;
    var radius = 42;
    var circumference = 2 * Math.PI * radius;
    var offset = 0;
    var svg = '<svg viewBox="0 0 120 120" class="aw-donut" role="img" aria-label="' + escapeHtml(widget.title) + '">' +
      '<circle cx="60" cy="60" r="' + radius + '" fill="none" stroke="#e8ecf1" stroke-width="18"/>';
    for (var i = 0; i < data.items.length; i++) {
      var dash = (Math.abs(data.items[i].value) / total) * circumference;
      svg += '<circle cx="60" cy="60" r="' + radius + '" fill="none" stroke="' + PALETTE[i % PALETTE.length] +
        '" stroke-width="18" stroke-dasharray="' + dash.toFixed(2) + ' ' + (circumference - dash).toFixed(2) +
        '" stroke-dashoffset="' + (-offset).toFixed(2) + '" transform="rotate(-90 60 60)"><title>' +
        escapeHtml(data.items[i].label) + '</title></circle>';
      offset += dash;
    }
    svg += '</svg>';
    var legend = '<div class="aw-legend">';
    for (var j = 0; j < data.items.length; j++) {
      var share = (Math.abs(data.items[j].value) / total) * 100;
      legend += '<div class="aw-legend-row"><span class="aw-dot" style="background:' + PALETTE[j % PALETTE.length] + '"></span>' +
        '<span class="aw-legend-name" title="' + escapeHtml(data.items[j].label) + '">' + escapeHtml(data.items[j].label) + '</span>' +
        '<span class="aw-legend-value">' + share.toFixed(0) + '%</span></div>';
    }
    return '<div class="aw-donut-wrap">' + svg + legend + '</div></div>';
  }

  function renderLine(widget, data) {
    if (data.empty) return '<div class="aw-empty">Belum ada data</div>';
    if (!data.points.length) return '<div class="aw-empty">Tidak ada nilai</div>';
    var width = 320;
    var height = 120;
    var padX = 10;
    var padY = 14;
    var span = (data.max - data.min) || 1;
    var step = data.points.length > 1 ? (width - padX * 2) / (data.points.length - 1) : 0;
    var coords = data.points.map(function (point, index) {
      return {
        x: padX + index * step,
        y: padY + (height - padY * 2) * (1 - (point.value - data.min) / span),
        point: point
      };
    });
    var line = coords.map(function (item) { return item.x.toFixed(1) + ',' + item.y.toFixed(1); }).join(' ');
    var baseline = (height - padY).toFixed(1);
    var lastX = (padX + (data.points.length - 1) * step).toFixed(1);
    var area = padX + ',' + baseline + ' ' + line + ' ' + lastX + ',' + baseline;
    var svg = '<svg viewBox="0 0 ' + width + ' ' + height + '" class="aw-line" preserveAspectRatio="none" role="img" aria-label="' + escapeHtml(widget.title) + '">' +
      '<polygon points="' + area + '" fill="rgba(46,117,182,0.12)"/>' +
      '<polyline points="' + line + '" fill="none" stroke="#2E75B6" stroke-width="2" stroke-linejoin="round"/>';
    for (var i = 0; i < coords.length; i++) {
      svg += '<circle cx="' + coords[i].x.toFixed(1) + '" cy="' + coords[i].y.toFixed(1) + '" r="2.6" fill="#1F4E79"><title>' +
        escapeHtml(coords[i].point.label + ': ' + formatCompact(coords[i].point.value)) + '</title></circle>';
    }
    svg += '</svg>';
    return svg + '<div class="aw-axis"><span>' + escapeHtml(data.points[0].label) + '</span><span>' +
      escapeHtml(data.points[data.points.length - 1].label) + '</span></div>';
  }

  function renderTable(widget, data) {
    if (data.empty) return '<div class="aw-empty">Belum ada data</div>';
    if (!data.rows.length) return '<div class="aw-empty">Tidak ada nilai</div>';
    var html = '<table class="aw-table"><tbody>';
    for (var i = 0; i < data.rows.length; i++) {
      var row = data.rows[i];
      var label = row.note || row.category || 'Transaksi';
      var when = row.date ? row.date.toISOString().slice(0, 10) : '';
      html += '<tr><td class="aw-td-main" title="' + escapeHtml(label) + '">' + escapeHtml(label) +
        (when ? '<span class="aw-td-date">' + escapeHtml(when) + '</span>' : '') + '</td>' +
        '<td class="aw-td-value ' + (row.value < 0 ? 'neg' : 'pos') + '">' + escapeHtml(formatCompact(row.value)) + '</td></tr>';
    }
    return html + '</tbody></table>';
  }

  function renderProgress(widget, data) {
    if (data.empty) return '<div class="aw-empty">Belum ada data</div>';
    var ratio = Math.max(0, Math.min(1, data.ratio));
    return '<div class="aw-progress"><span class="aw-progress-fill" style="width:' + (ratio * 100).toFixed(1) + '%"></span></div>' +
      '<div class="aw-sub">' + escapeHtml(formatCompact(data.value)) + ' dari ' + escapeHtml(formatCompact(data.target)) +
      ' · ' + (data.ratio * 100).toFixed(1).replace('.', ',') + '%</div>';
  }

  function renderWidgetBody(widget, data) {
    switch (widget.type) {
      case 'kpi': return renderKpi(widget, data);
      case 'bar': return renderBar(widget, data);
      case 'donut': return renderDonut(widget, data);
      case 'line': return renderLine(widget, data);
      case 'table': return renderTable(widget, data);
      case 'progress': return renderProgress(widget, data);
      default: return '';
    }
  }

  function renderWidgets(container, widgets, ledger) {
    if (!widgets.length) {
      container.innerHTML = '<div class="aw-blank">Belum ada widget. Tulis permintaan di kotak atas lalu tekan <b>Buat widget</b>, atau tekan <b>Standar</b> untuk susunan siap pakai.</div>';
      return;
    }
    var html = '';
    for (var i = 0; i < widgets.length; i++) {
      var widget = widgets[i];
      var wide = widget.type !== 'kpi' && widget.type !== 'progress';
      html += '<section class="aw-card' + (wide ? ' wide' : '') + '" data-widget="' + escapeHtml(widget.id) + '">' +
        '<header class="aw-head"><h3 class="aw-title">' + escapeHtml(widget.title) + '</h3>' +
        '<button type="button" class="aw-del" data-del="' + escapeHtml(widget.id) + '" title="Hapus widget" aria-label="Hapus widget">' +
        '<svg viewBox="0 0 16 16" width="12" height="12" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" aria-hidden="true">' +
        '<path d="M4 4l8 8M12 4l-8 8"/></svg></button></header>' +
        '<div class="aw-body">' + renderWidgetBody(widget, computeWidget(widget, ledger)) + '</div></section>';
    }
    container.innerHTML = html;
  }

  /* ── source panel ────────────────────────────────────── */

  function columnOptions(sheet, selected) {
    var width = Math.max(Array.isArray(sheet.columns) ? sheet.columns.length : 0, 1);
    var html = '<option value="">Tidak dipakai</option>';
    for (var c = 0; c < width; c++) {
      var header = '';
      var rows = Array.isArray(sheet.rows) ? sheet.rows : [];
      for (var r = 0; r < Math.min(rows.length, 12); r++) {
        var text = toText(rawValue(rows[r] ? rows[r][c] : undefined));
        if (text && !/^-?[\d.,]+$/.test(text)) { header = text; break; }
      }
      html += '<option value="' + c + '"' + (selected === c ? ' selected' : '') + '>' +
        escapeHtml(columnLabel(c) + (header ? ': ' + header.slice(0, 24) : '')) + '</option>';
    }
    return html;
  }

  function fieldSelect(key, label, optionsHtml, extraAttr) {
    return '<label class="aw-field"><span>' + escapeHtml(label) + '</span>' +
      '<select data-field="' + key + '"' + (extraAttr || '') + '>' + optionsHtml + '</select></label>';
  }

  function renderSourcePanel(container, dataset, fields) {
    var sheets = (dataset && Array.isArray(dataset.sheets)) ? dataset.sheets : [];
    var sheetHtml = '';
    for (var s = 0; s < sheets.length; s++) {
      sheetHtml += '<option value="' + s + '"' + (fields && fields.sheet === s ? ' selected' : '') + '>' +
        escapeHtml(sheets[s].name) + '</option>';
    }
    var sheet = fields ? sheets[fields.sheet] : null;
    var rowCount = sheet && Array.isArray(sheet.rows) ? sheet.rows.length : 0;

    container.innerHTML =
      '<div class="aw-panel-head"><h2>Sumber Data</h2>' +
      '<span class="aw-panel-note">Tentukan kolom mana yang berarti apa. Semua angka dihitung dari kolom ini.</span></div>' +
      '<div class="aw-fields">' +
        fieldSelect('sheet', 'Sheet', sheetHtml, ' data-sheet="1"') +
        fieldSelect('date', 'Tanggal', sheet ? columnOptions(sheet, fields.date) : '') +
        fieldSelect('income', 'Pemasukan', sheet ? columnOptions(sheet, fields.income) : '') +
        fieldSelect('expense', 'Pengeluaran', sheet ? columnOptions(sheet, fields.expense) : '') +
        fieldSelect('amount', 'Nominal (+/-)', sheet ? columnOptions(sheet, fields.amount) : '') +
        fieldSelect('category', 'Kategori', sheet ? columnOptions(sheet, fields.category) : '') +
        fieldSelect('note', 'Keterangan', sheet ? columnOptions(sheet, fields.note) : '') +
      '</div>' +
      '<div class="aw-panel-foot">' + (rowCount ? rowCount.toLocaleString('id-ID') + ' baris di sheet ini' : 'Sheet kosong') + '</div>';
  }

  /* Keyword-only widget choice, mirrored from the server so the static demo
     behaves like the real endpoint when no AI is reachable. */
  function keywordWidgets(prompt) {
    var text = String(prompt || '').toLowerCase();
    var found = [];
    function has(pattern) { return pattern.test(text); }
    var wantsIncome = has(/pemasukan|pendapatan|income|masuk|omset|penjualan|revenue/);
    var wantsExpense = has(/pengeluaran|belanja|biaya|keluar|expense|cost|spending/);
    var wantsNet = has(/untung|laba|profit|margin|net|bersih|selisih/);
    var wantsCategory = has(/kategori|category|jenis|pos|akun|per bagian/);
    var wantsMonth = has(/bulan|tren|trend|waktu|periode|monthly/);
    var wantsShare = has(/komposisi|proporsi|porsi|persentase|share|donat|pie/);
    var wantsTop = has(/terbesar|tertinggi|top|terbanyak|ranking/);
    var wantsTarget = has(/target|progres|progress|capai/);

    function push(type, title, metric, groupBy, extra) {
      if (found.length >= 5) return;
      var widget = { id: 'k' + (found.length + 1), type: type, title: title, metric: metric, agg: 'sum', groupBy: groupBy || 'none', format: 'currency', limit: 8 };
      if (extra) for (var key in extra) widget[key] = extra[key];
      found.push(widget);
    }

    if (wantsNet) push('kpi', 'Untung Bersih', 'net');
    if (wantsIncome) push('kpi', 'Pemasukan', 'income');
    if (wantsExpense) push('kpi', 'Pengeluaran', 'expense');
    if (wantsShare) push('donut', 'Komposisi Pengeluaran', wantsExpense ? 'expense' : 'income', 'category');
    else if (wantsCategory) push('bar', 'Per Kategori', wantsExpense ? 'expense' : 'income', 'category');
    if (wantsMonth) push('line', 'Tren per Bulan', wantsNet ? 'net' : 'income', 'month');
    if (wantsTop) push('table', 'Transaksi Terbesar', 'net');
    if (wantsTarget) push('progress', 'Target', 'income', 'none', { target: 1000000, format: 'percent' });
    if (!found.length) {
      push('kpi', 'Untung Bersih', 'net');
      push('bar', 'Pengeluaran per Kategori', 'expense', 'category');
    }
    return found;
  }

  /* ── status bar ──────────────────────────────────────── */

  /* One thin row that answers "which book, which period, how fresh". It is
     the first thing on screen, so it must never be empty. */
  function renderStatus(container, dataset, ledger) {
    if (!container) return;
    var title = toText(dataset && dataset.title) || 'Workspace';
    var records = ledger.records || [];

    var period = '';
    var months = records.filter(function (record) { return record.month; })
      .map(function (record) { return record.month; }).sort();
    if (months.length) {
      var first = months[0];
      var last = months[months.length - 1];
      period = first === last ? monthLabel(first) : monthLabel(first) + ' – ' + monthLabel(last);
    }

    var updated = dataset && dataset.calculated_at ? new Date(dataset.calculated_at) : null;
    var updatedText = updated && !isNaN(updated.getTime())
      ? 'Diperbarui ' + updated.toLocaleString('id-ID', {
        day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
      })
      : '';

    var chips = [];
    if (period) chips.push('<span class="aw-chip">' + escapeHtml(period) + '</span>');
    if (records.length) {
      chips.push('<span class="aw-chip">' + records.length.toLocaleString('id-ID') + ' transaksi</span>');
    }
    chips.push('<span class="aw-chip ghost">' + escapeHtml(ledger.sheetName ? 'Sheet: ' + ledger.sheetName : 'Sheet kosong') + '</span>');

    container.innerHTML =
      '<div class="aw-status-main">' +
        '<h2 class="aw-status-title">' + escapeHtml(title) + '</h2>' +
        '<div class="aw-status-chips">' + chips.join('') + '</div>' +
      '</div>' +
      '<div class="aw-status-side">' +
        (updatedText ? '<span class="aw-status-updated">' + escapeHtml(updatedText) + '</span>' : '') +
        '<button type="button" class="aw-status-btn" data-mode="sheet">Buka Sheet</button>' +
      '</div>';
  }

  /* ── notes ───────────────────────────────────────────── */

  function medianOf(values) {
    if (!values.length) return 0;
    var sorted = values.slice().sort(function (a, b) { return a - b; });
    var mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }

  /* Three plain-language observations, computed locally. Mirrors the server
     rule in src/finance/analytics.js: an expense above 3x the median counts as
     unusual. No AI call, so the first screen costs nothing. */
  function buildInsights(ledger) {
    var records = (ledger && ledger.records) || [];
    if (!records.length) return [];
    var notes = [];

    var expenses = records.filter(function (record) { return record.expense > 0; });
    var byCategory = {};
    for (var i = 0; i < expenses.length; i++) {
      var key = expenses[i].category || 'Lainnya';
      byCategory[key] = (byCategory[key] || 0) + expenses[i].expense;
    }
    var categories = Object.keys(byCategory).map(function (name) {
      return { label: name, value: byCategory[name] };
    }).sort(function (a, b) { return b.value - a.value; });
    var totalExpense = categories.reduce(function (sum, item) { return sum + item.value; }, 0);
    if (categories.length && totalExpense > 0) {
      notes.push({
        tone: 'info',
        title: 'Pengeluaran terbesar: ' + categories[0].label,
        detail: formatCompact(categories[0].value) + ' · ' +
          ((categories[0].value / totalExpense) * 100).toFixed(0) + '% dari total pengeluaran.'
      });
    }

    var byMonth = {};
    for (var m = 0; m < records.length; m++) {
      if (!records[m].month) continue;
      byMonth[records[m].month] = (byMonth[records[m].month] || 0) + records[m].net;
    }
    var monthKeys = Object.keys(byMonth).sort();
    if (monthKeys.length >= 2) {
      var best = monthKeys[0];
      var worst = monthKeys[0];
      for (var k = 0; k < monthKeys.length; k++) {
        if (byMonth[monthKeys[k]] > byMonth[best]) best = monthKeys[k];
        if (byMonth[monthKeys[k]] < byMonth[worst]) worst = monthKeys[k];
      }
      notes.push({
        tone: byMonth[best] < 0 ? 'warn' : 'good',
        title: 'Bulan terbaik: ' + monthLabel(best),
        detail: 'Untung bersih ' + formatCompact(byMonth[best]) +
          ' · terendah ' + monthLabel(worst) + ' ' + formatCompact(byMonth[worst]) + '.'
      });
    }

    var median = medianOf(expenses.map(function (record) { return record.expense; }));
    if (median > 0) {
      var unusual = expenses.filter(function (record) { return record.expense > median * 3; })
        .sort(function (a, b) { return b.expense - a.expense; });
      if (unusual.length) {
        notes.push({
          tone: 'warn',
          title: unusual.length + ' pengeluaran jauh di atas wajar',
          detail: 'Di atas 3× median (' + formatCompact(median) + '): ' +
            unusual.slice(0, 2).map(function (record) {
              return (record.note || record.category) + ' ' + formatCompact(record.expense);
            }).join(', ') + '.'
        });
      }
    }
    return notes.slice(0, 3);
  }

  function renderInsights(container, ledger) {
    if (!container) return;
    var notes = buildInsights(ledger);
    if (!notes.length) {
      container.innerHTML = '';
      container.hidden = true;
      return;
    }
    container.hidden = false;
    var html = '<h2 class="aw-insight-title">Catatan</h2><ul class="aw-insight-list">';
    for (var i = 0; i < notes.length; i++) {
      html += '<li class="aw-insight-item ' + escapeHtml(notes[i].tone) + '">' +
        '<strong>' + escapeHtml(notes[i].title) + '</strong>' +
        '<span>' + escapeHtml(notes[i].detail) + '</span></li>';
    }
    container.innerHTML = html + '</ul>';
  }

  /* ── public entry point ──────────────────────────────── */

  function render(root, dataset, analysis) {
    var fields = analysis && analysis.fields ? analysis.fields : detectFields(dataset);
    var widgets = analysis && Array.isArray(analysis.widgets) && analysis.widgets.length
      ? analysis.widgets
      : autoWidgets();
    var ledger = buildLedger(dataset, fields);

    var source = root.querySelector('#aw-source');
    var grid = root.querySelector('#aw-grid');
    var status = root.querySelector('#aw-status');
    var insight = root.querySelector('#aw-insight');
    if (source) renderSourcePanel(source, dataset, fields);
    if (grid) renderWidgets(grid, widgets, ledger);
    if (status) renderStatus(status, dataset, ledger);
    if (insight) renderInsights(insight, ledger);

    var meta = root.querySelector('#aw-meta');
    if (meta) {
      if (!fields) {
        meta.textContent = 'Kolom pemasukan/pengeluaran belum terdeteksi. Buka "Atur analisis" untuk memilih kolom.';
        meta.className = 'aw-meta warn';
      } else if (!ledger.records.length) {
        meta.textContent = 'Sheet terpilih belum punya baris data.';
        meta.className = 'aw-meta warn';
      } else if (fields.confidence === 'guess') {
        // No labelled header: say so, because the numbers may be meaningless.
        meta.textContent = ledger.records.length.toLocaleString('id-ID') +
          ' baris ditebak dari ' + ledger.sheetName + ' — kolom belum berlabel, periksa "Atur analisis".';
        meta.className = 'aw-meta warn';
      } else {
        meta.textContent = ledger.records.length.toLocaleString('id-ID') + ' transaksi terbaca dari ' + ledger.sheetName;
        meta.className = 'aw-meta';
      }
    }
    return { fields: fields, widgets: widgets, ledger: ledger };
  }

  global.WorkspaceAnalysis = {
    render: render,
    renderWidgets: renderWidgets,
    detectFields: detectFields,
    detectForSheet: detectForSheet,
    autoWidgets: autoWidgets,
    buildLedger: buildLedger,
    computeWidget: computeWidget,
    columnLabel: columnLabel,
    formatNumber: formatNumber,
    formatCompact: formatCompact,
    toNumber: toNumber,
    toDate: toDate,
    keywordWidgets: keywordWidgets,
    renderStatus: renderStatus,
    renderInsights: renderInsights,
    buildInsights: buildInsights
  };
})(window);

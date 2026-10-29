/* Excel-like workspace client. Vanilla JS, no dependencies. */
(function () {
  'use strict';

  var boot = window.__WORKSPACE__ || {};
  var BASE = boot.base || '';
  var api = function (path) { return BASE + path; };

  var MIN_ROWS = 120;
  var MIN_COLS = 26;
  var EXTRA_ROWS = 30;
  var EXTRA_COLS = 8;
  var DEFAULT_COL_W = 96;
  var MIN_COL_W = 54;
  var MAX_COL_W = 340;
  var ROW_SCAN_LIMIT = 400;
  var SAVE_DELAY = 650;

  var workbook = null;
  var activeSheet = 0;
  var sel = { r: 0, c: 0 };
  var anchor = { r: 0, c: 0 };
  var editing = false;
  var formulaDirty = false;
  var saving = false;
  var saveTimer = null;
  var undoStack = [];
  var colWidths = {};
  var autoWidths = {};
  var rowHeadW = 46;
  var cellTds = [];
  var rowTds = [];
  var colTds = [];

  var gridWrap = document.getElementById('grid-wrap');
  var table = document.getElementById('sheet');
  var colHead = document.getElementById('col-head');
  var body = document.getElementById('body');
  var tabs = document.getElementById('tabs');
  var nameBox = document.getElementById('name-box');
  var formulaInput = document.getElementById('formula-input');
  var statusEl = document.getElementById('status');
  var noticeEl = document.getElementById('notice');
  var menuEl = document.getElementById('menu');

  /* ── helpers ─────────────────────────────────────────── */

  function colName(index) {
    var name = '';
    var n = index + 1;
    while (n > 0) {
      var rem = (n - 1) % 26;
      name = String.fromCharCode(65 + rem) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  }

  function colIndex(letters) {
    var total = 0;
    var text = String(letters).toUpperCase();
    for (var i = 0; i < text.length; i++) {
      var code = text.charCodeAt(i);
      if (code < 65 || code > 90) return -1;
      total = total * 26 + (code - 64);
    }
    return total - 1;
  }

  function cellRef(r, c) { return colName(c) + (r + 1); }

  function parseRef(text) {
    var match = /^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})\s*$/.exec(String(text || ''));
    if (!match) return null;
    var c = colIndex(match[1]);
    var r = parseInt(match[2], 10) - 1;
    if (c < 0 || r < 0) return null;
    return { r: r, c: c };
  }

  function sheet() { return workbook.sheets[activeSheet]; }

  function getCell(r, c) {
    var rows = sheet().rows;
    if (!rows[r]) return undefined;
    return rows[r][c];
  }

  function ensureCell(r, c) {
    var s = sheet();
    while (s.rows.length <= r) s.rows.push([]);
    var row = s.rows[r];
    while (row.length <= c) row.push('');
    return row;
  }

  function isFormula(cell) {
    return Boolean(cell && typeof cell === 'object' && typeof cell.formula === 'string');
  }

  function cellValue(cell) {
    if (cell === null || cell === undefined) return '';
    if (typeof cell === 'object') {
      if (isFormula(cell)) return cell.value === undefined || cell.value === null ? '' : cell.value;
      return '';
    }
    return cell;
  }

  function display(cell) {
    var value = cellValue(cell);
    if (value === '') return '';
    if (typeof value === 'number') {
      if (!isFinite(value)) return '';
      return String(value);
    }
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    return String(value);
  }

  function rawText(cell) {
    if (isFormula(cell)) return '=' + cell.formula;
    var value = cellValue(cell);
    if (value === '') return '';
    return String(value);
  }

  function isError(cell) {
    var value = cellValue(cell);
    return typeof value === 'string' && value.charAt(0) === '#';
  }

  /* A formula whose result has not come back from the server yet. */
  function isPending(cell) {
    return isFormula(cell) && (cell.value === undefined || cell.value === null);
  }

  function isNumericCell(cell) {
    return typeof cellValue(cell) === 'number';
  }

  /* ── cell styles ─────────────────────────────────────── */

  var FONT_STACK = '"Segoe UI", system-ui, -apple-system, sans-serif';

  function styleOf(r, c) {
    var styles = sheet().styles;
    if (!styles) return null;
    return styles[r + ',' + c] || null;
  }

  function setStyleAt(r, c, style) {
    var s = sheet();
    if (!s.styles) s.styles = {};
    var key = r + ',' + c;
    if (!style || !Object.keys(style).length) delete s.styles[key];
    else s.styles[key] = style;
    if (!Object.keys(s.styles).length) delete s.styles;
  }

  /* Build the inline CSS for one cell. Only properties the style actually
     carries are emitted, so an unstyled cell keeps the grid defaults. */
  var BORDER_CSS = {
    thin: '1px solid',
    medium: '2px solid',
    thick: '3px solid',
    hair: '1px solid',
    dotted: '1px dotted',
    dashed: '1px dashed',
    double: '3px double',
    mediumDashed: '2px dashed',
    dashDot: '1px dashed',
    mediumDashDot: '2px dashed',
    dashDotDot: '1px dotted',
    mediumDashDotDot: '2px dotted'
  };

  function styleCss(style) {
    if (!style) return '';
    var css = '';
    if (style.fill) css += 'background:' + style.fill + ';';
    if (style.color) css += 'color:' + style.color + ';';
    if (style.bold) css += 'font-weight:700;';
    if (style.italic) css += 'font-style:italic;';
    if (style.underline && style.strike) css += 'text-decoration:underline line-through;';
    else if (style.underline) css += 'text-decoration:underline;';
    else if (style.strike) css += 'text-decoration:line-through;';
    if (style.font) css += 'font-family:"' + String(style.font).replace(/"/g, '') + '",' + FONT_STACK + ';';
    if (style.size) css += 'font-size:' + style.size + 'px;';
    if (style.align) css += 'text-align:' + style.align + ';';
    if (style.valign) css += 'vertical-align:' + style.valign + ';';
    if (style.wrap) css += 'white-space:normal;';
    if (style.border) {
      for (var side in style.border) {
        if (!Object.prototype.hasOwnProperty.call(style.border, side)) continue;
        var edge = style.border[side];
        var rule = BORDER_CSS[edge.style] || BORDER_CSS.thin;
        css += 'border-' + side + ':' + rule + ' ' + (edge.color || '#000000') + ';';
      }
    }
    return css;
  }

  function styleAttr(style) {
    var css = styleCss(style);
    return css ? ' style="' + escapeHtml(css) + '"' : '';
  }

  /* Apply a patch to every cell in the selection. A null value removes the
     property, which is how "no fill" and "reset" work. */
  function applyStyleToSelection(patch) {
    var box = rangeOf();
    var changed = 0;
    snapshot();
    for (var r = box.r1; r <= box.r2; r++) {
      for (var c = box.c1; c <= box.c2; c++) {
        var current = styleOf(r, c);
        var next = {};
        if (current) {
          for (var key in current) {
            if (Object.prototype.hasOwnProperty.call(current, key)) next[key] = current[key];
          }
        }
        for (var prop in patch) {
          if (!Object.prototype.hasOwnProperty.call(patch, prop)) continue;
          if (patch[prop] === null || patch[prop] === undefined) delete next[prop];
          else next[prop] = patch[prop];
        }
        setStyleAt(r, c, Object.keys(next).length ? next : null);
        changed++;
      }
    }
    render();
    updateToolbarState();
    scheduleSave();
    return changed;
  }

  /* Toolbar controls mirror the active cell's formatting. */
  function updateToolbarState() {
    var style = styleOf(sel.r, sel.c) || {};
    var map = {
      'cmd-bold': !!style.bold,
      'cmd-italic': !!style.italic,
      'cmd-underline': !!style.underline,
      'cmd-strike': !!style.strike
    };
    for (var id in map) {
      var el = document.getElementById(id);
      if (el) el.setAttribute('aria-pressed', map[id] ? 'true' : 'false');
    }
    var fillSwatch = document.getElementById('fill-swatch');
    if (fillSwatch) fillSwatch.style.background = style.fill || 'transparent';
    var colorSwatch = document.getElementById('color-swatch');
    if (colorSwatch) colorSwatch.style.background = style.color || 'currentColor';
    var family = document.getElementById('font-family');
    if (family && style.font) family.value = style.font;
    else if (family) family.value = '';
    var size = document.getElementById('font-size');
    if (size) size.value = style.size ? String(style.size) : '';
    var alignButtons = document.querySelectorAll('[data-align]');
    for (var i = 0; i < alignButtons.length; i++) {
      alignButtons[i].setAttribute('aria-pressed', alignButtons[i].dataset.align === style.align ? 'true' : 'false');
    }
  }

  function escapeHtml(text) {
    return String(text === null || text === undefined ? '' : text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /* ── text measurement ────────────────────────────────── */

  var measureCtx = null;

  function textWidth(text, bold) {
    if (!measureCtx) {
      var canvas = document.createElement('canvas');
      measureCtx = canvas.getContext('2d');
    }
    measureCtx.font = (bold ? '600 ' : '') + '12.5px "Segoe UI", system-ui, -apple-system, sans-serif';
    return measureCtx.measureText(String(text)).width;
  }

  /* Widest value in a column, so the column can size itself to its content.
     Values too wide to ever fit are skipped: those are titles or notes, and
     letting them drive the width would flatten every column to the maximum. */
  function autoColWidth(c) {
    var s = sheet();
    var widest = textWidth(colName(c), true) + 26;
    var limit = Math.min(s.rows.length, ROW_SCAN_LIMIT);
    for (var r = 0; r < limit; r++) {
      var cell = s.rows[r] ? s.rows[r][c] : undefined;
      if (cell === undefined || cell === null || cell === '') continue;
      var text = display(cell);
      if (!text) continue;
      var w = textWidth(text, false) + 18;
      if (w > MAX_COL_W) continue;
      if (w > widest) widest = w;
      if (widest >= MAX_COL_W) break;
    }
    return Math.max(MIN_COL_W, Math.min(MAX_COL_W, Math.ceil(widest)));
  }

  /* Excel merges were flattened on import, so a merged title now repeats the
     same long text in every covered cell. Group those runs back into one
     spanning cell for display. The stored data is never modified. */
  function rowSpans(row, cols) {
    var spans = {};
    var skip = {};
    var c = 0;
    while (c < cols) {
      var cell = row ? row[c] : undefined;
      var text = display(cell);
      if (text && !isFormula(cell) && textWidth(text, false) > MAX_COL_W * 0.75) {
        var run = 1;
        while (c + run < cols && display(row[c + run]) === text) run++;
        if (run >= 2) {
          spans[c] = run;
          for (var k = 1; k < run; k++) skip[c + k] = true;
          c += run;
          continue;
        }
      }
      c++;
    }
    return { spans: spans, skip: skip };
  }

  function rowHeadWidth(rows) {
    return Math.max(38, Math.ceil(textWidth(String(rows), false)) + 22);
  }

  function computeAutoWidths(cols) {
    autoWidths = {};
    for (var c = 0; c < cols; c++) autoWidths[c] = autoColWidth(c);
  }

  /* ── notices ─────────────────────────────────────────── */

  var noticeTimer = null;
  function notify(message, tone) {
    noticeEl.textContent = message;
    noticeEl.dataset.tone = tone || 'success';
    noticeEl.classList.add('show');
    if (noticeTimer) clearTimeout(noticeTimer);
    noticeTimer = setTimeout(function () { noticeEl.classList.remove('show'); }, 2600);
  }

  /* ── api ─────────────────────────────────────────────── */

  function call(path, options) {
    return fetch(api(path), Object.assign({ cache: 'no-store' }, options || {}))
      .then(function (response) {
        return response.json().catch(function () { return {}; }).then(function (result) {
          if (!response.ok) throw new Error(result.error || 'Permintaan gagal.');
          return result;
        });
      });
  }

  /* ── rendering ───────────────────────────────────────── */

  function visibleRows() {
    return Math.max(sheet().rows.length + EXTRA_ROWS, MIN_ROWS);
  }

  function visibleCols() {
    var width = sheet().columns.length;
    for (var i = 0; i < sheet().rows.length; i++) {
      if (sheet().rows[i].length > width) width = sheet().rows[i].length;
    }
    return Math.max(width + EXTRA_COLS, MIN_COLS);
  }

  /* A manually resized column wins; otherwise the auto-fit width applies. */
  function colWidth(c) {
    if (colWidths[c]) return colWidths[c];
    if (autoWidths[c]) return autoWidths[c];
    return DEFAULT_COL_W;
  }

  function renderTabs() {
    var html = '';
    for (var i = 0; i < workbook.sheets.length; i++) {
      html += '<button type="button" data-sheet="' + i + '" aria-selected="' + (i === activeSheet) + '">' +
        escapeHtml(workbook.sheets[i].name) + '</button>';
    }
    tabs.innerHTML = html;
  }

  function render() {
    var rows = visibleRows();
    var cols = visibleCols();

    // Size every column to its content before laying the table out.
    computeAutoWidths(cols);
    rowHeadW = rowHeadWidth(rows);

    // Fixed layout needs an explicit table width, otherwise the browser sizes
    // columns from content and ignores the per-column widths.
    var totalWidth = rowHeadW;
    for (var w = 0; w < cols; w++) totalWidth += colWidth(w);
    table.style.width = totalWidth + 'px';
    table.style.setProperty('--rowhead-w', rowHeadW + 'px');

    var colgroup = '<colgroup><col style="width:' + rowHeadW + 'px">';
    for (var c = 0; c < cols; c++) colgroup += '<col style="width:' + colWidth(c) + 'px">';
    colgroup += '</colgroup>';

    var head = '<tr><th class="corner"></th>';
    for (var c2 = 0; c2 < cols; c2++) {
      head += '<th data-col="' + c2 + '">' + colName(c2) +
        '<span class="col-grip" data-grip="' + c2 + '"></span></th>';
    }
    head += '</tr>';

    var rowsHtml = '';
    for (var r = 0; r < rows; r++) {
      rowsHtml += '<tr><th data-row="' + r + '">' + (r + 1) + '</th>';
      var layout = rowSpans(sheet().rows[r], cols);
      for (var c3 = 0; c3 < cols; c3++) {
        if (layout.skip[c3]) continue;
        rowsHtml += cellHtml(r, c3, layout.spans[c3] || 1);
      }
      rowsHtml += '</tr>';
    }

    table.innerHTML = colgroup + '<thead>' + head + '</thead><tbody>' + rowsHtml + '</tbody>';

    colHead = table.querySelector('thead');
    body = table.querySelector('tbody');

    cacheElements(rows, cols);
    updateSelection();
    updateFormulaBar();
    renderTabs();
  }

  function cellHtml(r, c, colspan) {
    var cell = getCell(r, c);
    var style = styleOf(r, c);
    var classes = 'cell';
    if (isNumericCell(cell)) classes += ' num';
    if (isError(cell)) classes += ' err';
    if (isPending(cell)) classes += ' pending';
    if (colspan > 1) classes += ' span';
    if (style && style.align) classes += ' align-' + style.align;
    var text = isPending(cell) ? '…' : display(cell);
    var spanAttr = colspan > 1 ? ' colspan="' + colspan + '"' : '';
    var titleAttr = colspan > 1 ? ' title="' + escapeHtml(text) + '"' : '';
    return '<td class="' + classes + '" data-r="' + r + '" data-c="' + c + '"' + spanAttr + titleAttr +
      styleAttr(style) + '>' + escapeHtml(text) + '</td>';
  }

  function cacheElements(rows, cols) {
    cellTds = [];
    rowTds = [];
    colTds = [];
    var trs = body.querySelectorAll('tr');
    for (var r = 0; r < trs.length; r++) {
      var tr = trs[r];
      rowTds[r] = tr.querySelector('th');
      cellTds[r] = [];
      var tds = tr.querySelectorAll('td');
      // Index by data-c: spanning cells make the DOM order skip columns.
      for (var c = 0; c < tds.length; c++) {
        cellTds[r][Number(tds[c].dataset.c)] = tds[c];
      }
    }
    var ths = colHead.querySelectorAll('th');
    for (var i = 1; i < ths.length; i++) colTds[i - 1] = ths[i];
  }

  function rangeOf() {
    return {
      r1: Math.min(anchor.r, sel.r), r2: Math.max(anchor.r, sel.r),
      c1: Math.min(anchor.c, sel.c), c2: Math.max(anchor.c, sel.c)
    };
  }

  function updateSelection() {
    var previous = table.querySelectorAll('.active, .in-range, .hl');
    for (var i = 0; i < previous.length; i++) {
      previous[i].classList.remove('active', 'in-range', 'hl');
    }
    var box = rangeOf();
    var multi = box.r1 !== box.r2 || box.c1 !== box.c2;

    for (var r = box.r1; r <= box.r2; r++) {
      if (cellTds[r]) {
        for (var c = box.c1; c <= box.c2; c++) {
          var td = cellTds[r][c];
          if (td && multi) td.classList.add('in-range');
        }
      }
    }
    if (cellTds[sel.r] && cellTds[sel.r][sel.c]) cellTds[sel.r][sel.c].classList.add('active');

    for (var r2 = box.r1; r2 <= box.r2; r2++) {
      if (rowTds[r2]) rowTds[r2].classList.add('hl');
    }
    for (var c2 = box.c1; c2 <= box.c2; c2++) {
      if (colTds[c2]) colTds[c2].classList.add('hl');
    }

    nameBox.textContent = multi
      ? cellRef(box.r1, box.c1) + ':' + cellRef(box.r2, box.c2)
      : cellRef(sel.r, sel.c);

    updateStatus(box, multi);
    updateToolbarState();
  }

  function updateStatus(box, multi) {
    if (!multi) { statusEl.textContent = ''; return; }
    var sum = 0, count = 0, numeric = 0;
    for (var r = box.r1; r <= box.r2; r++) {
      for (var c = box.c1; c <= box.c2; c++) {
        var cell = getCell(r, c);
        var value = cellValue(cell);
        if (value === '' || value === null || value === undefined) continue;
        count++;
        if (typeof value === 'number') { sum += value; numeric++; }
      }
    }
    var parts = ['Terisi: ' + count];
    if (numeric) {
      parts.push('Jumlah: ' + round(sum));
      parts.push('Rata-rata: ' + round(sum / numeric));
    }
    statusEl.textContent = parts.join('  |  ');
  }

  function round(value) {
    if (!isFinite(value)) return '0';
    return String(Math.round(value * 10000) / 10000);
  }

  function updateFormulaBar() {
    if (editing) return;
    formulaInput.textContent = rawText(getCell(sel.r, sel.c));
    formulaDirty = false;
  }

  /* Apply whatever is typed in the formula bar to the selected cell. */
  function commitFormulaBar() {
    if (!formulaDirty) return;
    formulaDirty = false;
    applyValue(sel.r, sel.c, formulaInput.textContent);
    refreshCell(sel.r, sel.c);
    updateFormulaBar();
  }

  function scrollIntoView(r, c) {
    var td = cellTds[r] && cellTds[r][c];
    if (!td) return;
    var wrapRect = gridWrap.getBoundingClientRect();
    var tdRect = td.getBoundingClientRect();
    if (tdRect.top < wrapRect.top + 26) gridWrap.scrollTop -= (wrapRect.top + 26 - tdRect.top);
    else if (tdRect.bottom > wrapRect.bottom) gridWrap.scrollTop += (tdRect.bottom - wrapRect.bottom);
    if (tdRect.left < wrapRect.left + rowHeadW) gridWrap.scrollLeft -= (wrapRect.left + rowHeadW - tdRect.left);
    else if (tdRect.right > wrapRect.right) gridWrap.scrollLeft += (tdRect.right - wrapRect.right);
  }

  /* ── selection ───────────────────────────────────────── */

  function select(r, c, extend) {
    if (editing) commitEdit(null);
    else commitFormulaBar();
    sel.r = Math.max(0, r);
    sel.c = Math.max(0, c);
    if (!extend) { anchor.r = sel.r; anchor.c = sel.c; }
    updateSelection();
    updateFormulaBar();
    scrollIntoView(sel.r, sel.c);
  }

  function move(dr, dc, extend) {
    select(sel.r + dr, sel.c + dc, extend);
  }

  /* ── editing ─────────────────────────────────────────── */

  function startEdit(seed) {
    var td = cellTds[sel.r] && cellTds[sel.r][sel.c];
    if (!td) return;
    editing = true;
    td.setAttribute('contenteditable', 'true');
    td.textContent = seed !== undefined ? seed : rawText(getCell(sel.r, sel.c));
    td.classList.add('active');
    td.focus();
    if (seed !== undefined) {
      var range = document.createRange();
      range.selectNodeContents(td);
      range.collapse(false);
      var selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    } else {
      var range2 = document.createRange();
      range2.selectNodeContents(td);
      var selection2 = window.getSelection();
      selection2.removeAllRanges();
      selection2.addRange(range2);
    }
  }

  function commitEdit(direction) {
    if (!editing) return;
    var td = cellTds[sel.r] && cellTds[sel.r][sel.c];
    editing = false;
    var text = td ? td.textContent : '';
    if (td) {
      td.setAttribute('contenteditable', 'false');
      td.classList.remove('active');
    }
    applyValue(sel.r, sel.c, text);
    refreshCell(sel.r, sel.c);
    refitColumn(sel.c);
    updateFormulaBar();

    if (direction === 'down') move(1, 0);
    else if (direction === 'up') move(-1, 0);
    else if (direction === 'right') move(0, 1);
    else if (direction === 'left') move(0, -1);
    else updateSelection();
  }

  function cancelEdit() {
    if (!editing) return;
    var td = cellTds[sel.r] && cellTds[sel.r][sel.c];
    editing = false;
    if (td) {
      td.setAttribute('contenteditable', 'false');
      td.classList.remove('active');
    }
    refreshCell(sel.r, sel.c);
    updateFormulaBar();
    updateSelection();
  }

  function refreshCell(r, c) {
    var td = cellTds[r] && cellTds[r][c];
    if (!td) return;
    var cell = getCell(r, c);
    var style = styleOf(r, c);
    td.textContent = isPending(cell) ? '…' : display(cell);
    td.classList.toggle('num', isNumericCell(cell));
    td.classList.toggle('err', isError(cell));
    td.classList.toggle('pending', isPending(cell));
    td.classList.toggle('align-left', !!style && style.align === 'left');
    td.classList.toggle('align-center', !!style && style.align === 'center');
    td.classList.toggle('align-right', !!style && style.align === 'right');
    var css = styleCss(style);
    if (css) td.setAttribute('style', css);
    else td.removeAttribute('style');
  }

  /* Re-size one column after an edit, without rebuilding the whole grid so the
     scroll position and selection survive. */
  function refitColumn(c) {
    if (colWidths[c]) return; // a manual width is intentional
    autoWidths[c] = autoColWidth(c);
    var cols = table.querySelectorAll('colgroup col');
    var target = cols[c + 1];
    if (!target) return;
    target.style.width = autoWidths[c] + 'px';
    var total = rowHeadW;
    for (var i = 0; i < cols.length - 1; i++) total += colWidth(i);
    table.style.width = total + 'px';
  }

  function applyValue(r, c, text) {
    snapshot();
    var row = ensureCell(r, c);
    var trimmed = String(text === null || text === undefined ? '' : text);
    if (trimmed.charAt(0) === '=') {
      row[c] = { formula: trimmed.slice(1) };
    } else if (trimmed === '') {
      row[c] = '';
    } else if (/^-?\d+(\.\d+)?$/.test(trimmed.trim())) {
      row[c] = Number(trimmed.trim());
    } else {
      row[c] = trimmed;
    }
    scheduleSave();
  }

  /* ── undo ────────────────────────────────────────────── */

  function snapshot() {
    undoStack.push(JSON.stringify(workbook.sheets));
    if (undoStack.length > 60) undoStack.shift();
  }

  function undo() {
    if (!undoStack.length) { notify('Tidak ada yang bisa dibatalkan.', 'busy'); return; }
    var previous = JSON.parse(undoStack.pop());
    workbook.sheets = previous;
    if (activeSheet >= workbook.sheets.length) activeSheet = workbook.sheets.length - 1;
    render();
    scheduleSave();
    notify('Dibatalkan.', 'busy');
  }

  /* ── saving ──────────────────────────────────────────── */

  function scheduleSave() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(function () { saveTimer = null; save(true); }, SAVE_DELAY);
  }

  function save(silent) {
    if (saving) return Promise.resolve();
    saving = true;
    var payload = { title: workbook.title, sheets: workbook.sheets };
    if (workbook.analysis) payload.analysis = workbook.analysis;
    return call('/workbook', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(function (result) {
      if (result && Array.isArray(result.sheets)) {
        workbook.sheets = result.sheets;
        if (result.analysis) workbook.analysis = result.analysis;
        render();
        if (!analysisEl.hidden) renderAnalysis();
      }
      if (!silent) notify('Tersimpan.', 'success');
    }).catch(function (error) {
      notify(error.message, 'error');
    }).then(function () {
      saving = false;
    });
  }

  /* ── structure changes ───────────────────────────────── */

  /* Styles live in a "row,col" map, so structural edits must shift the keys
     or every colour would end up on the wrong cell. */
  function shiftStyles(mode, at, delta) {
    var s = sheet();
    if (!s.styles) return;
    var next = {};
    for (var key in s.styles) {
      if (!Object.prototype.hasOwnProperty.call(s.styles, key)) continue;
      var parts = key.split(',');
      var r = Number(parts[0]);
      var c = Number(parts[1]);
      if (mode === 'row') {
        if (r === at && delta < 0) continue;
        if (r >= at) r += delta;
      } else {
        if (c === at && delta < 0) continue;
        if (c >= at) c += delta;
      }
      if (r < 0 || c < 0) continue;
      next[r + ',' + c] = s.styles[key];
    }
    if (Object.keys(next).length) s.styles = next;
    else delete s.styles;
  }

  function insertRow(at) {
    snapshot();
    var s = sheet();
    var width = visibleCols();
    s.rows.splice(at, 0, new Array(width).fill(''));
    shiftStyles('row', at, 1);
    render();
    select(at, sel.c);
    scheduleSave();
  }

  function deleteRow(at) {
    snapshot();
    var s = sheet();
    s.rows.splice(at, 1);
    if (!s.rows.length) s.rows.push([]);
    shiftStyles('row', at, -1);
    render();
    select(Math.min(at, s.rows.length - 1), sel.c);
    scheduleSave();
  }

  function insertColumn(at) {
    snapshot();
    var s = sheet();
    s.columns.splice(at, 0, colName(at));
    for (var r = 0; r < s.rows.length; r++) s.rows[r].splice(at, 0, '');
    shiftStyles('col', at, 1);
    render();
    select(sel.r, at);
    scheduleSave();
  }

  function deleteColumn(at) {
    snapshot();
    var s = sheet();
    if (s.columns.length <= 1) { notify('Minimal satu kolom.', 'error'); return; }
    s.columns.splice(at, 1);
    for (var r = 0; r < s.rows.length; r++) s.rows[r].splice(at, 1);
    shiftStyles('col', at, -1);
    render();
    select(sel.r, Math.min(at, s.columns.length - 1));
    scheduleSave();
  }

  function clearRange() {
    var box = rangeOf();
    snapshot();
    for (var r = box.r1; r <= box.r2; r++) {
      for (var c = box.c1; c <= box.c2; c++) {
        var row = ensureCell(r, c);
        row[c] = '';
        // Clearing content clears the formatting too, like Excel's Delete.
        setStyleAt(r, c, null);
      }
    }
    render();
    scheduleSave();
  }

  function addSheet() {
    var name = window.prompt('Nama sheet', 'Sheet' + (workbook.sheets.length + 1));
    if (!name) return;
    snapshot();
    var width = Math.max(MIN_COLS, 8);
    var columns = [];
    for (var i = 0; i < width; i++) columns.push(colName(i));
    workbook.sheets.push({ name: String(name).slice(0, 60), columns: columns, rows: [] });
    activeSheet = workbook.sheets.length - 1;
    sel = { r: 0, c: 0 };
    anchor = { r: 0, c: 0 };
    render();
    scheduleSave();
  }

  function renameSheet() {
    var s = sheet();
    var name = window.prompt('Nama sheet', s.name);
    if (!name) return;
    snapshot();
    s.name = String(name).slice(0, 60);
    render();
    scheduleSave();
  }

  function deleteSheet() {
    if (workbook.sheets.length <= 1) { notify('Minimal satu sheet.', 'error'); return; }
    if (!window.confirm('Hapus sheet "' + sheet().name + '"?')) return;
    snapshot();
    workbook.sheets.splice(activeSheet, 1);
    activeSheet = Math.max(0, activeSheet - 1);
    sel = { r: 0, c: 0 };
    anchor = { r: 0, c: 0 };
    render();
    scheduleSave();
  }

  function addRows(count) {
    var s = sheet();
    var width = visibleCols();
    snapshot();
    for (var i = 0; i < count; i++) s.rows.push(new Array(width).fill(''));
    render();
    scheduleSave();
  }

  function addColumns(count) {
    var s = sheet();
    snapshot();
    for (var i = 0; i < count; i++) s.columns.push(colName(s.columns.length));
    render();
    scheduleSave();
  }

  /* ── section ──────────────────────────────────────── */

  function copyRange() {
    var box = rangeOf();
    var lines = [];
    for (var r = box.r1; r <= box.r2; r++) {
      var cells = [];
      for (var c = box.c1; c <= box.c2; c++) {
        cells.push(display(getCell(r, c)));
      }
      lines.push(cells.join('\t'));
    }
    var text = lines.join('\n');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () {
        notify('Disalin: ' + (box.r2 - box.r1 + 1) + '×' + (box.c2 - box.c1 + 1) + ' sel.', 'busy');
      }).catch(function () { notify('Gagal menyalin.', 'error'); });
    }
    return text;
  }

  function pasteText(text) {
    if (!text) return;
    var rows = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    if (rows.length && rows[rows.length - 1] === '') rows.pop();
    snapshot();
    for (var r = 0; r < rows.length; r++) {
      var cells = rows[r].split('\t');
      for (var c = 0; c < cells.length; c++) {
        var value = cells[c];
        var row = ensureCell(sel.r + r, sel.c + c);
        if (value.charAt(0) === '=') row[sel.c + c] = { formula: value.slice(1) };
        else if (value === '') row[sel.c + c] = '';
        else if (/^-?\d+(\.\d+)?$/.test(value.trim())) row[sel.c + c] = Number(value.trim());
        else row[sel.c + c] = value;
      }
    }
    render();
    scheduleSave();
    notify('Ditempel: ' + rows.length + ' baris.', 'busy');
  }

  /* ── context menu ────────────────────────────────────── */

  function openMenu(x, y) {
    menuEl.style.left = Math.min(x, window.innerWidth - 200) + 'px';
    menuEl.style.top = Math.min(y, window.innerHeight - 260) + 'px';
    menuEl.classList.add('open');
  }

  function closeMenu() { menuEl.classList.remove('open'); }

  /* ── events ──────────────────────────────────────────── */

  gridWrap.addEventListener('mousedown', function (event) {
    var grip = event.target.closest('.col-grip');
    if (grip) {
      event.preventDefault();
      var c = Number(grip.dataset.grip);
      var startX = event.clientX;
      var startW = colWidth(c);
      function onMove(moveEvent) {
        var width = Math.max(40, startW + (moveEvent.clientX - startX));
        colWidths[c] = width;
        var col = table.querySelectorAll('colgroup col')[c + 1];
        if (col) col.style.width = width + 'px';
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
      return;
    }

    var td = event.target.closest('td[data-r]');
    if (td) {
      if (editing) commitEdit(null);
      var r = Number(td.dataset.r);
      var c2 = Number(td.dataset.c);
      select(r, c2, event.shiftKey);
      return;
    }
    var rowHead = event.target.closest('th[data-row]');
    if (rowHead) {
      var rr = Number(rowHead.dataset.row);
      anchor = { r: rr, c: 0 };
      sel = { r: rr, c: visibleCols() - 1 };
      updateSelection();
      return;
    }
    var colHeader = event.target.closest('th[data-col]');
    if (colHeader) {
      var cc = Number(colHeader.dataset.col);
      anchor = { r: 0, c: cc };
      sel = { r: visibleRows() - 1, c: cc };
      updateSelection();
      return;
    }
    if (event.target.closest('th.corner')) {
      anchor = { r: 0, c: 0 };
      sel = { r: visibleRows() - 1, c: visibleCols() - 1 };
      updateSelection();
    }
  });

  gridWrap.addEventListener('dblclick', function (event) {
    // Double-clicking the column grip restores the auto-fit width, like Excel.
    var grip = event.target.closest('.col-grip');
    if (grip) {
      event.preventDefault();
      var gc = Number(grip.dataset.grip);
      delete colWidths[gc];
      autoWidths[gc] = autoColWidth(gc);
      var cols = table.querySelectorAll('colgroup col');
      if (cols[gc + 1]) cols[gc + 1].style.width = autoWidths[gc] + 'px';
      var total = rowHeadW;
      for (var i = 0; i < cols.length - 1; i++) total += colWidth(i);
      table.style.width = total + 'px';
      return;
    }

    var td = event.target.closest('td[data-r]');
    if (!td) return;
    startEdit();
  });

  gridWrap.addEventListener('contextmenu', function (event) {
    var td = event.target.closest('td[data-r]');
    if (!td) return;
    event.preventDefault();
    if (editing) commitEdit(null);
    select(Number(td.dataset.r), Number(td.dataset.c));
    openMenu(event.clientX, event.clientY);
  });

  document.addEventListener('click', function (event) {
    if (!menuEl.contains(event.target)) closeMenu();
  });

  menuEl.addEventListener('click', function (event) {
    var button = event.target.closest('button[data-action]');
    if (!button) return;
    closeMenu();
    var action = button.dataset.action;
    if (action === 'row-above') insertRow(sel.r);
    else if (action === 'row-below') insertRow(sel.r + 1);
    else if (action === 'row-delete') deleteRow(sel.r);
    else if (action === 'col-left') insertColumn(sel.c);
    else if (action === 'col-right') insertColumn(sel.c + 1);
    else if (action === 'col-delete') deleteColumn(sel.c);
    else if (action === 'clear') clearRange();
    else if (action === 'rename-sheet') renameSheet();
    else if (action === 'delete-sheet') deleteSheet();
  });

  tabs.addEventListener('click', function (event) {
    var button = event.target.closest('button[data-sheet]');
    if (!button) return;
    if (editing) commitEdit(null);
    activeSheet = Number(button.dataset.sheet);
    sel = { r: 0, c: 0 };
    anchor = { r: 0, c: 0 };
    render();
  });

  tabs.addEventListener('dblclick', function (event) {
    if (event.target.closest('button[data-sheet]')) renameSheet();
  });

  document.addEventListener('keydown', function (event) {
    var target = event.target;
    var inNameBox = target === nameBox;
    var inFormula = target === formulaInput;

    if (inNameBox) {
      if (event.key === 'Enter') {
        event.preventDefault();
        var ref = parseRef(nameBox.textContent);
        if (ref) select(ref.r, ref.c);
        else updateSelection();
        gridWrap.focus();
      }
      return;
    }

    if (inFormula) {
      if (event.key === 'Enter') {
        event.preventDefault();
        formulaDirty = true;
        commitFormulaBar();
        move(1, 0);
        gridWrap.focus();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        updateFormulaBar();
        gridWrap.focus();
      } else if (event.key === 'Tab') {
        event.preventDefault();
        formulaDirty = true;
        commitFormulaBar();
        move(0, event.shiftKey ? -1 : 1);
        gridWrap.focus();
      }
      return;
    }

    var ctrl = event.ctrlKey || event.metaKey;

    if (ctrl && event.key.toLowerCase() === 'z') { event.preventDefault(); undo(); return; }
    if (ctrl && event.key.toLowerCase() === 's') { event.preventDefault(); save(false); return; }
    if (ctrl && event.key.toLowerCase() === 'c') { event.preventDefault(); copyRange(); return; }
    if (ctrl && event.key.toLowerCase() === 'b') { event.preventDefault(); toggleFlag('bold'); return; }
    if (ctrl && event.key.toLowerCase() === 'i') { event.preventDefault(); toggleFlag('italic'); return; }
    if (ctrl && event.key.toLowerCase() === 'u') { event.preventDefault(); toggleFlag('underline'); return; }
    if (ctrl && event.key === 'Home') { event.preventDefault(); select(0, 0); return; }

    if (editing) {
      if (event.key === 'Enter') { event.preventDefault(); commitEdit('down'); }
      else if (event.key === 'Tab') { event.preventDefault(); commitEdit('right'); }
      else if (event.key === 'Escape') { event.preventDefault(); cancelEdit(); }
      return;
    }

    if (event.key === 'ArrowUp') { event.preventDefault(); move(-1, 0, event.shiftKey); }
    else if (event.key === 'ArrowDown') { event.preventDefault(); move(1, 0, event.shiftKey); }
    else if (event.key === 'ArrowLeft') { event.preventDefault(); move(0, -1, event.shiftKey); }
    else if (event.key === 'ArrowRight') { event.preventDefault(); move(0, 1, event.shiftKey); }
    else if (event.key === 'Tab') { event.preventDefault(); move(0, event.shiftKey ? -1 : 1); }
    else if (event.key === 'Home') { event.preventDefault(); select(sel.r, 0); }
    else if (event.key === 'End') { event.preventDefault(); select(sel.r, visibleCols() - 1); }
    else if (event.key === 'PageDown') { event.preventDefault(); move(20, 0, event.shiftKey); }
    else if (event.key === 'PageUp') { event.preventDefault(); move(-20, 0, event.shiftKey); }
    else if (event.key === 'Enter') { event.preventDefault(); startEdit(); }
    else if (event.key === 'F2') { event.preventDefault(); startEdit(); }
    else if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clearRange(); }
    else if (event.key.length === 1 && !ctrl && !event.altKey) {
      event.preventDefault();
      startEdit(event.key);
    }
  });

  gridWrap.setAttribute('tabindex', '0');

  /* ── paste ───────────────────────────────────────────── */

  gridWrap.addEventListener('paste', function (event) {
    if (editing) return;
    var text = (event.clipboardData || window.clipboardData).getData('text');
    if (!text) return;
    event.preventDefault();
    pasteText(text);
  });

  /* ── formula bar ─────────────────────────────────────── */

  formulaInput.addEventListener('input', function () { formulaDirty = true; });

  formulaInput.addEventListener('blur', function () {
    if (formulaDirty) commitFormulaBar();
  });

  nameBox.addEventListener('focus', function () {
    var range = document.createRange();
    range.selectNodeContents(nameBox);
    var selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  });

  /* ── format toolbar ──────────────────────────────────── */

  function toggleFlag(prop) {
    var style = styleOf(sel.r, sel.c) || {};
    var patch = {};
    patch[prop] = style[prop] ? null : true;
    applyStyleToSelection(patch);
  }

  document.getElementById('cmd-bold').addEventListener('click', function () { toggleFlag('bold'); });
  document.getElementById('cmd-italic').addEventListener('click', function () { toggleFlag('italic'); });
  document.getElementById('cmd-underline').addEventListener('click', function () { toggleFlag('underline'); });
  document.getElementById('cmd-strike').addEventListener('click', function () { toggleFlag('strike'); });

  document.getElementById('font-family').addEventListener('change', function (event) {
    var value = event.target.value;
    applyStyleToSelection({ font: value || null });
  });

  document.getElementById('font-size').addEventListener('change', function (event) {
    var value = event.target.value;
    applyStyleToSelection({ size: value ? Number(value) : null });
  });

  // The visible button opens the hidden colour input, which carries the picker.
  document.getElementById('fill-button').addEventListener('click', function () {
    document.getElementById('fill-color').click();
  });
  document.getElementById('fill-color').addEventListener('input', function (event) {
    applyStyleToSelection({ fill: event.target.value.toUpperCase() });
  });
  document.getElementById('fill-color').addEventListener('change', function (event) {
    applyStyleToSelection({ fill: event.target.value.toUpperCase() });
  });

  document.getElementById('color-button').addEventListener('click', function () {
    document.getElementById('text-color').click();
  });
  document.getElementById('text-color').addEventListener('input', function (event) {
    applyStyleToSelection({ color: event.target.value.toUpperCase() });
  });
  document.getElementById('text-color').addEventListener('change', function (event) {
    applyStyleToSelection({ color: event.target.value.toUpperCase() });
  });

  document.getElementById('cmd-clear-fill').addEventListener('click', function () {
    applyStyleToSelection({ fill: null });
  });

  document.getElementById('cmd-reset').addEventListener('click', function () {
    var box = rangeOf();
    snapshot();
    for (var r = box.r1; r <= box.r2; r++) {
      for (var c = box.c1; c <= box.c2; c++) setStyleAt(r, c, null);
    }
    render();
    scheduleSave();
    notify('Format direset.', 'busy');
  });

  var alignButtons = document.querySelectorAll('[data-align]');
  for (var ai = 0; ai < alignButtons.length; ai++) {
    alignButtons[ai].addEventListener('click', function (event) {
      var want = event.currentTarget.dataset.align;
      var style = styleOf(sel.r, sel.c) || {};
      applyStyleToSelection({ align: style.align === want ? null : want });
    });
  }

  /* ── buttons ─────────────────────────────────────────── */

  document.getElementById('btn-save').addEventListener('click', function () { save(false); });
  document.getElementById('btn-add-row').addEventListener('click', function () { addRows(20); });
  document.getElementById('btn-add-col').addEventListener('click', function () { addColumns(1); });
  document.getElementById('btn-add-sheet').addEventListener('click', addSheet);

  var importInput = document.createElement('input');
  importInput.type = 'file';
  importInput.accept = '.xlsx';
  importInput.style.display = 'none';
  document.body.appendChild(importInput);

  document.getElementById('btn-import').addEventListener('click', function () { importInput.click(); });

  importInput.addEventListener('change', function () {
    var file = importInput.files[0];
    if (!file) return;
    importInput.value = '';
    var button = document.getElementById('btn-import');
    button.disabled = true;
    notify('Mengimpor…', 'busy');
    var form = new FormData();
    form.append('file', file);
    fetch(api('/import'), { method: 'POST', body: form })
      .then(function (response) {
        return response.json().catch(function () { return {}; }).then(function (result) {
          if (!response.ok) throw new Error(result.error || 'Import gagal.');
          return result;
        });
      })
      .then(function (result) {
        return call('/data').then(function (data) {
          activeSheet = 0;
          sel = { r: 0, c: 0 };
          anchor = { r: 0, c: 0 };
          applyData(data);
          notify('Diimpor: ' + result.sheet_count + ' sheet, ' + result.row_count + ' baris.', 'success');
        });
      })
      .catch(function (error) { notify(error.message, 'error'); })
      .then(function () { button.disabled = false; });
  });

  document.getElementById('btn-export').addEventListener('click', function () {
    var button = document.getElementById('btn-export');
    button.disabled = true;
    notify('Menyiapkan file…', 'busy');
    fetch(api('/export'), { cache: 'no-store' })
      .then(function (response) {
        if (!response.ok) throw new Error('Export gagal.');
        var disposition = response.headers.get('content-disposition') || '';
        var match = /filename="?([^"]+)"?/.exec(disposition);
        return response.blob().then(function (blob) {
          return { blob: blob, name: match ? match[1] : 'workspace.xlsx' };
        });
      })
      .then(function (file) {
        var url = URL.createObjectURL(file.blob);
        var link = document.createElement('a');
        link.href = url;
        link.download = file.name;
        link.click();
        URL.revokeObjectURL(url);
        notify('File diunduh.', 'success');
      })
      .catch(function (error) { notify(error.message, 'error'); })
      .then(function () { button.disabled = false; });
  });

  /* ── analysis view ───────────────────────────────────── */

  var analysisEl = document.getElementById('analysis');
  var analysisDirty = false;

  /* Remembered per browser so an accountant who prefers the grid keeps it,
     while a first-time visitor always lands on Workspace. */
  var MODE_KEY = 'workspace.view';

  function rememberMode(mode) {
    try { window.localStorage.setItem(MODE_KEY, mode); } catch (error) { /* private mode */ }
  }

  function savedMode() {
    try { return window.localStorage.getItem(MODE_KEY); } catch (error) { return null; }
  }

  function analysisFieldsFromPanel() {
    var panel = document.getElementById('aw-source');
    if (!panel) return null;
    function val(key) {
      var el = panel.querySelector('[data-field="' + key + '"]');
      if (!el || el.value === '') return null;
      return Number(el.value);
    }
    var fields = {
      sheet: Number((panel.querySelector('[data-sheet]') || {}).value || 0),
      headerRow: workbook.analysis && workbook.analysis.fields ? workbook.analysis.fields.headerRow : 0,
      date: val('date'),
      income: val('income'),
      expense: val('expense'),
      amount: val('amount'),
      category: val('category'),
      note: val('note')
    };
    if (fields.income === null && fields.expense === null && fields.amount === null) return null;
    return fields;
  }

  function renderAnalysis() {
    if (!window.WorkspaceAnalysis || !analysisEl) return;
    var state = window.WorkspaceAnalysis.render(analysisEl, workbook, workbook.analysis);
    if (!workbook.analysis || !workbook.analysis.fields) {
      // First visit on a workbook that has never been analysed: keep the
      // detection, but do not write it back until the user changes something.
      if (state.fields) workbook.analysis = { fields: state.fields, widgets: state.widgets };
      else workbook.analysis = workbook.analysis || { widgets: state.widgets };
    }
  }

  function setMode(mode) {
    var isAnalysis = mode === 'analysis';
    analysisEl.hidden = !isAnalysis;
    gridWrap.hidden = isAnalysis;
    document.querySelector('.formula-row').hidden = isAnalysis;
    document.querySelector('.toolbar').hidden = isAnalysis;
    document.querySelector('.tabbar').hidden = isAnalysis;
    document.getElementById('mode-sheet').setAttribute('aria-selected', isAnalysis ? 'false' : 'true');
    document.getElementById('mode-analysis').setAttribute('aria-selected', isAnalysis ? 'true' : 'false');
    if (isAnalysis) {
      renderAnalysis();
    } else {
      gridWrap.focus();
    }
  }

  /* A workbook with no money column has nothing to show on Workspace, so it
     opens on the grid instead of an empty dashboard. */
  function canAnalyse() {
    if (!window.WorkspaceAnalysis) return false;
    if (workbook.analysis && workbook.analysis.fields) return true;
    return Boolean(window.WorkspaceAnalysis.detectFields(workbook));
  }

  function applyInitialMode() {
    if (savedMode() === 'sheet') { setMode('sheet'); return; }
    setMode(canAnalyse() ? 'analysis' : 'sheet');
  }

  function saveAnalysis() {
    if (!workbook.analysis || !workbook.analysis.fields) return Promise.resolve();
    return call('/analysis/widgets/save', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ fields: workbook.analysis.fields, widgets: workbook.analysis.widgets })
    }).catch(function (error) { notify(error.message, 'error'); });
  }

  function generateWidgets() {
    var input = document.getElementById('aw-input');
    var button = document.getElementById('aw-generate');
    var prompt = (input.value || '').trim();
    if (!prompt) { notify('Tulis dulu mau lihat analisis apa.', 'error'); input.focus(); return; }
    button.disabled = true;
    button.textContent = 'Menyusun…';
    call('/analysis/widgets', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ prompt: prompt })
    }).then(function (result) {
      if (!Array.isArray(result.widgets) || !result.widgets.length) {
        throw new Error(result.note || 'AI tidak mengembalikan widget.');
      }
      if (!workbook.analysis || !workbook.analysis.fields) {
        workbook.analysis = { fields: window.WorkspaceAnalysis.detectFields(workbook), widgets: [] };
      }
      workbook.analysis.widgets = result.widgets;
      renderAnalysis();
      return saveAnalysis().then(function () {
        notify(result.source === 'ai'
          ? 'Widget dibuat AI.'
          : (result.note || 'Widget dibuat dari kata kunci.'), result.source === 'ai' ? 'success' : 'busy');
      });
    }).catch(function (error) {
      notify(error.message, 'error');
    }).then(function () {
      button.disabled = false;
      button.textContent = 'Buat widget';
    });
  }

  document.getElementById('mode-sheet').addEventListener('click', function () { setMode('sheet'); rememberMode('sheet'); });
  document.getElementById('mode-analysis').addEventListener('click', function () { setMode('analysis'); rememberMode('analysis'); });

  // "Buka Sheet" inside the Workspace status bar.
  var statusHost = document.getElementById('aw-status');
  if (statusHost) {
    statusHost.addEventListener('click', function (event) {
      var button = event.target.closest('[data-mode]');
      if (!button) return;
      setMode(button.dataset.mode);
      rememberMode(button.dataset.mode);
    });
  }
  document.getElementById('aw-generate').addEventListener('click', generateWidgets);

  document.getElementById('aw-input').addEventListener('keydown', function (event) {
    if (event.key === 'Enter') { event.preventDefault(); generateWidgets(); }
  });

  document.getElementById('aw-auto').addEventListener('click', function () {
    var fields = analysisFieldsFromPanel() || window.WorkspaceAnalysis.detectFields(workbook);
    if (!fields) { notify('Kolom pemasukan/pengeluaran belum ketemu. Pilih manual di panel Sumber Data.', 'error'); return; }
    workbook.analysis = { fields: fields, widgets: window.WorkspaceAnalysis.autoWidgets() };
    renderAnalysis();
    saveAnalysis();
    notify('Widget standar dimuat.', 'busy');
  });

  // Any change in the Sumber Data panel re-binds the columns and re-renders.
  document.getElementById('aw-source').addEventListener('change', function (event) {
    var select = event.target.closest('select');
    if (!select) return;

    // Switching sheet invalidates every column index, so re-detect for it.
    if (select.dataset.sheet) {
      var detected = window.WorkspaceAnalysis.detectForSheet(workbook, Number(select.value));
      if (!detected) { notify('Sheet ini belum punya data.', 'error'); renderAnalysis(); return; }
      workbook.analysis = {
        fields: detected,
        widgets: (workbook.analysis && workbook.analysis.widgets && workbook.analysis.widgets.length)
          ? workbook.analysis.widgets
          : window.WorkspaceAnalysis.autoWidgets()
      };
      renderAnalysis();
      persistAnalysis();
      return;
    }

    var fields = analysisFieldsFromPanel();
    if (!fields) {
      notify('Pilih minimal kolom Pemasukan, Pengeluaran, atau Nominal.', 'error');
      renderAnalysis();
      return;
    }
    workbook.analysis = workbook.analysis || {};
    workbook.analysis.fields = fields;
    if (!workbook.analysis.widgets || !workbook.analysis.widgets.length) {
      workbook.analysis.widgets = window.WorkspaceAnalysis.autoWidgets();
    }
    renderAnalysis();
    persistAnalysis();
  });

  /* Store the binding, then the widget list, and report once. */
  function persistAnalysis() {
    if (!workbook.analysis || !workbook.analysis.fields) return;
    call('/analysis/fields', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ fields: workbook.analysis.fields })
    }).then(function () {
      return call('/analysis/widgets/save', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ fields: workbook.analysis.fields, widgets: workbook.analysis.widgets || [] })
      });
    }).then(function () {
      notify('Sumber data disimpan.', 'success');
    }).catch(function (error) { notify(error.message, 'error'); });
  }

  // Widget delete buttons are re-created on every render, so listen on the grid.
  document.getElementById('aw-grid').addEventListener('click', function (event) {
    var button = event.target.closest('[data-del]');
    if (!button) return;
    var id = button.dataset.del;
    if (!workbook.analysis || !Array.isArray(workbook.analysis.widgets)) return;
    workbook.analysis.widgets = workbook.analysis.widgets.filter(function (widget) { return widget.id !== id; });
    renderAnalysis();
    saveAnalysis();
    notify('Widget dihapus.', 'busy');
  });

  /* ── boot ────────────────────────────────────────────── */

  function applyData(data) {
    workbook = data && Array.isArray(data.sheets) && data.sheets.length
      ? data
      : { title: boot.title || 'Workspace', sheets: [{ name: 'Sheet1', columns: ['A', 'B'], rows: [] }] };
    workbook.sheets.forEach(function (s) {
      s.rows = Array.isArray(s.rows) ? s.rows : [];
      s.columns = Array.isArray(s.columns) && s.columns.length ? s.columns : ['A', 'B'];
    });
    if (activeSheet >= workbook.sheets.length) activeSheet = 0;
    render();
  }

  applyData(boot.data);
  applyInitialMode();

  call('/data').then(function (data) {
    applyData(data);
    applyInitialMode();
  }).catch(function (error) {
    notify(error.message, 'error');
  });
})();

import { readFileSync, writeFileSync } from 'node:fs';
import { importWorkspaceXlsx } from './src/workspace-sheets.js';

const dataset = await importWorkspaceXlsx(
  readFileSync('Master_Excel_Semua_Fitur_Rumus_Sistem.xlsx'),
  'Master Excel',
);
const css = readFileSync('public/workspace-app.css', 'utf8') + '\n' + readFileSync('public/workspace-analysis.css', 'utf8');
const app = readFileSync('public/workspace-app.js', 'utf8');
const analysis = readFileSync('public/workspace-analysis.js', 'utf8');
const shell = readFileSync('src/workspace-page.js', 'utf8');
const markup = /<body>([\s\S]*?)<script>window\.__WORKSPACE__/.exec(shell)[1];
const dataJson = JSON.stringify(dataset).replace(/</g, '\\u003c');

const harness = `<script>
(function () {
  var out = [];
  function ok(name, condition) { out.push((condition ? "PASS " : "FAIL ") + name); }
  function report() {
    var pre = document.createElement("pre");
    pre.id = "harness-result";
    pre.textContent = out.join("\\n");
    document.body.appendChild(pre);
  }
  try {

  // Default view: Workspace when the workbook can be analysed, grid otherwise.
  var analysisEl = document.getElementById("analysis");
  var gridEl = document.getElementById("grid-wrap");
  var probeFields = window.WorkspaceAnalysis.detectFields(window.__WORKSPACE__.data);
  var canAnalyse = Boolean(probeFields);
  ok("default view follows the data (" + (canAnalyse ? "workspace" : "sheet") + ")",
    analysisEl.hidden === !canAnalyse && gridEl.hidden === canAnalyse);
  ok("sheet toolbar hidden on the default view",
    document.querySelector(".toolbar").hidden === canAnalyse);
  if (canAnalyse) {
    ok("Workspace tab is selected",
      document.getElementById("mode-analysis").getAttribute("aria-selected") === "true");
    ok("status bar names the book", (document.getElementById("aw-status").textContent || "").length > 8);
    ok("status bar offers a way back to the grid",
      Boolean(document.querySelector("#aw-status [data-mode=sheet]")));
    var probeLedger = window.WorkspaceAnalysis.buildLedger(window.__WORKSPACE__.data, probeFields);
    var probeNotes = window.WorkspaceAnalysis.buildInsights(probeLedger).length;
    ok("notes rendered (" + probeNotes + ")",
      document.querySelectorAll("#aw-insight .aw-insight-item").length === probeNotes);
    ok("source panel starts collapsed",
      document.getElementById("aw-source").closest("details").open === false);
    ok("default widgets include the latest transactions",
      document.querySelectorAll('#aw-grid [data-widget="w7"]').length === 1);

    // Layout: nothing may overflow sideways at this window width.
    ok("no page overflow at " + window.innerWidth + "px",
      document.documentElement.scrollWidth <= window.innerWidth + 1);
    ok("analysis pane does not scroll sideways",
      analysisEl.scrollWidth <= analysisEl.clientWidth + 1);
    var cards = Array.prototype.slice.call(document.querySelectorAll("#aw-grid .aw-card"));
    ok("every card has a usable width",
      cards.length > 0 && cards.every(function (card) { return card.getBoundingClientRect().width > 60; }));
    var perTop = {};
    cards.forEach(function (card) {
      var top = Math.round(card.getBoundingClientRect().top);
      perTop[top] = (perTop[top] || 0) + 1;
    });
    var widestRow = Math.max.apply(null, Object.keys(perTop).map(function (top) { return perTop[top]; }));
    ok("cards share a row at " + window.innerWidth + "px (max " + widestRow + ")", widestRow >= 2);
    ok("status bar stays inside the pane",
      document.getElementById("aw-status").getBoundingClientRect().width <= analysisEl.clientWidth + 1);
  }
  document.getElementById("mode-sheet").click();
  ok("grid visible after switching to Sheet", !gridEl.hidden && analysisEl.hidden);
  try { window.localStorage.clear(); } catch (error) { /* file:// may deny storage */ }

  var nameBox = document.getElementById("name-box");
  var formulaInput = document.getElementById("formula-input");
  var tabs = document.getElementById("tabs");
  var menu = document.getElementById("menu");

  function cellAt(r, c) {
    return document.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
  }
  function click(el, extra) {
    el.dispatchEvent(new MouseEvent("mousedown", Object.assign({ bubbles: true, clientX: 10, clientY: 10 }, extra || {})));
  }
  function key(k, opts) {
    document.dispatchEvent(new KeyboardEvent("keydown", Object.assign({ key: k, bubbles: true }, opts || {})));
  }

  ok("grid has cells", document.querySelectorAll("td[data-r]").length > 1000);
  ok("no raw formula text", !/[>]=[A-Z]/.test(document.body.innerHTML));

  var colA = document.querySelector('th[data-col="0"]');
  var colC = document.querySelector('th[data-col="2"]');
  ok("column header A", colA && colA.textContent.trim().indexOf("A") === 0);
  ok("column header C", colC && colC.textContent.trim().indexOf("C") === 0);
  var rowHead = document.querySelector('th[data-row="4"]');
  ok("row header 5", rowHead && rowHead.textContent.trim() === "5");

  ok("merged title spans columns", document.querySelectorAll('td.span').length > 0);

  // Auto-fit: widths must follow content, not a fixed default.
  var widths = {};
  document.querySelectorAll('th[data-col]').forEach(function (th) {
    widths[th.dataset.col] = Math.round(th.getBoundingClientRect().width);
  });
  var distinct = {};
  Object.keys(widths).forEach(function (k) { distinct[widths[k]] = true; });
  ok("columns are auto-sized (varied widths)", Object.keys(distinct).length > 3);
  ok("no column narrower than minimum", Object.keys(widths).every(function (k) { return widths[k] >= 54; }));
  ok("no column wider than maximum", Object.keys(widths).every(function (k) { return widths[k] <= 340; }));

  // Nothing that has content should be clipped.
  var clipped = 0;
  document.querySelectorAll('td[data-r]').forEach(function (td) {
    if (!td.textContent) return;
    if (td.scrollWidth > td.clientWidth + 1) clipped++;
  });
  ok("filled cells are not truncated (" + clipped + " clipped)", clipped === 0);

  var target = cellAt(4, 3);
  click(target);
  ok("name box shows D5", nameBox.textContent === "D5");
  ok("active cell marked", target.classList.contains("active"));
  ok("formula bar shows formula", formulaInput.textContent.indexOf("=") === 0);

  key("ArrowDown");
  ok("arrow down moves to D6", nameBox.textContent === "D6");
  key("ArrowRight");
  ok("arrow right moves to E6", nameBox.textContent === "E6");

  nameBox.textContent = "B10";
  nameBox.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
  ok("name box jump to B10", nameBox.textContent === "B10");

  click(cellAt(0, 0));
  click(cellAt(4, 1), { shiftKey: true });
  ok("range label A1:B5", nameBox.textContent === "A1:B5");

  click(cellAt(0, 0));
  key("Z");
  var active = document.querySelector('td[contenteditable="true"]');
  ok("edit mode entered", Boolean(active));
  if (active) {
    active.textContent = "HALO";
    key("Enter");
    ok("edited value stored", cellAt(0, 0).textContent === "HALO");
  }

  click(cellAt(1, 0));
  key("=");
  var editor = document.querySelector('td[contenteditable="true"]');
  if (editor) {
    editor.textContent = "=1+2";
    key("Enter");
    ok("formula cell not raw", cellAt(1, 0).textContent !== "=1+2");
  }

  var tabButtons = tabs.querySelectorAll("button");
  ok("six sheet tabs", tabButtons.length === 6);
  tabButtons[1].click();
  ok("switched to sheet 2", tabs.querySelector('button[aria-selected="true"]').textContent === "02_Rumus_Matematika_Statistik");
  ok("second sheet has data", cellAt(0, 0).textContent.length > 0);

  cellAt(2, 2).dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, clientX: 40, clientY: 40 }));
  ok("context menu opens", menu.classList.contains("open"));
  document.body.click();

  // â”€â”€ formatting â”€â”€
  // The imported workbook already carries real Excel formatting, so these
  // checks assert that a control *changes* the cell, not that it starts bare.
  function weight(el) { return getComputedStyle(el).fontWeight; }

  var fmtCell = cellAt(6, 4);
  click(fmtCell);
  var weightBefore = weight(cellAt(6, 4));
  document.getElementById("cmd-bold").click();
  var weightAfter = weight(cellAt(6, 4));
  ok("bold toggles the cell", weightBefore !== weightAfter);
  ok("bold button reflects state", document.getElementById("cmd-bold").getAttribute("aria-pressed") ===
    (weightAfter === "700" ? "true" : "false"));
  document.getElementById("cmd-bold").click();
  ok("bold toggles back", weight(cellAt(6, 4)) === weightBefore);

  click(cellAt(6, 5));
  var fill = document.getElementById("fill-color");
  fill.value = "#fff2cc";
  fill.dispatchEvent(new Event("input", { bubbles: true }));
  var fillCell = cellAt(6, 5);
  ok("fill applied [" + (fillCell && fillCell.getAttribute("style")) + "]",
    /#FFF2CC/i.test(fillCell.getAttribute("style") || ""));
  ok("fill swatch updated", /255|fff2cc/i.test(document.getElementById("fill-swatch").style.background));

  var textColor = document.getElementById("text-color");
  textColor.value = "#c00000";
  textColor.dispatchEvent(new Event("input", { bubbles: true }));
  ok("text colour applied [" + cellAt(6, 5).getAttribute("style") + "]",
    /#C00000/i.test(cellAt(6, 5).getAttribute("style") || ""));

  click(cellAt(6, 6));
  document.getElementById("font-family").value = "Arial";
  document.getElementById("font-family").dispatchEvent(new Event("change", { bubbles: true }));
  ok("font family applied", /Arial/i.test(cellAt(6, 6).style.fontFamily));

  document.getElementById("font-size").value = "18";
  document.getElementById("font-size").dispatchEvent(new Event("change", { bubbles: true }));
  ok("font size applied", cellAt(6, 6).style.fontSize === "18px");

  click(cellAt(6, 7));
  document.querySelector('[data-align="center"]').click();
  ok("centre align applied", getComputedStyle(cellAt(6, 7)).textAlign === "center");

  // Format a range at once
  click(cellAt(8, 3));
  click(cellAt(9, 4), { shiftKey: true });
  document.getElementById("cmd-italic").click();
  ok("range italic (8,3)", getComputedStyle(cellAt(8, 3)).fontStyle === "italic");
  ok("range italic (9,4)", getComputedStyle(cellAt(9, 4)).fontStyle === "italic");

  // Clear fill returns the cell to the sheet default background
  click(cellAt(6, 5));
  document.getElementById("cmd-clear-fill").click();
  ok("clear fill works [" + cellAt(6, 5).getAttribute("style") + "]",
    !/#FFF2CC/i.test(cellAt(6, 5).getAttribute("style") || ""));

  // Reset removes every property
  click(cellAt(6, 7));
  document.getElementById("cmd-reset").click();
  ok("reset clears formatting", !cellAt(6, 7).getAttribute("style"));

  // Ctrl+B works as a shortcut
  click(cellAt(10, 6));
  var beforeCtrl = weight(cellAt(10, 6));
  var neighbourBefore = cellAt(10, 7).getAttribute("style");
  key("b", { ctrlKey: true });
  ok("ctrl+B toggles bold", weight(cellAt(10, 6)) !== beforeCtrl);

  // Styles must not leak into neighbouring cells
  ok("neighbour untouched", cellAt(10, 7).getAttribute("style") === neighbourBefore);

  // Imported Excel colours must actually reach the DOM
  var painted = 0;
  document.querySelectorAll("td[data-r]").forEach(function (td) {
    var bg = getComputedStyle(td).backgroundColor;
    if (bg && bg !== "rgba(0, 0, 0, 0)") painted++;
  });
  ok("imported fills are rendered (" + painted + ")", painted > 20);

  var wrap = document.getElementById("grid-wrap");
  wrap.scrollTop = 400;
  wrap.scrollLeft = 250;
  var wrapRect = wrap.getBoundingClientRect();
  var headRect = document.querySelector('th[data-col="0"]').getBoundingClientRect();
  var rowRect = document.querySelector('th[data-row="0"]').getBoundingClientRect();
  ok("column header stays at top", Math.abs(headRect.y - wrapRect.y) < 2);
  ok("row header stays at left", Math.abs(rowRect.x - wrapRect.x) < 2);

  report();
  } catch (error) {
    out.push("ERROR " + (error && error.stack ? error.stack.split("\\n").slice(0, 4).join(" | ") : error));
    report();
  }
})();
</script>`;

function page(extra) {
  return `<!doctype html>
<html lang="id"><head><meta charset="utf-8"><title>Pratinjau Workspace</title>
<style>${css}</style></head>
<body>${markup}
<script>window.__WORKSPACE__ = { base: "", title: "Master Excel", data: ${dataJson} };
(function () {
  function reply(payload, status) {
    return Promise.resolve(new Response(JSON.stringify(payload), {
      status: status || 200, headers: { "content-type": "application/json" } }));
  }
  window.fetch = function (url, options) {
    var target = String(url), opts = options || {};
    if (target.indexOf("/data") === 0) return reply(window.__WORKSPACE__.data);
    if (target.indexOf("/workbook") === 0) {
      var body = JSON.parse(opts.body || "{}");
      return reply({ saved: true, sheets: body.sheets || [], sheet_names: [] });
    }
    return reply({ error: "Pratinjau statis: import/export hanya jalan di server." }, 400);
  };
})();</script>
<script>${analysis}</script>
<script>${app}</script>
${extra}
</body></html>`;
}

writeFileSync('workspace-preview.html', page(''));
writeFileSync('workspace-harness.html', page(harness));
console.log('preview + harness written; sheets:', dataset.sheets.length);

import assert from "node:assert/strict";
import test from "node:test";
import { existsSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { exportWorkspaceXlsx, importWorkspaceXlsx, validateWorkspace } from "../src/workspace-sheets.js";
import { sanitizeStyle, sanitizeStyles, readExcelStyle, toExcelStyle } from "../src/cell-styles.js";

// Optional fixture: the repo copy may be absent on a deploy machine, so the
// round-trip test that needs it skips rather than reporting a false failure.
const masterPath = fileURLToPath(new URL("../Master_Excel_Semua_Fitur_Rumus_Sistem.xlsx", import.meta.url));
const needsMaster = existsSync(masterPath) ? {} : { skip: `fixture not found at ${masterPath}` };

test("cell styles survive an export and import round trip", async () => {
  const data = validateWorkspace({
    title: "Warna",
    sheets: [{
      name: "Data",
      columns: ["A", "B"],
      rows: [["Judul", "Nilai"], ["Baris", 10]],
      styles: {
        "0,0": { fill: "#FFF2CC", bold: true, color: "#C00000", size: 14, font: "Arial" },
        "0,1": { italic: true, align: "center" },
        "1,0": { underline: true, strike: true },
        "1,1": { fill: "#D9E1F2", align: "right", format: "#,##0" },
      },
    }],
  });

  const bytes = await exportWorkspaceXlsx(data);
  assert.deepEqual([...bytes].slice(0, 2), [0x50, 0x4b]);

  const reimported = await importWorkspaceXlsx(bytes, "Warna");
  const styles = reimported.sheets[0].styles;

  assert.equal(styles["0,0"].fill, "#FFF2CC");
  assert.equal(styles["0,0"].color, "#C00000");
  assert.equal(styles["0,0"].bold, true);
  assert.equal(styles["0,0"].font, "Arial");
  assert.equal(styles["0,0"].size, 14);
  assert.equal(styles["0,1"].italic, true);
  assert.equal(styles["0,1"].align, "center");
  assert.equal(styles["1,0"].underline, true);
  assert.equal(styles["1,0"].strike, true);
  assert.equal(styles["1,1"].fill, "#D9E1F2");
  assert.equal(styles["1,1"].align, "right");
  assert.equal(styles["1,1"].format, "#,##0");
});

test("unstyled sheets stay free of a styles map", async () => {
  const data = validateWorkspace({ title: "Polos", sheets: [{ name: "S", columns: ["A"], rows: [["x"]] }] });
  assert.equal(data.sheets[0].styles, undefined);
  const bytes = await exportWorkspaceXlsx(data);
  const back = await importWorkspaceXlsx(bytes, "Polos");
  assert.equal(back.sheets[0].styles, undefined);
});

test("styles on cells without values are still kept", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("S");
  ws.getCell("B2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC6EFCE" } };
  ws.getCell("B2").font = { bold: true, color: { argb: "FF006100" } };
  const bytes = Buffer.from(await wb.xlsx.writeBuffer());

  const imported = await importWorkspaceXlsx(bytes, "Kosong");
  const style = imported.sheets[0].styles["1,1"];
  assert.equal(style.fill, "#C6EFCE");
  assert.equal(style.color, "#006100");
  assert.equal(style.bold, true);
});

test("only Excel-supported style properties are accepted", () => {
  assert.equal(sanitizeStyle({ madeUp: true }), null);
  assert.equal(sanitizeStyle({ fill: "not-a-colour" }), null);
  assert.equal(sanitizeStyle({ fill: "#ff0000" }).fill, "#FF0000");
  assert.deepEqual(sanitizeStyle({ bold: true, align: "justify" }), { bold: true });
  assert.deepEqual(sanitizeStyle({ align: "center", size: 999 }), { align: "center", size: 96 });
  assert.deepEqual(sanitizeStyle({ size: 2 }), { size: 6 });
  assert.equal(sanitizeStyle({ wrap: "yes" }), null);

  const map = sanitizeStyles({ "0,0": { bold: true }, "bad-key": { bold: true }, "1,2": { fill: "#00FF00" } });
  assert.deepEqual(Object.keys(map).sort(), ["0,0", "1,2"]);
});

test("Excel theme colours are resolved to hex", () => {
  assert.equal(readExcelStyle({ font: { color: { theme: 1 } } }).color, "#000000");
  assert.equal(readExcelStyle({ font: { color: { theme: 0 } } }).color, "#FFFFFF");
  const tinted = readExcelStyle({ fill: { type: "pattern", pattern: "solid", fgColor: { theme: 4, tint: 0.4 } } });
  assert.match(tinted.fill, /^#[0-9A-F]{6}$/);
});

test("gradient and pattern fills are ignored, not corrupted", () => {
  assert.equal(readExcelStyle({ fill: { type: "gradient" } }), null);
  assert.equal(readExcelStyle({ fill: { type: "pattern", pattern: "none" } }), null);
});

test("toExcelStyle writes real Excel formatting objects", () => {
  const style = toExcelStyle({ fill: "#FFF2CC", color: "#C00000", bold: true, size: 14, align: "center", format: "0.00" });
  assert.equal(style.fill.pattern, "solid");
  assert.equal(style.fill.fgColor.argb, "FFFFF2CC");
  assert.equal(style.font.color.argb, "FFC00000");
  assert.equal(style.font.bold, true);
  assert.equal(style.font.size, 14);
  assert.equal(style.alignment.horizontal, "center");
  assert.equal(style.numFmt, "0.00");
});

test("borders round-trip through Excel", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("S");
  ws.getCell("A1").value = "kotak";
  ws.getCell("A1").border = {
    top: { style: "medium", color: { argb: "FF1F4E79" } },
    bottom: { style: "double", color: { argb: "FFC00000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
  };
  const imported = await importWorkspaceXlsx(Buffer.from(await wb.xlsx.writeBuffer()), "Border");
  const border = imported.sheets[0].styles["0,0"].border;
  assert.equal(border.top.style, "medium");
  assert.equal(border.top.color, "#1F4E79");
  assert.equal(border.bottom.style, "double");
  assert.equal(border.bottom.color, "#C00000");
  assert.equal(border.left.style, "thin");
  assert.equal(border.right, undefined);

  const bytes = await exportWorkspaceXlsx(imported);
  const back = await importWorkspaceXlsx(bytes, "Border");
  const again = back.sheets[0].styles["0,0"].border;
  assert.equal(again.top.style, "medium");
  assert.equal(again.top.color, "#1F4E79");
  assert.equal(again.bottom.style, "double");
});

test("the full Master Excel keeps its formatting through export and import", needsMaster, async () => {
  const source = readFileSync(masterPath);
  const first = await importWorkspaceXlsx(source, "Master");
  const styledBefore = first.sheets.reduce((n, s) => n + Object.keys(s.styles ?? {}).length, 0);
  assert.ok(styledBefore > 500, "source should carry formatting, got " + styledBefore);

  const bytes = await exportWorkspaceXlsx(first);
  const second = await importWorkspaceXlsx(bytes, "Master");
  const styledAfter = second.sheets.reduce((n, s) => n + Object.keys(s.styles ?? {}).length, 0);
  assert.ok(styledAfter >= styledBefore * 0.98, "styles lost on export: " + styledBefore + " -> " + styledAfter);

  // Values and formulas must survive too, not just formatting.
  for (let i = 0; i < first.sheets.length; i++) {
    assert.equal(second.sheets[i].name, first.sheets[i].name);
    assert.equal(second.sheets[i].rows.length, first.sheets[i].rows.length);
  }

  // Every colour in the export must be a colour the source already had.
  const coloursBefore = new Set();
  for (const sheet of first.sheets) for (const style of Object.values(sheet.styles ?? {})) {
    if (style.fill) coloursBefore.add(style.fill);
  }
  const coloursAfter = new Set();
  for (const sheet of second.sheets) for (const style of Object.values(sheet.styles ?? {})) {
    if (style.fill) coloursAfter.add(style.fill);
  }
  for (const colour of coloursAfter) {
    assert.ok(coloursBefore.has(colour), "export invented a colour the source did not have: " + colour);
  }
});

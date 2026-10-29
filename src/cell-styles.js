/** Cell styling that round-trips to Excel.
 *
 *  Only properties Excel itself has are stored, so an exported file never
 *  contains formatting the source workbook did not have. Everything is
 *  optional: a cell with no style entry exports exactly as before.
 */

const MAX_STYLE_ENTRIES = 200000;
const MAX_FONT_NAME = 60;
const MAX_FORMAT = 80;
const HEX = /^#[0-9a-fA-F]{6}$/;

export const ALIGN_VALUES = new Set(["left", "center", "right"]);
export const VALIGN_VALUES = new Set(["top", "middle", "bottom"]);

/** Excel border styles we round-trip. Anything else falls back to thin. */
export const BORDER_STYLES = new Set([
  "thin", "medium", "thick", "hair", "dotted", "dashed", "double",
  "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot",
]);
export const BORDER_SIDES = ["top", "right", "bottom", "left"];

/** Default Office theme palette, used when a file stores a theme colour. */
const THEME_COLORS = [
  [255, 255, 255],
  [0, 0, 0],
  [242, 242, 242],
  [127, 127, 127],
  [68, 114, 196],
  [237, 125, 49],
  [165, 165, 165],
  [255, 192, 0],
  [91, 155, 213],
  [112, 173, 71],
  [5, 99, 193],
  [149, 79, 114],
];

function clampByte(value) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function toHex(rgb) {
  return "#" + rgb.map((channel) => clampByte(channel).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** Excel stores colours as ARGB. */
function argbToHex(argb) {
  if (typeof argb !== "string") return null;
  const value = argb.replace(/^#/, "");
  if (value.length === 8) {
    const hex = "#" + value.slice(2).toUpperCase();
    return HEX.test(hex) ? hex : null;
  }
  if (value.length === 6) {
    const hex = "#" + value.toUpperCase();
    return HEX.test(hex) ? hex : null;
  }
  return null;
}

/** Apply an Excel tint (-1 darkens to black, 1 lightens to white). */
function applyTint(rgb, tint) {
  if (!tint) return rgb;
  if (tint < 0) return rgb.map((channel) => channel * (1 + tint));
  return rgb.map((channel) => channel + (255 - channel) * tint);
}

/** Read any ExcelJS colour object into #RRGGBB. */
function readColor(color) {
  if (!color) return null;
  if (typeof color === "string") return argbToHex(color);
  if (typeof color.argb === "string") return argbToHex(color.argb);
  if (typeof color.theme === "number") {
    const base = THEME_COLORS[color.theme];
    if (!base) return null;
    return toHex(applyTint(base, Number(color.tint) || 0));
  }
  if (typeof color.indexed === "number") {
    const base = THEME_COLORS[color.indexed];
    if (base) return toHex(applyTint(base, Number(color.tint) || 0));
  }
  return null;
}

/** Extract the style Excel actually applied to a cell. Returns null when the
 *  cell carries no formatting, so unstyled sheets stay small. */
export function readExcelStyle(cell) {
  const style = {};

  const fill = cell.fill;
  if (fill && fill.type === "pattern" && fill.pattern === "solid") {
    const hex = readColor(fill.fgColor);
    if (hex) style.fill = hex;
  }

  const font = cell.font;
  if (font) {
    const color = readColor(font.color);
    if (color) style.color = color;
    if (font.bold) style.bold = true;
    if (font.italic) style.italic = true;
    if (font.underline) style.underline = true;
    if (font.strike) style.strike = true;
    if (typeof font.name === "string" && font.name.trim()) style.font = font.name.trim().slice(0, MAX_FONT_NAME);
    if (typeof font.size === "number" && font.size > 0) style.size = Math.min(96, Math.max(6, Math.round(font.size * 10) / 10));
  }

  const alignment = cell.alignment;
  if (alignment) {
    if (ALIGN_VALUES.has(alignment.horizontal)) style.align = alignment.horizontal;
    if (VALIGN_VALUES.has(alignment.vertical)) style.valign = alignment.vertical;
    if (alignment.wrapText) style.wrap = true;
  }

  if (typeof cell.numFmt === "string" && cell.numFmt && cell.numFmt !== "General") {
    style.format = cell.numFmt.slice(0, MAX_FORMAT);
  }

  const border = readBorder(cell.border);
  if (border) style.border = border;

  return Object.keys(style).length ? style : null;
}

/** Read the four edges Excel actually drew. Cells with no drawn edge return
 *  null so the map stays small on plain sheets. */
function readBorder(border) {
  if (!border || typeof border !== "object") return null;
  const out = {};
  for (const side of BORDER_SIDES) {
    const edge = border[side];
    if (!edge || !edge.style) continue;
    const hex = readColor(edge.color) ?? "#000000";
    out[side] = { style: BORDER_STYLES.has(edge.style) ? edge.style : "thin", color: hex };
  }
  return Object.keys(out).length ? out : null;
}

function writeBorder(border) {
  if (!border || typeof border !== "object") return null;
  const out = {};
  for (const side of BORDER_SIDES) {
    const edge = border[side];
    if (!edge || !BORDER_STYLES.has(edge.style)) continue;
    const color = typeof edge.color === "string" && HEX.test(edge.color) ? edge.color : "#000000";
    out[side] = { style: edge.style, color: { argb: "FF" + color.slice(1).toUpperCase() } };
  }
  return Object.keys(out).length ? out : null;
}

/** Build the ExcelJS style object for one of our style entries. */
export function toExcelStyle(style) {
  const out = {};

  if (style.fill && HEX.test(style.fill)) {
    out.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + style.fill.slice(1).toUpperCase() } };
  }

  const font = {};
  if (style.color && HEX.test(style.color)) font.color = { argb: "FF" + style.color.slice(1).toUpperCase() };
  if (style.bold) font.bold = true;
  if (style.italic) font.italic = true;
  if (style.underline) font.underline = true;
  if (style.strike) font.strike = true;
  if (typeof style.font === "string" && style.font.trim()) font.name = style.font.trim().slice(0, MAX_FONT_NAME);
  if (typeof style.size === "number" && style.size > 0) font.size = Math.min(96, Math.max(6, style.size));
  if (Object.keys(font).length) out.font = font;

  const alignment = {};
  if (ALIGN_VALUES.has(style.align)) alignment.horizontal = style.align;
  if (VALIGN_VALUES.has(style.valign)) alignment.vertical = style.valign;
  if (style.wrap) alignment.wrapText = true;
  if (Object.keys(alignment).length) out.alignment = alignment;

  if (typeof style.format === "string" && style.format.trim()) out.numFmt = style.format.slice(0, MAX_FORMAT);

  const border = writeBorder(style.border);
  if (border) out.border = border;

  return out;
}

/** Validate one incoming style object from the browser. Unknown keys are
 *  dropped, so the client cannot smuggle in anything Excel does not support. */
export function sanitizeStyle(raw) {
  if (!raw || typeof raw !== "object") return null;
  const style = {};

  if (typeof raw.fill === "string" && HEX.test(raw.fill)) style.fill = raw.fill.toUpperCase();
  if (typeof raw.color === "string" && HEX.test(raw.color)) style.color = raw.color.toUpperCase();
  if (raw.bold === true) style.bold = true;
  if (raw.italic === true) style.italic = true;
  if (raw.underline === true) style.underline = true;
  if (raw.strike === true) style.strike = true;
  if (typeof raw.font === "string" && raw.font.trim()) style.font = raw.font.trim().slice(0, MAX_FONT_NAME);
  if (typeof raw.size === "number" && Number.isFinite(raw.size)) {
    style.size = Math.min(96, Math.max(6, Math.round(raw.size * 10) / 10));
  }
  if (ALIGN_VALUES.has(raw.align)) style.align = raw.align;
  if (VALIGN_VALUES.has(raw.valign)) style.valign = raw.valign;
  if (raw.wrap === true) style.wrap = true;
  if (typeof raw.format === "string" && raw.format.trim()) style.format = raw.format.trim().slice(0, MAX_FORMAT);

  const border = sanitizeBorder(raw.border);
  if (border) style.border = border;

  return Object.keys(style).length ? style : null;
}

function sanitizeBorder(raw) {
  if (!raw || typeof raw !== "object") return null;
  const out = {};
  for (const side of BORDER_SIDES) {
    const edge = raw[side];
    if (!edge || typeof edge !== "object") continue;
    if (!BORDER_STYLES.has(edge.style)) continue;
    const color = typeof edge.color === "string" && HEX.test(edge.color) ? edge.color.toUpperCase() : "#000000";
    out[side] = { style: edge.style, color };
  }
  return Object.keys(out).length ? out : null;
}

/** Sanitize a whole `styles` map keyed by "row,col". */
export function sanitizeStyles(raw) {
  if (!raw || typeof raw !== "object") return {};
  const out = {};
  let count = 0;
  for (const key of Object.keys(raw)) {
    if (count >= MAX_STYLE_ENTRIES) break;
    const match = /^(\d{1,5}),(\d{1,4})$/.exec(key);
    if (!match) continue;
    const style = sanitizeStyle(raw[key]);
    if (!style) continue;
    out[key] = style;
    count++;
  }
  return out;
}

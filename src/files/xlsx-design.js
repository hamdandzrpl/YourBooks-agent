// Coloured panels and native charts for agent-designed workbooks.
//
// ExcelJS can style cells but cannot author drawings, so "make this look like the
// user asked" stops at flat grids. This module patches the finished .xlsx: it
// adds text shapes (coloured panels, KPI boxes, notes) and native charts to any
// worksheet, merging into a drawing the sheet already owns instead of replacing
// it — a finance report ships with charts, and an agent edit must not wipe them.
//
// Same rules as src/finance/xlsx/charts.js: no new dependencies, node:zlib only.
import { readZip, writeZip } from "../finance/xlsx/zip.js";
import { chartXml, sheetPathByName } from "../finance/xlsx/charts.js";

const contentTypesPath = "[Content_Types].xml";
const drawingRelationType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const chartRelationType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const drawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
const chartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function attribute(tag, name) {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
}

function parseRelationships(xml) {
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b[^>]*\/>/g)].map((match) => ({
    id: attribute(match[0], "Id"),
    type: attribute(match[0], "Type"),
    target: attribute(match[0], "Target"),
  })).filter((entry) => entry.id && entry.type);
}

function relationshipXml(entries) {
  return `${xmlHeader}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + entries.map((entry) => `<Relationship Id="${entry.id}" Type="${entry.type}" Target="${entry.target}"/>`).join("")
    + "</Relationships>";
}

function resolveTarget(partPath, target) {
  if (target.startsWith("/")) return target.slice(1);
  const base = partPath.split("/").slice(0, -1);
  for (const segment of target.split("/")) {
    if (segment === "..") base.pop();
    else if (segment !== ".") base.push(segment);
  }
  return base.join("/");
}

function relsPathFor(partPath) {
  const slash = partPath.lastIndexOf("/");
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
}

function nextPartIndex(files, prefix) {
  let max = 0;
  for (const name of files.keys()) {
    const match = new RegExp(`^${prefix}(\\d+)\\.xml$`).exec(name);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return max + 1;
}

function nextRelId(entries) {
  return entries.reduce((max, entry) => {
    const number = Number.parseInt(String(entry.id).replace(/\D/g, ""), 10);
    return Number.isFinite(number) ? Math.max(max, number) : max;
  }, 0) + 1;
}

function emptyDrawing() {
  return `${xmlHeader}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" `
    + `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>`;
}

function anchorXml(from, to) {
  return `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>0</xdr:colOff>`
    + `<xdr:row>${from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`
    + `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>0</xdr:colOff>`
    + `<xdr:row>${to.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;
}

const alignMap = new Map([["left", "l"], ["center", "ctr"], ["centre", "ctr"], ["right", "r"], ["justify", "just"]]);
const anchorMap = new Map([["top", "t"], ["middle", "ctr"], ["center", "ctr"], ["bottom", "b"]]);

function shapeTextBody(shape) {
  const paragraphs = String(shape.text ?? "").split(/\r?\n/);
  const size = Math.round(Number(shape.fontSize ?? 11) * 100);
  const bold = shape.bold === false ? "" : ' b="1"';
  const color = String(shape.textColor ?? "FFFFFF").replace(/^#/, "").toUpperCase();
  const align = alignMap.get(String(shape.align ?? "left").toLowerCase()) ?? "l";
  const body = paragraphs.map((line) => (line
    ? `<a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="id-ID" sz="${size}"${bold}>`
      + `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr>`
      + `<a:t>${escapeXml(line)}</a:t></a:r></a:p>`
    : `<a:p><a:endParaRPr lang="id-ID" sz="${size}"/></a:p>`)).join("");
  return `<xdr:txBody>`
    + `<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" `
    + `lIns="91440" tIns="45720" rIns="91440" bIns="45720" `
    + `anchor="${anchorMap.get(String(shape.valign ?? "middle").toLowerCase()) ?? "ctr"}"/>`
    + `<a:lstStyle/>${body}</xdr:txBody>`;
}

function shapeXml(shape, id) {
  const fill = shape.fill ? String(shape.fill).replace(/^#/, "").toUpperCase() : null;
  const borderColor = shape.border ? String(shape.border).replace(/^#/, "").toUpperCase() : null;
  const geometry = shape.rounded === false ? "rect" : "roundRect";
  const fillXml = fill
    ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`
    : "<a:noFill/>";
  const lineXml = borderColor
    ? `<a:ln w="${Math.round(Number(shape.borderWidth ?? 1) * 12700)}"><a:solidFill><a:srgbClr val="${borderColor}"/></a:solidFill></a:ln>`
    : "<a:ln><a:noFill/></a:ln>";
  const shadowXml = shape.shadow
    ? `<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" algn="tl" rotWithShape="0">`
      + `<a:srgbClr val="0F172A"><a:alpha val="18000"/></a:srgbClr></a:outerShdw></a:effectLst>`
    : "";
  return `<xdr:twoCellAnchor editAs="oneCell">${anchorXml(shape.from, shape.to)}`
    + `<xdr:sp macro="" textlink="">`
    + `<xdr:nvSpPr><xdr:cNvPr id="${id}" name="Panel ${id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>`
    + `<xdr:spPr><a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>`
    + `${fillXml}${lineXml}${shadowXml}</xdr:spPr>`
    + `${shapeTextBody(shape)}</xdr:sp>`
    + `<xdr:clientData/></xdr:twoCellAnchor>`;
}

function chartAnchorXml(spec, relationshipId, objectId) {
  return `<xdr:twoCellAnchor editAs="oneCell">${anchorXml(spec.from, spec.to)}`
    + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>`
    + `<xdr:cNvPr id="${objectId}" name="Chart ${objectId}"/><xdr:cNvGraphicFramePr/>`
    + `</xdr:nvGraphicFramePr>`
    + `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">`
    + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" `
    + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relationshipId}"/>`
    + `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

function maxDrawingObjectId(drawingXml) {
  let max = 1;
  for (const match of drawingXml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)) {
    max = Math.max(max, Number(match[1]));
  }
  return max;
}

function ensureDrawingNamespaces(xml) {
  return xml.replace(/<xdr:wsDr\b([^>]*)>/, (full, attributes) => {
    let next = attributes;
    if (!/xmlns:xdr=/.test(next)) {
      next += ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"';
    }
    if (!/xmlns:a=/.test(next)) {
      next += ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    }
    return `<xdr:wsDr${next}>`;
  });
}

function normalizeAnchor(value, label) {
  const from = value?.from;
  const to = value?.to;
  const valid = (point) => point && Number.isInteger(point.col) && Number.isInteger(point.row)
    && point.col >= 0 && point.row >= 0;
  if (!valid(from) || !valid(to)) {
    throw new Error(`${label} needs from/to anchors as {col, row} zero-based integers`);
  }
  if (to.col <= from.col || to.row <= from.row) {
    throw new Error(`${label} anchor must end below and right of its start`);
  }
  return { from: { col: from.col, row: from.row }, to: { col: to.col, row: to.row } };
}

function normalizeChartSpec(spec) {
  const type = String(spec.type ?? "").toLowerCase();
  if (!["bar", "line", "doughnut"].includes(type)) {
    throw new Error(`Unsupported chart type "${spec.type}": use bar, line, or doughnut`);
  }
  if (!spec.sheet) throw new Error("Each chart needs a sheet name");
  const anchor = normalizeAnchor(spec, `Chart "${spec.title ?? type}"`);
  if (!spec.categories?.ref || !Array.isArray(spec.categories.values)) {
    throw new Error("Each chart needs categories {ref, values}");
  }
  if (!Array.isArray(spec.series) || spec.series.length === 0) {
    throw new Error("Each chart needs at least one series");
  }
  if (type === "doughnut" && spec.series.length !== 1) {
    throw new Error("A doughnut chart takes exactly one series");
  }
  return {
    ...spec,
    type,
    from: anchor.from,
    to: anchor.to,
    series: spec.series.map((entry) => ({
      name: entry.name ?? "Nilai",
      nameRef: entry.nameRef ?? null,
      ref: entry.ref,
      values: Array.isArray(entry.values) ? entry.values : [],
      color: String(entry.color ?? "0F766E").replace(/^#/, "").toUpperCase(),
    })),
  };
}

/**
 * Add shapes and native charts to an .xlsx buffer.
 *
 * @param {Buffer} bytes workbook bytes
 * @param {{shapes?: Array<object>, charts?: Array<object>}} design
 * @returns {{bytes: Buffer, shapes: number, charts: number}}
 */
export function injectDesign(bytes, { shapes = [], charts = [] } = {}) {
  if (!shapes.length && !charts.length) return { bytes, shapes: 0, charts: 0 };
  const files = readZip(bytes);
  const bySheet = new Map();
  for (const shape of shapes) {
    if (!shape?.sheet) throw new Error("Each shape needs a sheet name");
    if (!bySheet.has(shape.sheet)) bySheet.set(shape.sheet, { shapes: [], charts: [] });
    bySheet.get(shape.sheet).shapes.push({ ...shape, ...normalizeAnchor(shape, `Shape "${shape.text ?? ""}"`) });
  }
  for (const spec of charts) {
    const normalized = normalizeChartSpec(spec);
    if (!bySheet.has(normalized.sheet)) bySheet.set(normalized.sheet, { shapes: [], charts: [] });
    bySheet.get(normalized.sheet).charts.push(normalized);
  }

  const overrides = [];
  const sheetPatches = [];
  let chartCount = 0;
  let shapeCount = 0;

  for (const [sheetName, elements] of bySheet) {
    const sheetPath = sheetPathByName(files, sheetName);
    const sheetRelsPath = relsPathFor(sheetPath);
    const sheetRelationships = parseRelationships(files.get(sheetRelsPath)?.toString("utf8"));

    let drawingRelationship = sheetRelationships.find((entry) => entry.type === drawingRelationType);
    let drawingPath;
    let drawingXml;
    if (drawingRelationship) {
      drawingPath = resolveTarget(sheetPath, drawingRelationship.target);
      drawingXml = files.get(drawingPath)?.toString("utf8") ?? emptyDrawing();
    } else {
      const index = nextPartIndex(files, "xl/drawings/drawing");
      drawingPath = `xl/drawings/drawing${index}.xml`;
      drawingXml = emptyDrawing();
      drawingRelationship = {
        id: `rId${nextRelId(sheetRelationships)}`,
        type: drawingRelationType,
        target: `../drawings/${drawingPath.split("/").pop()}`,
      };
      sheetRelationships.push(drawingRelationship);
      files.set(sheetRelsPath, relationshipXml(sheetRelationships));
      overrides.push(`<Override PartName="/${drawingPath}" ContentType="${drawingContentType}"/>`);
    }

    const drawingRelsPath = relsPathFor(drawingPath);
    const drawingRelationships = parseRelationships(files.get(drawingRelsPath)?.toString("utf8"));
    let nextId = nextRelId(drawingRelationships);

    const anchors = [];
    const baseId = maxDrawingObjectId(drawingXml);
    for (const shape of elements.shapes) {
      shapeCount += 1;
      anchors.push(shapeXml(shape, baseId + anchors.length + 1));
    }
    for (const spec of elements.charts) {
      chartCount += 1;
      const chartIndex = nextPartIndex(files, "xl/charts/chart");
      const chartName = `chart${chartIndex}.xml`;
      files.set(`xl/charts/${chartName}`, chartXml(spec));
      overrides.push(`<Override PartName="/xl/charts/${chartName}" ContentType="${chartContentType}"/>`);
      const relationshipId = `rId${nextId}`;
      nextId += 1;
      drawingRelationships.push({
        id: relationshipId,
        type: chartRelationType,
        target: `../charts/${chartName}`,
      });
      anchors.push(chartAnchorXml(spec, relationshipId, baseId + anchors.length + 1));
    }

    drawingXml = ensureDrawingNamespaces(drawingXml).replace("</xdr:wsDr>", `${anchors.join("")}</xdr:wsDr>`);
    files.set(drawingPath, Buffer.from(drawingXml, "utf8"));
    if (drawingRelationships.length) {
      files.set(drawingRelsPath, Buffer.from(relationshipXml(drawingRelationships), "utf8"));
    }

    let sheetXml = files.get(sheetPath)?.toString("utf8");
    if (!sheetXml) throw new Error(`Worksheet part ${sheetPath} is missing`);
    if (!new RegExp(`<drawing\\b[^>]*r:id="${drawingRelationship.id}"`).test(sheetXml)) {
      const rootMatch = /<worksheet\b([^>]*)>/.exec(sheetXml);
      if (!rootMatch) throw new Error(`Worksheet part ${sheetPath} has no root element`);
      if (!/xmlns:r=/.test(rootMatch[1])) {
        sheetXml = sheetXml.replace(rootMatch[0], `<worksheet${rootMatch[1]} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`);
      }
      sheetXml = sheetXml.replace(/<drawing\b[^>]*\/>/g, "");
      sheetPatches.push({ sheetPath, sheetXml, relationshipId: drawingRelationship.id });
    }
  }

  for (const patch of sheetPatches) {
    let sheetXml = patch.sheetXml;
    const drawingTag = `<drawing r:id="${patch.relationshipId}"/>`;
    const candidates = ["legacyDrawing", "tableParts", "extLst"]
      .map((tag) => lastTopLevelTagStartSafe(sheetXml, tag))
      .filter((index) => index > -1);
    const insertAt = candidates.length ? Math.min(...candidates) : sheetXml.length - "</worksheet>".length;
    sheetXml = `${sheetXml.slice(0, insertAt)}${drawingTag}${sheetXml.slice(insertAt)}`;
    files.set(patch.sheetPath, Buffer.from(sheetXml, "utf8"));
  }

  if (overrides.length) {
    const contentTypes = files.get(contentTypesPath)?.toString("utf8");
    if (!contentTypes) throw new Error("Workbook is missing [Content_Types].xml");
    const missing = overrides.filter((entry) => !contentTypes.includes(entry));
    if (missing.length) {
      files.set(contentTypesPath, Buffer.from(contentTypes.replace("</Types>", `${missing.join("")}</Types>`), "utf8"));
    }
  }

  return { bytes: writeZip(files), shapes: shapeCount, charts: chartCount };
}

// Local copy so this module does not depend on chart internals beyond chartXml.
function lastTopLevelTagStartSafe(xml, tag) {
  const elementPattern = /<(\/?)([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
  let depth = 0;
  let found = -1;
  let match;
  while ((match = elementPattern.exec(xml)) !== null) {
    const [, closing, name, , selfClosing] = match;
    if (closing) {
      depth -= 1;
      continue;
    }
    if (depth === 1 && name === tag) found = match.index;
    if (!selfClosing) depth += 1;
  }
  return found;
}

/**
 * Describe the drawings a workbook already carries, per sheet.
 * Used by inspect so the agent knows what area of the sheet is already taken by
 * a chart or panel before it writes a new section on top of it.
 *
 * @returns {Map<string, {charts: number, shapes: number, images: number, anchors: string[]}>}
 */
export function describeDesign(bytes) {
  const result = new Map();
  let files;
  try {
    files = readZip(bytes);
  } catch {
    return result;
  }
  const workbook = files.get("xl/workbook.xml")?.toString("utf8");
  const workbookRels = files.get("xl/_rels/workbook.xml.rels")?.toString("utf8");
  if (!workbook || !workbookRels) return result;
  const targets = new Map(parseRelationships(workbookRels).map((entry) => [entry.id, entry.target]));
  for (const match of workbook.matchAll(/<sheet\b[^>]*\/>/g)) {
    const name = attribute(match[0], "name");
    const relId = attribute(match[0], "r:id");
    const target = targets.get(relId);
    if (!name || !target) continue;
    const sheetPath = resolveTarget("xl/workbook.xml", target);
    const sheetRelationships = parseRelationships(files.get(relsPathFor(sheetPath))?.toString("utf8"));
    const drawingRelationship = sheetRelationships.find((entry) => entry.type === drawingRelationType);
    if (!drawingRelationship) {
      result.set(name, { charts: 0, shapes: 0, images: 0, anchors: [] });
      continue;
    }
    const drawingPath = resolveTarget(sheetPath, drawingRelationship.target);
    const drawingXml = files.get(drawingPath)?.toString("utf8") ?? "";
    const drawingRelationships = parseRelationships(files.get(relsPathFor(drawingPath))?.toString("utf8"));
    result.set(name, {
      charts: drawingRelationships.filter((entry) => entry.type === chartRelationType).length,
      shapes: [...drawingXml.matchAll(/<xdr:sp\b/g)].length,
      images: [...drawingXml.matchAll(/<xdr:pic\b/g)].length,
      anchors: anchorAreas(drawingXml),
    });
  }
  return result;
}

function anchorAreas(drawingXml) {
  const areas = [];
  for (const match of drawingXml.matchAll(/<xdr:twoCellAnchor\b[^>]*>([\s\S]*?)<\/xdr:twoCellAnchor>/g)) {
    const block = match[1];
    const from = /<xdr:from>([\s\S]*?)<\/xdr:from>/.exec(block)?.[1] ?? "";
    const to = /<xdr:to>([\s\S]*?)<\/xdr:to>/.exec(block)?.[1] ?? "";
    const col = (xml) => Number(/<xdr:col>(\d+)<\/xdr:col>/.exec(xml)?.[1]);
    const row = (xml) => Number(/<xdr:row>(\d+)<\/xdr:row>/.exec(xml)?.[1]);
    if (![col(from), row(from), col(to), row(to)].every(Number.isInteger)) continue;
    const area = `${columnName(col(from) + 1)}${row(from) + 1}:${columnName(col(to) + 1)}${row(to) + 1}`;
    if (!areas.includes(area)) areas.push(area);
  }
  return areas.slice(0, 12);
}

function columnName(index) {
  let remaining = Number(index);
  let letters = "";
  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + modulo) + letters;
    remaining = Math.floor((remaining - modulo) / 26);
  }
  return letters;
}

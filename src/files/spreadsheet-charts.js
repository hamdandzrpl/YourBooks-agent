// Preserve native charts across ExcelJS round-trips.
//
// ExcelJS does not model charts or drawings. Loading a workbook with it and
// writing it back silently drops every chart, so editing one cell of a rendered
// finance report destroyed all of its charts. This module copies the chart,
// drawing, and relationship parts from the original archive into the rewritten
// one and re-attaches each drawing to its sheet.
import { readZip, writeZip } from "../finance/xlsx/zip.js";
import { lastTopLevelTagStart } from "../finance/xlsx/charts.js";

const drawingRelationType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const chartRelationType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const worksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const drawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
const chartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const relationshipPattern = /<Relationship\b[^>]*\/>/g;
const attribute = (tag, name) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];

function relationships(xml) {
  return [...xml.matchAll(relationshipPattern)].map((match) => ({
    tag: match[0],
    id: attribute(match[0], "Id"),
    type: attribute(match[0], "Type"),
    target: attribute(match[0], "Target"),
  }));
}

/** Resolve a relationship target relative to the part that declares it. */
function resolveTarget(partPath, target) {
  if (target.startsWith("/")) return target.slice(1);
  const base = partPath.split("/").slice(0, -1);
  for (const segment of target.split("/")) {
    if (segment === "..") base.pop();
    else if (segment !== ".") base.push(segment);
  }
  return base.join("/");
}

/** Map sheet name to its worksheet part path using workbook.xml + its rels. */
function sheetParts(files) {
  const workbook = files.get("xl/workbook.xml")?.toString("utf8");
  const rels = files.get("xl/_rels/workbook.xml.rels")?.toString("utf8");
  if (!workbook || !rels) return new Map();
  const byId = new Map(relationships(rels).map((rel) => [rel.id, rel.target]));
  const parts = new Map();
  for (const match of workbook.matchAll(/<sheet\b[^>]*\/>/g)) {
    const name = attribute(match[0], "name");
    const relId = attribute(match[0], "r:id");
    const target = byId.get(relId);
    if (name && target) parts.set(name, resolveTarget("xl/workbook.xml", target));
  }
  return parts;
}

function sheetRelsPath(sheetPath) {
  const slash = sheetPath.lastIndexOf("/");
  return `${sheetPath.slice(0, slash)}/_rels/${sheetPath.slice(slash + 1)}.rels`;
}

/**
 * Copy charts/drawings from `original` into `updated`.
 *
 * Returns `{ bytes, charts }` where `charts` is how many drawing parts were
 * re-attached. The updated bytes are returned untouched when the original had
 * no charts, so plain workbooks keep the cheap path.
 */
export function preserveCharts(originalBytes, updatedBytes) {
  const original = readZip(originalBytes);
  const updated = readZip(updatedBytes);

  const drawingParts = [...original.keys()].filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name));
  if (drawingParts.length === 0) return { bytes: updatedBytes, charts: 0 };

  const originalSheets = sheetParts(original);
  const updatedSheets = sheetParts(updated);
  let charts = 0;

  // Carry over every chart and drawing part, including their own relationships.
  for (const [name, data] of original) {
    if (/^xl\/(charts|drawings)\//.test(name) || /^xl\/drawings\/_rels\//.test(name)) {
      updated.set(name, data);
    }
  }

  for (const [sheetName, originalPath] of originalSheets) {
    const updatedPath = updatedSheets.get(sheetName);
    if (!updatedPath) continue; // sheet was removed by the edit
    const originalSheetXml = original.get(originalPath)?.toString("utf8") ?? "";
    const originalRelsXml = original.get(sheetRelsPath(originalPath))?.toString("utf8");
    if (!originalRelsXml) continue;
    const drawingRel = relationships(originalRelsXml).find((rel) => rel.type === drawingRelationType);
    if (!drawingRel) continue;

    // Point the copied drawing part at the sheet's new path for its charts.
    const drawingPath = resolveTarget(originalPath, drawingRel.target);
    const drawingRelsPath = sheetRelsPath(drawingPath);
    const drawingRelsXml = original.get(drawingRelsPath)?.toString("utf8");
    if (drawingRelsXml) updated.set(drawingRelsPath, Buffer.from(drawingRelsXml, "utf8"));

    // The rewritten sheet lost its <drawing> reference; put it back in the
    // schema-mandated position (before legacyDrawing/tableParts/extLst).
    let sheetXml = updated.get(updatedPath).toString("utf8");
    sheetXml = sheetXml.replace(/<drawing\b[^>]*\/>/g, "");
    const relId = "rIdCharts1";
    const relsPath = sheetRelsPath(updatedPath);
    const existingRels = updated.get(relsPath)?.toString("utf8")
      ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    if (!relationships(existingRels).some((rel) => rel.type === drawingRelationType)) {
      const patched = existingRels.replace("</Relationships>", `<Relationship Id="${relId}" Type="${drawingRelationType}" Target="../drawings/${drawingPath.split("/").pop()}"/></Relationships>`);
      updated.set(relsPath, Buffer.from(patched, "utf8"));
    }
    const usedRelId = relationships(updated.get(relsPath).toString("utf8")).find((rel) => rel.type === drawingRelationType)?.id ?? relId;
    const drawingTag = `<drawing r:id="${usedRelId}"/>`;
    // `<extLst>` also appears nested inside conditional-formatting rules, so the
    // anchor must be the top-level sibling or the drawing lands inside a cfRule
    // and Excel silently drops the chart.
    const insertAt = Math.min(
      ...[lastTopLevelTagStart(sheetXml, "legacyDrawing"), lastTopLevelTagStart(sheetXml, "tableParts"), lastTopLevelTagStart(sheetXml, "extLst")]
        .filter((index) => index > -1)
        .concat([sheetXml.length - "</worksheet>".length]),
    );
    sheetXml = `${sheetXml.slice(0, insertAt)}${drawingTag}${sheetXml.slice(insertAt)}`;
    updated.set(updatedPath, Buffer.from(sheetXml, "utf8"));
    if (originalSheetXml.includes("<drawing")) charts += 1;
  }

  // Content types for the copied parts. Every chart and drawing part needs its
  // own Override: ExcelJS drops them all on rewrite, and restoring only the first
  // one left chart2..N without a content type, which makes Excel report the file
  // as corrupt and repair it as empty.
  const contentTypesXml = updated.get("[Content_Types].xml")?.toString("utf8");
  if (contentTypesXml) {
    const declared = new Set(
      [...contentTypesXml.matchAll(/PartName="\/([^"]+)"/g)].map((match) => match[1]),
    );
    const additions = [];
    for (const name of updated.keys()) {
      const type = /^xl\/drawings\/drawing\d+\.xml$/.test(name)
        ? drawingContentType
        : /^xl\/charts\/chart\d+\.xml$/.test(name)
          ? chartContentType
          : /^xl\/worksheets\/sheet\d+\.xml$/.test(name)
            ? worksheetContentType
            : null;
      if (type && !declared.has(name)) additions.push(`<Override PartName="/${name}" ContentType="${type}"/>`);
    }
    if (additions.length) {
      updated.set(
        "[Content_Types].xml",
        Buffer.from(contentTypesXml.replace("</Types>", `${additions.join("")}</Types>`), "utf8"),
      );
    }
  }

  return { bytes: writeZip(updated), charts };
}

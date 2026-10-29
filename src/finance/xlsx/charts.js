// Native Excel charts for workbooks authored with ExcelJS.
//
// ExcelJS cannot write chart parts, so after the workbook is serialised we open
// the .xlsx ZIP, add the DrawingML chart parts, wire the relationships and put a
// <drawing> reference on the sheet. Everything is validated by reloading the
// result with ExcelJS in the caller; this module throws on anything unexpected.
import { readZip, writeZip } from "./zip.js";

const contentTypesPath = "[Content_Types].xml";
const workbookPath = "xl/workbook.xml";
const workbookRelsPath = "xl/_rels/workbook.xml.rels";

const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

/** Quote a sheet name for use inside a chart formula. */
export function sheetRef(sheetName, range) {
  return `'${String(sheetName).replaceAll("'", "''")}'!${range}`;
}

function numCache(values) {
  const points = values
    .map((value, index) => (value === null || value === undefined || Number.isNaN(Number(value))
      ? ""
      : `<c:pt idx="${index}"><c:v>${Number(value)}</c:v></c:pt>`))
    .join("");
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache>`;
}

function strCache(values) {
  const points = values
    .map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`)
    .join("");
  return `<c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache>`;
}

function solidFill(color) {
  return `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`;
}

function chartTitle(title) {
  if (!title) return "";
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1300" b="1"/></a:pPr>`
    + `<a:r><a:rPr lang="id-ID" sz="1300" b="1"/><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx>`
    + `<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
}

/** A literal series name is used when the workbook has no label cell to point at. */
function seriesNameXml(name, nameRef) {
  if (!nameRef) return `<c:tx><c:v>${escapeXml(name ?? `Seri`) }</c:v></c:tx>`;
  return `<c:tx><c:strRef><c:f>${escapeXml(nameRef)}</c:f>${strCache([name])}</c:strRef></c:tx>`;
}

function seriesXml({ index, name, nameRef, categories, values, color }) {
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>`
    + seriesNameXml(name, nameRef)
    + solidFill(color)
    + `<c:invertIfNegative val="0"/>`
    + `<c:cat><c:strRef><c:f>${escapeXml(categories.ref)}</c:f>${strCache(categories.values)}</c:strRef></c:cat>`
    + `<c:val><c:numRef><c:f>${escapeXml(values.ref)}</c:f>${numCache(values.values)}</c:numRef></c:val>`
    + `</c:ser>`;
}

function catAxis(axisId, crossId, { title } = {}) {
  return `<c:catAx><c:axId val="${axisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling>`
    + `<c:delete val="0"/><c:axPos val="b"/>`
    + (title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="id-ID" sz="900"/><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : "")
    + `<c:numFmt formatCode="General" sourceLinked="1"/>`
    + `<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`
    + `<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D6E0E0"/></a:solidFill></a:ln></c:spPr>`
    + `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>`
    + `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
}

function valAxis(axisId, crossId, formatCode) {
  return `<c:valAx><c:axId val="${axisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling>`
    + `<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines>`
    + `<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="EDF2F2"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>`
    + `<c:numFmt formatCode="${formatCode}" sourceLinked="0"/>`
    + `<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`
    + `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>`
    + `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>`
    + `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
}

function plotArea(chartBody, axes) {
  return `<c:plotArea><c:layout/>${chartBody}${axes}</c:plotArea>`;
}

function chartSpace(inner) {
  return `${xmlHeader}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" `
    + `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `
    + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + `<c:roundedCorners val="0"/><c:chart>${inner}</c:chart>`
    + `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>`
    + `<c:pageSetup/></c:printSettings></c:chartSpace>`;
}

function barChartXml(spec) {
  const axisA = 51_000_001;
  const axisB = 51_000_002;
  const series = spec.series.map((entry, index) => seriesXml({
    index,
    name: entry.name,
    nameRef: entry.nameRef,
    categories: spec.categories,
    values: entry,
    color: entry.color,
  })).join("");
  const body = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>`
    + series
    + `<c:gapWidth val="70"/><c:overlap val="-15"/>`
    + `<c:axId val="${axisA}"/><c:axId val="${axisB}"/></c:barChart>`;
  return chartSpace(
    chartTitle(spec.title)
    + plotArea(body, catAxis(axisA, axisB) + valAxis(axisB, axisA, spec.formatCode ?? "#,##0"))
    + `<c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:legend>`
    + `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>`,
  );
}

function lineChartXml(spec) {
  const axisA = 51_000_003;
  const axisB = 51_000_004;
  const series = spec.series.map((entry, index) => `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>`
    + seriesNameXml(entry.name, entry.nameRef)
    + `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${entry.color}"/></a:solidFill><a:round/></a:ln></c:spPr>`
    + `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${entry.color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>`
    + `<c:cat><c:strRef><c:f>${escapeXml(spec.categories.ref)}</c:f>${strCache(spec.categories.values)}</c:strRef></c:cat>`
    + `<c:val><c:numRef><c:f>${escapeXml(entry.ref)}</c:f>${numCache(entry.values)}</c:numRef></c:val>`
    + `<c:smooth val="0"/></c:ser>`).join("");
  const body = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}`
    + `<c:marker val="1"/><c:axId val="${axisA}"/><c:axId val="${axisB}"/></c:lineChart>`;
  return chartSpace(
    chartTitle(spec.title)
    + plotArea(body, catAxis(axisA, axisB) + valAxis(axisB, axisA, spec.formatCode ?? "#,##0"))
    + `<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>`
    + `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>`,
  );
}

function doughnutChartXml(spec) {
  const points = (spec.categories.values ?? [])
    .map((_, index) => (spec.palette[index % spec.palette.length]
      ? `<c:dPt><c:idx val="${index}"/><c:bubble3D val="0"/>${solidFill(spec.palette[index % spec.palette.length])}</c:dPt>`
      : ""))
    .join("");
  const series = `<c:ser><c:idx val="0"/><c:order val="0"/>`
    + `<c:tx><c:strRef><c:f>${escapeXml(spec.series[0].nameRef)}</c:f>${strCache([spec.series[0].name])}</c:strRef></c:tx>`
    + points
    + `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>`
    + `<c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>`
    + `<c:cat><c:strRef><c:f>${escapeXml(spec.categories.ref)}</c:f>${strCache(spec.categories.values)}</c:strRef></c:cat>`
    + `<c:val><c:numRef><c:f>${escapeXml(spec.series[0].ref)}</c:f>${numCache(spec.series[0].values)}</c:numRef></c:val>`
    + `</c:ser>`;
  const body = `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/><c:holeSize val="52"/></c:doughnutChart>`;
  return chartSpace(
    chartTitle(spec.title)
    + plotArea(body, "")
    + `<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>`
    + `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>`,
  );
}

export function chartXml(spec) {
  if (spec.type === "bar") return barChartXml(spec);
  if (spec.type === "line") return lineChartXml(spec);
  if (spec.type === "doughnut") return doughnutChartXml(spec);
  throw new Error(`Unsupported chart type: ${spec.type}`);
}

function drawingXml(anchors) {
  const frames = anchors.map((anchor, index) => {
    const { from, to } = anchor;
    return `<xdr:twoCellAnchor editAs="oneCell">`
      + `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`
      + `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`
      + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>`
      + `<xdr:cNvPr id="${index + 2}" name="Chart ${index + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
      + `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`
      + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">`
      + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" `
      + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${anchor.relationshipId}"/>`
      + `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  }).join("");
  return `${xmlHeader}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" `
    + `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${frames}</xdr:wsDr>`;
}

function relationshipXml(relationships) {
  return `${xmlHeader}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + relationships.map((entry) => `<Relationship Id="${entry.id}" Type="${entry.type}" Target="${entry.target}"/>`).join("")
    + `</Relationships>`;
}

/**
 * Index of the first *top-level* `<tag>` child of `<worksheet>`, or -1.
 *
 * `<extLst>` also appears nested inside conditional-formatting rules, so a plain
 * lastIndexOf (or a backward depth count, which cannot tell a complete nested
 * element from a top-level one) splices `<drawing>` into the middle of a cfRule
 * and Excel then silently drops the chart.
 */
export function lastTopLevelTagStart(xml, tag) {
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

/** Resolve a worksheet part path from its sheet name. */
export function sheetPathByName(files, sheetName) {  const workbook = files.get(workbookPath)?.toString("utf8") ?? "";
  const rels = files.get(workbookRelsPath)?.toString("utf8") ?? "";
  const sheetPattern = /<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"/g;
  let match;
  while ((match = sheetPattern.exec(workbook))) {
    if (match[1] !== sheetName) continue;
    const relationshipPattern = new RegExp(`<Relationship[^>]*Id="${match[2]}"[^>]*Target="([^"]*)"`);
    const target = relationshipPattern.exec(rels)?.[1];
    if (!target) break;
    const normalized = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
    return normalized.replaceAll("xl/xl/", "xl/");
  }
  throw new Error(`Worksheet ${sheetName} not found in workbook`);
}

/**
 * Add native charts to an existing .xlsx buffer.
 *
 * @param {Buffer} bytes source workbook
 * @param {Array<object>} specs chart specifications
 * @returns {Buffer} patched workbook
 */
export function injectCharts(bytes, specs) {
  if (!specs?.length) return bytes;
  const files = readZip(bytes);
  const bySheet = new Map();
  for (const spec of specs) {
    if (!bySheet.has(spec.sheet)) bySheet.set(spec.sheet, []);
    bySheet.get(spec.sheet).push(spec);
  }

  const overrides = [];
  const sheetPatches = [];
  let chartIndex = 0;
  let drawingIndex = 0;

  for (const [sheetName, sheetSpecs] of bySheet) {
    const sheetPath = sheetPathByName(files, sheetName);
    const sheetBase = sheetPath.split("/").at(-1);
    const relsPath = `xl/worksheets/_rels/${sheetBase}.rels`;
    const existingRels = files.get(relsPath)?.toString("utf8") ?? "";
    const relationships = [];
    const relPattern = /<Relationship[^>]*Id="([^"]*)"[^>]*Type="([^"]*)"[^>]*Target="([^"]*)"/g;
    let relMatch;
    while ((relMatch = relPattern.exec(existingRels))) {
      relationships.push({ id: relMatch[1], type: relMatch[2], target: relMatch[3] });
    }
    let nextId = relationships.reduce((max, entry) => {
      const number = Number.parseInt(entry.id.replace(/\D/g, ""), 10);
      return Number.isFinite(number) ? Math.max(max, number) : max;
    }, 0) + 1;

    drawingIndex += 1;
    const drawingName = `drawing${drawingIndex}.xml`;
    const drawingRels = [];
    const anchors = [];

    for (const spec of sheetSpecs) {
      chartIndex += 1;
      const chartName = `chart${chartIndex}.xml`;
      files.set(`xl/charts/${chartName}`, chartXml(spec));
      overrides.push(`<Override PartName="/xl/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);

      const chartRelId = `rId${chartIndex}`;
      drawingRels.push({
        id: chartRelId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        target: `../charts/${chartName}`,
      });
      anchors.push({ from: spec.from, to: spec.to, relationshipId: chartRelId });
    }

    files.set(`xl/drawings/${drawingName}`, drawingXml(anchors));
    files.set(`xl/drawings/_rels/${drawingName}.rels`, relationshipXml(drawingRels));
    overrides.push(`<Override PartName="/xl/drawings/${drawingName}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);

    const drawingRelId = `rId${nextId}`;
    relationships.push({
      id: drawingRelId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
      target: `../drawings/${drawingName}`,
    });
    files.set(relsPath, relationshipXml(relationships));
    sheetPatches.push({ sheetPath, drawingRelId });
  }

  for (const patch of sheetPatches) {
    let sheetXml = files.get(patch.sheetPath)?.toString("utf8");
    if (!sheetXml) throw new Error(`Worksheet part ${patch.sheetPath} is missing`);
    const rootMatch = /<worksheet\b([^>]*)>/.exec(sheetXml);
    if (!rootMatch) throw new Error(`Worksheet part ${patch.sheetPath} has no root element`);
    if (!/xmlns:r=/.test(rootMatch[1])) {
      sheetXml = sheetXml.replace(rootMatch[0], `<worksheet${rootMatch[1]} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`);
    }
    sheetXml = sheetXml.replace(/<drawing[^>]*\/>/g, "");
    // CT_Worksheet order is strict: <drawing> must sit before legacyDrawing,
    // tableParts and extLst. Appending it just before </worksheet> makes Excel
    // report the file as corrupt and recover it as empty.
    const drawingTag = `<drawing r:id="${patch.drawingRelId}"/>`;
    const tableIndex = lastTopLevelTagStart(sheetXml, "tableParts");
    const extIndex = lastTopLevelTagStart(sheetXml, "extLst");
    const candidates = [tableIndex, extIndex].filter((index) => index > -1);
    const insertAt = candidates.length ? Math.min(...candidates) : sheetXml.length - "</worksheet>".length;
    sheetXml = `${sheetXml.slice(0, insertAt)}${drawingTag}${sheetXml.slice(insertAt)}`;
    files.set(patch.sheetPath, sheetXml);
  }

  const contentTypes = files.get(contentTypesPath)?.toString("utf8");
  if (!contentTypes) throw new Error("Workbook is missing [Content_Types].xml");
  files.set(contentTypesPath, contentTypes.replace("</Types>", `${overrides.join("")}</Types>`));

  return writeZip(files);
}

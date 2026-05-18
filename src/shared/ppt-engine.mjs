import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  relAttr: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  pkgRel: "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypes: "http://schemas.openxmlformats.org/package/2006/content-types",
  app: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
};

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const DEFAULT_MIN_FONT_SIZE = 6000;
const FONT_STEP = 200;
const EMU_PER_POINT = 12700;
const FULL_WIDTH_FACTOR = 0.92;
const LINE_HEIGHT_FACTOR = 1.15;
const OVERLAY_TOLERANCE = 2;
const CJK_CHAR_UNITS = 1.1;
const LONG_CJK_LABEL_UNITS = 1.3;
const DEFAULT_FONT_SIZE = 16600;

export async function generatePresentationBuffer({ templateBytes, lines, accentColor = null }) {
  const normalizedLines = normalizeLines(lines);
  if (normalizedLines.length === 0) {
    throw new Error("input does not contain any lines");
  }

  if (accentColor) {
    assertAccentColor(accentColor);
  }

  const zip = await JSZip.loadAsync(templateBytes);
  const slideSources = await loadSlideSources(zip, [1, 2, 3]);
  const slide1Doc = parseXml(slideSources.get(1).xml);
  const baseLayout = getLayoutMetrics(slide1Doc);
  const standardShapeLayout = captureStandardShapeLayout(slide1Doc);
  const pages = planPages(normalizedLines, baseLayout);
  const summary = summarizePlan(baseLayout, normalizedLines, pages);

  removeExistingSlides(zip);
  await writeSlides(zip, slideSources, pages, accentColor, standardShapeLayout);
  await rewritePresentationManifest(zip, pages.length);
  await rewriteContentTypes(zip, pages.length);
  await rewriteAppProperties(zip, pages.length);

  const outputBuffer = await zip.generateAsync({ type: "uint8array" });
  return {
    outputBuffer,
    summary
  };
}

export async function analyzeTemplateInput({ templateBytes, lines }) {
  const normalizedLines = normalizeLines(lines);
  if (normalizedLines.length === 0) {
    return {
      lineCount: 0,
      pageCount: 0,
      didShrink: false,
      didSplit: false,
      splitCount: 0
    };
  }

  const zip = await JSZip.loadAsync(templateBytes);
  const slideSources = await loadSlideSources(zip, [1, 2, 3]);
  const baseLayout = getLayoutMetrics(parseXml(slideSources.get(1).xml));
  const pages = planPages(normalizedLines, baseLayout);
  return summarizePlan(baseLayout, normalizedLines, pages);
}

export function parseInputLines(inputText) {
  const lines = inputText.replace(/\r\n/g, "\n").split("\n");
  if (lines[lines.length - 1] === "") {
    lines.pop();
  }
  return lines;
}

export function normalizeLines(lines) {
  return Array.isArray(lines) ? [...lines] : [];
}

export function assertAccentColor(accentColor) {
  if (!/^[0-9A-Fa-f]{6}$/.test(accentColor)) {
    throw new Error("accentColor must be a 6-digit hex value");
  }
}

function summarizePlan(layout, lines, pages) {
  return {
    lineCount: lines.length,
    pageCount: pages.length,
    didShrink: pages.some((page) => page.fontSize < layout.defaultFontSize),
    didSplit: pages.length > lines.length,
    splitCount: Math.max(0, pages.length - lines.length)
  };
}

async function loadSlideSources(zip, indices) {
  const entries = await Promise.all(
    indices.map(async (index) => [
      index,
      {
        xml: await zip.file(`ppt/slides/slide${index}.xml`).async("string"),
        rels: await zip.file(`ppt/slides/_rels/slide${index}.xml.rels`).async("string")
      }
    ])
  );
  return new Map(entries);
}

function getLayoutMetrics(slideDoc) {
  const editableShapes = findEditableShapeGroup(slideDoc);
  if (editableShapes.length !== 3) {
    throw new Error(`expected 3 editable text shapes, found ${editableShapes.length}`);
  }

  const sampleShape = editableShapes[0];
  return {
    widthEmu: Number(sampleShape.ext.getAttribute("cx")),
    heightEmu: Number(sampleShape.ext.getAttribute("cy")),
    defaultFontSize: Number(sampleShape.runProps.getAttribute("sz")) || DEFAULT_FONT_SIZE,
    minFontSize: DEFAULT_MIN_FONT_SIZE
  };
}

function planPages(lines, layout) {
  const pages = [];

  for (const line of lines) {
    const bestFit = findBestFontSize(line, layout);
    if (bestFit !== null) {
      pages.push({ text: line, fontSize: bestFit });
      continue;
    }

    let remaining = line;
    const minMetrics = getFitMetrics(layout, layout.minFontSize);
    while (remaining.length > 0) {
      const { chunk, rest } = sliceTextForCapacity(remaining, minMetrics.unitsPerLine, minMetrics.maxLines);
      pages.push({ text: chunk, fontSize: layout.minFontSize });
      remaining = rest;
    }
  }

  return pages;
}

function findBestFontSize(text, layout) {
  for (let fontSize = layout.defaultFontSize; fontSize >= layout.minFontSize; fontSize -= FONT_STEP) {
    if (fitsOnSlide(text, layout, fontSize)) {
      return fontSize;
    }
  }
  return null;
}

function fitsOnSlide(text, layout, fontSize) {
  const metrics = getFitMetrics(layout, fontSize);
  return countWrappedLines(text, metrics.unitsPerLine) <= metrics.maxLines;
}

function getFitMetrics(layout, fontSize) {
  const fontPoints = fontSize / 100;
  const unitsPerLine = layout.widthEmu / (fontPoints * EMU_PER_POINT * FULL_WIDTH_FACTOR);
  const maxLines = Math.max(
    1,
    Math.floor(layout.heightEmu / (fontPoints * EMU_PER_POINT * LINE_HEIGHT_FACTOR))
  );

  return { unitsPerLine, maxLines };
}

function countWrappedLines(text, unitsPerLine) {
  return text.split(/\n/).reduce((total, line) => total + Math.max(1, Math.ceil(measureTextUnits(line) / unitsPerLine)), 0);
}

function measureTextUnits(text) {
  const chars = Array.from(text);
  if (isDenseCjkLabel(chars)) {
    return chars.length * LONG_CJK_LABEL_UNITS;
  }

  return chars.reduce((sum, char) => sum + getCharUnits(char), 0);
}

function getCharUnits(char) {
  if (char === " ") {
    return 0.35;
  }

  if (/[\u4E00-\u9FFF\u3040-\u30FF\uAC00-\uD7AF]/u.test(char)) {
    return CJK_CHAR_UNITS;
  }

  if (/[A-Z0-9]/.test(char)) {
    return 0.7;
  }

  if (/[a-z]/.test(char)) {
    return 0.6;
  }

  if (/[，。、“”‘’！？：；,.!?;:()（）【】《》\-_/]/u.test(char)) {
    return 0.5;
  }

  return 0.8;
}

function isDenseCjkLabel(chars) {
  return chars.length >= 5 && chars.every((char) => /[\u4E00-\u9FFF]/u.test(char));
}

function sliceTextForCapacity(text, unitsPerLine, maxLines) {
  const characters = Array.from(text);
  let index = 0;
  let currentLineUnits = 0;
  let linesUsed = 1;

  while (index < characters.length) {
    const char = characters[index];
    if (char === "\n") {
      if (linesUsed === maxLines) {
        break;
      }
      linesUsed += 1;
      currentLineUnits = 0;
      index += 1;
      continue;
    }

    const charUnits = getCharUnits(char);
    if (currentLineUnits > 0 && currentLineUnits + charUnits > unitsPerLine) {
      if (linesUsed === maxLines) {
        break;
      }
      linesUsed += 1;
      currentLineUnits = 0;
      continue;
    }

    currentLineUnits += charUnits;
    index += 1;
  }

  const safeIndex = Math.max(index, 1);
  return {
    chunk: characters.slice(0, safeIndex).join(""),
    rest: characters.slice(safeIndex).join("")
  };
}

function removeExistingSlides(zip) {
  for (const fileName of Object.keys(zip.files)) {
    if (/^ppt\/slides\/slide\d+\.xml$/.test(fileName) || /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(fileName)) {
      zip.remove(fileName);
    }
  }
}

async function writeSlides(zip, slideSources, pages, accentColor, standardShapeLayout) {
  for (let index = 0; index < pages.length; index += 1) {
    const slideNumber = index + 1;
    const sourceIndex = slideNumber <= 3 ? slideNumber : 1;
    const source = slideSources.get(sourceIndex);
    const slideDoc = parseXml(source.xml);

    applyPageContent(slideDoc, pages[index], accentColor, standardShapeLayout);

    zip.file(`ppt/slides/slide${slideNumber}.xml`, serializeXml(slideDoc));
    zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, source.rels);
  }
}

function applyPageContent(slideDoc, page, accentColor, standardShapeLayout) {
  const editableShapes = findEditableShapeGroup(slideDoc);
  const accentShape = editableShapes.find(isAccentOutlineShape);

  for (const [index, shape] of editableShapes.entries()) {
    replaceShapeText(shape.shapeNode, page.text);
    updateShapeFontSize(shape.shapeNode, page.fontSize);
    if (standardShapeLayout[index]) {
      applyStandardShapeLayout(shape.shapeNode, standardShapeLayout[index]);
    }
  }

  if (accentColor && accentShape) {
    updateAccentColor(accentShape.shapeNode, accentColor);
  }
}

function captureStandardShapeLayout(slideDoc) {
  return findEditableShapeGroup(slideDoc).map((shape) => ({
    x: shape.off.getAttribute("x"),
    y: shape.off.getAttribute("y"),
    cx: shape.ext.getAttribute("cx"),
    cy: shape.ext.getAttribute("cy")
  }));
}

function applyStandardShapeLayout(shapeNode, layout) {
  const xfrm = shapeNode.getElementsByTagNameNS(NS.a, "xfrm")[0];
  const off = xfrm && xfrm.getElementsByTagNameNS(NS.a, "off")[0];
  const ext = xfrm && xfrm.getElementsByTagNameNS(NS.a, "ext")[0];

  if (off && ext) {
    off.setAttribute("x", layout.x);
    off.setAttribute("y", layout.y);
    ext.setAttribute("cx", layout.cx);
    ext.setAttribute("cy", layout.cy);
  }

  const bodyPr = shapeNode.getElementsByTagNameNS(NS.a, "bodyPr")[0];
  if (bodyPr) {
    bodyPr.setAttribute("anchor", "ctr");
  }

  const pPr = shapeNode.getElementsByTagNameNS(NS.a, "pPr")[0];
  if (pPr) {
    pPr.setAttribute("algn", "ctr");
  }
}

function findEditableShapeGroup(slideDoc) {
  const shapes = Array.from(slideDoc.getElementsByTagNameNS(NS.p, "sp"))
    .map((shapeNode) => describeShape(shapeNode))
    .filter(Boolean);

  let bestGroup = [];
  for (const shape of shapes) {
    const group = shapes.filter((candidate) => sameOverlayBox(shape, candidate));
    if (group.length > bestGroup.length) {
      bestGroup = group;
    }
  }

  return bestGroup.sort((left, right) => Number(right.lineWidth || 0) - Number(left.lineWidth || 0));
}

function describeShape(shapeNode) {
  const textNodes = shapeNode.getElementsByTagNameNS(NS.a, "t");
  if (textNodes.length === 0) {
    return null;
  }

  const xfrm = shapeNode.getElementsByTagNameNS(NS.a, "xfrm")[0];
  const off = xfrm && xfrm.getElementsByTagNameNS(NS.a, "off")[0];
  const ext = xfrm && xfrm.getElementsByTagNameNS(NS.a, "ext")[0];
  const runProps = shapeNode.getElementsByTagNameNS(NS.a, "rPr")[0];
  const line = runProps && runProps.getElementsByTagNameNS(NS.a, "ln")[0];
  const cNvSpPr = shapeNode.getElementsByTagNameNS(NS.p, "cNvSpPr")[0];

  if (!off || !ext || !runProps || !cNvSpPr || cNvSpPr.getAttribute("txBox") !== "1") {
    return null;
  }

  return {
    shapeNode,
    off,
    ext,
    runProps,
    lineWidth: line ? line.getAttribute("w") : null,
    lineColor: readColorValue(line)
  };
}

function isAccentOutlineShape(shape) {
  return Number(shape.lineWidth) > 0 && shape.lineColor !== "bg1";
}

function sameOverlayBox(left, right) {
  return (
    Math.abs(Number(left.off.getAttribute("x")) - Number(right.off.getAttribute("x"))) <= OVERLAY_TOLERANCE &&
    Math.abs(Number(left.off.getAttribute("y")) - Number(right.off.getAttribute("y"))) <= OVERLAY_TOLERANCE &&
    left.ext.getAttribute("cx") === right.ext.getAttribute("cx") &&
    left.ext.getAttribute("cy") === right.ext.getAttribute("cy")
  );
}

function replaceShapeText(shapeNode, text) {
  const textNodes = Array.from(shapeNode.getElementsByTagNameNS(NS.a, "t"));
  textNodes.forEach((node, index) => {
    node.textContent = index === 0 ? text : "";
  });
}

function updateShapeFontSize(shapeNode, fontSize) {
  const sizeValue = String(fontSize);
  for (const tagName of ["rPr", "endParaRPr"]) {
    const nodes = Array.from(shapeNode.getElementsByTagNameNS(NS.a, tagName));
    for (const node of nodes) {
      node.setAttribute("sz", sizeValue);
    }
  }
}

function updateAccentColor(shapeNode, accentColor) {
  const runProps = shapeNode.getElementsByTagNameNS(NS.a, "rPr")[0];
  if (!runProps) {
    return;
  }

  let line = runProps.getElementsByTagNameNS(NS.a, "ln")[0];
  if (!line) {
    line = shapeNode.ownerDocument.createElementNS(NS.a, "a:ln");
    line.setAttribute("w", "762000");
    runProps.insertBefore(line, runProps.firstChild);
  }

  while (line.firstChild) {
    line.removeChild(line.firstChild);
  }

  const solidFill = shapeNode.ownerDocument.createElementNS(NS.a, "a:solidFill");
  const srgbColor = shapeNode.ownerDocument.createElementNS(NS.a, "a:srgbClr");
  srgbColor.setAttribute("val", accentColor.toUpperCase());
  solidFill.appendChild(srgbColor);
  line.appendChild(solidFill);
}

function readColorValue(node) {
  if (!node) {
    return null;
  }

  const srgb = node.getElementsByTagNameNS(NS.a, "srgbClr")[0];
  if (srgb) {
    return srgb.getAttribute("val");
  }

  const scheme = node.getElementsByTagNameNS(NS.a, "schemeClr")[0];
  if (scheme) {
    return scheme.getAttribute("val");
  }

  return null;
}

async function rewritePresentationManifest(zip, slideCount) {
  const presentationText = await zip.file("ppt/presentation.xml").async("string");
  const relsText = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  const presentationDoc = parseXml(presentationText);
  const relsDoc = parseXml(relsText);

  rewriteSlideIdList(presentationDoc, slideCount);
  rewritePresentationRelationships(relsDoc, slideCount);

  zip.file("ppt/presentation.xml", serializeXml(presentationDoc));
  zip.file("ppt/_rels/presentation.xml.rels", serializeXml(relsDoc));
}

function rewriteSlideIdList(presentationDoc, slideCount) {
  const sldIdList = presentationDoc.getElementsByTagNameNS(NS.p, "sldIdLst")[0];
  while (sldIdList.firstChild) {
    sldIdList.removeChild(sldIdList.firstChild);
  }

  for (let index = 0; index < slideCount; index += 1) {
    const slideId = presentationDoc.createElementNS(NS.p, "p:sldId");
    slideId.setAttribute("id", String(256 + index));
    slideId.setAttributeNS(NS.relAttr, "r:id", `rId${9 + index}`);
    sldIdList.appendChild(slideId);
  }
}

function rewritePresentationRelationships(relsDoc, slideCount) {
  const relationships = relsDoc.documentElement;
  const existing = Array.from(relationships.getElementsByTagNameNS(NS.pkgRel, "Relationship"));

  for (const relationship of existing) {
    if (relationship.getAttribute("Type") === SLIDE_REL_TYPE) {
      relationships.removeChild(relationship);
    }
  }

  for (let index = 0; index < slideCount; index += 1) {
    const relationship = relsDoc.createElementNS(NS.pkgRel, "Relationship");
    relationship.setAttribute("Id", `rId${9 + index}`);
    relationship.setAttribute("Type", SLIDE_REL_TYPE);
    relationship.setAttribute("Target", `slides/slide${index + 1}.xml`);
    relationships.appendChild(relationship);
  }
}

async function rewriteContentTypes(zip, slideCount) {
  const contentTypesText = await zip.file("[Content_Types].xml").async("string");
  const contentTypesDoc = parseXml(contentTypesText);
  const root = contentTypesDoc.documentElement;
  const overrides = Array.from(root.getElementsByTagNameNS(NS.contentTypes, "Override"));

  for (const override of overrides) {
    if (/^\/ppt\/slides\/slide\d+\.xml$/.test(override.getAttribute("PartName"))) {
      root.removeChild(override);
    }
  }

  for (let index = 0; index < slideCount; index += 1) {
    const override = contentTypesDoc.createElementNS(NS.contentTypes, "Override");
    override.setAttribute("PartName", `/ppt/slides/slide${index + 1}.xml`);
    override.setAttribute("ContentType", SLIDE_CONTENT_TYPE);
    root.appendChild(override);
  }

  zip.file("[Content_Types].xml", serializeXml(contentTypesDoc));
}

async function rewriteAppProperties(zip, slideCount) {
  const appText = await zip.file("docProps/app.xml").async("string");
  const appDoc = parseXml(appText);
  const previousSlideCount = readAppSlideTitleCount(appDoc);
  const slidesNode = appDoc.getElementsByTagNameNS(NS.app, "Slides")[0];
  if (slidesNode) {
    slidesNode.textContent = String(slideCount);
  }
  rewriteHeadingPairs(appDoc, slideCount);
  rewriteTitlesOfParts(appDoc, previousSlideCount, slideCount);
  zip.file("docProps/app.xml", serializeXml(appDoc));
}

function readAppSlideTitleCount(appDoc) {
  const headingPairs = appDoc.getElementsByTagNameNS(NS.app, "HeadingPairs")[0];
  if (!headingPairs) {
    return 0;
  }

  const vector = headingPairs.getElementsByTagNameNS(NS.vt, "vector")[0];
  const variants = Array.from(vector.getElementsByTagNameNS(NS.vt, "variant"));

  for (let index = 0; index < variants.length - 1; index += 2) {
    const label = firstChildElement(variants[index]);
    const count = firstChildElement(variants[index + 1]);
    if (!label || !count) {
      continue;
    }
    if (/(幻灯片标题|Slide Titles)/i.test(label.textContent)) {
      return Number(count.textContent) || 0;
    }
  }

  return 0;
}

function rewriteHeadingPairs(appDoc, slideCount) {
  const headingPairs = appDoc.getElementsByTagNameNS(NS.app, "HeadingPairs")[0];
  if (!headingPairs) {
    return;
  }

  const vector = headingPairs.getElementsByTagNameNS(NS.vt, "vector")[0];
  const variants = Array.from(vector.getElementsByTagNameNS(NS.vt, "variant"));

  for (let index = 0; index < variants.length - 1; index += 2) {
    const label = firstChildElement(variants[index]);
    const count = firstChildElement(variants[index + 1]);
    if (!label || !count) {
      continue;
    }
    if (/(幻灯片标题|Slide Titles)/i.test(label.textContent)) {
      count.textContent = String(slideCount);
      return;
    }
  }
}

function rewriteTitlesOfParts(appDoc, previousSlideCount, slideCount) {
  const titlesOfParts = appDoc.getElementsByTagNameNS(NS.app, "TitlesOfParts")[0];
  if (!titlesOfParts) {
    return;
  }

  const vector = titlesOfParts.getElementsByTagNameNS(NS.vt, "vector")[0];
  const entries = directChildElements(vector, NS.vt, "lpstr");
  const staticEntryCount = Math.max(entries.length - previousSlideCount, 0);
  const existingSlideEntries = entries.slice(staticEntryCount);
  const fallbackTitle = existingSlideEntries[0]?.textContent || "PowerPoint 演示文稿";

  for (let index = entries.length - 1; index >= staticEntryCount; index -= 1) {
    vector.removeChild(entries[index]);
  }

  for (let index = 0; index < slideCount; index += 1) {
    const entry = appDoc.createElementNS(NS.vt, "vt:lpstr");
    entry.textContent = existingSlideEntries[index]?.textContent || fallbackTitle;
    vector.appendChild(entry);
  }

  vector.setAttribute("size", String(staticEntryCount + slideCount));
}

function directChildElements(node, namespaceUri, localName) {
  return Array.from(node.childNodes).filter(
    (child) => child.nodeType === 1 && child.namespaceURI === namespaceUri && child.localName === localName
  );
}

function firstChildElement(node) {
  return Array.from(node.childNodes).find((child) => child.nodeType === 1) || null;
}

function parseXml(xmlText) {
  return new DOMParser().parseFromString(xmlText, "application/xml");
}

function serializeXml(doc) {
  const xmlBody = new XMLSerializer().serializeToString(doc).replace(/^\s*<\?xml[^>]*>\s*/i, "");
  return `${XML_HEADER}${xmlBody}`;
}

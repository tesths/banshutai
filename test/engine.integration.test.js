const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const test = require("node:test");
const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");

const TEMPLATE_PATH = path.join(__dirname, "..", "public", "default-template.pptx");
const ENGINE_PATH = pathToFileURL(path.join(__dirname, "..", "src", "shared", "ppt-engine.mjs")).href;
const XML_NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  app: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
};

test("generates one updated slide for one input item and keeps all three layers in sync", async () => {
  const zip = await generatePresentationZip({
    lines: ["统一文案测试"]
  });
  const slidePaths = await getOrderedSlidePaths(zip);

  assert.equal(slidePaths.length, 1);

  const slideDoc = await loadXml(zip, slidePaths[0]);
  const textShapes = getTextShapes(slideDoc);
  assert.equal(textShapes.length, 3);
  assert.deepEqual(
    textShapes.map((shape) => shape.text),
    ["统一文案测试", "统一文案测试", "统一文案测试"]
  );
  assert.equal(new Set(textShapes.map((shape) => shape.extKey)).size, 1);
});

test("creates additional slides beyond the template and reuses slide 1 layout for slide 4+", async () => {
  const zip = await generatePresentationZip({
    lines: ["第一页", "第二页", "第三页", "第四页"]
  });
  const slidePaths = await getOrderedSlidePaths(zip);

  assert.equal(slidePaths.length, 4);

  const slide1Doc = await loadXml(zip, slidePaths[0]);
  const slide4Doc = await loadXml(zip, slidePaths[3]);
  const slide1Shapes = getTextShapes(slide1Doc);
  const slide4Shapes = getTextShapes(slide4Doc);

  assert.deepEqual(
    slide4Shapes.map((shape) => shape.text),
    ["第四页", "第四页", "第四页"]
  );
  assert.deepEqual(
    slide4Shapes.map((shape) => shape.layoutKey),
    slide1Shapes.map((shape) => shape.layoutKey)
  );
});

test("keeps the third slide text boxes in the standard centered region", async () => {
  const zip = await generatePresentationZip({
    lines: ["第一页", "第二页", "第三页"]
  });
  const slidePaths = await getOrderedSlidePaths(zip);

  const slide1Shapes = getTextShapes(await loadXml(zip, slidePaths[0]));
  const slide3Shapes = getTextShapes(await loadXml(zip, slidePaths[2]));

  assert.equal(slide3Shapes.length, 3);
  assert.deepEqual(
    slide3Shapes.map((shape) => shape.text),
    ["第三页", "第三页", "第三页"]
  );
  assert.deepEqual(
    slide3Shapes.map((shape) => shape.layoutKey),
    slide1Shapes.map((shape) => shape.layoutKey)
  );
  assert.deepEqual(
    slide3Shapes.map((shape) => shape.bodyAnchor),
    ["ctr", "ctr", "ctr"]
  );
  assert.deepEqual(
    slide3Shapes.map((shape) => shape.paragraphAlign),
    ["ctr", "ctr", "ctr"]
  );
});

test("applies the provided accent color only to the blue outline layer", async () => {
  const zip = await generatePresentationZip({
    lines: ["颜色测试"],
    accent: "FF6600"
  });
  const slidePaths = await getOrderedSlidePaths(zip);
  const slideDoc = await loadXml(zip, slidePaths[0]);
  const textShapes = getTextShapes(slideDoc);
  const blueShape = textShapes.find((shape) => shape.lineWidth === "762000");
  const whiteShape = textShapes.find((shape) => shape.lineWidth === "381000");

  assert.ok(blueShape, "expected to find the blue accent layer");
  assert.ok(whiteShape, "expected to find the white outline layer");
  assert.equal(blueShape.lineColor, "FF6600");
  assert.equal(whiteShape.lineColor, "bg1");
});

test("splits very long text across multiple slides and preserves full content order", async () => {
  const longText = "这是一个需要自动拆分页的超长标题".repeat(18);
  const zip = await generatePresentationZip({
    lines: [longText]
  });
  const slidePaths = await getOrderedSlidePaths(zip);

  assert.ok(slidePaths.length > 1, "expected overflow text to produce additional slides");

  const chunks = [];
  for (const slidePath of slidePaths) {
    const slideDoc = await loadXml(zip, slidePath);
    const textShapes = getTextShapes(slideDoc);
    assert.equal(textShapes.length, 3);
    assert.equal(textShapes[0].text, textShapes[1].text);
    assert.equal(textShapes[1].text, textShapes[2].text);
    chunks.push(textShapes[0].text);
  }

  assert.equal(chunks.join(""), longText);
});

test("shrinks font size for longer text before it needs to split into another slide", async () => {
  const zip = await generatePresentationZip({
    lines: ["这是一个需要缩小字号但还不需要分页的稍长标题文案"]
  });
  const slidePaths = await getOrderedSlidePaths(zip);
  const slideDoc = await loadXml(zip, slidePaths[0]);
  const textShapes = getTextShapes(slideDoc);

  assert.equal(slidePaths.length, 1);
  assert.ok(
    textShapes.every((shape) => Number(shape.fontSize) < 16600),
    "expected all text layers to use a smaller font size than the template default"
  );
});

test("preserves blank lines and surrounding spaces as exact page content", async () => {
  const zip = await generatePresentationZip({
    lines: ["  前后空格  ", "", "第三页"]
  });
  const slidePaths = await getOrderedSlidePaths(zip);

  assert.equal(slidePaths.length, 3);

  const slide1 = getTextShapes(await loadXml(zip, slidePaths[0]));
  const slide2 = getTextShapes(await loadXml(zip, slidePaths[1]));
  const slide3 = getTextShapes(await loadXml(zip, slidePaths[2]));

  assert.deepEqual(slide1.map((shape) => shape.text), ["  前后空格  ", "  前后空格  ", "  前后空格  "]);
  assert.deepEqual(slide2.map((shape) => shape.text), ["", "", ""]);
  assert.deepEqual(slide3.map((shape) => shape.text), ["第三页", "第三页", "第三页"]);
});

test("shrinks five-character CJK labels enough to stay on a single line", async () => {
  const zip = await generatePresentationZip({
    lines: ["纠结拖延兽"]
  });
  const slideDoc = await loadXml(zip, (await getOrderedSlidePaths(zip))[0]);
  const textShapes = getTextShapes(slideDoc);

  assert.ok(
    textShapes.every((shape) => Number(shape.fontSize) <= 12400),
    "expected five-character CJK labels to keep a safety margin below the wrap-prone size"
  );
});

test("shrinks six-character CJK labels with extra headroom for Safari", async () => {
  const zip = await generatePresentationZip({
    lines: ["春江花月夜里"]
  });
  const slideDoc = await loadXml(zip, (await getOrderedSlidePaths(zip))[0]);
  const textShapes = getTextShapes(slideDoc);

  assert.ok(
    textShapes.every((shape) => Number(shape.fontSize) <= 10000),
    "expected six-character CJK labels to keep extra headroom against browser-specific wrapping"
  );
});

test("preserves emoji and other astral characters when long text is split across slides", async () => {
  const longText = "这是一个🙂需要安全拆分的超长标题🚀".repeat(14);
  const zip = await generatePresentationZip({
    lines: [longText]
  });
  const slidePaths = await getOrderedSlidePaths(zip);

  assert.ok(slidePaths.length > 1);

  const chunks = [];
  for (const slidePath of slidePaths) {
    const slideDoc = await loadXml(zip, slidePath);
    const textShapes = getTextShapes(slideDoc);
    chunks.push(textShapes[0].text);
  }

  assert.equal(chunks.join(""), longText);
  assert.ok(!chunks.join("").includes("\uFFFD"));
});

test("keeps app metadata aligned with the generated slide count", async () => {
  const zip = await generatePresentationZip({
    lines: ["第一页", "第二页", "第三页", "第四页"]
  });
  const appDoc = await loadXml(zip, "docProps/app.xml");
  const metadata = getAppMetadata(appDoc);

  assert.equal(metadata.slides, "4");
  assert.equal(metadata.headingPairsSlideCount, "4");
  assert.equal(metadata.titlesOfPartsSize, "9");
});

test("rejects empty input before generating a presentation", async () => {
  const { generatePresentationBuffer } = await loadEngine();
  const templateBytes = new Uint8Array(await fs.readFile(TEMPLATE_PATH));

  await assert.rejects(
    () =>
      generatePresentationBuffer({
        templateBytes,
        lines: []
      }),
    /input does not contain any lines/
  );
});

async function loadEngine() {
  return import(ENGINE_PATH);
}

async function generatePresentationZip({ lines, accent }) {
  const { generatePresentationBuffer } = await loadEngine();
  const templateBytes = new Uint8Array(await fs.readFile(TEMPLATE_PATH));
  const { outputBuffer } = await generatePresentationBuffer({
    templateBytes,
    lines,
    accentColor: accent || null
  });

  return JSZip.loadAsync(outputBuffer);
}

async function loadXml(zip, filePath) {
  const entry = zip.file(filePath);
  assert.ok(entry, `missing XML file: ${filePath}`);
  return parseXml(await entry.async("string"));
}

function parseXml(xmlText) {
  return new DOMParser().parseFromString(xmlText, "application/xml");
}

async function getOrderedSlidePaths(zip) {
  const presentation = await loadXml(zip, "ppt/presentation.xml");
  const rels = await loadXml(zip, "ppt/_rels/presentation.xml.rels");
  const relMap = new Map(
    Array.from(rels.getElementsByTagNameNS(XML_NS.rel, "Relationship")).map((node) => [
      node.getAttribute("Id"),
      node.getAttribute("Target")
    ])
  );

  return Array.from(presentation.getElementsByTagNameNS(XML_NS.p, "sldId")).map((node) => {
    const rid = node.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id"
    );
    const target = relMap.get(rid);
    assert.ok(target, `missing relationship target for ${rid}`);
    return path.posix.join("ppt", target);
  });
}

function getTextShapes(slideDoc) {
  return Array.from(slideDoc.getElementsByTagNameNS(XML_NS.p, "sp"))
    .map((shapeNode) => {
      const texts = Array.from(shapeNode.getElementsByTagNameNS(XML_NS.a, "t")).map((node) => node.textContent);
      if (texts.length === 0) {
        return null;
      }

      const xfrm = shapeNode.getElementsByTagNameNS(XML_NS.a, "xfrm")[0];
      const off = xfrm && xfrm.getElementsByTagNameNS(XML_NS.a, "off")[0];
      const ext = xfrm && xfrm.getElementsByTagNameNS(XML_NS.a, "ext")[0];
      const bodyPr = shapeNode.getElementsByTagNameNS(XML_NS.a, "bodyPr")[0];
      const runProps = shapeNode.getElementsByTagNameNS(XML_NS.a, "rPr")[0];
      const lineNode = runProps && runProps.getElementsByTagNameNS(XML_NS.a, "ln")[0];

      return {
        text: texts.join(""),
        fontSize: runProps ? runProps.getAttribute("sz") : null,
        lineWidth: lineNode ? lineNode.getAttribute("w") : null,
        lineColor: readColorValue(lineNode),
        bodyAnchor: bodyPr ? bodyPr.getAttribute("anchor") : null,
        paragraphAlign: shapeNode.getElementsByTagNameNS(XML_NS.a, "pPr")[0]?.getAttribute("algn") ?? null,
        extKey: ext ? `${ext.getAttribute("cx")}x${ext.getAttribute("cy")}` : "",
        layoutKey: [off?.getAttribute("x"), off?.getAttribute("y"), ext?.getAttribute("cx"), ext?.getAttribute("cy")].join(":")
      };
    })
    .filter(Boolean);
}

function readColorValue(node) {
  if (!node) {
    return null;
  }
  const srgb = node.getElementsByTagNameNS(XML_NS.a, "srgbClr")[0];
  if (srgb) {
    return srgb.getAttribute("val");
  }
  const scheme = node.getElementsByTagNameNS(XML_NS.a, "schemeClr")[0];
  if (scheme) {
    return scheme.getAttribute("val");
  }
  return null;
}

function getAppMetadata(appDoc) {
  const slides = appDoc.getElementsByTagNameNS(XML_NS.app, "Slides")[0]?.textContent ?? null;
  const headingPairs = Array.from(appDoc.getElementsByTagNameNS(XML_NS.vt, "i4")).map((node) => node.textContent);
  const titlesOfPartsNode = appDoc.getElementsByTagNameNS(XML_NS.app, "TitlesOfParts")[0];
  const titlesOfPartsSize = titlesOfPartsNode
    ? titlesOfPartsNode.getElementsByTagNameNS(XML_NS.vt, "lpstr").length
    : 0;

  return {
    slides,
    headingPairsSlideCount: headingPairs[0] ?? null,
    titlesOfPartsSize: String(titlesOfPartsSize)
  };
}

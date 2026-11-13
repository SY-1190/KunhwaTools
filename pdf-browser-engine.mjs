// SPDX-License-Identifier: AGPL-3.0-or-later
import mupdf from "./assets/vendor/mupdf-1.28.6/dist/mupdf.js";

const { PDFPage, Matrix } = mupdf;
const MAX_PAGES = 150;
const MAX_BYTES = 55 * 1024 * 1024;
const clamp = (value, max) => Math.max(0, Math.min(value, max));
const number = (n) => Number(n.toFixed(5));
const hex = (n) => Math.round(clamp(n, 1) * 255).toString(16).padStart(2, "0");
const rgb = (value) => {
  if (!/^#[0-9a-f]{6}$/i.test(value)) throw new Error("색상 값이 올바르지 않습니다.");
  return [1, 3, 5].map((i) => number(parseInt(value.slice(i, i + 2), 16) / 255)).join(" ");
};

function openDocument(data) {
  if (!data.byteLength || data.byteLength > MAX_BYTES) throw new Error("PDF는 55MB 이하만 열 수 있습니다.");
  let doc;
  try { doc = new mupdf.PDFDocument(data); }
  catch { throw new Error("올바른 PDF 파일을 열어 주세요."); }
  try {
    doc.disableJS();
    if (doc.needsPassword() || !doc.hasPermission("edit")) throw new Error("암호가 있거나 편집이 제한된 PDF입니다.");
    const count = doc.countPages();
    if (!count || count > MAX_PAGES) throw new Error("PDF는 1~150페이지까지 열 수 있습니다.");
    return doc;
  } catch (error) { doc.destroy(); throw error; }
}

function geometry(page) {
  const [x, y, right, bottom] = page.getBounds();
  return { x, y, width: right - x, height: bottom - y };
}

function pageObjects(page, geometry, includeImages = false) {
  const { x, y, width, height } = geometry;
  const local = (r) => [clamp(r[0] - x, width), clamp(r[1] - y, height), clamp(r[2] - x, width), clamp(r[3] - y, height)];
  const visible = (r) => r[2] - r[0] > .1 && r[3] - r[1] > .1;
  const objects = [], images = new Map();
  let line, textIndex = 0, imageIndex = 0;
  const structured = page.toStructuredText("preserve-images");
  try {
    structured.walk({
      beginLine(bbox, wmode, direction) {
        line = { id: `t${textIndex++}`, type: "text", rect: local(bbox), text: "", direction, size: 12, color: "#17202a" };
      },
      onChar(char, _origin, font, size, _quad, color) {
        font.destroy();
        if (!line.text) {
          line.size = size;
          if (color.length === 3) line.color = "#" + color.map(hex).join("");
          else if (color.length === 1) line.color = "#" + hex(color[0]).repeat(3);
        }
        line.text += char;
      },
      endLine() {
        if (line.text.trim() && visible(line.rect)) objects.push(line);
      },
      onImageBlock(bbox, _transform, image) {
        const id = `i${imageIndex++}`, rect = local(bbox);
        if (visible(rect)) {
          objects.push({ id, type: "image", rect });
          if (includeImages) { images.set(id, image); return; }
        }
        image.destroy();
      },
    });
    return { objects, images };
  } finally { structured.destroy(); }
}

function raster(page, { width, height }, limit) {
  if (!(width > 0 && height > 0) || !Number.isFinite(width + height)) throw new Error("PDF 페이지 크기가 올바르지 않습니다.");
  const scale = Math.min(1.6, limit / Math.max(width, height));
  const pix = page.toPixmap(Matrix.scale(scale, scale), mupdf.ColorSpace.DeviceRGB);
  try { return new Blob([pix.asPNG()], { type: "image/png" }); }
  finally { pix.destroy(); }
}

export function inspectPdf(data) {
  const doc = openDocument(data);
  try {
    const count = doc.countPages(), pages = [];
    const limit = Math.min(1200, Math.floor(Math.sqrt(32_000_000 / count)));
    for (let i = 0; i < count; i++) {
      const page = doc.loadPage(i);
      try {
        const widgets = page.getWidgets();
        const hasSignature = widgets.some((widget) => widget.getFieldType() === "signature");
        widgets.forEach((widget) => widget.destroy());
        if (hasSignature) {
          throw new Error("전자서명이 있는 PDF는 편집할 수 없습니다.");
        }
        const box = geometry(page);
        pages.push({ width: box.width, height: box.height, image: raster(page, box, limit), objects: pageObjects(page, box).objects });
      } finally { page.destroy(); }
    }
    return { pages, pageCount: count };
  } finally { doc.destroy(); }
}

function validateRect(rect, { width, height }) {
  if (!Array.isArray(rect) || rect.length !== 4 || !rect.every(Number.isFinite) ||
      rect[0] < 0 || rect[1] < 0 || rect[2] > width || rect[3] > height ||
      rect[2] - rect[0] < 1 || rect[3] - rect[1] < 1) throw new Error("항목의 위치와 크기를 확인해 주세요.");
  return rect;
}

function eraseRect(source) {
  const [x, y, right, bottom] = source.rect;
  if (source.type === "image") return source.rect;
  if (Math.abs(source.direction[0]) >= Math.abs(source.direction[1])) {
    const inset = (bottom - y) * .2;
    return [x, y + inset, right, bottom - inset];
  }
  const inset = (right - x) * .2;
  return [x + inset, y, right - inset, bottom];
}

function cloneDictionary(doc, from) {
  const out = doc.newDictionary();
  if (from.isDictionary()) from.forEach((value, key) => out.put(key, value));
  return out;
}

function appendContent(doc, page, resources, content) {
  const pageObject = page.getObject();
  const previous = pageObject.get("Contents");
  const streams = doc.newArray();
  if (!previous.isNull()) {
    streams.push(doc.addStream("q\n", {}));
    if (previous.isArray()) previous.forEach((stream) => streams.push(stream));
    else streams.push(previous);
    content = "Q\n" + content;
  }
  streams.push(doc.addStream(content, {}));
  pageObject.put("Resources", resources);
  pageObject.put("Contents", streams);
}

function resource(dictionary, ref) {
  let i = 0, name;
  do { name = "KTools" + i++; } while (!dictionary.get(name).isNull());
  dictionary.put(name, ref);
  return "/" + name;
}

function layoutText(text, size, width, font) {
  const lines = [];
  for (const paragraph of text.replace(/\r\n?/g, "\n").split("\n")) {
    let glyphs = "", used = 0;
    for (const char of paragraph.replace(/\t/g, "    ")) {
      const glyph = font.encodeCharacter(char);
      if (!glyph) throw new Error("지원되지 않는 글자가 있습니다. 한글·영문 및 일반 기호를 사용해 주세요.");
      const advance = font.advanceGlyph(glyph) * size;
      if (advance > width) throw new Error("글자 영역이 너무 좁습니다.");
      if (glyphs && used + advance > width + .01) { lines.push(glyphs); glyphs = ""; used = 0; }
      glyphs += glyph.toString(16).padStart(4, "0");
      used += advance;
    }
    lines.push(glyphs);
  }
  return lines;
}

function readImage(value) {
  if (typeof value !== "string" || !/^data:image\/(png|jpeg|webp);base64,/i.test(value) || value.length > 16 * 1024 * 1024 + 40) {
    throw new Error("12MB 이하의 PNG, JPG 또는 WebP 이미지를 선택해 주세요.");
  }
  const data = Uint8Array.from(atob(value.slice(value.indexOf(",") + 1)), (c) => c.charCodeAt(0));
  try { return new mupdf.Image(data); }
  catch { throw new Error("이미지를 읽을 수 없습니다."); }
}

function editPage(doc, page, edits, fontState) {
  const box = geometry(page);
  const { objects, images } = pageObjects(page, box, true);
  try {
    const byId = new Map(objects.map((item) => [item.id, item]));
    const seen = new Set();
    const paint = [], content = [];
    const toPdf = Matrix.invert(page.getTransform()).map(number).join(" ");
    for (const op of edits) {
      if (op.type !== "text" && op.type !== "image") throw new Error("지원되지 않는 편집 항목입니다.");
      const source = byId.get(op.sourceId);
      if (op.sourceId && (source?.type !== op.type || seen.has(op.sourceId))) throw new Error("원본 항목을 다시 선택해 주세요.");
      if (source) {
        seen.add(source.id);
        const [x, y, r, b] = eraseRect(source);
        const annot = page.createAnnotation("Redact");
        try {
          annot.setRect([x + box.x, y + box.y, r + box.x, b + box.y]);
          annot.applyRedaction(0, op.type === "image" ? PDFPage.REDACT_IMAGE_PIXELS : PDFPage.REDACT_IMAGE_NONE,
            PDFPage.REDACT_LINE_ART_NONE, op.type === "image" ? PDFPage.REDACT_TEXT_NONE : PDFPage.REDACT_TEXT_REMOVE);
        } finally { annot.destroy(); }
        if (op.type === "text") paint.push(`${rgb(op.fill || "#ffffff")} rg ${number(x)} ${number(y)} ${number(r - x)} ${number(b - y)} re f`);
      }
    }
    const inherited = page.getObject().getInheritable("Resources");
    const resources = cloneDictionary(doc, inherited);
    const fonts = cloneDictionary(doc, inherited.get("Font"));
    const xobjects = cloneDictionary(doc, inherited.get("XObject"));
    let fontName;
    for (const op of edits) {
      if (op.deleted) continue;
      const [x, y, r, b] = validateRect(op.rect, box).map(number);
      if (op.type === "image") {
        const image = op.image ? readImage(op.image) : images.get(op.sourceId);
        if (!image) throw new Error("추가할 이미지를 선택해 주세요.");
        try {
          const name = resource(xobjects, doc.addImage(image));
          content.push(`q ${number(r - x)} 0 0 ${number(y - b)} ${x} ${b} cm ${name} Do Q`);
        } finally { if (op.image) image.destroy(); }
      } else {
        const size = Number(op.size);
        if (typeof op.text !== "string" || op.text.length > 10000 || !(size >= 6 && size <= 120)) throw new Error("글자 내용 또는 크기가 올바르지 않습니다.");
        if (!fontState.font) { fontState.font = new mupdf.Font("ko"); fontState.ref = doc.addFont(fontState.font); }
        fontName ||= resource(fonts, fontState.ref);
        const lines = layoutText(op.text, size, r - x, fontState.font);
        if (y + size * (1.24 + (lines.length - 1) * 1.22) > b + .01) throw new Error("글자가 영역보다 큽니다. 너비·높이를 늘리거나 글자 크기를 줄여 주세요.");
        const show = lines.map((line, i) => `1 0 0 -1 ${x} ${number(y + size * (1 + i * 1.22))} Tm <${line}> Tj`).join("\n");
        content.push(`BT ${fontName} ${size} Tf ${rgb(op.color)} rg\n${show}\nET`);
      }
    }
    if (fontName) resources.put("Font", fonts);
    if (edits.some((op) => op.type === "image" && !op.deleted)) resources.put("XObject", xobjects);
    if (paint.length || content.length) {
      // getTransform maps PDF user space into the displayed page, including crop and rotation.
      appendContent(doc, page, resources, `q ${toPdf} cm 1 0 0 1 ${number(box.x)} ${number(box.y)} cm\n${paint.concat(content).join("\n")}\nQ\n`);
    }
  } finally { images.forEach((image) => image.destroy()); }
}

export function editPdf(data, operations, preview = false) {
  if (!Array.isArray(operations) || operations.length > 500) throw new Error("편집 항목은 500개 이하만 저장할 수 있습니다.");
  const doc = openDocument(data);
  const fontState = {};
  try {
    const byPage = new Map();
    for (const op of operations) {
      if (!Number.isInteger(op.page) || op.page < 0 || op.page >= doc.countPages()) throw new Error("편집할 페이지가 올바르지 않습니다.");
      if (!byPage.has(op.page)) byPage.set(op.page, []);
      byPage.get(op.page).push(op);
    }
    const images = {};
    const limit = Math.min(1200, Math.floor(Math.sqrt(32_000_000 / Math.max(1, byPage.size))));
    for (const [index, edits] of byPage) {
      const page = doc.loadPage(index);
      try { editPage(doc, page, edits, fontState); }
      finally { page.destroy(); }
      if (preview) {
        const updated = doc.loadPage(index);
        try { images[index] = { image: raster(updated, geometry(updated), limit) }; }
        finally { updated.destroy(); }
      }
    }
    if (preview) return images;
    if (fontState.font) doc.subsetFonts();
    const output = doc.saveToBuffer("garbage=4,compress");
    try { return new Blob([output.asUint8Array()], { type: "application/pdf" }); }
    finally { output.destroy(); }
  } finally { fontState.font?.destroy(); doc.destroy(); }
}

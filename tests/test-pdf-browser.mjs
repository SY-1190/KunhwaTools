import assert from "node:assert/strict";
import test from "node:test";
import mupdf from "../assets/vendor/mupdf-1.28.6/dist/mupdf.js";
import { inspectPdf, editPdf } from "../pdf-browser-engine.mjs";

const png = (rgb) => {
  const pix = new mupdf.Pixmap(mupdf.ColorSpace.DeviceRGB, [0, 0, 8, 8], false);
  const pixels = pix.getPixels();
  for (let i = 0; i < pixels.length; i += 3) pixels.set(rgb, i);
  const bytes = new Uint8Array(pix.asPNG());
  pix.destroy();
  return bytes;
};
const url = (rgb) => "data:image/png;base64," + Buffer.from(png(rgb)).toString("base64");

function fixture(rotation = 0) {
  const doc = new mupdf.PDFDocument();
  const font = new mupdf.Font("Helvetica");
  const image = new mupdf.Image(png([210, 10, 10]));
  const page = doc.addPage([0, 0, 300, 400], rotation, {
    Font: { F1: doc.addSimpleFont(font) },
    XObject: { I1: doc.addImage(image) },
  }, "BT /F1 15 Tf 30 340 Td (Original line) Tj 0 -28 Td (Keep this line) Tj ET\nq 50 0 0 50 45 210 cm /I1 Do Q");
  doc.insertPage(-1, page);
  const saved = doc.saveToBuffer("");
  const result = new Uint8Array(saved.asUint8Array());
  saved.destroy(); image.destroy(); font.destroy(); doc.destroy();
  return result;
}

const bytes = async (blob) => new Uint8Array(await blob.arrayBuffer());
const view = async (blob, fn) => {
  const doc = new mupdf.PDFDocument(await bytes(blob)), page = doc.loadPage(0);
  try { return fn(page); } finally { page.destroy(); doc.destroy(); }
};

test("replaces real text with searchable Korean and keeps adjacent text", async () => {
  const original = fixture();
  const [line] = inspectPdf(original).pages[0].objects.filter((item) => item.type === "text");
  const op = { ...line, sourceId: line.id, page: 0, rect: [30, 40, 220, 69], text: "한글 수정", color: "#000000" };
  assert.ok(editPdf(original, [op], true)[0].image.size > 0);
  await view(editPdf(original, [op]), (page) => {
    const text = page.toStructuredText().asText();
    assert.match(text, /한글 수정/);
    assert.match(text, /Keep this line/);
    assert.doesNotMatch(text, /Original line/);
  });
  assert.equal(inspectPdf(original).pages[0].objects[0].text, "Original line");
});

test("replaces original image and adds another image without flattening text", async () => {
  const original = fixture();
  const image = inspectPdf(original).pages[0].objects.find((item) => item.type === "image");
  const edited = editPdf(original, [
    { ...image, sourceId: image.id, page: 0, image: url([0, 170, 0]) },
    { type: "image", page: 0, rect: [120, 140, 170, 190], image: url([0, 0, 190]) },
  ]);
  await view(edited, (page) => {
    assert.match(page.toStructuredText().asText(), /Keep this line/);
    const pix = page.toPixmap(mupdf.Matrix.identity, mupdf.ColorSpace.DeviceRGB);
    const at = (x, y) => [...pix.getPixels().slice(y * pix.getStride() + 3 * x, y * pix.getStride() + 3 * x + 3)];
    assert.deepEqual(at(55, 155), [0, 170, 0]);
    assert.deepEqual(at(130, 155), [0, 0, 190]);
    pix.destroy();
  });
});

test("moves a source image and adds searchable text on a rotated page", async () => {
  const original = fixture(90);
  const image = inspectPdf(original).pages[0].objects.find((item) => item.type === "image");
  await view(editPdf(original, [
    { ...image, sourceId: image.id, page: 0, rect: [240, 145, 290, 195] },
    { type: "text", page: 0, rect: [90, 65, 230, 95], size: 16, text: "회전 한글", color: "#000000" },
  ]), (page) => {
    assert.equal(page.getObject().getInheritable("Rotate").asNumber(), 90);
    assert.match(page.toStructuredText().asText(), /회전 한글/);
    const pix = page.toPixmap(mupdf.Matrix.identity, mupdf.ColorSpace.DeviceRGB);
    const at = (x, y) => [...pix.getPixels().slice(y * pix.getStride() + 3 * x, y * pix.getStride() + 3 * x + 3)];
    assert.deepEqual(at(250, 160), [210, 10, 10]);
    assert.notDeepEqual(at(image.rect[0] + 5, image.rect[1] + 5), [210, 10, 10]);
    pix.destroy();
  });
});

test("rejects invalid and encrypted input and off-page edits", () => {
  assert.throws(() => inspectPdf(new Uint8Array([1, 2, 3])));
  const doc = new mupdf.PDFDocument(fixture());
  const encrypted = doc.saveToBuffer("encrypt=aes-256,user-password=test");
  assert.throws(() => inspectPdf(encrypted.asUint8Array()), /암호/);
  encrypted.destroy(); doc.destroy();
  assert.throws(() => editPdf(fixture(), [{type:"text",page:0,rect:[0,0,400,80],size:16,text:"test",color:"#000000"}]), /위치와 크기/);
});

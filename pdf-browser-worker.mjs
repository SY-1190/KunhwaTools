// SPDX-License-Identifier: AGPL-3.0-or-later
let engine;
let source;

self.onmessage = async ({ data: { id, action, file, operations } }) => {
  const report = (phase, done, total) => self.postMessage({ id, progress: { phase, done, total } });
  try {
    if (!engine) {
      if (action === "open") report("engine");
      try { engine = await import("./pdf-browser-engine.mjs?v=0.2.3"); }
      catch { throw new Error("PDF 편집 엔진을 불러오지 못했습니다. 페이지를 새로고침해 주세요."); }
    }
    let result;
    if (action === "open") {
      report("read");
      const bytes = new Uint8Array(await file.arrayBuffer());
      result = engine.inspectPdf(bytes, (done, total) => report("pages", done, total));
      source = bytes;
    } else {
      if (!source) throw new Error("PDF 파일을 다시 열어 주세요.");
      result = engine.editPdf(source, operations, action === "preview");
    }
    self.postMessage({ id, result });
  } catch (error) {
    self.postMessage({ id, error: error?.message || "PDF를 처리하지 못했습니다." });
  }
};

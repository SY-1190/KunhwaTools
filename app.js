pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";

const $ = (id) => document.getElementById(id);

const CANCELLED_ERROR = "__CANCELLED__";

const operationConfigs = {
  pdfToImage: {
    cancelBtnId: "cancelPdfToImage",
    progressId: "progressPdfToImage",
    progressLabelId: "progressPdfToImageLabel",
    etaLabelId: "etaPdfToImageLabel",
    statusId: "pdfToImageStatus",
  },
  imageToPdf: {
    cancelBtnId: "cancelImageToPdf",
    progressId: "progressImageToPdf",
    progressLabelId: "progressImageToPdfLabel",
    etaLabelId: "etaImageToPdfLabel",
    statusId: "imageToPdfStatus",
  },
  arrange: {
    cancelBtnId: "cancelArrange",
    progressId: "progressArrange",
    progressLabelId: "progressArrangeLabel",
    etaLabelId: "etaArrangeLabel",
    statusId: "arrangePdfStatus",
  },
  mergePdf: {
    cancelBtnId: "cancelMergePdf",
    progressId: "progressMergePdf",
    progressLabelId: "progressMergePdfLabel",
    etaLabelId: "etaMergePdfLabel",
    statusId: "mergePdfStatus",
  },
  pdfCompress: {
    cancelBtnId: "cancelPdfCompress",
    progressId: "progressPdfCompress",
    progressLabelId: "progressPdfCompressLabel",
    etaLabelId: "etaPdfCompressLabel",
    statusId: "pdfCompressStatus",
  },
  dwgToPdf: {
    cancelBtnId: "cancelDwgToPdf",
    progressId: "progressDwgToPdf",
    progressLabelId: "progressDwgToPdfLabel",
    etaLabelId: "etaDwgToPdfLabel",
    statusId: "dwgToPdfStatus",
  },
  resize: {
    cancelBtnId: "cancelResize",
    progressId: "progressResize",
    progressLabelId: "progressResizeLabel",
    etaLabelId: "etaResizeLabel",
    statusId: "resizeStatus",
  },
  format: {
    cancelBtnId: "cancelFormat",
    progressId: "progressFormat",
    progressLabelId: "progressFormatLabel",
    etaLabelId: "etaFormatLabel",
    statusId: "formatStatus",
  },
  batchRename: {
    cancelBtnId: "cancelBatchRename",
    progressId: "progressBatchRename",
    progressLabelId: "progressBatchRenameLabel",
    etaLabelId: "etaBatchRenameLabel",
    statusId: "batchRenameStatus",
  },
};

const operationStates = {};
const globalBusyState = {
  manualCount: 0,
  message: "작업 처리 중입니다...",
};

const refreshGlobalBusy = (message) => {
  const busyWrap = $("globalBusy");
  const busyText = $("globalBusyText");
  if (!busyWrap) return;
  if (message) globalBusyState.message = message;
  const anyOpRunning = Object.values(operationStates).some((s) => s.running);
  const active = anyOpRunning || globalBusyState.manualCount > 0;
  busyWrap.classList.toggle("hidden", !active);
  if (busyText) busyText.textContent = globalBusyState.message;
};

const setGlobalBusyMessage = (message) => {
  if (!message) return;
  globalBusyState.message = message;
  const busyText = $("globalBusyText");
  if (busyText) busyText.textContent = message;
};

const beginGlobalBusy = (message) => {
  globalBusyState.manualCount += 1;
  refreshGlobalBusy(message);
};

const endGlobalBusy = () => {
  globalBusyState.manualCount = Math.max(0, globalBusyState.manualCount - 1);
  refreshGlobalBusy();
};

const formatBytes = (bytes) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
};

const formatDurationShort = (seconds) => {
  if (!Number.isFinite(seconds) || seconds < 0) return "--:--";
  const s = Math.round(seconds);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  if (h > 0) return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(sec).padStart(2, "0")}`;
  return `${String(m).padStart(2, "0")}:${String(sec).padStart(2, "0")}`;
};

const setStatus = (id, text) => {
  const el = $(id);
  if (el) el.textContent = text;
  if (!text) return;
  const opEntry = Object.entries(operationConfigs).find(([, cfg]) => cfg.statusId === id);
  if (!opEntry) return;
  const [opKey] = opEntry;
  if (operationStates[opKey]?.running) setGlobalBusyMessage(text);
};

const downloadBlob = (blob, fileName) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
};

const readAsArrayBuffer = (file) =>
  new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsArrayBuffer(file);
  });

const readAsDataURL = (file) =>
  new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsDataURL(file);
  });

const readAsText = (file) =>
  new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsText(file, "utf-8");
  });

const isHeicLikeFile = (file) => {
  const ext = (file?.name?.split(".").pop() || "").toLowerCase();
  return ext === "heic" || ext === "heif" || /image\/hei(c|f)/i.test(file?.type || "");
};

const dedupeFiles = (files) => {
  const seen = new Set();
  const out = [];
  files.forEach((f) => {
    const key = `${f.name}__${f.size}__${f.lastModified}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push(f);
  });
  return out;
};

const sameFileSignature = (a, b) =>
  a.length === b.length &&
  a.every((f, i) => f.name === b[i].name && f.size === b[i].size && f.lastModified === b[i].lastModified);

const ICONS = {
  trash3: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M6.5 1h3a.5.5 0 0 1 .5.5v1H6v-1a.5.5 0 0 1 .5-.5M11 2.5v-1A1.5 1.5 0 0 0 9.5 0h-3A1.5 1.5 0 0 0 5 1.5v1H1.5a.5.5 0 0 0 0 1h.538l.853 10.66A2 2 0 0 0 4.885 16h6.23a2 2 0 0 0 1.994-1.84l.853-10.66h.538a.5.5 0 0 0 0-1zm1.958 1-.846 10.58a1 1 0 0 1-.997.92h-6.23a1 1 0 0 1-.997-.92L3.042 3.5zm-7.487 1a.5.5 0 0 1 .528.47l.5 8.5a.5.5 0 0 1-.998.06L5 5.03a.5.5 0 0 1 .47-.53Zm5.058 0a.5.5 0 0 1 .47.53l-.5 8.5a.5.5 0 1 1-.998-.06l.5-8.5a.5.5 0 0 1 .528-.47M8 4.5a.5.5 0 0 1 .5.5v8.5a.5.5 0 0 1-1 0V5a.5.5 0 0 1 .5-.5"/></svg>`,
  undo: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path fill-rule="evenodd" d="M8 3a5 5 0 1 1-4.546 2.914.5.5 0 0 0-.908-.417A6 6 0 1 0 8 2z"/><path d="M8 4.466V.534a.25.25 0 0 0-.41-.192L5.23 2.308a.25.25 0 0 0 0 .384l2.36 1.966A.25.25 0 0 0 8 4.466"/></svg>`,
  download: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5"/><path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708z"/></svg>`,
  house: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="m8.354 1.146 6.5 6.5A.5.5 0 0 1 14.5 8.5H13v5a1 1 0 0 1-1 1h-2.5a.5.5 0 0 1-.5-.5V11a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v3a.5.5 0 0 1-.5.5H4a1 1 0 0 1-1-1v-5H1.5a.5.5 0 0 1-.354-.854l6.5-6.5a.5.5 0 0 1 .708 0"/></svg>`,
  person: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M8 8a3 3 0 1 0 0-6 3 3 0 0 0 0 6"/><path fill-rule="evenodd" d="M8 9a6 6 0 0 0-5.468 3.516A.75.75 0 0 0 3.205 14h9.59a.75.75 0 0 0 .673-1.484A6 6 0 0 0 8 9"/></svg>`,
  x: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M2.146 2.146a.5.5 0 0 1 .708 0L8 7.293l5.146-5.147a.5.5 0 0 1 .708.708L8.707 8l5.147 5.146a.5.5 0 0 1-.708.708L8 8.707l-5.146 5.147a.5.5 0 0 1-.708-.708L7.293 8 2.146 2.854a.5.5 0 0 1 0-.708"/></svg>`,
};

const setIconButton = (id, iconKey) => {
  const btn = $(id);
  if (!btn || !ICONS[iconKey]) return;
  btn.innerHTML = ICONS[iconKey];
};

const syncFilesToInput = (inputId, files) => {
  const input = $(inputId);
  if (!input) return;
  const dt = new DataTransfer();
  files.forEach((f) => dt.items.add(f));
  input.files = dt.files;
  input.dataset.replaceFilesOnce = "1";
  input.dispatchEvent(new Event("change", { bubbles: true }));
};

const setupDropZones = () => {
  const fileStoreByInput = new Map();

  const updateDropLabel = (inputId) => {
    const input = $(inputId);
    const nameEl = $(`${inputId}DropName`);
    if (!input || !nameEl) return;
    const files = [...(input.files || [])];
    if (!files.length) {
      nameEl.textContent = "선택된 파일 없음";
      return;
    }
    if (files.length === 1) {
      nameEl.textContent = files[0].name;
      return;
    }
    nameEl.textContent = `${files[0].name} 외 ${files.length - 1}개`;
  };

  document.querySelectorAll(".drop-zone[data-file-input]").forEach((zone) => {
    const inputId = zone.dataset.fileInput;
    const input = $(inputId);
    if (!input) return;
    const preserveFiles = input.multiple && zone.dataset.preserveFiles === "true";

    zone.addEventListener("click", () => input.click());
    zone.addEventListener("dragover", (e) => {
      e.preventDefault();
      zone.classList.add("dragover");
    });
    zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
    zone.addEventListener("drop", (e) => {
      e.preventDefault();
      zone.classList.remove("dragover");
      const dropped = [...(e.dataTransfer?.files || [])];
      if (!dropped.length) return;
      const allowMultiple = input.multiple;
      const dt = new DataTransfer();
      if (allowMultiple) {
        [...(input.files || [])].forEach((f) => dt.items.add(f));
        dropped.forEach((f) => dt.items.add(f));
      } else {
        dt.items.add(dropped[0]);
      }
      input.files = dt.files;
      fileStoreByInput.set(inputId, [...input.files]);
      input.dispatchEvent(new Event("change", { bubbles: true }));
      updateDropLabel(inputId);
    });

    input.addEventListener("change", () => {
      const current = [...(input.files || [])];
      if (!preserveFiles) {
        fileStoreByInput.set(inputId, current);
        updateDropLabel(inputId);
        return;
      }
      if (input.dataset.replaceFilesOnce === "1") {
        input.dataset.replaceFilesOnce = "0";
        fileStoreByInput.set(inputId, current);
        updateDropLabel(inputId);
        return;
      }
      const prev = fileStoreByInput.get(inputId) || [];
      const merged = dedupeFiles([...prev, ...current]);
      if (!sameFileSignature(current, merged)) {
        const dt = new DataTransfer();
        merged.forEach((f) => dt.items.add(f));
        input.files = dt.files;
      }
      fileStoreByInput.set(inputId, [...(input.files || [])]);
      updateDropLabel(inputId);
    });
    updateDropLabel(inputId);
  });
};

const loadImageFromFile = async (file) => {
  const dataUrl = await readAsDataURL(file);
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ img, dataUrl });
    img.onerror = reject;
    img.src = dataUrl;
  });
};

const dataUrlToUint8Array = (dataUrl) => {
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

const escapeRegExp = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const splitFileName = (fileName) => {
  const lastDot = fileName.lastIndexOf(".");
  if (lastDot <= 0) return { stem: fileName, ext: "" };
  return {
    stem: fileName.slice(0, lastDot),
    ext: fileName.slice(lastDot),
  };
};

const sanitizeFileStem = (name) => {
  let out = String(name || "")
    .replace(/[\u0000-\u001f\u0080-\u009f]/g, "")
    .replace(/[<>:"/\\|?*]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "")
    .replace(/^[. ]+/g, "");
  if (!out || out === "." || out === "..") out = "file";
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(out)) out = `${out}_`;
  if (out.length > 180) out = out.slice(0, 180);
  return out;
};

const initOperations = () => {
  Object.entries(operationConfigs).forEach(([key, cfg]) => {
    operationStates[key] = { running: false, cancelled: false, startedAt: 0 };
    const cancelBtn = $(cfg.cancelBtnId);
    if (cancelBtn) {
      cancelBtn.disabled = true;
      cancelBtn.addEventListener("click", () => {
        if (!operationStates[key].running) return;
        operationStates[key].cancelled = true;
        setStatus(
          cfg.statusId,
          "취소 요청됨: 현재 파일 단계가 끝나면 안전하게 중단합니다."
        );
      });
    }
    updateProgress(key, 0, 100);
  });
};

const updateProgress = (opKey, done, total) => {
  const cfg = operationConfigs[opKey];
  const state = operationStates[opKey];
  if (!cfg) return;
  const progress = $(cfg.progressId);
  const label = $(cfg.progressLabelId);
  const etaLabel = $(cfg.etaLabelId);
  const percent = total > 0 ? Math.min(100, Math.round((done / total) * 100)) : 0;
  if (progress) progress.value = percent;
  if (label) label.textContent = `${percent}%`;
  if (!etaLabel) return;

  if (!state?.running || done <= 0 || total <= 0 || done >= total) {
    etaLabel.textContent = done >= total ? "ETA 00:00" : "ETA --:--";
    return;
  }

  const elapsedSec = (Date.now() - state.startedAt) / 1000;
  const doneRatio = done / total;
  if (elapsedSec < 0.3 || doneRatio <= 0) {
    etaLabel.textContent = "ETA --:--";
    return;
  }
  const totalEstimatedSec = elapsedSec / doneRatio;
  const remainSec = Math.max(0, totalEstimatedSec - elapsedSec);
  etaLabel.textContent = `ETA ${formatDurationShort(remainSec)}`;
};

const startOperation = (opKey, statusText) => {
  const cfg = operationConfigs[opKey];
  const state = operationStates[opKey];
  if (!cfg || !state) return;
  state.running = true;
  state.cancelled = false;
  state.startedAt = Date.now();
  const cancelBtn = $(cfg.cancelBtnId);
  if (cancelBtn) cancelBtn.disabled = false;
  updateProgress(opKey, 0, 100);
  setStatus(cfg.statusId, statusText);
  refreshGlobalBusy(statusText || "작업 처리 중입니다...");
};

const endOperation = (opKey, statusText) => {
  const cfg = operationConfigs[opKey];
  const state = operationStates[opKey];
  if (!cfg || !state) return;
  state.running = false;
  state.startedAt = 0;
  const cancelBtn = $(cfg.cancelBtnId);
  if (cancelBtn) cancelBtn.disabled = true;
  if (statusText) setStatus(cfg.statusId, statusText);
  refreshGlobalBusy(statusText || "작업 처리 중입니다...");
};

const checkCancelled = (opKey) => {
  const state = operationStates[opKey];
  if (state?.cancelled) throw new Error(CANCELLED_ERROR);
};

const handleOperationError = (opKey, err) => {
  const cfg = operationConfigs[opKey];
  if (!cfg) return;
  if (err.message === CANCELLED_ERROR) {
    endOperation(opKey, "작업이 취소되었습니다.");
    return;
  }
  endOperation(opKey, `오류: ${err.message}`);
};

const parseSplitGroups = (input, maxPages) => {
  const groups = input.split("|").map((g) => g.trim()).filter(Boolean);
  return groups.map((groupText) => {
    const pages = [];
    const chunks = groupText.split(",").map((s) => s.trim()).filter(Boolean);
    chunks.forEach((chunk) => {
      if (chunk.includes("-")) {
        const [a, b] = chunk.split("-").map((n) => Number(n.trim()));
        if (!Number.isInteger(a) || !Number.isInteger(b)) return;
        const start = Math.min(a, b);
        const end = Math.max(a, b);
        for (let i = start; i <= end; i += 1) {
          if (i >= 1 && i <= maxPages) pages.push(i);
        }
      } else {
        const n = Number(chunk);
        if (Number.isInteger(n) && n >= 1 && n <= maxPages) pages.push(n);
      }
    });
    return pages;
  });
};

const parsePageTokens = (value, maxPages) => {
  const txt = (value || "").trim();
  if (!txt) return Array.from({ length: maxPages }, (_, i) => i + 1);
  const result = [];
  txt.split(",").map((v) => v.trim()).filter(Boolean).forEach((token) => {
    if (token.includes("-")) {
      const [aRaw, bRaw] = token.split("-");
      const a = Number(aRaw.trim());
      const b = Number(bRaw.trim());
      if (!Number.isInteger(a) || !Number.isInteger(b)) return;
      const start = Math.min(a, b);
      const end = Math.max(a, b);
      for (let i = start; i <= end; i += 1) {
        if (i >= 1 && i <= maxPages) result.push(i);
      }
    } else {
      const n = Number(token);
      if (Number.isInteger(n) && n >= 1 && n <= maxPages) result.push(n);
    }
  });
  const uniq = [...new Set(result)];
  return uniq.length ? uniq : Array.from({ length: maxPages }, (_, i) => i + 1);
};

const pdfToImageState = {
  files: [],
  pages: [],
  deletedStack: [],
  draggingPageId: null,
  placeholder: null,
};

const ensurePdfToImagePlaceholder = () => {
  if (pdfToImageState.placeholder) return pdfToImageState.placeholder;
  const ph = document.createElement("div");
  ph.className = "drag-placeholder";
  pdfToImageState.placeholder = ph;
  return ph;
};

const removePdfToImagePlaceholder = () => {
  const ph = pdfToImageState.placeholder;
  if (ph?.parentElement) ph.parentElement.removeChild(ph);
};

const renderPdfToImageGrid = () => {
  const previewBox = $("pdfToImagePreview");
  if (!previewBox) return;
  previewBox.innerHTML = "";
  pdfToImageState.pages.forEach((page) => {
    const item = document.createElement("div");
    item.className = "thumb-item";
    item.draggable = true;
    item.dataset.pageId = page.id;
    item.innerHTML = `<button class="thumb-delete" type="button" title="페이지 제외" aria-label="페이지 제외">${ICONS.trash3}</button><div class="thumb-label">${page.fileLabel} · p.${page.pageNo}</div>`;
    const img = document.createElement("img");
    img.src = page.thumbDataUrl;
    img.alt = `page-${page.pageNo}`;
    img.draggable = false;
    img.style.width = "100%";
    img.style.border = "1px solid #d4e2f1";
    img.style.borderRadius = "6px";
    item.prepend(img);
    previewBox.appendChild(item);
  });
};

const getPdfToImageDragAfterElement = (container, x, y) => {
  const items = [...container.querySelectorAll(".thumb-item:not(.dragging)")];
  if (!items.length) return { afterEl: null, nearestEl: null, before: false };
  const rects = items.map((el) => {
    const r = el.getBoundingClientRect();
    return {
      el,
      top: r.top,
      bottom: r.bottom,
      left: r.left,
      right: r.right,
      midX: r.left + r.width / 2,
      midY: r.top + r.height / 2,
      height: r.height,
    };
  });

  // 1) Build visual rows (wrapped grid friendly).
  const sortedByVisual = [...rects].sort((a, b) => (a.top - b.top) || (a.left - b.left));
  const rows = [];
  const rowTolerance = Math.max(10, Math.round((sortedByVisual[0]?.height || 80) * 0.35));
  sortedByVisual.forEach((cell) => {
    const current = rows[rows.length - 1];
    if (!current || Math.abs(cell.top - current.anchorTop) > rowTolerance) {
      rows.push({ anchorTop: cell.top, cells: [cell] });
    } else {
      current.cells.push(cell);
    }
  });
  rows.forEach((row) => {
    row.cells.sort((a, b) => a.left - b.left);
    row.top = Math.min(...row.cells.map((c) => c.top));
    row.bottom = Math.max(...row.cells.map((c) => c.bottom));
    row.midY = (row.top + row.bottom) / 2;
  });

  // 2) Pick active row by Y with stable boundaries.
  let rowIndex = 0;
  if (rows.length > 1) {
    if (y <= rows[0].midY) {
      rowIndex = 0;
    } else if (y >= rows[rows.length - 1].midY) {
      rowIndex = rows.length - 1;
    } else {
      for (let i = 0; i < rows.length - 1; i += 1) {
        const boundary = (rows[i].midY + rows[i + 1].midY) / 2;
        if (y < boundary) {
          rowIndex = i;
          break;
        }
        rowIndex = i + 1;
      }
    }
  }

  const row = rows[rowIndex];
  const rowCells = row.cells;
  const nearestEl = rowCells.reduce((near, c) => {
    if (!near) return c;
    const dNear = Math.abs(near.midX - x);
    const dCur = Math.abs(c.midX - x);
    return dCur < dNear ? c : near;
  }, null)?.el || null;

  // 3) In-row insertion by X midpoint.
  if (x <= rowCells[0].midX) {
    return { afterEl: rowCells[0].el, nearestEl: rowCells[0].el, before: true };
  }
  for (let i = 1; i < rowCells.length; i += 1) {
    if (x < rowCells[i].midX) {
      return { afterEl: rowCells[i].el, nearestEl, before: true };
    }
  }

  // End of row: insert before next row head if exists, otherwise append.
  const nextRow = rows[rowIndex + 1];
  if (nextRow?.cells?.length) {
    return { afterEl: nextRow.cells[0].el, nearestEl, before: false };
  }
  return { afterEl: null, nearestEl, before: false };
};

const applyPdfToImageDrop = () => {
  const grid = $("pdfToImagePreview");
  const ph = pdfToImageState.placeholder;
  if (!grid || !ph?.parentElement || !pdfToImageState.draggingPageId) return;
  const nextThumb = ph.nextElementSibling?.closest?.(".thumb-item");
  const moving = pdfToImageState.pages.find((p) => p.id === pdfToImageState.draggingPageId);
  if (!moving) return;
  const filtered = pdfToImageState.pages.filter((p) => p.id !== moving.id);
  if (nextThumb) {
    const nextId = nextThumb.dataset.pageId;
    const idx = filtered.findIndex((p) => p.id === nextId);
    if (idx >= 0) filtered.splice(idx, 0, moving);
    else filtered.push(moving);
  } else {
    filtered.push(moving);
  }
  pdfToImageState.pages = filtered;
  renderPdfToImageGrid();
};

const setupPdfToImagePreviewDnD = () => {
  const grid = $("pdfToImagePreview");
  if (!grid) return;
  grid.addEventListener("dragstart", (e) => {
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    pdfToImageState.draggingPageId = item.dataset.pageId;
    item.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  });
  grid.addEventListener("dragend", () => {
    pdfToImageState.draggingPageId = null;
    removePdfToImagePlaceholder();
    grid.querySelectorAll(".thumb-item.dragging").forEach((el) => el.classList.remove("dragging"));
  });
  grid.addEventListener("dragover", (e) => {
    if (!pdfToImageState.draggingPageId) return;
    e.preventDefault();
    const placeholder = ensurePdfToImagePlaceholder();
    const intent = getPdfToImageDragAfterElement(grid, e.clientX, e.clientY);
    if (!intent.afterEl) grid.appendChild(placeholder);
    else grid.insertBefore(placeholder, intent.afterEl);
  });
  grid.addEventListener("drop", (e) => {
    if (!pdfToImageState.draggingPageId) return;
    e.preventDefault();
    applyPdfToImageDrop();
    removePdfToImagePlaceholder();
  });
  grid.addEventListener("click", (e) => {
    const del = e.target.closest(".thumb-delete");
    if (!del) return;
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    const pageId = item.dataset.pageId;
    const idx = pdfToImageState.pages.findIndex((p) => p.id === pageId);
    if (idx < 0) return;
    const removed = pdfToImageState.pages[idx];
    pdfToImageState.deletedStack.push({ page: removed, index: idx });
    pdfToImageState.pages.splice(idx, 1);
    renderPdfToImageGrid();
    setStatus("pdfToImageStatus", `${removed.fileLabel} p.${removed.pageNo}가 제외되었습니다.`);
  });
};

const renderPdfToImagePreview = async (files) => {
  const previewBox = $("pdfToImagePreview");
  if (!previewBox) return;
  previewBox.innerHTML = "";
  pdfToImageState.files = [...files];
  pdfToImageState.pages = [];
  pdfToImageState.deletedStack = [];
  if (!files.length) return;
  beginGlobalBusy("PDF 미리보기를 준비 중입니다...");
  try {
    let renderedCount = 0;
    let totalPages = 0;
    const docs = [];
    for (let fi = 0; fi < files.length; fi += 1) {
      const buffer = await readAsArrayBuffer(files[fi]);
      const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
      docs.push(pdf);
      totalPages += pdf.numPages;
    }
    for (let fi = 0; fi < docs.length; fi += 1) {
      const pdf = docs[fi];
      const fileLabel = files[fi].name;
      for (let i = 1; i <= pdf.numPages; i += 1) {
        renderedCount += 1;
        setGlobalBusyMessage(`PDF 미리보기 생성 중 (${renderedCount}/${totalPages})`);
        const page = await pdf.getPage(i);
        const baseViewport = page.getViewport({ scale: 1 });
        const scale = 130 / baseViewport.width;
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement("canvas");
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
        pdfToImageState.pages.push({
          id: `${fi}:${i}`,
          fileIndex: fi,
          fileLabel,
          pageNo: i,
          thumbDataUrl: canvas.toDataURL("image/png"),
        });
      }
    }
    renderPdfToImageGrid();
    setStatus("pdfToImageStatus", `미리보기 완료: ${files.length}개 PDF, 총 ${pdfToImageState.pages.length}페이지`);
  } catch (err) {
    const note = document.createElement("div");
    note.className = "thumb-label";
    note.textContent = `미리보기 오류: ${err.message}`;
    previewBox.appendChild(note);
  } finally {
    endGlobalBusy();
  }
};

const buildNav = () => {
  const cards = [...document.querySelectorAll(".tool-card")];
  $("toolNav").innerHTML = cards
    .map(
      (card) =>
        `<a href="#${card.id}" data-tool-id="${card.id}">${card.querySelector("h2").textContent}</a>`
    )
    .join("");
};

const setupThemeToggle = () => {
  const button = $("themeToggle");
  if (!button) return;
  const key = "kunhwa-tools-theme";
  const setThemeButtonLabel = () => {
    const isDark = document.body.classList.contains("dark");
    button.textContent = isDark ? "🌙" : "☀";
    button.setAttribute("aria-label", isDark ? "다크 모드" : "화이트 모드");
    button.title = isDark ? "다크 모드" : "화이트 모드";
  };
  const saved = localStorage.getItem(key);
  if (saved === "dark") document.body.classList.add("dark");
  setThemeButtonLabel();
  button.addEventListener("click", () => {
    document.body.classList.toggle("dark");
    localStorage.setItem(key, document.body.classList.contains("dark") ? "dark" : "light");
    setThemeButtonLabel();
  });
};

const setupNavActive = () => {
  const page = document.body.dataset.page;
  if (!page) return;
  document.querySelectorAll(".tool-nav a[data-page]").forEach((link) => {
    link.classList.toggle("active", link.dataset.page === page);
  });
};

const setupHashStageRouter = () => {
  const hub = $("quickHub");
  const stageBar = $("stageBar");
  const stageTitle = $("stageTitleText");
  const stageMenu = $("stageMenuScroll");
  const stages = [...document.querySelectorAll(".tool-stage")];
  const validIds = new Set(stages.map((s) => s.id));
  if (!hub || !stageBar || !stageTitle || !stages.length) return;

  const setMenuActive = (id) => {
    if (!stageMenu) return;
    const chips = [...stageMenu.querySelectorAll(".stage-chip[data-stage-target]")];
    chips.forEach((chip) => chip.classList.toggle("active", chip.dataset.stageTarget === id));
    const active = chips.find((chip) => chip.dataset.stageTarget === id);
    active?.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
  };

  const setHome = (replaceHash = true) => {
    document.body.classList.remove("stage-mode");
    document.body.classList.add("home-mode");
    hub.classList.remove("hidden");
    stageBar.classList.add("hidden");
    stages.forEach((s) => s.classList.remove("active-stage"));
    setMenuActive("");
    if (replaceHash) history.replaceState(null, "", "#");
  };

  const setStage = (id, replaceHash = true) => {
    const target = stages.find((s) => s.id === id);
    if (!target) {
      setHome(replaceHash);
      return;
    }
    document.body.classList.remove("home-mode");
    document.body.classList.add("stage-mode");
    hub.classList.add("hidden");
    stageBar.classList.remove("hidden");
    stages.forEach((s) => s.classList.toggle("active-stage", s.id === id));
    setMenuActive(id);
    stageTitle.textContent = target.querySelector("h2")?.textContent || id;
    if (replaceHash) history.replaceState(null, "", `#${id}`);
    target.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  $("backToHub")?.addEventListener("click", () => setHome());

  document.querySelectorAll("a[href^='#']").forEach((link) => {
    link.addEventListener("click", (e) => {
      const id = (link.getAttribute("href") || "").replace("#", "").trim();
      if (!id) return;
      if (!validIds.has(id)) return;
      e.preventDefault();
      setStage(id);
    });
  });

  const applyFromHash = () => {
    const hash = (window.location.hash || "").replace("#", "").trim();
    if (!hash || !validIds.has(hash)) {
      setHome(false);
      return;
    }
    setStage(hash, false);
  };

  window.addEventListener("hashchange", applyFromHash);
  applyFromHash();
};

const setupStageRouter = () => {
  const nav = $("toolNav");
  const stageHeader = $("stageHeader");
  const stageTitle = $("activeStageTitle");
  const hubGuide = $("hubGuide");
  const cards = [...document.querySelectorAll(".tool-card")];

  const clearNavActive = () => {
    [...nav.querySelectorAll("a")].forEach((a) => a.classList.remove("active"));
  };

  const enterHub = (replaceHash = true) => {
    document.body.classList.remove("stage-mode");
    document.body.classList.add("hub-mode");
    stageHeader.classList.add("hidden");
    hubGuide.classList.remove("hidden");
    cards.forEach((card) => card.classList.remove("active-stage"));
    clearNavActive();
    if (replaceHash) history.replaceState(null, "", "#hub");
  };

  const enterStage = (toolId, replaceHash = true) => {
    const target = cards.find((c) => c.id === toolId);
    if (!target) {
      enterHub(replaceHash);
      return;
    }
    document.body.classList.remove("hub-mode");
    document.body.classList.add("stage-mode");
    stageHeader.classList.remove("hidden");
    hubGuide.classList.add("hidden");
    cards.forEach((card) => card.classList.toggle("active-stage", card.id === toolId));
    clearNavActive();
    const activeLink = nav.querySelector(`a[data-tool-id="${toolId}"]`);
    if (activeLink) activeLink.classList.add("active");
    stageTitle.textContent = target.querySelector("h2").textContent;
    if (replaceHash) history.replaceState(null, "", `#${toolId}`);
  };

  nav.addEventListener("click", (e) => {
    const link = e.target.closest("a[data-tool-id]");
    if (!link) return;
    e.preventDefault();
    enterStage(link.dataset.toolId);
  });

  $("backToHub").addEventListener("click", () => enterHub());

  const initialHash = window.location.hash.replace("#", "").trim();
  if (!initialHash || initialHash === "hub") enterHub(false);
  else enterStage(initialHash, false);

  window.addEventListener("hashchange", () => {
    const hash = window.location.hash.replace("#", "").trim();
    if (!hash || hash === "hub") enterHub(false);
    else enterStage(hash, false);
  });
};

const setupPdfToImage = () => {
  if (!$("runPdfToImage")) return;
  setIconButton("undoPdfToImageDelete", "undo");
  setIconButton("runPdfToImageFromPreview", "download");
  setupPdfToImagePreviewDnD();

  $("pdfToImageFile").addEventListener("change", () => {
    const files = [...$("pdfToImageFile").files];
    setStatus("pdfToImageStatus", files.length ? "PDF를 읽는 중..." : "");
    renderPdfToImagePreview(files);
  });

  $("undoPdfToImageDelete")?.addEventListener("click", () => {
    const last = pdfToImageState.deletedStack.pop();
    if (!last) {
      setStatus("pdfToImageStatus", "되돌릴 삭제 내역이 없습니다.");
      return;
    }
    const insertAt = Math.max(0, Math.min(last.index, pdfToImageState.pages.length));
    pdfToImageState.pages.splice(insertAt, 0, last.page);
    renderPdfToImageGrid();
    setStatus("pdfToImageStatus", `페이지 ${last.page.pageNo} 복원 완료`);
  });

  $("runPdfToImageFromPreview")?.addEventListener("click", () => {
    $("runPdfToImage").click();
  });

  $("runPdfToImage").addEventListener("click", async () => {
    const files = [...$("pdfToImageFile").files];
    const format = $("pdfToImageFormat").value;
    const dpi = Number($("pdfToImageDpi").value || 200);
    const quality = Math.min(1, Math.max(0.01, Number($("pdfToImageQuality").value || 80) / 100));
    const scale = Math.max(1, dpi / 96);
    if (!files.length) {
      setStatus("pdfToImageStatus", "PDF 파일을 선택해주세요.");
      return;
    }

    startOperation("pdfToImage", "PDF 로딩 중...");
    try {
      const pageFilter = ($("pdfToImagePages").value || "").trim();
      const docs = [];
      for (let fi = 0; fi < files.length; fi += 1) {
        checkCancelled("pdfToImage");
        const buffer = await readAsArrayBuffer(files[fi]);
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        docs.push(pdf);
      }

      const finalPages = pdfToImageState.pages.length
        ? [...pdfToImageState.pages]
        : docs.flatMap((pdf, fi) =>
            Array.from({ length: pdf.numPages }, (_, i) => ({
              id: `${fi}:${i + 1}`,
              fileIndex: fi,
              fileLabel: files[fi].name,
              pageNo: i + 1,
            }))
          );

      const filteredFinalPages =
        pageFilter && docs.length === 1
          ? (() => {
              const selected = new Set(parsePageTokens(pageFilter, docs[0].numPages));
              return finalPages.filter((p) => selected.has(p.pageNo));
            })()
          : finalPages;

      if (!filteredFinalPages.length) {
        throw new Error("변환할 페이지가 없습니다. 페이지 필터/삭제 상태를 확인해주세요.");
      }
      const outputs = [];

      for (let i = 0; i < filteredFinalPages.length; i += 1) {
        checkCancelled("pdfToImage");
        const pageMeta = filteredFinalPages[i];
        const pageNo = pageMeta.pageNo;
        setStatus("pdfToImageStatus", `페이지 변환 중 (${i + 1}/${filteredFinalPages.length})`);
        const page = await docs[pageMeta.fileIndex].getPage(pageNo);
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement("canvas");
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({
          canvasContext: canvas.getContext("2d"),
          viewport,
        }).promise;
        const mime = format === "png" ? "image/png" : format === "webp" ? "image/webp" : "image/jpeg";
        const dataUrl = canvas.toDataURL(mime, quality);
        const realMime = dataUrl.startsWith("data:image/png") && mime !== "image/png" ? "image/png" : mime;
        const ext = realMime === "image/jpeg" ? "jpg" : realMime === "image/webp" ? "webp" : "png";
        const outBytes = dataUrlToUint8Array(dataUrl);
        const safeName = pageMeta.fileLabel.replace(/\.[^.]+$/, "").replace(/[\\/:*?"<>|]/g, "_");
        outputs.push({
          name: `${safeName}_p${pageNo}.${ext}`,
          blob: new Blob([outBytes], { type: realMime }),
        });
        updateProgress("pdfToImage", i + 1, filteredFinalPages.length);
      }

      checkCancelled("pdfToImage");
      if (outputs.length === 1) {
        downloadBlob(outputs[0].blob, outputs[0].name);
      } else {
        const zip = new JSZip();
        outputs.forEach((out) => zip.file(out.name, out.blob));
        const zipBlob = await zip.generateAsync({ type: "blob" });
        downloadBlob(zipBlob, "pdf-to-images.zip");
      }
      updateProgress("pdfToImage", 100, 100);
      const totalBytes = outputs.reduce((sum, out) => sum + (out.blob?.size || 0), 0);
      endOperation(
        "pdfToImage",
        outputs.length === 1
          ? `완료: 1페이지 단건 다운로드 (${formatBytes(totalBytes)})`
          : `완료: ${outputs.length}페이지 (ZIP 다운로드, ${formatBytes(totalBytes)})`
      );
    } catch (err) {
      handleOperationError("pdfToImage", err);
    }
  });
};

const setupLoginModal = () => {
  const modal = $("loginModal");
  const openBtn = $("loginBtn");
  const closeBtn = $("closeLoginModal");
  const cancelBtn = $("loginCancelBtn");
  const submitBtn = $("loginSubmitBtn");
  const status = $("loginModalStatus");
  if (!modal || !openBtn || !closeBtn || !cancelBtn || !submitBtn) return;

  const openModal = () => {
    modal.classList.remove("hidden");
    modal.setAttribute("aria-hidden", "false");
    status.textContent = "";
    setTimeout(() => $("loginEmail")?.focus(), 40);
  };

  const closeModal = () => {
    modal.classList.add("hidden");
    modal.setAttribute("aria-hidden", "true");
  };

  setIconButton("loginBtnIcon", "person");
  setIconButton("closeLoginModal", "x");

  openBtn.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  cancelBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", (e) => {
    if (e.target.closest("[data-close-login='true']")) closeModal();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !modal.classList.contains("hidden")) closeModal();
  });
  submitBtn.addEventListener("click", () => {
    const email = $("loginEmail")?.value?.trim() || "";
    const pw = $("loginPassword")?.value?.trim() || "";
    if (!email || !pw) {
      status.textContent = "이메일과 비밀번호를 입력해주세요.";
      return;
    }
    status.textContent = "로그인 확인 중...";
    setTimeout(() => {
      status.textContent = "로그인 성공(데모): 연동 기능 준비 중입니다.";
    }, 450);
  });
};

const toolFileState = {
  imageToPdf: [],
  mergePdf: [],
  pdfCompress: [],
  dwgToPdf: [],
  resize: [],
  format: [],
  rename: [],
};

const pdfThumbCache = new Map();
const pdfPageCountCache = new Map();
const PDF_THUMB_CACHE_MAX = 120;

const getFileCacheKey = (file) => `${file.name}__${file.size}__${file.lastModified}`;

const getCachedPdfThumb = (key) => {
  if (!pdfThumbCache.has(key)) return null;
  const value = pdfThumbCache.get(key);
  // LRU touch
  pdfThumbCache.delete(key);
  pdfThumbCache.set(key, value);
  return value;
};

const setCachedPdfThumb = (key, value) => {
  if (pdfThumbCache.has(key)) pdfThumbCache.delete(key);
  pdfThumbCache.set(key, value);
  if (pdfThumbCache.size <= PDF_THUMB_CACHE_MAX) return;
  const oldest = pdfThumbCache.keys().next().value;
  if (oldest) pdfThumbCache.delete(oldest);
};

const getCachedPdfPageCount = (key) => pdfPageCountCache.get(key) || null;
const setCachedPdfPageCount = (key, value) => {
  pdfPageCountCache.set(key, value);
};

const loadImageFromAnyFile = async (file) => {
  if (!isHeicLikeFile(file)) return loadImageFromFile(file);
  if (typeof heic2any !== "function") {
    throw new Error("HEIC 변환 라이브러리(heic2any)를 불러오지 못했습니다.");
  }
  const converted = await heic2any({
    blob: file,
    toType: "image/png",
    quality: 0.92,
  });
  const blob = Array.isArray(converted) ? converted[0] : converted;
  const dataUrl = await readAsDataURL(blob);
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ img, dataUrl });
    img.onerror = reject;
    img.src = dataUrl;
  });
};

const loadThumbFromImageFile = async (file, maxSize = 130) => {
  const { img } = await loadImageFromAnyFile(file);
  const ratio = img.width / img.height;
  const width = ratio >= 1 ? maxSize : Math.round(maxSize * ratio);
  const height = ratio >= 1 ? Math.round(maxSize / ratio) : maxSize;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  canvas.getContext("2d").drawImage(img, 0, 0, width, height);
  return canvas.toDataURL("image/png");
};

const loadPdfFrontThumb = async (file, maxWidth = 130) => {
  const cacheKey = `${getFileCacheKey(file)}__${maxWidth}`;
  const cached = getCachedPdfThumb(cacheKey);
  const cachedPageCount = getCachedPdfPageCount(cacheKey);
  if (cached && cachedPageCount) {
    return { thumb: cached, pageCount: cachedPageCount };
  }
  const buffer = await readAsArrayBuffer(file);
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pageCount = pdf.numPages;
  const page = await pdf.getPage(1);
  const base = page.getViewport({ scale: 1 });
  const scale = maxWidth / base.width;
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
  const thumb = canvas.toDataURL("image/png");
  setCachedPdfThumb(cacheKey, thumb);
  setCachedPdfPageCount(cacheKey, pageCount);
  return { thumb, pageCount };
};

const removeFileAtIndex = (stateKey, inputId, index) => {
  const files = toolFileState[stateKey];
  if (!files?.length) return;
  files.splice(index, 1);
  syncFilesToInput(inputId, files);
};

const renderImageThumbPreview = async (previewId, stateKey, inputId, reorderable = false) => {
  const grid = $(previewId);
  if (!grid) return;
  grid.innerHTML = "";
  const files = toolFileState[stateKey];
  if (!files.length) return;
  beginGlobalBusy("이미지 미리보기를 준비 중입니다...");
  try {
    for (let i = 0; i < files.length; i += 1) {
      setGlobalBusyMessage(`이미지 미리보기 생성 중 (${i + 1}/${files.length})`);
      const item = document.createElement("div");
      item.className = "thumb-item";
      item.draggable = reorderable;
      item.dataset.idx = String(i);
      item.innerHTML = `<button class="thumb-delete" type="button" title="파일 제거" aria-label="파일 제거">${ICONS.trash3}</button><div class="thumb-label">${files[i].name}</div>`;
      try {
        const dataUrl = await loadThumbFromImageFile(files[i]);
        const img = document.createElement("img");
        img.src = dataUrl;
        img.alt = files[i].name;
        img.draggable = false;
        img.style.width = "100%";
        img.style.border = "1px solid #d4e2f1";
        img.style.borderRadius = "6px";
        item.prepend(img);
      } catch {
        const stub = document.createElement("div");
        stub.className = "thumb-label";
        stub.textContent = "미리보기 불가";
        item.prepend(stub);
      }
      grid.appendChild(item);
    }
    grid.onclick = (e) => {
      const del = e.target.closest(".thumb-delete");
      if (!del) return;
      const cell = e.target.closest(".thumb-item");
      if (!cell) return;
      removeFileAtIndex(stateKey, inputId, Number(cell.dataset.idx));
    };
  } finally {
    endGlobalBusy();
  }
  if (!reorderable) return;
  let dragIdx = -1;
  const placeholder = document.createElement("div");
  placeholder.className = "drag-placeholder";
  grid.ondragstart = (e) => {
    const cell = e.target.closest(".thumb-item");
    if (!cell) return;
    dragIdx = Number(cell.dataset.idx);
    cell.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  };
  grid.ondragend = () => {
    dragIdx = -1;
    if (placeholder.parentElement) placeholder.parentElement.removeChild(placeholder);
    grid.querySelectorAll(".thumb-item.dragging").forEach((el) => el.classList.remove("dragging"));
  };
  grid.ondragover = (e) => {
    if (dragIdx < 0) return;
    e.preventDefault();
    const intent = getPdfToImageDragAfterElement(grid, e.clientX, e.clientY);
    if (!intent.afterEl) grid.appendChild(placeholder);
    else grid.insertBefore(placeholder, intent.afterEl);
  };
  grid.ondrop = (e) => {
    if (dragIdx < 0) return;
    e.preventDefault();
    const moving = toolFileState[stateKey][dragIdx];
    const filtered = toolFileState[stateKey].filter((_, i) => i !== dragIdx);
    const next = placeholder.nextElementSibling?.closest?.(".thumb-item");
    if (next) {
      const nextIdx = Number(next.dataset.idx);
      filtered.splice(nextIdx, 0, moving);
    } else {
      filtered.push(moving);
    }
    toolFileState[stateKey] = filtered;
    syncFilesToInput(inputId, filtered);
  };
};

const renderMergePdfPreview = async (previewId, stateKey, inputId) => {
  const grid = $(previewId);
  if (!grid) return;
  grid.innerHTML = "";
  const files = toolFileState[stateKey];
  if (!files.length) return;
  beginGlobalBusy("PDF 썸네일을 준비 중입니다...");
  try {
    for (let idx = 0; idx < files.length; idx += 1) {
      setGlobalBusyMessage(`PDF 목록 썸네일 생성 중 (${idx + 1}/${files.length})`);
      const file = files[idx];
      const item = document.createElement("div");
      item.className = "thumb-item";
      item.draggable = true;
      item.dataset.idx = String(idx);
      item.innerHTML = `<button class="thumb-delete" type="button" title="파일 제거" aria-label="파일 제거">${ICONS.trash3}</button><div class="thumb-label">${idx + 1}. ${file.name}</div>`;
      try {
        const { thumb, pageCount } = await loadPdfFrontThumb(file, 130);
        const img = document.createElement("img");
        img.src = thumb;
        img.alt = `${file.name} first page`;
        img.draggable = false;
        img.style.width = "100%";
        img.style.border = "1px solid #d4e2f1";
        img.style.borderRadius = "6px";
        item.prepend(img);
        const pageTag = document.createElement("div");
        pageTag.className = "thumb-label";
        pageTag.textContent = `${pageCount}페이지`;
        item.appendChild(pageTag);
      } catch {
        const stub = document.createElement("div");
        stub.className = "thumb-label";
        stub.textContent = "미리보기 불가";
        item.prepend(stub);
      }
      grid.appendChild(item);
    }
  } finally {
    endGlobalBusy();
  }
  grid.onclick = (e) => {
    const del = e.target.closest(".thumb-delete");
    if (!del) return;
    const cell = e.target.closest(".thumb-item");
    if (!cell) return;
    removeFileAtIndex(stateKey, inputId, Number(cell.dataset.idx));
  };
  let dragIdx = -1;
  const placeholder = document.createElement("div");
  placeholder.className = "drag-placeholder";
  grid.ondragstart = (e) => {
    const cell = e.target.closest(".thumb-item");
    if (!cell) return;
    dragIdx = Number(cell.dataset.idx);
    cell.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  };
  grid.ondragend = () => {
    dragIdx = -1;
    if (placeholder.parentElement) placeholder.parentElement.removeChild(placeholder);
    grid.querySelectorAll(".thumb-item.dragging").forEach((el) => el.classList.remove("dragging"));
  };
  grid.ondragover = (e) => {
    if (dragIdx < 0) return;
    e.preventDefault();
    const intent = getPdfToImageDragAfterElement(grid, e.clientX, e.clientY);
    if (!intent.afterEl) grid.appendChild(placeholder);
    else grid.insertBefore(placeholder, intent.afterEl);
  };
  grid.ondrop = (e) => {
    if (dragIdx < 0) return;
    e.preventDefault();
    const moving = toolFileState[stateKey][dragIdx];
    const filtered = toolFileState[stateKey].filter((_, i) => i !== dragIdx);
    const next = placeholder.nextElementSibling?.closest?.(".thumb-item");
    if (next) {
      const nextIdx = Number(next.dataset.idx);
      filtered.splice(nextIdx, 0, moving);
    } else {
      filtered.push(moving);
    }
    toolFileState[stateKey] = filtered;
    syncFilesToInput(inputId, filtered);
  };
};

const setupImageToPdf = () => {
  if (!$("runImageToPdf")) return;
  setIconButton("runImageToPdfFromPreview", "download");
  $("imageToPdfFiles")?.addEventListener("change", async () => {
    toolFileState.imageToPdf = [...$("imageToPdfFiles").files];
    await renderImageThumbPreview("imageToPdfPreview", "imageToPdf", "imageToPdfFiles", true);
  });
  $("runImageToPdfFromPreview")?.addEventListener("click", () => $("runImageToPdf").click());

  $("runImageToPdf").addEventListener("click", async () => {
    const files = [...$("imageToPdfFiles").files];
    if (!files.length) {
      setStatus("imageToPdfStatus", "이미지 파일을 선택해주세요.");
      return;
    }

    startOperation("imageToPdf", "PDF 생성 준비 중...");
    try {
      const pdfDoc = await PDFLib.PDFDocument.create();
      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("imageToPdf");
        setStatus("imageToPdfStatus", `이미지 처리 중 (${i + 1}/${files.length})`);
        const file = files[i];
        let embed;
        if (file.type.includes("png")) {
          embed = await pdfDoc.embedPng(await readAsArrayBuffer(file));
        } else if (file.type.includes("jpeg") || file.type.includes("jpg")) {
          embed = await pdfDoc.embedJpg(await readAsArrayBuffer(file));
        } else {
          const { img } = await loadImageFromFile(file);
          const canvas = document.createElement("canvas");
          canvas.width = img.width;
          canvas.height = img.height;
          canvas.getContext("2d").drawImage(img, 0, 0);
          embed = await pdfDoc.embedPng(
            dataUrlToUint8Array(canvas.toDataURL("image/png"))
          );
        }
        const page = pdfDoc.addPage([embed.width, embed.height]);
        page.drawImage(embed, {
          x: 0,
          y: 0,
          width: embed.width,
          height: embed.height,
        });
        updateProgress("imageToPdf", i + 1, files.length);
      }
      checkCancelled("imageToPdf");
      const out = await pdfDoc.save();
      const blob = new Blob([out], { type: "application/pdf" });
      downloadBlob(blob, "images-to-pdf.pdf");
      updateProgress("imageToPdf", 100, 100);
      endOperation("imageToPdf", `완료: ${files.length}개 이미지 병합`);
    } catch (err) {
      handleOperationError("imageToPdf", err);
    }
  });
};

const arrangeState = {
  file: null,
  pages: [],
  deletedStack: [],
  reorderOrder: [],
  splitBuckets: [],
  nextBucketId: 1,
  selection: {
    source: new Set(),
    reorder: new Set(),
  },
  anchorIndex: {
    source: null,
    reorder: null,
  },
  dragCtx: null,
  placeholder: null,
};

const ensureArrangePlaceholder = () => {
  if (arrangeState.placeholder) return arrangeState.placeholder;
  const ph = document.createElement("div");
  ph.className = "drag-placeholder";
  arrangeState.placeholder = ph;
  return ph;
};

const getArrangePageByNo = (pageNo) => arrangeState.pages.find((p) => p.pageNo === pageNo);
const isPageDeleted = (pageNo) => !!getArrangePageByNo(pageNo)?.deleted;
const getAvailablePageNos = () =>
  arrangeState.pages.filter((p) => !p.deleted).map((p) => p.pageNo);

const cleanArrangeState = () => {
  const available = new Set(getAvailablePageNos());
  arrangeState.reorderOrder = arrangeState.reorderOrder.filter((n) => available.has(n));
  arrangeState.splitBuckets.forEach((bucket) => {
    bucket.pages = bucket.pages.filter((n) => available.has(n));
  });
  arrangeState.selection.source = new Set(
    [...arrangeState.selection.source].filter((n) => available.has(n))
  );
  arrangeState.selection.reorder = new Set(
    [...arrangeState.selection.reorder].filter((n) => available.has(n))
  );
};

const updateArrangeOrderText = () => {
  $("arrangeOrderText").textContent = arrangeState.reorderOrder.length
    ? arrangeState.reorderOrder.join(", ")
    : "-";
};

const createThumbNode = (pageNo, pane) => {
  const page = getArrangePageByNo(pageNo);
  if (!page || page.deleted) return null;
  const item = document.createElement("div");
  item.className = "thumb-item";
  item.draggable = true;
  item.dataset.page = String(pageNo);
  item.dataset.pane = pane;
  item.innerHTML = `<button class="thumb-delete" type="button" title="페이지 삭제" aria-label="페이지 삭제">${ICONS.trash3}</button><div class="thumb-label">p.${pageNo}</div>`;
  const img = document.createElement("img");
  img.src = page.thumbDataUrl;
  img.alt = `page-${pageNo}`;
  img.draggable = false;
  img.style.width = "100%";
  img.style.border = "1px solid #d4e2f1";
  img.style.borderRadius = "6px";
  item.prepend(img);
  const selectedSet =
    pane === "source"
      ? arrangeState.selection.source
      : pane === "reorder"
        ? arrangeState.selection.reorder
        : null;
  if (selectedSet?.has(pageNo)) item.classList.add("selected-range");
  return item;
};

const renderSourceGrid = () => {
  const grid = $("pdfThumbGrid");
  if (!grid) return;
  grid.innerHTML = "";
  getAvailablePageNos().forEach((pageNo) => {
    const node = createThumbNode(pageNo, "source");
    if (node) grid.appendChild(node);
  });
};

const renderReorderGrid = () => {
  const grid = $("reorderThumbGrid");
  if (!grid) return;
  grid.innerHTML = "";
  arrangeState.reorderOrder.forEach((pageNo) => {
    if (isPageDeleted(pageNo)) return;
    const node = createThumbNode(pageNo, "reorder");
    if (node) grid.appendChild(node);
  });
  updateArrangeOrderText();
};

const renderSplitBuckets = () => {
  const wrap = $("splitBucketWrap");
  if (!wrap) return;
  wrap.innerHTML = "";
  arrangeState.splitBuckets.forEach((bucket, idx) => {
    const panel = document.createElement("div");
    panel.className = "split-bucket";
    panel.dataset.bucketId = String(bucket.id);
    panel.innerHTML = `
      <div class="split-bucket-head">
        <span class="split-bucket-title">분할 ${idx + 1}</span>
        <div class="split-bucket-actions">
          <button type="button" data-bucket-action="clear">비우기</button>
          <button type="button" data-bucket-action="remove">삭제</button>
        </div>
      </div>
      <div class="thumb-grid split-bucket-grid drop-target-grid" data-bucket-grid="${bucket.id}"></div>
    `;
    const grid = panel.querySelector(".split-bucket-grid");
    bucket.pages.forEach((pageNo) => {
      if (isPageDeleted(pageNo)) return;
      const node = createThumbNode(pageNo, "split");
      if (node) {
        node.dataset.bucketId = String(bucket.id);
        grid.appendChild(node);
      }
    });
    wrap.appendChild(panel);
  });
  if (!arrangeState.splitBuckets.length) {
    wrap.innerHTML = `<div class="split-bucket"><div class="split-bucket-title">분할 박스가 없습니다. "분할 박스 추가"를 눌러주세요.</div></div>`;
  }
};

const rerenderArrangeWorkspace = () => {
  cleanArrangeState();
  renderSourceGrid();
  renderReorderGrid();
  renderSplitBuckets();
};

const removeArrangePage = (pageNo) => {
  const page = getArrangePageByNo(pageNo);
  if (!page) return;
  const idx = arrangeState.pages.findIndex((p) => p.pageNo === pageNo);
  if (idx >= 0) {
    arrangeState.deletedStack.push({
      pageNo,
      pageSnapshot: { ...arrangeState.pages[idx] },
      sourceIndex: idx,
    });
  }
  page.deleted = true;
  cleanArrangeState();
  rerenderArrangeWorkspace();
};

const undoArrangeDelete = () => {
  const last = arrangeState.deletedStack.pop();
  if (!last) return null;
  const target = getArrangePageByNo(last.pageNo);
  if (target) {
    target.deleted = false;
  } else {
    const insertAt = Math.max(0, Math.min(last.sourceIndex, arrangeState.pages.length));
    arrangeState.pages.splice(insertAt, 0, last.pageSnapshot);
  }
  cleanArrangeState();
  rerenderArrangeWorkspace();
  return last.pageNo;
};

const setPaneSelection = (pane, selectedNos) => {
  if (pane === "source") arrangeState.selection.source = new Set(selectedNos);
  if (pane === "reorder") arrangeState.selection.reorder = new Set(selectedNos);
};

const applyShiftSelection = (pane, clickedPageNo, shiftKey) => {
  const grid = pane === "source" ? $("pdfThumbGrid") : $("reorderThumbGrid");
  if (!grid) return;
  const items = [...grid.querySelectorAll(".thumb-item")];
  const indexMap = items.map((el) => Number(el.dataset.page));
  const clickedIndex = indexMap.indexOf(clickedPageNo);
  if (clickedIndex < 0) return;
  if (!shiftKey || arrangeState.anchorIndex[pane] === null) {
    arrangeState.anchorIndex[pane] = clickedIndex;
    setPaneSelection(pane, [clickedPageNo]);
    rerenderArrangeWorkspace();
    return;
  }
  const start = Math.min(arrangeState.anchorIndex[pane], clickedIndex);
  const end = Math.max(arrangeState.anchorIndex[pane], clickedIndex);
  setPaneSelection(pane, indexMap.slice(start, end + 1));
  rerenderArrangeWorkspace();
};

const getDragAfterElement = (container, x, y) => {
  return getPdfToImageDragAfterElement(container, x, y).afterEl;
};

const placePlaceholderInGrid = (grid, x, y) => {
  const placeholder = ensureArrangePlaceholder();
  const intent = getPdfToImageDragAfterElement(grid, x, y);
  const afterEl = intent.afterEl;
  if (!afterEl) grid.appendChild(placeholder);
  else grid.insertBefore(placeholder, afterEl);
};

const removePlaceholder = () => {
  const ph = arrangeState.placeholder;
  if (ph && ph.parentElement) ph.parentElement.removeChild(ph);
};

const setupArrangeDnD = () => {
  const sourceGrid = $("pdfThumbGrid");
  const reorderGrid = $("reorderThumbGrid");
  const splitWrap = $("splitBucketWrap");
  if (!sourceGrid || !reorderGrid || !splitWrap) return;

  const onThumbClick = (e) => {
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    if (e.target.closest(".thumb-delete")) return;
    const pane = item.dataset.pane;
    const pageNo = Number(item.dataset.page);
    if (pane === "source" || pane === "reorder") {
      applyShiftSelection(pane, pageNo, e.shiftKey);
    }
  };

  const onDeleteClick = (e) => {
    const btn = e.target.closest(".thumb-delete");
    if (!btn) return;
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    const pageNo = Number(item.dataset.page);
    removeArrangePage(pageNo);
    setStatus("arrangePdfStatus", `페이지 ${pageNo} 삭제됨`);
  };

  const handleDragStart = (e) => {
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    const pane = item.dataset.pane;
    const pageNo = Number(item.dataset.page);
    if (pane === "source" && !arrangeState.selection.source.has(pageNo)) {
      setPaneSelection("source", [pageNo]);
    }
    if (pane === "reorder" && !arrangeState.selection.reorder.has(pageNo)) {
      setPaneSelection("reorder", [pageNo]);
    }

    document.querySelectorAll("#pdfThumbGrid .thumb-item, #reorderThumbGrid .thumb-item").forEach((el) => {
      const elPane = el.dataset.pane;
      const elPage = Number(el.dataset.page);
      const selected =
        (elPane === "source" && arrangeState.selection.source.has(elPage)) ||
        (elPane === "reorder" && arrangeState.selection.reorder.has(elPage));
      el.classList.toggle("selected-range", selected);
    });

    const selected =
      pane === "source"
        ? [...arrangeState.selection.source]
        : pane === "reorder"
          ? [...arrangeState.selection.reorder]
          : [pageNo];
    arrangeState.dragCtx = { pane, pageNos: selected, bucketId: item.dataset.bucketId || null };
    item.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  };

  const handleDragEnd = () => {
    removePlaceholder();
    arrangeState.dragCtx = null;
    document.querySelectorAll(".thumb-item.dragging").forEach((el) => el.classList.remove("dragging"));
  };

  const applyDropToReorder = () => {
    const ph = arrangeState.placeholder;
    const ctx = arrangeState.dragCtx;
    if (!ctx || !ph || !ph.parentElement) return;
    const moving = ctx.pageNos.filter((n) => !isPageDeleted(n));
    if (!moving.length) return;
    const filtered = arrangeState.reorderOrder.filter((n) => !moving.includes(n));
    const nextThumb = ph.nextElementSibling?.closest?.(".thumb-item");
    if (nextThumb) {
      const nextPage = Number(nextThumb.dataset.page);
      const idx = filtered.indexOf(nextPage);
      if (idx >= 0) filtered.splice(idx, 0, ...moving);
      else filtered.push(...moving);
    } else {
      filtered.push(...moving);
    }
    arrangeState.reorderOrder = filtered;
    rerenderArrangeWorkspace();
  };

  const applyDropToSplitBucket = (bucketId) => {
    const ctx = arrangeState.dragCtx;
    const ph = arrangeState.placeholder;
    if (!ctx || !ph || !ph.parentElement) return;
    const bucket = arrangeState.splitBuckets.find((b) => b.id === bucketId);
    if (!bucket) return;
    const moving = ctx.pageNos.filter((n) => !isPageDeleted(n));
    if (!moving.length) return;
    const filtered = bucket.pages.filter((n) => !moving.includes(n));
    const nextThumb = ph.nextElementSibling?.closest?.(".thumb-item");
    if (nextThumb) {
      const nextPage = Number(nextThumb.dataset.page);
      const idx = filtered.indexOf(nextPage);
      if (idx >= 0) filtered.splice(idx, 0, ...moving);
      else filtered.push(...moving);
    } else {
      filtered.push(...moving);
    }
    bucket.pages = [...new Set(filtered)];
    rerenderArrangeWorkspace();
  };

  [sourceGrid, reorderGrid].forEach((grid) => {
    grid.addEventListener("click", onThumbClick);
    grid.addEventListener("click", onDeleteClick);
    grid.addEventListener("dragstart", handleDragStart);
    grid.addEventListener("dragend", handleDragEnd);
  });

  reorderGrid.addEventListener("dragover", (e) => {
    if (!arrangeState.dragCtx) return;
    e.preventDefault();
    placePlaceholderInGrid(reorderGrid, e.clientX, e.clientY);
  });
  reorderGrid.addEventListener("drop", (e) => {
    if (!arrangeState.dragCtx) return;
    e.preventDefault();
    applyDropToReorder();
    removePlaceholder();
  });

  splitWrap.addEventListener("click", (e) => {
    const actionBtn = e.target.closest("[data-bucket-action]");
    if (!actionBtn) return;
    const bucketEl = e.target.closest(".split-bucket");
    if (!bucketEl) return;
    const bucketId = Number(bucketEl.dataset.bucketId);
    const bucket = arrangeState.splitBuckets.find((b) => b.id === bucketId);
    if (!bucket) return;
    const action = actionBtn.dataset.bucketAction;
    if (action === "clear") {
      bucket.pages = [];
    } else if (action === "remove") {
      arrangeState.splitBuckets = arrangeState.splitBuckets.filter((b) => b.id !== bucketId);
    }
    rerenderArrangeWorkspace();
  });

  splitWrap.addEventListener("click", onDeleteClick);
  splitWrap.addEventListener("dragstart", handleDragStart);
  splitWrap.addEventListener("dragend", handleDragEnd);
  splitWrap.addEventListener("dragover", (e) => {
    const grid = e.target.closest(".split-bucket-grid");
    if (!arrangeState.dragCtx || !grid) return;
    e.preventDefault();
    placePlaceholderInGrid(grid, e.clientX, e.clientY);
  });
  splitWrap.addEventListener("drop", (e) => {
    const grid = e.target.closest(".split-bucket-grid");
    if (!arrangeState.dragCtx || !grid) return;
    e.preventDefault();
    const bucketId = Number(grid.dataset.bucketGrid);
    applyDropToSplitBucket(bucketId);
    removePlaceholder();
  });
};

const renderArrangeThumbs = async (file) => {
  startOperation("arrange", "썸네일 자동 불러오는 중...");
  try {
    const buffer = await readAsArrayBuffer(file);
    checkCancelled("arrange");
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
    arrangeState.file = file;
    arrangeState.pages = [];
    arrangeState.deletedStack = [];
    arrangeState.reorderOrder = [];
    arrangeState.splitBuckets = [];
    arrangeState.nextBucketId = 1;
    arrangeState.selection.source = new Set();
    arrangeState.selection.reorder = new Set();
    arrangeState.anchorIndex.source = null;
    arrangeState.anchorIndex.reorder = null;

    for (let i = 1; i <= pdf.numPages; i += 1) {
      checkCancelled("arrange");
      setStatus("arrangePdfStatus", `썸네일 렌더링 중 (${i}/${pdf.numPages})`);
      const page = await pdf.getPage(i);
      const baseViewport = page.getViewport({ scale: 1 });
      const targetWidth = 130;
      const scale = targetWidth / baseViewport.width;
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({
        canvasContext: canvas.getContext("2d"),
        viewport,
      }).promise;
      const thumbDataUrl = canvas.toDataURL("image/png");
      arrangeState.pages.push({ pageNo: i, thumbDataUrl, deleted: false });
      arrangeState.reorderOrder.push(i);
      updateProgress("arrange", i, pdf.numPages);
    }

    if (!arrangeState.splitBuckets.length) {
      arrangeState.splitBuckets.push({ id: arrangeState.nextBucketId++, pages: [] });
    }
    rerenderArrangeWorkspace();
    updateProgress("arrange", 100, 100);
    endOperation("arrange", `완료: ${pdf.numPages}개 페이지 자동 로드됨`);
  } catch (err) {
    handleOperationError("arrange", err);
  }
};

const setupPdfArrange = () => {
  if (!$("runReorderPdf") || !$("pdfThumbGrid")) return;
  setIconButton("undoArrangeDelete", "undo");
  setIconButton("runReorderPdf", "download");
  setIconButton("runSplitPdf", "download");
  setupArrangeDnD();
  rerenderArrangeWorkspace();

  $("arrangePdfFile").addEventListener("change", async () => {
    const file = $("arrangePdfFile").files[0];
    if (!file) return;
    await renderArrangeThumbs(file);
  });

  $("undoArrangeDelete")?.addEventListener("click", () => {
    const restored = undoArrangeDelete();
    if (!restored) {
      setStatus("arrangePdfStatus", "되돌릴 삭제 내역이 없습니다.");
      return;
    }
    setStatus("arrangePdfStatus", `페이지 ${restored} 복원 완료`);
  });

  $("addSplitBucket").addEventListener("click", () => {
    arrangeState.splitBuckets.push({ id: arrangeState.nextBucketId++, pages: [] });
    rerenderArrangeWorkspace();
  });

  $("clearSplitBuckets").addEventListener("click", () => {
    arrangeState.splitBuckets = [{ id: arrangeState.nextBucketId++, pages: [] }];
    rerenderArrangeWorkspace();
  });

  $("runReorderPdf").addEventListener("click", async () => {
    const file = $("arrangePdfFile").files[0];
    if (!file) {
      setStatus("arrangePdfStatus", "PDF 파일을 먼저 선택해주세요.");
      return;
    }
    if (!arrangeState.pages.length) await renderArrangeThumbs(file);

    const finalOrder = arrangeState.reorderOrder.filter((n) => !isPageDeleted(n));
    if (!finalOrder.length) {
      setStatus("arrangePdfStatus", "저장할 페이지가 없습니다.");
      return;
    }

    startOperation("arrange", "정렬 순서로 PDF 생성 중...");
    try {
      const src = await PDFLib.PDFDocument.load(await readAsArrayBuffer(file));
      const out = await PDFLib.PDFDocument.create();
      const copied = await out.copyPages(
        src,
        finalOrder.map((n) => n - 1)
      );
      copied.forEach((p, idx) => {
        checkCancelled("arrange");
        out.addPage(p);
        updateProgress("arrange", idx + 1, copied.length);
      });
      const result = await out.save();
      downloadBlob(new Blob([result], { type: "application/pdf" }), "reordered.pdf");
      updateProgress("arrange", 100, 100);
      endOperation("arrange", "완료: 순서 변경 PDF 저장");
    } catch (err) {
      handleOperationError("arrange", err);
    }
  });

  $("runSplitPdf").addEventListener("click", async () => {
    const file = $("arrangePdfFile").files[0];
    const splitText = $("splitInput").value.trim();
    if (!file) {
      setStatus("arrangePdfStatus", "PDF 파일을 선택해주세요.");
      return;
    }
    if (!arrangeState.pages.length) await renderArrangeThumbs(file);

    startOperation("arrange", "PDF 분할 처리 중...");
    try {
      const src = await PDFLib.PDFDocument.load(await readAsArrayBuffer(file));
      let groups = arrangeState.splitBuckets
        .map((b) => b.pages.filter((n) => !isPageDeleted(n)))
        .filter((g) => g.length);

      if (!groups.length) {
        if (!splitText) throw new Error("분할 박스가 비어있습니다. 페이지를 끌어다 놓거나 텍스트 분할값을 입력해주세요.");
        groups = parseSplitGroups(splitText, src.getPageCount()).filter((g) => g.length);
      }
      if (!groups.length) throw new Error("유효한 분할 대상이 없습니다.");

      const outputs = [];
      for (let i = 0; i < groups.length; i += 1) {
        checkCancelled("arrange");
        const out = await PDFLib.PDFDocument.create();
        const pages = await out.copyPages(
          src,
          groups[i].map((n) => n - 1)
        );
        pages.forEach((p) => out.addPage(p));
        outputs.push({
          name: `split-${i + 1}.pdf`,
          bytes: await out.save(),
        });
        updateProgress("arrange", i + 1, groups.length);
      }
      checkCancelled("arrange");
      if (outputs.length === 1) {
        downloadBlob(new Blob([outputs[0].bytes], { type: "application/pdf" }), outputs[0].name);
      } else {
        const zip = new JSZip();
        outputs.forEach((out) => zip.file(out.name, out.bytes));
        const blob = await zip.generateAsync({ type: "blob" });
        downloadBlob(blob, "split-pdfs.zip");
      }
      updateProgress("arrange", 100, 100);
      endOperation(
        "arrange",
        outputs.length === 1
          ? "완료: 1개 분할 PDF 단건 다운로드"
          : `완료: ${outputs.length}개 파일로 분할 (ZIP 다운로드)`
      );
    } catch (err) {
      handleOperationError("arrange", err);
    }
  });
};

const setupPdfMerge = () => {
  if (!$("runMergePdf")) return;
  $("mergePdfFiles")?.addEventListener("change", async () => {
    toolFileState.mergePdf = [...$("mergePdfFiles").files];
    await renderMergePdfPreview("mergePdfPreview", "mergePdf", "mergePdfFiles");
  });
  $("runMergePdf").addEventListener("click", async () => {
    const files = [...$("mergePdfFiles").files];
    if (!files.length) {
      setStatus("mergePdfStatus", "병합할 PDF를 선택해주세요.");
      return;
    }

    startOperation("mergePdf", "PDF 병합 중...");
    try {
      const merged = await PDFLib.PDFDocument.create();
      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("mergePdf");
        setStatus("mergePdfStatus", `파일 병합 중 (${i + 1}/${files.length})`);
        const doc = await PDFLib.PDFDocument.load(await readAsArrayBuffer(files[i]));
        const pages = await merged.copyPages(doc, [...Array(doc.getPageCount()).keys()]);
        pages.forEach((p) => merged.addPage(p));
        updateProgress("mergePdf", i + 1, files.length);
      }
      checkCancelled("mergePdf");
      const blob = new Blob([await merged.save()], { type: "application/pdf" });
      downloadBlob(blob, "merged.pdf");
      updateProgress("mergePdf", 100, 100);
      endOperation("mergePdf", `완료: ${files.length}개 파일 병합 (${formatBytes(blob.size)})`);
    } catch (err) {
      handleOperationError("mergePdf", err);
    }
  });
};

const renderPdfCompressResult = (rows) => {
  const wrap = $("pdfCompressResult");
  if (!wrap) return;
  wrap.innerHTML = "";
  if (!rows.length) {
    const empty = document.createElement("p");
    empty.className = "status";
    empty.textContent = "아직 실행 결과가 없습니다.";
    wrap.appendChild(empty);
    return;
  }

  const totalBefore = rows.reduce((sum, row) => sum + row.beforeSize, 0);
  const totalAfter = rows.reduce((sum, row) => sum + row.afterSize, 0);
  const totalSaved = totalBefore - totalAfter;
  const totalRate = totalBefore > 0 ? (totalSaved / totalBefore) * 100 : 0;

  const summary = document.createElement("p");
  summary.className = "order-line";
  summary.textContent =
    `총 ${rows.length}개 파일 · 원본 ${formatBytes(totalBefore)} → 압축 ${formatBytes(totalAfter)} ` +
    `(절감 ${formatBytes(Math.max(0, totalSaved))}, ${totalRate.toFixed(1)}%)`;
  wrap.appendChild(summary);

  rows.forEach((row) => {
    const item = document.createElement("div");
    item.className = "compress-result-item";
    const saved = row.beforeSize - row.afterSize;
    const savedRate = row.beforeSize > 0 ? (saved / row.beforeSize) * 100 : 0;
    item.innerHTML = `
      <p class="compress-name">${row.name}</p>
      <p class="compress-meta">페이지: ${row.selectedPages}/${row.totalPages}</p>
      <p class="compress-meta">원본: ${formatBytes(row.beforeSize)} → 압축: ${formatBytes(row.afterSize)}</p>
      <p class="compress-meta compress-emphasis">절감: ${formatBytes(Math.max(0, saved))} (${savedRate.toFixed(1)}%)</p>
    `;
    wrap.appendChild(item);
  });
};

const setupPdfCompress = () => {
  if (!$("runPdfCompress")) return;
  $("pdfCompressFiles")?.addEventListener("change", async () => {
    toolFileState.pdfCompress = [...$("pdfCompressFiles").files];
    await renderMergePdfPreview("pdfCompressPreview", "pdfCompress", "pdfCompressFiles");
  });

  $("runPdfCompress").addEventListener("click", async () => {
    const files = [...$("pdfCompressFiles").files];
    if (!files.length) {
      setStatus("pdfCompressStatus", "압축할 PDF를 선택해주세요.");
      return;
    }

    const preset = $("pdfCompressPreset")?.value || "balanced";
    const baseDpi = Number($("pdfCompressDpi")?.value || 150);
    const baseQuality = Number($("pdfCompressQuality")?.value || 72);
    const pageInput = ($("pdfCompressPages")?.value || "").trim();

    const presetMap = {
      mild: { dpiMul: 1.0, qualityMul: 1.0 },
      balanced: { dpiMul: 0.85, qualityMul: 0.9 },
      strong: { dpiMul: 0.7, qualityMul: 0.78 },
    };
    const presetOpt = presetMap[preset] || presetMap.balanced;
    const effectiveDpi = Math.max(72, Math.round(baseDpi * presetOpt.dpiMul));
    const effectiveQuality = Math.min(95, Math.max(30, Math.round(baseQuality * presetOpt.qualityMul)));

    const meta = [];
    let totalSteps = 0;
    for (let i = 0; i < files.length; i += 1) {
      const cacheKey = `${getFileCacheKey(files[i])}__130`;
      let totalPages = getCachedPdfPageCount(cacheKey);
      if (!totalPages) {
        const { pageCount } = await loadPdfFrontThumb(files[i], 130);
        totalPages = pageCount;
      }
      const selectedPages = parsePageTokens(pageInput, totalPages);
      totalSteps += selectedPages.length;
      meta.push({
        totalPages,
        selectedPages,
      });
    }

    if (!totalSteps) {
      setStatus("pdfCompressStatus", "유효한 페이지가 없습니다. 페이지 범위를 확인해주세요.");
      return;
    }

    startOperation("pdfCompress", "PDF 용량 압축 준비 중...");
    renderPdfCompressResult([]);
    try {
      const rows = [];
      const outputs = [];
      let processedSteps = 0;

      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("pdfCompress");
        const file = files[i];
        const beforeSize = file.size;
        const srcBytes = await readAsArrayBuffer(file);
        const pdf = await pdfjsLib.getDocument({ data: srcBytes }).promise;
        const out = await PDFLib.PDFDocument.create();
        const selectedPages = meta[i].selectedPages;

        for (let p = 0; p < selectedPages.length; p += 1) {
          checkCancelled("pdfCompress");
          const pageNo = selectedPages[p];
          setStatus(
            "pdfCompressStatus",
            `${file.name} 압축 중 (${p + 1}/${selectedPages.length}) · DPI ${effectiveDpi} · Q ${effectiveQuality}`
          );
          const srcPage = await pdf.getPage(pageNo);
          const baseViewport = srcPage.getViewport({ scale: 1 });
          const scale = effectiveDpi / 72;
          const viewport = srcPage.getViewport({ scale });
          const canvas = document.createElement("canvas");
          canvas.width = Math.max(1, Math.floor(viewport.width));
          canvas.height = Math.max(1, Math.floor(viewport.height));
          await srcPage.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;

          const jpgDataUrl = canvas.toDataURL("image/jpeg", effectiveQuality / 100);
          const jpgImage = await out.embedJpg(dataUrlToUint8Array(jpgDataUrl));
          const outPage = out.addPage([baseViewport.width, baseViewport.height]);
          outPage.drawImage(jpgImage, {
            x: 0,
            y: 0,
            width: baseViewport.width,
            height: baseViewport.height,
          });

          processedSteps += 1;
          updateProgress("pdfCompress", processedSteps, totalSteps);
        }

        checkCancelled("pdfCompress");
        const outBytes = await out.save();
        const outName = `${String(i + 1).padStart(2, "0")}_${file.name.replace(/\.pdf$/i, "")}_compressed.pdf`;
        outputs.push({
          name: outName,
          bytes: outBytes,
        });

        rows.push({
          name: `${file.name} → ${outName}`,
          beforeSize,
          afterSize: outBytes.length,
          totalPages: meta[i].totalPages,
          selectedPages: selectedPages.length,
        });
      }

      checkCancelled("pdfCompress");
      if (outputs.length === 1) {
        downloadBlob(new Blob([outputs[0].bytes], { type: "application/pdf" }), outputs[0].name);
      } else {
        const zip = new JSZip();
        outputs.forEach((out) => zip.file(out.name, out.bytes));
        const zipBlob = await zip.generateAsync({ type: "blob" });
        downloadBlob(zipBlob, "compressed-pdfs.zip");
      }
      updateProgress("pdfCompress", totalSteps, totalSteps);
      renderPdfCompressResult(rows);
      endOperation(
        "pdfCompress",
        rows.length === 1
          ? `완료: 1개 PDF 압축 (단건 다운로드)`
          : `완료: ${rows.length}개 압축 (ZIP 다운로드) · DPI ${effectiveDpi}, JPEG 품질 ${effectiveQuality}`
      );
    } catch (err) {
      handleOperationError("pdfCompress", err);
    }
  });
};

const renderGenericFilePreview = (previewId, stateKey, inputId, prefix = "") => {
  const grid = $(previewId);
  if (!grid) return;
  grid.innerHTML = "";
  const files = toolFileState[stateKey];
  if (!files?.length) return;

  for (let i = 0; i < files.length; i += 1) {
    const item = document.createElement("div");
    item.className = "thumb-item";
    item.draggable = true;
    item.dataset.idx = String(i);
    item.innerHTML = `<button class="thumb-delete" type="button" title="파일 제거" aria-label="파일 제거">${ICONS.trash3}</button>`;
    const stub = document.createElement("div");
    stub.className = "file-type-stub";
    stub.textContent = prefix || (files[i].name.split(".").pop() || "FILE").toUpperCase();
    const label = document.createElement("div");
    label.className = "thumb-label";
    label.textContent = `${String(i + 1).padStart(2, "0")}. ${files[i].name}`;
    item.appendChild(stub);
    item.appendChild(label);
    grid.appendChild(item);
  }

  grid.onclick = (e) => {
    const del = e.target.closest(".thumb-delete");
    if (!del) return;
    const cell = e.target.closest(".thumb-item");
    if (!cell) return;
    removeFileAtIndex(stateKey, inputId, Number(cell.dataset.idx));
  };

  let dragIdx = -1;
  const placeholder = document.createElement("div");
  placeholder.className = "drag-placeholder";
  grid.ondragstart = (e) => {
    const cell = e.target.closest(".thumb-item");
    if (!cell) return;
    dragIdx = Number(cell.dataset.idx);
    cell.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  };
  grid.ondragend = () => {
    dragIdx = -1;
    if (placeholder.parentElement) placeholder.parentElement.removeChild(placeholder);
    grid.querySelectorAll(".thumb-item.dragging").forEach((el) => el.classList.remove("dragging"));
  };
  grid.ondragover = (e) => {
    if (dragIdx < 0) return;
    e.preventDefault();
    const intent = getPdfToImageDragAfterElement(grid, e.clientX, e.clientY);
    if (!intent.afterEl) grid.appendChild(placeholder);
    else grid.insertBefore(placeholder, intent.afterEl);
  };
  grid.ondrop = (e) => {
    if (dragIdx < 0) return;
    e.preventDefault();
    const moving = toolFileState[stateKey][dragIdx];
    const filtered = toolFileState[stateKey].filter((_, i) => i !== dragIdx);
    const next = placeholder.nextElementSibling?.closest?.(".thumb-item");
    if (next) {
      const nextIdx = Number(next.dataset.idx);
      filtered.splice(nextIdx, 0, moving);
    } else {
      filtered.push(moving);
    }
    toolFileState[stateKey] = filtered;
    syncFilesToInput(inputId, filtered);
  };
};

const renderDwgResult = (rows) => {
  const wrap = $("dwgToPdfResult");
  if (!wrap) return;
  wrap.innerHTML = "";
  if (!rows.length) {
    const empty = document.createElement("p");
    empty.className = "status";
    empty.textContent = "아직 실행 결과가 없습니다.";
    wrap.appendChild(empty);
    return;
  }

  const totalBefore = rows.reduce((sum, row) => sum + row.beforeSize, 0);
  const totalAfter = rows.reduce((sum, row) => sum + row.afterSize, 0);
  const summary = document.createElement("p");
  summary.className = "order-line";
  summary.textContent = `총 ${rows.length}개 파일 · 원본 ${formatBytes(totalBefore)} → PDF ${formatBytes(totalAfter)}`;
  wrap.appendChild(summary);

  rows.forEach((row) => {
    const item = document.createElement("div");
    item.className = "compress-result-item";
    item.innerHTML = `
      <p class="compress-name">${row.name}</p>
      <p class="compress-meta">원본: ${formatBytes(row.beforeSize)} · PDF: ${formatBytes(row.afterSize)}</p>
    `;
    wrap.appendChild(item);
  });
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const setupDwgToPdf = () => {
  if (!$("runDwgToPdf")) return;

  $("dwgFiles")?.addEventListener("change", () => {
    toolFileState.dwgToPdf = [...$("dwgFiles").files];
    renderGenericFilePreview("dwgToPdfPreview", "dwgToPdf", "dwgFiles", "DWG");
  });

  const cloudConvertRequest = async (baseUrl, path, apiKey, options = {}) => {
    const response = await fetch(`${baseUrl}${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${apiKey}`,
        ...(options.headers || {}),
      },
    });
    let payload = null;
    try {
      payload = await response.json();
    } catch {
      payload = null;
    }
    if (!response.ok) {
      const details = payload?.message || payload?.data?.message || `HTTP ${response.status}`;
      throw new Error(`CloudConvert 오류: ${details}`);
    }
    return payload?.data ?? payload;
  };

  const waitForCloudConvertJob = async (baseUrl, jobId, apiKey) => {
    while (true) {
      checkCancelled("dwgToPdf");
      const job = await cloudConvertRequest(
        baseUrl,
        `/jobs/${jobId}?include=tasks`,
        apiKey,
        { method: "GET" }
      );
      if (job.status === "finished") return job;
      if (job.status === "error") {
        const taskError = (job.tasks || []).find((t) => t.status === "error");
        throw new Error(taskError?.message || "변환 작업이 실패했습니다.");
      }
      await sleep(1400);
    }
  };

  $("runDwgToPdf").addEventListener("click", async () => {
    const files = [...$("dwgFiles").files];
    const apiKey = ($("dwgApiKey")?.value || "").trim();
    const baseUrl = ($("dwgApiEndpoint")?.value || "https://api.cloudconvert.com/v2").trim();
    if (!apiKey) {
      setStatus("dwgToPdfStatus", "CloudConvert API Key를 입력해주세요.");
      return;
    }
    if (!files.length) {
      setStatus("dwgToPdfStatus", "DWG 파일을 선택해주세요.");
      return;
    }

    startOperation("dwgToPdf", "DWG → PDF 변환 준비 중...");
    renderDwgResult([]);
    try {
      const rows = [];
      const outputs = [];
      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("dwgToPdf");
        const file = files[i];
        if (!/\.(dwg|dxf)$/i.test(file.name)) {
          throw new Error(`${file.name}: DWG/DXF 파일만 변환 가능합니다.`);
        }
        const inputFormat = /\.dxf$/i.test(file.name) ? "dxf" : "dwg";

        setStatus("dwgToPdfStatus", `${file.name} 업로드/변환 요청 중... (${i + 1}/${files.length})`);
        const base64 = (await readAsDataURL(file)).split(",")[1];

        const created = await cloudConvertRequest(baseUrl, "/jobs", apiKey, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            tasks: {
              import_file: {
                operation: "import/base64",
                file: base64,
                filename: file.name,
              },
              convert_file: {
                operation: "convert",
                input: "import_file",
                input_format: inputFormat,
                output_format: "pdf",
                filename: `${file.name.replace(/\.[^.]+$/i, "")}.pdf`,
              },
              export_file: {
                operation: "export/url",
                input: "convert_file",
              },
            },
            tag: `kunhwa-dwg-${Date.now()}-${i + 1}`,
          }),
        });

        setStatus("dwgToPdfStatus", `${file.name} 변환 대기 중...`);
        const doneJob = await waitForCloudConvertJob(baseUrl, created.id, apiKey);
        const exportTask = (doneJob.tasks || []).find(
          (task) => task.operation === "export/url" && task.status === "finished"
        );
        const outFile = exportTask?.result?.files?.[0];
        if (!outFile?.url) throw new Error(`${file.name}: 결과 다운로드 URL을 찾지 못했습니다.`);
        const outRes = await fetch(outFile.url);
        if (!outRes.ok) throw new Error(`${file.name}: 결과 파일 다운로드 실패`);
        const outBlob = await outRes.blob();
        const pdfName = `${String(i + 1).padStart(2, "0")}_${file.name.replace(/\.[^.]+$/i, "")}.pdf`;
        outputs.push({
          name: pdfName,
          blob: outBlob,
        });

        rows.push({
          name: `${file.name} → ${pdfName}`,
          beforeSize: file.size,
          afterSize: outBlob.size,
        });
        updateProgress("dwgToPdf", i + 1, files.length);
      }

      checkCancelled("dwgToPdf");
      if (outputs.length === 1) {
        downloadBlob(outputs[0].blob, outputs[0].name);
      } else {
        const zip = new JSZip();
        outputs.forEach((out) => zip.file(out.name, out.blob));
        const zipBlob = await zip.generateAsync({ type: "blob" });
        downloadBlob(zipBlob, "dwg-to-pdf.zip");
      }
      updateProgress("dwgToPdf", files.length, files.length);
      renderDwgResult(rows);
      endOperation(
        "dwgToPdf",
        rows.length === 1
          ? "완료: 1개 DWG/DXF 파일 단건 PDF 다운로드"
          : `완료: ${rows.length}개 파일을 PDF로 변환 (ZIP 다운로드)`
      );
    } catch (err) {
      handleOperationError("dwgToPdf", err);
    }
  });
};

const setupImageResize = () => {
  if (!$("runResize")) return;
  $("resizeFiles")?.addEventListener("change", async () => {
    toolFileState.resize = [...$("resizeFiles").files];
    await renderImageThumbPreview("resizePreview", "resize", "resizeFiles");
  });
  $("runResize").addEventListener("click", async () => {
    const files = [...$("resizeFiles").files];
    const width = Number($("resizeWidth").value);
    const height = Number($("resizeHeight").value);
    const fmt = $("resizeFormat")?.value || "webp";
    const quality = Math.min(1, Math.max(0.01, Number($("resizeQuality")?.value || 82) / 100));
    if (!files.length || (!width && !height)) {
      setStatus("resizeStatus", "이미지 파일과 너비/높이 중 하나 이상을 입력해주세요.");
      return;
    }

    startOperation("resize", "이미지 리사이즈 중...");
    try {
      const outputs = [];
      const canvasToBlobAsync = (canvas, mime, q) =>
        new Promise((resolve, reject) => {
          canvas.toBlob(
            (blob) => {
              if (!blob) {
                reject(new Error(`브라우저가 ${mime} 인코딩을 지원하지 않습니다.`));
                return;
              }
              resolve(blob);
            },
            mime,
            q
          );
        });
      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("resize");
        setStatus("resizeStatus", `리사이즈 처리 중 (${i + 1}/${files.length})`);
        const { img } = await loadImageFromFile(files[i]);
        const ratio = img.width / img.height;
        const targetW = width || Math.round(height * ratio);
        const targetH = height || Math.round(width / ratio);
        const canvas = document.createElement("canvas");
        canvas.width = targetW;
        canvas.height = targetH;
        canvas.getContext("2d").drawImage(img, 0, 0, targetW, targetH);
        const mime = fmt === "png" ? "image/png" : fmt === "jpeg" ? "image/jpeg" : "image/webp";
        const ext = fmt === "jpeg" ? "jpg" : fmt;
        const outBlob = await canvasToBlobAsync(canvas, mime, quality);
        outputs.push({
          name: `${files[i].name.replace(/\.[^.]+$/, "")}_${targetW}x${targetH}.${ext}`,
          blob: outBlob,
        });
        updateProgress("resize", i + 1, files.length);
      }
      checkCancelled("resize");
      if (outputs.length === 1) {
        downloadBlob(outputs[0].blob, outputs[0].name);
      } else {
        const zip = new JSZip();
        outputs.forEach((out) => zip.file(out.name, out.blob));
        const blob = await zip.generateAsync({ type: "blob" });
        downloadBlob(blob, "resized-images.zip");
      }
      updateProgress("resize", 100, 100);
      endOperation(
        "resize",
        outputs.length === 1
          ? "완료: 1개 파일 단건 다운로드"
          : `완료: ${outputs.length}개 변환 (ZIP 다운로드)`
      );
    } catch (err) {
      handleOperationError("resize", err);
    }
  });
};

const setupImageFormat = () => {
  if (!$("runFormatConvert")) return;
  const updateHeicDirectNotice = (files) => {
    const notice = $("heicDirectNotice");
    if (!notice) return;
    const hasHeic = files.some((file) => isHeicLikeFile(file));
    notice.classList.toggle("visible", hasHeic);
    notice.setAttribute("aria-hidden", hasHeic ? "false" : "true");
  };

  $("formatFiles")?.addEventListener("change", async () => {
    toolFileState.format = [...$("formatFiles").files];
    await renderImageThumbPreview("formatPreview", "format", "formatFiles");
    updateHeicDirectNotice(toolFileState.format);
  });

  const isCanvasMimeSupported = (mime) => {
    try {
      const c = document.createElement("canvas");
      c.width = 2;
      c.height = 2;
      const url = c.toDataURL(mime, 0.9);
      return url.startsWith(`data:${mime}`);
    } catch {
      return false;
    }
  };

  const canvasToBlobAsync = (canvas, mime, quality) =>
    new Promise((resolve, reject) => {
      canvas.toBlob(
        (blob) => {
          if (!blob) {
            reject(new Error(`브라우저가 ${mime} 인코딩을 지원하지 않습니다.`));
            return;
          }
          resolve(blob);
        },
        mime,
        quality
      );
    });

  const renderGifBlobFromCanvas = (canvas) =>
    new Promise((resolve, reject) => {
      if (typeof GIF !== "function") {
        reject(new Error("GIF 변환 라이브러리(gif.js)를 불러오지 못했습니다."));
        return;
      }
      const gif = new GIF({
        workers: 1,
        quality: 10,
        width: canvas.width,
        height: canvas.height,
        workerScript: "https://cdnjs.cloudflare.com/ajax/libs/gif.js/0.2.0/gif.worker.js",
      });
      gif.on("finished", (blob) => resolve(blob));
      gif.on("abort", () => reject(new Error("GIF 렌더링이 중단되었습니다.")));
      gif.addFrame(canvas, { copy: true, delay: 0 });
      gif.render();
    });

  const outputExt = (fmt) => (fmt === "jpeg" ? "jpg" : fmt);
  const mimeByFormat = {
    png: "image/png",
    jpeg: "image/jpeg",
    webp: "image/webp",
    avif: "image/avif",
  };

  $("runFormatConvert").addEventListener("click", async () => {
    const files = [...$("formatFiles").files];
    const fmt = $("targetFormat").value;
    const quality = Number($("targetQuality").value);
    const hasHeicInput = files.some((file) => isHeicLikeFile(file));
    if (!files.length) {
      setStatus("formatStatus", "이미지를 선택해주세요.");
      return;
    }

    startOperation("format", "이미지 포맷 변환 중...");
    try {
      const outputs = [];
      const isGifTarget = fmt === "gif";
      if (fmt === "avif" && !isCanvasMimeSupported("image/avif")) {
        throw new Error("현재 브라우저는 AVIF 내보내기를 지원하지 않습니다.");
      }
      if (!isGifTarget && !mimeByFormat[fmt]) {
        throw new Error(`지원하지 않는 대상 포맷입니다: ${fmt}`);
      }

      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("format");
        setStatus("formatStatus", `포맷 변환 중 (${i + 1}/${files.length})`);
        const { img } = await loadImageFromAnyFile(files[i]);
        const canvas = document.createElement("canvas");
        canvas.width = img.width;
        canvas.height = img.height;
        canvas.getContext("2d").drawImage(img, 0, 0);

        if (isGifTarget) {
          const gifBlob = await renderGifBlobFromCanvas(canvas);
          outputs.push({
            name: `${files[i].name.replace(/\.[^.]+$/, "")}.gif`,
            blob: gifBlob,
          });
        } else {
          const mime = mimeByFormat[fmt];
          const outBlob = await canvasToBlobAsync(canvas, mime, quality);
          outputs.push({
            name: `${files[i].name.replace(/\.[^.]+$/, "")}.${outputExt(fmt)}`,
            blob: outBlob,
          });
        }
        updateProgress("format", i + 1, files.length);
      }

      checkCancelled("format");
      if (hasHeicInput) {
        outputs.forEach((out) => downloadBlob(out.blob, out.name));
      } else if (outputs.length === 1) {
        downloadBlob(outputs[0].blob, outputs[0].name);
      } else {
        const zip = new JSZip();
        outputs.forEach((out) => zip.file(out.name, out.blob));
        const blob = await zip.generateAsync({ type: "blob" });
        downloadBlob(blob, "converted-images.zip");
      }
      updateProgress("format", 100, 100);
      endOperation(
        "format",
        hasHeicInput
          ? `완료: HEIC/HEIF 포함 감지 · ${outputs.length}개 파일 개별 직접 다운로드`
          : outputs.length === 1
          ? `완료: 1개 파일 단건 다운로드`
          : `완료: ${outputs.length}개 변환 (ZIP 다운로드)`
      );
    } catch (err) {
      handleOperationError("format", err);
    }
  });
};

const setupBatchRename = () => {
  if (!$("runBatchRename")) return;
  setIconButton("runBatchRenameFromPreview", "download");

  const getOptions = () => {
    const startRaw = Number($("renameStart")?.value || 1);
    const stepRaw = Number($("renameStep")?.value || 1);
    const padRaw = Number($("renamePad")?.value || 0);
    return {
      findText: $("renameFind")?.value || "",
      replaceText: $("renameReplace")?.value || "",
      prefix: $("renamePrefix")?.value || "",
      suffix: $("renameSuffix")?.value || "",
      addSeq: Boolean($("renameAddSeq")?.checked),
      useRegex: Boolean($("renameUseRegex")?.checked),
      caseSensitive: Boolean($("renameCaseSensitive")?.checked),
      start: Number.isFinite(startRaw) ? startRaw : 1,
      step: Number.isFinite(stepRaw) ? stepRaw : 1,
      pad: Number.isFinite(padRaw) ? Math.max(0, Math.trunc(padRaw)) : 0,
      seqPos: ($("renameSeqPos")?.value || "prefix") === "suffix" ? "suffix" : "prefix",
      sep: $("renameSep")?.value || "",
    };
  };

  const buildPlan = (files, options) => {
    if (!files.length) {
      return { rows: [], changedCount: 0, collisions: 0, error: "" };
    }

    let replaceRegex = null;
    if (options.findText) {
      if (options.useRegex) {
        try {
          replaceRegex = new RegExp(options.findText, options.caseSensitive ? "g" : "gi");
        } catch {
          return {
            rows: [],
            changedCount: 0,
            collisions: 0,
            error: "정규식 문법 오류입니다. 패턴을 확인해주세요.",
          };
        }
      } else {
        replaceRegex = new RegExp(escapeRegExp(options.findText), options.caseSensitive ? "g" : "gi");
      }
    }

    const rows = [];
    const used = new Map();
    let collisions = 0;
    let changedCount = 0;

    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      const { stem, ext } = splitFileName(file.name);
      let nextStem = stem;

      if (replaceRegex) {
        nextStem = nextStem.replace(replaceRegex, options.replaceText);
      }

      nextStem = `${options.prefix}${nextStem}${options.suffix}`;

      if (options.addSeq) {
        const raw = Math.trunc(options.start + options.step * i);
        const abs = String(Math.abs(raw));
        const num =
          options.pad > 0
            ? `${raw < 0 ? "-" : ""}${abs.padStart(Math.max(1, Math.trunc(options.pad)), "0")}`
            : String(raw);
        const token = options.sep ? `${num}${options.sep}` : num;
        nextStem = options.seqPos === "prefix" ? `${token}${nextStem}` : `${nextStem}${options.sep}${num}`;
      }

      nextStem = sanitizeFileStem(nextStem);
      let nextName = `${nextStem}${ext}`;
      let duplicateIndex = 1;
      while (used.has(nextName.toLowerCase())) {
        duplicateIndex += 1;
        collisions += 1;
        nextName = `${sanitizeFileStem(`${nextStem} (${duplicateIndex})`)}${ext}`;
      }
      used.set(nextName.toLowerCase(), true);

      const changed = nextName !== file.name;
      if (changed) changedCount += 1;
      rows.push({
        originalName: file.name,
        nextName,
        changed,
        duplicateIndex,
      });
    }

    return { rows, changedCount, collisions, error: "" };
  };

  const renderPlan = (plan) => {
    const list = $("renamePlanList");
    const summary = $("renamePlanSummary");
    if (!list || !summary) return;
    list.innerHTML = "";

    if (!toolFileState.rename.length) {
      summary.textContent = "규칙을 입력하면 미리보기가 표시됩니다.";
      return;
    }
    if (plan.error) {
      summary.textContent = plan.error;
      return;
    }

    summary.textContent = `총 ${plan.rows.length}개 · 변경 ${plan.changedCount}개 · 중복 자동 보정 ${plan.collisions}건`;
    plan.rows.forEach((row, idx) => {
      const item = document.createElement("div");
      item.className = `rename-plan-item${row.changed ? " changed" : ""}`;
      item.innerHTML = `
        <p class="rename-plan-index">${String(idx + 1).padStart(2, "0")}</p>
        <p class="rename-plan-original" title="${row.originalName}">${row.originalName}</p>
        <p class="rename-plan-arrow">→</p>
        <p class="rename-plan-next" title="${row.nextName}">${row.nextName}</p>
      `;
      list.appendChild(item);
    });
  };

  let latestPlan = { rows: [], changedCount: 0, collisions: 0, error: "" };
  const refreshPlan = () => {
    const files = [...($("renameFiles")?.files || [])];
    toolFileState.rename = files;
    latestPlan = buildPlan(files, getOptions());
    renderPlan(latestPlan);
    if (latestPlan.error) setStatus("batchRenameStatus", latestPlan.error);
    else setStatus("batchRenameStatus", "");
  };

  $("renameFiles")?.addEventListener("change", refreshPlan);
  [
    "renameFind",
    "renameReplace",
    "renamePrefix",
    "renameSuffix",
    "renameStart",
    "renameStep",
    "renamePad",
    "renameSeqPos",
    "renameSep",
    "renameAddSeq",
    "renameUseRegex",
    "renameCaseSensitive",
  ].forEach((id) => {
    const el = $(id);
    if (!el) return;
    el.addEventListener("input", refreshPlan);
    el.addEventListener("change", refreshPlan);
  });

  $("runBatchRenameFromPreview")?.addEventListener("click", () => {
    $("runBatchRename").click();
  });

  $("runBatchRename").addEventListener("click", async () => {
    const files = [...($("renameFiles")?.files || [])];
    if (!files.length) {
      setStatus("batchRenameStatus", "파일을 선택해주세요.");
      return;
    }

    latestPlan = buildPlan(files, getOptions());
    renderPlan(latestPlan);
    if (latestPlan.error) {
      setStatus("batchRenameStatus", latestPlan.error);
      return;
    }

    startOperation("batchRename", "파일 이름 변경 다운로드 준비 중...");
    try {
      const outputs = [];
      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("batchRename");
        setStatus("batchRenameStatus", `파일 처리 중 (${i + 1}/${files.length})`);
        const bytes = await readAsArrayBuffer(files[i]);
        outputs.push({
          name: latestPlan.rows[i].nextName,
          bytes,
        });
        updateProgress("batchRename", i + 1, files.length);
      }

      checkCancelled("batchRename");
      if (outputs.length === 1) {
        downloadBlob(new Blob([outputs[0].bytes]), outputs[0].name);
      } else {
        const zip = new JSZip();
        outputs.forEach((out) => zip.file(out.name, out.bytes));
        const blob = await zip.generateAsync({ type: "blob" });
        downloadBlob(blob, "renamed-files.zip");
      }
      updateProgress("batchRename", files.length, files.length);
      endOperation(
        "batchRename",
        outputs.length === 1
          ? "완료: 1개 파일 단건 다운로드"
          : `완료: ${outputs.length}개 파일 이름 변경 (ZIP 다운로드)`
      );
    } catch (err) {
      handleOperationError("batchRename", err);
    }
  });

  refreshPlan();
};

const processRecords = [];

const escapeCsv = (value) => {
  const str = String(value ?? "");
  if (str.includes(",") || str.includes('"') || str.includes("\n")) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
};

const renderProcessMediaPreview = () => {
  const box = $("capturedMediaPreview");
  const input = $("procMediaFiles");
  if (!box || !input) return;
  box.innerHTML = "";
  const files = [...input.files];
  files.forEach((file, idx) => {
    const card = document.createElement("div");
    card.className = "media-item";
    const img = document.createElement("img");
    img.src = URL.createObjectURL(file);
    img.alt = file.name;
    const btn = document.createElement("button");
    btn.className = "thumb-delete";
    btn.type = "button";
    btn.title = "사진 제거";
    btn.setAttribute("aria-label", "사진 제거");
    btn.innerHTML = ICONS.trash3;
    btn.addEventListener("click", () => {
      const remain = files.filter((_, i) => i !== idx);
      syncFilesToInput("procMediaFiles", remain);
    });
    const name = document.createElement("div");
    name.className = "media-name";
    name.textContent = file.name;
    card.appendChild(btn);
    card.appendChild(img);
    card.appendChild(name);
    box.appendChild(card);
  });
};

const setupProcessTimer = () => {
  if (
    !$("timerDisplay") ||
    !$("timerStart") ||
    !$("timerPause") ||
    !$("timerResume") ||
    !$("timerStop")
  ) {
    return;
  }
  $("procMediaFiles")?.addEventListener("change", renderProcessMediaPreview);

  let running = false;
  let startTime = 0;
  let elapsed = 0;
  let interval = null;

  const renderTime = () => {
    const total = running ? Date.now() - startTime + elapsed : elapsed;
    const sec = Math.floor(total / 1000);
    const h = String(Math.floor(sec / 3600)).padStart(2, "0");
    const m = String(Math.floor((sec % 3600) / 60)).padStart(2, "0");
    const s = String(sec % 60).padStart(2, "0");
    $("timerDisplay").textContent = `${h}:${m}:${s}`;
  };

  const startTick = () => {
    if (interval) clearInterval(interval);
    interval = setInterval(renderTime, 300);
  };

  const addSessionCard = ({ durationText, name, customer, memo, mediaItems, timestamp }) => {
    const wrap = document.createElement("div");
    wrap.className = "session-item";
    const now = new Date(timestamp).toLocaleString("ko-KR");
    wrap.innerHTML = `
      <p class="session-title">${name || "공정명 미입력"} (${durationText})</p>
      <p class="session-meta">고객사: ${customer || "-"}</p>
      <p class="session-meta">기록시간: ${now}</p>
      <p class="session-meta">메모: ${memo || "-"}</p>
      <p class="session-media-count">첨부 미디어: ${mediaItems.length}개</p>
    `;

    if (mediaItems.length) {
      const grid = document.createElement("div");
      grid.className = "media-grid";
      mediaItems.forEach((m) => {
        const card = document.createElement("div");
        card.className = "media-item";
        card.innerHTML = `<img src="${m.url}" alt="${m.name}" /><div class="media-name">${m.name}</div>`;
        grid.appendChild(card);
      });
      wrap.appendChild(grid);
    }

    $("timerLog").prepend(wrap);
  };

  const gatherAllMedia = () =>
    [...$("procMediaFiles").files].map((f) => ({
      name: f.name,
      type: f.type || "application/octet-stream",
      url: URL.createObjectURL(f),
    }));

  $("timerStart").addEventListener("click", () => {
    if (running) return;
    running = true;
    elapsed = 0;
    startTime = Date.now();
    startTick();
    renderTime();
    setStatus("timerStatus", "작업 시간 측정을 시작했습니다.");
  });

  $("timerPause").addEventListener("click", () => {
    if (!running) return;
    elapsed += Date.now() - startTime;
    running = false;
    clearInterval(interval);
    renderTime();
    setStatus("timerStatus", "일시정지되었습니다.");
  });

  $("timerResume").addEventListener("click", () => {
    if (running) return;
    running = true;
    startTime = Date.now();
    startTick();
    setStatus("timerStatus", "재개되었습니다.");
  });

  $("timerStop").addEventListener("click", () => {
    if (running) {
      elapsed += Date.now() - startTime;
      running = false;
    }
    clearInterval(interval);
    renderTime();
    const durationText = $("timerDisplay").textContent;
    const durationSec =
      Number(durationText.slice(0, 2)) * 3600 +
      Number(durationText.slice(3, 5)) * 60 +
      Number(durationText.slice(6, 8));
    const timestamp = Date.now();
    const name = $("procName").value.trim();
    const customer = $("procCustomer").value.trim();
    const memo = $("procMemo").value.trim();
    const mediaItems = gatherAllMedia();

    processRecords.push({
      timestamp,
      datetime_local: new Date(timestamp).toLocaleString("ko-KR"),
      duration_text: durationText,
      duration_seconds: durationSec,
      process_name: name || "",
      customer: customer || "",
      memo: memo || "",
      media_count: mediaItems.length,
      media_names: mediaItems.map((m) => m.name).join(" | "),
      media_types: mediaItems.map((m) => m.type).join(" | "),
    });

    addSessionCard({
      durationText,
      name,
      customer,
      memo,
      mediaItems,
      timestamp,
    });

    elapsed = 0;
    renderTime();
    $("procMediaFiles").value = "";
    renderProcessMediaPreview();
    setStatus("timerStatus", "기록이 저장되었습니다.");
  });

  $("exportProcessJson").addEventListener("click", () => {
    if (!processRecords.length) {
      setStatus("timerStatus", "내보낼 기록이 없습니다.");
      return;
    }
    const blob = new Blob([JSON.stringify(processRecords, null, 2)], {
      type: "application/json;charset=utf-8",
    });
    downloadBlob(blob, `process-records-${Date.now()}.json`);
    setStatus("timerStatus", "JSON 내보내기를 완료했습니다.");
  });

  $("exportProcessCsv").addEventListener("click", () => {
    if (!processRecords.length) {
      setStatus("timerStatus", "내보낼 기록이 없습니다.");
      return;
    }
    const headers = [
      "timestamp",
      "datetime_local",
      "duration_text",
      "duration_seconds",
      "process_name",
      "customer",
      "memo",
      "media_count",
      "media_names",
      "media_types",
    ];
    const rows = processRecords.map((r) =>
      headers.map((h) => escapeCsv(r[h])).join(",")
    );
    const csv = [headers.join(","), ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    downloadBlob(blob, `process-records-${Date.now()}.csv`);
    setStatus("timerStatus", "CSV 내보내기를 완료했습니다.");
  });
};

const setupQr = () => {
  if (
    !$("runQr") ||
    !$("saveQr") ||
    !$("qrPreview") ||
    !$("runQrBulk") ||
    !$("downloadQrBulkTemplate")
  ) {
    return;
  }

  const state = {
    lastBlob: null,
    lastExt: "png",
    renderUrl: null,
  };

  const syncQrColorUi = () => {
    const fg = ($("qrFg")?.value || "#000000").toUpperCase();
    const bg = ($("qrBg")?.value || "#FFFFFF").toUpperCase();
    const transparent = !!$("qrTransparentBg")?.checked;
    const fgHex = $("qrFgHex");
    const bgHex = $("qrBgHex");
    const fgSwatch = $("qrFgSwatch");
    const bgSwatch = $("qrBgSwatch");
    const bgLabel = $("qrBg")?.closest(".qr-color-label");
    if (fgHex) fgHex.textContent = fg;
    if (bgHex) bgHex.textContent = transparent ? `${bg} (투명 처리됨)` : bg;
    if (fgSwatch) fgSwatch.style.background = fg;
    if (bgSwatch) {
      bgSwatch.style.background = transparent
        ? "repeating-conic-gradient(#dde6f2 0% 25%, #ffffff 0% 50%) 50% / 10px 10px"
        : bg;
    }
    if (bgLabel) bgLabel.classList.toggle("qr-bg-disabled", transparent);
  };

  const getQrOptions = () => {
    const fmt = document.querySelector("input[name='qrFormat']:checked")?.value || "png";
    return {
      text: $("qrInput").value.trim(),
      size: Math.max(120, Math.min(2000, Number($("qrSize").value || 1000))),
      margin: Math.max(0, Math.min(200, Number($("qrMargin").value || 40))),
      fg: $("qrFg").value || "#000000",
      bg: $("qrBg").value || "#ffffff",
      transparent: !!$("qrTransparentBg").checked,
      format: fmt,
    };
  };

  const makeQrCanvas = (options) =>
    new Promise((resolve, reject) => {
      try {
        const tmp = document.createElement("div");
        tmp.style.position = "fixed";
        tmp.style.left = "-9999px";
        document.body.appendChild(tmp);
        // qrcodejs renders immediately to canvas/img.
        new QRCode(tmp, {
          text: options.text,
          width: options.size,
          height: options.size,
          colorDark: options.fg,
          colorLight: options.transparent ? "rgba(0,0,0,0)" : options.bg,
          correctLevel: QRCode.CorrectLevel.M,
        });
        setTimeout(() => {
          const srcCanvas = tmp.querySelector("canvas");
          const srcImg = tmp.querySelector("img");
          const out = document.createElement("canvas");
          const outSize = options.size + options.margin * 2;
          out.width = outSize;
          out.height = outSize;
          const ctx = out.getContext("2d");
          if (!options.transparent) {
            ctx.fillStyle = options.bg;
            ctx.fillRect(0, 0, out.width, out.height);
          } else {
            ctx.clearRect(0, 0, out.width, out.height);
          }
          if (srcCanvas) {
            ctx.drawImage(srcCanvas, options.margin, options.margin, options.size, options.size);
          } else if (srcImg) {
            ctx.drawImage(srcImg, options.margin, options.margin, options.size, options.size);
          }
          tmp.remove();
          resolve(out);
        }, 0);
      } catch (err) {
        reject(err);
      }
    });

  const canvasToBlob = (canvas, mime, quality) =>
    new Promise((resolve) => {
      canvas.toBlob((blob) => resolve(blob), mime, quality);
    });

  const buildQrAsset = async (options) => {
    const canvas = await makeQrCanvas(options);
    if (options.format === "svg") {
      const pngData = canvas.toDataURL("image/png");
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${canvas.width}" height="${canvas.height}" viewBox="0 0 ${canvas.width} ${canvas.height}"><image href="${pngData}" width="${canvas.width}" height="${canvas.height}" /></svg>`;
      return {
        blob: new Blob([svg], { type: "image/svg+xml;charset=utf-8" }),
        previewUrl: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`,
        ext: "svg",
      };
    }
    const mime =
      options.format === "jpeg"
        ? "image/jpeg"
        : options.format === "webp"
          ? "image/webp"
          : "image/png";
    const blob = await canvasToBlob(canvas, mime, 0.92);
    const previewUrl = URL.createObjectURL(blob);
    return {
      blob,
      previewUrl,
      ext: options.format === "jpeg" ? "jpg" : options.format,
    };
  };

  const renderQrPreview = (url) => {
    const box = $("qrPreview");
    box.innerHTML = "";
    const img = document.createElement("img");
    img.src = url;
    img.alt = "QR preview";
    img.style.width = "100%";
    img.style.maxWidth = "360px";
    img.style.height = "auto";
    box.appendChild(img);
  };

  const refreshQrPreviewLive = async () => {
    const options = getQrOptions();
    if (!options.text) return;
    try {
      const asset = await buildQrAsset(options);
      state.lastBlob = asset.blob;
      state.lastExt = asset.ext;
      state.renderUrl = asset.previewUrl;
      renderQrPreview(asset.previewUrl);
    } catch {
      // Ignore live-preview errors; explicit "QR 생성" will surface details.
    }
  };

  ["qrFg", "qrBg", "qrTransparentBg"].forEach((id) => {
    $(id)?.addEventListener("input", () => {
      syncQrColorUi();
      refreshQrPreviewLive();
    });
    $(id)?.addEventListener("change", () => {
      syncQrColorUi();
      refreshQrPreviewLive();
    });
  });
  syncQrColorUi();

  $("runQr").addEventListener("click", async () => {
    const options = getQrOptions();
    if (!options.text) {
      setStatus("qrStatus", "QR 내용을 입력해주세요.");
      return;
    }
    beginGlobalBusy("QR 이미지를 생성 중입니다...");
    try {
      setGlobalBusyMessage("QR 코드를 렌더링 중입니다...");
      if (state.renderUrl?.startsWith("blob:")) URL.revokeObjectURL(state.renderUrl);
      const asset = await buildQrAsset(options);
      state.lastBlob = asset.blob;
      state.lastExt = asset.ext;
      state.renderUrl = asset.previewUrl;
      renderQrPreview(asset.previewUrl);
      setStatus("qrStatus", "QR 코드 생성 완료");
    } catch (err) {
      setStatus("qrStatus", `QR 생성 오류: ${err.message}`);
    } finally {
      endGlobalBusy();
    }
  });

  $("saveQr").addEventListener("click", () => {
    if (!state.lastBlob) {
      setStatus("qrStatus", "먼저 QR 코드를 생성해주세요.");
      return;
    }
    downloadBlob(state.lastBlob, `qrcode.${state.lastExt}`);
  });

  $("downloadQrBulkTemplate").addEventListener("click", () => {
    const csv = [
      "text,filename,size,margin,fg,bg,transparent,format",
      "https://tools.mytory.net,example-1,800,40,#000000,#ffffff,false,png",
      "HELLO QR,example-2,600,20,#1f4f8f,#ffffff,false,webp",
    ].join("\n");
    downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), "qr-bulk-template.csv");
    setStatus("qrStatus", "CSV 양식을 다운로드했습니다.");
  });

  $("runQrBulk").addEventListener("click", async () => {
    const file = $("qrBulkFile").files[0];
    if (!file) {
      setStatus("qrStatus", "먼저 CSV 양식 파일을 첨부해주세요.");
      return;
    }
    beginGlobalBusy("벌크 QR를 준비 중입니다...");
    try {
      const raw = await readAsText(file);
      const lines = raw.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
      if (lines.length <= 1) throw new Error("데이터 행이 없습니다.");
      const zip = new JSZip();
      const optionsBase = getQrOptions();
      const rows = lines.slice(1);
      for (let i = 0; i < rows.length; i += 1) {
        setGlobalBusyMessage(`벌크 QR 생성 중 (${i + 1}/${rows.length})`);
        const parts = rows[i].split(",").map((v) => v.trim());
        const text = parts[0] || "";
        const filename = (parts[1] || `qr-${i + 1}`).replace(/[\\/:*?"<>|]/g, "_");
        const rowSize = Number(parts[2] || optionsBase.size);
        const rowMargin = Number(parts[3] || optionsBase.margin);
        const rowFg = /^#[0-9a-fA-F]{6}$/.test(parts[4] || "") ? parts[4] : optionsBase.fg;
        const rowBg = /^#[0-9a-fA-F]{6}$/.test(parts[5] || "") ? parts[5] : optionsBase.bg;
        const rowTransparent = (parts[6] || "").toLowerCase() === "true" || (parts[6] || "") === "1";
        const rowFormatRaw = (parts[7] || optionsBase.format).toLowerCase();
        const rowFormat = ["png", "jpeg", "jpg", "webp", "svg"].includes(rowFormatRaw)
          ? rowFormatRaw === "jpg"
            ? "jpeg"
            : rowFormatRaw
          : optionsBase.format;
        if (!text) continue;
        const asset = await buildQrAsset({
          ...optionsBase,
          text,
          size: Number.isFinite(rowSize) ? Math.max(120, Math.min(2000, rowSize)) : optionsBase.size,
          margin: Number.isFinite(rowMargin) ? Math.max(0, Math.min(200, rowMargin)) : optionsBase.margin,
          fg: rowFg,
          bg: rowBg,
          transparent: rowTransparent,
          format: rowFormat,
        });
        zip.file(`${filename}.${asset.ext}`, asset.blob);
      }
      const blob = await zip.generateAsync({ type: "blob" });
      downloadBlob(blob, "qr-bulk.zip");
      setStatus("qrStatus", "벌크 QR ZIP 생성 완료");
    } catch (err) {
      setStatus("qrStatus", `벌크 처리 오류: ${err.message}`);
    } finally {
      endGlobalBusy();
    }
  });
};

const init = () => {
  document.body.classList.add("home-mode");
  initOperations();
  setupThemeToggle();
  setupLoginModal();
  setIconButton("backToHub", "house");
  setupNavActive();
  setupDropZones();
  setupHashStageRouter();
  setupPdfToImage();
  setupImageToPdf();
  setupPdfArrange();
  setupPdfMerge();
  setupPdfCompress();
  setupImageResize();
  setupImageFormat();
  setupBatchRename();
  setupProcessTimer();
  setupQr();
};

init();

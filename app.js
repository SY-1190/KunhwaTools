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

const getErrorMessage = (err, fallback = "알 수 없는 오류가 발생했습니다.") => {
  if (!err) return fallback;
  if (typeof err === "string") return err;
  if (typeof err.message === "string" && err.message.trim()) return err.message;
  try {
    return JSON.stringify(err);
  } catch {
    return fallback;
  }
};

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
  arrowCounterclockwise: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path fill-rule="evenodd" d="M8 3a5 5 0 1 1-4.546 2.914.5.5 0 0 0-.908-.417A6 6 0 1 0 8 2z"/><path d="M8 4.466V.534a.25.25 0 0 0-.41-.192L5.23 2.308a.25.25 0 0 0 0 .384l2.36 1.966A.25.25 0 0 0 8 4.466"/></svg>`,
  arrowClockwise: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path fill-rule="evenodd" d="M8 3a5 5 0 1 0 4.546 2.914.5.5 0 0 1 .908-.417A6 6 0 1 1 8 2z"/><path d="M8 4.466V.534a.25.25 0 0 1 .41-.192l2.36 1.966a.25.25 0 0 1 0 .384L8.41 4.658A.25.25 0 0 1 8 4.466"/></svg>`,
  check2: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0"/></svg>`,
  download: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5"/><path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708z"/></svg>`,
  house: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="m8.354 1.146 6.5 6.5A.5.5 0 0 1 14.5 8.5H13v5a1 1 0 0 1-1 1h-2.5a.5.5 0 0 1-.5-.5V11a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v3a.5.5 0 0 1-.5.5H4a1 1 0 0 1-1-1v-5H1.5a.5.5 0 0 1-.354-.854l6.5-6.5a.5.5 0 0 1 .708 0"/></svg>`,
  person: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M8 8a3 3 0 1 0 0-6 3 3 0 0 0 0 6"/><path fill-rule="evenodd" d="M8 9a6 6 0 0 0-5.468 3.516A.75.75 0 0 0 3.205 14h9.59a.75.75 0 0 0 .673-1.484A6 6 0 0 0 8 9"/></svg>`,
  journal: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M3 1.5A1.5 1.5 0 0 1 4.5 0h7A1.5 1.5 0 0 1 13 1.5v13A1.5 1.5 0 0 1 11.5 16h-7A1.5 1.5 0 0 1 3 14.5zM4.5 1a.5.5 0 0 0-.5.5V2h8v-.5a.5.5 0 0 0-.5-.5zm7.5 2h-8v11.5a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 .5-.5z"/><path d="M5 5h6v1H5zm0 2h6v1H5zm0 2h4v1H5z"/></svg>`,
  link45: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M4.715 6.542 3.343 7.914a3 3 0 1 0 4.243 4.243l1.828-1.829A3 3 0 0 0 8.586 5.5L8 6.086a1 1 0 0 0-.154.199 2 2 0 0 1 .861 3.337L6.88 11.45a2 2 0 1 1-2.83-2.83l.793-.792a4 4 0 0 1-.128-1.287z"/><path d="M6.586 4.672A3 3 0 0 0 7.414 9.5l.775-.776a2 2 0 0 1-.896-3.346L9.12 3.55a2 2 0 1 1 2.83 2.83l-.793.792c.112.42.155.855.128 1.287l1.372-1.372a3 3 0 1 0-4.243-4.243z"/></svg>`,
  wifi: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M15.384 6.115a.485.485 0 0 0-.047-.736A12.44 12.44 0 0 0 8 3C5.259 3 2.723 3.882.663 5.379a.485.485 0 0 0-.048.736.52.52 0 0 0 .668.05A11.45 11.45 0 0 1 8 4c2.507 0 4.827.802 6.716 2.164.205.148.49.13.668-.049"/><path d="M13.229 8.271a.482.482 0 0 0-.063-.745A9.46 9.46 0 0 0 8 6c-1.905 0-3.68.56-5.166 1.526a.48.48 0 0 0-.063.745.525.525 0 0 0 .652.065A8.46 8.46 0 0 1 8 7a8.46 8.46 0 0 1 4.576 1.336c.206.132.48.108.653-.065m-2.183 2.183c.226-.226.185-.605-.1-.75A6.5 6.5 0 0 0 8 9c-1.06 0-2.062.254-2.946.704-.285.145-.326.524-.1.75l.015.015c.16.16.407.19.611.09A5.5 5.5 0 0 1 8 10c.868 0 1.69.201 2.42.56.203.1.45.07.61-.091zM9.06 12.44c.196-.196.198-.52-.04-.66A2 2 0 0 0 8 11.5a2 2 0 0 0-1.02.28c-.238.14-.236.464-.04.66l.706.706a.5.5 0 0 0 .707 0l.707-.707z"/></svg>`,
  calendarEvent: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M11 6.5a.5.5 0 0 1 .5-.5h1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-1a.5.5 0 0 1-.5-.5z"/><path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5M1 4v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V4z"/></svg>`,
  ticket: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M4 4.85v.9h1v-.9zm7 0v.9h1v-.9zm-7 1.8v.9h1v-.9zm7 0v.9h1v-.9zm-7 1.8v.9h1v-.9zm7 0v.9h1v-.9zm-7 1.8v.9h1v-.9zm7 0v.9h1v-.9z"/><path d="M1.5 3A1.5 1.5 0 0 0 0 4.5V6a.5.5 0 0 0 .5.5 1.5 1.5 0 1 1 0 3 .5.5 0 0 0-.5.5v1.5A1.5 1.5 0 0 0 1.5 13h13a1.5 1.5 0 0 0 1.5-1.5V10a.5.5 0 0 0-.5-.5 1.5 1.5 0 0 1 0-3A.5.5 0 0 0 16 6V4.5A1.5 1.5 0 0 0 14.5 3zM1 4.5a.5.5 0 0 1 .5-.5h13a.5.5 0 0 1 .5.5v1.05a2.5 2.5 0 0 0 0 4.9v1.05a.5.5 0 0 1-.5.5h-13a.5.5 0 0 1-.5-.5v-1.05a2.5 2.5 0 0 0 0-4.9z"/></svg>`,
  pencil: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M15.502 1.94a.5.5 0 0 1 0 .706l-1 1a.5.5 0 0 1-.707 0l-1.439-1.439a.5.5 0 0 1 0-.707l1-1a.5.5 0 0 1 .707 0z"/><path d="M13.5 3.207 5 11.707V13h1.293l8.5-8.5z"/><path fill-rule="evenodd" d="M1 13.5A1.5 1.5 0 0 0 2.5 15h11a1.5 1.5 0 0 0 1.5-1.5v-6a.5.5 0 0 0-1 0v6a.5.5 0 0 1-.5.5h-11a.5.5 0 0 1-.5-.5v-11a.5.5 0 0 1 .5-.5H9a.5.5 0 0 0 0-1H2.5A1.5 1.5 0 0 0 1 2.5z"/></svg>`,
  arrowLeft: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path fill-rule="evenodd" d="M15 8a.5.5 0 0 0-.5-.5H2.707l3.147-3.146a.5.5 0 1 0-.708-.708l-4 4a.5.5 0 0 0 0 .708l4 4a.5.5 0 0 0 .708-.708L2.707 8.5H14.5A.5.5 0 0 0 15 8"/></svg>`,
  arrowRight: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path fill-rule="evenodd" d="M1 8a.5.5 0 0 1 .5-.5h11.793l-3.147-3.146a.5.5 0 0 1 .708-.708l4 4a.5.5 0 0 1 0 .708l-4 4a.5.5 0 0 1-.708-.708L13.293 8.5H1.5A.5.5 0 0 1 1 8"/></svg>`,
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

const setupLogoGameEntry = () => {
  const mascot = $("brandMascot");
  if (!mascot) return;
  const goGameHub = () => {
    window.location.hash = "#gameHub";
  };
  mascot.addEventListener("dblclick", goGameHub);
  mascot.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" && e.key !== " ") return;
    e.preventDefault();
    goGameHub();
  });
};

const setupYachtGame = () => {
  if (!$("gameYacht")) return;
  const diceFaces = ["", "⚀", "⚁", "⚂", "⚃", "⚄", "⚅"];
  const categories = [
    { key: "ones", label: "에이스(1)" },
    { key: "twos", label: "듀스(2)" },
    { key: "threes", label: "트리(3)" },
    { key: "fours", label: "포(4)" },
    { key: "fives", label: "파이브(5)" },
    { key: "sixes", label: "식스(6)" },
    { key: "threeKind", label: "3 of a kind" },
    { key: "fourKind", label: "4 of a kind" },
    { key: "fullHouse", label: "풀하우스" },
    { key: "smallStraight", label: "스몰 스트레이트" },
    { key: "largeStraight", label: "라지 스트레이트" },
    { key: "yacht", label: "야추" },
    { key: "chance", label: "찬스" },
  ];
  const upperKeys = ["ones", "twos", "threes", "fours", "fives", "sixes"];
  const lowerKeys = ["threeKind", "fourKind", "fullHouse", "smallStraight", "largeStraight", "yacht", "chance"];
  const modeEl = $("yachtMode");

  const state = {
    mode: "solo",
    dice: [1, 1, 1, 1, 1],
    held: [false, false, false, false, false],
    rollsLeft: 3,
    scoresPlayer: {},
    scoresCpu: {},
    finished: false,
  };
  categories.forEach((c) => {
    state.scoresPlayer[c.key] = null;
    state.scoresCpu[c.key] = null;
  });

  const getCounts = (dice) => {
    const counts = [0, 0, 0, 0, 0, 0, 0];
    dice.forEach((d) => {
      counts[d] += 1;
    });
    return counts;
  };

  const calcScore = (key, dice) => {
    const sum = dice.reduce((a, b) => a + b, 0);
    const counts = getCounts(dice);
    const uniq = [...new Set(dice)].sort((a, b) => a - b);
    const hasSeq = (arr) => arr.every((n) => uniq.includes(n));

    switch (key) {
      case "ones":
        return counts[1] * 1;
      case "twos":
        return counts[2] * 2;
      case "threes":
        return counts[3] * 3;
      case "fours":
        return counts[4] * 4;
      case "fives":
        return counts[5] * 5;
      case "sixes":
        return counts[6] * 6;
      case "threeKind":
        return counts.some((v) => v >= 3) ? sum : 0;
      case "fourKind":
        return counts.some((v) => v >= 4) ? sum : 0;
      case "fullHouse":
        return counts.includes(3) && counts.includes(2) ? 25 : 0;
      case "smallStraight":
        return hasSeq([1, 2, 3, 4]) || hasSeq([2, 3, 4, 5]) || hasSeq([3, 4, 5, 6]) ? 30 : 0;
      case "largeStraight":
        return hasSeq([1, 2, 3, 4, 5]) || hasSeq([2, 3, 4, 5, 6]) ? 40 : 0;
      case "yacht":
        return counts.some((v) => v === 5) ? 50 : 0;
      case "chance":
        return sum;
      default:
        return 0;
    }
  };

  const totalsOf = (scoreMap) => {
    const upper = upperKeys.reduce((sum, key) => sum + (scoreMap[key] ?? 0), 0);
    const bonus = upper >= 63 ? 35 : 0;
    const lower = lowerKeys.reduce((sum, key) => sum + (scoreMap[key] ?? 0), 0);
    return { upper, bonus, lower, grand: upper + bonus + lower };
  };
  const isAllScored = (scoreMap) => categories.every((c) => scoreMap[c.key] !== null);

  const runCpuTurn = () => {
    const remaining = categories.filter((c) => state.scoresCpu[c.key] === null);
    if (!remaining.length) return null;
    let best = null;
    for (let r = 0; r < 3; r += 1) {
      const dice = Array.from({ length: 5 }, () => Math.floor(Math.random() * 6) + 1);
      remaining.forEach((cat) => {
        const score = calcScore(cat.key, dice);
        if (!best || score > best.score) best = { key: cat.key, label: cat.label, score };
      });
    }
    if (!best) return null;
    state.scoresCpu[best.key] = best.score;
    return best;
  };

  const setYachtStatus = (text) => {
    setStatus("yachtStatus", text);
  };

  const render = () => {
    const diceRow = $("yachtDiceRow");
    const table = $("yachtScoreTable");
    const rollsLeft = $("yachtRollsLeft");
    const cpuMeta = $("yachtCpuMeta");
    const turnMeta = $("yachtTurnMeta");
    if (!diceRow || !table || !rollsLeft || !cpuMeta || !turnMeta) return;

    diceRow.innerHTML = "";
    state.dice.forEach((value, idx) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `yacht-die${state.held[idx] ? " held" : ""}`;
      btn.textContent = diceFaces[value];
      btn.title = state.held[idx] ? "홀드 해제" : "홀드";
      btn.disabled = state.finished || state.rollsLeft === 3;
      btn.addEventListener("click", () => {
        if (state.finished || state.rollsLeft === 3) return;
        state.held[idx] = !state.held[idx];
        render();
      });
      diceRow.appendChild(btn);
    });

    rollsLeft.textContent = `남은 굴림: ${state.rollsLeft}`;
    const pTotals = totalsOf(state.scoresPlayer);
    const cTotals = totalsOf(state.scoresCpu);
    const filledTurns = categories.filter((c) => state.scoresPlayer[c.key] !== null).length;
    turnMeta.textContent = `턴 ${Math.min(filledTurns + 1, 13)}/13`;
    cpuMeta.textContent =
      state.mode === "cpu"
        ? `CPU 총점: ${cTotals.grand}`
        : `솔로 모드 총점: ${pTotals.grand}`;

    table.innerHTML = "";
    const headerRow = document.createElement("div");
    headerRow.className = "yacht-row header";
    headerRow.innerHTML = `
      <span class="label">카테고리</span>
      <span class="preview">예상</span>
      <span class="score">나</span>
      <span class="score cpu">${state.mode === "cpu" ? "CPU" : "-"}</span>
    `;
    table.appendChild(headerRow);

    const appendSection = (label) => {
      const row = document.createElement("div");
      row.className = "yacht-row section";
      row.innerHTML = `
        <span class="label">${label}</span>
        <span class="preview"></span>
        <span class="score"></span>
        <span class="score cpu"></span>
      `;
      table.appendChild(row);
    };

    appendSection("상단 섹션");
    categories.forEach((cat) => {
      if (cat.key === "threeKind") appendSection("하단 섹션");
      const lockedScore = state.scoresPlayer[cat.key];
      const cpuScore = state.scoresCpu[cat.key];
      const preview = state.rollsLeft === 3 ? "-" : calcScore(cat.key, state.dice);
      const row = document.createElement("button");
      row.type = "button";
      row.className = `yacht-row selectable${lockedScore !== null ? " locked" : ""}`;
      row.innerHTML = `
        <span class="label">${cat.label}</span>
        <span class="preview">${lockedScore === null ? `예상 ${preview}` : "기록됨"}</span>
        <span class="score">${lockedScore ?? "-"}</span>
        <span class="score cpu">${state.mode === "cpu" ? cpuScore ?? "-" : "-"}</span>
      `;
      row.disabled = state.finished || lockedScore !== null;
      row.addEventListener("click", () => {
        if (state.finished || state.scoresPlayer[cat.key] !== null) return;
        if (state.rollsLeft === 3) {
          setYachtStatus("먼저 주사위를 굴려주세요.");
          return;
        }
        const bestAvailable = categories
          .filter((c) => state.scoresPlayer[c.key] === null)
          .reduce((best, c) => Math.max(best, calcScore(c.key, state.dice)), 0);
        const score = calcScore(cat.key, state.dice);
        if (score === 0 && bestAvailable > 0) {
          const ok = window.confirm(
            `이 선택은 0점입니다. 현재 주사위로 최대 ${bestAvailable}점 선택이 가능합니다. 그래도 진행할까요?`
          );
          if (!ok) return;
        }
        state.scoresPlayer[cat.key] = score;

        let cpuNote = "";
        if (state.mode === "cpu") {
          const cpuPlay = runCpuTurn();
          if (cpuPlay) cpuNote = ` · CPU ${cpuPlay.label} ${cpuPlay.score}점`;
        }

        if (isAllScored(state.scoresPlayer)) {
          state.finished = true;
          const playerGrand = totalsOf(state.scoresPlayer).grand;
          if (state.mode === "cpu") {
            const cpuGrand = totalsOf(state.scoresCpu).grand;
            const result =
              playerGrand > cpuGrand ? "승리" : playerGrand < cpuGrand ? "패배" : "무승부";
            setYachtStatus(`게임 종료! 나 ${playerGrand} : CPU ${cpuGrand} (${result})`);
          } else {
            setYachtStatus(`게임 종료! 총점 ${playerGrand}점`);
          }
        } else {
          state.rollsLeft = 3;
          state.held = [false, false, false, false, false];
          setYachtStatus(`${cat.label} ${score}점 기록${cpuNote}. 다음 턴 시작`);
        }
        render();
      });
      table.appendChild(row);
    });

    const addMetaRow = (label, pValue, cValue = "-") => {
      const row = document.createElement("div");
      row.className = "yacht-row meta";
      row.innerHTML = `
        <span class="label">${label}</span>
        <span class="preview"></span>
        <span class="score">${pValue}</span>
        <span class="score cpu">${state.mode === "cpu" ? cValue : "-"}</span>
      `;
      table.appendChild(row);
    };

    addMetaRow("상단 합계", pTotals.upper, cTotals.upper);
    addMetaRow("보너스(63+)", pTotals.bonus, cTotals.bonus);
    addMetaRow("하단 합계", pTotals.lower, cTotals.lower);
    addMetaRow("총점", pTotals.grand, cTotals.grand);
  };

  $("yachtRollBtn")?.addEventListener("click", () => {
    if (state.finished) {
      setYachtStatus("게임이 종료되었습니다. 새 게임을 눌러 다시 시작하세요.");
      return;
    }
    if (state.rollsLeft <= 0) {
      setYachtStatus("남은 굴림이 없습니다. 점수판에서 카테고리를 선택하세요.");
      return;
    }

    for (let i = 0; i < state.dice.length; i += 1) {
      if (!state.held[i]) {
        state.dice[i] = Math.floor(Math.random() * 6) + 1;
      }
    }
    state.rollsLeft -= 1;
    if (state.rollsLeft === 0) setYachtStatus("굴림 종료. 점수판에서 카테고리를 선택하세요.");
    else setYachtStatus(`굴림 완료. 남은 굴림 ${state.rollsLeft}회`);
    render();
  });

  const resetGame = () => {
    state.mode = modeEl?.value === "cpu" ? "cpu" : "solo";
    state.dice = [1, 1, 1, 1, 1];
    state.held = [false, false, false, false, false];
    state.rollsLeft = 3;
    state.finished = false;
    categories.forEach((c) => {
      state.scoresPlayer[c.key] = null;
      state.scoresCpu[c.key] = null;
    });
    setYachtStatus(
      state.mode === "cpu"
        ? "CPU 대결 모드 시작! 굴리기를 눌러주세요."
        : "솔로 모드 시작! 굴리기를 눌러주세요."
    );
    render();
  };

  $("yachtResetBtn")?.addEventListener("click", resetGame);
  modeEl?.addEventListener("change", resetGame);

  resetGame();
};

const setupRpsGame = () => {
  if (!$("gameRps")) return;
  const scoreText = $("rpsScoreText");
  const statusEl = $("rpsStatus");
  const choiceButtons = [...document.querySelectorAll("button[data-rps]")];
  const state = {
    win: 0,
    lose: 0,
    draw: 0,
    round: 0,
    maxRound: 5,
    done: false,
  };
  const labelMap = {
    rock: "바위",
    scissors: "가위",
    paper: "보",
  };
  const winMap = {
    rock: "scissors",
    scissors: "paper",
    paper: "rock",
  };
  const choices = ["rock", "scissors", "paper"];

  const render = () => {
    if (scoreText) scoreText.textContent = `내 점수 ${state.win} : ${state.lose} 컴퓨터 · ${state.round}/${state.maxRound}판`;
    choiceButtons.forEach((btn) => {
      btn.disabled = state.done;
    });
  };

  const setRpsStatus = (text) => {
    if (statusEl) statusEl.textContent = text;
  };

  const finishIfNeeded = () => {
    if (state.round < state.maxRound) return false;
    state.done = true;
    if (state.win > state.lose) setRpsStatus(`게임 종료! 승리 (${state.win}:${state.lose})`);
    else if (state.win < state.lose) setRpsStatus(`게임 종료! 패배 (${state.win}:${state.lose})`);
    else setRpsStatus(`게임 종료! 무승부 (${state.win}:${state.lose})`);
    render();
    return true;
  };

  choiceButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (state.done) return;
      const mine = btn.dataset.rps;
      const cpu = choices[Math.floor(Math.random() * choices.length)];
      state.round += 1;
      if (mine === cpu) {
        state.draw += 1;
        setRpsStatus(`라운드 ${state.round}: 비김 · 나(${labelMap[mine]}) vs 컴퓨터(${labelMap[cpu]})`);
      } else if (winMap[mine] === cpu) {
        state.win += 1;
        setRpsStatus(`라운드 ${state.round}: 승리 · 나(${labelMap[mine]}) vs 컴퓨터(${labelMap[cpu]})`);
      } else {
        state.lose += 1;
        setRpsStatus(`라운드 ${state.round}: 패배 · 나(${labelMap[mine]}) vs 컴퓨터(${labelMap[cpu]})`);
      }
      render();
      finishIfNeeded();
    });
  });

  $("rpsResetBtn")?.addEventListener("click", () => {
    state.win = 0;
    state.lose = 0;
    state.draw = 0;
    state.round = 0;
    state.done = false;
    setRpsStatus("선택해서 게임을 시작하세요.");
    render();
  });

  render();
};

const setupTttGame = () => {
  if (!$("gameTtt")) return;
  const boardEl = $("tttBoard");
  const statusEl = $("tttStatus");
  const modeEl = $("tttMode");
  if (!boardEl || !statusEl || !modeEl) return;

  const wins = [
    [0, 1, 2],
    [3, 4, 5],
    [6, 7, 8],
    [0, 3, 6],
    [1, 4, 7],
    [2, 5, 8],
    [0, 4, 8],
    [2, 4, 6],
  ];
  const state = {
    board: Array(9).fill(""),
    turn: "X",
    mode: "local",
    done: false,
  };

  const winnerOf = () => {
    for (let i = 0; i < wins.length; i += 1) {
      const [a, b, c] = wins[i];
      if (state.board[a] && state.board[a] === state.board[b] && state.board[a] === state.board[c]) return state.board[a];
    }
    if (state.board.every(Boolean)) return "draw";
    return null;
  };

  const evaluateMove = (board, mark) => {
    for (let i = 0; i < wins.length; i += 1) {
      const [a, b, c] = wins[i];
      const line = [board[a], board[b], board[c]];
      const markCount = line.filter((v) => v === mark).length;
      const emptyIdx = [a, b, c].find((idx) => !board[idx]);
      if (markCount === 2 && Number.isInteger(emptyIdx)) return emptyIdx;
    }
    return -1;
  };

  const pickCpuMove = () => {
    const board = state.board;
    let idx = evaluateMove(board, "O");
    if (idx >= 0) return idx;
    idx = evaluateMove(board, "X");
    if (idx >= 0) return idx;
    const center = 4;
    if (!board[center]) return center;
    const corners = [0, 2, 6, 8].filter((i) => !board[i]);
    if (corners.length) return corners[Math.floor(Math.random() * corners.length)];
    const empties = board.map((v, i) => (!v ? i : -1)).filter((i) => i >= 0);
    if (!empties.length) return -1;
    return empties[Math.floor(Math.random() * empties.length)];
  };

  const applyResultOrNext = () => {
    const winner = winnerOf();
    if (winner === "draw") {
      state.done = true;
      statusEl.textContent = "무승부입니다.";
      return;
    }
    if (winner) {
      state.done = true;
      if (state.mode === "cpu" && winner === "O") statusEl.textContent = "CPU 승리!";
      else if (state.mode === "cpu" && winner === "X") statusEl.textContent = "플레이어 승리!";
      else statusEl.textContent = `플레이어 ${winner} 승리!`;
      return;
    }
    state.turn = state.turn === "X" ? "O" : "X";
    if (state.mode === "cpu" && state.turn === "O") statusEl.textContent = "CPU 생각 중...";
    else statusEl.textContent = `플레이어 ${state.turn} 차례`;
  };

  const render = () => {
    boardEl.innerHTML = "";
    for (let i = 0; i < 9; i += 1) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "ttt-cell";
      btn.textContent = state.board[i];
      btn.disabled = state.done || Boolean(state.board[i]);
      btn.addEventListener("click", () => {
        if (state.done || state.board[i]) return;
        if (state.mode === "cpu" && state.turn !== "X") return;
        state.board[i] = state.turn;
        applyResultOrNext();
        render();
        if (!state.done && state.mode === "cpu" && state.turn === "O") {
          setTimeout(() => {
            const cpuIdx = pickCpuMove();
            if (cpuIdx < 0 || state.done || state.board[cpuIdx]) return;
            state.board[cpuIdx] = "O";
            applyResultOrNext();
            render();
          }, 230);
        }
      });
      boardEl.appendChild(btn);
    }
  };

  const resetTtt = () => {
    state.mode = modeEl.value === "cpu" ? "cpu" : "local";
    state.board = Array(9).fill("");
    state.turn = "X";
    state.done = false;
    statusEl.textContent = state.mode === "cpu" ? "플레이어(X) 차례" : "플레이어 X 차례";
    render();
  };

  $("tttResetBtn")?.addEventListener("click", resetTtt);
  modeEl.addEventListener("change", resetTtt);

  resetTtt();
};

const setupUpdownGame = () => {
  if (!$("gameUpdown")) return;
  const input = $("updownGuessInput");
  const submitBtn = $("updownSubmitBtn");
  const resetBtn = $("updownResetBtn");
  const meta = $("updownMeta");
  const statusEl = $("updownStatus");
  const historyEl = $("updownHistory");
  if (!input || !submitBtn || !resetBtn || !meta || !statusEl || !historyEl) return;

  const state = {
    target: Math.floor(Math.random() * 100) + 1,
    left: 10,
    done: false,
    tries: [],
  };

  const render = () => {
    meta.textContent = `남은 시도: ${state.left}`;
    historyEl.innerHTML = "";
    state.tries
      .slice()
      .reverse()
      .forEach((item) => {
        const p = document.createElement("p");
        p.textContent = item;
        historyEl.appendChild(p);
      });
    submitBtn.disabled = state.done;
    input.disabled = state.done;
  };

  const reset = () => {
    state.target = Math.floor(Math.random() * 100) + 1;
    state.left = 10;
    state.done = false;
    state.tries = [];
    input.value = "";
    statusEl.textContent = "숫자를 입력하고 확인을 누르세요.";
    render();
  };

  submitBtn.addEventListener("click", () => {
    if (state.done) return;
    const guess = Number(input.value);
    if (!Number.isInteger(guess) || guess < 1 || guess > 100) {
      statusEl.textContent = "1~100 사이 정수를 입력해주세요.";
      return;
    }
    state.left -= 1;
    if (guess === state.target) {
      state.done = true;
      state.tries.push(`${guess} 정답`);
      statusEl.textContent = `정답! ${10 - state.left}번 만에 맞췄습니다.`;
      render();
      return;
    }
    const hint = guess < state.target ? "UP" : "DOWN";
    state.tries.push(`${guess} → ${hint}`);
    if (state.left <= 0) {
      state.done = true;
      statusEl.textContent = `실패! 정답은 ${state.target}였습니다.`;
    } else {
      statusEl.textContent = `${hint}! 다시 시도하세요.`;
    }
    render();
  });

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") submitBtn.click();
  });

  resetBtn.addEventListener("click", reset);
  reset();
};

const isTypingTarget = (el) => ["INPUT", "TEXTAREA", "SELECT"].includes(el?.tagName || "");
const activeHashIs = (id) => window.location.hash === `#${id}`;

const setup2048Game = () => {
  if (!$("game2048")) return;
  const boardEl = $("g2048Board");
  const scoreEl = $("g2048Score");
  const statusEl = $("g2048Status");
  const size = 4;
  const state = { board: [], score: 0, over: false };

  const emptyCells = () => {
    const cells = [];
    state.board.forEach((row, y) => row.forEach((v, x) => { if (!v) cells.push([x, y]); }));
    return cells;
  };
  const addTile = () => {
    const cells = emptyCells();
    if (!cells.length) return;
    const [x, y] = cells[Math.floor(Math.random() * cells.length)];
    state.board[y][x] = Math.random() < 0.9 ? 2 : 4;
  };
  const compress = (line) => {
    const nums = line.filter(Boolean);
    const out = [];
    for (let i = 0; i < nums.length; i += 1) {
      if (nums[i] === nums[i + 1]) {
        out.push(nums[i] * 2);
        state.score += nums[i] * 2;
        i += 1;
      } else {
        out.push(nums[i]);
      }
    }
    while (out.length < size) out.push(0);
    return out;
  };
  const canMove = () => {
    if (emptyCells().length) return true;
    for (let y = 0; y < size; y += 1) {
      for (let x = 0; x < size; x += 1) {
        const v = state.board[y][x];
        if (state.board[y]?.[x + 1] === v || state.board[y + 1]?.[x] === v) return true;
      }
    }
    return false;
  };
  const render = () => {
    boardEl.innerHTML = "";
    state.board.flat().forEach((v) => {
      const tile = document.createElement("div");
      tile.className = `g2048-tile${v ? ` v${v}` : ""}`;
      tile.textContent = v || "";
      boardEl.appendChild(tile);
    });
    scoreEl.textContent = `점수 ${state.score}`;
    if (!canMove()) {
      state.over = true;
      statusEl.textContent = "게임 종료. 새 게임을 눌러주세요.";
    }
  };
  const reset = () => {
    state.board = Array.from({ length: size }, () => Array(size).fill(0));
    state.score = 0;
    state.over = false;
    addTile();
    addTile();
    statusEl.textContent = "타일을 합쳐 2048을 만들어보세요.";
    render();
  };
  const move = (dir) => {
    if (state.over) return;
    const before = JSON.stringify(state.board);
    if (dir === "left" || dir === "right") {
      state.board = state.board.map((row) => {
        const line = dir === "left" ? row : [...row].reverse();
        const next = compress(line);
        return dir === "left" ? next : next.reverse();
      });
    } else {
      for (let x = 0; x < size; x += 1) {
        const col = state.board.map((row) => row[x]);
        const line = dir === "up" ? col : col.reverse();
        const next = compress(line);
        const finalCol = dir === "up" ? next : next.reverse();
        finalCol.forEach((v, y) => { state.board[y][x] = v; });
      }
    }
    if (before !== JSON.stringify(state.board)) addTile();
    if (state.board.flat().includes(2048)) statusEl.textContent = "2048 달성!";
    render();
  };
  document.addEventListener("keydown", (e) => {
    if (!activeHashIs("game2048") || isTypingTarget(e.target)) return;
    const map = { ArrowLeft: "left", a: "left", A: "left", ArrowRight: "right", d: "right", D: "right", ArrowUp: "up", w: "up", W: "up", ArrowDown: "down", s: "down", S: "down" };
    if (!map[e.key]) return;
    e.preventDefault();
    move(map[e.key]);
  });
  $("g2048ResetBtn")?.addEventListener("click", reset);
  reset();
};

const setupSnakeGame = () => {
  if (!$("gameSnake")) return;
  const canvas = $("snakeCanvas");
  const ctx = canvas.getContext("2d");
  const scoreEl = $("snakeScore");
  const statusEl = $("snakeStatus");
  const cell = 16;
  const cols = canvas.width / cell;
  const rows = canvas.height / cell;
  const state = { snake: [], food: { x: 0, y: 0 }, dir: { x: 1, y: 0 }, next: { x: 1, y: 0 }, score: 0, timer: null, running: false };
  const placeFood = () => {
    do {
      state.food = { x: Math.floor(Math.random() * cols), y: Math.floor(Math.random() * rows) };
    } while (state.snake.some((p) => p.x === state.food.x && p.y === state.food.y));
  };
  const draw = () => {
    ctx.fillStyle = "#101722";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#78d08a";
    state.snake.forEach((p) => ctx.fillRect(p.x * cell + 1, p.y * cell + 1, cell - 2, cell - 2));
    ctx.fillStyle = "#f05f6b";
    ctx.fillRect(state.food.x * cell + 2, state.food.y * cell + 2, cell - 4, cell - 4);
    scoreEl.textContent = `점수 ${state.score}`;
  };
  const reset = () => {
    clearInterval(state.timer);
    state.snake = [{ x: 8, y: 10 }, { x: 7, y: 10 }, { x: 6, y: 10 }];
    state.dir = { x: 1, y: 0 };
    state.next = { x: 1, y: 0 };
    state.score = 0;
    state.running = false;
    placeFood();
    statusEl.textContent = "시작을 눌러주세요.";
    draw();
  };
  const step = () => {
    state.dir = state.next;
    const head = { x: state.snake[0].x + state.dir.x, y: state.snake[0].y + state.dir.y };
    if (head.x < 0 || head.y < 0 || head.x >= cols || head.y >= rows || state.snake.some((p) => p.x === head.x && p.y === head.y)) {
      clearInterval(state.timer);
      state.running = false;
      statusEl.textContent = "게임 종료.";
      return;
    }
    state.snake.unshift(head);
    if (head.x === state.food.x && head.y === state.food.y) {
      state.score += 10;
      placeFood();
    } else {
      state.snake.pop();
    }
    draw();
  };
  const start = () => {
    if (state.running) return;
    state.running = true;
    statusEl.textContent = "진행 중";
    state.timer = setInterval(step, 115);
  };
  document.addEventListener("keydown", (e) => {
    if (!activeHashIs("gameSnake") || isTypingTarget(e.target)) return;
    const dirs = { ArrowLeft: { x: -1, y: 0 }, a: { x: -1, y: 0 }, A: { x: -1, y: 0 }, ArrowRight: { x: 1, y: 0 }, d: { x: 1, y: 0 }, D: { x: 1, y: 0 }, ArrowUp: { x: 0, y: -1 }, w: { x: 0, y: -1 }, W: { x: 0, y: -1 }, ArrowDown: { x: 0, y: 1 }, s: { x: 0, y: 1 }, S: { x: 0, y: 1 } };
    const next = dirs[e.key];
    if (!next) return;
    e.preventDefault();
    if (next.x + state.dir.x !== 0 || next.y + state.dir.y !== 0) state.next = next;
  });
  $("snakeStartBtn")?.addEventListener("click", start);
  $("snakeResetBtn")?.addEventListener("click", reset);
  reset();
};

const setupTetrisGame = () => {
  if (!$("gameTetris")) return;
  const canvas = $("tetrisCanvas");
  const ctx = canvas.getContext("2d");
  const scoreEl = $("tetrisScore");
  const statusEl = $("tetrisStatus");
  const cell = 24;
  const w = 10;
  const h = 20;
  const shapes = [
    [[1, 1, 1, 1]],
    [[1, 1], [1, 1]],
    [[0, 1, 0], [1, 1, 1]],
    [[1, 0, 0], [1, 1, 1]],
    [[0, 0, 1], [1, 1, 1]],
    [[1, 1, 0], [0, 1, 1]],
    [[0, 1, 1], [1, 1, 0]],
  ];
  const colors = ["#63b3ff", "#f1c84b", "#b886f5", "#f08d5f", "#5fa8f0", "#67c587", "#ef6f87"];
  const state = { board: [], piece: null, score: 0, lines: 0, timer: null, running: false };
  const rotate = (m) => m[0].map((_, i) => m.map((row) => row[i]).reverse());
  const spawn = () => {
    const idx = Math.floor(Math.random() * shapes.length);
    state.piece = { shape: shapes[idx], color: colors[idx], x: 3, y: 0 };
    if (collides(state.piece, 0, 0)) gameOver();
  };
  const collides = (p, dx, dy, shape = p.shape) => shape.some((row, y) => row.some((v, x) => v && (p.x + x + dx < 0 || p.x + x + dx >= w || p.y + y + dy >= h || state.board[p.y + y + dy]?.[p.x + x + dx])));
  const merge = () => {
    state.piece.shape.forEach((row, y) => row.forEach((v, x) => {
      if (v && state.board[state.piece.y + y]) state.board[state.piece.y + y][state.piece.x + x] = state.piece.color;
    }));
    let cleared = 0;
    state.board = state.board.filter((row) => {
      if (row.every(Boolean)) { cleared += 1; return false; }
      return true;
    });
    while (state.board.length < h) state.board.unshift(Array(w).fill(""));
    if (cleared) {
      state.lines += cleared;
      state.score += [0, 100, 300, 500, 800][cleared] || cleared * 200;
    }
    spawn();
  };
  const draw = () => {
    ctx.fillStyle = "#101722";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    const drawCell = (x, y, color) => {
      ctx.fillStyle = color;
      ctx.fillRect(x * cell + 1, y * cell + 1, cell - 2, cell - 2);
    };
    state.board.forEach((row, y) => row.forEach((v, x) => { if (v) drawCell(x, y, v); }));
    if (state.piece) state.piece.shape.forEach((row, y) => row.forEach((v, x) => { if (v) drawCell(state.piece.x + x, state.piece.y + y, state.piece.color); }));
    scoreEl.textContent = `점수 ${state.score} · 줄 ${state.lines}`;
  };
  const step = () => {
    if (!state.piece) return;
    if (!collides(state.piece, 0, 1)) state.piece.y += 1;
    else merge();
    draw();
  };
  const reset = () => {
    clearInterval(state.timer);
    state.board = Array.from({ length: h }, () => Array(w).fill(""));
    state.score = 0;
    state.lines = 0;
    state.running = false;
    spawn();
    statusEl.textContent = "시작을 눌러주세요.";
    draw();
  };
  const start = () => {
    if (state.running) return;
    state.running = true;
    statusEl.textContent = "진행 중";
    state.timer = setInterval(step, 520);
  };
  const gameOver = () => {
    clearInterval(state.timer);
    state.running = false;
    statusEl.textContent = "게임 종료.";
  };
  document.addEventListener("keydown", (e) => {
    if (!activeHashIs("gameTetris") || isTypingTarget(e.target) || !state.piece) return;
    const key = e.key;
    if (key === "ArrowLeft" || key === "a" || key === "A") { e.preventDefault(); if (!collides(state.piece, -1, 0)) state.piece.x -= 1; }
    if (key === "ArrowRight" || key === "d" || key === "D") { e.preventDefault(); if (!collides(state.piece, 1, 0)) state.piece.x += 1; }
    if (key === "ArrowDown" || key === "s" || key === "S") { e.preventDefault(); step(); }
    if (key === "ArrowUp" || key === "w" || key === "W") {
      e.preventDefault();
      const rotated = rotate(state.piece.shape);
      if (!collides(state.piece, 0, 0, rotated)) state.piece.shape = rotated;
    }
    draw();
  });
  $("tetrisStartBtn")?.addEventListener("click", start);
  $("tetrisResetBtn")?.addEventListener("click", reset);
  reset();
};

const setupPacmanGame = () => {
  if (!$("gamePacman")) return;
  const canvas = $("pacmanCanvas");
  const ctx = canvas.getContext("2d");
  const scoreEl = $("pacmanScore");
  const statusEl = $("pacmanStatus");
  const cell = 20;
  const mapText = [
    "################",
    "#..............#",
    "#.####.##.####.#",
    "#..............#",
    "#.##.######.##.#",
    "#..............#",
    "####.##..##.####",
    "#..............#",
    "#.####.##.####.#",
    "#..............#",
    "################",
  ];
  const state = { map: [], p: { x: 1, y: 1 }, g: { x: 14, y: 9 }, dir: { x: 1, y: 0 }, next: { x: 1, y: 0 }, score: 0, timer: null, running: false };
  const reset = () => {
    clearInterval(state.timer);
    state.map = mapText.map((r) => r.split(""));
    state.p = { x: 1, y: 1 };
    state.g = { x: 14, y: 9 };
    state.dir = { x: 1, y: 0 };
    state.next = { x: 1, y: 0 };
    state.score = 0;
    state.running = false;
    statusEl.textContent = "시작을 눌러주세요.";
    draw();
  };
  const wall = (x, y) => state.map[y]?.[x] === "#";
  const draw = () => {
    ctx.fillStyle = "#101722";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    state.map.forEach((row, y) => row.forEach((v, x) => {
      if (v === "#") { ctx.fillStyle = "#315b9b"; ctx.fillRect(x * cell, y * cell, cell, cell); }
      if (v === ".") { ctx.fillStyle = "#f4d35e"; ctx.beginPath(); ctx.arc(x * cell + 10, y * cell + 10, 2.5, 0, Math.PI * 2); ctx.fill(); }
    }));
    ctx.fillStyle = "#ffd447";
    ctx.beginPath();
    ctx.arc(state.p.x * cell + 10, state.p.y * cell + 10, 8, 0.2 * Math.PI, 1.8 * Math.PI);
    ctx.lineTo(state.p.x * cell + 10, state.p.y * cell + 10);
    ctx.fill();
    ctx.fillStyle = "#f05f6b";
    ctx.beginPath();
    ctx.arc(state.g.x * cell + 10, state.g.y * cell + 10, 8, 0, Math.PI * 2);
    ctx.fill();
    scoreEl.textContent = `점수 ${state.score}`;
  };
  const moveGhost = () => {
    const dirs = [{ x: 1, y: 0 }, { x: -1, y: 0 }, { x: 0, y: 1 }, { x: 0, y: -1 }]
      .filter((d) => !wall(state.g.x + d.x, state.g.y + d.y))
      .sort((a, b) => Math.abs(state.p.x - (state.g.x + a.x)) + Math.abs(state.p.y - (state.g.y + a.y)) - (Math.abs(state.p.x - (state.g.x + b.x)) + Math.abs(state.p.y - (state.g.y + b.y))));
    const d = Math.random() < 0.75 ? dirs[0] : dirs[Math.floor(Math.random() * dirs.length)];
    if (d) { state.g.x += d.x; state.g.y += d.y; }
  };
  const step = () => {
    if (!wall(state.p.x + state.next.x, state.p.y + state.next.y)) state.dir = state.next;
    if (!wall(state.p.x + state.dir.x, state.p.y + state.dir.y)) {
      state.p.x += state.dir.x;
      state.p.y += state.dir.y;
    }
    if (state.map[state.p.y][state.p.x] === ".") {
      state.map[state.p.y][state.p.x] = " ";
      state.score += 10;
    }
    moveGhost();
    if (state.p.x === state.g.x && state.p.y === state.g.y) {
      clearInterval(state.timer);
      state.running = false;
      statusEl.textContent = "게임 종료.";
    } else if (!state.map.flat().includes(".")) {
      clearInterval(state.timer);
      state.running = false;
      statusEl.textContent = "클리어!";
    }
    draw();
  };
  const start = () => {
    if (state.running) return;
    state.running = true;
    statusEl.textContent = "진행 중";
    state.timer = setInterval(step, 170);
  };
  document.addEventListener("keydown", (e) => {
    if (!activeHashIs("gamePacman") || isTypingTarget(e.target)) return;
    const dirs = { ArrowLeft: { x: -1, y: 0 }, a: { x: -1, y: 0 }, A: { x: -1, y: 0 }, ArrowRight: { x: 1, y: 0 }, d: { x: 1, y: 0 }, D: { x: 1, y: 0 }, ArrowUp: { x: 0, y: -1 }, w: { x: 0, y: -1 }, W: { x: 0, y: -1 }, ArrowDown: { x: 0, y: 1 }, s: { x: 0, y: 1 }, S: { x: 0, y: 1 } };
    if (!dirs[e.key]) return;
    e.preventDefault();
    state.next = dirs[e.key];
  });
  $("pacmanStartBtn")?.addEventListener("click", start);
  $("pacmanResetBtn")?.addEventListener("click", reset);
  reset();
};

const setupBreakoutGame = () => {
  if (!$("gameBreakout")) return;
  const canvas = $("breakoutCanvas");
  const ctx = canvas.getContext("2d");
  const state = {
    timer: null,
    running: false,
    score: 0,
    paddleX: 125,
    ball: { x: 160, y: 245, vx: 2.2, vy: -2.8, r: 5 },
    keys: { left: false, right: false },
    bricks: []
  };
  const paddle = { w: 70, h: 9, y: 288 };
  const brick = { rows: 5, cols: 8, w: 34, h: 14, gap: 4, top: 42, left: 14 };

  const buildBricks = () => {
    state.bricks = [];
    for (let r = 0; r < brick.rows; r += 1) {
      for (let c = 0; c < brick.cols; c += 1) {
        state.bricks.push({
          x: brick.left + c * (brick.w + brick.gap),
          y: brick.top + r * (brick.h + brick.gap),
          alive: true,
          tone: r
        });
      }
    }
  };

  const updateMeta = () => {
    const left = state.bricks.filter((b) => b.alive).length;
    $("breakoutScore").textContent = `점수 ${state.score} · 남은 벽돌 ${left}`;
  };

  const draw = () => {
    ctx.clearRect(0, 0, 320, 320);
    ctx.fillStyle = "#101722";
    ctx.fillRect(0, 0, 320, 320);
    ctx.fillStyle = "#d8ecff";
    ctx.font = "11px sans-serif";
    ctx.fillText("BRICK BREAKER", 12, 22);

    state.bricks.forEach((b) => {
      if (!b.alive) return;
      const colors = ["#72b8ff", "#62d0a2", "#ffd36a", "#ff9b77", "#c7a4ff"];
      ctx.fillStyle = colors[b.tone % colors.length];
      ctx.fillRect(b.x, b.y, brick.w, brick.h);
      ctx.fillStyle = "rgba(255,255,255,0.35)";
      ctx.fillRect(b.x, b.y, brick.w, 3);
    });

    ctx.fillStyle = "#f5fbff";
    ctx.fillRect(state.paddleX, paddle.y, paddle.w, paddle.h);
    ctx.beginPath();
    ctx.arc(state.ball.x, state.ball.y, state.ball.r, 0, Math.PI * 2);
    ctx.fillStyle = "#ffed8a";
    ctx.fill();
  };

  const stop = (message) => {
    if (state.timer) clearInterval(state.timer);
    state.timer = null;
    state.running = false;
    if (message) $("breakoutStatus").textContent = message;
  };

  const reset = () => {
    stop();
    state.score = 0;
    state.paddleX = 125;
    state.ball = { x: 160, y: 245, vx: 2.2, vy: -2.8, r: 5 };
    buildBricks();
    updateMeta();
    $("breakoutStatus").textContent = "시작을 눌러주세요.";
    draw();
  };

  const step = () => {
    if (state.keys.left) state.paddleX -= 5;
    if (state.keys.right) state.paddleX += 5;
    state.paddleX = Math.max(0, Math.min(320 - paddle.w, state.paddleX));

    const ball = state.ball;
    ball.x += ball.vx;
    ball.y += ball.vy;
    if (ball.x - ball.r <= 0 || ball.x + ball.r >= 320) ball.vx *= -1;
    if (ball.y - ball.r <= 0) ball.vy *= -1;

    if (
      ball.y + ball.r >= paddle.y &&
      ball.y - ball.r <= paddle.y + paddle.h &&
      ball.x >= state.paddleX &&
      ball.x <= state.paddleX + paddle.w &&
      ball.vy > 0
    ) {
      const hit = (ball.x - (state.paddleX + paddle.w / 2)) / (paddle.w / 2);
      ball.vx = hit * 3.8;
      ball.vy = -Math.abs(ball.vy) - 0.03;
    }

    for (const b of state.bricks) {
      if (!b.alive) continue;
      if (
        ball.x + ball.r >= b.x &&
        ball.x - ball.r <= b.x + brick.w &&
        ball.y + ball.r >= b.y &&
        ball.y - ball.r <= b.y + brick.h
      ) {
        b.alive = false;
        state.score += 10;
        ball.vy *= -1;
        break;
      }
    }

    if (ball.y - ball.r > 320) {
      updateMeta();
      draw();
      stop("공을 놓쳤습니다. 리셋 후 다시 도전하세요.");
      return;
    }

    updateMeta();
    draw();
    if (!state.bricks.some((b) => b.alive)) stop("클리어! 모든 벽돌을 깼습니다.");
  };

  const start = () => {
    if (state.timer) return;
    state.running = true;
    $("breakoutStatus").textContent = "진행 중 · 패들을 움직여 공을 받아주세요.";
    state.timer = setInterval(step, 16);
  };

  const movePointer = (clientX) => {
    const rect = canvas.getBoundingClientRect();
    const x = ((clientX - rect.left) / rect.width) * 320;
    state.paddleX = Math.max(0, Math.min(320 - paddle.w, x - paddle.w / 2));
    if (!state.running) draw();
  };

  canvas.addEventListener("pointermove", (e) => movePointer(e.clientX));
  document.addEventListener("keydown", (e) => {
    if (!activeHashIs("gameBreakout") || isTypingTarget(e.target)) return;
    if (["ArrowLeft", "a", "A"].includes(e.key)) {
      e.preventDefault();
      state.keys.left = true;
    }
    if (["ArrowRight", "d", "D"].includes(e.key)) {
      e.preventDefault();
      state.keys.right = true;
    }
  });
  document.addEventListener("keyup", (e) => {
    if (["ArrowLeft", "a", "A"].includes(e.key)) state.keys.left = false;
    if (["ArrowRight", "d", "D"].includes(e.key)) state.keys.right = false;
  });
  $("breakoutStartBtn")?.addEventListener("click", start);
  $("breakoutResetBtn")?.addEventListener("click", reset);
  reset();
};

const setupMemoryGame = () => {
  if (!$("gameMemory")) return;
  const symbols = ["★", "◆", "●", "▲", "■", "♣", "♥", "☀"];
  const state = { deck: [], open: [], matched: new Set(), moves: 0, locked: false };

  const shuffle = (items) => {
    const copy = [...items];
    for (let i = copy.length - 1; i > 0; i -= 1) {
      const j = Math.floor(Math.random() * (i + 1));
      [copy[i], copy[j]] = [copy[j], copy[i]];
    }
    return copy;
  };

  const updateMeta = () => {
    $("memoryScore").textContent = `이동 ${state.moves} · 발견 ${state.matched.size}/8`;
  };

  const render = () => {
    $("memoryBoard").innerHTML = "";
    state.deck.forEach((symbol, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "memory-card";
      button.dataset.index = String(index);
      button.textContent = symbol;
      const isOpen = state.open.includes(index);
      const isMatched = state.matched.has(symbol);
      button.classList.toggle("is-hidden", !isOpen && !isMatched);
      button.classList.toggle("is-matched", isMatched);
      button.disabled = isMatched || state.locked;
      button.addEventListener("click", () => flip(index));
      $("memoryBoard").appendChild(button);
    });
    updateMeta();
  };

  const flip = (index) => {
    if (state.locked || state.open.includes(index) || state.matched.has(state.deck[index])) return;
    state.open.push(index);
    if (state.open.length === 2) {
      state.moves += 1;
      const [a, b] = state.open;
      if (state.deck[a] === state.deck[b]) {
        state.matched.add(state.deck[a]);
        state.open = [];
        $("memoryStatus").textContent = state.matched.size === 8 ? "완료! 모든 쌍을 찾았습니다." : "정답입니다. 다음 쌍을 찾아보세요.";
      } else {
        state.locked = true;
        $("memoryStatus").textContent = "다른 카드입니다. 잠시 후 닫힙니다.";
        setTimeout(() => {
          state.open = [];
          state.locked = false;
          $("memoryStatus").textContent = "카드 2장을 선택하세요.";
          render();
        }, 650);
      }
    }
    render();
  };

  const reset = () => {
    state.deck = shuffle([...symbols, ...symbols]);
    state.open = [];
    state.matched = new Set();
    state.moves = 0;
    state.locked = false;
    $("memoryStatus").textContent = "카드 2장을 선택하세요.";
    render();
  };

  $("memoryResetBtn")?.addEventListener("click", reset);
  reset();
};

const setupMinesweeperGame = () => {
  if (!$("gameMinesweeper")) return;
  const rows = 9;
  const cols = 9;
  const mineCount = 10;
  const state = { mines: null, open: new Set(), flags: new Set(), over: false };
  let longPressTimer = null;

  const keyOf = (r, c) => `${r}:${c}`;
  const parseKey = (key) => key.split(":").map(Number);
  const inBounds = (r, c) => r >= 0 && c >= 0 && r < rows && c < cols;
  const around = (r, c) => {
    const cells = [];
    for (let dr = -1; dr <= 1; dr += 1) {
      for (let dc = -1; dc <= 1; dc += 1) {
        if (!dr && !dc) continue;
        const nr = r + dr;
        const nc = c + dc;
        if (inBounds(nr, nc)) cells.push([nr, nc]);
      }
    }
    return cells;
  };

  const buildMines = (safeR, safeC) => {
    const safe = new Set([keyOf(safeR, safeC), ...around(safeR, safeC).map(([r, c]) => keyOf(r, c))]);
    const mines = new Set();
    while (mines.size < mineCount) {
      const r = Math.floor(Math.random() * rows);
      const c = Math.floor(Math.random() * cols);
      const key = keyOf(r, c);
      if (!safe.has(key)) mines.add(key);
    }
    state.mines = mines;
  };

  const countNear = (r, c) => around(r, c).filter(([nr, nc]) => state.mines?.has(keyOf(nr, nc))).length;

  const updateMeta = () => {
    $("mineScore").textContent = `지뢰 ${mineCount} · 깃발 ${state.flags.size}`;
  };

  const revealAllMines = () => {
    state.mines?.forEach((key) => state.open.add(key));
  };

  const checkWin = () => {
    if (state.open.size === rows * cols - mineCount) {
      state.over = true;
      $("mineStatus").textContent = "클리어! 모든 안전 칸을 열었습니다.";
      return true;
    }
    return false;
  };

  const reveal = (r, c) => {
    if (state.over) return;
    if (!state.mines) buildMines(r, c);
    const start = keyOf(r, c);
    if (state.flags.has(start) || state.open.has(start)) return;
    if (state.mines.has(start)) {
      state.over = true;
      revealAllMines();
      $("mineStatus").textContent = "지뢰를 밟았습니다. 새 게임으로 다시 도전하세요.";
      render();
      return;
    }
    const queue = [[r, c]];
    while (queue.length) {
      const [cr, cc] = queue.shift();
      const current = keyOf(cr, cc);
      if (state.open.has(current) || state.flags.has(current)) continue;
      state.open.add(current);
      if (countNear(cr, cc) === 0) {
        around(cr, cc).forEach(([nr, nc]) => {
          const next = keyOf(nr, nc);
          if (!state.open.has(next) && !state.flags.has(next)) queue.push([nr, nc]);
        });
      }
    }
    $("mineStatus").textContent = "좋습니다. 숫자를 보고 다음 칸을 선택하세요.";
    checkWin();
    render();
  };

  const toggleFlag = (r, c) => {
    if (state.over) return;
    const key = keyOf(r, c);
    if (state.open.has(key)) return;
    if (state.flags.has(key)) state.flags.delete(key);
    else state.flags.add(key);
    $("mineStatus").textContent = state.flags.has(key) ? "깃발을 표시했습니다." : "깃발을 해제했습니다.";
    render();
  };

  const render = () => {
    $("mineBoard").innerHTML = "";
    for (let r = 0; r < rows; r += 1) {
      for (let c = 0; c < cols; c += 1) {
        const key = keyOf(r, c);
        const button = document.createElement("button");
        button.type = "button";
        button.className = "mine-cell";
        button.dataset.key = key;
        const isOpen = state.open.has(key);
        const isFlagged = state.flags.has(key);
        button.classList.toggle("is-open", isOpen);
        button.classList.toggle("is-flagged", isFlagged);
        if (isOpen && state.mines?.has(key)) {
          button.classList.add("is-mine");
          button.textContent = "*";
        } else if (isOpen) {
          const near = countNear(r, c);
          button.textContent = near ? String(near) : "";
        } else if (isFlagged) {
          button.textContent = "⚑";
        }
        button.disabled = state.over && !isFlagged;
        button.addEventListener("click", () => reveal(r, c));
        button.addEventListener("contextmenu", (e) => {
          e.preventDefault();
          toggleFlag(r, c);
        });
        button.addEventListener("pointerdown", () => {
          longPressTimer = setTimeout(() => toggleFlag(r, c), 520);
        });
        button.addEventListener("pointerup", () => clearTimeout(longPressTimer));
        button.addEventListener("pointerleave", () => clearTimeout(longPressTimer));
        $("mineBoard").appendChild(button);
      }
    }
    updateMeta();
  };

  const reset = () => {
    state.mines = null;
    state.open = new Set();
    state.flags = new Set();
    state.over = false;
    $("mineStatus").textContent = "첫 칸을 열어 시작하세요.";
    render();
  };

  $("mineResetBtn")?.addEventListener("click", reset);
  reset();
};

const setupPongGame = () => {
  if (!$("gamePong")) return;
  const canvas = $("pongCanvas");
  const ctx = canvas.getContext("2d");
  const state = {
    timer: null,
    playerY: 130,
    cpuY: 130,
    playerScore: 0,
    cpuScore: 0,
    ball: { x: 160, y: 160, vx: 3, vy: 2.1 },
    keys: { up: false, down: false }
  };
  const paddle = { w: 9, h: 58 };

  const updateMeta = () => {
    $("pongScore").textContent = `플레이어 ${state.playerScore} : ${state.cpuScore} CPU`;
  };

  const resetBall = (dir = 1) => {
    state.ball = { x: 160, y: 160, vx: 3 * dir, vy: (Math.random() > 0.5 ? 1 : -1) * 2.1 };
  };

  const draw = () => {
    ctx.clearRect(0, 0, 320, 320);
    ctx.fillStyle = "#101722";
    ctx.fillRect(0, 0, 320, 320);
    ctx.strokeStyle = "rgba(216,236,255,0.35)";
    ctx.setLineDash([6, 8]);
    ctx.beginPath();
    ctx.moveTo(160, 10);
    ctx.lineTo(160, 310);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = "#f5fbff";
    ctx.fillRect(18, state.playerY, paddle.w, paddle.h);
    ctx.fillRect(293, state.cpuY, paddle.w, paddle.h);
    ctx.beginPath();
    ctx.arc(state.ball.x, state.ball.y, 5, 0, Math.PI * 2);
    ctx.fillStyle = "#72b8ff";
    ctx.fill();
  };

  const stop = (message) => {
    if (state.timer) clearInterval(state.timer);
    state.timer = null;
    if (message) $("pongStatus").textContent = message;
  };

  const reset = () => {
    stop();
    state.playerY = 130;
    state.cpuY = 130;
    state.playerScore = 0;
    state.cpuScore = 0;
    resetBall(Math.random() > 0.5 ? 1 : -1);
    updateMeta();
    $("pongStatus").textContent = "시작을 눌러주세요.";
    draw();
  };

  const score = (playerWonPoint) => {
    if (playerWonPoint) state.playerScore += 1;
    else state.cpuScore += 1;
    updateMeta();
    if (state.playerScore >= 5 || state.cpuScore >= 5) {
      stop(state.playerScore > state.cpuScore ? "승리! 5점을 먼저 달성했습니다." : "CPU 승리. 다시 도전해보세요.");
      return;
    }
    resetBall(playerWonPoint ? -1 : 1);
  };

  const step = () => {
    if (state.keys.up) state.playerY -= 5;
    if (state.keys.down) state.playerY += 5;
    state.playerY = Math.max(0, Math.min(320 - paddle.h, state.playerY));

    const cpuTarget = state.ball.y - paddle.h / 2;
    state.cpuY += Math.max(-3.2, Math.min(3.2, cpuTarget - state.cpuY));
    state.cpuY = Math.max(0, Math.min(320 - paddle.h, state.cpuY));

    const ball = state.ball;
    ball.x += ball.vx;
    ball.y += ball.vy;
    if (ball.y <= 5 || ball.y >= 315) ball.vy *= -1;

    if (ball.x <= 27 && ball.x >= 15 && ball.y >= state.playerY && ball.y <= state.playerY + paddle.h && ball.vx < 0) {
      const hit = (ball.y - (state.playerY + paddle.h / 2)) / (paddle.h / 2);
      ball.vx = Math.abs(ball.vx) + 0.18;
      ball.vy = hit * 3.6;
    }
    if (ball.x >= 293 && ball.x <= 305 && ball.y >= state.cpuY && ball.y <= state.cpuY + paddle.h && ball.vx > 0) {
      const hit = (ball.y - (state.cpuY + paddle.h / 2)) / (paddle.h / 2);
      ball.vx = -Math.abs(ball.vx) - 0.18;
      ball.vy = hit * 3.6;
    }
    if (ball.x < -8) score(false);
    if (ball.x > 328) score(true);
    draw();
  };

  const start = () => {
    if (state.timer) return;
    $("pongStatus").textContent = "진행 중 · 먼저 5점을 달성하세요.";
    state.timer = setInterval(step, 16);
  };

  const movePointer = (clientY) => {
    const rect = canvas.getBoundingClientRect();
    const y = ((clientY - rect.top) / rect.height) * 320;
    state.playerY = Math.max(0, Math.min(320 - paddle.h, y - paddle.h / 2));
    if (!state.timer) draw();
  };

  canvas.addEventListener("pointermove", (e) => movePointer(e.clientY));
  document.addEventListener("keydown", (e) => {
    if (!activeHashIs("gamePong") || isTypingTarget(e.target)) return;
    if (["ArrowUp", "w", "W"].includes(e.key)) {
      e.preventDefault();
      state.keys.up = true;
    }
    if (["ArrowDown", "s", "S"].includes(e.key)) {
      e.preventDefault();
      state.keys.down = true;
    }
  });
  document.addEventListener("keyup", (e) => {
    if (["ArrowUp", "w", "W"].includes(e.key)) state.keys.up = false;
    if (["ArrowDown", "s", "S"].includes(e.key)) state.keys.down = false;
  });
  $("pongStartBtn")?.addEventListener("click", start);
  $("pongResetBtn")?.addEventListener("click", reset);
  reset();
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
  setIconButton("patchNotesBtnIcon", "journal");
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
const pdfFrontRotationCache = new Map();
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
  const cachedRotation = pdfFrontRotationCache.get(cacheKey);
  if (cached && cachedPageCount) {
    return { thumb: cached, pageCount: cachedPageCount, firstPageRotation: cachedRotation || 0 };
  }
  const buffer = await readAsArrayBuffer(file);
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pageCount = pdf.numPages;
  const page = await pdf.getPage(1);
  const firstPageRotation = ((page.rotate % 360) + 360) % 360;
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
  pdfFrontRotationCache.set(cacheKey, firstPageRotation);
  return { thumb, pageCount, firstPageRotation };
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

const mergeFileRotationState = new Map();

const getMergeFileRotation = (file) => {
  const key = getFileCacheKey(file);
  if (!mergeFileRotationState.has(key)) {
    mergeFileRotationState.set(key, { mode: "relative", angle: 0 });
  }
  return mergeFileRotationState.get(key);
};

const getMergePreviewRotation = (file, firstPageRotation) => {
  const rotation = getMergeFileRotation(file);
  return rotation.mode === "absolute"
    ? normalizeArrangeRotation(rotation.angle - firstPageRotation)
    : rotation.angle;
};

const formatMergeRotation = (file) => {
  const rotation = getMergeFileRotation(file);
  return rotation.mode === "absolute"
    ? `${rotation.angle}° 고정`
    : `${rotation.angle}° 상대 회전`;
};

const applyMergeRotation = (page, file) => {
  const rotation = getMergeFileRotation(file);
  if (rotation.mode === "absolute") {
    page.setRotation(PDFLib.degrees(rotation.angle));
    return;
  }
  if (!rotation.angle) return;
  const sourceRotation = page.getRotation().angle || 0;
  page.setRotation(PDFLib.degrees(normalizeArrangeRotation(sourceRotation + rotation.angle)));
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
      const rotationActions =
        stateKey === "mergePdf"
          ? `<div class="thumb-rotate-actions">
              <button class="thumb-rotate merge-file-rotate" type="button" data-rotate="-90" title="파일 전체를 왼쪽으로 90도 회전" aria-label="${file.name} 전체 페이지 왼쪽으로 90도 회전">${ICONS.arrowCounterclockwise}</button>
              <button class="thumb-rotate merge-file-rotate" type="button" data-rotate="90" title="파일 전체를 오른쪽으로 90도 회전" aria-label="${file.name} 전체 페이지 오른쪽으로 90도 회전">${ICONS.arrowClockwise}</button>
            </div>`
          : "";
      item.innerHTML = `<button class="thumb-delete" type="button" title="파일 제거" aria-label="파일 제거">${ICONS.trash3}</button>${rotationActions}<div class="thumb-label">${idx + 1}. ${file.name}</div>`;
      try {
        const { thumb, pageCount, firstPageRotation } = await loadPdfFrontThumb(file, 130);
        const img = document.createElement("img");
        img.src =
          stateKey === "mergePdf"
            ? await renderRotatedThumb(thumb, getMergePreviewRotation(file, firstPageRotation))
            : thumb;
        img.alt = `${file.name} first page`;
        img.draggable = false;
        img.style.width = "100%";
        img.style.border = "1px solid #d4e2f1";
        img.style.borderRadius = "6px";
        item.prepend(img);
        const pageTag = document.createElement("div");
        pageTag.className = "thumb-label";
        pageTag.textContent =
          stateKey === "mergePdf"
            ? `${pageCount}페이지 · ${formatMergeRotation(file)}`
            : `${pageCount}페이지`;
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
  grid.onclick = async (e) => {
    const rotate = e.target.closest(".merge-file-rotate");
    if (rotate && stateKey === "mergePdf") {
      const cell = rotate.closest(".thumb-item");
      if (!cell) return;
      const file = toolFileState.mergePdf[Number(cell.dataset.idx)];
      if (!file) return;
      const rotation = getMergeFileRotation(file);
      rotation.angle = normalizeArrangeRotation(rotation.angle + Number(rotate.dataset.rotate));
      await renderMergePdfPreview(previewId, stateKey, inputId);
      setStatus("mergePdfStatus", `${file.name}: ${formatMergeRotation(file)}`);
      return;
    }
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
    if (e.target.closest("button")) {
      e.preventDefault();
      return;
    }
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
const normalizeArrangeRotation = (angle) => ((angle % 360) + 360) % 360;

const renderRotatedThumb = (sourceDataUrl, rotation) =>
  new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const angle = normalizeArrangeRotation(rotation);
      if (angle === 0) {
        resolve(sourceDataUrl);
        return;
      }
      const swapSides = angle === 90 || angle === 270;
      const canvas = document.createElement("canvas");
      canvas.width = swapSides ? img.height : img.width;
      canvas.height = swapSides ? img.width : img.height;
      const ctx = canvas.getContext("2d");
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate((angle * Math.PI) / 180);
      ctx.drawImage(img, -img.width / 2, -img.height / 2);
      resolve(canvas.toDataURL("image/png"));
    };
    img.onerror = () => reject(new Error("회전 미리보기 생성에 실패했습니다."));
    img.src = sourceDataUrl;
  });

const rotateArrangePages = async (pageNos, delta) => {
  const targets = pageNos
    .map((pageNo) => getArrangePageByNo(pageNo))
    .filter((page) => page && !page.deleted);
  for (const page of targets) {
    page.rotation = normalizeArrangeRotation(page.rotation + delta);
    page.thumbDataUrl = await renderRotatedThumb(page.originalThumbDataUrl, page.rotation);
  }
  rerenderArrangeWorkspace();
  return targets.length;
};

const setArrangePagesAbsoluteRotation = async (pageNos, targetRotation) => {
  const targets = pageNos
    .map((pageNo) => getArrangePageByNo(pageNo))
    .filter((page) => page && !page.deleted);
  for (const page of targets) {
    page.rotation = normalizeArrangeRotation(targetRotation - page.sourceRotation);
    page.thumbDataUrl = await renderRotatedThumb(page.originalThumbDataUrl, page.rotation);
  }
  rerenderArrangeWorkspace();
  return targets.length;
};

const applyArrangeRotation = (copiedPage, pageNo) => {
  const rotation = getArrangePageByNo(pageNo)?.rotation || 0;
  if (!rotation) return;
  const sourceRotation = copiedPage.getRotation().angle || 0;
  copiedPage.setRotation(PDFLib.degrees(normalizeArrangeRotation(sourceRotation + rotation)));
};

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
  item.innerHTML = `
    <button class="thumb-delete" type="button" title="페이지 삭제" aria-label="페이지 ${pageNo} 삭제">${ICONS.trash3}</button>
    <div class="thumb-rotate-actions">
      <button class="thumb-rotate" type="button" data-rotate="-90" title="왼쪽으로 90도 회전" aria-label="페이지 ${pageNo} 왼쪽으로 90도 회전">${ICONS.arrowCounterclockwise}</button>
      <button class="thumb-rotate" type="button" data-rotate="90" title="오른쪽으로 90도 회전" aria-label="페이지 ${pageNo} 오른쪽으로 90도 회전">${ICONS.arrowClockwise}</button>
    </div>
    <div class="thumb-label">p.${pageNo}${page.rotation ? ` · ${page.rotation}°` : ""}</div>
  `;
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
    if (e.target.closest(".thumb-delete, .thumb-rotate")) return;
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

  const onRotateClick = async (e) => {
    const btn = e.target.closest(".thumb-rotate");
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    const item = btn.closest(".thumb-item");
    if (!item) return;
    const pageNo = Number(item.dataset.page);
    const delta = Number(btn.dataset.rotate);
    await rotateArrangePages([pageNo], delta);
    setStatus("arrangePdfStatus", `페이지 ${pageNo}을(를) ${delta < 0 ? "왼쪽" : "오른쪽"}으로 90° 회전했습니다.`);
  };

  const handleDragStart = (e) => {
    if (e.target.closest("button")) {
      e.preventDefault();
      return;
    }
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
    grid.addEventListener("click", onRotateClick);
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
  splitWrap.addEventListener("click", onRotateClick);
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
      arrangeState.pages.push({
        pageNo: i,
        originalThumbDataUrl: thumbDataUrl,
        thumbDataUrl,
        sourceRotation: normalizeArrangeRotation(page.rotate || 0),
        rotation: 0,
        deleted: false,
      });
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
  setIconButton("rotateAllLeftIcon", "arrowCounterclockwise");
  setIconButton("rotateAllRightIcon", "arrowClockwise");
  setIconButton("applyArrangeAbsoluteRotationIcon", "check2");
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

  $("rotateAllLeft").addEventListener("click", async () => {
    const count = await rotateArrangePages(getAvailablePageNos(), -90);
    setStatus("arrangePdfStatus", `전체 ${count}개 페이지를 왼쪽으로 90° 회전했습니다.`);
  });

  $("rotateAllRight").addEventListener("click", async () => {
    const count = await rotateArrangePages(getAvailablePageNos(), 90);
    setStatus("arrangePdfStatus", `전체 ${count}개 페이지를 오른쪽으로 90° 회전했습니다.`);
  });

  $("applyArrangeAbsoluteRotation").addEventListener("click", async () => {
    const targetRotation = Number($("arrangeAbsoluteRotation").value);
    const count = await setArrangePagesAbsoluteRotation(getAvailablePageNos(), targetRotation);
    setStatus("arrangePdfStatus", `전체 ${count}개 페이지 방향을 ${targetRotation}°로 강제 통일했습니다.`);
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
        applyArrangeRotation(p, finalOrder[idx]);
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
        pages.forEach((p, pageIndex) => {
          applyArrangeRotation(p, groups[i][pageIndex]);
          out.addPage(p);
        });
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
  setIconButton("applyMergeAbsoluteRotationIcon", "check2");
  $("mergePdfFiles")?.addEventListener("change", async () => {
    toolFileState.mergePdf = [...$("mergePdfFiles").files];
    toolFileState.mergePdf.forEach((file) => getMergeFileRotation(file));
    await renderMergePdfPreview("mergePdfPreview", "mergePdf", "mergePdfFiles");
  });
  $("applyMergeAbsoluteRotation").addEventListener("click", async () => {
    const files = toolFileState.mergePdf;
    if (!files.length) {
      setStatus("mergePdfStatus", "병합할 PDF를 먼저 선택해주세요.");
      return;
    }
    const targetRotation = Number($("mergeAbsoluteRotation").value);
    files.forEach((file) => {
      const rotation = getMergeFileRotation(file);
      rotation.mode = "absolute";
      rotation.angle = targetRotation;
    });
    await renderMergePdfPreview("mergePdfPreview", "mergePdf", "mergePdfFiles");
    setStatus("mergePdfStatus", `모든 파일의 전체 페이지 방향을 ${targetRotation}°로 강제 통일했습니다.`);
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
        pages.forEach((p) => {
          applyMergeRotation(p, files[i]);
          merged.addPage(p);
        });
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
    const csv = [headers.join(","), ...rows].join("\r\n");
    // UTF-8 BOM을 붙여 Excel에서 한글 깨짐을 방지합니다.
    const blob = new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" });
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
    currentStep: 1,
    selectedType: "url",
    formValues: {},
    previewRequest: 0,
    previewTimer: null,
  };

  const qrPageTypes = {
    url: {
      label: "URL 링크",
      title: "URL 링크 정보",
      fields: [
        { key: "title", label: "페이지 제목", placeholder: "페이지 제목", span: 2 },
        { key: "description", label: "페이지 설명", type: "textarea", placeholder: "페이지 설명", span: 2 },
        { key: "url", label: "웹사이트 URL", placeholder: "https://example.com", required: true, span: 2 },
      ],
    },
    contact: {
      label: "명함",
      title: "명함 정보",
      fields: [
        { key: "name", label: "이름", placeholder: "홍길동", required: true },
        { key: "company", label: "회사", placeholder: "회사명" },
        { key: "phone", label: "전화번호", type: "tel", placeholder: "010-0000-0000", required: true },
        { key: "email", label: "이메일", type: "email", placeholder: "name@example.com" },
        { key: "url", label: "웹사이트", placeholder: "https://example.com", span: 2 },
      ],
    },
    menu: {
      label: "메뉴판",
      title: "메뉴판 정보",
      fields: [
        { key: "title", label: "메뉴판 제목", placeholder: "메뉴판", required: true, span: 2 },
        { key: "description", label: "설명", type: "textarea", placeholder: "메뉴 또는 매장 설명", span: 2 },
        { key: "url", label: "메뉴 링크", placeholder: "https://example.com/menu", required: true, span: 2 },
      ],
    },
    invite: {
      label: "초대장",
      title: "초대장 정보",
      fields: [
        { key: "title", label: "행사명", placeholder: "행사명", required: true, span: 2 },
        { key: "date", label: "날짜", type: "date", required: true },
        { key: "time", label: "시간", type: "time" },
        { key: "location", label: "장소", placeholder: "행사 장소", span: 2 },
        { key: "details", label: "안내 내용", type: "textarea", placeholder: "행사 안내", span: 2 },
      ],
    },
    wifi: {
      label: "와이파이",
      title: "와이파이 정보",
      fields: [
        { key: "ssid", label: "네트워크 이름", placeholder: "Wi-Fi 이름", required: true, span: 2 },
        { key: "security", label: "보안 방식", type: "select", options: [["WPA", "WPA/WPA2"], ["WEP", "WEP"], ["nopass", "비밀번호 없음"]] },
        { key: "password", label: "비밀번호", type: "password", placeholder: "Wi-Fi 비밀번호" },
        { key: "hidden", label: "숨김 네트워크", type: "checkbox", span: 2 },
      ],
    },
    coupon: {
      label: "쿠폰",
      title: "쿠폰 정보",
      fields: [
        { key: "title", label: "쿠폰명", placeholder: "쿠폰명", required: true, span: 2 },
        { key: "code", label: "쿠폰 코드", placeholder: "COUPON-001", required: true },
        { key: "expiry", label: "사용 기한", type: "date" },
        { key: "url", label: "관련 링크", placeholder: "https://example.com", span: 2 },
      ],
    },
    guide: {
      label: "안내문",
      title: "안내문 정보",
      fields: [
        { key: "title", label: "제목", placeholder: "안내 제목", required: true, span: 2 },
        { key: "details", label: "안내 내용", type: "textarea", placeholder: "안내할 내용을 입력하세요.", required: true, span: 2 },
        { key: "url", label: "관련 링크", placeholder: "https://example.com", span: 2 },
      ],
    },
    custom: {
      label: "직접 제작",
      title: "직접 제작 정보",
      fields: [
        { key: "content", label: "QR 내용", type: "textarea", placeholder: "QR 코드에 담을 내용을 입력하세요.", required: true, span: 2 },
      ],
    },
  };

  const normalizeQrUrl = (value) => {
    const url = String(value || "").trim();
    if (!url || /^[a-z][a-z0-9+.-]*:/i.test(url)) return url;
    return `https://${url}`;
  };

  const escapeVcard = (value) => String(value || "").replace(/([\\;,])/g, "\\$1").replace(/\n/g, "\\n");
  const escapeWifi = (value) => String(value || "").replace(/([\\;,:"])/g, "\\$1");

  const buildQrPayload = () => {
    const values = state.formValues[state.selectedType] || {};
    const config = qrPageTypes[state.selectedType];
    if (config.fields.some((field) => field.required && !String(values[field.key] || "").trim())) return "";
    switch (state.selectedType) {
      case "url":
      case "menu":
        return normalizeQrUrl(values.url);
      case "contact":
        return [
          "BEGIN:VCARD",
          "VERSION:3.0",
          `FN:${escapeVcard(values.name)}`,
          values.company ? `ORG:${escapeVcard(values.company)}` : "",
          values.phone ? `TEL:${escapeVcard(values.phone)}` : "",
          values.email ? `EMAIL:${escapeVcard(values.email)}` : "",
          values.url ? `URL:${normalizeQrUrl(values.url)}` : "",
          "END:VCARD",
        ].filter(Boolean).join("\n");
      case "invite": {
        const date = String(values.date || "").replace(/-/g, "");
        const time = String(values.time || "").replace(/:/g, "") || "0000";
        return [
          "BEGIN:VEVENT",
          `SUMMARY:${escapeVcard(values.title)}`,
          date ? `DTSTART:${date}T${time}00` : "",
          values.location ? `LOCATION:${escapeVcard(values.location)}` : "",
          values.details ? `DESCRIPTION:${escapeVcard(values.details)}` : "",
          "END:VEVENT",
        ].filter(Boolean).join("\n");
      }
      case "wifi":
        return `WIFI:T:${values.security || "WPA"};S:${escapeWifi(values.ssid)};P:${escapeWifi(values.password)};H:${values.hidden ? "true" : "false"};;`;
      case "coupon":
        return [
          values.title,
          values.code ? `쿠폰 코드: ${values.code}` : "",
          values.expiry ? `사용 기한: ${values.expiry}` : "",
          values.url ? normalizeQrUrl(values.url) : "",
        ].filter(Boolean).join("\n");
      case "guide":
        return [values.title, values.details, values.url ? normalizeQrUrl(values.url) : ""].filter(Boolean).join("\n");
      case "custom":
        return String(values.content || "").trim();
      default:
        return "";
    }
  };

  const validateQrFields = () => {
    const config = qrPageTypes[state.selectedType];
    const values = state.formValues[state.selectedType] || {};
    const missing = config.fields.find((field) => field.required && !String(values[field.key] || "").trim());
    return missing ? `${missing.label}을(를) 입력해주세요.` : "";
  };

  const renderQrFields = () => {
    const config = qrPageTypes[state.selectedType];
    const values = state.formValues[state.selectedType] || {};
    state.formValues[state.selectedType] = values;
    $("qrFieldsTitle").textContent = config.title;
    $("qrPreviewType").textContent = config.label;
    const box = $("qrPageFields");
    box.innerHTML = "";

    config.fields.forEach((field) => {
      const label = document.createElement("label");
      if (field.span === 2) label.classList.add("span-2");

      if (field.type === "checkbox") {
        label.classList.add("inline-check");
        const input = document.createElement("input");
        input.type = "checkbox";
        input.checked = !!values[field.key];
        input.addEventListener("change", () => {
          values[field.key] = input.checked;
          syncQrPayload();
        });
        label.append(input, document.createTextNode(field.label));
        box.appendChild(label);
        return;
      }

      label.appendChild(document.createTextNode(field.label));
      let input;
      if (field.type === "textarea") {
        input = document.createElement("textarea");
      } else if (field.type === "select") {
        input = document.createElement("select");
        field.options.forEach(([value, text]) => {
          const option = document.createElement("option");
          option.value = value;
          option.textContent = text;
          input.appendChild(option);
        });
      } else {
        input = document.createElement("input");
        input.type = field.type || "text";
      }
      input.id = `qrField-${field.key}`;
      if (field.placeholder) input.placeholder = field.placeholder;
      input.required = !!field.required;
      input.value = values[field.key] || (field.key === "security" ? "WPA" : "");
      values[field.key] = input.value;
      input.addEventListener("input", () => {
        values[field.key] = input.value;
        syncQrPayload();
      });
      input.addEventListener("change", () => {
        values[field.key] = input.value;
        syncQrPayload();
      });
      label.appendChild(input);
      box.appendChild(label);
    });
    syncQrPayload();
  };

  const updateQrStepUi = () => {
    document.querySelectorAll("[data-qr-step-target]").forEach((button) => {
      const step = Number(button.dataset.qrStepTarget);
      button.classList.toggle("is-active", step === state.currentStep);
      button.classList.toggle("is-complete", step < state.currentStep);
      if (step === state.currentStep) button.setAttribute("aria-current", "step");
      else button.removeAttribute("aria-current");
    });
    document.querySelectorAll("[data-qr-step-panel]").forEach((panel) => {
      const active = Number(panel.dataset.qrStepPanel) === state.currentStep;
      panel.hidden = !active;
      panel.classList.toggle("is-active", active);
    });
    $("qrPrevStep").hidden = state.currentStep === 1;
    $("qrNextStep").hidden = state.currentStep === 3;
    $("qrFinalActions").hidden = state.currentStep !== 3;
  };

  const goToQrStep = (step) => {
    state.currentStep = Math.max(1, Math.min(3, Number(step) || 1));
    if (state.currentStep === 3) renderQrFields();
    updateQrStepUi();
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
    img.alt = `${qrPageTypes[state.selectedType].label} QR 코드 미리보기`;
    box.appendChild(img);
  };

  const renderQrPreviewEmpty = () => {
    $("qrPreview").innerHTML = '<div class="qr-preview-empty"><span aria-hidden="true">▦</span><p>정보를 입력하면 QR 코드가 표시됩니다.</p></div>';
  };

  const refreshQrPreviewLive = async () => {
    const options = getQrOptions();
    if (!options.text) {
      state.previewRequest += 1;
      if (state.renderUrl?.startsWith("blob:")) URL.revokeObjectURL(state.renderUrl);
      state.lastBlob = null;
      state.renderUrl = null;
      renderQrPreviewEmpty();
      return;
    }
    const request = ++state.previewRequest;
    try {
      const asset = await buildQrAsset(options);
      if (request !== state.previewRequest) {
        if (asset.previewUrl.startsWith("blob:")) URL.revokeObjectURL(asset.previewUrl);
        return;
      }
      if (state.renderUrl?.startsWith("blob:")) URL.revokeObjectURL(state.renderUrl);
      state.lastBlob = asset.blob;
      state.lastExt = asset.ext;
      state.renderUrl = asset.previewUrl;
      renderQrPreview(asset.previewUrl);
    } catch {
      // Ignore live-preview errors; explicit "QR 생성" will surface details.
    }
  };

  const queueQrPreview = () => {
    clearTimeout(state.previewTimer);
    state.previewTimer = setTimeout(refreshQrPreviewLive, 180);
  };

  const syncQrPayload = () => {
    const payload = buildQrPayload();
    $("qrInput").value = payload;
    const values = state.formValues[state.selectedType] || {};
    const heading = values.title || values.name || values.ssid || qrPageTypes[state.selectedType].label;
    const excerpt = payload.replace(/\n/g, " · ").slice(0, 180);
    $("qrPayloadPreview").textContent = payload ? `${heading} · ${excerpt}` : "연결할 정보를 입력해주세요.";
    queueQrPreview();
  };

  ["qrFg", "qrBg", "qrTransparentBg", "qrSize", "qrMargin"].forEach((id) => {
    $(id)?.addEventListener("input", () => {
      syncQrColorUi();
      document.querySelectorAll("[data-qr-preset]").forEach((button) => {
        button.classList.remove("is-selected");
        button.setAttribute("aria-pressed", "false");
      });
      queueQrPreview();
    });
    $(id)?.addEventListener("change", () => {
      syncQrColorUi();
      queueQrPreview();
    });
  });

  document.querySelectorAll("input[name='qrFormat']").forEach((input) => {
    input.addEventListener("change", queueQrPreview);
  });

  const qrPresets = {
    basic: { fg: "#000000", bg: "#ffffff", margin: 40, transparent: false },
    navy: { fg: "#173c68", bg: "#ffffff", margin: 48, transparent: false },
    green: { fg: "#126b50", bg: "#f6fff9", margin: 40, transparent: false },
    reverse: { fg: "#ffffff", bg: "#1d2a38", margin: 48, transparent: false },
  };

  document.querySelectorAll("[data-qr-preset]").forEach((button) => {
    button.addEventListener("click", () => {
      const preset = qrPresets[button.dataset.qrPreset];
      $("qrFg").value = preset.fg;
      $("qrBg").value = preset.bg;
      $("qrMargin").value = preset.margin;
      $("qrTransparentBg").checked = preset.transparent;
      document.querySelectorAll("[data-qr-preset]").forEach((item) => {
        const selected = item === button;
        item.classList.toggle("is-selected", selected);
        item.setAttribute("aria-pressed", String(selected));
      });
      syncQrColorUi();
      queueQrPreview();
    });
  });

  document.querySelectorAll("[data-qr-page-type]").forEach((button) => {
    const icon = button.querySelector(".qr-type-icon");
    if (icon && ICONS[button.dataset.qrIcon]) icon.innerHTML = ICONS[button.dataset.qrIcon];
    button.addEventListener("click", () => {
      state.selectedType = button.dataset.qrPageType;
      document.querySelectorAll("[data-qr-page-type]").forEach((item) => {
        const selected = item === button;
        item.classList.toggle("is-selected", selected);
        item.setAttribute("aria-checked", String(selected));
      });
      $("qrPreviewType").textContent = qrPageTypes[state.selectedType].label;
      syncQrPayload();
    });
  });

  document.querySelectorAll("[data-qr-step-target]").forEach((button) => {
    button.addEventListener("click", () => goToQrStep(button.dataset.qrStepTarget));
  });
  $("qrPrevStep").addEventListener("click", () => goToQrStep(state.currentStep - 1));
  $("qrNextStep").addEventListener("click", () => goToQrStep(state.currentStep + 1));
  $("qrPrevIcon").innerHTML = ICONS.arrowLeft;
  $("qrNextIcon").innerHTML = ICONS.arrowRight;
  syncQrColorUi();
  state.formValues.url = { title: "", description: "", url: "" };
  renderQrFields();
  updateQrStepUi();

  $("runQr").addEventListener("click", async () => {
    const options = getQrOptions();
    const validationMessage = validateQrFields();
    if (validationMessage || !options.text) {
      setStatus("qrStatus", validationMessage || "QR 내용을 입력해주세요.");
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

  $("saveQr").addEventListener("click", async () => {
    const options = getQrOptions();
    const validationMessage = validateQrFields();
    if (validationMessage || !options.text) {
      setStatus("qrStatus", validationMessage || "QR 내용을 입력해주세요.");
      return;
    }
    beginGlobalBusy("QR 파일을 준비 중입니다...");
    try {
      const asset = await buildQrAsset(options);
      if (state.renderUrl?.startsWith("blob:")) URL.revokeObjectURL(state.renderUrl);
      state.lastBlob = asset.blob;
      state.lastExt = asset.ext;
      state.renderUrl = asset.previewUrl;
      renderQrPreview(asset.previewUrl);
      downloadBlob(asset.blob, `qrcode.${asset.ext}`);
      setStatus("qrStatus", `QR ${asset.ext.toUpperCase()} 저장 완료`);
    } catch (err) {
      setStatus("qrStatus", `QR 저장 오류: ${err.message}`);
    } finally {
      endGlobalBusy();
    }
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

const setupHangulWebEditor = () => {
  const fileInput = $("hangulEditorFile");
  const formatSelect = $("hangulConvertFormat");
  const convertBtn = $("runHangulConvert");
  const previewBox = $("hangulFilePreview");
  const openWebEditorBtn = $("openRhwpWebEditor");
  const editorStatus = $("hangulEditorStatus");
  if (!fileInput || !formatSelect || !convertBtn || !previewBox || !openWebEditorBtn || !editorStatus) return;

  let activeFile = null;
  let rhwpMod = null;
  let rhwpReady = false;
  let extractedText = "";

  const normalizeExtractedText = (raw) => {
    const src = String(raw || "").replace(/\r/g, "");
    if (!src.trim()) return "";
    const lines = src.split("\n");
    const nonEmpty = lines.filter((l) => l.trim().length > 0);
    if (!nonEmpty.length) return src.trim();
    const shortCount = nonEmpty.filter((l) => l.trim().length <= 1).length;
    const shortRatio = shortCount / nonEmpty.length;
    if (shortRatio < 0.6) return src.trim();

    // 글자 단위 줄바꿈으로 추출된 경우 문장 단위로 재조합
    const compact = nonEmpty.map((l) => l.trim()).join("");
    return compact
      .replace(/([.!?])\s*/g, "$1\n")
      .replace(/([다요죠니다임])\s*(?=[A-Za-z0-9가-힣])/g, "$1\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  };

  const ensureRhwp = async () => {
    if (rhwpReady && rhwpMod) return rhwpMod;
    if (!rhwpMod) rhwpMod = await import("https://cdn.jsdelivr.net/npm/@rhwp/core@0.7.10/rhwp.js");
    const initFn = rhwpMod.default || rhwpMod.init;
    if (typeof initFn !== "function" || typeof rhwpMod.HwpDocument !== "function") throw new Error("rhwp core load failed");
    if (!globalThis.measureTextWidth) {
      globalThis.measureTextWidth = (font, text) => {
        const c = document.createElement("canvas");
        const ctx = c.getContext("2d");
        ctx.font = font || "14px sans-serif";
        return ctx.measureText(text || "").width;
      };
    }
    await initFn({ module_or_path: "https://cdn.jsdelivr.net/npm/@rhwp/core@0.7.10/rhwp_bg.wasm" });
    rhwpReady = true;
    return rhwpMod;
  };

  const readTextFromHwpx = async (file) => {
    const zip = await JSZip.loadAsync(await readAsArrayBuffer(file));
    const prv = zip.file("Preview/PrvText.txt");
    if (prv) {
      const txt = (await prv.async("string")).trim();
      if (txt) return txt;
    }
    const sections = Object.keys(zip.files).filter((n) => /^Contents\/section\d+\.xml$/i.test(n)).sort();
    if (!sections.length) throw new Error("HWPX section file not found");
    const parser = new DOMParser();
    const out = [];
    for (const s of sections) {
      const xml = await zip.file(s).async("string");
      const doc = parser.parseFromString(xml, "application/xml");
      const lines = [...doc.getElementsByTagName("*")]
        .filter((el) => el.localName === "t")
        .map((el) => (el.textContent || "").trim())
        .filter(Boolean);
      if (lines.length) out.push(lines.join("\n"));
    }
    return out.join("\n\n").trim();
  };

  const readTextFromHwp = async (file) => {
    const mod = await ensureRhwp();
    const doc = new mod.HwpDocument(new Uint8Array(await readAsArrayBuffer(file)));
    const parser = new DOMParser();
    const parts = [];
    for (let i = 0; i < 300; i += 1) {
      let svg = "";
      try {
        svg = doc.renderPageSvg(i);
      } catch {
        break;
      }
      if (!svg) break;
      const xml = parser.parseFromString(svg, "image/svg+xml");
      const lines = [...xml.querySelectorAll("text")]
        .map((n) => (n.textContent || "").replace(/\s+/g, " ").trim())
        .filter(Boolean);
      if (lines.length) parts.push(lines.join("\n"));
    }
    const merged = parts.join("\n\n").trim();
    if (!merged) throw new Error("HWP text extraction failed");
    return normalizeExtractedText(merged);
  };

  const collectRenderedPageSvgs = async (file) => {
    const mod = await ensureRhwp();
    const doc = new mod.HwpDocument(new Uint8Array(await readAsArrayBuffer(file)));
    const pages = [];

    const tryCollect = (startIndex) => {
      for (let i = startIndex; i < startIndex + 2000; i += 1) {
        let svg = "";
        try {
          svg = doc.renderPageSvg(i);
        } catch {
          break;
        }
        if (!svg) break;
        pages.push(svg);
      }
    };

    tryCollect(0);
    if (!pages.length) tryCollect(1);
    return pages;
  };

  const renderPreview = async (file) => {
    previewBox.innerHTML = '<p class="status">미리보기 생성 중...</p>';
    try {
      const pages = await collectRenderedPageSvgs(file);
      const svg = pages[0];
      if (!svg) throw new Error("first page render failed");
      previewBox.innerHTML = "";
      const meta = document.createElement("p");
      meta.className = "status";
      meta.textContent = `미리보기: ${file.name} (1페이지)`;
      const wrap = document.createElement("div");
      wrap.className = "hangul-preview-svg";
      wrap.innerHTML = svg;
      previewBox.append(meta, wrap);
    } catch (err) {
      previewBox.innerHTML = `<p class="status">미리보기 실패: ${err.message}</p>`;
    }
  };

  const extractText = async (file) => {
    const ext = (file.name.split(".").pop() || "").toLowerCase();
    if (ext === "hwpx") return readTextFromHwpx(file);
    if (ext === "hwp") return readTextFromHwp(file);
    throw new Error("지원 확장자는 .hwp/.hwpx 입니다.");
  };

  const loadFile = async (file) => {
    if (!file) return;
    activeFile = file;
    setStatus("hangulEditorStatus", `파일 선택됨: ${file.name}`);
    try {
      extractedText = normalizeExtractedText(await extractText(file));
      setStatus("hangulEditorStatus", `문서 로드 완료: ${file.name}`);
    } catch (err) {
      extractedText = "";
      setStatus("hangulEditorStatus", `문서 로드 실패: ${err.message}`);
    }
    await renderPreview(file);
  };

  fileInput.addEventListener("change", () => loadFile(fileInput.files?.[0] || null));

  convertBtn.addEventListener("click", async () => {
    if (!activeFile) {
      setStatus("hangulConvertStatus", "먼저 HWP/HWPX 파일을 넣어 주세요.");
      return;
    }
    beginGlobalBusy("문서 변환 중...");
    try {
      const text = extractedText?.trim() ? extractedText : await extractText(activeFile);
      const stem = activeFile.name.replace(/\.[^.]+$/, "");
      const fmt = formatSelect.value;
      if (fmt === "txt") {
        downloadBlob(new Blob([text], { type: "text/plain;charset=utf-8" }), `${stem}.txt`);
      } else if (fmt === "md") {
        downloadBlob(new Blob([text], { type: "text/markdown;charset=utf-8" }), `${stem}.md`);
      } else if (fmt === "html") {
        const escaped = String(text || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;");
        const htmlDoc = `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${stem}</title>
  <style>
    body { margin: 24px; font-family: "Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",sans-serif; line-height: 1.7; color: #111; }
    .doc { white-space: pre-wrap; word-break: break-word; }
  </style>
</head>
<body>
  <div class="doc">${escaped}</div>
</body>
</html>`;
        downloadBlob(new Blob([htmlDoc], { type: "text/html;charset=utf-8" }), `${stem}.html`);
      } else if (fmt === "csv") {
        const rows = text.split(/\r?\n/).map((l) => `"${l.replaceAll('"', '""')}"`);
        downloadBlob(new Blob([`line\n${rows.join("\n")}\n`], { type: "text/csv;charset=utf-8" }), `${stem}.csv`);
      } else if (fmt === "xlsx") {
        if (!globalThis.XLSX) throw new Error("XLSX 라이브러리 로드 실패");
        const rows = String(text || "")
          .split(/\r?\n/)
          .map((line, idx) => [idx + 1, line]);
        const aoa = [["No", "내용"], ...rows];
        const ws = globalThis.XLSX.utils.aoa_to_sheet(aoa);
        ws["!cols"] = [{ wch: 8 }, { wch: 80 }];
        const wb = globalThis.XLSX.utils.book_new();
        globalThis.XLSX.utils.book_append_sheet(wb, ws, "문서");
        const out = globalThis.XLSX.write(wb, { bookType: "xlsx", type: "array" });
        downloadBlob(
          new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
          `${stem}.xlsx`,
        );
      } else if (fmt === "pdf") {
        const pageSvgs = await collectRenderedPageSvgs(activeFile);

        const parseSvgSize = (svgText) => {
          const viewBoxMatch = svgText.match(/viewBox\s*=\s*"([^"]+)"/i);
          if (viewBoxMatch) {
            const nums = viewBoxMatch[1].trim().split(/\s+/).map(Number);
            if (nums.length === 4 && Number.isFinite(nums[2]) && Number.isFinite(nums[3]) && nums[2] > 0 && nums[3] > 0) {
              return { width: nums[2], height: nums[3] };
            }
          }
          const widthMatch = svgText.match(/width\s*=\s*"([\d.]+)(px)?"/i);
          const heightMatch = svgText.match(/height\s*=\s*"([\d.]+)(px)?"/i);
          const w = widthMatch ? Number(widthMatch[1]) : 595.28;
          const h = heightMatch ? Number(heightMatch[1]) : 841.89;
          return {
            width: Number.isFinite(w) && w > 0 ? w : 595.28,
            height: Number.isFinite(h) && h > 0 ? h : 841.89,
          };
        };

        const svgToPngDataUrl = (svgText, width, height) =>
          new Promise((resolve, reject) => {
            const blob = new Blob([svgText], { type: "image/svg+xml;charset=utf-8" });
            const url = URL.createObjectURL(blob);
            const img = new Image();
            img.onload = () => {
              try {
                const canvas = document.createElement("canvas");
                const scale = 2;
                canvas.width = Math.max(1, Math.round(width * scale));
                canvas.height = Math.max(1, Math.round(height * scale));
                const ctx = canvas.getContext("2d");
                ctx.fillStyle = "#ffffff";
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
                resolve(canvas.toDataURL("image/png"));
              } catch (e) {
                reject(e);
              } finally {
                URL.revokeObjectURL(url);
              }
            };
            img.onerror = () => {
              URL.revokeObjectURL(url);
              reject(new Error("SVG 렌더링 실패"));
            };
            img.src = url;
          });

        if (pageSvgs.length) {
          const pdfDoc = await PDFLib.PDFDocument.create();
          for (let i = 0; i < pageSvgs.length; i += 1) {
            setStatus("hangulConvertStatus", `PDF 렌더링 중... (${i + 1}/${pageSvgs.length})`);
            const svg = pageSvgs[i];
            const { width, height } = parseSvgSize(svg);
            const pngDataUrl = await svgToPngDataUrl(svg, width, height);
            const pngImage = await pdfDoc.embedPng(pngDataUrl);
            const page = pdfDoc.addPage([width, height]);
            page.drawImage(pngImage, { x: 0, y: 0, width, height });
          }
          downloadBlob(new Blob([await pdfDoc.save()], { type: "application/pdf" }), `${stem}.pdf`);
        } else {
          // 렌더 실패 시 텍스트 기반 PDF로 폴백
          const fallbackText = normalizeExtractedText(text);
          const pageW = 595.28;
          const pageH = 841.89;
          const renderW = 1240;
          const renderH = 1754;
          const margin = 72;
          const fontPx = 24;
          const lineH = 36;
          const maxWidth = renderW - margin * 2;

          const wrapLine = (ctx, line) => {
            const words = String(line || "").split(/\s+/);
            if (!words.length) return [""];
            const out = [];
            let cur = words[0] || "";
            for (let i = 1; i < words.length; i += 1) {
              const next = `${cur} ${words[i]}`;
              if (ctx.measureText(next).width <= maxWidth) cur = next;
              else {
                out.push(cur);
                cur = words[i];
              }
            }
            out.push(cur);
            return out;
          };

          const canvas = document.createElement("canvas");
          canvas.width = renderW;
          canvas.height = renderH;
          const ctx = canvas.getContext("2d");
          ctx.textBaseline = "top";
          ctx.font = `${fontPx}px "Noto Sans KR", sans-serif`;

          const lines = [];
          fallbackText.split(/\r?\n/).forEach((line) => {
            const wrapped = wrapLine(ctx, line);
            wrapped.forEach((w) => lines.push(w));
            if (!wrapped.length) lines.push("");
          });

          const linesPerPage = Math.max(1, Math.floor((renderH - margin * 2) / lineH));
          const chunks = [];
          for (let i = 0; i < lines.length; i += linesPerPage) chunks.push(lines.slice(i, i + linesPerPage));
          if (!chunks.length) chunks.push([""]);

          const pdfDoc = await PDFLib.PDFDocument.create();
          for (const chunk of chunks) {
            ctx.fillStyle = "#ffffff";
            ctx.fillRect(0, 0, renderW, renderH);
            ctx.fillStyle = "#111111";
            let y = margin;
            chunk.forEach((line) => {
              ctx.fillText(line, margin, y);
              y += lineH;
            });
            const png = canvas.toDataURL("image/png");
            const pngImage = await pdfDoc.embedPng(png);
            const page = pdfDoc.addPage([pageW, pageH]);
            page.drawImage(pngImage, { x: 0, y: 0, width: pageW, height: pageH });
          }
          downloadBlob(new Blob([await pdfDoc.save()], { type: "application/pdf" }), `${stem}.pdf`);
        }
      }
      setStatus("hangulConvertStatus", `변환 완료: ${stem}.${fmt}`);
    } catch (err) {
      setStatus("hangulConvertStatus", `변환 실패: ${getErrorMessage(err)}`);
    } finally {
      endGlobalBusy();
    }
  });

  openWebEditorBtn.addEventListener("click", () => {
    window.open("https://edwardkim.github.io/rhwp/", "_blank", "noopener,noreferrer");
    if (activeFile) {
      setStatus("hangulEditorStatus", `웹에디터 새창 열림. 동일 파일(${activeFile.name})을 새창에 드래그앤드롭하세요.`);
      return;
    }
    setStatus("hangulEditorStatus", "웹에디터 새창 열림. 파일을 새창에 드래그앤드롭하세요.");
  });

};

const init = () => {
  document.body.classList.add("home-mode");
  initOperations();
  setupThemeToggle();
  setupLoginModal();
  setupLogoGameEntry();
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
  setupHangulWebEditor();
  setupYachtGame();
  setupRpsGame();
  setupTttGame();
  setupUpdownGame();
  setup2048Game();
  setupSnakeGame();
  setupTetrisGame();
  setupPacmanGame();
  setupBreakoutGame();
  setupMemoryGame();
  setupMinesweeperGame();
  setupPongGame();
};

init();

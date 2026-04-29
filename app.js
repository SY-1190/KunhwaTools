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
};

const operationStates = {};

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
    button.textContent = isDark ? "🌙 다크 모드" : "☀ 화이트 모드";
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
  $("runPdfToImage").addEventListener("click", async () => {
    const file = $("pdfToImageFile").files[0];
    const format = $("pdfToImageFormat").value;
    const scale = Number($("pdfToImageScale").value);
    if (!file) {
      setStatus("pdfToImageStatus", "PDF 파일을 선택해주세요.");
      return;
    }

    startOperation("pdfToImage", "PDF 로딩 중...");
    try {
      const buffer = await readAsArrayBuffer(file);
      checkCancelled("pdfToImage");
      const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
      const zip = new JSZip();

      for (let i = 1; i <= pdf.numPages; i += 1) {
        checkCancelled("pdfToImage");
        setStatus("pdfToImageStatus", `페이지 변환 중 (${i}/${pdf.numPages})`);
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement("canvas");
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({
          canvasContext: canvas.getContext("2d"),
          viewport,
        }).promise;
        const mime = format === "png" ? "image/png" : "image/jpeg";
        const base64 = canvas.toDataURL(mime, 0.95).split(",")[1];
        zip.file(`page-${i}.${format === "png" ? "png" : "jpg"}`, base64, { base64: true });
        updateProgress("pdfToImage", i, pdf.numPages);
      }

      checkCancelled("pdfToImage");
      const blob = await zip.generateAsync({ type: "blob" });
      downloadBlob(blob, "pdf-to-images.zip");
      updateProgress("pdfToImage", 100, 100);
      endOperation("pdfToImage", `완료: ${pdf.numPages}페이지 (${formatBytes(blob.size)})`);
    } catch (err) {
      handleOperationError("pdfToImage", err);
    }
  });
};

const setupImageToPdf = () => {
  if (!$("runImageToPdf")) return;
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
  pageOrder: [],
  renderDoc: null,
  file: null,
  splitSelectMode: false,
  splitBlocks: [],
};

const getThumbItems = () => [...$("pdfThumbGrid").querySelectorAll(".thumb-item")];

const clearThumbSelection = () => {
  getThumbItems().forEach((el) => el.classList.remove("selected-range"));
};

const refreshArrangeOrderText = () => {
  const order = getThumbItems().map((el) => Number(el.dataset.page));
  arrangeState.pageOrder = order;
  $("arrangeOrderText").textContent = order.length ? order.join(", ") : "-";
};

const renderSplitBlocks = () => {
  const box = $("splitBlocksList");
  if (!arrangeState.splitBlocks.length) {
    box.innerHTML = `<div class="split-block-item"><span class="block-pages">등록된 블록 없음</span></div>`;
    return;
  }
  box.innerHTML = arrangeState.splitBlocks
    .map(
      (block, idx) => `
        <div class="split-block-item" data-block-index="${idx}">
          <span class="block-pages">블록 ${idx + 1}: ${block.join(", ")}</span>
          <button class="remove-block" data-block-index="${idx}" type="button">삭제</button>
        </div>
      `
    )
    .join("");
};

const setSplitSelectionMode = (enabled) => {
  arrangeState.splitSelectMode = enabled;
  $("toggleSplitSelectMode").textContent = `분할 선택 모드: ${enabled ? "ON" : "OFF"}`;
  getThumbItems().forEach((item) => {
    item.draggable = !enabled;
  });
  if (!enabled) clearThumbSelection();
};

const applyRangeSelection = (aIdx, bIdx) => {
  const items = getThumbItems();
  const start = Math.min(aIdx, bIdx);
  const end = Math.max(aIdx, bIdx);
  items.forEach((item, idx) => {
    item.classList.toggle("selected-range", idx >= start && idx <= end);
  });
};

const setupThumbDnD = () => {
  const grid = $("pdfThumbGrid");
  let dragging = null;
  let selecting = false;
  let anchorIndex = -1;

  const getDragAfterElement = (container, x, y) => {
    const items = [...container.querySelectorAll(".thumb-item:not(.dragging)")];
    return items.reduce(
      (closest, child) => {
        const box = child.getBoundingClientRect();
        const centerY = box.top + box.height / 2;
        const centerX = box.left + box.width / 2;
        const offset = y - centerY + (x - centerX) * 0.1;
        if (offset < 0 && offset > closest.offset) return { offset, element: child };
        return closest;
      },
      { offset: Number.NEGATIVE_INFINITY, element: null }
    ).element;
  };

  grid.addEventListener("dragstart", (e) => {
    if (arrangeState.splitSelectMode) {
      e.preventDefault();
      return;
    }
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    dragging = item;
    item.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  });

  grid.addEventListener("dragover", (e) => {
    if (!dragging || arrangeState.splitSelectMode) return;
    e.preventDefault();
    const afterElement = getDragAfterElement(grid, e.clientX, e.clientY);
    if (afterElement == null) grid.appendChild(dragging);
    else grid.insertBefore(dragging, afterElement);
  });

  grid.addEventListener("drop", (e) => {
    if (!dragging || arrangeState.splitSelectMode) return;
    e.preventDefault();
    refreshArrangeOrderText();
  });

  grid.addEventListener("dragend", () => {
    if (!dragging) return;
    dragging.classList.remove("dragging");
    dragging = null;
    refreshArrangeOrderText();
  });

  grid.addEventListener("mousedown", (e) => {
    if (!arrangeState.splitSelectMode) return;
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    const idx = getThumbItems().indexOf(item);
    if (idx < 0) return;
    selecting = true;
    anchorIndex = idx;
    applyRangeSelection(anchorIndex, idx);
  });

  grid.addEventListener("mouseover", (e) => {
    if (!arrangeState.splitSelectMode || !selecting) return;
    const item = e.target.closest(".thumb-item");
    if (!item) return;
    const idx = getThumbItems().indexOf(item);
    if (idx < 0) return;
    applyRangeSelection(anchorIndex, idx);
  });

  document.addEventListener("mouseup", () => {
    selecting = false;
  });
};

const renderArrangeThumbs = async (file) => {
  startOperation("arrange", "썸네일 렌더링 준비 중...");
  try {
    const grid = $("pdfThumbGrid");
    grid.innerHTML = "";
    const buffer = await readAsArrayBuffer(file);
    checkCancelled("arrange");
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
    arrangeState.renderDoc = pdf;
    arrangeState.file = file;
    arrangeState.pageOrder = [];
    arrangeState.splitBlocks = [];
    renderSplitBlocks();
    setSplitSelectionMode(false);

    for (let i = 1; i <= pdf.numPages; i += 1) {
      checkCancelled("arrange");
      setStatus("arrangePdfStatus", `썸네일 렌더링 중 (${i}/${pdf.numPages})`);
      const page = await pdf.getPage(i);
      const baseViewport = page.getViewport({ scale: 1 });
      const targetWidth = 120;
      const scale = targetWidth / baseViewport.width;
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({
        canvasContext: canvas.getContext("2d"),
        viewport,
      }).promise;

      const item = document.createElement("div");
      item.className = "thumb-item";
      item.draggable = true;
      item.dataset.page = String(i);
      item.innerHTML = `<div class="thumb-label">p.${i}</div>`;
      item.prepend(canvas);
      grid.appendChild(item);
      updateProgress("arrange", i, pdf.numPages);
    }

    refreshArrangeOrderText();
    updateProgress("arrange", 100, 100);
    endOperation("arrange", `완료: ${pdf.numPages}개 페이지를 드래그 정렬할 수 있습니다.`);
  } catch (err) {
    handleOperationError("arrange", err);
  }
};

const setupPdfArrange = () => {
  if (!$("runReorderPdf") || !$("pdfThumbGrid")) return;
  setupThumbDnD();
  renderSplitBlocks();

  $("toggleSplitSelectMode").addEventListener("click", () => {
    setSplitSelectionMode(!arrangeState.splitSelectMode);
  });

  $("addSplitBlock").addEventListener("click", () => {
    const selectedItems = getThumbItems().filter((item) =>
      item.classList.contains("selected-range")
    );
    if (!selectedItems.length) {
      setStatus("arrangePdfStatus", "분할 선택 모드에서 썸네일 구간을 먼저 드래그 선택해주세요.");
      return;
    }
    const blockPages = selectedItems.map((item) => Number(item.dataset.page));
    arrangeState.splitBlocks.push(blockPages);
    renderSplitBlocks();
    clearThumbSelection();
    setStatus("arrangePdfStatus", `블록 ${arrangeState.splitBlocks.length} 추가됨`);
  });

  $("clearSplitBlocks").addEventListener("click", () => {
    arrangeState.splitBlocks = [];
    renderSplitBlocks();
    clearThumbSelection();
    setStatus("arrangePdfStatus", "분할 블록이 초기화되었습니다.");
  });

  $("splitBlocksList").addEventListener("click", (e) => {
    const btn = e.target.closest(".remove-block");
    if (!btn) return;
    const idx = Number(btn.dataset.blockIndex);
    if (!Number.isInteger(idx)) return;
    arrangeState.splitBlocks.splice(idx, 1);
    renderSplitBlocks();
  });

  $("loadArrangePdf").addEventListener("click", async () => {
    const file = $("arrangePdfFile").files[0];
    if (!file) {
      setStatus("arrangePdfStatus", "PDF 파일을 선택해주세요.");
      return;
    }
    await renderArrangeThumbs(file);
  });

  $("runReorderPdf").addEventListener("click", async () => {
    const file = $("arrangePdfFile").files[0];
    if (!file) {
      setStatus("arrangePdfStatus", "PDF 파일을 먼저 선택해주세요.");
      return;
    }
    if (!arrangeState.pageOrder.length) {
      await renderArrangeThumbs(file);
    }

    startOperation("arrange", "정렬 순서로 PDF 생성 중...");
    try {
      const src = await PDFLib.PDFDocument.load(await readAsArrayBuffer(file));
      const out = await PDFLib.PDFDocument.create();
      const copied = await out.copyPages(
        src,
        arrangeState.pageOrder.map((n) => n - 1)
      );
      copied.forEach((p, idx) => {
        checkCancelled("arrange");
        out.addPage(p);
        updateProgress("arrange", idx + 1, copied.length);
      });
      const result = await out.save();
      downloadBlob(new Blob([result], { type: "application/pdf" }), "reordered.pdf");
      updateProgress("arrange", 100, 100);
      endOperation("arrange", "완료: 드래그 순서대로 저장했습니다.");
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

    startOperation("arrange", "PDF 분할 처리 중...");
    try {
      const src = await PDFLib.PDFDocument.load(await readAsArrayBuffer(file));
      let groups = [];

      if (arrangeState.splitBlocks.length) {
        groups = arrangeState.splitBlocks;
      } else {
        if (!splitText) throw new Error("분할 블록이 없으면 분할 입력값을 입력해주세요.");
        groups = parseSplitGroups(splitText, src.getPageCount()).filter((g) => g.length);
      }

      if (!groups.length) throw new Error("유효한 분할값이 없습니다.");

      const zip = new JSZip();
      for (let i = 0; i < groups.length; i += 1) {
        checkCancelled("arrange");
        const out = await PDFLib.PDFDocument.create();
        const pages = await out.copyPages(
          src,
          groups[i].map((n) => n - 1)
        );
        pages.forEach((p) => out.addPage(p));
        zip.file(`split-${i + 1}.pdf`, await out.save());
        updateProgress("arrange", i + 1, groups.length);
      }
      checkCancelled("arrange");
      const blob = await zip.generateAsync({ type: "blob" });
      downloadBlob(blob, "split-pdfs.zip");
      updateProgress("arrange", 100, 100);
      endOperation("arrange", `완료: ${groups.length}개 파일로 분할했습니다.`);
    } catch (err) {
      handleOperationError("arrange", err);
    }
  });
};

const setupPdfMerge = () => {
  if (!$("runMergePdf")) return;
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

const setupImageResize = () => {
  if (!$("runResize")) return;
  $("runResize").addEventListener("click", async () => {
    const files = [...$("resizeFiles").files];
    const width = Number($("resizeWidth").value);
    const height = Number($("resizeHeight").value);
    if (!files.length || (!width && !height)) {
      setStatus("resizeStatus", "이미지 파일과 너비/높이 중 하나 이상을 입력해주세요.");
      return;
    }

    startOperation("resize", "이미지 리사이즈 중...");
    try {
      const zip = new JSZip();
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
        zip.file(
          `${files[i].name.replace(/\.[^.]+$/, "")}_${targetW}x${targetH}.png`,
          canvas.toDataURL("image/png").split(",")[1],
          { base64: true }
        );
        updateProgress("resize", i + 1, files.length);
      }
      checkCancelled("resize");
      const blob = await zip.generateAsync({ type: "blob" });
      downloadBlob(blob, "resized-images.zip");
      updateProgress("resize", 100, 100);
      endOperation("resize", `완료: ${files.length}개 변환 (${formatBytes(blob.size)})`);
    } catch (err) {
      handleOperationError("resize", err);
    }
  });
};

const setupImageFormat = () => {
  if (!$("runFormatConvert")) return;
  $("runFormatConvert").addEventListener("click", async () => {
    const files = [...$("formatFiles").files];
    const fmt = $("targetFormat").value;
    const quality = Number($("targetQuality").value);
    if (!files.length) {
      setStatus("formatStatus", "이미지를 선택해주세요.");
      return;
    }

    startOperation("format", "이미지 포맷 변환 중...");
    try {
      const zip = new JSZip();
      for (let i = 0; i < files.length; i += 1) {
        checkCancelled("format");
        setStatus("formatStatus", `포맷 변환 중 (${i + 1}/${files.length})`);
        const { img } = await loadImageFromFile(files[i]);
        const canvas = document.createElement("canvas");
        canvas.width = img.width;
        canvas.height = img.height;
        canvas.getContext("2d").drawImage(img, 0, 0);
        const mime = fmt === "png" ? "image/png" : fmt === "jpeg" ? "image/jpeg" : "image/webp";
        const base64 = canvas.toDataURL(mime, quality).split(",")[1];
        const ext = fmt === "jpeg" ? "jpg" : fmt;
        zip.file(`${files[i].name.replace(/\.[^.]+$/, "")}.${ext}`, base64, { base64: true });
        updateProgress("format", i + 1, files.length);
      }
      checkCancelled("format");
      const blob = await zip.generateAsync({ type: "blob" });
      downloadBlob(blob, "converted-images.zip");
      updateProgress("format", 100, 100);
      endOperation("format", `완료: ${files.length}개 변환 (${formatBytes(blob.size)})`);
    } catch (err) {
      handleOperationError("format", err);
    }
  });
};

const processRecords = [];

const escapeCsv = (value) => {
  const str = String(value ?? "");
  if (str.includes(",") || str.includes('"') || str.includes("\n")) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
};

const processMediaState = {
  captured: [],
  stream: null,
  recorder: null,
  chunks: [],
};

const renderCapturedMediaPreview = () => {
  const box = $("capturedMediaPreview");
  box.innerHTML = "";
  processMediaState.captured.forEach((item) => {
    const card = document.createElement("div");
    card.className = "media-item";
    const mediaTag =
      item.type.startsWith("video")
        ? `<video src="${item.url}" controls></video>`
        : `<img src="${item.url}" alt="${item.name}" />`;
    card.innerHTML = `${mediaTag}<div class="media-name">${item.name}</div>`;
    box.appendChild(card);
  });
};

const addCapturedBlob = (blob, name, type) => {
  const url = URL.createObjectURL(blob);
  processMediaState.captured.push({ blob, name, type, url });
  renderCapturedMediaPreview();
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

  const addSessionCard = ({
    durationText,
    name,
    customer,
    memo,
    mediaItems,
    timestamp,
  }) => {
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
        card.innerHTML = m.type.startsWith("video")
          ? `<video src="${m.url}" controls></video><div class="media-name">${m.name}</div>`
          : `<img src="${m.url}" alt="${m.name}" /><div class="media-name">${m.name}</div>`;
        grid.appendChild(card);
      });
      wrap.appendChild(grid);
    }

    $("timerLog").prepend(wrap);
  };

  const gatherAllMedia = () => {
    const fileMedia = [...$("procMediaFiles").files, ...$("procCapture").files].map((f) => ({
      name: f.name,
      type: f.type || "application/octet-stream",
      url: URL.createObjectURL(f),
    }));
    const captured = processMediaState.captured.map((m) => ({
      name: m.name,
      type: m.type,
      url: m.url,
    }));
    return [...fileMedia, ...captured];
  };

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

  $("startCamera").addEventListener("click", async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: true, audio: true });
      processMediaState.stream = stream;
      $("cameraPreview").srcObject = stream;
      setStatus("timerStatus", "카메라가 시작되었습니다.");
    } catch (err) {
      setStatus("timerStatus", `카메라 시작 실패: ${err.message}`);
    }
  });

  $("stopCamera").addEventListener("click", () => {
    if (!processMediaState.stream) return;
    processMediaState.stream.getTracks().forEach((track) => track.stop());
    processMediaState.stream = null;
    $("cameraPreview").srcObject = null;
    setStatus("timerStatus", "카메라를 종료했습니다.");
  });

  $("capturePhoto").addEventListener("click", () => {
    const video = $("cameraPreview");
    if (!video.srcObject) {
      setStatus("timerStatus", "카메라를 먼저 시작해주세요.");
      return;
    }
    const canvas = document.createElement("canvas");
    canvas.width = video.videoWidth || 1280;
    canvas.height = video.videoHeight || 720;
    canvas.getContext("2d").drawImage(video, 0, 0, canvas.width, canvas.height);
    canvas.toBlob((blob) => {
      if (!blob) return;
      addCapturedBlob(blob, `captured-${Date.now()}.png`, "image/png");
      setStatus("timerStatus", "사진이 추가되었습니다.");
    }, "image/png");
  });

  $("startVideoRec").addEventListener("click", () => {
    if (!processMediaState.stream) {
      setStatus("timerStatus", "카메라를 먼저 시작해주세요.");
      return;
    }
    if (typeof MediaRecorder === "undefined") {
      setStatus("timerStatus", "이 브라우저는 영상 녹화를 지원하지 않습니다.");
      return;
    }
    try {
      processMediaState.chunks = [];
      processMediaState.recorder = new MediaRecorder(processMediaState.stream);
      processMediaState.recorder.ondataavailable = (e) => {
        if (e.data.size > 0) processMediaState.chunks.push(e.data);
      };
      processMediaState.recorder.onstop = () => {
        const blob = new Blob(processMediaState.chunks, { type: "video/webm" });
        addCapturedBlob(blob, `recorded-${Date.now()}.webm`, "video/webm");
        setStatus("timerStatus", "영상이 추가되었습니다.");
      };
      processMediaState.recorder.start();
      setStatus("timerStatus", "영상 기록을 시작했습니다.");
    } catch (err) {
      setStatus("timerStatus", `영상 기록 시작 실패: ${err.message}`);
    }
  });

  $("stopVideoRec").addEventListener("click", () => {
    if (!processMediaState.recorder) return;
    if (processMediaState.recorder.state !== "inactive") {
      processMediaState.recorder.stop();
    }
  });
};

const setupQr = () => {
  if (!$("runQr") || !$("saveQr") || !$("qrPreview")) return;
  let qr = null;
  $("runQr").addEventListener("click", () => {
    const text = $("qrInput").value.trim();
    const size = Number($("qrSize").value || 220);
    if (!text) {
      setStatus("qrStatus", "QR 입력값을 입력해주세요.");
      return;
    }
    $("qrPreview").innerHTML = "";
    qr = new QRCode($("qrPreview"), {
      text,
      width: size,
      height: size,
      colorDark: "#111",
      colorLight: "#fff",
      correctLevel: QRCode.CorrectLevel.M,
    });
    setStatus("qrStatus", "QR 코드 생성 완료");
  });

  $("saveQr").addEventListener("click", () => {
    if (!qr) {
      setStatus("qrStatus", "먼저 QR 코드를 생성해주세요.");
      return;
    }
    const canvas = $("qrPreview").querySelector("canvas");
    const img = $("qrPreview").querySelector("img");
    if (canvas) {
      canvas.toBlob((blob) => downloadBlob(blob, "qrcode.png"));
      return;
    }
    if (img) {
      fetch(img.src)
        .then((res) => res.blob())
        .then((blob) => downloadBlob(blob, "qrcode.png"));
    }
  });
};

const init = () => {
  initOperations();
  setupThemeToggle();
  setupNavActive();
  setupPdfToImage();
  setupImageToPdf();
  setupPdfArrange();
  setupPdfMerge();
  setupImageResize();
  setupImageFormat();
  setupProcessTimer();
  setupQr();
};

init();

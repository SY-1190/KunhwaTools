(() => {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const state = { file: null, pages: [], preview: null, page: 0, selected: null, history: [[]], step: 0, scale: 1, generation: 0, busy: false };
  const operations = () => state.history[state.step];
  const hasUnsaved = () => state.file && (operations() !== state.savedOperations || state.inline);
  const status = (text, error = false) => { $("pdfEditStatus").textContent = text; $("pdfEditStatus").classList.toggle("error", error); };
  const localError = "로컬 PDF 엔진에 연결할 수 없습니다. start-kunhwa-tools.ps1로 실행한 주소에서 다시 열어 주세요.";
  $("pdfEditOpen").onclick = () => $("pdfEditFile").click();
  ["pdfEditFile", "pdfWordFile"].forEach((id) => {
    const zone = document.querySelector(`[data-file-input="${id}"]`);
    zone.tabIndex = 0; zone.setAttribute("role", "button"); zone.setAttribute("aria-label", "PDF 파일 선택");
    zone.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); $(id).click(); } });
  });
  async function request(action, file, fields = {}) {
    const form = new FormData(); form.append("file", file);
    Object.entries(fields).forEach(([key, value]) => form.append(key, value));
    let response;
    try { response = await fetch(`/api/pdf-document/${action}`, { method: "POST", body: form }); }
    catch { throw new Error(localError); }
    if (!response.ok) {
      const body = await response.json().catch(() => null);
      throw new Error(body?.error || localError);
    }
    return response;
  }
  function download(blob, name) {
    const url = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
  function items() {
    const page = state.pages[state.page]; if (!page) return [];
    const edits = operations().filter((op) => op.page === state.page);
    const result = page.objects.map((item) => ({ ...item, sourceId: item.id, page: state.page, fill: "#ffffff" }));
    edits.forEach((op) => {
      const index = result.findIndex((item) => item.id === op.id);
      if (index >= 0) result[index] = op; else result.push(op);
    });
    return result.filter((item) => !item.deleted);
  }
  function selected() { return items().find((item) => item.id === state.selected); }
  function commit(op) {
    const list = structuredClone(operations());
    const index = list.findIndex((item) => item.id === op.id && item.page === op.page);
    if (index >= 0) list[index] = op; else list.push(op);
    state.history = state.history.slice(0, state.step + 1); state.history.push(list); state.step++; state.preview = null;
    render(); refreshPreview();
  }
  async function refreshPreview() {
    const generation = ++state.generation;
    if (!operations().length) {
      state.preview = null; render(); status("원본 상태입니다.");
      $("pdfEditSave").disabled = state.busy || !state.file;
      return;
    }
    status("편집 결과 확인 중…"); $("pdfEditSave").disabled = true;
    try {
      const result = await request("edit", state.file, { operations: JSON.stringify(operations()) });
      const blob = await result.blob();
      if (generation !== state.generation) return;
      const response = await request("analyze", new File([blob], "preview.pdf", { type: "application/pdf" }));
      const preview = await response.json();
      if (generation !== state.generation) return;
      state.preview = preview.pages; state.previewGeneration = generation; render(); status(`편집 ${operations().length}개 · 원본은 변경되지 않았습니다.`);
    } catch (error) {
      if (generation === state.generation) { state.preview = null; render(); status(error.message, true); }
    } finally { if (generation === state.generation) $("pdfEditSave").disabled = !state.file || state.busy; }
  }
  function render() {
    const page = state.pages[state.page]; if (!page) return;
    const viewport = $("pdfEditViewport"), canvas = $("pdfEditCanvas");
    state.scale = $("pdfEditZoom").value === "fit" ? Math.min(1.6, Math.max(.2, (viewport.clientWidth - 40) / page.width)) : Number($("pdfEditZoom").value);
    const inlineFocused = state.inline?.form.contains(document.activeElement);
    canvas.replaceChildren(); canvas.style.width = `${page.width * state.scale}px`; canvas.style.height = `${page.height * state.scale}px`;
    const bg = document.createElement("img"); bg.className = "studio-background"; bg.src = (state.preview?.[state.page] || page).image; bg.alt = `PDF ${state.page + 1}페이지`; bg.draggable = false; canvas.append(bg);
    const validPreview = state.preview && state.previewGeneration === state.generation;
    if (!validPreview && operations().length) {
      operations().filter((op) => op.page === state.page).forEach((op) => {
        const source = page.objects.find((item) => item.id === op.sourceId);
        if (source) { const cover = document.createElement("div"); cover.className = "studio-draft-cover"; position(cover, source.rect); cover.style.background = op.fill; canvas.append(cover); }
        if (!op.deleted) {
          const draft = document.createElement(op.type === "image" ? "img" : "div"); draft.className = "studio-draft";
          position(draft, op.rect);
          if (op.type === "image") draft.src = op.image;
          else { draft.textContent = op.text; draft.style.fontSize = `${op.size * state.scale}px`; draft.style.color = op.color; }
          canvas.append(draft);
        }
      });
    }
    items().forEach((item) => {
      const button = document.createElement("button"); button.className = `studio-object ${item.type}${item.id === state.selected ? " selected" : ""}`;
      button.setAttribute("aria-label", item.type === "text" ? `글자: ${item.text}` : "이미지 선택"); button.title = item.type === "text" ? item.text : "이미지"; button.dataset.objectId = item.id;
      position(button, item.rect); button.addEventListener("click", () => selectItem(item.id));
      if (item.type === "text") button.addEventListener("dblclick", () => beginInlineEdit(item));
      button.addEventListener("pointerdown", (event) => beginDrag(event, item)); canvas.append(button);
    });
    if (state.inline?.item.page === state.page) {
      placeInlineEditor();
      if (inlineFocused) state.inline.form.querySelector("textarea").focus({ preventScroll: true });
    }
    $("pdfEditPageLabel").textContent = `${state.page + 1} / ${state.pages.length}`;
    $("pdfEditPrev").disabled = state.page === 0; $("pdfEditNext").disabled = state.page === state.pages.length - 1;
    $("pdfEditUndo").disabled = state.step === 0; $("pdfEditRedo").disabled = state.step === state.history.length - 1;
    $("pdfEditDelete").disabled = !selected();
    $("pdfEditInline").disabled = selected()?.type !== "text";
    $("pdfEditPages").replaceChildren(...state.pages.map((p, number) => {
      const button = document.createElement("button"); button.className = `studio-thumbnail${number === state.page ? " active" : ""}`; button.setAttribute("aria-label", `${number + 1}페이지`);
      const image = document.createElement("img"); image.src = (state.preview?.[number] || p).image; image.alt = "";
      const label = document.createElement("span"); label.textContent = String(number + 1); button.append(image, label);
      button.onclick = () => { state.page = number; state.selected = null; render(); }; return button;
    }));
    inspector();
  }
  function position(element, rect) {
    Object.assign(element.style, { left: `${rect[0] * state.scale}px`, top: `${rect[1] * state.scale}px`, width: `${(rect[2] - rect[0]) * state.scale}px`, height: `${(rect[3] - rect[1]) * state.scale}px` });
  }
  function selectItem(id) {
    if (state.inline) return;
    state.selected = id;
    $("pdfEditCanvas").querySelectorAll(".studio-object").forEach((button) => {
      button.classList.toggle("selected", button.dataset.objectId === id);
    });
    $("pdfEditDelete").disabled = !selected();
    $("pdfEditInline").disabled = selected()?.type !== "text";
    inspector();
  }
  function placeInlineEditor() {
    const { item, form } = state.inline;
    const canvas = $("pdfEditCanvas"), width = canvas.clientWidth;
    const editorWidth = Math.min(width, Math.max(180, (item.rect[2] - item.rect[0]) * state.scale + 12));
    form.style.width = `${editorWidth}px`;
    form.style.left = `${Math.max(0, Math.min(width - editorWidth, item.rect[0] * state.scale - 6))}px`;
    canvas.append(form);
    form.style.top = `${Math.max(0, Math.min(canvas.clientHeight - form.offsetHeight, item.rect[1] * state.scale - 6))}px`;
  }
  function beginInlineEdit(item) {
    if (state.busy || item?.type !== "text") return;
    if (state.inline) {
      state.page = state.inline.item.page;
      render();
      state.inline.form.querySelector("textarea").focus();
      return;
    }
    selectItem(item.id);
    const form = document.createElement("form");
    form.className = "studio-inline-editor";
    form.innerHTML = '<textarea rows="2" maxlength="10000" aria-label="PDF 글자 편집"></textarea><div class="studio-inline-actions"><button type="button" class="btn ghost">취소</button><button class="btn primary" type="submit">적용</button></div>';
    const input = form.querySelector("textarea");
    input.value = item.text;
    const close = () => {
      form.remove(); state.inline = null;
      [...$("pdfEditCanvas").querySelectorAll(".studio-object")].find((node) => node.dataset.objectId === item.id)?.focus({ preventScroll: true });
    };
    form.querySelector('button[type="button"]').onclick = close;
    form.onkeydown = (event) => {
      if (event.isComposing) return;
      if (event.key === "Escape") { event.preventDefault(); close(); }
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) { event.preventDefault(); form.requestSubmit(); }
    };
    form.onsubmit = (event) => {
      event.preventDefault();
      if (state.busy) return;
      const text = input.value;
      if (text === item.text) { close(); return; }
      if (!text.trim()) { status("글자를 입력하거나 선택 항목 삭제를 사용해 주세요.", true); return; }
      const page = state.pages[item.page];
      const [x, y] = item.rect;
      const measure = document.createElement("canvas").getContext("2d");
      measure.font = `${item.size}px "Malgun Gothic", sans-serif`;
      const longest = Math.max(...text.split("\n").map((line) => measure.measureText(line).width));
      const right = Math.min(page.width, Math.max(item.rect[2], x + longest * 1.06 + 2));
      const lines = text.split("\n").reduce((n, line) => n + Math.max(1, Math.ceil(measure.measureText(line).width / Math.max(1, right - x - 2))), 0);
      const bottom = Math.min(page.height, Math.max(item.rect[3], y + item.size * (1.8 + (lines - 1) * 1.25)));
      close(); commit({ ...item, text, rect: [x, y, right, bottom] });
    };
    state.inline = { item, form };
    placeInlineEditor(); input.focus(); input.select();
  }
  function inspector() {
    const item = selected(); $("pdfEditProperties").hidden = !item;
    $("pdfEditSelection").textContent = item ? (item.type === "text" ? "글자" : "이미지") : "항목을 선택하세요.";
    if (!item) return;
    if (state.inspected === state.selected && $("pdfEditProperties").contains(document.activeElement)) return;
    state.inspected = state.selected;
    $("pdfEditText").parentElement.hidden = item.type !== "text";
    $("pdfEditSize").parentElement.hidden = $("pdfEditColor").parentElement.hidden = item.type !== "text";
    $("pdfEditReplaceImage").hidden = item.type !== "image";
    $("pdfEditText").value = item.text || ""; $("pdfEditSize").value = item.size || 12; $("pdfEditColor").value = item.color || "#17202a"; $("pdfEditFill").value = item.fill || "#ffffff";
    ["X", "Y", "W", "H"].forEach((key, i) => {
      let value = i < 2 ? item.rect[i] : item.rect[i] - item.rect[i - 2];
      if (i === 3 && item.type === "text" && !operations().some((op) => op.id === item.id)) value = Math.min(state.pages[state.page].height - item.rect[1], Math.max(value, item.size * 1.8));
      $("pdfEdit" + key).value = Math.round(value * 10) / 10;
    });
  }
  function beginDrag(event, item) {
    if (event.button !== 0 || state.busy || state.inline) return;
    const target = event.currentTarget, startX = event.clientX, startY = event.clientY;
    target.setPointerCapture(event.pointerId); let rect = [...item.rect], moved = false;
    const move = (e) => {
      const dx = (e.clientX - startX) / state.scale, dy = (e.clientY - startY) / state.scale;
      if (Math.abs(dx) + Math.abs(dy) < 3) return;
      moved = true; const page = state.pages[state.page], width = rect[2] - rect[0], height = rect[3] - rect[1];
      const x = Math.max(0, Math.min(page.width - width, item.rect[0] + dx)), y = Math.max(0, Math.min(page.height - height, item.rect[1] + dy));
      rect = [x, y, x + width, y + height]; position(target, rect);
    };
    const finish = (e) => {
      target.removeEventListener("pointermove", move); target.removeEventListener("pointerup", finish); target.removeEventListener("pointercancel", finish);
      if (e.type === "pointercancel") { position(target, item.rect); return; }
      if (moved) { state.selected = item.id; commit({ ...item, rect }); }
    };
    target.addEventListener("pointermove", move); target.addEventListener("pointerup", finish); target.addEventListener("pointercancel", finish);
  }
  $("pdfEditFile").addEventListener("change", async (event) => {
    const file = event.target.files[0]; if (!file) return;
    if (state.busy) { event.target.value = ""; return; }
    if (hasUnsaved() && !confirm("저장하지 않은 편집이 있습니다. 다른 파일을 열까요?")) { event.target.value = ""; return; }
    const generation = ++state.generation; state.busy = true; $("pdfEditSave").disabled = true; status("PDF 분석 중…");
    try {
      if (file.size > 55 * 1024 * 1024) throw new Error("PDF는 55MB 이하만 열 수 있습니다.");
      const response = await request("analyze", file); const data = await response.json(); if (generation !== state.generation) return;
      Object.assign(state, { file, pages: data.pages, preview: null, page: 0, selected: null, inline: null, history: [[]], step: 0 });
      state.savedOperations = operations();
      $("pdfEditFileName").textContent = file.name; $("pdfEditDrop").hidden = true; $("pdfStudio").hidden = false;
      render(); status(`${data.pageCount}페이지 · ${data.pages.reduce((n, p) => n + p.objects.length, 0)}개 항목`);
    } catch (error) { if (generation === state.generation) status(error.message, true); }
    finally { if (generation === state.generation) { state.busy = false; $("pdfEditSave").disabled = !state.file; } }
  });
  $("pdfEditApply").onclick = () => {
    if (state.busy || state.inline) return;
    const item = selected(); if (!item) return;
    const [x, y, w, h] = ["X", "Y", "W", "H"].map((key) => Number($("pdfEdit" + key).value));
    const rect = [x, y, x + w, y + h], page = state.pages[state.page];
    if (![x, y, w, h].every(Number.isFinite) || x < 0 || y < 0 || w < 1 || h < 1 || rect[2] > page.width || rect[3] > page.height) { status("항목이 페이지 밖에 있습니다. 위치와 크기를 확인해 주세요.", true); return; }
    commit({ ...item, rect, text: $("pdfEditText").value, size: Number($("pdfEditSize").value), color: $("pdfEditColor").value, fill: $("pdfEditFill").value });
  };
  $("pdfEditDelete").onclick = () => { const item = selected(); if (item && !state.busy && !state.inline) { state.selected = null; commit({ ...item, deleted: true }); } };
  $("pdfEditUndo").onclick = () => { if (!state.busy && !state.inline && state.step) { state.step--; state.selected = null; state.preview = null; render(); refreshPreview(); } };
  $("pdfEditRedo").onclick = () => { if (!state.busy && !state.inline && state.step < state.history.length - 1) { state.step++; state.preview = null; render(); refreshPreview(); } };
  $("pdfEditPrev").onclick = () => { if (state.page) { state.page--; state.selected = null; render(); } };
  $("pdfEditNext").onclick = () => { if (state.page + 1 < state.pages.length) { state.page++; state.selected = null; render(); } };
  $("pdfEditZoom").onchange = render;
  $("pdfEditSelect").onclick = () => { state.selected = null; render(); };
  $("pdfEditInline").onclick = () => beginInlineEdit(selected());
  document.addEventListener("keydown", (event) => {
    if (location.hash !== "#pdfEditor" || /INPUT|TEXTAREA|SELECT/.test(event.target.tagName)) return;
    if (event.key === "Delete" && selected()) { event.preventDefault(); $("pdfEditDelete").click(); }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") { event.preventDefault(); $(event.shiftKey ? "pdfEditRedo" : "pdfEditUndo").click(); }
  });
  $("pdfEditAddText").onclick = () => {
    const page = state.pages[state.page]; if (!page || state.busy || state.inline) return;
    const item = { id: crypto.randomUUID(), page: state.page, type: "text", rect: [30, 30, Math.min(page.width - 10, 260), 70], text: "새 글자", size: 16, color: "#17202a", fill: "#ffffff" };
    state.selected = item.id; commit(item); $("pdfEditText").focus(); $("pdfEditText").select();
  };
  let replaceImage = false;
  $("pdfEditAddImage").onclick = () => { if (!state.busy && !state.inline) { replaceImage = false; $("pdfEditImage").click(); } };
  $("pdfEditReplaceImage").onclick = () => { if (!state.busy && !state.inline) { replaceImage = true; $("pdfEditImage").click(); } };
  $("pdfEditImage").onchange = async (event) => {
    const file = event.target.files[0], target = selected(); if (!file) return;
    try {
      if (file.size > 12 * 1024 * 1024) throw new Error("이미지는 12MB 이하만 사용할 수 있습니다.");
      const image = await new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = reject; reader.readAsDataURL(file); });
      const bitmap = await createImageBitmap(file), page = state.pages[state.page];
      const width = Math.min(200, page.width - 60), height = Math.min(page.height - 60, width * bitmap.height / bitmap.width); bitmap.close();
      const item = replaceImage && target?.type === "image" ? { ...target, image } : { id: crypto.randomUUID(), page: state.page, type: "image", rect: [30, 30, 30 + width, 30 + height], image, fill: "#ffffff" };
      state.selected = item.id; commit(item);
    } catch (error) { status(error.message || "이미지를 읽을 수 없습니다.", true); }
    event.target.value = "";
  };
  $("pdfEditSave").onclick = async () => {
    if (!state.file || state.busy) return;
    if (state.inline) {
      state.page = state.inline.item.page; render(); state.inline.form.querySelector("textarea").focus();
      status("글자 편집을 적용하거나 취소한 뒤 저장해 주세요.", true); return;
    }
    state.busy = true; $("pdfEditSave").disabled = true; status("PDF 저장 중…");
    const saved = operations();
    try { const response = await request("edit", state.file, { operations: JSON.stringify(saved) }); download(await response.blob(), state.file.name.replace(/\.pdf$/i, "_편집.pdf")); state.savedOperations = saved; status("편집 PDF 다운로드 완료 · 원본 파일은 그대로 유지됩니다."); }
    catch (error) { status(error.message, true); }
    finally { state.busy = false; $("pdfEditSave").disabled = false; }
  };
  let resizeTimer; new ResizeObserver(() => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { if (state.pages.length && $("pdfEditor").classList.contains("active-stage")) render(); }, 100); }).observe($("pdfEditViewport"));
  window.addEventListener("hashchange", () => { if (location.hash === "#pdfEditor" && state.pages.length) requestAnimationFrame(render); });
  window.addEventListener("tools:routechange", () => { if (location.hash === "#pdfEditor" && state.pages.length) requestAnimationFrame(render); });
  window.addEventListener("beforeunload", (event) => { if (hasUnsaved()) { event.preventDefault(); event.returnValue = ""; } });

  let wordFile = null, wordBusy = false, wordGeneration = 0;
  const wordStatus = (text, error = false) => { $("pdfWordStatus").textContent = text; $("pdfWordStatus").classList.toggle("error", error); };
  $("pdfWordFile").onchange = async (event) => {
    const file = event.target.files[0]; if (!file) return; const generation = ++wordGeneration; wordFile = null; $("pdfWordConvert").disabled = true; wordStatus("PDF 분석 중…");
    try {
      const response = await request("analyze", file), data = await response.json(); if (generation !== wordGeneration) return;
      wordFile = file; $("pdfWordPreview").replaceChildren();
      const image = document.createElement("img"); image.src = data.pages[0].image; image.alt = `${file.name} 첫 페이지`;
      const label = document.createElement("p"); label.textContent = `${file.name} · ${data.pageCount}페이지`;
      $("pdfWordPreview").append(image, label); wordStatus("변환 준비 완료");
    } catch (error) { if (generation === wordGeneration) wordStatus(error.message, true); }
    finally { if (generation === wordGeneration) $("pdfWordConvert").disabled = !wordFile || wordBusy; }
  };
  $("pdfWordConvert").onclick = async () => {
    if (!wordFile || wordBusy) return; wordBusy = true; $("pdfWordConvert").disabled = true; wordStatus("Word 문서 생성 중…");
    try {
      const file = wordFile, response = await request("word", file, { mode: $("pdfWordMode").value });
      const encoded = response.headers.get("X-Kunhwa-Conversion-Report");
      const report = encoded ? JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(encoded.replace(/-/g, "+").replace(/_/g, "/")), (c) => c.charCodeAt(0)))) : null;
      download(await response.blob(), file.name.replace(/\.pdf$/i, ".docx"));
      wordStatus(report?.scannedPages.length ? `다운로드 완료 · ${report.scannedPages.join(", ")}페이지는 이미지 PDF로, 이미지로 유지됩니다. OCR은 적용되지 않습니다.` : `다운로드 완료 · ${report?.pages || ""}페이지${report?.mode === "appearance" ? " · 원본 모양 유지" : ` · 본문 ${report?.paragraphs}개 · 표 ${report?.tables}개`}`);
    } catch (error) { wordStatus(error.message, true); }
    finally { wordBusy = false; $("pdfWordConvert").disabled = !wordFile; }
  };
})();

(() => {
  const groups = [
    { label: "문서", tools: [["pdfEditor", "PDF 편집", "pencil-square"], ["pdfToWord", "PDF → Word", "file-earmark-word"], ["pdfToExcel", "PDF → Excel", "file-earmark-spreadsheet"], ["pdfToImage", "PDF → 이미지", "file-earmark-pdf"], ["imageToPdf", "이미지 → PDF", "images"], ["pdfArrange", "PDF 순서 · 나누기 · 회전", "layers"], ["pdfMerge", "PDF 합치기", "intersect"], ["pdfCompress", "PDF 압축", "file-zip"], ["hangulWebEditor", "한글(hwp/hwpx) 포맷 에디터", "file-earmark-text"]] },
    { label: "이미지", tools: [["imageResize", "이미지 크기 조정", "aspect-ratio"], ["imageFormat", "이미지 포맷 변환", "image"]] },
    { label: "업무", tools: [["fileRename", "파일명 일괄 변경", "input-cursor-text"], ["processTimer", "공정 작업시간 측정", "stopwatch"], ["qrGenerator", "QR 코드 생성기", "qr-code"]] },
  ];
  const icon = (name) => `<svg viewBox="0 0 16 16" aria-hidden="true"><use href="./assets/icons.svg#${name}"></use></svg>`;
  window.toolsIcon = icon;
  const nav = document.createElement("aside");
  nav.className = "workspace-nav";
  nav.id = "workspaceNav";
  nav.innerHTML = `<a class="workspace-home" href="#">${icon("grid-1x2")}<span>모든 도구</span></a><nav aria-label="도구 메뉴">${groups.map(group => `<section class="nav-group"><h2>${group.label}</h2>${group.tools.map(([id, title, name]) => `<a href="#${id}" data-nav-tool="${id}">${icon(name)}<span>${title}</span></a>`).join("")}</section>`).join("")}</nav><div class="workspace-nav-foot"><span class="privacy-dot"></span>로컬 작업 공간</div>`;
  document.querySelector(".single-shell").before(nav);
  const menuButton = document.createElement("button");
  menuButton.className = "icon-btn mobile-menu";
  menuButton.type = "button";
  menuButton.setAttribute("aria-label", "도구 메뉴 열기");
  menuButton.setAttribute("aria-expanded", "false");
  menuButton.setAttribute("aria-controls", nav.id);
  menuButton.innerHTML = icon("list");
  document.querySelector(".brand-wrap").prepend(menuButton);
  const closeNav = () => { document.body.classList.remove("nav-open"); menuButton.setAttribute("aria-expanded", "false"); };
  menuButton.addEventListener("click", () => {
    const open = document.body.classList.toggle("nav-open");
    menuButton.setAttribute("aria-expanded", String(open));
  });
  nav.addEventListener("click", (event) => { if (event.target.closest("a")) closeNav(); });
  document.addEventListener("keydown", (event) => { if (event.key === "Escape") closeNav(); });
  document.addEventListener("click", (event) => { if (!nav.contains(event.target) && !menuButton.contains(event.target)) closeNav(); });
  const descriptions = {
    pdfEditor: "PDF", pdfToWord: "PDF / DOCX", pdfToExcel: "PDF / XLSX",
    pdfToImage: "PDF / PNG · JPG · WebP", imageToPdf: "PNG · JPG / PDF", pdfArrange: "PDF", pdfMerge: "PDF", pdfCompress: "PDF", hangulWebEditor: "HWP · HWPX",
    imageResize: "PNG · JPG · WebP", imageFormat: "PNG · JPG · WebP · GIF · HEIC", fileRename: "ZIP", processTimer: "JSON · CSV", qrGenerator: "PNG · JPG · WebP · SVG",
  };
  const allTools = groups.flatMap(group => group.tools);
  document.querySelectorAll(".quick-btn").forEach(link => {
    const id = link.hash.slice(1);
    const tool = allTools.find(item => item[0] === id);
    if (!tool) return;
    link.innerHTML = `<span class="tool-glyph">${icon(tool[2])}</span><strong>${tool[1]}</strong><small>${descriptions[id]}</small><span class="tool-arrow">${icon("arrow-up-right")}</span>`;
    link.dataset.category = groups.find(group => group.tools.includes(tool)).label;
  });
  document.querySelectorAll(".tool-stage:not(.orbit-stage) > h2").forEach(heading => {
    const id = heading.parentElement.id;
    const tool = allTools.find(item => item[0] === id);
    if (tool) heading.innerHTML = `<span class="icon">${icon(tool[2])}</span>${tool[1]}`;
  });
  document.querySelectorAll("[data-ui-icon]").forEach(element => { element.innerHTML = icon(element.dataset.uiIcon); });
  const categories = document.createElement("div");
  categories.className = "hub-filters";
  categories.setAttribute("role", "group");
  categories.setAttribute("aria-label", "도구 분류");
  categories.innerHTML = ["전체", ...groups.map(group => group.label)].map((name, index) => `<button type="button" data-filter="${name}" aria-pressed="${index === 0}" class="${index === 0 ? "active" : ""}">${name}</button>`).join("");
  document.querySelector(".quick-grid").before(categories);
  const empty = document.createElement("p");
  empty.className = "search-empty";
  empty.textContent = "검색 결과가 없습니다.";
  empty.hidden = true;
  document.querySelector(".quick-grid").after(empty);
  let filter = "전체";
  const filterTools = () => {
    const query = document.getElementById("toolSearch").value.trim().toLocaleLowerCase();
    let count = 0;
    document.querySelectorAll(".quick-btn").forEach(link => {
      link.hidden = (filter !== "전체" && link.dataset.category !== filter) || !link.textContent.toLocaleLowerCase().includes(query);
      if (!link.hidden) count += 1;
    });
    empty.hidden = count !== 0;
  };
  categories.addEventListener("click", (event) => {
    const button = event.target.closest("[data-filter]");
    if (!button) return;
    filter = button.dataset.filter;
    categories.querySelectorAll("button").forEach(item => { item.classList.toggle("active", item === button); item.setAttribute("aria-pressed", String(item === button)); });
    filterTools();
  });
  document.getElementById("toolSearch").addEventListener("input", filterTools);
  const syncNav = () => nav.querySelectorAll("a").forEach(link => {
    const selected = link.hash === location.hash || (link.getAttribute("href") === "#" && (!location.hash || location.hash === "#"));
    link.classList.toggle("active", selected);
    if (selected) link.setAttribute("aria-current", "page"); else link.removeAttribute("aria-current");
  });
  window.addEventListener("hashchange", syncNav);
  window.addEventListener("tools:routechange", syncNav);
  const activeHeading = document.querySelector(".active-stage > h2");
  if (activeHeading) document.getElementById("stageTitleText").textContent = activeHeading.textContent;
  document.querySelectorAll(".quick-btn,.stage-chip").forEach(link => link.addEventListener("click", () => { requestAnimationFrame(syncNav); }));
  syncNav();
})();

(function (global) {
  "use strict";

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));
  const dom = {};
  const state = {
    project: null,
    currentSlideIndex: 0,
    selectedNodeId: null,
    history: [],
    historyIndex: -1,
    frameReady: false,
    frameRenderToken: 0,
    drag: null,
    toastTimer: null,
    aiStep: 1,
  };
  const MOVE_THRESHOLD = 6;

  function deepClone(value) {
    return global.PPTDLProjectIO.deepClone(value);
  }

  function currentSlide() {
    return state.project?.slides?.[state.currentSlideIndex] || null;
  }

  function frameDocument(slide, css) {
    return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>
      html,body{margin:0;padding:0;width:1920px;height:1080px;overflow:hidden;background:#fff}
      body{position:relative}
      ${css || ""}
      .pptdlab-frame-root{position:relative}
      .pptdlab-frame-root>[data-pptdlab-slide-root="true"]{display:block!important;visibility:visible!important;opacity:1!important}
      [data-editor-selection]{position:absolute!important;z-index:2147483646!important;box-sizing:border-box!important;border:2px solid #e83e65!important;border-radius:3px!important;background:rgba(232,62,101,.04)!important;pointer-events:none!important;}
      [data-editor-selection] .editor-resize-handle{position:absolute!important;right:-7px!important;bottom:-7px!important;width:12px!important;height:12px!important;border:2px solid #fff!important;border-radius:50%!important;background:#e83e65!important;box-shadow:0 1px 4px rgba(0,0,0,.25)!important;pointer-events:auto!important;cursor:nwse-resize!important}
      [contenteditable="true"]{outline:2px dashed #e83e65!important;outline-offset:4px!important}
    </style></head><body><div class="pptdlab-frame-root">${slide?.html || ""}</div></body></html>`;
  }

  function initDom() {
    [
      "slideFrame", "canvasFrame", "canvasViewport", "canvasArea", "slideList", "slideCountLabel", "projectTitleLabel", "dirtyIndicator",
      "statusMessage", "selectionSummary", "inspectorTitle", "clearSelectionButton", "pageInspector", "elementInspector", "pageLabelInput",
      "pageBackgroundInput", "pageBackgroundValue", "duplicatePageButton", "deletePageButton", "moveSlideUpButton", "moveSlideDownButton", "undoButton", "redoButton", "zoomLabel",
      "selectedElementLabel", "elementTextInput", "fontFamilyInput", "fontSizeInput", "textColorInput", "textColorValue", "fontWeightInput",
      "textAlignInput", "lineHeightInput", "elementXInput", "elementYInput", "elementWInput", "elementHInput", "alignLeftButton",
      "alignCenterButton", "alignTopButton", "replaceImageInput", "htmlFileInput", "projectFileInput", "htmlInput", "importStatus",
      "aiTopicInput", "aiAudienceInput", "aiDurationInput", "aiPagesInput", "aiLanguageInput", "aiUseInput", "aiStyleInput", "aiMaterialsInput",
      "outlinePromptOutput", "outlineInput", "outlineCopyStatus", "htmlPromptOutput", "htmlCopyStatus", "aiHtmlInput", "toast", "desktopHint",
    ].forEach((id) => { dom[id] = document.getElementById(id); });
    dom.importModal = document.getElementById("importModal");
    dom.aiModal = document.getElementById("aiModal");
    dom.privacyModal = document.getElementById("privacyModal");
  }

  function setStatus(message, tone) {
    dom.statusMessage.textContent = message;
    dom.statusMessage.className = tone ? `is-${tone}` : "";
  }

  function showToast(message, tone = "") {
    dom.toast.textContent = message;
    dom.toast.className = `toast${tone ? ` is-${tone}` : ""}`;
    dom.toast.hidden = false;
    if (state.toastTimer) clearTimeout(state.toastTimer);
    state.toastTimer = setTimeout(() => { dom.toast.hidden = true; }, 4200);
  }

  function setModal(modal, open) {
    modal.hidden = !open;
    if (open) {
      const focusable = modal.querySelector("input,textarea,button");
      if (focusable) setTimeout(() => focusable.focus(), 0);
    }
  }

  function setDirty(isDirty) {
    dom.dirtyIndicator.hidden = !isDirty;
  }

  function snapshotForHistory() {
    syncFrameToProject();
    return deepClone(state.project);
  }

  function resetHistory() {
    state.history = [snapshotForHistory()];
    state.historyIndex = 0;
    setDirty(false);
    updateHistoryButtons();
  }

  function recordHistory(message = "已更新") {
    const next = snapshotForHistory();
    const previous = state.history[state.historyIndex];
    if (previous && JSON.stringify(previous) === JSON.stringify(next)) return;
    state.history = state.history.slice(0, state.historyIndex + 1);
    state.history.push(next);
    if (state.history.length > global.PPTDLProjectIO.MAX_HISTORY) state.history.shift();
    state.historyIndex = state.history.length - 1;
    setDirty(true);
    updateHistoryButtons();
    setStatus(message);
  }

  function updateHistoryButtons() {
    dom.undoButton.disabled = state.historyIndex <= 0;
    dom.redoButton.disabled = state.historyIndex < 0 || state.historyIndex >= state.history.length - 1;
  }

  function undo() {
    if (state.historyIndex <= 0) return;
    syncFrameToProject();
    state.historyIndex -= 1;
    state.project = deepClone(state.history[state.historyIndex]);
    state.currentSlideIndex = Math.min(state.currentSlideIndex, state.project.slides.length - 1);
    state.selectedNodeId = null;
    setDirty(true);
    renderAll();
    setStatus("已撤销上一项修改。");
  }

  function redo() {
    if (state.historyIndex >= state.history.length - 1) return;
    syncFrameToProject();
    state.historyIndex += 1;
    state.project = deepClone(state.history[state.historyIndex]);
    state.currentSlideIndex = Math.min(state.currentSlideIndex, state.project.slides.length - 1);
    state.selectedNodeId = null;
    setDirty(true);
    renderAll();
    setStatus("已恢复上一项修改。");
  }

  function getFrameRoot() {
    const doc = dom.slideFrame.contentDocument;
    return doc?.querySelector(".pptdlab-frame-root > .slide, .pptdlab-frame-root > [data-slide]") || null;
  }

  function getSelectedNode() {
    const root = getFrameRoot();
    if (!root || !state.selectedNodeId) return null;
    return root.querySelector(`[data-editor-id="${CSS.escape(state.selectedNodeId)}"]`);
  }

  function removeOverlay() {
    const doc = dom.slideFrame.contentDocument;
    doc?.getElementById("__editor_selection")?.remove();
  }

  function syncFrameToProject() {
    if (!state.project || !state.frameReady || !currentSlide()) return;
    const root = getFrameRoot();
    if (!root) return;
    const clone = root.cloneNode(true);
    global.PPTDLSanitizer.removeEditorArtifacts(clone);
    currentSlide().html = clone.outerHTML;
    state.project.metadata.updatedAt = new Date().toISOString();
  }

  function ensureOverlay() {
    const doc = dom.slideFrame.contentDocument;
    if (!doc) return null;
    let overlay = doc.getElementById("__editor_selection");
    if (!overlay) {
      overlay = doc.createElement("div");
      overlay.id = "__editor_selection";
      overlay.setAttribute("data-editor-selection", "true");
      const handle = doc.createElement("span");
      handle.className = "editor-resize-handle";
      handle.setAttribute("data-editor-resize", "true");
      overlay.appendChild(handle);
      doc.body.appendChild(overlay);
      handle.addEventListener("pointerdown", onResizeStart);
    }
    return overlay;
  }

  function updateOverlay() {
    const node = getSelectedNode();
    const root = getFrameRoot();
    if (!node || !root) {
      removeOverlay();
      return;
    }
    const overlay = ensureOverlay();
    const rootRect = root.getBoundingClientRect();
    const rect = node.getBoundingClientRect();
    overlay.style.left = `${Math.round(rect.left - rootRect.left)}px`;
    overlay.style.top = `${Math.round(rect.top - rootRect.top)}px`;
    overlay.style.width = `${Math.max(2, Math.round(rect.width))}px`;
    overlay.style.height = `${Math.max(2, Math.round(rect.height))}px`;
  }

  function setSelectedNode(node) {
    const root = getFrameRoot();
    if (!node || !root || node === root || node === dom.slideFrame.contentDocument.body) {
      state.selectedNodeId = null;
      removeOverlay();
      renderInspector();
      dom.selectionSummary.textContent = "未选择元素";
      return;
    }
    if (!node.getAttribute("data-editor-id")) {
      global.PPTDLSanitizer.assignEditorIds(root);
    }
    state.selectedNodeId = node.getAttribute("data-editor-id");
    renderInspector();
    updateOverlay();
    dom.selectionSummary.textContent = `${node.tagName.toLowerCase()} · 已选择`;
  }

  function renderFrame() {
    const slide = currentSlide();
    if (!slide) return;
    state.frameReady = false;
    state.selectedNodeId = null;
    removeOverlay();
    const token = ++state.frameRenderToken;
    dom.slideFrame.onload = () => {
      if (token !== state.frameRenderToken) return;
      state.frameReady = true;
      bindFrameEvents();
      renderInspector();
      updateOverlay();
    };
    dom.slideFrame.srcdoc = frameDocument(slide, state.project.css);
    dom.slideFrame.setAttribute("aria-label", `正在编辑第 ${state.currentSlideIndex + 1} 页`);
  }

  function renderSlideList() {
    dom.slideList.innerHTML = "";
    state.project.slides.forEach((slide, index) => {
      const item = document.createElement("li");
      item.className = `slide-list-item${index === state.currentSlideIndex ? " is-current" : ""}${state.project.warnings?.length ? "" : ""}`;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "slide-thumb-button";
      button.setAttribute("aria-label", `切换到第 ${index + 1} 页：${slide.label}`);
      const thumb = document.createElement("div");
      thumb.className = "slide-thumb";
      // Keep imported CSS inside a shadow root so selectors such as `body`,
      // `button` or `*` cannot restyle or obscure the editor chrome.
      const shadow = thumb.attachShadow({ mode: "closed" });
      const style = document.createElement("style");
      style.textContent = `${state.project.css || ""}\n:host{display:block;width:100%;height:100%;pointer-events:none}.thumbnail-stage{width:1920px;height:1080px;transform:scale(.101);transform-origin:top left;overflow:hidden}.thumbnail-stage>[data-pptdlab-slide-root="true"]{display:block!important;visibility:visible!important;opacity:1!important}`;
      shadow.appendChild(style);
      const stage = document.createElement("div");
      stage.className = "thumbnail-stage";
      const holder = document.createElement("div");
      holder.innerHTML = slide.html;
      const root = holder.firstElementChild;
      if (root) stage.appendChild(root);
      shadow.appendChild(stage);
      button.appendChild(thumb);
      const label = document.createElement("span");
      label.className = "slide-thumb-label";
      label.textContent = slide.label || `第 ${index + 1} 页`;
      button.appendChild(label);
      button.addEventListener("click", () => switchSlide(index));
      item.appendChild(button);
      if ((slide.html || "").includes("data-local-placeholder") || (state.project.warnings || []).length) {
        const mark = document.createElement("span");
        mark.className = "slide-warning-mark";
        mark.title = "此项目导入时存在处理提示";
        mark.textContent = "!";
        item.classList.add("has-warning");
        item.appendChild(mark);
      }
      dom.slideList.appendChild(item);
    });
    dom.slideCountLabel.textContent = `${state.project.slides.length} 页`;
  }

  function renderPageInspector() {
    const slide = currentSlide();
    if (!slide) return;
    const root = getFrameRoot();
    const label = slide.label || String(state.currentSlideIndex + 1).padStart(2, "0");
    dom.pageLabelInput.value = label;
    const bg = root ? colorToHex(root.ownerDocument.defaultView.getComputedStyle(root).backgroundColor) : "#FFFFFF";
    dom.pageBackgroundInput.value = /^#[0-9a-f]{6}$/i.test(bg) ? bg : "#FFFFFF";
    dom.pageBackgroundValue.textContent = dom.pageBackgroundInput.value.toUpperCase();
    dom.deletePageButton.disabled = state.project.slides.length <= 1;
    dom.moveSlideUpButton.disabled = state.currentSlideIndex <= 0;
    dom.moveSlideDownButton.disabled = state.currentSlideIndex >= state.project.slides.length - 1;
  }

  function renderInspector() {
    const node = getSelectedNode();
    const hasNode = Boolean(node);
    dom.pageInspector.hidden = hasNode;
    dom.elementInspector.hidden = !hasNode;
    dom.clearSelectionButton.hidden = !hasNode;
    dom.inspectorTitle.textContent = hasNode ? "元素属性" : "页面属性";
    if (!hasNode) {
      renderPageInspector();
      return;
    }
    const computed = node.ownerDocument.defaultView.getComputedStyle(node);
    const rect = node.getBoundingClientRect();
    const root = getFrameRoot();
    const rootRect = root?.getBoundingClientRect();
    const isImage = node.tagName.toLowerCase() === "img";
    const text = isImage ? "" : node.textContent || "";
    dom.selectedElementLabel.textContent = `${node.tagName.toLowerCase()}${node.className && typeof node.className === "string" ? `.${node.className.split(/\s+/)[0]}` : ""}`;
    dom.elementTextInput.value = text.trim();
    dom.elementTextInput.disabled = isImage;
    dom.fontFamilyInput.value = findFontOption(computed.fontFamily);
    dom.fontSizeInput.value = Math.round(parseFloat(computed.fontSize) || 24);
    const color = colorToHex(computed.color);
    dom.textColorInput.value = /^#[0-9a-f]{6}$/i.test(color) ? color : "#172033";
    dom.textColorValue.textContent = dom.textColorInput.value.toUpperCase();
    dom.fontWeightInput.value = ["400", "500", "600", "700", "800"].includes(computed.fontWeight) ? computed.fontWeight : "400";
    dom.textAlignInput.value = ["left", "center", "right", "justify"].includes(computed.textAlign) ? computed.textAlign : "left";
    dom.lineHeightInput.value = lineHeightNumber(computed.lineHeight, computed.fontSize);
    dom.elementXInput.value = Math.round((rect.left - (rootRect?.left || 0)) || 0);
    dom.elementYInput.value = Math.round((rect.top - (rootRect?.top || 0)) || 0);
    dom.elementWInput.value = Math.round(rect.width || 0);
    dom.elementHInput.value = Math.round(rect.height || 0);
  }

  function findFontOption(fontFamily) {
    const normalized = String(fontFamily || "").replace(/["']/g, "").split(",")[0].trim();
    const options = Array.from(dom.fontFamilyInput.options).map((option) => option.value);
    return options.includes(normalized) ? normalized : "Aptos";
  }

  function lineHeightNumber(value, fontSize) {
    if (value === "normal") return "1.2";
    const numeric = parseFloat(value);
    const size = parseFloat(fontSize) || 24;
    return Number.isFinite(numeric) ? (numeric > 4 ? (numeric / size).toFixed(2) : numeric.toFixed(2)) : "1.2";
  }

  function colorToHex(value) {
    const input = String(value || "").trim();
    if (/^#[0-9a-f]{6}$/i.test(input)) return input;
    const match = input.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
    if (!match) return "#172033";
    return `#${[match[1], match[2], match[3]].map((part) => Number(part).toString(16).padStart(2, "0")).join("")}`;
  }

  function applyElementStyle(property, value, message, commit = false) {
    const node = getSelectedNode();
    if (!node) return;
    node.style[property] = value;
    updateOverlay();
    renderInspector();
    if (commit) recordHistory(message);
  }

  function setElementText(value, commit = false) {
    const node = getSelectedNode();
    if (!node || node.tagName.toLowerCase() === "img") return;
    node.textContent = value;
    updateOverlay();
    if (commit) recordHistory("已更新文字内容");
  }

  function elementPosition(node) {
    const root = getFrameRoot();
    if (!node || !root) return { x: 0, y: 0, width: 0, height: 0 };
    const rootRect = root.getBoundingClientRect();
    const rect = node.getBoundingClientRect();
    return { x: rect.left - rootRect.left, y: rect.top - rootRect.top, width: rect.width, height: rect.height };
  }

  function offsetParentPosition(node) {
    const root = getFrameRoot();
    const parent = node?.offsetParent || root;
    if (!node || !parent) return { parent: null, left: 0, top: 0 };
    const nodeRect = node.getBoundingClientRect();
    const parentRect = parent.getBoundingClientRect();
    const borderLeft = Number(parent.clientLeft) || 0;
    const borderTop = Number(parent.clientTop) || 0;
    return {
      parent,
      left: nodeRect.left - parentRect.left - borderLeft + (parent.scrollLeft || 0),
      top: nodeRect.top - parentRect.top - borderTop + (parent.scrollTop || 0),
    };
  }

  function setElementGeometry(values, message, commit = false) {
    const node = getSelectedNode();
    if (!node) return;
    const root = getFrameRoot();
    if (!root) return;
    const current = elementPosition(node);
    node.style.position = "absolute";
    if (values.x != null) node.style.left = `${Math.round(Math.max(0, Math.min(1920 - 24, Number(values.x) || 0)))}px`;
    if (values.y != null) node.style.top = `${Math.round(Math.max(0, Math.min(1080 - 24, Number(values.y) || 0)))}px`;
    if (values.width != null) node.style.width = `${Math.max(24, Math.round(Number(values.width) || current.width))}px`;
    if (values.height != null) node.style.height = `${Math.max(24, Math.round(Number(values.height) || current.height))}px`;
    updateOverlay();
    renderInspector();
    if (commit) recordHistory(message);
  }

  function beginMove(node, event) {
    if (!node || node === getFrameRoot() || node.tagName.toLowerCase() === "img" && event.target === node && event.detail === 2) return;
    const pos = offsetParentPosition(node);
    state.drag = {
      type: "move",
      node,
      offsetParent: pos.parent,
      startX: event.clientX,
      startY: event.clientY,
      startLeft: pos.left,
      startTop: pos.top,
      moved: false,
    };
  }

  function onResizeStart(event) {
    event.preventDefault();
    event.stopPropagation();
    const node = getSelectedNode();
    if (!node) return;
    const pos = elementPosition(node);
    state.drag = { type: "resize", node, startX: event.clientX, startY: event.clientY, startWidth: pos.width, startHeight: pos.height, moved: false };
  }

  function onFramePointerMove(event) {
    const drag = state.drag;
    if (!drag) return;
    const dx = event.clientX - drag.startX;
    const dy = event.clientY - drag.startY;
    if (!drag.moved) {
      if (Math.hypot(dx, dy) < MOVE_THRESHOLD) return;
      drag.moved = true;
    }
    if (drag.type === "move") {
      drag.node.style.position = "absolute";
      drag.node.style.left = `${Math.round(drag.startLeft + dx)}px`;
      drag.node.style.top = `${Math.round(drag.startTop + dy)}px`;
    } else {
      drag.node.style.width = `${Math.round(Math.max(24, Math.min(1920, drag.startWidth + dx)))}px`;
      drag.node.style.height = `${Math.round(Math.max(24, Math.min(1080, drag.startHeight + dy)))}px`;
    }
    updateOverlay();
    renderInspector();
  }

  function onFramePointerUp() {
    const drag = state.drag;
    state.drag = null;
    if (drag?.moved) recordHistory(drag.type === "move" ? "已移动元素" : "已调整元素尺寸");
  }

  function onFrameClick(event) {
    if (state.drag?.moved) return;
    const target = event.target;
    if (!target || target.nodeType !== 1) return;
    if (target.closest("[data-editor-selection]")) return;
    const root = getFrameRoot();
    if (!root || !root.contains(target)) return;
    setSelectedNode(target === root ? null : target);
  }

  function onFrameDoubleClick(event) {
    const target = event.target;
    if (!target || target.nodeType !== 1 || target === getFrameRoot() || target.tagName.toLowerCase() === "img") return;
    setSelectedNode(target);
    if (target.children.length > 0 || !target.textContent.trim()) return;
    target.contentEditable = "true";
    target.spellcheck = false;
    target.focus();
    const selection = target.ownerDocument.getSelection();
    const range = target.ownerDocument.createRange();
    range.selectNodeContents(target);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
    const finish = () => {
      target.contentEditable = "false";
      target.removeEventListener("blur", finish);
      recordHistory("已编辑文字");
      renderInspector();
    };
    target.addEventListener("blur", finish);
  }

  function bindFrameEvents() {
    const doc = dom.slideFrame.contentDocument;
    if (!doc || doc.body.dataset.bound === "true") return;
    doc.body.dataset.bound = "true";
    doc.addEventListener("click", onFrameClick);
    doc.addEventListener("dblclick", onFrameDoubleClick);
    doc.addEventListener("pointerdown", (event) => {
      const target = event.target;
      if (!target || target.nodeType !== 1 || target.closest("[data-editor-selection]")) return;
      if (event.detail > 1) return;
      const root = getFrameRoot();
      const node = target.closest("[data-editor-id]");
      if (node && node !== root && !target.isContentEditable) {
        setSelectedNode(node);
        beginMove(node, event);
      }
    });
    doc.addEventListener("pointermove", onFramePointerMove);
    doc.addEventListener("pointerup", onFramePointerUp);
    doc.addEventListener("pointercancel", onFramePointerUp);
    doc.addEventListener("keydown", (event) => {
      if (!state.selectedNodeId || event.target?.isContentEditable) return;
      if (!["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(event.key)) return;
      event.preventDefault();
      const node = getSelectedNode();
      if (!node) return;
      const pos = elementPosition(node);
      const step = event.shiftKey ? 10 : 1;
      const x = pos.x + (event.key === "ArrowLeft" ? -step : event.key === "ArrowRight" ? step : 0);
      const y = pos.y + (event.key === "ArrowUp" ? -step : event.key === "ArrowDown" ? step : 0);
      setElementGeometry({ x, y }, "已微调元素位置", true);
    });
  }

  function switchSlide(index) {
    if (!state.project.slides[index] || index === state.currentSlideIndex) return;
    syncFrameToProject();
    state.currentSlideIndex = index;
    state.selectedNodeId = null;
    renderSlideList();
    renderPageInspector();
    renderFrame();
    setStatus(`已切换到第 ${index + 1} 页。`);
  }

  function renderAll() {
    dom.projectTitleLabel.textContent = state.project.metadata?.title || "未命名演示文稿";
    renderSlideList();
    renderFrame();
    renderInspector();
    updateHistoryButtons();
    updateCanvasScale();
  }

  function loadProject(project, message = "项目已载入。") {
    state.project = project;
    state.currentSlideIndex = 0;
    state.selectedNodeId = null;
    renderAll();
    setTimeout(resetHistory, 60);
    const warningCount = state.project.warnings?.length || 0;
    const pageSummary = `已识别 ${state.project.slides.length} 页。`;
    setStatus(warningCount ? `${message} ${pageSummary} 已生成 ${warningCount} 条导入提示，请留意页面标记。` : `${message} ${pageSummary}`);
    if (warningCount) showToast(`载入完成：${warningCount} 项内容已按本地安全规则处理。`, "success");
  }

  function newSlide() {
    syncFrameToProject();
    const index = state.project.slides.length;
    const raw = `<section class="slide" data-slide-id="slide-${String(index + 1).padStart(2, "0")}" data-screen-label="${String(index + 1).padStart(2, "0")}" style="background:#ffffff"><h2 style="position:absolute;left:140px;top:120px;font-size:64px">新页面标题</h2><p style="position:absolute;left:145px;top:260px;width:1100px;font-size:34px;line-height:1.45;color:#667085">在这里添加页面内容。</p></section>`;
    const warnings = { add() {} };
    const slide = global.PPTDLSanitizer.sanitizeSlideHtml(raw, index, warnings);
    state.project.slides.push(slide);
    state.currentSlideIndex = index;
    state.selectedNodeId = null;
    renderSlideList();
    renderFrame();
    recordHistory("已新增页面");
  }

  function duplicateSlide() {
    syncFrameToProject();
    const source = currentSlide();
    if (!source) return;
    const index = state.currentSlideIndex + 1;
    const clone = deepClone(source);
    clone.id = `slide-${String(state.project.slides.length + 1).padStart(2, "0")}-${Date.now().toString(36).slice(-4)}`;
    clone.label = `${source.label || `页面 ${state.currentSlideIndex + 1}`} 副本`;
    const parser = new DOMParser();
    const doc = parser.parseFromString(clone.html, "text/html");
    const root = doc.body.firstElementChild;
    if (root) {
      root.setAttribute("data-slide-id", clone.id);
      root.setAttribute("id", clone.id);
      root.setAttribute("data-screen-label", clone.label);
      clone.html = root.outerHTML;
    }
    state.project.slides.splice(index, 0, clone);
    state.currentSlideIndex = index;
    state.selectedNodeId = null;
    renderSlideList();
    renderFrame();
    recordHistory("已复制页面");
  }

  function deleteSlide() {
    if (state.project.slides.length <= 1) {
      showToast("至少保留一页幻灯片。", "error");
      return;
    }
    if (!global.confirm(`确定删除第 ${state.currentSlideIndex + 1} 页吗？`)) return;
    syncFrameToProject();
    state.project.slides.splice(state.currentSlideIndex, 1);
    state.currentSlideIndex = Math.min(state.currentSlideIndex, state.project.slides.length - 1);
    state.selectedNodeId = null;
    renderSlideList();
    renderFrame();
    recordHistory("已删除页面");
  }

  function moveSlide(direction) {
    const from = state.currentSlideIndex;
    const to = from + direction;
    if (to < 0 || to >= state.project.slides.length) return;
    syncFrameToProject();
    const [slide] = state.project.slides.splice(from, 1);
    state.project.slides.splice(to, 0, slide);
    state.currentSlideIndex = to;
    renderSlideList();
    renderFrame();
    recordHistory(direction < 0 ? "已上移页面" : "已下移页面");
  }

  function updatePageLabel(value) {
    const slide = currentSlide();
    if (!slide) return;
    slide.label = value.trim() || String(state.currentSlideIndex + 1).padStart(2, "0");
    const root = getFrameRoot();
    if (root) root.setAttribute("data-screen-label", slide.label);
    renderSlideList();
    recordHistory("已更新页面名称");
  }

  function updatePageBackground(value) {
    const root = getFrameRoot();
    if (!root) return;
    root.style.backgroundColor = value;
    dom.pageBackgroundValue.textContent = value.toUpperCase();
    recordHistory("已更新页面背景");
  }

  function replaceImage(file) {
    const node = getSelectedNode();
    if (!node || node.tagName.toLowerCase() !== "img") {
      showToast("请先在画布中选择一张图片。", "error");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      node.src = String(reader.result || "");
      node.removeAttribute("srcset");
      node.removeAttribute("data-local-placeholder");
      recordHistory("已替换本地图片");
      showToast("图片已嵌入当前项目，不会上传。", "success");
    };
    reader.onerror = () => showToast("读取图片失败。", "error");
    reader.readAsDataURL(file);
  }

  function updateCanvasScale() {
    if (!dom.canvasViewport || !dom.canvasFrame) return;
    const width = dom.canvasViewport.clientWidth;
    const height = dom.canvasViewport.clientHeight;
    if (!width || !height) return;
    const scale = Math.min((width - 42) / 1920, (height - 42) / 1080, 1);
    dom.canvasFrame.style.transform = "translate(-50%, -50%)";
    // Chrome and Edge support CSS zoom, which keeps iframe hit-testing aligned
    // with the visible slide at reduced workspace sizes.
    dom.canvasFrame.style.zoom = String(Math.max(.16, scale));
    dom.zoomLabel.textContent = `${Math.round(scale * 100)}%`;
  }

  function isNarrowScreen() {
    return window.matchMedia && window.matchMedia("(max-width: 840px)").matches;
  }

  function makePromptContext() {
    return {
      topic: dom.aiTopicInput.value.trim() || "未指定主题",
      audience: dom.aiAudienceInput.value.trim() || "业务相关听众",
      duration: dom.aiDurationInput.value.trim() || "未指定时长",
      pages: dom.aiPagesInput.value.trim() || "8",
      language: dom.aiLanguageInput.value,
      use: dom.aiUseInput.value.trim() || "内部沟通",
      style: dom.aiStyleInput.value.trim() || "清晰、克制、专业",
      materials: dom.aiMaterialsInput.value.trim() || "请先询问我需要补充的事实，不要自行编造数据。",
    };
  }

  function generateOutlinePrompt() {
    const context = makePromptContext();
    dom.outlinePromptOutput.value = `你是资深演示文稿策划。请为以下需求设计一份 ${context.pages} 页左右的演示文稿大纲。\n\n主题：${context.topic}\n受众：${context.audience}\n演讲时长：${context.duration}\n语言：${context.language}\n用途：${context.use}\n视觉方向：${context.style}\n已有材料：\n${context.materials}\n\n请按“页码｜页面目的｜一句话结论｜关键内容｜建议视觉结构”的格式输出。只使用我提供的事实；缺失处标注“待补充”，不要编造业务数据、品牌事实或图片链接。`; 
    setAiStep(2);
  }

  function generateHtmlPrompt() {
    const context = makePromptContext();
    const outline = dom.outlineInput.value.trim() || "请根据上述需求自行组织合理的大纲，并对未提供的事实标注待补充。";
    dom.htmlPromptOutput.value = `请把下面的演示文稿大纲转换成一个可直接保存为 .html 的完整静态 HTML 幻灯片文件。\n\n【需求】\n主题：${context.topic}\n受众：${context.audience}\n时长：${context.duration}\n页数：${context.pages}\n语言：${context.language}\n用途：${context.use}\n风格：${context.style}\n\n【大纲】\n${outline}\n\n【必须遵守的 HTML 契约】\n1. 只返回完整 HTML，不要 Markdown 代码围栏、解释文字或外部链接。\n2. 使用 <main class="deck" data-deck>，每一页使用 <section class="slide" id="slide-01" data-slide-id="slide-01" data-screen-label="01">；id 与 data-slide-id 必须完全一致，页面专属 CSS 使用对应的 #slide-01 选择器。\n3. 每页固定 1920×1080、16:9；正文模块使用绝对定位，避免复杂脚本和 Canvas。\n4. 所有文字放在独立的 h1/h2/p/span 等叶子元素中，方便后续可视化编辑。画布是 1920×1080 CSS px（约 2px≈1pt），字号必须按信息密度和文本角色设置，不能把所有文字按同一倍率全局放大：\n   - 标准档：封面主标题 84–96px，普通页标题 58–68px，副标题/强调 40–48px，正文 30–36px，注释/页脚 24–28px，标签/表格 22–24px。\n   - 稀疏页可适度放大标题和关键数字；密集页正文可用 28–32px，但一般不要低于 22px。关键数字、标题、正文、注释应分别设置，不要用一个全局 font-size 或统一缩放因子代替层级。\n   - 字号变大时必须同步扩容文本框、增高卡片、换行或重新分栏；禁止只放大字号而保持原容器尺寸。\n   若页面同时包含 .section-label 与 h1/h2，章节标签放在 top:20px 左右，主标题放在 top:50px 或更低，两者必须留有间距且不得重叠。重复卡片、三栏信息和时间线必须使用统一的内边距、图标/标记列宽、文字起始 x 坐标与行高；同一组标题、标签和正文分别对齐到一致的左边线和基线，不要让图标列挤压文字或让时间线标记与文字重叠。\n   - 双栏布局必须显式划分左右列，列间至少保留 80px gutter；圆形主视觉、投票框或大图不能继续水平居中到右栏区域，必须给右侧文字面板预留完整宽度。\n5. CSS 写在 <style> 中，使用变量管理颜色和字号；不要加载字体、图片、CSS、JavaScript 或任何 http(s) 资源。图片位置用带 aria-label 的本地占位符或纯 CSS 图形。\n6. 每个 section.slide 必须默认独立可见；不要使用 display:none、visibility:hidden、opacity:0、data-active 或脚本来控制页面显示。\n7. 不要生成内部翻页按钮、导航圆点、nav-controls 或 nav-dot；PPT Design Lab 会负责缩略图与翻页。\n8. 不要使用 <script>、iframe、表单、动画、视频、外部 API 或自动请求。\n9. 只使用我提供的事实；未知内容标注“待补充”，不要编造。\n10. 生成 HTML 前和输出前都要做版式自检：检查元素重叠、内容裁切、越界、意外换行、标题与小标签重叠、卡片高度不足及左右栏间距；发现问题时优先调整容器尺寸、换行策略或分栏，不要只继续放大字号。\n\n请输出可以直接粘贴回 PPT Design Lab 的 HTML。`;
    setAiStep(3);
  }

  function setAiStep(step) {
    state.aiStep = Number(step) || 1;
    $$('[data-ai-pane]').forEach((pane) => { pane.hidden = pane.dataset.aiPane !== String(state.aiStep); });
    $$('[data-ai-step]').forEach((button) => { button.classList.toggle("is-active", button.dataset.aiStep === String(state.aiStep)); });
  }

  async function copyText(text, statusNode) {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
      } else {
        const helper = document.createElement("textarea");
        helper.value = text;
        helper.style.position = "fixed";
        helper.style.opacity = "0";
        document.body.appendChild(helper);
        helper.focus();
        helper.select();
        document.execCommand("copy");
        helper.remove();
      }
      statusNode.textContent = "已复制到剪贴板";
      statusNode.className = "copy-status is-success";
      showToast("提示词已复制。现在可以粘贴到 LorealGPT。", "success");
    } catch (error) {
      statusNode.textContent = "浏览器未允许自动复制，请手工选择文本";
      statusNode.className = "copy-status is-error";
    }
  }

  function importHTMLText(text, closeModal) {
    try {
      const project = global.PPTDLProjectIO.createFromHTML(text);
      if (closeModal) setModal(closeModal, false);
      loadProject(project, "HTML 已清洗并载入。");
      dom.htmlInput.value = "";
      dom.aiHtmlInput.value = "";
    } catch (error) {
      const message = error?.message || "HTML 导入失败。";
      if (closeModal === dom.importModal) {
        dom.importStatus.textContent = message;
        dom.importStatus.className = "modal-status is-error";
      } else {
        showToast(message, "error");
      }
    }
  }

  async function openHTMLFile(file) {
    try {
      const text = await global.PPTDLProjectIO.readFile(file);
      importHTMLText(text);
    } catch (error) {
      showToast(error.message || "HTML 文件打开失败。", "error");
    }
  }

  async function openProjectFile(file) {
    try {
      const text = await global.PPTDLProjectIO.readFile(file);
      const raw = JSON.parse(text);
      const project = global.PPTDLProjectIO.normalizeProject(raw);
      loadProject(project, "本地项目文件已打开。");
    } catch (error) {
      showToast(error.message || "项目文件打开失败。", "error");
    }
  }

  function handleSaveProject() {
    syncFrameToProject();
    global.PPTDLProjectIO.downloadProject(state.project);
    setDirty(false);
    setStatus("项目文件已下载到本机。");
  }

  function handleExportHTML() {
    syncFrameToProject();
    global.PPTDLExporter.downloadStandaloneHTML(state.project);
    setDirty(false);
    setStatus("独立 HTML 已下载到本机。");
  }

  async function handleExportPptx() {
    syncFrameToProject();
    const button = document.getElementById("exportPptxButton");
    button.disabled = true;
    button.textContent = "正在生成…";
    setStatus("正在逐页生成图片型 PPTX，请稍候。");
    try {
      const filename = await global.PPTDLExporter.exportPptx(state.project);
      setDirty(false);
      setStatus(`PPTX 已下载：${filename}`);
      showToast("PPTX 已生成并下载。页面内容仍只在本机处理。", "success");
    } catch (error) {
      const message = error?.message || "PPTX 导出失败。";
      setStatus(message, "error");
      showToast(message, "error");
    } finally {
      button.disabled = false;
      button.textContent = "导出 PPTX";
    }
  }

  function wireEvents() {
    document.getElementById("newProjectButton").addEventListener("click", () => {
      if (dom.dirtyIndicator.hidden || global.confirm("当前修改尚未保存，确定载入内置示例吗？")) loadProject(global.PPTDLProjectIO.makeDemoProject(), "已载入内置演示示例。");
    });
    document.getElementById("openAiButton").addEventListener("click", () => { setModal(dom.aiModal, true); setAiStep(1); });
    document.getElementById("importHtmlButton").addEventListener("click", () => { dom.importStatus.textContent = ""; dom.importStatus.className = "modal-status"; setModal(dom.importModal, true); });
    document.getElementById("openHtmlButton").addEventListener("click", () => dom.htmlFileInput.click());
    document.getElementById("openProjectButton").addEventListener("click", () => dom.projectFileInput.click());
    document.getElementById("saveProjectButton").addEventListener("click", handleSaveProject);
    document.getElementById("exportHtmlButton").addEventListener("click", handleExportHTML);
    document.getElementById("exportPptxButton").addEventListener("click", handleExportPptx);
    document.getElementById("addSlideButton").addEventListener("click", newSlide);
    document.getElementById("duplicatePageButton").addEventListener("click", duplicateSlide);
    document.getElementById("deletePageButton").addEventListener("click", deleteSlide);
    document.getElementById("moveSlideUpButton").addEventListener("click", () => moveSlide(-1));
    document.getElementById("moveSlideDownButton").addEventListener("click", () => moveSlide(1));
    document.getElementById("undoButton").addEventListener("click", undo);
    document.getElementById("redoButton").addEventListener("click", redo);
    document.getElementById("clearSelectionButton").addEventListener("click", () => setSelectedNode(null));
    dom.htmlFileInput.addEventListener("change", () => { openHTMLFile(dom.htmlFileInput.files[0]); dom.htmlFileInput.value = ""; });
    dom.projectFileInput.addEventListener("change", () => { openProjectFile(dom.projectFileInput.files[0]); dom.projectFileInput.value = ""; });
    dom.replaceImageInput.addEventListener("change", () => { replaceImage(dom.replaceImageInput.files[0]); dom.replaceImageInput.value = ""; });
    dom.pageLabelInput.addEventListener("change", () => updatePageLabel(dom.pageLabelInput.value));
    dom.pageBackgroundInput.addEventListener("input", () => updatePageBackground(dom.pageBackgroundInput.value));
    dom.pageBackgroundInput.addEventListener("change", () => recordHistory("已更新页面背景"));
    dom.elementTextInput.addEventListener("input", () => setElementText(dom.elementTextInput.value));
    dom.elementTextInput.addEventListener("blur", () => recordHistory("已更新文字内容"));
    dom.fontFamilyInput.addEventListener("change", () => applyElementStyle("fontFamily", dom.fontFamilyInput.value, "已更新字体", true));
    dom.fontSizeInput.addEventListener("input", () => applyElementStyle("fontSize", `${Math.max(8, Number(dom.fontSizeInput.value) || 24)}px`, "已更新字号"));
    dom.fontSizeInput.addEventListener("change", () => recordHistory("已更新字号"));
    dom.textColorInput.addEventListener("input", () => { dom.textColorValue.textContent = dom.textColorInput.value.toUpperCase(); applyElementStyle("color", dom.textColorInput.value, "已更新文字颜色"); });
    dom.textColorInput.addEventListener("change", () => recordHistory("已更新文字颜色"));
    dom.fontWeightInput.addEventListener("change", () => applyElementStyle("fontWeight", dom.fontWeightInput.value, "已更新字重", true));
    dom.textAlignInput.addEventListener("change", () => applyElementStyle("textAlign", dom.textAlignInput.value, "已更新对齐", true));
    dom.lineHeightInput.addEventListener("change", () => applyElementStyle("lineHeight", dom.lineHeightInput.value, "已更新行高", true));
    ["elementXInput", "elementYInput", "elementWInput", "elementHInput"].forEach((id) => dom[id].addEventListener("change", () => {
      setElementGeometry({ x: dom.elementXInput.value, y: dom.elementYInput.value, width: dom.elementWInput.value, height: dom.elementHInput.value }, "已更新元素布局", true);
    }));
    dom.alignLeftButton.addEventListener("click", () => setElementGeometry({ x: 0 }, "已左对齐元素", true));
    dom.alignCenterButton.addEventListener("click", () => { const node = getSelectedNode(); const root = getFrameRoot(); const width = node?.getBoundingClientRect().width || 0; setElementGeometry({ x: (1920 - width) / 2 }, "已水平居中元素", true); });
    dom.alignTopButton.addEventListener("click", () => setElementGeometry({ y: 0 }, "已顶端对齐元素", true));
    document.getElementById("sanitizeImportButton").addEventListener("click", () => importHTMLText(dom.htmlInput.value, dom.importModal));
    document.getElementById("generateOutlineButton").addEventListener("click", generateOutlinePrompt);
    document.getElementById("generateHtmlPromptButton").addEventListener("click", generateHtmlPrompt);
    document.getElementById("copyOutlinePromptButton").addEventListener("click", () => copyText(dom.outlinePromptOutput.value, dom.outlineCopyStatus));
    document.getElementById("copyHtmlPromptButton").addEventListener("click", () => copyText(dom.htmlPromptOutput.value, dom.htmlCopyStatus));
    document.getElementById("loadAiHtmlButton").addEventListener("click", () => importHTMLText(dom.aiHtmlInput.value, dom.aiModal));
    $$('[data-ai-step]').forEach((button) => button.addEventListener("click", () => setAiStep(button.dataset.aiStep)));
    document.getElementById("showPrivacyButton").addEventListener("click", () => setModal(dom.privacyModal, true));
    $$('[data-close-modal]').forEach((button) => button.addEventListener("click", () => setModal(document.getElementById(button.dataset.closeModal), false)));
    [dom.importModal, dom.aiModal, dom.privacyModal].forEach((modal) => modal.addEventListener("click", (event) => { if (event.target === modal) setModal(modal, false); }));
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") [dom.importModal, dom.aiModal, dom.privacyModal].forEach((modal) => { if (!modal.hidden) setModal(modal, false); });
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "z") { event.preventDefault(); event.shiftKey ? redo() : undo(); }
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "y") { event.preventDefault(); redo(); }
    });
    const resizeObserver = new ResizeObserver(updateCanvasScale);
    resizeObserver.observe(dom.canvasViewport);
    window.addEventListener("resize", updateCanvasScale);
  }

  function init() {
    initDom();
    state.project = global.PPTDLProjectIO.makeDemoProject();
    wireEvents();
    if (isNarrowScreen()) dom.desktopHint.hidden = false;
    renderAll();
    setTimeout(resetHistory, 80);
    setStatus("内置演示示例已载入。试试双击文字或拖动卡片。", "success");
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})(window);

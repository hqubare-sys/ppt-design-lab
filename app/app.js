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
    aiTheme: "consulting-blue",
    aiMaterialFiles: [],
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
      "aiTopicInput", "aiAudienceInput", "aiDurationInput", "aiPagesInput", "aiLanguageInput", "aiUseInput", "aiScenarioInput", "aiStyleInput", "aiMaterialsInput", "aiMaterialsFileInput", "aiMaterialsFileStatus",
      "htmlPromptOutput", "htmlCopyStatus", "aiHtmlInput", "toast", "desktopHint", "runDesignCheckButton", "startDesignCheckButton", "designCheckResults", "designCheckStatus",
    ].forEach((id) => { dom[id] = document.getElementById(id); });
    dom.importModal = document.getElementById("importModal");
    dom.aiModal = document.getElementById("aiModal");
    dom.privacyModal = document.getElementById("privacyModal");
    dom.designCheckModal = document.getElementById("designCheckModal");
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
    const theme = global.PPTDLDesignSystem?.getTheme(state.aiTheme);
    return {
      topic: dom.aiTopicInput.value.trim() || "未指定主题",
      audience: dom.aiAudienceInput.value.trim() || "业务相关听众",
      duration: dom.aiDurationInput.value.trim() || "未指定时长",
      pages: dom.aiPagesInput.value.trim() || "8",
      language: dom.aiLanguageInput.value,
      use: dom.aiUseInput.value.trim() || "内部沟通",
      scenarioId: dom.aiScenarioInput.value || "custom",
      scenarioName: dom.aiScenarioInput.options[dom.aiScenarioInput.selectedIndex]?.textContent || "自定义",
      style: dom.aiStyleInput.value.trim() || "清晰、克制、专业",
      materials: dom.aiMaterialsInput.value.trim() || "请先询问我需要补充的事实，不要自行编造数据。",
      themeId: theme?.id || state.aiTheme,
    };
  }

  function generatePrompt() {
    const context = makePromptContext();
    dom.htmlPromptOutput.value = global.PPTDLDesignSystem?.buildLorealPrompt(context) || `请根据以下需求生成完整静态 HTML 幻灯片文件：\n主题：${context.topic}\n页数：${context.pages}\n已有材料：${context.materials}`;
    setAiStep(2);
  }

  function setAiStep(step) {
    state.aiStep = Number(step) || 1;
    $$('[data-ai-pane]').forEach((pane) => { pane.hidden = pane.dataset.aiPane !== String(state.aiStep); });
    $$('[data-ai-step]').forEach((button) => { button.classList.toggle("is-active", button.dataset.aiStep === String(state.aiStep)); });
  }

  const TEXT_MATERIAL_LIMIT_BYTES = 2 * 1024 * 1024;
  const BINARY_MATERIAL_LIMIT_BYTES = 20 * 1024 * 1024;
  const MATERIAL_TOTAL_LIMIT_BYTES = 40 * 1024 * 1024;
  const TEXT_MATERIAL_EXTENSIONS = new Set(["txt", "md", "markdown", "csv", "tsv", "json", "html", "htm", "xml"]);

  function fileExtension(file) {
    return String(file?.name || "").split(".").pop().toLowerCase();
  }

  function readLocalText(file) {
    if (typeof file?.text === "function") return file.text();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(reader.error || new Error("读取文件失败。"));
      reader.readAsText(file);
    });
  }

  function htmlToPlainText(source) {
    const parser = new DOMParser();
    const parsed = parser.parseFromString(String(source || ""), "text/html");
    parsed.querySelectorAll("script,style,noscript").forEach((node) => node.remove());
    return (parsed.body?.textContent || parsed.documentElement?.textContent || "").replace(/\s+/g, " ").trim();
  }

  function parserResultEntries(value) {
    if (Array.isArray(value)) return value;
    if (Array.isArray(value?.files)) return value.files;
    if (Array.isArray(value?.results)) return value.results;
    return value ? [value] : [];
  }

  function parserEntryText(entry) {
    return String(entry?.text ?? entry?.content ?? entry?.plainText ?? entry?.body ?? "").trim();
  }

  function parserEntrySummary(entry, text) {
    const raw = entry?.summary;
    if (typeof raw === "string" && raw.trim()) return raw.trim();
    const details = [];
    if (entry?.pages != null) details.push(`${entry.pages} 页`);
    if (entry?.pageCount != null) details.push(`${entry.pageCount} 页`);
    if (entry?.sheets != null) details.push(`${entry.sheets} 个工作表`);
    if (entry?.sheetCount != null) details.push(`${entry.sheetCount} 个工作表`);
    if (entry?.worksheets != null) details.push(`${entry.worksheets} 个工作表`);
    if (entry?.characters != null) details.push(`${entry.characters} 字符`);
    if (entry?.charCount != null) details.push(`${entry.charCount} 字符`);
    if (!details.length && text) details.push(`${text.length} 字符`);
    return details.join("，");
  }

  async function readMaterialsFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    let totalBytes = 0;
    const validFiles = [];
    const rejected = [];
    files.forEach((file) => {
      const ext = fileExtension(file);
      if (!TEXT_MATERIAL_EXTENSIONS.has(ext) && !["pdf", "docx", "xlsx", "xls"].includes(ext)) {
        rejected.push(`${file.name}：格式不支持`);
        return;
      }
      const isText = TEXT_MATERIAL_EXTENSIONS.has(ext);
      const fileLimit = isText ? TEXT_MATERIAL_LIMIT_BYTES : BINARY_MATERIAL_LIMIT_BYTES;
      if (file.size > fileLimit) {
        rejected.push(`${file.name}：超过 ${isText ? "2MB" : "20MB"}`);
        return;
      }
      if (totalBytes + file.size > MATERIAL_TOTAL_LIMIT_BYTES) {
        rejected.push(`${file.name}：本次选择总计超过 40MB`);
        return;
      }
      totalBytes += file.size;
      validFiles.push(file);
    });
    if (!validFiles.length) {
      dom.aiMaterialsFileStatus.textContent = rejected.join("；") || "没有可读取的文件";
      dom.aiMaterialsFileStatus.className = "materials-file-status is-error";
      return;
    }
    dom.aiMaterialsFileStatus.textContent = `正在读取 ${validFiles.length} 个本地文件…`;
    dom.aiMaterialsFileStatus.className = "materials-file-status";
    const output = [];
    const fileWarnings = [];
    const parsedFiles = new Set();
    try {
      if (global.PPTDLFileParser?.parseFiles) {
        const parsed = parserResultEntries(await global.PPTDLFileParser.parseFiles(validFiles));
        parsed.forEach((entry, index) => {
          const fileName = String(entry?.name || entry?.fileName || validFiles[index]?.name || "本地材料");
          const text = parserEntryText(entry);
          const summary = parserEntrySummary(entry, text);
          parsedFiles.add(fileName);
          if (entry?.error || entry?.status === "error" || entry?.ok === false || !text) {
            const warning = Array.isArray(entry?.warnings) ? entry.warnings.join("；") : "";
            fileWarnings.push(`${fileName}：${entry?.error || entry?.message || warning || "未返回可用文字内容"}`);
            return;
          }
          const source = [`【本地材料：${fileName}${summary ? `｜${summary}` : ""}】`, text].filter(Boolean).join("\n");
          if (source) output.push(source);
        });
      }
      for (const file of validFiles) {
        if (parsedFiles.has(file.name)) continue;
        const ext = fileExtension(file);
        if (!TEXT_MATERIAL_EXTENSIONS.has(ext)) {
          fileWarnings.push(`${file.name}：当前环境未返回可用文字内容，请稍后重试或复制文字内容到材料框。`);
          continue;
        }
        const raw = await readLocalText(file);
        const text = ["html", "htm"].includes(ext) ? htmlToPlainText(raw) : raw.trim();
        if (text) output.push(`【本地材料：${file.name}｜${text.length} 字符】\n${text}`);
        else fileWarnings.push(`${file.name}：文件没有可用文字内容。`);
      }
      const existing = dom.aiMaterialsInput.value.trim();
      dom.aiMaterialsInput.value = [existing, ...output].filter(Boolean).join("\n\n");
      state.aiMaterialFiles = validFiles.map((file) => file.name);
      const accepted = validFiles.map((file) => file.name).join("、");
      const warnings = [...rejected, ...fileWarnings];
      const suffix = warnings.length ? `；${warnings.join("；")}` : "";
      dom.aiMaterialsFileStatus.textContent = `已读取：${accepted}${suffix}`;
      dom.aiMaterialsFileStatus.className = warnings.length ? "materials-file-status is-error" : "materials-file-status is-success";
    } catch (error) {
      dom.aiMaterialsFileStatus.textContent = error?.message || "本地材料读取失败";
      dom.aiMaterialsFileStatus.className = "materials-file-status is-error";
      showToast(dom.aiMaterialsFileStatus.textContent, "error");
    }
  }

  function escapeHtml(value) {
    return String(value || "").replace(/[&<>'"]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[character]));
  }

  function issueNodeLabel(node) {
    const text = String(node?.textContent || "").replace(/\s+/g, " ").trim();
    return text ? `“${text.slice(0, 34)}${text.length > 34 ? "…" : ""}”` : node?.tagName?.toLowerCase() || "元素";
  }

  function checkSlideDocument(doc) {
    const root = doc?.querySelector(".pptdlab-frame-root > .slide, .pptdlab-frame-root > [data-slide]");
    if (!root) return ["未找到可检查的幻灯片根节点。"];
    const win = doc.defaultView;
    const rootRect = root.getBoundingClientRect();
    const issues = [];
    root.querySelectorAll("*").forEach((node) => {
      if (node.hasAttribute("data-editor-selection") || node.id === "__editor_selection") return;
      const style = win.getComputedStyle(node);
      const label = issueNodeLabel(node);
      const isLeafText = node.children.length === 0 && node.textContent.trim() && !["svg", "path", "style"].includes(node.tagName.toLowerCase());
      const allowsSmallText = Boolean(node.closest('[data-design-check~="allow-small"], .system-slide'));
      if (isLeafText && !allowsSmallText && parseFloat(style.fontSize) < 22) issues.push(`${label}：字号 ${Math.round(parseFloat(style.fontSize))}px，小于 22px。`);
      if (isLeafText && node.clientWidth > 0 && node.clientHeight > 0 && (node.scrollWidth > node.clientWidth + 2 || node.scrollHeight > node.clientHeight + 2)) issues.push(`${label}：文字可能溢出文本框。`);
      const rect = node.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0 && (rect.left < rootRect.left - 1 || rect.top < rootRect.top - 1 || rect.right > rootRect.right + 1 || rect.bottom > rootRect.bottom + 1)) issues.push(`${label}：元素超出 1920×1080 画布。`);
    });
    root.querySelectorAll("[data-local-placeholder], .image-placeholder, [aria-label*='占位'], [aria-label*='待补充']").forEach((node) => {
      issues.push(`${issueNodeLabel(node)}：存在缺图或本地图片占位。`);
    });
    return Array.from(new Set(issues));
  }

  function inspectSlideLocally(slide) {
    return new Promise((resolve) => {
      const host = document.createElement("iframe");
      host.setAttribute("sandbox", "allow-same-origin");
      host.setAttribute("aria-hidden", "true");
      host.style.cssText = "position:fixed;left:-10000px;top:-10000px;width:1920px;height:1080px;border:0;visibility:hidden;pointer-events:none";
      let finished = false;
      const finish = () => {
        if (finished) return;
        finished = true;
        const issues = checkSlideDocument(host.contentDocument);
        host.remove();
        resolve(issues);
      };
      host.addEventListener("load", () => setTimeout(finish, 40), { once: true });
      document.body.appendChild(host);
      host.srcdoc = frameDocument(slide, state.project.css);
      setTimeout(finish, 1600);
    });
  }

  function renderDesignCheckResults(results) {
    const issueCount = results.reduce((sum, item) => sum + item.issues.length, 0);
    dom.designCheckResults.innerHTML = `<div class="check-summary${issueCount ? " has-issues" : ""}">${issueCount ? `发现 ${issueCount} 项需要留意的问题` : "检查完成：未发现基础布局问题"}</div>${results.map((item) => `<section class="check-page${item.issues.length ? "" : " is-clean"}"><div class="check-page-title">第 ${item.index + 1} 页 · ${escapeHtml(item.label)}<span>${item.issues.length ? `${item.issues.length} 项` : "通过"}</span></div>${item.issues.length ? `<ul>${item.issues.map((issue) => `<li>${escapeHtml(issue)}</li>`).join("")}</ul>` : `<div class="check-clean-label">字号、溢出、越界和缺图占位检查通过</div>`}</section>`).join("")}`;
    dom.designCheckStatus.textContent = `共检查 ${results.length} 页`;
  }

  async function runDesignCheck() {
    if (!state.project?.slides?.length) return;
    syncFrameToProject();
    dom.startDesignCheckButton.disabled = true;
    dom.designCheckStatus.textContent = "正在逐页检查…";
    dom.designCheckResults.innerHTML = "<div class=\"check-empty\">正在创建本地临时画布，请稍候…</div>";
    const results = [];
    for (let index = 0; index < state.project.slides.length; index += 1) {
      const slide = state.project.slides[index];
      dom.designCheckStatus.textContent = `正在检查第 ${index + 1} / ${state.project.slides.length} 页…`;
      results.push({ index, label: slide.label || `页面 ${index + 1}`, issues: await inspectSlideLocally(slide) });
    }
    renderDesignCheckResults(results);
    dom.startDesignCheckButton.disabled = false;
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
    dom.runDesignCheckButton.addEventListener("click", () => { dom.designCheckResults.innerHTML = "<div class=\"check-empty\">点击“开始检查”扫描当前项目。</div>"; dom.designCheckStatus.textContent = ""; setModal(dom.designCheckModal, true); });
    dom.startDesignCheckButton.addEventListener("click", runDesignCheck);
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
    dom.aiMaterialsFileInput.addEventListener("change", () => { readMaterialsFiles(dom.aiMaterialsFileInput.files); dom.aiMaterialsFileInput.value = ""; });
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
    document.getElementById("generatePromptButton").addEventListener("click", generatePrompt);
    document.getElementById("copyHtmlPromptButton").addEventListener("click", () => copyText(dom.htmlPromptOutput.value, dom.htmlCopyStatus));
    document.getElementById("loadAiHtmlButton").addEventListener("click", () => importHTMLText(dom.aiHtmlInput.value, dom.aiModal));
    $$("[data-theme-id]").forEach((button) => button.addEventListener("click", () => {
      state.aiTheme = button.dataset.themeId || "consulting-blue";
      $$("[data-theme-id]").forEach((card) => { const selected = card === button; card.classList.toggle("is-selected", selected); card.setAttribute("aria-pressed", String(selected)); });
    }));
    $$('[data-ai-step]').forEach((button) => button.addEventListener("click", () => setAiStep(button.dataset.aiStep)));
    document.getElementById("showPrivacyButton").addEventListener("click", () => setModal(dom.privacyModal, true));
    $$('[data-close-modal]').forEach((button) => button.addEventListener("click", () => setModal(document.getElementById(button.dataset.closeModal), false)));
    [dom.importModal, dom.aiModal, dom.privacyModal, dom.designCheckModal].forEach((modal) => modal.addEventListener("click", (event) => { if (event.target === modal) setModal(modal, false); }));
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") [dom.importModal, dom.aiModal, dom.privacyModal, dom.designCheckModal].forEach((modal) => { if (!modal.hidden) setModal(modal, false); });
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

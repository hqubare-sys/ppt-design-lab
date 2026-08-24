(function (global) {
  "use strict";

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function safeFilename(title, extension) {
    const base = String(title || "ppt-design-lab")
      .replace(/[^\w\-\u4e00-\u9fff]+/g, "-")
      .replace(/^-+|-+$/g, "") || "ppt-design-lab";
    return `${base}.${extension}`;
  }

  function buildStandaloneHTML(project) {
    const title = escapeHtml(project.metadata?.title || "PPT Design Lab 演示文稿");
    const css = String(project.css || "");
    const slides = (project.slides || []).map((slide) => slide.html).join("\n");
    const playerCSS = `
      html,body{margin:0;min-height:100%;background:#111827;color:#fff;font-family:Arial,sans-serif}
      body{display:grid;place-items:center;overflow:hidden}
      .pptdlab-player{position:relative;width:100vw;height:100vh;display:grid;place-items:center}
      .pptdlab-stage{position:relative;width:min(100vw,177.7778vh);height:min(56.25vw,100vh);overflow:hidden;background:#fff;color:#111}
      .pptdlab-stage>.slide{display:none!important;visibility:visible!important;opacity:1!important;width:1920px;height:1080px;transform-origin:top left}
      .pptdlab-stage>.slide.is-active{display:block!important}
      .pptdlab-controls{position:fixed;left:50%;bottom:18px;transform:translateX(-50%);display:flex;gap:8px;align-items:center;padding:8px 12px;border-radius:99px;background:rgba(17,24,39,.86);font:14px/1 Arial,sans-serif;z-index:10}
      .pptdlab-controls button{border:0;border-radius:6px;background:#374151;color:#fff;padding:6px 10px;cursor:pointer}
      .pptdlab-controls button:hover{background:#4b5563}
      .pptdlab-count{min-width:54px;text-align:center;color:#d1d5db}
      @media (max-width:720px){.pptdlab-controls{bottom:8px}.pptdlab-controls button{padding:8px}}
    `;
    const player = `
      <script>
        (function(){
          var stage=document.querySelector('.pptdlab-stage');
          var slides=[].slice.call(stage.querySelectorAll(':scope>.slide'));
          var index=0;
          var count=document.querySelector('.pptdlab-count');
          function fit(){var scale=Math.min(stage.clientWidth/1920,stage.clientHeight/1080);slides.forEach(function(s){s.style.transform='scale('+scale+')';});}
          function show(next){if(!slides.length)return;index=(next+slides.length)%slides.length;slides.forEach(function(s,i){s.classList.toggle('is-active',i===index);});count.textContent=(index+1)+' / '+slides.length;fit();}
          document.querySelector('[data-prev]').addEventListener('click',function(){show(index-1)});
          document.querySelector('[data-next]').addEventListener('click',function(){show(index+1)});
          document.querySelector('[data-fullscreen]').addEventListener('click',function(){(document.documentElement.requestFullscreen||function(){}).call(document.documentElement)});
          window.addEventListener('resize',fit);
          document.addEventListener('keydown',function(e){if(e.key==='ArrowRight'||e.key==='PageDown'||e.key===' '){e.preventDefault();show(index+1)}else if(e.key==='ArrowLeft'||e.key==='PageUp'){e.preventDefault();show(index-1)}else if(e.key==='Home'){show(0)}else if(e.key==='End'){show(slides.length-1)}});
          show(0);
        }());
      </script>
    `;
    return `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${title}</title><style>${css}${playerCSS}</style></head>
<body><main class="pptdlab-player"><div class="pptdlab-stage">${slides}</div><nav class="pptdlab-controls" aria-label="幻灯片控制"><button type="button" data-prev>上一页</button><span class="pptdlab-count">1 / 1</span><button type="button" data-next>下一页</button><button type="button" data-fullscreen>全屏</button></nav></main>${player}</body></html>`;
  }

  function downloadStandaloneHTML(project) {
    const html = buildStandaloneHTML(project);
    global.PPTDLProjectIO.triggerDownload(
      new Blob([html], { type: "text/html;charset=utf-8" }),
      safeFilename(project.metadata?.title, "html")
    );
  }

  function getPptxConstructor() {
    const candidate = global.PptxGenJS || global.pptxgen;
    if (typeof candidate === "function") return candidate;
    if (candidate && typeof candidate.default === "function") return candidate.default;
    return null;
  }

  function getHtmlToImage() {
    const candidate = global.htmlToImage || global.htmltoimage;
    if (!candidate || typeof candidate.toPng !== "function") return null;
    return candidate;
  }

  function renderFrameDocument(project, slideHtml) {
    return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>
      html,body{margin:0;padding:0;width:1920px;height:1080px;overflow:hidden;background:#fff}
      ${project.css || ""}
      body>.slide{display:block!important;visibility:visible!important;opacity:1!important;width:1920px;height:1080px}
    </style></head><body>${slideHtml}</body></html>`;
  }

  function waitForRenderFrames(view) {
    return new Promise((resolve) => {
      view.requestAnimationFrame(() => view.requestAnimationFrame(resolve));
    });
  }

  function isTransparentColor(value) {
    const input = String(value || "").trim().toLowerCase();
    if (!input || input === "transparent") return true;
    const rgba = input.match(/^rgba?\(\s*[^,]+\s*,\s*[^,]+\s*,\s*[^,]+(?:\s*,\s*([\d.]+)\s*)?\)$/);
    if (!rgba || rgba[1] == null) return false;
    return Number(rgba[1]) <= 0;
  }

  function exportBackgroundColor(root) {
    const doc = root?.ownerDocument;
    const view = doc?.defaultView;
    if (!doc || !view) return "#ffffff";
    const candidates = [root, doc.body, doc.documentElement].filter(Boolean);
    for (const node of candidates) {
      const computed = view.getComputedStyle(node);
      if (!isTransparentColor(computed.backgroundColor)) return computed.backgroundColor;
    }
    // Keep gradients in the DOM, but give transparent pixels a real page/body
    // color underneath them instead of letting html-to-image default to white.
    for (const node of candidates) {
      const computed = view.getComputedStyle(node);
      if (computed.backgroundImage && computed.backgroundImage !== "none") return "transparent";
    }
    return "#ffffff";
  }

  function hasExplicitLineBreak(node) {
    if (node.querySelector("br")) return true;
    const whiteSpace = node.ownerDocument?.defaultView?.getComputedStyle(node).whiteSpace || "normal";
    return /[\r\n]/.test(node.textContent || "") && /^(?:pre|pre-wrap|pre-line|break-spaces)$/.test(whiteSpace);
  }

  function hasSingleVisualLine(node) {
    const doc = node.ownerDocument;
    const range = doc.createRange();
    range.selectNodeContents(node);
    const rects = Array.from(range.getClientRects()).filter((rect) => rect.width > 0 && rect.height > 0);
    if (!rects.length) return false;
    const lineTops = [];
    rects.forEach((rect) => {
      if (!lineTops.some((top) => Math.abs(top - rect.top) < 1)) lineTops.push(rect.top);
    });
    return lineTops.length <= 1;
  }

  function stabilizeSingleLineText(root) {
    if (!root?.ownerDocument) return 0;
    const view = root.ownerDocument.defaultView;
    let count = 0;
    root.querySelectorAll("*").forEach((node) => {
      if (node.children.length || !node.textContent?.trim() || hasExplicitLineBreak(node)) return;
      const tag = node.tagName.toLowerCase();
      if (["script", "style", "textarea", "option", "svg", "path"].includes(tag)) return;
      const computed = view.getComputedStyle(node);
      if (computed.whiteSpace !== "normal") return;
      if (!hasSingleVisualLine(node)) return;
      node.style.whiteSpace = "nowrap";
      node.style.wordBreak = "keep-all";
      node.style.overflowWrap = "normal";
      count += 1;
    });
    return count;
  }

  async function makeRenderHost(project, slideHtml) {
    const host = document.createElement("iframe");
    host.className = "pptdlab-render-host";
    host.setAttribute("sandbox", "allow-same-origin");
    host.setAttribute("aria-hidden", "true");
    host.style.cssText = "position:fixed;left:-10000px;top:0;width:1920px;height:1080px;border:0;overflow:hidden;pointer-events:none;z-index:-1;";
    const loaded = new Promise((resolve, reject) => {
      host.onload = resolve;
      host.onerror = () => reject(new Error("PPTX 离屏渲染页面载入失败。"));
    });
    try {
      host.srcdoc = renderFrameDocument(project, slideHtml);
      document.body.appendChild(host);
      await loaded;
      const renderDocument = host.contentDocument;
      const renderWindow = host.contentWindow;
      if (!renderDocument || !renderWindow) throw new Error("浏览器未允许创建 PPTX 离屏渲染页面。");
      const slide = renderDocument.body.firstElementChild;
      if (!slide) throw new Error("页面内容为空，无法渲染 PPTX。");
      slide.querySelectorAll("[data-editor-selection],#__editor_selection").forEach((node) => node.remove());
      if (renderDocument.fonts?.ready) await renderDocument.fonts.ready;
      await waitForRenderFrames(renderWindow);
      stabilizeSingleLineText(slide);
      return { host, slide };
    } catch (error) {
      host.remove();
      throw error;
    }
  }

  async function renderSlideToPng(project, slide) {
    const htmlToImage = getHtmlToImage();
    if (!htmlToImage) throw new Error("未加载 html-to-image 本地依赖，暂不能导出图片型 PPTX。");
    const { host, slide: root } = await makeRenderHost(project, slide.html);
    try {
      return await htmlToImage.toPng(root, {
        width: 1920,
        height: 1080,
        pixelRatio: 1,
        cacheBust: false,
        skipFonts: true,
        backgroundColor: exportBackgroundColor(root),
      });
    } finally {
      host.remove();
    }
  }

  async function exportPptx(project) {
    const PptxGenJS = getPptxConstructor();
    if (!PptxGenJS) throw new Error("未加载 PptxGenJS 本地依赖，暂不能导出 PPTX。请确认 app/vendor 文件已完整部署。");
    if (!getHtmlToImage()) throw new Error("未加载 html-to-image 本地依赖，暂不能导出图片型 PPTX。请确认 app/vendor 文件已完整部署。");
    if (!Array.isArray(project.slides) || !project.slides.length) throw new Error("没有可导出的幻灯片。");

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "PPT Design Lab";
    pptx.company = "Local-first tool";
    pptx.subject = "Image-based presentation export";
    pptx.title = project.metadata?.title || "PPT Design Lab";
    pptx.lang = "zh-CN";
    for (const slide of project.slides) {
      const dataUrl = await renderSlideToPng(project, slide);
      const page = pptx.addSlide();
      page.background = { color: "FFFFFF" };
      page.addImage({ data: dataUrl, x: 0, y: 0, w: 13.333333, h: 7.5 });
    }
    const filename = safeFilename(project.metadata?.title, "pptx");
    if (typeof pptx.writeFile !== "function") throw new Error("当前 PptxGenJS 版本不支持浏览器文件导出。");
    await pptx.writeFile({ fileName: filename });
    return filename;
  }

  global.PPTDLExporter = {
    buildStandaloneHTML,
    downloadStandaloneHTML,
    exportPptx,
    renderSlideToPng,
    exportBackgroundColor,
    stabilizeSingleLineText,
    getPptxConstructor,
    getHtmlToImage,
  };
})(window);

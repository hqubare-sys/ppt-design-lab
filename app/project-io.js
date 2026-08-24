(function (global) {
  "use strict";

  const VERSION = "0.1.0";
  const MAX_HISTORY = 40;

  function now() {
    return new Date().toISOString();
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function deepClone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function makeProjectFromSanitized(result) {
    const createdAt = now();
    return {
      metadata: {
        title: result.title || "未命名演示文稿",
        createdAt,
        updatedAt: createdAt,
        version: VERSION,
      },
      headHtml: result.headHtml || "",
      css: result.css || "",
      slides: (result.slides || []).map((slide, index) => ({
        id: slide.id || `slide-${String(index + 1).padStart(2, "0")}`,
        label: slide.label || String(index + 1).padStart(2, "0"),
        html: slide.html,
      })),
      warnings: Array.from(new Set(result.warnings || [])),
    };
  }

  function createFromHTML(html) {
    return makeProjectFromSanitized(global.PPTDLSanitizer.sanitizeDocument(html));
  }

  function normalizeProject(input) {
    if (!input || typeof input !== "object") throw new Error("项目文件不是有效的 JSON 对象。");
    const rawSlides = Array.isArray(input.slides) ? input.slides : [];
    if (!rawSlides.length) throw new Error("项目文件中没有幻灯片。");
    const warningSet = new Set(Array.isArray(input.warnings) ? input.warnings.map(String) : []);
    const normalizedSlides = rawSlides.map((rawSlide, index) => {
      if (!rawSlide || typeof rawSlide.html !== "string") throw new Error(`第 ${index + 1} 页内容无效。`);
      const slide = global.PPTDLSanitizer.sanitizeSlideHtml(rawSlide.html, index, {
        add(message) {
          warningSet.add(message);
        },
      });
      slide.id = String(rawSlide.id || slide.id || `slide-${String(index + 1).padStart(2, "0")}`);
      slide.label = String(rawSlide.label || slide.label || String(index + 1).padStart(2, "0"));
      return slide;
    });
    const timestamp = now();
    const title = String(input.metadata?.title || "未命名演示文稿");
    return {
      metadata: {
        title,
        createdAt: String(input.metadata?.createdAt || timestamp),
        updatedAt: timestamp,
        version: VERSION,
      },
      // Do not trust headHtml from a project file. Reconstruct only the
      // minimal inert head metadata; exporter currently uses metadata/css.
      headHtml: `<meta charset="utf-8"><title>${escapeHtml(title)}</title>`,
      css: global.PPTDLSanitizer.safeCss(String(input.css || ""), {
        add(message) {
          warningSet.add(message);
        },
      }),
      slides: normalizedSlides,
      warnings: Array.from(warningSet),
    };
  }

  function snapshot(project) {
    return deepClone(project);
  }

  function updateTimestamp(project) {
    project.metadata = project.metadata || {};
    project.metadata.updatedAt = now();
    project.metadata.version = VERSION;
  }

  function serialize(project) {
    updateTimestamp(project);
    return JSON.stringify(project, null, 2);
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.rel = "noopener";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function downloadProject(project) {
    const title = (project.metadata?.title || "ppt-design-lab")
      .replace(/[^\w\-\u4e00-\u9fff]+/g, "-")
      .replace(/^-+|-+$/g, "") || "ppt-design-lab";
    const filename = `${title}.pptdlab`;
    triggerDownload(new Blob([serialize(project)], { type: "application/x-ppt-design-lab+json;charset=utf-8" }), filename);
    return filename;
  }

  function readFile(file) {
    return new Promise((resolve, reject) => {
      if (!file) return reject(new Error("没有选择文件。"));
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(new Error("读取文件失败。"));
      reader.readAsText(file);
    });
  }

  function makeDemoProject() {
    const demo = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>本地优先的演示文稿</title>
  <style>
    :root { --ink:#172033; --muted:#667085; --accent:#e83e65; --paper:#fff8f6; --soft:#f5d7df; }
    * { box-sizing:border-box; }
    html, body { margin:0; padding:0; }
    body { font-family: "Aptos", "Arial", sans-serif; color:var(--ink); }
    .deck { width:1920px; }
    .slide { position:relative; width:1920px; height:1080px; overflow:hidden; background:var(--paper); padding:110px 140px; }
    .slide::after { content:""; position:absolute; right:-160px; bottom:-220px; width:680px; height:680px; border-radius:50%; background:var(--soft); opacity:.7; }
    .eyebrow { position:absolute; top:110px; left:140px; color:var(--accent); font-size:28px; letter-spacing:4px; font-weight:700; text-transform:uppercase; }
    h1 { position:absolute; left:140px; top:260px; width:1260px; margin:0; font-size:96px; line-height:1.05; letter-spacing:-2px; }
    h2 { position:absolute; left:140px; top:120px; width:1300px; margin:0; font-size:64px; line-height:1.1; }
    p { position:absolute; left:145px; top:620px; width:900px; margin:0; font-size:36px; line-height:1.45; color:var(--muted); }
    .card { position:absolute; left:140px; top:370px; width:760px; min-height:280px; padding:48px; border-radius:30px; background:#fff; box-shadow:0 18px 50px rgba(23,32,51,.12); }
    .card strong { display:block; font-size:44px; margin-bottom:18px; }
    .card span { display:block; color:var(--muted); font-size:28px; line-height:1.4; }
    .stat { position:absolute; right:220px; top:320px; width:430px; height:300px; padding:44px; border-radius:30px; background:var(--accent); color:#fff; }
    .stat b { display:block; font-size:100px; line-height:1; }
    .stat span { display:block; margin-top:20px; font-size:28px; }
    .footer { position:absolute; left:140px; bottom:82px; color:var(--muted); font-size:24px; }
  </style>
</head>
<body>
  <main class="deck" data-deck>
    <section class="slide" data-slide-id="slide-01" data-screen-label="01">
      <div class="eyebrow">PPT DESIGN LAB</div>
      <h1>把 HTML 变成<br><span style="color:#e83e65">可编辑的演示文稿</span></h1>
      <p>这是一个安全、轻量、只在本机处理内容的工作台。先让 LorealGPT 生成结构，再在画布里完成最后一公里。</p>
      <div class="footer">演示页面 · 可直接拖动标题、编辑文字和替换图片</div>
    </section>
    <section class="slide" data-slide-id="slide-02" data-screen-label="02" style="background:#f7f9fc">
      <h2>一条清晰的本地工作流</h2>
      <div class="card"><strong>01 · 生成</strong><span>在公司批准的 LorealGPT 中手工生成结构化 HTML，不让业务内容经过第三方服务。</span></div>
      <div class="card" style="left:980px"><strong>02 · 编辑</strong><span>把 HTML 粘贴进工作台，拖动模块、调整字号、替换本地图片。</span></div>
      <div class="stat"><b>03</b><span>导出为独立 HTML 或图片型 PPTX</span></div>
      <div class="footer">本地演示数据 · 不代表公司业务内容</div>
    </section>
  </main>
</body>
</html>`;
    return createFromHTML(demo);
  }

  global.PPTDLProjectIO = {
    VERSION,
    MAX_HISTORY,
    deepClone,
    createFromHTML,
    makeProjectFromSanitized,
    makeDemoProject,
    normalizeProject,
    snapshot,
    updateTimestamp,
    serialize,
    triggerDownload,
    downloadProject,
    projectFilename(title) {
      const base = String(title || "ppt-design-lab")
        .replace(/[^\w\-\u4e00-\u9fff]+/g, "-")
        .replace(/^-+|-+$/g, "") || "ppt-design-lab";
      return `${base}.pptdlab`;
    },
    readFile,
  };
})(window);

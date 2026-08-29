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
  <title>PPT Design Lab · 本地优先演示</title>
  <style>
    :root {
      --ink:#172033;
      --muted:#657089;
      --accent:#e83e65;
      --accent-dark:#b82f51;
      --paper:#fff8f6;
      --paper-cool:#f7f9fc;
      --soft:#f5d7df;
      --line:#dfe5ee;
      --blue:#254d7d;
      --blue-soft:#eaf2fb;
      --gold:#c9992d;
      --gold-soft:#fbf2dc;
      --rice:#b85c43;
      --rice-soft:#f5e7df;
    }
    * { box-sizing:border-box; }
    html, body { margin:0; padding:0; }
    body { font-family:"Aptos","Arial","Microsoft YaHei",sans-serif; color:var(--ink); }
    .deck { width:1920px; }
    .slide { position:relative; width:1920px; height:1080px; overflow:hidden; }
    .slide-footer { position:absolute; left:120px; bottom:44px; color:#8994a8; font-size:22px; letter-spacing:1px; }

    /* Page 1: preserve the original built-in introduction. */
    .intro-slide { background:var(--paper); padding:110px 140px; }
    .intro-slide::after { content:""; position:absolute; right:-160px; bottom:-220px; width:680px; height:680px; border-radius:50%; background:var(--soft); opacity:.7; }
    .intro-eyebrow { position:absolute; top:110px; left:140px; color:var(--accent); font-size:28px; letter-spacing:4px; font-weight:700; text-transform:uppercase; }
    .intro-title { position:absolute; left:140px; top:260px; width:1260px; margin:0; font-size:96px; line-height:1.05; letter-spacing:-2px; }
    .intro-lead { position:absolute; left:145px; top:620px; width:900px; margin:0; font-size:36px; line-height:1.45; color:var(--muted); }
    .intro-start { position:absolute; right:170px; bottom:108px; color:var(--accent-dark); font-size:25px; font-weight:800; z-index:1; }
    .intro-footer { position:absolute; left:140px; bottom:82px; color:var(--muted); font-size:24px; }

    /* Page 2: the local quick-start path, using the same warm brand system as page 1. */
    .workflow-slide { background:var(--paper); }
    .workflow-slide::before { content:""; position:absolute; left:-230px; top:-300px; width:620px; height:620px; border-radius:50%; background:#ffe7ed; opacity:.74; }
    .workflow-slide::after { content:""; position:absolute; right:-180px; bottom:-260px; width:600px; height:600px; border-radius:50%; background:#f9e2e8; opacity:.72; }
    .workflow-kicker { position:absolute; left:120px; top:66px; color:var(--accent-dark); font-size:23px; font-weight:800; letter-spacing:3px; }
    .workflow-title { position:absolute; left:120px; top:120px; width:1250px; margin:0; color:var(--ink); font-size:64px; line-height:1.08; letter-spacing:-1px; }
    .workflow-lead { position:absolute; left:125px; top:206px; width:1450px; margin:0; color:var(--muted); font-size:30px; line-height:1.4; }
    .workflow-grid { position:absolute; left:120px; top:335px; display:grid; grid-template-columns:repeat(4,1fr); gap:28px; width:1680px; height:445px; }
    .workflow-step { position:relative; padding:31px 30px 27px; border:1px solid #f0cbd6; border-top:8px solid var(--accent); border-radius:26px; background:rgba(255,255,255,.96); box-shadow:0 16px 38px rgba(184,47,81,.09); }
    .workflow-step::after { content:""; position:absolute; left:30px; right:30px; bottom:84px; height:1px; background:#f2dfe4; }
    .workflow-number { display:flex; align-items:center; justify-content:center; width:58px; height:58px; border-radius:50%; background:var(--accent); color:#fff; font-size:24px; font-weight:800; }
    .workflow-step:nth-child(2) .workflow-number { background:#c85b78; }
    .workflow-step:nth-child(3) .workflow-number { background:#3d8a83; }
    .workflow-step:nth-child(4) .workflow-number { background:var(--gold); }
    .workflow-step h3 { margin:25px 0 0; color:var(--ink); font-size:32px; line-height:1.18; }
    .workflow-copy { display:block; margin-top:19px; color:var(--muted); font-size:23px; line-height:1.5; }
    .workflow-tag { position:absolute; left:30px; bottom:27px; color:var(--accent-dark); font-size:22px; font-weight:800; }
    .workflow-arrow { position:absolute; top:548px; width:40px; color:#e79aae; font-size:30px; font-weight:800; text-align:center; }
    .workflow-arrow-1 { left:524px; }
    .workflow-arrow-2 { left:952px; }
    .workflow-arrow-3 { left:1380px; }
    .workflow-local { position:absolute; left:120px; top:835px; display:flex; align-items:center; width:1680px; height:145px; padding:0 34px; border:1px solid #f0cbd6; border-radius:25px; background:rgba(255,255,255,.92); box-shadow:0 12px 30px rgba(184,47,81,.06); }
    .workflow-local-title { flex:0 0 325px; color:var(--accent-dark); font-size:25px; font-weight:800; line-height:1.3; }
    .workflow-path { flex:1; color:var(--ink); font-size:23px; line-height:1.45; }
    .workflow-path strong { color:var(--accent-dark); }
    .workflow-privacy { flex:0 0 370px; padding-left:28px; border-left:1px solid #f0d5dd; color:var(--muted); font-size:22px; line-height:1.42; }

    /* Page 3: the design-system map shown on the home page. */
    .system-slide { background:var(--paper); }
    .system-slide::after { content:""; position:absolute; right:-190px; top:-230px; width:530px; height:530px; border-radius:50%; background:#ffe7ed; opacity:.62; }
    .system-kicker { position:absolute; left:72px; top:52px; color:var(--accent); font-size:22px; font-weight:800; letter-spacing:3px; }
    .system-title { position:absolute; left:72px; top:100px; width:1150px; margin:0; color:var(--ink); font-size:58px; line-height:1.1; }
    .system-intro { position:absolute; left:77px; top:168px; width:1450px; margin:0; color:var(--muted); font-size:25px; line-height:1.4; }
    .system-callout { position:absolute; right:91px; top:52px; z-index:1; padding:13px 26px; border:1px solid #f1b7c5; border-radius:999px; color:var(--accent-dark); background:#fff9fa; font-size:20px; font-weight:800; line-height:1; }
    .system-panel { position:absolute; z-index:1; border:1px solid #f0dce2; border-radius:24px; background:#fff; box-shadow:0 12px 30px rgba(184,47,81,.07); }
    .panel-heading { display:block; margin:0; color:var(--ink); font-size:27px; line-height:1.1; }
    .panel-note { display:block; margin-top:6px; color:var(--muted); font-size:18px; line-height:1.2; }
    .themes-panel { left:72px; top:225px; width:405px; height:600px; padding:27px 25px; }
    .theme-stack { display:grid; gap:14px; margin-top:20px; }
    .theme-card { position:relative; height:145px; padding:16px 18px; overflow:hidden; border-radius:18px; box-shadow:0 8px 20px rgba(23,32,51,.08); }
    .theme-card::before { content:""; position:absolute; left:0; top:0; right:0; height:7px; border-radius:18px 18px 0 0; }
    .theme-card.consulting { background:var(--blue-soft); }
    .theme-card.consulting::before { background:var(--blue); }
    .theme-card.black-gold { background:#f3f1ed; }
    .theme-card.black-gold::before { background:#1d1d1f; }
    .theme-card.warm-rice { background:var(--rice-soft); }
    .theme-card.warm-rice::before { background:var(--rice); }
    .theme-card strong { display:block; color:var(--ink); font-size:27px; line-height:1.1; }
    .theme-card span { display:block; margin-top:5px; color:var(--muted); font-size:17px; line-height:1.2; }
    .theme-mini { position:absolute; left:17px; right:17px; bottom:13px; height:61px; padding:10px 12px; border-radius:12px; background:rgba(255,255,255,.82); }
    .theme-mini-lines { display:block; width:43%; height:7px; margin-top:0 !important; border-radius:99px; background:var(--blue); }
    .theme-mini-lines::after { content:""; display:block; width:67%; height:7px; margin-top:10px; border-radius:99px; background:currentColor; opacity:.78; }
    .consulting .theme-mini-lines { color:#3d8a83; }
    .black-gold .theme-mini-lines { width:49%; background:#d3ad50; color:#b4852a; }
    .warm-rice .theme-mini-lines { width:48%; background:#b85c43; color:#d18c68; }
    .theme-mini-bars { position:absolute; left:13px; right:auto; bottom:10px; display:flex; align-items:end; gap:5px; height:31px; margin-top:0 !important; }
    .theme-mini-bars i { display:block; width:13px; border-radius:3px 3px 0 0; background:#3d8a83; }
    .theme-mini-bars i:nth-child(1) { height:12px; }
    .theme-mini-bars i:nth-child(2) { height:22px; }
    .theme-mini-bars i:nth-child(3) { height:29px; }
    .theme-mini-bars i:nth-child(4) { height:18px; }
    .black-gold .theme-mini-bars i { background:#d3ad50; }
    .warm-rice .theme-mini-bars i { background:#c78663; }
    .theme-mini-kpi { position:absolute; right:12px; top:9px; display:grid; gap:4px; width:78px; margin-top:0 !important; color:transparent !important; font-size:0 !important; }
    .theme-mini-kpi i { display:block; height:20px; border-radius:5px; color:var(--blue); background:#edf3f7; font-size:12px; font-style:normal; font-weight:800; line-height:20px; text-align:center; }
    .black-gold .theme-mini-kpi i { color:#e8ce82; background:#433d2f; }
    .warm-rice .theme-mini-kpi i { color:#9f593f; background:#f2dfcf; }
    .structure-panel { left:501px; top:225px; width:734px; height:600px; padding:27px 27px; }
    .structure-list { display:grid; grid-template-columns:1fr 1fr; column-gap:14px; row-gap:10px; margin-top:19px; }
    .structure-item { display:flex; align-items:center; min-height:80px; padding:10px; border:1px solid #f0e8ea; border-radius:14px; color:var(--ink); background:#fff; }
    .structure-item b { display:flex; align-items:center; justify-content:center; flex:0 0 91px; width:91px; height:56px; margin-right:11px; border:1px solid #e8dfe2; border-radius:8px; background:#fbfcfd; color:transparent; font-size:0; position:relative; overflow:hidden; }
    .structure-item b::before { content:""; position:absolute; left:7px; right:7px; top:7px; height:7px; border-radius:2px; background:var(--accent); }
    .structure-item b::after { content:""; position:absolute; left:8px; right:8px; bottom:8px; height:27px; border-radius:3px; background:linear-gradient(90deg,#dce4ef 0 32%,transparent 32% 38%,#69a9a0 38% 68%,transparent 68% 74%,#dce4ef 74%); }
    .structure-item:nth-child(2n) b::before { background:var(--blue); }
    .structure-item:nth-child(3n) b::before { background:#3d8a83; }
    .structure-item:nth-child(1) b::after { background:#dce4ef; }
    .structure-item:nth-child(2) b::after { background:linear-gradient(90deg,#dce4ef 0 62%,#69a9a0 62%); }
    .structure-item:nth-child(3) b::after { background:linear-gradient(90deg,#dce4ef 0 47%,transparent 47% 53%,#dce4ef 53% 100%),linear-gradient(90deg,#69a9a0 0 47%,transparent 47% 53%,#dce4ef 53% 100%); background-size:100% 43%,100% 43%; background-position:0 0,0 100%; background-repeat:no-repeat; }
    .structure-item:nth-child(4) b::after { background:linear-gradient(90deg,#69a9a0 0 62%,#dce4ef 62%); }
    .structure-item:nth-child(5) b::after { background:linear-gradient(90deg,#dce4ef 0 29%,transparent 29% 36%,#d45873 36% 64%,transparent 64% 71%,#dce4ef 71%); }
    .structure-item:nth-child(6) b::after { background:linear-gradient(90deg,#dce4ef 0 47%,transparent 47% 53%,#dce4ef 53% 100%),linear-gradient(90deg,#dce4ef 0 47%,transparent 47% 53%,#d45873 53% 100%); background-size:100% 43%,100% 43%; background-position:0 0,0 100%; background-repeat:no-repeat; }
    .structure-item:nth-child(7) b::after { background:linear-gradient(90deg,#d45873 0 22%,#dce4ef 22% 44%,#69a9a0 44% 72%,#dce4ef 72%); }
    .structure-item:nth-child(8) b::after { background:linear-gradient(90deg,#dce4ef 0 22%,#254d7d 22% 47%,#dce4ef 47% 72%,#69a9a0 72%); }
    .structure-item:nth-child(9) b::after { background:linear-gradient(90deg,#d45873 0 30%,#dce4ef 30% 64%,#69a9a0 64%); }
    .structure-item:nth-child(10) b::after { background:#d45873; }
    .structure-item strong { display:block; font-size:20px; line-height:1.12; }
    .structure-item span { display:block; margin-top:4px; color:var(--muted); font-size:16px; line-height:1.12; }
    .module-panel { left:1259px; top:225px; width:552px; height:320px; padding:26px 24px; }
    .module-grid { display:grid; gap:6px; margin-top:15px; }
    .module-group { min-height:37px; padding:6px 10px; border-radius:11px; background:#f8fafc; }
    .module-group strong { display:block; color:var(--accent-dark); font-size:17px; line-height:1.1; }
    .module-chips { display:flex; flex-wrap:wrap; gap:5px; margin-top:4px; }
    .module-chips span { display:inline-block; padding:3px 7px; border:1px solid #e4e9ef; border-radius:999px; color:var(--muted); background:#fff; font-size:14px; line-height:1; }
    .blueprint-panel { left:1259px; top:565px; width:552px; height:260px; padding:20px 24px; border-color:#172033; background:var(--ink); box-shadow:0 12px 30px rgba(23,32,51,.14); }
    .blueprint-panel .panel-heading { color:#fff; }
    .blueprint-panel .panel-note { color:#c6cedd; }
    .blueprint-code { margin-top:10px; padding:8px 12px; border:1px solid rgba(255,255,255,.12); border-radius:12px; background:#101725; color:#f6cbd5; font-family:"SFMono-Regular","Consolas","Liberation Mono",monospace; font-size:12px; line-height:1.25; }
    .blueprint-code span { display:block; }
    .blueprint-code b { color:#f4ce6b; font-weight:700; }
    .blueprint-free { display:block; margin-top:8px; padding:6px 10px; border-radius:9px; color:#ffd7df; background:rgba(232,62,101,.23); font-size:13px; line-height:1.2; }
    .system-pipeline { position:absolute; left:72px; top:850px; z-index:1; display:grid; grid-template-columns:1fr 34px 1fr 34px 1fr 34px 1fr; align-items:center; width:1739px; height:155px; padding:20px 25px; border:1px solid #f0dce2; border-radius:24px; background:rgba(255,255,255,.9); box-shadow:0 12px 30px rgba(184,47,81,.06); }
    .pipeline-step { display:block; min-width:0; height:113px; padding:15px 16px; border:1px solid #f1e3e6; border-radius:16px; background:#fff; box-shadow:0 8px 18px rgba(23,32,51,.05); }
    .pipeline-step b { display:flex; align-items:center; justify-content:center; width:38px; height:38px; border-radius:50%; background:#fde0e7; color:var(--accent-dark); font-size:17px; }
    .pipeline-step strong { display:block; margin-top:10px; color:var(--ink); font-size:21px; line-height:1.1; }
    .pipeline-step span { display:block; margin-top:6px; color:var(--muted); font-size:16px; line-height:1.2; }
    .pipeline-arrow { color:var(--accent); font-size:27px; font-style:normal; text-align:center; }
    .system-slide .slide-footer { left:auto; right:91px; bottom:15px; font-size:18px; }
  </style>
</head>
<body>
  <main class="deck" data-deck>
    <section class="slide intro-slide" data-slide-id="slide-01" data-screen-label="01">
      <div class="intro-eyebrow">PPT DESIGN LAB</div>
      <h1 class="intro-title">把 HTML 变成<br><span style="color:#e83e65">可编辑的演示文稿</span></h1>
      <p class="intro-lead">这是一个安全、轻量、只在本机处理内容的工作台。先让 LorealGPT 生成结构，再在画布里完成最后一公里。</p>
      <div class="intro-footer">演示页面 · 可直接拖动标题、编辑文字和替换图片 · 01 / 03</div>
      <div class="intro-start">从顶部「AI 生成」开始 →</div>
    </section>
    <section class="slide workflow-slide" data-slide-id="slide-02" data-screen-label="02">
      <div class="workflow-kicker">QUICK START / 本地工作流</div>
      <h2 class="workflow-title">四步完成一份本地 PPT</h2>
      <p class="workflow-lead">工具负责提示词、渲染和导出；LorealGPT 负责内容规划与自由创作，中间只通过你的剪贴板连接。</p>
      <div class="workflow-grid">
        <div class="workflow-step"><b class="workflow-number">01</b><h3>填写需求</h3><span class="workflow-copy">点击顶部「AI 生成」，填写主题、受众、页数与已有材料；可点「读取本地材料」载入 TXT、PDF、DOCX 或 Excel。</span><span class="workflow-tag">输入材料</span></div>
        <div class="workflow-step"><b class="workflow-number">02</b><h3>复制完整提示词</h3><span class="workflow-copy">工具一次生成完整提示词，点击「复制提示词」，粘贴到公司 LorealGPT。</span><span class="workflow-tag">手工复制</span></div>
        <div class="workflow-step"><b class="workflow-number">03</b><h3>粘贴返回 HTML</h3><span class="workflow-copy">LorealGPT 同一次完成规划和自由 HTML；把结果粘贴回「粘贴 HTML」。</span><span class="workflow-tag">载入结果</span></div>
        <div class="workflow-step"><b class="workflow-number">04</b><h3>编辑、保存、导出</h3><span class="workflow-copy">双击文字、拖动模块、替换图片；点击「保存项目」或导出 HTML / PPTX。</span><span class="workflow-tag">交付文件</span></div>
      </div>
      <div class="workflow-arrow workflow-arrow-1">→</div><div class="workflow-arrow workflow-arrow-2">→</div><div class="workflow-arrow workflow-arrow-3">→</div>
      <div class="workflow-local"><strong class="workflow-local-title">一条清晰的本地路径</strong><span class="workflow-path"><strong>公司电脑材料</strong> → 当前工具生成提示词 → <strong>LorealGPT（手工复制）</strong> → 返回 HTML → 本地编辑与导出</span><span class="workflow-privacy">不上传文件<br>不连接公司 AI<br>项目只保存到本机</span></div>
      <div class="slide-footer">内置演示 02 / 03 · 先从顶部「AI 生成」开始</div>
    </section>
    <section class="slide system-slide" data-slide-id="slide-03" data-screen-label="03" data-design-check="allow-small">
      <div class="system-kicker">PPT DESIGN LAB · CONCEPT MAP</div>
      <h2 class="system-title">高级视觉生成系统</h2>
      <p class="system-intro">LorealGPT 自由规划与创作，工具用可复用的视觉语言在本地生成、编辑与导出。</p>
      <div class="system-callout">一键生成优先 · 自由 HTML 作为自动能力</div>

      <div class="system-panel themes-panel">
        <strong class="panel-heading">01 · 三套视觉系统</strong>
        <span class="panel-note">决定气质，不限制结构</span>
        <div class="theme-stack">
          <div class="theme-card consulting">
            <strong>咨询蓝</strong>
            <span>理性、清晰、数据驱动</span>
            <div class="theme-mini"><span class="theme-mini-lines"></span><span class="theme-mini-bars"><i></i><i></i><i></i><i></i></span><span class="theme-mini-kpi"><i>72%</i><i>+18%</i></span></div>
          </div>
          <div class="theme-card black-gold">
            <strong>黑金</strong>
            <span>高级、权威、管理层表达</span>
            <div class="theme-mini"><span class="theme-mini-lines"></span><span class="theme-mini-bars"><i></i><i></i><i></i><i></i></span><span class="theme-mini-kpi"><i>3.2B</i><i>TOP 1</i></span></div>
          </div>
          <div class="theme-card warm-rice">
            <strong>暖米</strong>
            <span>温暖、编辑感、叙事友好</span>
            <div class="theme-mini"><span class="theme-mini-lines"></span><span class="theme-mini-bars"><i></i><i></i><i></i><i></i></span><span class="theme-mini-kpi"><i>Q3</i><i>NEXT</i></span></div>
          </div>
        </div>
      </div>

      <div class="system-panel structure-panel" aria-label="10 种页面结构原型">
        <strong class="panel-heading">02 · 10 种页面结构原型</strong>
        <span class="panel-note">构图骨架可自由组合</span>
        <div class="structure-list">
          <div class="structure-item"><b>01</b><div><strong>核心结论</strong><span>一句洞察＋关键证据</span></div></div>
          <div class="structure-item"><b>02</b><div><strong>主分析＋证据</strong><span>主视觉＋排名或指标</span></div></div>
          <div class="structure-item"><b>03</b><div><strong>数据仪表盘</strong><span>KPI＋趋势＋结构</span></div></div>
          <div class="structure-item"><b>04</b><div><strong>趋势＋原因</strong><span>变化、曲线、驱动因素</span></div></div>
          <div class="structure-item"><b>05</b><div><strong>对比与选择</strong><span>方案、时机、目标对比</span></div></div>
          <div class="structure-item"><b>06</b><div><strong>矩阵与组合</strong><span>客户、产品、优先级</span></div></div>
          <div class="structure-item"><b>07</b><div><strong>流程与方法</strong><span>步骤、业务链路、机制</span></div></div>
          <div class="structure-item"><b>08</b><div><strong>路线图</strong><span>阶段、里程碑、计划</span></div></div>
          <div class="structure-item"><b>09</b><div><strong>问题—原因—行动</strong><span>复盘与改善闭环</span></div></div>
          <div class="structure-item"><b>10</b><div><strong>总结与决策</strong><span>发现、建议、责任与时间</span></div></div>
        </div>
      </div>

      <div class="system-panel module-panel">
        <strong class="panel-heading">03 · 视觉模块库</strong>
        <span class="panel-note">25+ 视觉模块 · 第一批可按内容组合</span>
        <div class="module-grid">
          <div class="module-group"><strong>数据分析</strong><div class="module-chips"><span>大数字</span><span>折线图</span><span>柱状图</span><span>环形图</span><span>瀑布图</span><span>排名表</span><span>气泡矩阵</span></div></div>
          <div class="module-group"><strong>逻辑表达</strong><div class="module-chips"><span>战略支柱</span><span>问题树</span><span>因果链</span><span>因果瀑</span><span>风险矩阵</span></div></div>
          <div class="module-group"><strong>项目管理</strong><div class="module-chips"><span>时间线</span><span>路线图</span><span>里程碑</span><span>角色分工</span><span>行动计划</span></div></div>
          <div class="module-group"><strong>叙事与品牌</strong><div class="module-chips"><span>图片洞察</span><span>案例卡片</span><span>前后对比</span><span>结论条</span><span>章节页</span></div></div>
        </div>
      </div>

      <div class="system-panel blueprint-panel">
        <strong class="panel-heading">04 · 页面蓝图格式</strong>
        <span class="panel-note">LorealGPT 一次返回；特殊表达返回 free_html</span>
        <div class="blueprint-code">
          <span><b>page_goal:</b> 说明增长转向经销商驱动</span>
          <span><b>headline:</b> 经销商贡献 72% 销售额</span>
          <span><b>structure:</b> 主分析＋辅助证据</span>
          <span><b>main_visual:</b> 客户组合＋渠道矩阵</span>
          <span><b>support:</b> Top5 表＋环形图＋大数字</span>
          <span><b>takeaway:</b> 业务模式正在重塑</span>
          <span><b>#</b> 数据、文字、注释与配色一并返回</span>
        </div>
        <span class="blueprint-free">如需更特殊表达：同时返回 free_html，工具直接导入并继续编辑</span>
      </div>

      <div class="system-pipeline">
        <div class="pipeline-step"><b>01</b><strong>用户材料</strong><span>内容、数据、目标、受众</span></div>
        <em class="pipeline-arrow">→</em>
        <div class="pipeline-step"><b>02</b><strong>工具生成提示词</strong><span>附带主题、结构、模块与蓝图协议</span></div>
        <em class="pipeline-arrow">→</em>
        <div class="pipeline-step"><b>03</b><strong>LorealGPT 自由创作</strong><span>完成分析、叙事、页面规划与 HTML</span></div>
        <em class="pipeline-arrow">→</em>
        <div class="pipeline-step"><b>04</b><strong>本地生成与编辑</strong><span>渲染、检查、拖动修改、重新设计、导出</span></div>
      </div>
      <div class="slide-footer">内置演示 03 / 03 · 主题、结构和模块是方法，不是公司内容</div>
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

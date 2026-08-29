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

    /* Page 2: the local quick-start path. */
    .workflow-slide { background:var(--paper-cool); }
    .workflow-slide::before { content:""; position:absolute; left:-230px; top:-300px; width:620px; height:620px; border-radius:50%; background:#e9eef7; opacity:.78; }
    .workflow-kicker { position:absolute; left:120px; top:66px; color:var(--blue); font-size:23px; font-weight:800; letter-spacing:3px; }
    .workflow-title { position:absolute; left:120px; top:120px; width:1250px; margin:0; color:var(--ink); font-size:64px; line-height:1.08; letter-spacing:-1px; }
    .workflow-lead { position:absolute; left:125px; top:206px; width:1450px; margin:0; color:var(--muted); font-size:30px; line-height:1.4; }
    .workflow-grid { position:absolute; left:120px; top:335px; display:grid; grid-template-columns:repeat(4,1fr); gap:28px; width:1680px; height:445px; }
    .workflow-step { position:relative; padding:31px 30px 27px; border:1px solid #d6deea; border-radius:26px; background:#fff; box-shadow:0 16px 38px rgba(36,60,96,.09); }
    .workflow-step::after { content:""; position:absolute; left:30px; right:30px; bottom:84px; height:1px; background:#e5eaf2; }
    .workflow-step:nth-child(1) { border-top:8px solid var(--accent); }
    .workflow-step:nth-child(2) { border-top:8px solid var(--blue); }
    .workflow-step:nth-child(3) { border-top:8px solid #3d8a83; }
    .workflow-step:nth-child(4) { border-top:8px solid var(--gold); }
    .workflow-number { display:flex; align-items:center; justify-content:center; width:58px; height:58px; border-radius:50%; background:var(--ink); color:#fff; font-size:24px; font-weight:800; }
    .workflow-step:nth-child(1) .workflow-number { background:var(--accent); }
    .workflow-step:nth-child(2) .workflow-number { background:var(--blue); }
    .workflow-step:nth-child(3) .workflow-number { background:#3d8a83; }
    .workflow-step:nth-child(4) .workflow-number { background:var(--gold); }
    .workflow-step h3 { margin:25px 0 0; font-size:32px; line-height:1.18; }
    .workflow-copy { display:block; margin-top:19px; color:var(--muted); font-size:23px; line-height:1.5; }
    .workflow-tag { position:absolute; left:30px; bottom:27px; color:var(--ink); font-size:22px; font-weight:800; }
    .workflow-arrow { position:absolute; top:548px; width:40px; color:#aeb9ca; font-size:34px; font-weight:800; text-align:center; }
    .workflow-arrow-1 { left:524px; }
    .workflow-arrow-2 { left:952px; }
    .workflow-arrow-3 { left:1380px; }
    .workflow-local { position:absolute; left:120px; top:835px; display:flex; align-items:center; width:1680px; height:145px; padding:0 34px; border:1px solid #c8d4e5; border-radius:25px; background:#eef4fb; }
    .workflow-local-title { flex:0 0 325px; color:var(--blue); font-size:25px; font-weight:800; line-height:1.3; }
    .workflow-path { flex:1; color:var(--ink); font-size:23px; line-height:1.45; }
    .workflow-path strong { color:var(--blue); }
    .workflow-privacy { flex:0 0 370px; padding-left:28px; border-left:1px solid #c8d4e5; color:var(--muted); font-size:22px; line-height:1.42; }

    /* Page 3: the design-system map shown on the home page. */
    .system-slide { background:#fbfcfe; }
    .system-kicker { position:absolute; left:120px; top:58px; color:var(--gold); font-size:22px; font-weight:800; letter-spacing:3px; }
    .system-title { position:absolute; left:120px; top:106px; width:1200px; margin:0; color:var(--ink); font-size:58px; line-height:1.1; }
    .system-intro { position:absolute; left:125px; top:176px; width:1680px; margin:0; color:var(--muted); font-size:25px; line-height:1.4; }
    .theme-strip { position:absolute; left:120px; top:250px; display:grid; grid-template-columns:repeat(3,1fr); gap:24px; width:1680px; height:120px; }
    .theme-card { position:relative; padding:23px 28px; border-radius:23px; background:#fff; box-shadow:0 10px 26px rgba(23,32,51,.08); }
    .theme-card::before { content:""; position:absolute; left:0; top:0; bottom:0; width:8px; border-radius:23px 0 0 23px; }
    .theme-card.consulting { background:var(--blue-soft); }
    .theme-card.consulting::before { background:var(--blue); }
    .theme-card.black-gold { background:#f3f1ed; }
    .theme-card.black-gold::before { background:#1d1d1f; }
    .theme-card.warm-rice { background:var(--rice-soft); }
    .theme-card.warm-rice::before { background:var(--rice); }
    .theme-card strong { display:block; font-size:29px; line-height:1.1; }
    .theme-card span { display:block; margin-top:9px; color:var(--muted); font-size:21px; line-height:1.2; }
    .system-panel { position:absolute; border:1px solid #dce3ed; border-radius:24px; background:#fff; box-shadow:0 12px 30px rgba(23,32,51,.06); }
    .panel-heading { display:block; margin:0; color:var(--ink); font-size:27px; line-height:1.1; }
    .panel-note { display:block; margin-top:6px; color:var(--muted); font-size:19px; line-height:1.2; }
    .structure-panel { left:120px; top:400px; width:800px; height:405px; padding:27px 31px; }
    .structure-list { display:grid; grid-template-columns:1fr 1fr; column-gap:34px; row-gap:9px; margin-top:22px; }
    .structure-item { display:flex; align-items:center; min-height:48px; border-bottom:1px solid #edf0f5; color:var(--ink); font-size:22px; line-height:1.2; }
    .structure-item b { display:inline-flex; align-items:center; justify-content:center; width:32px; height:32px; margin-right:11px; border-radius:50%; background:var(--gold-soft); color:var(--gold); font-size:17px; }
    .module-panel { left:952px; top:400px; width:848px; height:200px; padding:26px 30px; }
    .module-grid { display:grid; grid-template-columns:repeat(2,1fr); gap:10px 18px; margin-top:17px; }
    .module-group { min-height:45px; padding:9px 13px; border-radius:13px; background:#f6f8fb; }
    .module-group strong { display:block; color:var(--blue); font-size:20px; line-height:1.1; }
    .module-group span { display:block; margin-top:3px; color:var(--muted); font-size:18px; line-height:1.1; }
    .blueprint-panel { left:952px; top:612px; width:848px; height:220px; padding:26px 30px; }
    .blueprint-fields { display:flex; flex-wrap:wrap; gap:10px; margin-top:18px; }
    .blueprint-field { padding:9px 14px; border:1px solid #d8e0ec; border-radius:999px; color:var(--blue); background:#f6f9fd; font-size:20px; line-height:1; }
    .system-pipeline { position:absolute; left:120px; top:850px; display:flex; align-items:center; width:1680px; height:145px; padding:0 30px; border-radius:24px; background:var(--ink); color:#fff; }
    .pipeline-label { flex:0 0 190px; color:#f4ce6b; font-size:23px; font-weight:800; line-height:1.25; }
    .pipeline-step { display:flex; align-items:center; flex:1; min-width:0; }
    .pipeline-step b { display:flex; align-items:center; justify-content:center; flex:0 0 42px; width:42px; height:42px; border-radius:50%; background:#fff; color:var(--ink); font-size:18px; }
    .pipeline-step span { display:block; margin-left:12px; color:#fff; font-size:21px; line-height:1.3; }
    .pipeline-arrow { flex:0 0 40px; color:#f4ce6b; font-size:28px; text-align:center; }
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
      <div class="system-kicker">DESIGN SYSTEM / 自由创作的边界</div>
      <h2 class="system-title">高级视觉生成系统</h2>
      <p class="system-intro">不把 PPT 锁死在一套模板里：LorealGPT 自由规划页面，工具用主题、结构和视觉模块把想法稳定地画出来。</p>
      <div class="theme-strip">
        <div class="theme-card consulting"><strong>咨询蓝</strong><span>清晰、克制、适合运营分析与管理层汇报</span></div>
        <div class="theme-card black-gold"><strong>黑金</strong><span>高对比、强调结论，适合战略与提案</span></div>
        <div class="theme-card warm-rice"><strong>暖米</strong><span>温暖、编辑感，适合项目与活动叙事</span></div>
      </div>
      <div class="system-panel structure-panel"><strong class="panel-heading">10 种页面结构原型</strong><span class="panel-note">先确定信息关系，再让视觉自由组合</span><div class="structure-list"><div class="structure-item"><b>01</b>核心结论</div><div class="structure-item"><b>06</b>矩阵与组合</div><div class="structure-item"><b>02</b>主分析＋辅助证据</div><div class="structure-item"><b>07</b>流程与方法</div><div class="structure-item"><b>03</b>数据仪表盘</div><div class="structure-item"><b>08</b>路线图与时间线</div><div class="structure-item"><b>04</b>趋势＋原因</div><div class="structure-item"><b>09</b>问题—原因—行动</div><div class="structure-item"><b>05</b>对比与选择</div><div class="structure-item"><b>10</b>总结与决策</div></div></div>
      <div class="system-panel module-panel"><strong class="panel-heading">25+ 视觉模块 · 四大模块组</strong><span class="panel-note">按内容需要组合，不要求每页长得一样</span><div class="module-grid"><div class="module-group"><strong>数据分析</strong><span>KPI · 趋势图 · 排名表 · 矩阵</span></div><div class="module-group"><strong>逻辑表达</strong><span>问题树 · 因果链 · 流程图 · 风险矩阵</span></div><div class="module-group"><strong>项目管理</strong><span>时间线 · 路线图 · 里程碑 · 行动计划</span></div><div class="module-group"><strong>叙事与品牌</strong><span>图片卡 · 案例 · 结论条 · 前后对比</span></div></div></div>
      <div class="system-panel blueprint-panel"><strong class="panel-heading">页面蓝图 · LorealGPT 一次返回</strong><span class="panel-note">蓝图告诉工具“这一页要表达什么”，自由 HTML 同时作为特殊视觉结果返回</span><div class="blueprint-fields"><span class="blueprint-field">page_goal</span><span class="blueprint-field">key_message</span><span class="blueprint-field">structure</span><span class="blueprint-field">modules</span><span class="blueprint-field">data</span><span class="blueprint-field">copy</span><span class="blueprint-field">visual_style</span><span class="blueprint-field">free_html</span></div></div>
      <div class="system-pipeline"><strong class="pipeline-label">一键生成逻辑</strong><div class="pipeline-step"><b>01</b><span>用户材料<br>与数据</span></div><em class="pipeline-arrow">→</em><div class="pipeline-step"><b>02</b><span>工具生成<br>完整提示词</span></div><em class="pipeline-arrow">→</em><div class="pipeline-step"><b>03</b><span>LorealGPT<br>自由创作</span></div><em class="pipeline-arrow">→</em><div class="pipeline-step"><b>04</b><span>本地工具生成、<br>编辑与导出</span></div></div>
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

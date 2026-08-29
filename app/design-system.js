(function (global) {
  "use strict";

  // This file is deliberately data-first. The UI uses it to render the theme
  // cards and the prompt builder uses the same names and contracts, so what a
  // user selects is exactly what LorealGPT receives.
  const themes = [
    {
      id: "consulting-blue",
      name: "咨询蓝",
      shortName: "CONSULTING BLUE",
      description: "清晰、克制、证据驱动，适合运营报表、战略与管理层汇报。",
      palette: {
        ink: "#102A56",
        accent: "#1B5EA7",
        secondary: "#2A9D9F",
        highlight: "#D9A928",
        paper: "#F7FAFD",
        panel: "#FFFFFF",
        soft: "#EAF2FA",
        muted: "#5B6B84",
      },
      visualRules: [
        "白底或极浅冷灰底，深蓝标题，青绿色作为第二数据色。",
        "使用细线、清晰网格和少量圆角；避免大面积渐变和厚重阴影。",
        "优先使用结论型标题、主分析图、辅助证据和底部 Key Takeaway。",
        "图表强调对比关系与变化，标签短、图例少，保留足够留白。",
      ],
    },
    {
      id: "black-gold",
      name: "黑金",
      shortName: "BLACK & GOLD",
      description: "沉稳、聚焦、有仪式感，适合战略提案、决策和重要成果。",
      palette: {
        ink: "#F7F1E6",
        accent: "#D5A83A",
        secondary: "#B9822B",
        highlight: "#F2C866",
        paper: "#111722",
        panel: "#1B2432",
        soft: "#2B3443",
        muted: "#B9C0CB",
      },
      visualRules: [
        "深炭黑或墨蓝底，金色只用于关键数字、结论和行动提示。",
        "用大留白、强对比和少量细金线建立高级感；不要把所有文字做成金色。",
        "每页只保留一个视觉焦点，辅助数据用低对比灰白层级呈现。",
        "图片和图表采用深色底适配方式，避免白色卡片堆叠成仪表盘。",
      ],
    },
    {
      id: "warm-editorial",
      name: "暖米",
      shortName: "WARM EDITORIAL",
      description: "温暖、从容、有编辑感，适合活动方案、品牌故事和工作复盘。",
      palette: {
        ink: "#28384C",
        accent: "#B84C32",
        secondary: "#D08B35",
        highlight: "#E7B541",
        paper: "#FBF4E9",
        panel: "#FFFDFC",
        soft: "#F1E2CF",
        muted: "#6C7280",
      },
      visualRules: [
        "浅米色作为页面底色，砖红用于标题和重点，暖金用于小面积强调。",
        "使用杂志式留白、图片裁切和轻阴影；卡片边界要轻，避免拥挤。",
        "标题可以有更强的叙事感，正文保持短句和清晰的阅读节奏。",
        "适合图片、时间线、方案对比和结论条组合，但每页只保留一个主叙事。",
      ],
    },
  ];

  // Free-design marker used in the one-shot contract (kept as data so the
  // prompt and static checks share the same explicit vocabulary).
  const FREE_DESIGN_LAYOUT_ID = "free_html";
  const FREE_DESIGN_LABEL = "自由HTML";

  const layoutPrototypes = [
    { id: "core-conclusion", name: "核心结论", summary: "结论型标题 + 3 个支撑证据", bestFor: "管理层摘要、关键发现、工作汇报", modules: ["key-takeaway", "kpi-card", "evidence-list"] },
    { id: "analysis-evidence", name: "主分析＋辅助证据", summary: "左侧主图，右侧排名、指标或解释", bestFor: "运营报表、客户/产品组合、渠道分析", modules: ["matrix", "ranking-table", "donut-chart", "key-takeaway"] },
    { id: "dashboard", name: "数据仪表盘", summary: "KPI + 趋势 + 结构 + 异常", bestFor: "月报、季度报、运营看板", modules: ["kpi-card", "line-chart", "bar-chart", "alert-callout"] },
    { id: "trend-drivers", name: "趋势＋驱动因素", summary: "上方变化趋势，下方解释原因", bestFor: "增长/下降分析、业务复盘", modules: ["line-chart", "driver-cards", "waterfall-chart", "key-takeaway"] },
    { id: "comparison-choice", name: "对比与选择", summary: "两列或多列比较方案、目标和结果", bestFor: "方案提案、年度对比、选型决策", modules: ["comparison-card", "pros-cons", "recommendation-callout"] },
    { id: "matrix-portfolio", name: "矩阵与组合", summary: "用二维坐标或分层区域解释组合关系", bestFor: "客户、产品、优先级、风险机会", modules: ["matrix", "bubble-chart", "quadrant-labels", "legend"] },
    { id: "process-method", name: "流程与方法", summary: "步骤、角色和输入输出一眼可读", bestFor: "工作方法、业务链路、项目介绍", modules: ["process-flow", "role-lane", "input-output", "step-card"] },
    { id: "roadmap-timeline", name: "路线图与时间线", summary: "阶段、里程碑和责任人按时间展开", bestFor: "项目计划、战略落地、季度行动", modules: ["timeline", "milestone", "owner-chip", "status-tag"] },
    { id: "problem-action", name: "问题—原因—行动", summary: "把复盘从现象推进到可执行动作", bestFor: "运营复盘、问题解决、工作汇报", modules: ["issue-card", "cause-chain", "action-plan", "owner-chip"] },
    { id: "decision-summary", name: "总结与决策", summary: "发现、建议、下一步和待决事项", bestFor: "结尾页、决策页、资源申请", modules: ["decision-card", "recommendation-callout", "action-plan", "key-takeaway"] },
  ];

  const moduleGroups = {
    dataAnalysis: {
      name: "数据分析",
      modules: [
        { id: "kpi-card", name: "KPI 大数字", description: "突出单项指标、目标与变化。" },
        { id: "line-chart", name: "折线/面积趋势", description: "表现时间序列、目标与实际。" },
        { id: "bar-chart", name: "柱状/堆叠柱状图", description: "比较类别、构成与排序。" },
        { id: "donut-chart", name: "环形/结构图", description: "表达占比，配合关键数字。" },
        { id: "waterfall-chart", name: "瀑布图", description: "解释增减因素对最终结果的贡献。" },
        { id: "ranking-table", name: "排名表", description: "呈现 Top N、排名和关键属性。" },
        { id: "bubble-chart", name: "气泡图", description: "同时表达两维位置和规模大小。" },
        { id: "matrix", name: "矩阵/四象限", description: "按两个维度划分组合、机会或风险。" },
      ],
    },
    logicalExpression: {
      name: "逻辑表达",
      modules: [
        { id: "key-takeaway", name: "Key Takeaway 结论条", description: "在底部收束页面核心结论。" },
        { id: "evidence-list", name: "证据列表", description: "用短句和数值支撑主张。" },
        { id: "driver-cards", name: "驱动因素卡片", description: "拆解趋势背后的 2–4 个原因。" },
        { id: "cause-chain", name: "因果链/问题树", description: "从现象到原因再到影响。" },
        { id: "comparison-card", name: "对比卡片", description: "并列呈现方案、状态或选择。" },
        { id: "pros-cons", name: "优劣势/取舍", description: "清晰呈现利弊与适用边界。" },
        { id: "quadrant-labels", name: "象限标签", description: "为矩阵区域提供结论化命名。" },
      ],
    },
    projectManagement: {
      name: "项目管理",
      modules: [
        { id: "process-flow", name: "流程箭头", description: "展示步骤、依赖和输入输出。" },
        { id: "role-lane", name: "角色泳道", description: "按角色分配责任和动作。" },
        { id: "input-output", name: "输入/输出", description: "表达每个阶段的交付物。" },
        { id: "step-card", name: "步骤卡片", description: "将复杂方法拆成可读步骤。" },
        { id: "timeline", name: "时间线/路线图", description: "按月份、季度或阶段安排工作。" },
        { id: "milestone", name: "里程碑", description: "标出关键节点与完成标准。" },
        { id: "owner-chip", name: "责任人标签", description: "显示 owner、协作方或决策人。" },
        { id: "status-tag", name: "状态标签", description: "标记进行中、已完成、风险等状态。" },
      ],
    },
    narrativeBrand: {
      name: "叙事与品牌",
      modules: [
        { id: "issue-card", name: "问题卡片", description: "以清晰问题开启一页分析。" },
        { id: "recommendation-callout", name: "推荐/建议框", description: "突出建议、判断和下一步。" },
        { id: "action-plan", name: "行动计划", description: "包含动作、负责人、时间和状态。" },
        { id: "decision-card", name: "决策卡片", description: "呈现需要拍板的事项和选项。" },
        { id: "image-caption", name: "图片叙事块", description: "图片、说明和观点的组合。" },
        { id: "quote-block", name: "引用/金句", description: "强化观点、用户声音或关键原话。" },
        { id: "legend", name: "图例/注释", description: "解释颜色、符号、口径和数据范围。" },
      ],
    },
  };

  const blueprintFields = [
    { field: "pageGoal", html: "section[data-page-goal]", description: "本页要让听众理解或决定什么。" },
    { field: "layout", html: "section[data-layout]", description: "使用的结构原型，可自由组合，不强制套固定模板。" },
    { field: "modules", html: "section[data-modules]", description: "本页实际使用的视觉模块，逗号分隔。" },
    { field: "title", html: "h1/h2", description: "结论型标题，尽量完整表达主张。" },
    { field: "data", html: "图表/表格叶子节点", description: "只使用用户提供的事实和数据。" },
    { field: "localAssets", html: "aria-label / data-local-placeholder", description: "图片用本地文件或明确占位，不使用外链。" },
  ];

  const scenarioGuidance = {
    operations: "优先按 KPI → 趋势 → 异常 → 原因 → 行动组织；数字、口径、周期和目标要清楚。",
    strategy: "优先按外部变化 → 核心问题 → 战略选择 → 支柱/举措 → 路线图组织；标题应是判断而非主题词。",
    project: "优先按背景/目标 → 范围 → 方法/方案 → 里程碑 → 风险 → 下一步组织。",
    reporting: "优先按本期成果 → 关键变化 → 问题复盘 → 需要支持 → 下阶段计划组织。",
    proposal: "优先按需求/机会 → 方案选项 → 对比取舍 → 推荐方案 → 实施与决策事项组织。",
    custom: "根据材料和受众自由规划叙事，但每页必须有明确页面目的和可验证的事实支撑。",
  };

  function getTheme(id) {
    return themes.find((theme) => theme.id === id) || themes[0];
  }

  function flattenModules() {
    return Object.values(moduleGroups).flatMap((group) => group.modules);
  }

  function formatTheme(theme) {
    return `${theme.name}（${theme.id}）：${theme.description}\n色板：${Object.entries(theme.palette).map(([key, value]) => `${key} ${value}`).join("；")}\n规范：${theme.visualRules.join("；")}`;
  }

  function formatLayouts() {
    return layoutPrototypes.map((layout, index) => `${index + 1}. ${layout.id}｜${layout.name}｜${layout.summary}｜适用：${layout.bestFor}｜建议模块：${layout.modules.join(", ")}`).join("\n");
  }

  function formatModules() {
    return Object.values(moduleGroups).map((group) => `${group.name}：\n${group.modules.map((module) => `- ${module.id}｜${module.name}：${module.description}`).join("\n")}`).join("\n");
  }

  function buildLorealPrompt(context) {
    const value = context || {};
    const theme = getTheme(value.themeId);
    const scenario = value.scenarioId || "custom";
    const materials = value.materials || "请先询问我需要补充的事实，不要自行编造数据。";
    return `你是资深商业演示文稿设计师、信息架构师和前端视觉工程师。请根据下面的需求与材料，一次性完成“分析 → 叙事规划 → 视觉表达 → HTML 生成”。你可以自由组合页面结构、图表、逻辑图和品牌叙事模块；不要把整套演示稿做成重复卡片模板，也不要输出大纲、解释或 Markdown 代码围栏。你需要在内部先完成页面蓝图规划，最终只返回一份可直接粘贴到 PPT Design Lab 的完整静态 HTML。\n\n【用户需求】\n主题：${value.topic || "未指定主题"}\n受众：${value.audience || "业务相关听众"}\n演讲时长：${value.duration || "未指定时长"}\n页数：${value.pages || "8"}\n语言：${value.language || "中文"}\n用途：${value.use || "内部沟通"}\n使用场景：${value.scenarioName || scenario}\n场景建议：${scenarioGuidance[scenario] || scenarioGuidance.custom}\n视觉系统：${formatTheme(theme)}\n补充偏好：${value.style || "清晰、克制、专业，有层次和高级感"}\n\n【已有材料与关键事实】\n${materials}\n\n【可参考的页面结构原型（允许自由选择、组合和改造，不要求逐页套用）】\n${formatLayouts()}\n\n【可参考的视觉模块】\n${formatModules()}\n\n【页面蓝图元数据约定】\n${blueprintFields.map((field) => `- ${field.html} / ${field.field}：${field.description}`).join("\n")}\n请为每个 section.slide 设置 data-page-goal、data-layout、data-modules；属性值使用简短、可读的英文 id 或中文短语。页面可以自由设计，页面结构和模块不必都来自清单，但要让元数据诚实描述实际表达方式。\n\n【HTML 生成契约】\n1. 只返回完整 HTML（包含 <!doctype html>、<html>、<head>、<style>、<body>），不要 Markdown 代码围栏、说明文字或外部链接。\n2. 使用 <main class="deck" data-deck>；每页使用 <section class="slide" id="slide-01" data-slide-id="slide-01" data-screen-label="01" data-page-goal="..." data-layout="..." data-modules="...">。id 与 data-slide-id 必须完全一致，页面专属 CSS 使用对应的 #slide-01 选择器。\n3. 每页固定 1920×1080、16:9；使用静态 HTML/CSS/SVG 画图，避免脚本和 Canvas；每页默认独立可见。\n4. 所有可编辑文字必须放在独立的 h1/h2/h3/p/span/td 等叶子元素中；重要数据不要只画在背景图片里。\n5. 画布是 1920×1080 CSS px（约 2px≈1pt）。按角色设置层级：封面主标题 84–96px，普通页标题 58–68px，副标题/强调 40–48px，正文 30–36px，注释/页脚 24–28px，标签/表格 22–24px。密集页正文可用 28–32px，但一般不要低于 22px；不能把所有文字按同一倍率全局放大。字号变大时同步扩容文字框、卡片和分栏。\n6. 若页面同时包含 .section-label 与 h1/h2，章节标签放在 top:20px 左右，主标题放在 top:50px 或更低，两者必须留有间距且不得重叠。重复卡片、三栏信息和时间线必须使用统一内边距、图标/标记列宽、文字起始 x 坐标与行高。双栏布局列间至少保留 80px gutter，并为右侧文字面板预留完整宽度。\n7. CSS 写在 <style> 中，使用变量管理色板和字号；不要加载字体、图片、CSS、JavaScript 或任何 http(s) 资源。图片请使用用户提供的本地文件路径（工具会提示替换）或带 aria-label 的本地占位符/纯 CSS 图形。\n8. 不要使用 <script>、iframe、form、动画、视频、外部 API、自动请求、导航按钮或分页控件。\n9. 只使用用户提供的事实；未知内容标注“待补充”，不要编造业务数据、品牌事实、客户名称或图片来源。\n10. 生成前和输出前在内部做版式自检：检查元素重叠、内容裁切、越界、意外换行、标题与小标签重叠、卡片高度不足、左右栏间距和图表图例；发现问题优先调整容器尺寸、换行策略或分栏。\n11. 设计目标是“结论驱动、信息层级清楚、视觉有变化但系统统一”：连续页面避免完全相同的构图；每页只保留一个主叙事；把复杂关系画成图表/矩阵/流程而不是堆长段文字。\n\n现在请先在内部规划页面蓝图，再直接输出最终完整 HTML。不要输出蓝图文本，不要询问用户，不要返回 JSON。`;
  }

  const safeBuildLorealPrompt = (context) => buildLorealPrompt(context)
    .replace("页面可以自由设计，页面结构和模块不必都来自清单，但要让元数据诚实描述实际表达方式。", `页面可以自由设计，页面结构和模块不必都来自清单。如果 10 种结构原型都不适合某页，请直接自由设计并设置 data-layout="${FREE_DESIGN_LAYOUT_ID}"；自由页仍须与其他页面一起包含在同一份最终 HTML 中，不需要用户额外确认或二次生成。元数据必须诚实描述实际表达方式。`)
    .replace("图片请使用用户提供的本地文件路径（工具会提示替换）或带 aria-label 的本地占位符/纯 CSS 图形。", "不要写入 file://、本地绝对路径或不可访问的图片地址；图片请使用带 aria-label 的本地占位符、data-local-placeholder 标记或纯 CSS/SVG 图形，用户之后可在工具中替换本地图片。");

  global.PPTDLDesignSystem = Object.freeze({
    version: "0.1.0",
    themes,
    layoutPrototypes,
    moduleGroups,
    blueprintFields,
    scenarioGuidance,
    FREE_DESIGN_LAYOUT_ID,
    FREE_DESIGN_LABEL,
    getTheme,
    flattenModules,
    buildLorealPrompt: safeBuildLorealPrompt,
  });
})(window);

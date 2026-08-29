---
record_id: PRJ-012
record_type: project
status: active
stage: build
progress: 80
version: v0.1-mvp
created: 2026-08-23
updated: 2026-08-24
last_progress_update: 2026-08-24
updated_by: codex
revision: 5
platforms: [codex, workbuddy]
maintainer: cindy
summary: 面向公司普通用户电脑的本地优先网页版 PPT 设计工具，使用 LorealGPT 手工桥接生成结构化 HTML 幻灯片，并支持可视化编辑及 HTML、PPTX 导出。
next_action: 在公司电脑真实 Chrome 或 Edge 中打开 GitHub Pages，验证导入、编辑、HTML/PPTX 导出和全程无内容外传。
definition_of_done: v0.1 可在公司浏览器中打开，完成 LorealGPT 提示词复制与 HTML 粘贴导入、结构化幻灯片可视化编辑、自包含 HTML 和图片型 PPTX 导出，并通过无业务内容外传的网络验收。
tags: [职业发展, PPT, HTML, 浏览器工具, 本地优先, LorealGPT]
related: [CAP-003]
---

# PPT Design Lab

> 一个无需安装、在公司电脑浏览器中使用的本地优先 PPT 设计工具。

## 项目目标

把 LorealGPT、Open Design / Baoyu Design、slide-maker 与 Beauty Lab 的有效思路组合成一条简单工作流：

1. 工具根据用户填写的主题、受众、页数与风格生成 LorealGPT 提示词。
2. 用户在公司批准的 LorealGPT 中手工粘贴一次完整提示词，由 LorealGPT 同时完成内容规划、页面蓝图和自由 HTML 创作。
3. 用户把 HTML 复制回工具；工具在浏览器本地解析并显示多页 16:9 幻灯片。
4. 用户通过可视化界面修改文字、字体、字号、颜色、图片、模块位置与尺寸。
5. 工具在本地导出自包含 HTML、图片型 PPTX；基础可编辑 PPTX 作为后续增强。

## v0.1 范围

- LorealGPT 提示词向导与复制粘贴桥接。
- 三套视觉系统（咨询蓝、黑金、暖米）、十种页面结构原型和第一批 25+ 个视觉模块。
- 页面蓝图字段与自由 HTML 约定：LorealGPT 在一次生成中返回页面目标、结论、结构、模块、数据、文案和完整 HTML。
- 固定 1920×1080、16:9 的结构化静态 HTML 幻灯片规范。
- HTML 粘贴或本地文件导入；首版支持文本、Markdown、CSV、TSV、JSON、HTML、XML、PDF、DOCX、XLSX、XLS 的基础读取。
- 页面缩略图、翻页、排序、复制与删除。
- 文字、字体、字号、颜色、背景、图片、位置和尺寸的可视化修改。
- 本地项目保存与重新打开（专用 `.pptdlab` 文件；兼容旧版 `.pptdlab.json` / `.json`）。
- 自包含 HTML 导出。
- 默认图片型 PPTX 导出。
- 浏览器本地质量检查：溢出、小字号、缺图、越界、低对比度和重复页。
- GitHub Pages 在线工具与离线 ZIP 发布方案。

其中 PDF 仅提取可复制的文字层（扫描件不 OCR），Word 仅支持 `.docx`，Excel 仅提取工作表单元格值；宏、图表和复杂公式语义不在首版范围内。解析失败时可直接粘贴材料摘要。

## 暂不纳入 v0.1

- LorealGPT API 或自动登录。
- 云端账号、云端保存或多人协作。
- 任意网页的无损导入。
- 所有 HTML 元素都转换为原生可编辑 PowerPoint 对象。
- 把公司 Logo、模板、字体、业务材料或生成后的 PPT 存入项目仓库。

## 单一事实源与资产根

| 类型 | 位置 | 说明 |
|---|---|---|
| 项目资产真源 | `/Users/cindy/Projects/04-职业发展/ppt-design-lab` | 需求、架构、实现、测试、文档与发布脚本的唯一主目录 |
| 项目控制面 | YYOB `项目管理/项目/PPT Design Lab.md`（`PRJ-012`） | 状态、版本、里程碑、阻塞和项目级下一步 |
| GitHub 源码仓库 | `https://github.com/hqubare-sys/ppt-design-lab` | 仅发布工具框架、通用模板和许可证，不含公司内容 |
| GitHub Pages | `https://hqubare-sys.github.io/ppt-design-lab/` | 静态工具发布地址，不作为用户内容存储 |
| 离线发行物 | 未来 GitHub Release / `dist/` | 只包含工具代码和通用资产 |
| 公司业务内容 | 公司电脑上的本地文件与浏览器会话 | 永不复制到 Projects、YYOB、GitHub 或工具服务器 |

## 隐私与安全边界

- GitHub 和网页只承载工具框架、通用提示词模板与非敏感演示数据。
- PPT 内容默认只存在于浏览器内存；只有用户主动保存或导出时才生成本机文件，v0.1 不使用 IndexedDB 自动保存。
- 打开、粘贴、编辑、保存和导出均使用浏览器本地能力，不设置业务后端。
- 不接入统计、遥测、云端错误上报或远程数据库。
- 导入 HTML 中的脚本、事件、iframe、外链资源与网络请求默认禁用或隔离。
- 公司模板、Logo、字体和图片由用户在公司电脑本地导入，不进入公开仓库。
- LorealGPT 是由用户手工操作的公司批准服务；工具不读取登录状态、不调用私有接口。
- 发布验收必须证明初始应用加载完成后，处理公司内容的全过程没有对外传输。

## 方法来源

- Beauty Lab：浏览器本地 HTML 导入、隔离预览与可视化编辑。
- Open Design / Baoyu Design：结构化静态幻灯片、固定画布、设计变量与 HTML/PPTX 双导出。
- `addsumtech/slides_maker`：需求访谈、内容规划、设计规划、质量检查和尽量可编辑的 PPTX；仅借鉴方法，浏览器版本不采用其 Python、LibreOffice 等本地运行依赖。

引用或改造第三方 MIT 代码时，必须保留相应许可证与版权声明。

## 快速审计

在项目根目录运行：

```bash
python3 scripts/audit_project.py
```

## GitHub Pages 使用

公开仓库 `hqubare-sys/ppt-design-lab` 的 `main` 分支会通过 `.github/workflows/pages.yml` 自动发布；也可以在 Actions 中手动运行同一个 workflow。发布 artifact 只取 `app/` 目录，因此 Pages 根路径会直接打开 `app/index.html`，无需额外安装 Node.js、Python、插件或桌面应用。

仓库地址为 `https://github.com/hqubare-sys/ppt-design-lab`，Pages 地址为 `https://hqubare-sys.github.io/ppt-design-lab/`。不要把公司业务文件加入仓库。

首次使用 Pages 版时，打开页面后直接在浏览器中读取本地材料、生成一份完整提示词，手工复制到 LorealGPT，再把同一次回答中的 HTML 粘贴回工具。导入、编辑、项目保存和 HTML/PPTX 导出均在当前浏览器本地完成；Pages 只提供静态框架，不提供上传接口、业务后端、统计或云端保存。

## 本地试用 v0.1

开发验证时，在项目根目录启动任意静态文件服务器并打开 `app/`。正式给公司电脑使用时，将发布同一套静态文件到 GitHub Pages，并同时提供离线 ZIP；公司电脑不需要安装开发工具。

第一版已经包含：三页首页内置演示、LorealGPT 一次完整提示词、三套视觉系统与页面蓝图说明、首版本地材料读取、HTML 安全导入、文字与布局编辑、页面管理、撤销恢复、本地 `.pptdlab` 项目文件（兼容旧扩展）、自包含 HTML 和图片型 PPTX 导出。

图片型 PPTX 使用与编辑预览一致的本地隔离 iframe 逐页渲染，确保 `:root` CSS 变量、主题色、背景、边框、圆角和阴影在截图中继续生效；渲染 iframe 禁止脚本且不发起网络请求。

导出器会保留页面实际背景色，并只对浏览器中本来就是单行的叶子文字增加导出稳定样式，减少中文尾字在图片转换时意外换行。编辑器只允许叶子文字进入双击编辑；结构容器保持完整，轻微点击不触发移动，嵌套模块拖动使用其父容器坐标。

导入器兼容常见的“HTML 自带翻页播放器”输出：它会按 deck 结构识别真正的 section 页面，忽略导航圆点和翻页控件，并在移除原脚本后让每页继续能够独立编辑和导出。独立 HTML 导出会把项目 CSS 与播放器 CSS 放在带标记的两个 style 块中；再次打开时只导入项目 CSS。对旧版本中两类 CSS 混在同一块的文件，导入器也会只清理已知的播放器选择器、深色 body 与居中规则，不会删除普通项目样式。若 LorealGPT 只提供 `data-slide-id`、但页面专属 CSS 使用 `#slide-01`，导入器会自动补齐安全且唯一的 DOM `id`，恢复页面字号、定位和卡片布局；直接位于页面根下的 `.section-label` 会被放到标题上方的安全区，避免两者共享相同 `top` 坐标而重叠。导入完成后状态栏会明确显示识别页数。

1920×1080 画布的字号使用 CSS px，约 `2px ≈ 1pt`。新版 LorealGPT 提示词按信息密度和文字角色给出标准档：封面主标题 84–96px、普通页标题 58–68px、副标题/强调 40–48px、正文 30–36px、注释/页脚 24–28px、标签/表格 22–24px；稀疏页可适度放大，密集页正文可降到 28–32px，但一般不低于 22px。不要按同一倍率全局放大所有文字：字号变化必须同步扩容文本框、调整卡片高度、换行或重新分栏。生成前和输出前都应检查重叠、裁切、越界、意外换行和对齐；双栏至少保留 80px 间距，圆形主视觉或投票框必须固定在所属列内，不能居中挤压右栏。重复卡片和时间线继续统一内边距、图标列宽、文字起始坐标与基线。

已知限制见 [`docs/known-issues.md`](docs/known-issues.md)。其中底部多段 inline 元数据在极窄剩余宽度下可能在图片型 PPTX 导出时整体下移一行，属于低优先级视觉细节，不阻塞公司电脑可用性测试。

## 当前状态

v0.2 预览版已完成本机开发与在线验收；公开仓库与 GitHub Pages 已发布并返回 HTTP 200，仓库树已核对为 38 个工具框架、品牌资源、本地依赖、文档、测试和许可证文件。公司电脑真实环境与网络零内容外传验收尚未完成，因此当前仍标记为测试版本。

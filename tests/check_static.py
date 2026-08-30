#!/usr/bin/env python3
"""Dependency-free smoke checks for the static PPT Design Lab prototype.

Run from the repository root:
    python3 tests/check_static.py

This intentionally checks source boundaries only. Browser interaction checks
belong to the acceptance workflow and should be run with a local HTTP server.
"""

from pathlib import Path
import re
import sys


ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "app"
REQUIRED = [
    "index.html",
    "styles.css",
    "app.js",
    "sanitizer.js",
    "project-io.js",
    "exporter.js",
]


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


def main() -> int:
    gitignore = ROOT / ".gitignore"
    if not gitignore.is_file():
        fail("missing .gitignore privacy boundary")
    ignored_paths = gitignore.read_text(encoding="utf-8")
    for marker in [".DS_Store", "*.pptx", "*.pptdlab", "*.pptdlab.json", "dist/", "exports/", "local-data/", "company-content/"]:
        if marker not in ignored_paths:
            fail(f".gitignore is missing privacy boundary: {marker}")

    for name in REQUIRED:
        path = APP / name
        if not path.is_file() or not path.read_text(encoding="utf-8").strip():
            fail(f"missing or empty app/{name}")
    if 'src="file-parser.js"' not in (APP / "index.html").read_text(encoding="utf-8"):
        fail("index.html does not load the local material parser")

    index = (APP / "index.html").read_text(encoding="utf-8")
    scripts = re.findall(r'<script[^>]+src="([^"]+)"', index)
    if any(url.startswith(("http:", "https:", "//")) for url in scripts):
        fail("index.html references a remote script")
    if "sandbox=\"allow-same-origin\"" not in index:
        fail("editor iframe is missing the no-script sandbox")
    if "仅本机处理" not in index:
        fail("privacy status is not visible in the UI")

    source = "\n".join((APP / name).read_text(encoding="utf-8") for name in REQUIRED)
    if re.search(r"\b(?:fetch|XMLHttpRequest|sendBeacon)\s*\(", source):
        fail("application source contains an outbound request path")
    if re.search(r"\b(?:localStorage|sessionStorage|indexedDB)\b", source):
        fail("application source uses browser persistence")
    if "data:image/svg+xml" not in (APP / "sanitizer.js").read_text(encoding="utf-8"):
        fail("sanitizer has no local image placeholder")
    if "writeFile" not in (APP / "exporter.js").read_text(encoding="utf-8"):
        fail("PPTX exporter has no explicit writeFile call")
    exporter = (APP / "exporter.js").read_text(encoding="utf-8")
    player_css_source = exporter.split("const playerCSS = `", 1)[1].split("`;", 1)[0]
    if re.search(r"html,body\{[^}]*(?:background|color|font-family)\s*:", player_css_source):
        fail("standalone player html/body still overrides project theme inheritance")
    if re.search(r"\.pptdlab-stage\{[^}]*(?:background|color|font-family)\s*:", player_css_source):
        fail("standalone player stage still overrides project theme inheritance")
    if "attachShadow" in exporter:
        fail("PPTX renderer still uses Shadow DOM, which breaks imported :root variables")
    for marker in ['document.createElement("iframe")', 'sandbox", "allow-same-origin"', "host.srcdoc", "waitForRenderFrames", "renderDocument.fonts?.ready", "exportBackgroundColor", "stabilizeSingleLineText", 'whiteSpace = "nowrap"', 'wordBreak = "keep-all"']:
        if marker not in exporter:
            fail(f"PPTX iframe renderer is missing: {marker}")
    if 'backgroundColor: "#ffffff"' in exporter:
        fail("PPTX renderer still forces a white background")
    render_harness = (ROOT / "tests" / "export-render-harness.html").read_text(encoding="utf-8")
    if ":root{--page:" not in render_harness or "renderSlideToPng" not in render_harness:
        fail("PPTX render harness does not exercise :root CSS variables")
    if "负责人：待确认" not in render_harness:
        fail("PPTX render harness must use neutral generic fixture copy")
    editor_harness = (ROOT / "tests" / "editor-interaction-harness.html").read_text(encoding="utf-8")
    editor_fixture = (ROOT / "tests" / "fixture-editor-interactions.html").read_text(encoding="utf-8")
    for marker in ["dblclick", "微小移动", "offsetParent", "nested-module", "浅米色页面背景"]:
        if marker not in editor_harness:
            fail(f"editor interaction harness is missing: {marker}")
    if ":root { --page: #fff4df" not in editor_fixture or 'data-test="nested-module"' not in editor_fixture:
        fail("editor interaction fixture is not generic or lacks nested module")
    if "负责人：待确认" not in editor_fixture:
        fail("editor interaction fixture must use neutral generic fixture copy")

    unsafe = (ROOT / "tests" / "fixture-unsafe.html").read_text(encoding="utf-8")
    if "<script>" not in unsafe or "https://example.invalid" not in unsafe:
        fail("unsafe fixture does not exercise blocked content")
    sanitizer = (APP / "sanitizer.js").read_text(encoding="utf-8")
    if 'name === "style"' not in sanitizer or "safeCss(value" not in sanitizer:
        fail("sanitizer does not clean inline style attributes")
    player_fixture = (ROOT / "tests" / "fixture-player-deck.html").read_text(encoding="utf-8")
    if player_fixture.count('<section class="slide"') != 3 or player_fixture.count('class="nav-dot"') != 3:
        fail("player-deck fixture must contain three pages and three navigation dots")
    roundtrip_harness = (ROOT / "tests" / "roundtrip-harness.html").read_text(encoding="utf-8")
    for marker in ["buildStandaloneHTML", "data-pptdlab-project-style", "data-pptdlab-player-style", "旧版导出 HTML", "projectFilename"]:
        if marker not in roundtrip_harness:
            fail(f"round-trip harness is missing: {marker}")
    if 'querySelectorAll(".slide, [data-slide]' in sanitizer:
        fail("sanitizer still treats generic data-slide controls as pages")
    if "findSlideNodes(parsed)" not in sanitizer or "data-pptdlab-slide-root" not in sanitizer:
        fail("sanitizer lacks structural page detection or visibility normalization")
    if "SAFE_DOM_ID_RE" not in sanitizer or 'slide.setAttribute("id", domId)' not in sanitizer:
        fail("sanitizer does not synchronize safe page DOM ids for page-specific CSS")
    if '>.section-label{top:20px!important;}' not in sanitizer:
        fail("sanitizer does not keep direct section labels above the title safe area")
    if "data-pptdlab-player-style" not in exporter or "data-pptdlab-project-style" not in exporter:
        fail("standalone HTML does not separate project and player styles")
    if "stripLegacyPlayerCss" not in sanitizer or ".pptdlab-stage" not in sanitizer:
        fail("sanitizer lacks standalone player CSS isolation or deck detection")
    project_io = (APP / "project-io.js").read_text(encoding="utf-8")
    if ".pptdlab`" not in project_io or "application/x-ppt-design-lab+json" not in project_io:
        fail("project download does not use the dedicated .pptdlab format")
    demo_markers = [
        "slide-03",
        "intro-slide",
        "workflow-slide",
        "system-slide",
        "咨询蓝",
        "黑金",
        "暖米",
        "10 种页面结构原型",
        "25+ 视觉模块",
        "页面蓝图",
        "free_html",
        "不上传文件",
    ]
    for marker in demo_markers:
        if marker not in project_io:
            fail(f"built-in three-page demo is missing: {marker}")
    if project_io.count('<section class="slide ') < 3:
        fail("built-in demo must contain three editable slides")
    for removed_first_page_copy in ["适合这些工作场景", "第一次使用，记住四件事"]:
        if removed_first_page_copy in project_io:
            fail(f"built-in first page must preserve the original simple layout: {removed_first_page_copy}")
    if ".pptdlab,.pptdlab.json" not in index:
        fail("project input does not accept new and legacy extensions")
    app_source = (APP / "app.js").read_text(encoding="utf-8")
    if "已识别 ${state.project.slides.length} 页" not in app_source:
        fail("import status does not report recognized page count")
    if 'root.setAttribute("id", clone.id)' not in app_source:
        fail("duplicated pages do not receive a fresh DOM id")
    design_system = (APP / "design-system.js").read_text(encoding="utf-8")
    prompt_source = f"{app_source}\n{design_system}"
    for marker in ["HTML_CODE_FENCE_OPEN", "一个且仅一个 Markdown HTML 代码块", "Copy code 按钮", "代码块外不得输出", "JavaScript 完全禁用时仍能静态呈现"]:
        if marker not in design_system:
            fail(f"LorealGPT prompt is missing copyable HTML delivery contract: {marker}")
    if "Copy code 按钮" not in index:
        fail("LorealGPT workflow UI does not explain the copyable html code block")
    if 'id="slide-01" data-slide-id="slide-01"' not in prompt_source:
        fail("LorealGPT prompt does not require matching DOM and project slide ids")
    for marker in ["封面主标题 84–96px", "普通页标题 58–68px", "正文 30–36px", "一般不要低于 22px", "不能把所有文字按同一倍率全局放大", "双栏布局列间至少保留 80px gutter", "版式自检"]:
        if marker not in prompt_source:
            fail(f"LorealGPT prompt is missing typography/layout guidance: {marker}")
    for marker in ["一次性完成", "页面蓝图", "咨询蓝", "黑金", "暖米", "不要把整套演示稿做成重复卡片模板"]:
        if marker not in prompt_source:
            fail(f"LorealGPT prompt is missing free-design guidance: {marker}")
    if "自由设计" not in prompt_source or "完整静态 HTML" not in prompt_source:
        fail("LorealGPT prompt is missing free HTML guidance")
    parser_path = APP / "file-parser.js"
    if not parser_path.is_file() or not parser_path.read_text(encoding="utf-8").strip():
        fail("missing local material parser entry")
    parser_source = parser_path.read_text(encoding="utf-8")
    for marker in ["FileReader", "arrayBuffer", ".pdf", ".docx", ".xlsx", ".xls", "本地", "不上传"]:
        if marker not in parser_source and marker not in index:
            fail(f"local material parser is missing: {marker}")
    for marker in [".pdf", ".docx", ".xlsx", ".xls", "文件只在当前浏览器本地读取，不上传"]:
        if marker not in index:
            fail(f"material input/privacy boundary is missing: {marker}")
    for marker in ["MOVE_THRESHOLD", "offsetParentPosition", "Math.hypot(dx, dy)", 'event.detail > 1', 'target.children.length > 0 || !target.textContent.trim()']:
        if marker not in app_source:
            fail(f"editor interaction guard is missing: {marker}")
    print("PASS: static source, privacy, sandbox, and fixture checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

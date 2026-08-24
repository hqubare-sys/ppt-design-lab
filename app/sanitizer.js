(function (global) {
  "use strict";

  const PLACEHOLDER_IMAGE =
    "data:image/svg+xml;charset=utf-8," +
    encodeURIComponent(
      '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540">' +
        '<rect width="960" height="540" fill="#e7eaf0"/>' +
        '<path d="M0 430 220 260l150 120 130-160 260 210H0Z" fill="#c6cbd7"/>' +
        '<circle cx="690" cy="180" r="70" fill="#b0b6c5"/>' +
        '<text x="480" y="500" text-anchor="middle" font-family="Arial,sans-serif" font-size="26" fill="#596273">本地图片占位符</text>' +
      '</svg>'
    );

  const BLOCKED_TAGS = ["script", "iframe", "object", "embed", "base", "form"];
  const REMOTE_RE = /^(?:https?:|\/\/|data:(?!image\/|font\/|application\/font-))/i;
  const JAVASCRIPT_RE = /^\s*javascript:/i;
  const SAFE_DOM_ID_RE = /^[A-Za-z][A-Za-z0-9_-]*$/;
  const EDITOR_SLIDE_CSS =
    '[data-pptdlab-slide-root="true"]{display:block!important;visibility:visible!important;opacity:1!important;}' +
    '[data-pptdlab-slide-root="true"]>.section-label{top:20px!important;}';
  const PAGE_CLASS_NAMES = new Set(["slide", "slide-page", "ppt-slide", "page"]);
  const CONTROL_CLASS_NAMES = new Set([
    "nav-dot",
    "nav-dots",
    "nav-controls",
    "slide-nav",
    "slide-navigation",
    "pagination",
    "pager",
    "controls",
  ]);

  function warningFactory() {
    const list = [];
    const seen = new Set();
    return {
      add(message) {
        if (!seen.has(message)) {
          seen.add(message);
          list.push(message);
        }
      },
      list,
    };
  }

  function safeCss(cssText, warnings) {
    let css = String(cssText || "");
    const withoutTags = css.replace(/<\/?[a-z][^>]*>/gi, "");
    if (withoutTags !== css) warnings.add("已移除 CSS 文本中的 HTML 标签。");
    css = withoutTags;
    const before = css;
    css = css.replace(/@import[^;]*(?:;|$)/gi, "");
    if (css !== before) warnings.add("已移除 CSS @import 规则。");
    css = css.replace(/url\(\s*(['"]?)([^'"\)]*)\1\s*\)/gi, (match, quote, url) => {
      const value = String(url || "").trim();
      if (REMOTE_RE.test(value) || JAVASCRIPT_RE.test(value)) {
        warnings.add("已移除 CSS 中的外部或活动资源地址。");
        return "none";
      }
      return `url(${quote}${value}${quote})`;
    });
    return css;
  }

  function hasExternalResource(value) {
    return REMOTE_RE.test(String(value || "")) || JAVASCRIPT_RE.test(String(value || ""));
  }

  function scrubAttributes(element, warnings) {
    Array.from(element.attributes || []).forEach((attribute) => {
      const name = attribute.name.toLowerCase();
      const value = attribute.value || "";
      if (name.startsWith("on")) {
        element.removeAttribute(attribute.name);
        warnings.add(`已移除事件属性 ${attribute.name}。`);
        return;
      }
      if (name === "style") {
        const cleaned = safeCss(value, warnings);
        if (cleaned) element.setAttribute("style", cleaned);
        else element.removeAttribute("style");
        return;
      }
      if (["src", "href", "xlink:href", "action", "formaction", "poster"].includes(name)) {
        if (JAVASCRIPT_RE.test(value)) {
          element.removeAttribute(attribute.name);
          warnings.add("已移除 javascript: 地址。");
          return;
        }
        if (hasExternalResource(value)) {
          if (element.tagName.toLowerCase() === "img" && name === "src") {
            element.setAttribute("src", PLACEHOLDER_IMAGE);
            element.setAttribute("data-local-placeholder", "true");
            warnings.add("外部图片未下载，已替换为本地占位符。");
          } else {
            element.removeAttribute(attribute.name);
            warnings.add("已移除外部链接或资源地址。");
          }
        }
        return;
      }
      if (name === "srcset") {
        element.removeAttribute(attribute.name);
        warnings.add("已移除图片 srcset 外部候选地址。");
      }
    });
  }

  function makeUniqueId(candidate, index, usedIds) {
    const fallback = `slide-${String(index + 1).padStart(2, "0")}`;
    const base = SAFE_DOM_ID_RE.test(String(candidate || "").trim()) ? String(candidate).trim() : fallback;
    let value = base;
    let suffix = 2;
    while (usedIds.has(value)) {
      value = `${base}-${suffix}`;
      suffix += 1;
    }
    usedIds.add(value);
    return value;
  }

  function ensureSlideIdentity(slide, index, usedSlideIds, usedDomIds) {
    const slideId = makeUniqueId(slide.getAttribute("data-slide-id"), index, usedSlideIds);
    slide.setAttribute("data-slide-id", slideId);
    if (!slide.getAttribute("data-screen-label")) {
      slide.setAttribute("data-screen-label", String(index + 1).padStart(2, "0"));
    }
    if (!slide.classList.contains("slide")) slide.classList.add("slide");
    const existingDomId = String(slide.getAttribute("id") || "").trim();
    const domId = makeUniqueId(existingDomId || slideId, index, usedDomIds);
    slide.setAttribute("id", domId);
    slide.setAttribute("data-pptdlab-slide-root", "true");
    slide.removeAttribute("data-active");
    slide.removeAttribute("hidden");
    slide.removeAttribute("aria-hidden");
    if (slide.style) {
      slide.style.removeProperty("display");
      slide.style.removeProperty("visibility");
      slide.style.removeProperty("opacity");
    }
  }

  function hasClassName(element, names) {
    return Array.from(element.classList || []).some((name) => names.has(String(name).toLowerCase()));
  }

  function isNavigationControl(element) {
    if (!element || element.nodeType !== 1) return true;
    const tagName = element.tagName.toLowerCase();
    if (tagName === "nav" || tagName === "button") return true;
    if (hasClassName(element, CONTROL_CLASS_NAMES)) return true;
    if (element.hasAttribute("data-slide-nav") || element.hasAttribute("data-slide-control")) return true;
    // A data-slide marker on a span is the common navigation-dot shape, not
    // a page root. Generic data-slide attributes are intentionally not used
    // as a standalone page selector.
    if (tagName === "span" && element.hasAttribute("data-slide")) return true;
    return false;
  }

  function hasPageMarker(element) {
    if (!element || element.nodeType !== 1 || isNavigationControl(element)) return false;
    if (hasClassName(element, PAGE_CLASS_NAMES)) return true;
    return ["data-slide-id", "data-page", "data-page-id"].some((attribute) => element.hasAttribute(attribute));
  }

  function isUnmarkedSection(element) {
    if (!element || element.nodeType !== 1 || isNavigationControl(element)) return false;
    const tagName = element.tagName.toLowerCase();
    return tagName === "section" || tagName === "article";
  }

  function directChildren(element) {
    return Array.from(element?.children || []);
  }

  function topLevelCandidates(nodes) {
    const unique = Array.from(new Set(nodes));
    return unique.filter((node) => !unique.some((other) => other !== node && other.contains(node)));
  }

  function findDeckContainers(parsed) {
    return topLevelCandidates(
      Array.from(parsed.querySelectorAll("[data-deck], .deck, deck-stage"))
        .filter((node) => node.nodeType === 1)
    );
  }

  function findSlideNodes(parsed) {
    const decks = findDeckContainers(parsed);

    // Prefer direct page children of a declared deck. This is deliberately
    // structural: a nav-dot span with data-slide is not eligible here.
    for (const deck of decks) {
      const children = directChildren(deck);
      const marked = children.filter(hasPageMarker);
      if (marked.length) return topLevelCandidates(marked);
      const sections = children.filter(isUnmarkedSection);
      if (sections.length) return topLevelCandidates(sections);
    }

    // Support common wrappers inside a deck, while still requiring an
    // explicit page marker and dropping nested page-like elements.
    const deckDescendants = decks.flatMap((deck) =>
      Array.from(deck.querySelectorAll("section, article, div"))
        .filter(hasPageMarker)
    );
    if (deckDescendants.length) return topLevelCandidates(deckDescendants);

    // No declared deck: accept explicit page roots at document level or in a
    // simple wrapper. This keeps legacy single-file HTML imports compatible.
    const explicitPages = Array.from(parsed.querySelectorAll("section, article, div")).filter(hasPageMarker);
    if (explicitPages.length) return topLevelCandidates(explicitPages);

    // A deck-less document made of top-level sections is a reasonable page
    // fallback; controls and ordinary spans are still excluded.
    const bodySections = directChildren(parsed.body).filter(isUnmarkedSection);
    return topLevelCandidates(bodySections);
  }

  function stripNavigationFromFallback(root, warnings) {
    const controls = Array.from(root.querySelectorAll("nav, button, [data-slide-nav], [data-slide-control]"))
      .concat(Array.from(root.querySelectorAll("*" )).filter((node) => isNavigationControl(node)));
    const unique = Array.from(new Set(controls));
    unique.forEach((node) => {
      if (node !== root && node.parentNode) {
        node.remove();
        warnings.add("已忽略导入 HTML 中的翻页导航控件。");
      }
    });
  }

  function assignEditorIds(root) {
    let serial = 1;
    const nodes = root.querySelectorAll("*");
    nodes.forEach((node) => {
      if (node.id === "__editor_selection" || node.hasAttribute("data-editor-selection")) return;
      if (!node.getAttribute("data-editor-id")) {
        node.setAttribute("data-editor-id", `node-${serial}`);
      }
      serial += 1;
    });
  }

  function removeEditorArtifacts(root) {
    root.querySelectorAll("[data-editor-selection], #__editor_selection").forEach((node) => node.remove());
    root.querySelectorAll("[contenteditable], [spellcheck]").forEach((node) => {
      node.removeAttribute("contenteditable");
      node.removeAttribute("spellcheck");
    });
  }

  function sanitizeDocument(input) {
    const warnings = warningFactory();
    const source = String(input || "").trim();
    if (!source) throw new Error("没有可导入的 HTML 内容。");
    const parser = new DOMParser();
    const parsed = parser.parseFromString(source, "text/html");

    BLOCKED_TAGS.forEach((tag) => {
      parsed.querySelectorAll(tag).forEach((node) => {
        node.remove();
        warnings.add(`已移除 ${tag} 活动元素。`);
      });
    });
    parsed.querySelectorAll('meta[http-equiv="refresh" i]').forEach((node) => {
      node.remove();
      warnings.add("已移除页面自动跳转 meta。");
    });
    parsed.querySelectorAll("link").forEach((node) => {
      const rel = (node.getAttribute("rel") || "").toLowerCase();
      const href = node.getAttribute("href") || "";
      if (rel.includes("stylesheet") || rel.includes("preload") || rel.includes("modulepreload") || hasExternalResource(href)) {
        node.remove();
        warnings.add("已移除外部 stylesheet、预加载或字体链接。");
      }
    });
    parsed.querySelectorAll("style").forEach((node) => {
      node.textContent = safeCss(node.textContent, warnings);
    });
    parsed.querySelectorAll("*").forEach((node) => scrubAttributes(node, warnings));

    let slideNodes = findSlideNodes(parsed);
    if (!slideNodes.length) {
      const fallback = parsed.createElement("section");
      fallback.className = "slide";
      fallback.innerHTML = parsed.body ? parsed.body.innerHTML : "";
      parsed.body.innerHTML = "";
      parsed.body.appendChild(fallback);
      stripNavigationFromFallback(fallback, warnings);
      slideNodes = [fallback];
      warnings.add("未发现标准 .slide 分页，已将正文作为单页导入。");
    }

    const usedSlideIds = new Set();
    const usedDomIds = new Set();
    const slides = slideNodes.map((node, index) => {
      ensureSlideIdentity(node, index, usedSlideIds, usedDomIds);
      removeEditorArtifacts(node);
      assignEditorIds(node);
      return {
        id: node.getAttribute("data-slide-id"),
        label: node.getAttribute("data-screen-label") || String(index + 1).padStart(2, "0"),
        html: node.outerHTML,
      };
    });

    const title = parsed.querySelector("title")?.textContent?.trim() || "未命名演示文稿";
    const headParts = [];
    parsed.querySelectorAll("meta, title, style").forEach((node) => {
      if (node.tagName.toLowerCase() === "meta") {
        const httpEquiv = node.getAttribute("http-equiv");
        if (!httpEquiv) headParts.push(node.outerHTML);
      } else if (node.tagName.toLowerCase() === "title") {
        headParts.push(`<title>${escapeText(node.textContent || title)}</title>`);
      } else {
        headParts.push(`<style>${safeCss(node.textContent, warnings)}</style>`);
      }
    });
    const css = `${Array.from(parsed.querySelectorAll("style"))
      .map((style) => safeCss(style.textContent || "", warnings))
      .join("\n")}\n${EDITOR_SLIDE_CSS}`;
    return {
      title,
      headHtml: headParts.join("\n"),
      css,
      slides,
      warnings: warnings.list,
    };
  }

  function escapeText(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function sanitizeSlideHtml(html, index, warnings) {
    const result = sanitizeDocument(`<main class="deck">${html}</main>`);
    const slide = result.slides[0];
    if (!slide) throw new Error("项目中的页面内容无效。");
    slide.id = slide.id || `slide-${String(index + 1).padStart(2, "0")}`;
    slide.label = slide.label || String(index + 1).padStart(2, "0");
    (result.warnings || []).forEach((warning) => warnings.add(warning));
    return slide;
  }

  global.PPTDLSanitizer = {
    PLACEHOLDER_IMAGE,
    sanitizeDocument,
    sanitizeSlideHtml,
    safeCss,
    assignEditorIds,
    removeEditorArtifacts,
  };
})(window);

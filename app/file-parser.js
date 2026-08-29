(function (global) {
  "use strict";

  // This parser deliberately works on user-selected File objects only. It has
  // no network client and does not persist the source files. Office/PDF
  // libraries are loaded from the local app/vendor directory on first use.
  const MB = 1024 * 1024;
  const CONFIG = Object.freeze({
    version: "0.1.0",
    maxTextFileBytes: 2 * MB,
    maxBinaryFileBytes: 20 * MB,
    maxSingleOutputChars: 60000,
    maxTotalOutputChars: 120000,
    maxPdfPages: 80,
    maxWorkbookSheets: 20,
    maxWorksheetRows: 5000,
    maxWorksheetColumns: 80,
  });

  const TEXT_EXTENSIONS = Object.freeze(["txt", "md", "markdown", "csv", "tsv", "json", "html", "htm", "xml"]);
  const SUPPORTED_EXTENSIONS = Object.freeze(TEXT_EXTENSIONS.concat(["pdf", "docx", "xlsx", "xls"]));
  const TYPE_BY_EXTENSION = Object.freeze({
    txt: "txt",
    md: "md",
    markdown: "markdown",
    csv: "csv",
    tsv: "tsv",
    json: "json",
    html: "html",
    htm: "html",
    xml: "xml",
    pdf: "pdf",
    docx: "docx",
    xlsx: "xlsx",
    xls: "xls",
  });
  const TYPE_BY_MIME = Object.freeze({
    "text/plain": "txt",
    "text/markdown": "md",
    "text/csv": "csv",
    "text/tab-separated-values": "tsv",
    "application/json": "json",
    "text/html": "html",
    "application/xhtml+xml": "html",
    "application/xml": "xml",
    "text/xml": "xml",
    "application/pdf": "pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/vnd.ms-excel": "xls",
  });
  const TYPE_INFO = Object.freeze({
    text: Object.freeze({
      label: "文字与结构化文本",
      extensions: TEXT_EXTENSIONS,
      maxBytes: CONFIG.maxTextFileBytes,
    }),
    pdf: Object.freeze({
      label: "PDF（文字层）",
      extensions: Object.freeze(["pdf"]),
      maxBytes: CONFIG.maxBinaryFileBytes,
    }),
    docx: Object.freeze({
      label: "Word（DOCX）",
      extensions: Object.freeze(["docx"]),
      maxBytes: CONFIG.maxBinaryFileBytes,
    }),
    spreadsheet: Object.freeze({
      label: "Excel",
      extensions: Object.freeze(["xlsx", "xls"]),
      maxBytes: CONFIG.maxBinaryFileBytes,
    }),
  });

  const dependencyPromises = new Map();

  function fileName(file) {
    return String(file?.name || "未命名文件");
  }

  function extensionOf(name) {
    const match = String(name || "").toLowerCase().match(/\.([a-z0-9]+)$/);
    return match ? match[1] : "";
  }

  function detectType(file) {
    const extension = extensionOf(fileName(file));
    if (TYPE_BY_EXTENSION[extension]) return TYPE_BY_EXTENSION[extension];
    return TYPE_BY_MIME[String(file?.type || "").toLowerCase()] || "unknown";
  }

  function normalizeFiles(input) {
    if (!input) return [];
    if (Array.isArray(input)) return input.filter(Boolean);
    if (typeof input !== "string" && typeof input[Symbol.iterator] === "function") {
      return Array.from(input).filter(Boolean);
    }
    return [input];
  }

  function makeResult(file, type, status, summary, text, warnings) {
    return {
      name: fileName(file),
      type: type || "unknown",
      status: status || "error",
      summary: String(summary || ""),
      text: String(text || ""),
      warnings: Array.isArray(warnings) ? warnings.filter(Boolean).map(String) : [],
    };
  }

  function normalizeText(value) {
    return String(value || "")
      .replace(/\r\n?/g, "\n")
      .replace(/\u0000/g, "")
      .replace(/\u00a0/g, " ")
      .split("\n")
      .map((line) => line.trimEnd())
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function decodeTextBuffer(buffer) {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer || 0);
    let encoding = "utf-8";
    let offset = 0;
    if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) offset = 3;
    else if (bytes[0] === 0xff && bytes[1] === 0xfe) { encoding = "utf-16le"; offset = 2; }
    else if (bytes[0] === 0xfe && bytes[1] === 0xff) { encoding = "utf-16be"; offset = 2; }
    if (typeof global.TextDecoder === "function") {
      try { return new global.TextDecoder(encoding).decode(bytes.subarray(offset)); } catch (error) { /* fallback below */ }
    }
    let binary = "";
    for (let index = offset; index < bytes.length; index += 1) binary += String.fromCharCode(bytes[index]);
    try { return decodeURIComponent(escape(binary)); } catch (error) { return binary; }
  }

  function readArrayBuffer(file, maxBytes) {
    if (!file) return Promise.reject(new Error("没有选择文件。"));
    if (Number.isFinite(file.size) && file.size > maxBytes) {
      return Promise.reject(new Error(`文件超过 ${Math.round(maxBytes / MB)} MB 限制。`));
    }
    if (typeof file.arrayBuffer === "function") return file.arrayBuffer();
    if (typeof global.FileReader !== "function") return Promise.reject(new Error("当前浏览器不支持本地文件读取。"));
    return new Promise((resolve, reject) => {
      const reader = new global.FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("读取文件失败。"));
      reader.readAsArrayBuffer(file);
    });
  }

  function clipText(text, maxChars) {
    const value = String(text || "");
    if (value.length <= maxChars) return { text: value, clipped: false };
    return { text: value.slice(0, Math.max(0, maxChars)), clipped: true };
  }

  function localUrl(relativePath) {
    if (typeof document === "undefined" || !document.baseURI) return relativePath;
    return new global.URL(relativePath, document.baseURI).href;
  }

  function loadLocalScript(key, relativePath, globalName) {
    if (global[globalName]) return Promise.resolve(global[globalName]);
    if (dependencyPromises.has(key)) return dependencyPromises.get(key);
    if (typeof document === "undefined") return Promise.reject(new Error(`${key} 依赖只能在浏览器中加载。`));
    const promise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.async = true;
      script.dataset.pptdlLocalDependency = key;
      script.src = localUrl(relativePath);
      script.onload = () => {
        if (global[globalName]) resolve(global[globalName]);
        else reject(new Error(`本地 ${key} 文件已加载，但未找到 ${globalName}。`));
      };
      script.onerror = () => reject(new Error(`无法加载本地 ${key} 文件，请确认 app/vendor 文件完整。`));
      (document.head || document.documentElement).appendChild(script);
    });
    dependencyPromises.set(key, promise);
    promise.catch(() => dependencyPromises.delete(key));
    return promise;
  }

  function toPlainMarkupText(markup) {
    const source = String(markup || "")
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/<script\b[^>]*>[\s\S]*?<\/script\s*>/gi, "")
      .replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi, "")
      .replace(/<noscript\b[^>]*>[\s\S]*?<\/noscript\s*>/gi, "");
    if (typeof global.DOMParser !== "function") {
      return normalizeText(source.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, " "));
    }
    const parsed = new global.DOMParser().parseFromString(source, "text/html");
    const root = parsed.body || parsed.documentElement;
    const ignored = new Set(["SCRIPT", "STYLE", "NOSCRIPT", "TEMPLATE", "IFRAME", "OBJECT", "EMBED"]);
    const blocks = new Set(["ADDRESS", "ARTICLE", "ASIDE", "BLOCKQUOTE", "DIV", "DL", "FIELDSET", "FIGCAPTION", "FIGURE", "FOOTER", "FORM", "H1", "H2", "H3", "H4", "H5", "H6", "HEADER", "HR", "LI", "MAIN", "NAV", "OL", "P", "PRE", "SECTION", "TABLE", "TBODY", "TD", "TFOOT", "TH", "THEAD", "TR", "UL"]);
    const output = [];
    function append(value) {
      if (!value) return;
      output.push(value);
    }
    function walk(node) {
      if (node.nodeType === 3) { append(node.nodeValue || ""); return; }
      if (node.nodeType !== 1 || ignored.has(node.tagName)) return;
      if (node.tagName === "BR" || node.tagName === "HR") append("\n");
      const isBlock = blocks.has(node.tagName);
      if (isBlock) append("\n");
      Array.from(node.childNodes || []).forEach(walk);
      if (isBlock) append("\n");
    }
    walk(root);
    return normalizeText(output.join(""));
  }

  async function parseTextFile(file, type, maxChars) {
    const buffer = await readArrayBuffer(file, CONFIG.maxTextFileBytes);
    let text = decodeTextBuffer(buffer);
    text = type === "html" ? toPlainMarkupText(text) : normalizeText(text);
    const clipped = clipText(text, maxChars);
    const warnings = [];
    if (clipped.clipped) warnings.push(`文本结果超过 ${maxChars.toLocaleString()} 字符，已截断。`);
    if (!clipped.text) warnings.push("文件没有可提取的文字。");
    return makeResult(file, type, clipped.clipped || warnings.length ? "partial" : "ok", `${type.toUpperCase()} · ${clipped.text.length.toLocaleString()} 字符`, clipped.text, warnings);
  }

  async function loadPdfJs() {
    const pdfjsLib = await loadLocalScript("PDF.js", "vendor/pdfjs/pdf.min.js", "pdfjsLib");
    if (pdfjsLib.GlobalWorkerOptions) pdfjsLib.GlobalWorkerOptions.workerSrc = localUrl("vendor/pdfjs/pdf.worker.min.js");
    return pdfjsLib;
  }

  async function parsePdfFile(file, maxChars) {
    const pdfjsLib = await loadPdfJs();
    const buffer = await readArrayBuffer(file, CONFIG.maxBinaryFileBytes);
    let loadingTask;
    const warnings = [];
    try {
      loadingTask = pdfjsLib.getDocument({
        data: new Uint8Array(buffer),
        useWorkerFetch: false,
        isEvalSupported: false,
        useSystemFonts: true,
      });
      const pdf = await loadingTask.promise;
      const pageCount = Number(pdf.numPages) || 0;
      const pagesToRead = Math.min(pageCount, CONFIG.maxPdfPages);
      if (pageCount > pagesToRead) warnings.push(`PDF 共 ${pageCount} 页，仅读取前 ${pagesToRead} 页。`);
      const parts = [];
      let length = 0;
      for (let pageNumber = 1; pageNumber <= pagesToRead; pageNumber += 1) {
        if (length >= maxChars) {
          warnings.push(`PDF 文字结果超过 ${maxChars.toLocaleString()} 字符，已截断。`);
          break;
        }
        try {
          const page = await pdf.getPage(pageNumber);
          const content = await page.getTextContent({ disableCombineTextItems: false });
          const pageText = normalizeText((content.items || []).map((item) => `${String(item?.str || "")}${item?.hasEOL ? "\n" : " "}`).join(" "));
          if (!pageText) warnings.push(`第 ${pageNumber} 页没有文字层，可能是扫描页或图片型 PDF。`);
          const section = `--- 第 ${pageNumber} 页 ---\n${pageText || "[无文字层]"}`;
          const available = Math.max(0, maxChars - length);
          const piece = section.slice(0, available);
          parts.push(piece);
          length += piece.length;
          if (piece.length < section.length) warnings.push(`PDF 文字结果超过 ${maxChars.toLocaleString()} 字符，已截断。`);
          if (typeof page.cleanup === "function") page.cleanup();
        } catch (error) {
          warnings.push(`第 ${pageNumber} 页读取失败，已跳过。`);
        }
      }
      const text = parts.join("\n\n");
      if (!text) warnings.push("PDF 没有可提取的文字。");
      const status = warnings.length ? "partial" : "ok";
      return makeResult(file, "pdf", status, `PDF · ${pageCount} 页，已提取 ${text.length.toLocaleString()} 字符`, text, warnings);
    } finally {
      if (loadingTask && typeof loadingTask.destroy === "function") {
        try { await loadingTask.destroy(); } catch (error) { /* cleanup best effort */ }
      }
    }
  }

  async function loadMammoth() {
    return loadLocalScript("Mammoth", "vendor/mammoth.browser.min.js", "mammoth");
  }

  async function parseDocxFile(file, maxChars) {
    const buffer = await readArrayBuffer(file, CONFIG.maxBinaryFileBytes);
    const mammoth = await loadMammoth();
    const converted = await mammoth.extractRawText({ arrayBuffer: buffer });
    const warnings = [];
    (converted.messages || []).forEach((message) => {
      const text = String(message?.message || "").trim();
      if (text) warnings.push(text.slice(0, 240));
    });
    const clipped = clipText(normalizeText(converted.value || ""), maxChars);
    if (clipped.clipped) warnings.push(`DOCX 结果超过 ${maxChars.toLocaleString()} 字符，已截断。`);
    if (!clipped.text) warnings.push("DOCX 没有可提取的正文文字。");
    const status = warnings.length ? "partial" : "ok";
    return makeResult(file, "docx", status, `DOCX · ${clipped.text.length.toLocaleString()} 字符（正文/表格文本）`, clipped.text, warnings);
  }

  async function loadSheetJs() {
    return loadLocalScript("SheetJS", "vendor/xlsx.full.min.js", "XLSX");
  }

  function cellText(cell) {
    if (!cell) return "";
    if (cell.w !== undefined && cell.w !== null) return String(cell.w);
    if (cell.v === undefined || cell.v === null) return "";
    if (cell.v instanceof Date) return cell.v.toISOString();
    if (typeof cell.v === "object") {
      try { return JSON.stringify(cell.v); } catch (error) { return String(cell.v); }
    }
    return String(cell.v);
  }

  async function parseSpreadsheetFile(file, type, maxChars) {
    const buffer = await readArrayBuffer(file, CONFIG.maxBinaryFileBytes);
    const XLSX = await loadSheetJs();
    const warnings = [];
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellNF: false,
      cellText: true,
      cellStyles: false,
      bookVBA: false,
      WTF: false,
    });
    const names = Array.isArray(workbook.SheetNames) ? workbook.SheetNames : [];
    const sheetsToRead = Math.min(names.length, CONFIG.maxWorkbookSheets);
    if (names.length > sheetsToRead) warnings.push(`工作簿共 ${names.length} 个工作表，仅读取前 ${sheetsToRead} 个。`);
    const parts = [];
    let length = 0;
    for (let sheetIndex = 0; sheetIndex < sheetsToRead; sheetIndex += 1) {
      if (length >= maxChars) {
        warnings.push(`Excel 文字结果超过 ${maxChars.toLocaleString()} 字符，已截断。`);
        break;
      }
      const name = String(names[sheetIndex] || `Sheet${sheetIndex + 1}`).replace(/[\r\n]+/g, " ");
      const sheet = workbook.Sheets?.[names[sheetIndex]];
      if (!sheet || !sheet["!ref"]) {
        warnings.push(`工作表“${name}”没有可读取的单元格范围。`);
        parts.push(`--- 工作表：${name} ---\n[空工作表]`);
        continue;
      }
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      const startRow = Math.max(0, range.s?.r || 0);
      const startColumn = Math.max(0, range.s?.c || 0);
      const endRow = Math.min(range.e?.r ?? startRow, startRow + CONFIG.maxWorksheetRows - 1);
      const endColumn = Math.min(range.e?.c ?? startColumn, startColumn + CONFIG.maxWorksheetColumns - 1);
      if ((range.e?.r ?? endRow) > endRow) warnings.push(`工作表“${name}”超过 ${CONFIG.maxWorksheetRows} 行，已截断。`);
      if ((range.e?.c ?? endColumn) > endColumn) warnings.push(`工作表“${name}”超过 ${CONFIG.maxWorksheetColumns} 列，已截断。`);
      const rows = [];
      for (let row = startRow; row <= endRow; row += 1) {
        const values = [];
        for (let column = startColumn; column <= endColumn; column += 1) {
          const address = XLSX.utils.encode_cell({ r: row, c: column });
          values.push(cellText(sheet[address]).replace(/[\t\r\n]+/g, " "));
        }
        rows.push(values.join("\t").replace(/\s+$/, ""));
      }
      const section = `--- 工作表：${name} ---\n${rows.join("\n")}`;
      const available = Math.max(0, maxChars - length);
      const piece = section.slice(0, available);
      parts.push(piece);
      length += piece.length;
      if (piece.length < section.length) warnings.push(`Excel 文字结果超过 ${maxChars.toLocaleString()} 字符，已截断。`);
    }
    if (!names.length) warnings.push("Excel 工作簿没有工作表。");
    const text = parts.join("\n\n");
    if (!text) warnings.push("Excel 没有可提取的单元格内容。");
    const status = warnings.length ? "partial" : "ok";
    return makeResult(file, type, status, `${type.toUpperCase()} · ${names.length} 个工作表，${text.length.toLocaleString()} 字符（TSV）`, text, warnings);
  }

  function unsupportedResult(file, type) {
    const extension = extensionOf(fileName(file));
    if (extension === "doc" || type === "doc") {
      return makeResult(file, "doc", "error", "不支持旧版 DOC 文件", "", ["当前版本只支持 .docx；请在 Word 中另存为 DOCX。"]);
    }
    const shown = extension ? `.${extension}` : "该文件类型";
    return makeResult(file, type, "error", `不支持 ${shown}`, "", [`支持类型：${SUPPORTED_EXTENSIONS.map((item) => `.${item}`).join("、")}`]);
  }

  async function parseOne(file, maxChars) {
    const type = detectType(file);
    if (!SUPPORTED_EXTENSIONS.includes(type)) return unsupportedResult(file, type);
    if (type === "docx") return parseDocxFile(file, maxChars);
    if (type === "pdf") return parsePdfFile(file, maxChars);
    if (type === "xlsx" || type === "xls") return parseSpreadsheetFile(file, type, maxChars);
    return parseTextFile(file, type, maxChars);
  }

  async function parseFiles(input) {
    const files = normalizeFiles(input);
    const results = [];
    let remaining = CONFIG.maxTotalOutputChars;
    for (const file of files) {
      if (remaining <= 0) {
        results.push(makeResult(file, detectType(file), "skipped", "达到本次解析总字符上限", "", [`本次解析最多保留 ${CONFIG.maxTotalOutputChars.toLocaleString()} 字符，已跳过该文件。`]));
        continue;
      }
      try {
        const result = await parseOne(file, Math.min(CONFIG.maxSingleOutputChars, remaining));
        if (result.text.length > remaining) {
          result.text = result.text.slice(0, remaining);
          result.status = "partial";
          result.warnings.push(`本次解析总结果超过 ${CONFIG.maxTotalOutputChars.toLocaleString()} 字符，已截断。`);
        }
        remaining -= result.text.length;
        results.push(result);
      } catch (error) {
        const type = detectType(file);
        const message = String(error?.message || "读取文件失败。").slice(0, 240);
        results.push(makeResult(file, type, "error", `${type.toUpperCase()} 读取失败`, "", [message]));
      }
    }
    return results;
  }

  function selfCheck() {
    const sample = "<h1>本地材料</h1><script>window.__shouldNotRun = true;</script><p>关键事实</p>";
    const text = toPlainMarkupText(sample);
    return {
      ok: text.includes("本地材料") && text.includes("关键事实") && !text.includes("shouldNotRun"),
      text,
    };
  }

  global.PPTDLFileParser = Object.freeze({
    version: CONFIG.version,
    supportedTypes: SUPPORTED_EXTENSIONS,
    supportedExtensions: SUPPORTED_EXTENSIONS,
    typeInfo: TYPE_INFO,
    limits: CONFIG,
    detectType,
    isSupported: (file) => SUPPORTED_EXTENSIONS.includes(detectType(file)),
    parseFiles,
    toPlainMarkupText,
    selfCheck,
  });
})(window);

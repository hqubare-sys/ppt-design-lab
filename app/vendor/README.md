# 本地浏览器依赖

这些文件随静态工具一同发布，运行时不会从 CDN 或其他网站下载代码。

- `pptxgen.bundle.js`：PptxGenJS 4.0.1，用于在浏览器中生成 PPTX。
- `html-to-image.js`：html-to-image 1.11.13，用于把每页 HTML 幻灯片渲染为 PNG。
- `pdfjs/pdf.min.js` 与 `pdfjs/pdf.worker.min.js`：Mozilla PDF.js 3.11.174 的 legacy 浏览器构建，用于提取文字型 PDF 的文字层。
- `mammoth.browser.min.js`：Mammoth 1.8.0 浏览器构建，用于提取 DOCX 正文和基本表格文字。
- `xlsx.full.min.js`：SheetJS CE 0.18.5 浏览器构建，用于把 XLSX/XLS 工作表读取为 TSV 文本。

对应许可证保存在同一目录：`LICENSE-pptxgenjs.txt`、`LICENSE-html-to-image.txt`、`LICENSE-pdfjs.txt`、`LICENSE-mammoth.txt`、`LICENSE-xlsx.txt`。上述新增构建均从 npm 官方包取得，版本、来源和许可证分别为：

- `pdfjs-dist@3.11.174`（Mozilla Foundation，Apache-2.0；来源：<https://www.npmjs.com/package/pdfjs-dist>）
- `mammoth@1.8.0`（Michael Williamson，BSD-2-Clause；来源：<https://www.npmjs.com/package/mammoth>）
- `xlsx@0.18.5`（SheetJS，Apache-2.0；来源：<https://www.npmjs.com/package/xlsx>）

升级依赖时，应重新检查许可证、浏览器兼容性、网络请求行为和导出结果。文件解析器只把用户主动选择的文件读入当前浏览器内存，不通过网络接口上传。

#!/usr/bin/env python3
"""Deterministic governance audit for PPT Design Lab."""

from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REQUIRED = (
    "README.md",
    "AGENTS.md",
    "docs/workflow.md",
    "docs/acceptance.md",
    "docs/architecture.md",
    "app/index.html",
    "app/styles.css",
    "app/app.js",
    "app/sanitizer.js",
    "app/project-io.js",
    "app/exporter.js",
    "app/file-parser.js",
    "app/vendor/pptxgen.bundle.js",
    "app/vendor/html-to-image.js",
    "app/vendor/LICENSE-pptxgenjs.txt",
    "app/vendor/LICENSE-html-to-image.txt",
    "app/vendor/pdfjs/pdf.min.js",
    "app/vendor/pdfjs/pdf.worker.min.js",
    "app/vendor/mammoth.browser.min.js",
    "app/vendor/xlsx.full.min.js",
    "app/vendor/LICENSE-pdfjs.txt",
    "app/vendor/LICENSE-mammoth.txt",
    "app/vendor/LICENSE-xlsx.txt",
    "tests/fixture-safe.html",
    "tests/fixture-unsafe.html",
    "tests/check_static.py",
    "scripts/audit_project.py",
)
FORBIDDEN_SUFFIXES = {".pptx", ".ppt", ".docx", ".xlsx", ".xls", ".pdf"}
SECRET_PATTERNS = (
    re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
    re.compile(r"(?i)(?:api[_-]?key|access[_-]?token|client[_-]?secret)\s*[:=]\s*['\"][^'\"]{8,}['\"]"),
)
NETWORK_API_PATTERNS = (
    re.compile(r"\bfetch\s*\("),
    re.compile(r"\bXMLHttpRequest\b"),
    re.compile(r"\bWebSocket\s*\("),
    re.compile(r"\bEventSource\s*\("),
    re.compile(r"\bsendBeacon\s*\("),
)


def main() -> int:
    errors: list[str] = []

    for relative in REQUIRED:
        if not (ROOT / relative).is_file():
            errors.append(f"missing required file: {relative}")

    for path in ROOT.rglob("*"):
        if path.is_symlink():
            try:
                path.resolve().relative_to(ROOT.resolve())
            except ValueError:
                errors.append(f"external symlink is not allowed: {path.relative_to(ROOT)}")
            continue
        if not path.is_file():
            continue
        relative = path.relative_to(ROOT)
        if path.suffix.lower() in FORBIDDEN_SUFFIXES:
            errors.append(f"office or PDF artifact must not be committed: {relative}")
        if path.stat().st_size > 2_000_000:
            errors.append(f"unexpected large file requires review: {relative}")
        if path.suffix.lower() not in {".md", ".py", ".js", ".ts", ".json", ".html", ".css", ".yml", ".yaml", ".txt"}:
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            errors.append(f"text-like file is not UTF-8: {relative}")
            continue
        for pattern in SECRET_PATTERNS:
            if pattern.search(text):
                errors.append(f"possible credential material requires review: {relative}")
                break

    required_statements = {
        "README.md": ("公司业务内容", "永不复制到 Projects、YYOB、GitHub"),
        "AGENTS.md": ("不可突破的隐私边界", "不逆向LorealGPT接口"),
        "docs/acceptance.md": ("没有携带用户内容的对外网络请求",),
    }
    for relative, statements in required_statements.items():
        path = ROOT / relative
        if not path.is_file():
            continue
        text = path.read_text(encoding="utf-8")
        for statement in statements:
            if statement not in text:
                errors.append(f"missing governance statement in {relative}: {statement}")

    # The shipped application must not contain code paths capable of sending
    # imported slide content over the network. Sanitizer patterns and test
    # fixtures are intentionally excluded from this executable-code check.
    for path in (ROOT / "app").glob("*.js"):
        if not path.is_file():
            continue
        text = path.read_text(encoding="utf-8")
        for pattern in NETWORK_API_PATTERNS:
            if pattern.search(text):
                errors.append(
                    f"network API is not allowed in shipped application: {path.relative_to(ROOT)}"
                )
                break

    index_path = ROOT / "app/index.html"
    if index_path.is_file():
        index_text = index_path.read_text(encoding="utf-8")
        remote_asset = re.compile(
            r"<(?:script|link)\b[^>]+(?:src|href)\s*=\s*['\"](?:https?:)?//",
            re.IGNORECASE,
        )
        if remote_asset.search(index_text):
            errors.append("remote script or stylesheet is not allowed in app/index.html")

    if errors:
        print("PPT Design Lab governance audit failed:")
        for error in errors:
            print(f"- {error}")
        return 1

    print("PPT Design Lab governance audit passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

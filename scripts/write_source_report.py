#!/usr/bin/env python3
"""Write 信息源报告.md from a smart-ppt plan and generation audit."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def load_json(path: str | Path | None) -> dict[str, Any]:
    if not path:
        return {}
    return json.loads(Path(path).read_text(encoding="utf-8"))


def output_pages(audit: dict[str, Any]) -> list[dict[str, Any]]:
    slides = audit.get("slides", [])
    return [{"number": i + 1, **slide} for i, slide in enumerate(slides)]


def write_report(output: Path, plan: dict[str, Any], audit: dict[str, Any]) -> Path:
    pages = output_pages(audit)
    sources = plan.get("sources", [])
    images = plan.get("image_sources", [])
    lines = ["# 信息源报告", "", f"- 对应演示文稿：`{output.name}`", f"- 幻灯片页数：{len(pages) if pages else '未提供审计清单'}", "", "## 按页信息来源", ""]
    for page in pages:
        source_page = page.get("source_page", page["number"] - 1)
        lines.append(f"### 第 {page['number']} 页")
        page_sources = [s for s in sources if s.get("page") == source_page]
        page_images = [s for s in images if s.get("page") == source_page]
        if not page_sources and not page_images:
            lines.append("- 未在计划中登记来源；请核对本页是否仅使用用户材料或原创分析。")
        for source in page_sources:
            title = source.get("title") or source.get("claim") or "未命名来源"
            url = source.get("url") or source.get("link")
            lines.append(f"- 信息来源：{title}" + (f" — {url}" if url else " — 未提供链接"))
            if source.get("accessed"):
                lines.append(f"  - 访问日期：{source['accessed']}")
            if source.get("claim") and source.get("title"):
                lines.append(f"  - 支持内容：{source['claim']}")
        for image in page_images:
            title = image.get("title") or image.get("description") or "配图"
            origin = image.get("url") or image.get("origin") or "来源未记录"
            license_name = image.get("license")
            lines.append(f"- 图片/插画：{title} — {origin}" + (f"；许可：{license_name}" if license_name else ""))
        lines.append("")
    lines.extend(["## 说明", "", "来源信息由生成流程根据实际使用的材料记录。AI生成的插画应标注为“AI生成”；用户提供的文件应标注为“用户提供”。不得将未核实的内容或链接写成已验证来源。", ""])
    report = output.with_name("信息源报告.md")
    report.write_text("\n".join(lines), encoding="utf-8")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path, help="生成的 PPTX 路径")
    parser.add_argument("--plan", type=Path, help="生成计划 JSON")
    parser.add_argument("--audit", type=Path, help="生成审计 JSON")
    args = parser.parse_args()
    report = write_report(args.output, load_json(args.plan), load_json(args.audit))
    print(report)


if __name__ == "__main__":
    main()

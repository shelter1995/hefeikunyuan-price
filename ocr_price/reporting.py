from __future__ import annotations

from pathlib import Path
from typing import Any


def _change(cell: dict[str, Any] | None) -> str:
    if not isinstance(cell, dict):
        return "空 → 空"
    old = cell.get("old")
    new = cell.get("new")
    old_text = old if old is not None else "空"
    new_text = new if new is not None else "空"
    return f"{old_text} → {new_text}"


def _section(title: str, prefix: str, summary: dict[str, Any]) -> list[str]:
    updated = summary.get("updated_items") or []
    skipped = summary.get("skipped_items") or []
    lines = [f"### {title}", ""]
    lines.append(f"**已更新（{len(updated)}家）**：")
    lines.append("")

    if prefix == "G":
        lines.extend(
            [
                "| 序号 | 厂家 | G1 | G3(盘螺) | G4(螺纹) |",
                "|------|------|-----|----------|----------|",
            ]
        )
        for idx, row in enumerate(updated, 1):
            lines.append(
                f"| {idx} | {row.get('项目文件Sheet', '')} | "
                f"{_change(row.get('G1'))} | {_change(row.get('G3'))} | {_change(row.get('G4'))} |"
            )
    else:
        lines.extend(
            [
                "| 序号 | 厂家 | H1 | H3(盘螺) | H4(螺纹) |",
                "|------|------|-----|----------|----------|",
            ]
        )
        for idx, row in enumerate(updated, 1):
            lines.append(
                f"| {idx} | {row.get('项目文件Sheet', '')} | "
                f"{_change(row.get('H1'))} | {_change(row.get('H3'))} | {_change(row.get('H4'))} |"
            )

    lines.append("")
    lines.append(f"**未更新（{len(skipped)}家）**：")
    if skipped:
        for row in skipped:
            lines.append(f"- {row.get('项目文件Sheet', '')} - {row.get('原因', '')}")
    else:
        lines.append("- 无")
    lines.append("")
    return lines


def render_single_report_markdown(result: dict[str, Any], json_report_path: str) -> str:
    project_name = Path(str(result.get("project") or "")).name
    started = str(result.get("started_at") or "")
    ended = str(result.get("ended_at") or "")
    mode = str(result.get("mode") or "")
    status = str(result.get("status") or "")

    lines = [
        "## 项目报价单文件更新完成",
        "",
        f"**执行时间**：{started} 至 {ended}",
        f"**更新模式**：{mode}",
        f"**执行结果**：{status}",
        "",
        f"## 一、{project_name}",
        "",
    ]

    web = result.get("web")
    if isinstance(web, dict) and web.get("status") == "ok":
        lines.extend(_section("1. 网价更新（G1/G3/G4）", "G", web.get("apply_summary") or {}))

    image_doc = result.get("image_doc")
    if isinstance(image_doc, dict) and image_doc.get("status") == "ok":
        lines.extend(_section("2. 图片/文档价更新（H1/H3/H4）", "H", image_doc.get("apply_summary") or {}))

    lines.extend(
        [
            "---",
            "",
            "## 报告文件",
            f"完整JSON报告已保存至：`{json_report_path}`",
            "",
        ]
    )
    return "\n".join(lines)

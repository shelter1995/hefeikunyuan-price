from __future__ import annotations

from pathlib import Path
from typing import Any


def _inventory_spec_text(item: dict[str, Any]) -> str:
    parts = [
        str(item.get("product") or "").strip(),
        str(item.get("spec") or "").strip(),
        str(item.get("length") or "").strip(),
        str(item.get("material") or "").strip(),
    ]
    return " ".join(x for x in parts if x) or "未识别规格"


def _inventory_rows(inventory_report: dict[str, Any]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for item in inventory_report.get("applied", []) or []:
        if not isinstance(item, dict):
            continue
        rows.append(
            {
                "mill": str(item.get("sheet_mill") or item.get("mill") or ""),
                "spec": _inventory_spec_text(item),
                "status": str(item.get("status") or ""),
                "cell": str(item.get("cell") or ""),
            }
        )
    return rows


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


def _pending_lines(flow_name: str, flow: dict[str, Any]) -> list[str]:
    if flow.get("status") != "pending_confirmation":
        return []
    lines = [f"### {flow_name}待确认事项", ""]
    reason = str(flow.get("reason") or "").strip()
    if reason:
        lines.append(f"**阻断原因**：{reason}")
        lines.append("")
    mapping = str(flow.get("pending_mapping_json") or "").strip()
    if mapping:
        lines.append(f"**待确认对照表**：`{mapping}`")
        lines.append("")

    details = flow.get("pending_details") if isinstance(flow.get("pending_details"), dict) else {}
    rows: list[dict[str, Any]] = []
    rows.extend(details.get("pending_matches") or [])
    rows.extend(details.get("pending_new") or [])
    if rows:
        lines.extend(
            [
                "| 序号 | 项目Sheet | 来源Sheet | 状态 | 说明 |",
                "|------|----------|----------|------|------|",
            ]
        )
        for idx, row in enumerate(rows, 1):
            lines.append(
                f"| {idx} | {row.get('项目文件Sheet', '')} | "
                f"{row.get('最新清单厂家Sheet', '')} | {row.get('状态', '')} | "
                f"{row.get('说明', '')} |"
            )
    else:
        lines.append("- 无可展示明细，请查看 JSON 报告。")
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

    if status == "ok":
        lines.extend(
            [
                "### 复用确认记录说明",
                "本次未发现新厂家或待确认项；已确认对照关系按历史记录复用。",
                "",
            ]
        )

    web = result.get("web")
    if isinstance(web, dict) and web.get("status") == "ok":
        lines.extend(_section("1. 网价更新（G1/G3/G4）", "G", web.get("apply_summary") or {}))
    elif isinstance(web, dict):
        lines.extend(_pending_lines("网价", web))

    image_doc = result.get("image_doc")
    if isinstance(image_doc, dict) and image_doc.get("status") == "ok":
        lines.extend(_section("2. 图片/文档价更新（H1/H3/H4）", "H", image_doc.get("apply_summary") or {}))
    elif isinstance(image_doc, dict):
        lines.extend(_pending_lines("图片/文档", image_doc))

    # Inventory color report
    inventory_report = image_doc.get("inventory_report") if isinstance(image_doc, dict) else None
    if isinstance(inventory_report, dict) and inventory_report.get("status") == "ok":
        rows = _inventory_rows(inventory_report)
        lines.append("")
        lines.append("### 3. 库存颜色标注明细")
        lines.append("")
        lines.append("| 厂家 | 钢材型号和规格 | 库存情况 | 单元格 |")
        lines.append("|------|----------------|----------|--------|")
        if rows:
            for row in rows:
                lines.append(f"| {row['mill']} | {row['spec']} | {row['status']} | {row['cell']} |")
        else:
            lines.append("| 无 | 无 | 无 | 无 |")
        lines.append("")
    elif isinstance(inventory_report, dict) and inventory_report.get("status") not in {None, "skipped"}:
        lines.append("")
        lines.append("### 3. 库存颜色标注异常")
        lines.append("")
        lines.append(str(inventory_report.get("error") or inventory_report.get("reason") or "未知错误"))
        lines.append("")

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

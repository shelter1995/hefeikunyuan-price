from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from ocr_price.inventory import InventoryItem, apply_inventory_to_project


def _build_quote_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"

    # Mill header used by inventory writeback.
    ws.cell(row=1, column=5, value="目标钢厂")

    # Row 12 is treated as rebar by current row mapping logic.
    ws.cell(row=12, column=2, value="12")
    ws.cell(row=12, column=3, value="9")
    ws.cell(row=12, column=4, value="HRB400E")

    # Another row to verify clear_existing_colors removes stale fill.
    ws.cell(row=13, column=2, value="14")
    ws.cell(row=13, column=3, value="9")
    ws.cell(row=13, column=4, value="HRB400E")
    ws.cell(row=13, column=5).fill = PatternFill(
        start_color="FF0000", end_color="FF0000", fill_type="solid"
    )

    wb.save(path)


def test_apply_inventory_uses_confirmed_mapping_and_clears_existing_colors(tmp_path: Path):
    project_path = tmp_path / "project.xlsx"
    _build_quote_workbook(project_path)

    mapping_path = tmp_path / "mapping.json"
    mapping_rows = [
        {
            "项目文件Sheet": "目标钢厂",
            "最新清单厂家Sheet": "来源钢厂",
            "状态": "已确认匹配",
            "说明": "",
        }
    ]
    mapping_path.write_text(json.dumps(mapping_rows, ensure_ascii=False), encoding="utf-8")

    mill_inventories = {
        "来源钢厂": [
            InventoryItem(
                product="螺纹",
                spec="12",
                length="9",
                material="HRB400E",
                status="告警",
            )
        ]
    }

    result = apply_inventory_to_project(
        project_excel=project_path,
        mill_inventories=mill_inventories,
        sheet_name="报价表",
        mapping_json_path=mapping_path,
        clear_existing_colors=True,
    )

    assert result["status"] == "ok"
    assert result["applied_count"] == 1
    assert result["cleared_count"] >= 2

    wb = load_workbook(project_path)
    ws = wb["报价表"]
    updated_cell = ws.cell(row=12, column=5)
    assert updated_cell.fill.fill_type == "solid"
    assert updated_cell.fill.start_color.rgb in {"00FFC000", "FFC000"}

    cleared_cell = ws.cell(row=13, column=5)
    assert cleared_cell.fill.fill_type is None

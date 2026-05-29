from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from ocr_price.inventory import (
    InventoryItem,
    apply_inventory_to_project,
    load_inventory_from_sources,
)


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

    applied = result["applied"][0]
    assert applied["mill"] == "来源钢厂"
    assert applied["sheet_mill"] == "目标钢厂"
    assert applied["product"] == "螺纹"
    assert applied["spec"] == "12"
    assert applied["length"] == "9"
    assert applied["material"] == "HRB400E"
    assert applied["status"] == "告警"
    assert applied["cell"] == "E12"

    wb = load_workbook(project_path)
    ws = wb["报价表"]
    updated_cell = ws.cell(row=12, column=5)
    assert updated_cell.fill.fill_type == "solid"
    assert updated_cell.fill.start_color.rgb in {"00FFC000", "FFC000"}

    cleared_cell = ws.cell(row=13, column=5)
    assert cleared_cell.fill.fill_type is None


def test_load_inventory_from_vision_json_preserves_structured_statuses(tmp_path: Path):
    source = tmp_path / "ocr价格提取_徐钢.json"
    source.write_text(
        json.dumps(
            {
                "meta": {"input_file": str(tmp_path / "徐钢.jpg")},
                "_vision_result": {
                    "库存情况": [
                        {"规格": "9米螺纹14E", "状态": "告警", "原始描述": "极少"},
                        {"规格": "12米螺纹16E", "状态": "告警", "原始描述": "28件"},
                        {"规格": "盘螺8E", "状态": "告警", "原始描述": "16件"},
                    ]
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    items = load_inventory_from_sources([source], "徐钢")

    assert InventoryItem(
        product="螺纹",
        spec="14",
        length="9",
        material=None,
        status="告警",
        note="极少",
    ) in items
    assert InventoryItem(
        product="螺纹",
        spec="16",
        length="12",
        material=None,
        status="告警",
        note="28件",
    ) in items
    assert InventoryItem(
        product="盘螺",
        spec="8",
        length=None,
        material=None,
        status="告警",
        note="16件",
    ) in items


def test_apply_inventory_matches_rebar_and_coil_rows(tmp_path: Path):
    project_path = tmp_path / "project.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    ws.cell(row=1, column=5, value="徐钢")

    ws.cell(row=13, column=2, value="8")
    ws.cell(row=13, column=3, value=None)
    ws.cell(row=13, column=4, value="HRB400E")

    ws.cell(row=19, column=2, value="14")
    ws.cell(row=19, column=3, value="9")
    ws.cell(row=19, column=4, value="HRB400E")

    wb.save(project_path)

    result = apply_inventory_to_project(
        project_excel=project_path,
        mill_inventories={
            "徐钢": [
                InventoryItem(
                    product="盘螺",
                    spec="8",
                    length=None,
                    material=None,
                    status="告警",
                    note="16件",
                ),
                InventoryItem(
                    product="螺纹",
                    spec="14",
                    length="9",
                    material=None,
                    status="告警",
                    note="极少",
                ),
            ]
        },
        sheet_name="报价表",
    )

    assert result["status"] == "ok"
    assert result["applied_count"] == 2

    wb = load_workbook(project_path)
    ws = wb["报价表"]
    assert ws.cell(row=13, column=5).fill.fill_type == "solid"
    assert ws.cell(row=13, column=5).fill.start_color.rgb in {"00FFC000", "FFC000"}
    assert ws.cell(row=19, column=5).fill.fill_type == "solid"
    assert ws.cell(row=19, column=5).fill.start_color.rgb in {"00FFC000", "FFC000"}

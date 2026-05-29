from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook

from ocr_price.writeback_image_doc import prepare_mapping


def test_prepare_image_mapping_treats_xugang_gang_gang_typo_as_pending_match(tmp_path: Path):
    project = tmp_path / "项目.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    wb.create_sheet("徐钢")
    wb.save(project)

    source = tmp_path / "ocr价格提取_徐刚.json"
    source.write_text(
        json.dumps(
            {
                "meta": {"input_file": "线下报价/徐刚.jpg"},
                "quote_date": "2026-05-29",
                "records": [
                    {
                        "location": "蚌埠",
                        "rebar_price": 3187,
                        "coil_price": 3347,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    mapping_json = tmp_path / "图片文档厂家对照表_安徽蚌埠_待确认.json"
    mapping_csv = tmp_path / "图片文档厂家对照表_安徽蚌埠_待确认.csv"
    report = tmp_path / "report.json"

    result = prepare_mapping(
        project_excel=project,
        source_json_paths=[source],
        location="蚌埠",
        mapping_json_out=mapping_json,
        mapping_csv_out=mapping_csv,
        report_out=report,
    )

    rows = json.loads(mapping_json.read_text(encoding="utf-8"))
    xugang = next(row for row in rows if row["项目文件Sheet"] == "徐钢")
    assert xugang["最新清单厂家Sheet"] == "徐刚"
    assert xugang["状态"] == "待确认匹配"
    assert result["pending_count"] == 1
    assert result["unmapped_sources"] == []

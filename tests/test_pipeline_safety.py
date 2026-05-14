from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook, load_workbook

from ocr_price import pipeline
from ocr_price.writeback_image_doc import apply_writeback


def _write_json(path: Path, payload: object) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _make_project(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    mill = wb.create_sheet("测试钢厂")
    mill["G3"] = 3300
    mill["G4"] = 3100
    mill["H3"] = 3300
    mill["H4"] = 3100
    wb.save(path)


def test_apply_writeback_dry_run_does_not_modify_workbook(tmp_path: Path):
    project = tmp_path / "project.xlsx"
    source = tmp_path / "source.json"
    mapping = tmp_path / "mapping.json"
    report = tmp_path / "report.json"
    _make_project(project)
    _write_json(
        source,
        {
            "meta": {"input_file": "测试钢厂报价.txt"},
            "quote_date": "2026-05-11",
            "records": [
                {"location": "蚌埠", "coil_price": 3400, "rebar_price": 3200},
            ],
        },
    )
    _write_json(
        mapping,
        [
            {
                "项目文件Sheet": "测试钢厂",
                "最新清单厂家Sheet": "测试钢厂",
                "状态": "已确认匹配",
                "说明": "",
            }
        ],
    )

    result = apply_writeback(
        project_excel=project,
        source_json_paths=[source],
        mapping_json_path=mapping,
        location="蚌埠",
        report_out=report,
        dry_run=True,
    )

    wb = load_workbook(project)
    ws = wb["测试钢厂"]
    assert result["dry_run"] is True
    assert result["updated_count"] == 1
    assert result["backup_file"] is None
    assert ws["H3"].value == 3300
    assert ws["H4"].value == 3100
    wb.close()


def test_file_sha256_changes_when_file_changes(tmp_path: Path):
    path = tmp_path / "project.xlsx"
    path.write_bytes(b"before")
    before = pipeline._file_sha256(path)

    path.write_bytes(b"after")
    after = pipeline._file_sha256(path)

    assert before != after


def test_assert_dry_run_unchanged_raises_on_modified_file(tmp_path: Path):
    path = tmp_path / "project.xlsx"
    path.write_bytes(b"before")
    before = pipeline._file_sha256(path)
    path.write_bytes(b"after")

    try:
        pipeline._assert_dry_run_unchanged(path, before)
    except RuntimeError as exc:
        assert "dry-run 修改了项目 Excel" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError when dry-run changes the workbook")


def test_pipeline_lock_blocks_existing_active_lock(tmp_path: Path):
    lock_path = tmp_path / ".quote_update.lock"
    lock_path.write_text(
        json.dumps({"pid": 999999, "project": "project.xlsx"}, ensure_ascii=False),
        encoding="utf-8",
    )

    try:
        with pipeline._pipeline_lock(lock_path, project=tmp_path / "project.xlsx"):
            raise AssertionError("lock should block")
    except RuntimeError as exc:
        assert "已有报价更新任务锁" in str(exc)

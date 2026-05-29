from __future__ import annotations

import json
from pathlib import Path

import pytest

from ocr_price import pipeline


def _manifest(tmp_path: Path, project: Path, source: Path) -> dict:
    artifact_dir = tmp_path / "artifacts"
    artifact_dir.mkdir()
    return {
        "project": str(project),
        "mode": "both",
        "artifact_dir": str(artifact_dir),
        "project_excel_hash_after": pipeline._file_sha256(project),
        "web": {
            "status": "ok",
            "location": "安徽合肥",
            "raw_price_excel": str(source),
            "raw_price_excel_sha256": pipeline._file_sha256(source),
        },
        "image_doc": {
            "status": "ok",
            "location": "安徽蚌埠",
            "source_jsons": [str(source)],
            "source_json_sha256": {str(source): pipeline._file_sha256(source)},
        },
    }


def test_validate_manifest_rejects_project_path_mismatch(tmp_path: Path):
    project = tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"
    other = tmp_path / "安徽合肥-安徽蚌埠-其他.xlsx"
    source = tmp_path / "source.json"
    project.write_bytes(b"project")
    other.write_bytes(b"other")
    source.write_text("{}", encoding="utf-8")

    manifest = _manifest(tmp_path, project, source)

    with pytest.raises(RuntimeError, match="manifest项目不匹配"):
        pipeline._validate_manifest_for_confirm(other, manifest, expected_mode="both")


def test_validate_manifest_rejects_changed_project_hash(tmp_path: Path):
    project = tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"
    source = tmp_path / "source.json"
    project.write_bytes(b"project-before")
    source.write_text("{}", encoding="utf-8")
    manifest = _manifest(tmp_path, project, source)
    project.write_bytes(b"project-after")

    with pytest.raises(RuntimeError, match="项目Excel已变化"):
        pipeline._validate_manifest_for_confirm(project, manifest, expected_mode="both")


def test_validate_manifest_rejects_changed_source_hash(tmp_path: Path):
    project = tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"
    source = tmp_path / "source.json"
    project.write_bytes(b"project")
    source.write_text('{"old": true}', encoding="utf-8")
    manifest = _manifest(tmp_path, project, source)
    source.write_text('{"new": true}', encoding="utf-8")

    with pytest.raises(RuntimeError, match="manifest源文件已变化"):
        pipeline._validate_manifest_for_confirm(project, manifest, expected_mode="both")


def test_write_manifest_records_artifact_hashes(tmp_path: Path):
    project = tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"
    source = tmp_path / "ocr价格提取_徐钢.json"
    raw = tmp_path / "安徽合肥2026-05-23建筑钢材原料价格清单.xlsx"
    project.write_bytes(b"project")
    source.write_text("{}", encoding="utf-8")
    raw.write_bytes(b"raw")
    result = {
        "project": str(project),
        "mode": "both",
        "web": {"raw_price_excel": str(raw)},
        "image_doc": {"source_jsons": [str(source)]},
    }
    out = tmp_path / "dry_run_manifest.json"

    pipeline._write_manifest(out, result, tmp_path)

    manifest = json.loads(out.read_text(encoding="utf-8"))
    assert manifest["web"]["raw_price_excel_sha256"] == pipeline._file_sha256(raw)
    assert manifest["image_doc"]["source_json_sha256"][str(source)] == pipeline._file_sha256(source)

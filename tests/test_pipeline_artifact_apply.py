from __future__ import annotations

import json
from pathlib import Path

from ocr_price import pipeline


def _ok_summary() -> dict:
    return {"updated_count": 0, "skipped_count": 0, "updated_items": [], "skipped_items": []}


def test_run_single_confirm_write_reuses_artifacts_by_default(monkeypatch, tmp_path: Path):
    calls: list[str] = []

    def fake_web_apply_from_artifacts(*args, **kwargs):
        calls.append("web_artifact_apply")
        return {"status": "ok", "apply_summary": _ok_summary()}

    def fake_image_apply_from_artifacts(*args, **kwargs):
        calls.append("image_artifact_apply")
        return {"status": "ok", "apply_summary": _ok_summary()}

    def fail_web_flow(*args, **kwargs):
        raise AssertionError("confirm-write默认应复用产物，不应重跑web flow")

    def fail_image_flow(*args, **kwargs):
        raise AssertionError("confirm-write默认应复用产物，不应重跑image flow")

    monkeypatch.setattr(pipeline, "_web_apply_from_artifacts", fake_web_apply_from_artifacts)
    monkeypatch.setattr(pipeline, "_image_apply_from_artifacts", fake_image_apply_from_artifacts)
    monkeypatch.setattr(pipeline, "_web_flow", fail_web_flow)
    monkeypatch.setattr(pipeline, "_image_flow", fail_image_flow)

    result = pipeline.run_single(
        project=tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx",
        mode="both",
        list_url=None,
        detail_url=None,
        account_file=tmp_path / "网站账号密码.txt",
        username=None,
        password=None,
        headless=True,
        image_inputs=[],
        image_jsons=[],
        artifact_dir=tmp_path / "运行产物",
        dry_run=False,
        confirm_write=True,
    )

    assert result["status"] == "ok"
    assert calls == ["web_artifact_apply", "image_artifact_apply"]


def test_run_single_confirm_write_with_refresh_uses_full_flow(monkeypatch, tmp_path: Path):
    calls: list[str] = []

    def fake_web_flow(*args, **kwargs):
        calls.append("web_flow")
        return {"status": "ok", "apply_summary": _ok_summary()}

    def fail_web_apply_from_artifacts(*args, **kwargs):
        raise AssertionError("refresh-web-artifacts后应重跑web flow")

    monkeypatch.setattr(pipeline, "_web_flow", fake_web_flow)
    monkeypatch.setattr(pipeline, "_web_apply_from_artifacts", fail_web_apply_from_artifacts)

    result = pipeline.run_single(
        project=tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx",
        mode="web",
        list_url=None,
        detail_url=None,
        account_file=tmp_path / "网站账号密码.txt",
        username=None,
        password=None,
        headless=True,
        image_inputs=[],
        image_jsons=[],
        artifact_dir=tmp_path / "运行产物",
        dry_run=False,
        confirm_write=True,
        refresh_web_artifacts=True,
    )

    assert result["status"] == "ok"
    assert calls == ["web_flow"]


def test_run_single_passes_manual_login_timeout_to_web_flow(monkeypatch, tmp_path: Path):
    seen: dict[str, int] = {}

    def fake_web_flow(*args, **kwargs):
        seen["manual_login_timeout_seconds"] = kwargs["manual_login_timeout_seconds"]
        return {"status": "ok", "apply_summary": _ok_summary()}

    monkeypatch.setattr(pipeline, "_web_flow", fake_web_flow)

    result = pipeline.run_single(
        project=tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx",
        mode="web",
        list_url=None,
        detail_url=None,
        account_file=tmp_path / "网站账号密码.txt",
        username=None,
        password=None,
        headless=False,
        image_inputs=[],
        image_jsons=[],
        artifact_dir=tmp_path / "运行产物",
        dry_run=True,
        manual_login_timeout_seconds=45,
    )

    assert result["status"] == "ok"
    assert seen["manual_login_timeout_seconds"] == 45


def test_web_apply_from_artifacts_blocks_when_mapping_pending(tmp_path: Path):
    project = tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"
    artifact_dir = tmp_path / "运行产物"
    artifact_dir.mkdir(parents=True, exist_ok=True)
    pending = artifact_dir / "厂家对照表_安徽合肥_待确认.json"
    pending.write_text(
        json.dumps(
            [
                {
                    "项目文件Sheet": "测试钢厂",
                    "最新清单厂家Sheet": "测试钢厂",
                    "状态": "待确认匹配",
                    "说明": "",
                }
            ],
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    result = pipeline._web_apply_from_artifacts(project=project, location="安徽合肥", artifact_dir=artifact_dir)
    assert result["status"] == "pending_confirmation"


def test_run_single_confirm_write_uses_manifest_paths(monkeypatch, tmp_path: Path):
    manifest = tmp_path / "manifest.json"
    manifest.write_text(
        json.dumps(
            {
                "project": str(tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"),
                "artifact_dir": str(tmp_path / "artifacts"),
                "web": {
                    "location": "安徽合肥",
                    "mapping_json": str(tmp_path / "artifacts" / "厂家对照表_安徽合肥_已确认.json"),
                    "raw_price_excel": str(tmp_path / "artifacts" / "安徽合肥2026-05-13建筑钢材原料价格清单.xlsx"),
                },
                "image_doc": {
                    "location": "安徽蚌埠",
                    "mapping_json": str(tmp_path / "artifacts" / "图片文档厂家对照表_安徽蚌埠_已确认.json"),
                    "source_jsons": [str(tmp_path / "artifacts" / "ocr价格提取_桂鑫报价.json")],
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    seen = {}

    def fake_web_apply_from_manifest(*args, **kwargs):
        seen["web_manifest"] = kwargs["manifest"]["web"]["raw_price_excel"]
        return {"status": "ok", "apply_summary": _ok_summary()}

    def fake_image_apply_from_manifest(*args, **kwargs):
        seen["image_manifest"] = kwargs["manifest"]["image_doc"]["source_jsons"]
        return {"status": "ok", "apply_summary": _ok_summary()}

    monkeypatch.setattr(pipeline, "_web_apply_from_manifest", fake_web_apply_from_manifest)
    monkeypatch.setattr(pipeline, "_image_apply_from_manifest", fake_image_apply_from_manifest)

    result = pipeline.run_single(
        project=tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx",
        mode="both",
        list_url=None,
        detail_url=None,
        account_file=tmp_path / "网站账号密码.txt",
        username=None,
        password=None,
        headless=True,
        image_inputs=[],
        image_jsons=[],
        artifact_dir=tmp_path / "artifacts",
        dry_run=False,
        confirm_write=True,
        manifest_path=manifest,
    )

    assert result["status"] == "ok"
    assert seen["web_manifest"].endswith("建筑钢材原料价格清单.xlsx")
    assert seen["image_manifest"][0].endswith("ocr价格提取_桂鑫报价.json")

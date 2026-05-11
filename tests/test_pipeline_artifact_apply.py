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

from __future__ import annotations

import json
import subprocess
import sys
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
    project = tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"
    artifact_dir = tmp_path / "artifacts"
    raw = artifact_dir / "安徽合肥2026-05-13建筑钢材原料价格清单.xlsx"
    source = artifact_dir / "ocr价格提取_桂鑫报价.json"
    artifact_dir.mkdir()
    project.write_bytes(b"project")
    raw.write_bytes(b"raw")
    source.write_text("{}", encoding="utf-8")
    manifest = tmp_path / "manifest.json"
    manifest.write_text(
        json.dumps(
            {
                "project": str(project),
                "mode": "both",
                "artifact_dir": str(artifact_dir),
                "project_excel_hash_after": pipeline._file_sha256(project),
                "web": {
                    "location": "安徽合肥",
                    "mapping_json": str(artifact_dir / "厂家对照表_安徽合肥_已确认.json"),
                    "raw_price_excel": str(raw),
                    "raw_price_excel_sha256": pipeline._file_sha256(raw),
                },
                "image_doc": {
                    "location": "安徽蚌埠",
                    "mapping_json": str(artifact_dir / "图片文档厂家对照表_安徽蚌埠_已确认.json"),
                    "source_jsons": [str(source)],
                    "source_json_sha256": {str(source): pipeline._file_sha256(source)},
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
        project=project,
        mode="both",
        list_url=None,
        detail_url=None,
        account_file=tmp_path / "网站账号密码.txt",
        username=None,
        password=None,
        headless=True,
        image_inputs=[],
        image_jsons=[],
        artifact_dir=artifact_dir,
        dry_run=False,
        confirm_write=True,
        manifest_path=manifest,
    )

    assert result["status"] == "ok"
    assert seen["web_manifest"].endswith("建筑钢材原料价格清单.xlsx")
    assert seen["image_manifest"][0].endswith("ocr价格提取_桂鑫报价.json")


def test_run_batch_wrapper_exposes_manifest_argument():
    script = Path(__file__).resolve().parents[1] / "skills" / "quote-update" / "scripts" / "run_batch.py"

    result = subprocess.run(
        [sys.executable, str(script), "--help"],
        capture_output=True,
    )

    assert result.returncode == 0
    assert b"--manifest" in result.stdout


def test_run_ocr_for_inputs_continues_and_reports_each_file(monkeypatch, tmp_path: Path):
    ok_input = tmp_path / "ok.jpg"
    bad_input = tmp_path / "bad.jpg"
    ok_input.write_bytes(b"ok")
    bad_input.write_bytes(b"bad")
    artifact_dir = tmp_path / "artifacts"

    def fake_run(cmd, check, env, capture_output, text, encoding, errors):
        out = Path(cmd[cmd.index("--output") + 1])
        if "ok" in str(cmd[cmd.index("--input") + 1]):
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text('{"records":[]}', encoding="utf-8")
            return subprocess.CompletedProcess(cmd, 0, stdout="ok", stderr="")
        return subprocess.CompletedProcess(cmd, 2, stdout="", stderr="MiniMax timeout")

    monkeypatch.setattr(pipeline.subprocess, "run", fake_run)

    report = pipeline._run_ocr_for_inputs(
        [ok_input, bad_input],
        location="蚌埠",
        artifact_dir=artifact_dir,
    )

    assert report["success_count"] == 1
    assert report["failed_count"] == 1
    assert report["outputs"] == [str(artifact_dir / "ocr价格提取_ok.json")]
    assert report["items"][0]["status"] == "ok"
    assert report["items"][1]["status"] == "failed"
    assert "MiniMax timeout" in report["items"][1]["stderr"]


def test_run_single_blocks_when_any_ocr_input_fails(monkeypatch, tmp_path: Path):
    project = tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"
    project.write_bytes(b"project")
    image = tmp_path / "bad.jpg"
    image.write_bytes(b"bad")

    def fake_run_ocr_for_inputs(*args, **kwargs):
        return {
            "outputs": [],
            "items": [
                {
                    "input": str(image),
                    "output": str(tmp_path / "artifacts" / "ocr价格提取_bad.json"),
                    "status": "failed",
                    "stderr": "MiniMax timeout",
                }
            ],
            "success_count": 0,
            "failed_count": 1,
        }

    def fail_image_flow(*args, **kwargs):
        raise AssertionError("OCR失败后不应继续图片/文档回写流程")

    monkeypatch.setattr(pipeline, "_run_ocr_for_inputs", fake_run_ocr_for_inputs)
    monkeypatch.setattr(pipeline, "_image_flow", fail_image_flow)

    result = pipeline.run_single(
        project=project,
        mode="image_doc",
        list_url=None,
        detail_url=None,
        account_file=tmp_path / "网站账号密码.txt",
        username=None,
        password=None,
        headless=True,
        image_inputs=[image],
        image_jsons=[],
        artifact_dir=tmp_path / "artifacts",
        dry_run=True,
    )

    assert result["status"] == "failed"
    assert result["image_doc"]["phase"] == "ocr"
    assert result["image_doc"]["ocr_report"]["failed_count"] == 1

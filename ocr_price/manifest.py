from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_sha256_if_exists(path_text: Any) -> str | None:
    if not path_text:
        return None
    path = Path(str(path_text))
    if not path.exists() or not path.is_file():
        return None
    return file_sha256(path)


def add_artifact_hashes(web: dict[str, Any] | None, image_doc: dict[str, Any] | None) -> None:
    if web is not None:
        raw_hash = file_sha256_if_exists(web.get("raw_price_excel"))
        if raw_hash:
            web["raw_price_excel_sha256"] = raw_hash
    if image_doc is not None:
        source_hashes: dict[str, str] = {}
        for source in image_doc.get("source_jsons", []) or []:
            source_hash = file_sha256_if_exists(source)
            if source_hash:
                source_hashes[str(source)] = source_hash
        if source_hashes:
            image_doc["source_json_sha256"] = source_hashes


def _same_resolved_path(left: Path, right_text: Any) -> bool:
    if not right_text:
        return False
    try:
        return left.resolve() == Path(str(right_text)).resolve()
    except OSError:
        return str(left) == str(right_text)


def _require_manifest_hash(path_text: Any, expected_hash: str | None, label: str) -> None:
    if not path_text:
        return
    if not expected_hash:
        raise RuntimeError(f"manifest缺少{label}hash，不能执行confirm-write。")
    path = Path(str(path_text))
    if not path.exists():
        raise RuntimeError(f"manifest源文件不存在：{path}")
    current_hash = file_sha256(path)
    if current_hash != expected_hash:
        raise RuntimeError(
            f"manifest源文件已变化：{path} expected={expected_hash} actual={current_hash}"
        )


def validate_manifest_for_confirm(project: Path, manifest: dict[str, Any], expected_mode: str) -> None:
    if not _same_resolved_path(project, manifest.get("project")):
        raise RuntimeError(
            f"manifest项目不匹配：当前={project} manifest={manifest.get('project')}"
        )
    manifest_mode = str(manifest.get("mode") or "").strip()
    if manifest_mode and manifest_mode != expected_mode:
        raise RuntimeError(f"manifest模式不匹配：当前={expected_mode} manifest={manifest_mode}")

    expected_project_hash = str(manifest.get("project_excel_hash_after") or "").strip()
    if not expected_project_hash:
        raise RuntimeError("manifest缺少project_excel_hash_after，不能执行confirm-write。")
    if not project.exists():
        raise RuntimeError(f"项目Excel不存在：{project}")
    current_project_hash = file_sha256(project)
    if current_project_hash != expected_project_hash:
        raise RuntimeError(
            f"项目Excel已变化，必须重新dry-run：{project} "
            f"expected={expected_project_hash} actual={current_project_hash}"
        )

    artifact_dir = Path(str(manifest.get("artifact_dir") or ""))
    if not artifact_dir.exists():
        raise RuntimeError(f"manifest产物目录不存在：{artifact_dir}")

    web = manifest.get("web") if isinstance(manifest.get("web"), dict) else {}
    if web and web.get("raw_price_excel"):
        _require_manifest_hash(
            web.get("raw_price_excel"),
            str(web.get("raw_price_excel_sha256") or "").strip(),
            "网价清单",
        )

    image_doc = manifest.get("image_doc") if isinstance(manifest.get("image_doc"), dict) else {}
    source_hashes = image_doc.get("source_json_sha256") if isinstance(image_doc, dict) else {}
    if image_doc and image_doc.get("source_jsons"):
        if not isinstance(source_hashes, dict) or not source_hashes:
            raise RuntimeError("manifest缺少OCR JSON源文件hash，不能执行confirm-write。")
        for source in image_doc.get("source_jsons", []) or []:
            _require_manifest_hash(source, str(source_hashes.get(str(source)) or "").strip(), "OCR JSON")

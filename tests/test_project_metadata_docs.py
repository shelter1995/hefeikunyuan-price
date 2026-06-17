from __future__ import annotations

import tomllib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_docs_do_not_advertise_pdf_offline_sources() -> None:
    doc_paths = [
        ROOT / "README.md",
        ROOT / "AGENTS.md",
        ROOT / "skills" / "quote-update" / "SKILL.md",
        ROOT / "skills" / "quote-update" / "references" / "commands.md",
        *sorted((ROOT / "doc").glob("*.md")),
    ]

    offenders: list[str] = []
    for path in doc_paths:
        text = path.read_text(encoding="utf-8")
        for token in ("PDF", ".pdf", "图片/PDF", "jpg/png/pdf"):
            if token in text:
                offenders.append(f"{path.relative_to(ROOT)} contains {token}")

    assert offenders == []


def test_pyproject_editable_install_limits_package_discovery() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    package_find = data["tool"]["setuptools"]["packages"]["find"]

    assert package_find["include"] == ["ocr_price*"]
    assert "项目报价*" in package_find["exclude"]
    assert "线下报价*" in package_find["exclude"]
    assert "运行产物*" in package_find["exclude"]

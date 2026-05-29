from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _script() -> Path:
    return Path(__file__).resolve().parents[1] / "skills" / "quote-update" / "scripts" / "make_agent_prompt.py"


def test_make_agent_prompt_single_dry_run_contains_safe_workflow():
    result = subprocess.run(
        [
            sys.executable,
            str(_script()),
            "single-dry-run",
            "--project",
            "项目报价/测试项目.xlsx",
            "--mode",
            "both",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    assert result.returncode == 0
    assert "python skills/quote-update/scripts/run_single.py" in result.stdout
    assert '--project "项目报价/测试项目.xlsx"' in result.stdout
    assert "--dry-run" in result.stdout
    assert "--confirm-write" not in result.stdout
    assert "禁止直接修改 Excel" in result.stdout
    assert "Manifest 路径" in result.stdout


def test_make_agent_prompt_single_confirm_requires_manifest():
    result = subprocess.run(
        [
            sys.executable,
            str(_script()),
            "single-confirm",
            "--project",
            "项目报价/测试项目.xlsx",
            "--manifest",
            "运行产物/测试项目/dry_run_manifest.json",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    assert result.returncode == 0
    assert "--confirm-write" in result.stdout
    assert '--manifest "运行产物/测试项目/dry_run_manifest.json"' in result.stdout
    assert "同一次 dry-run" in result.stdout


def test_make_agent_prompt_batch_dry_run_contains_batch_command():
    result = subprocess.run(
        [
            sys.executable,
            str(_script()),
            "batch-dry-run",
            "--project-dir",
            "项目报价",
            "--glob",
            "*.xlsx",
            "--mode",
            "web",
            "--no-headless",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    assert result.returncode == 0
    assert "python skills/quote-update/scripts/run_batch.py" in result.stdout
    assert '--project-dir "项目报价"' in result.stdout
    assert '--glob "*.xlsx"' in result.stdout
    assert "--mode web" in result.stdout
    assert "--headless" not in result.stdout

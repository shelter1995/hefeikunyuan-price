from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


def _script_path() -> Path:
    repo = Path(__file__).resolve().parents[1]
    return repo / "skills" / "quote-update" / "scripts" / "apply_confirmations.py"


def test_apply_web_confirmations_maps_liugang_and_skips_new_mills(tmp_path: Path):
    mapping = tmp_path / "厂家对照表_安徽合肥_待确认.json"
    mapping.write_text(
        json.dumps(
            [
                {"项目文件Sheet": "徐钢", "最新清单厂家Sheet": "徐钢", "状态": "待确认匹配", "说明": ""},
                {"项目文件Sheet": "六钢", "最新清单厂家Sheet": "", "状态": "未匹配(项目有)", "说明": ""},
                {"项目文件Sheet": "", "最新清单厂家Sheet": "六安钢铁", "状态": "待确认(新厂家)", "说明": ""},
                {"项目文件Sheet": "", "最新清单厂家Sheet": "宝钢", "状态": "待确认(新厂家)", "说明": ""},
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    script = _script_path()
    result = subprocess.run(
        [
            sys.executable,
            str(script),
            "--web-mapping-json",
            str(mapping),
            "--web-match",
            "徐钢=徐钢",
            "--web-match",
            "六钢=六安钢铁",
            "--skip-new-mills",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    print(result.stdout)
    print(result.stderr, file=sys.stderr)
    assert result.returncode == 0

    rows = json.loads(mapping.read_text(encoding="utf-8"))
    assert rows[0]["状态"] == "已确认匹配"
    assert rows[1]["最新清单厂家Sheet"] == "六安钢铁"
    assert rows[1]["状态"] == "已确认匹配"
    assert rows[2]["状态"] == "已确认不更新"
    assert rows[3]["状态"] == "已确认不更新"

    summary = json.loads(result.stdout)
    assert summary["pending_count"] == 0


def test_apply_confirmations_supports_image_doc_mapping(tmp_path: Path):
    mapping = tmp_path / "图片文档厂家对照表_安徽蚌埠_待确认.json"
    mapping.write_text(
        json.dumps(
            [
                {"项目文件Sheet": "徐钢", "最新清单厂家Sheet": "徐刚", "状态": "待确认匹配", "说明": ""},
                {"项目文件Sheet": "闽源", "最新清单厂家Sheet": "闽源", "状态": "待确认匹配", "说明": ""},
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    result = subprocess.run(
        [
            sys.executable,
            str(_script_path()),
            "--image-mapping-json",
            str(mapping),
            "--image-match",
            "徐钢=徐刚",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    assert result.returncode == 0
    rows = json.loads(mapping.read_text(encoding="utf-8"))
    assert rows[0]["状态"] == "已确认匹配"
    assert rows[0]["最新清单厂家Sheet"] == "徐刚"
    assert rows[1]["状态"] == "已确认匹配"

    summary = json.loads(result.stdout)
    assert summary["image"]["pending_count"] == 0

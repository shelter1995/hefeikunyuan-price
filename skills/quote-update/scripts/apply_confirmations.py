from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def _read_rows(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, list):
        raise SystemExit(f"mapping JSON must be a list: {path}")
    return [x for x in payload if isinstance(x, dict)]


def _write_rows(path: Path, rows: list[dict[str, Any]]) -> None:
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


def apply_web_confirmations(
    mapping_json: Path,
    confirmed_matches: dict[str, str],
    skip_new_mills: bool,
) -> dict[str, int]:
    rows = _read_rows(mapping_json)
    for row in rows:
        project_sheet = str(row.get("项目文件Sheet") or "").strip()
        source_sheet = str(row.get("最新清单厂家Sheet") or "").strip()
        status = str(row.get("状态") or "").strip()
        if project_sheet in confirmed_matches:
            row["最新清单厂家Sheet"] = confirmed_matches[project_sheet]
            row["状态"] = "已确认匹配"
            row["说明"] = "用户通过官方确认脚本确认"
        elif status.startswith("待确认匹配") and project_sheet and source_sheet:
            row["状态"] = "已确认匹配"
            row["说明"] = "用户通过官方确认脚本确认"
        elif skip_new_mills and status.startswith("待确认(新厂家)"):
            row["状态"] = "已确认不更新"
            row["说明"] = "用户通过官方确认脚本确认不更新"
    _write_rows(mapping_json, rows)
    pending_count = sum(1 for row in rows if str(row.get("状态") or "").startswith("待确认"))
    return {"total_count": len(rows), "pending_count": pending_count}


def _parse_pairs(items: list[str]) -> dict[str, str]:
    pairs: dict[str, str] = {}
    for item in items:
        if "=" not in item:
            raise SystemExit(f"确认项必须是 项目Sheet=来源Sheet 格式：{item}")
        left, right = item.split("=", 1)
        pairs[left.strip()] = right.strip()
    return pairs


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply quote-update mapping confirmations.")
    parser.add_argument("--web-mapping-json", required=True)
    parser.add_argument("--web-match", action="append", default=[], help="项目Sheet=来源Sheet")
    parser.add_argument("--skip-new-mills", action="store_true")
    args = parser.parse_args()
    summary = apply_web_confirmations(
        Path(args.web_mapping_json),
        confirmed_matches=_parse_pairs(args.web_match),
        skip_new_mills=args.skip_new_mills,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if summary["pending_count"] != 0:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

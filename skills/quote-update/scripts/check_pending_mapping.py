from __future__ import annotations

import argparse
import json
from pathlib import Path


def _count_pending(path: Path) -> tuple[int, int]:
    if not path.exists():
        return 0, 0
    rows = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(rows, list):
        return 0, 0
    total = len(rows)
    pending = 0
    for row in rows:
        if not isinstance(row, dict):
            continue
        if str(row.get("状态") or "").strip().startswith("待确认"):
            pending += 1
    return total, pending


def main() -> int:
    p = argparse.ArgumentParser(description="Check pending rows in mapping json.")
    p.add_argument("--mapping-json", required=True, help="对照表json路径")
    args = p.parse_args()
    path = Path(args.mapping_json)
    total, pending = _count_pending(path)
    print(f"Mapping: {path}")
    print(f"Total Rows: {total}")
    print(f"Pending Rows: {pending}")
    return 1 if pending > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())

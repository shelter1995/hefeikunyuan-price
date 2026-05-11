from __future__ import annotations

from pathlib import Path
from typing import Any

from .xlsx_utils import load_workbook_safe


def _same_value(left: Any, right: Any) -> bool:
    if left == right:
        return True
    if left is None or right is None:
        return False
    return str(left).strip() == str(right).strip()


def audit_image_doc_updates(project_excel: Path, updates: list[dict[str, Any]]) -> dict[str, Any]:
    wb = load_workbook_safe(project_excel, data_only=False)
    mismatches: list[dict[str, Any]] = []
    checked = 0

    for row in updates:
        sheet_name = str(row.get("项目文件Sheet") or "").strip()
        if not sheet_name or sheet_name not in wb.sheetnames:
            mismatches.append({"sheet": sheet_name, "cell": "", "expected": "sheet exists", "actual": "missing"})
            continue
        ws = wb[sheet_name]
        for key, cell in (("H3_new", "H3"), ("H4_new", "H4")):
            expected = row.get(key)
            if expected is None or (isinstance(expected, str) and "保留原值" in expected):
                continue
            checked += 1
            actual = ws[cell].value
            if not _same_value(actual, expected):
                mismatches.append({"sheet": sheet_name, "cell": cell, "expected": expected, "actual": actual})

    wb.close()
    return {
        "status": "ok" if not mismatches else "failed",
        "checked_count": checked,
        "mismatch_count": len(mismatches),
        "mismatches": mismatches,
    }

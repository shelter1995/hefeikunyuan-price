# MiniMax Offline Quote Pipeline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the legacy OCR-driven offline quote path with a MiniMax vision structured-extraction pipeline, then fix the known location parsing, inventory color, report, and dependency gaps so the project can run end to end with auditable results.

**Architecture:** Keep the existing `ocr_price.pipeline` orchestration, but make MiniMax vision the only image/PDF extraction engine and move business validation into deterministic Python code. Treat model output as untrusted structured input: normalize it, validate it, block unsafe rows for confirmation, then write prices and inventory colors to Excel.

**Tech Stack:** Python 3.14, openpyxl, Playwright, requests, pytest, MiniMax VLM API, JSON/CSV artifacts, Excel `.xlsx` workbooks.

---

## File Structure

- Modify `requirements-ocr.txt`: declare all runtime and test dependencies used by this project.
- Modify `.env.example`: document `MINIMAX_API_KEY` without exposing a real key.
- Modify `ocr_price/web_price.py`: fix ordered location parsing and preserve existing web writeback behavior.
- Modify `ocr_price/minimax_vision.py`: keep MiniMax as the image/PDF engine, remove hardcoded city conversion, and preserve inventory payloads.
- Create `ocr_price/offline_validation.py`: deterministic validation for MiniMax/text-extracted offline quote JSON.
- Modify `ocr_price/writeback_image_doc.py`: load only validated source prices, carry validation issues into reports, and keep supplemental text merge rules.
- Modify `ocr_price/inventory.py`: make inventory color application accept confirmed mapping data, clear quote-sheet colors, and save reliably.
- Modify `ocr_price/cli.py`: make image/PDF processing MiniMax-only and label outputs consistently.
- Modify `ocr_price/pipeline.py`: wire the new validation and reporting flow; save both JSON and Markdown reports.
- Create `ocr_price/reporting.py`: render the final user-facing Markdown report from pipeline result JSON.
- Create `tests/test_location_parse.py`: regression tests for file-name location parsing.
- Create `tests/test_minimax_conversion.py`: tests for dynamic target-city conversion and inventory preservation.
- Create `tests/test_offline_validation.py`: tests for price/status validation and blockable issues.
- Create `tests/test_inventory_writeback.py`: tests for inventory coloring on a generated workbook.
- Create `tests/test_reporting.py`: tests for final Markdown report shape.
- Modify `README.md` and `doc/线下报价识别与库存标注标准流程.md`: update the documented implementation path.

The repository currently has no `.git` directory. For every commit step below, run the listed `git` commands only when this project is inside a Git repository; otherwise record the changed files in the task notes and continue.

---

### Task 1: Dependency and Environment Documentation

**Files:**
- Modify: `requirements-ocr.txt`
- Modify: `.env.example`

- [ ] **Step 1: Update dependency list**

Replace `requirements-ocr.txt` with:

```text
requests>=2.31.0
openpyxl>=3.1.2
playwright>=1.45.0
pytest>=8.0.0
```

- [ ] **Step 2: Update environment example**

Replace `.env.example` with:

```text
# MiniMax vision API, used for offline quote images/PDFs.
MINIMAX_API_KEY=replace_with_your_minimax_key

# PaddleOCR is no longer used by the main offline quote pipeline.
# Keep these only if you run legacy diagnostic scripts manually.
PADDLEOCR_BASE_URL=https://4au0y5nbbamev9w3.aistudio-app.com/layout-parsing
PADDLEOCR_API_KEY=replace_with_your_token
PADDLEOCR_AUTH_SCHEME=token
```

- [ ] **Step 3: Verify dependency import availability**

Run:

```powershell
python - <<'PY'
import importlib.util
for name in ["requests", "openpyxl", "playwright", "pytest"]:
    print(name, "ok" if importlib.util.find_spec(name) else "missing")
PY
```

Expected: all four lines end with `ok`.

- [ ] **Step 4: Commit or record changes**

Run if Git is available:

```powershell
git add requirements-ocr.txt .env.example
git commit -m "chore: document minimax quote dependencies"
```

---

### Task 2: Fix Ordered Project Location Parsing

**Files:**
- Test: `tests/test_location_parse.py`
- Modify: `ocr_price/web_price.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_location_parse.py`:

```python
from ocr_price.web_price import _parse_location


def test_parse_dual_locations_in_filename_order():
    name = "安徽合肥-安徽蚌埠-蚌投(春和苑）询价表2026.4.24.xlsx"
    assert _parse_location(name) == ("安徽合肥", "安徽蚌埠")


def test_parse_single_location_uses_same_location_for_both_flows():
    name = "安徽蚌埠-蚌投项目询价表2026.4.24.xlsx"
    assert _parse_location(name) == ("安徽蚌埠", "安徽蚌埠")


def test_parse_overlapping_short_city_does_not_duplicate_long_city():
    name = "江苏南通-项目询价表.xlsx"
    assert _parse_location(name) == ("江苏南通", "江苏南通")
```

- [ ] **Step 2: Run tests and verify the first test fails**

Run:

```powershell
pytest tests/test_location_parse.py -q
```

Expected before implementation: `test_parse_dual_locations_in_filename_order` fails because the current parser returns `("安徽蚌埠", "安徽合肥")`.

- [ ] **Step 3: Implement ordered parsing**

In `ocr_price/web_price.py`, replace `_parse_location` with:

```python
def _parse_location(filename: str) -> tuple[str, str]:
    """
    Parse locations from the project file name in visual filename order.

    Returns (web_location, image_doc_location). If only one location is present,
    both flows use that location.
    """
    stem = Path(filename).stem
    matches: list[tuple[int, int, str]] = []

    for keyword in sorted(LOCATION_KEYWORDS, key=len, reverse=True):
        for match in re.finditer(re.escape(keyword), stem):
            start, end = match.span()
            overlaps = any(not (end <= s or start >= e) for s, e, _ in matches)
            if not overlaps:
                matches.append((start, end, keyword))

    if not matches:
        raise WebPriceError(f"无法从文件名识别地点: {filename}")

    ordered = [value for _, _, value in sorted(matches, key=lambda item: item[0])]
    if len(ordered) >= 2:
        return ordered[0], ordered[1]
    return ordered[0], ordered[0]
```

- [ ] **Step 4: Verify location tests pass**

Run:

```powershell
pytest tests/test_location_parse.py -q
```

Expected: `3 passed`.

- [ ] **Step 5: Commit or record changes**

Run if Git is available:

```powershell
git add tests/test_location_parse.py ocr_price/web_price.py
git commit -m "fix: parse quote locations in filename order"
```

---

### Task 3: Make MiniMax Conversion Dynamic and Preserve Inventory

**Files:**
- Test: `tests/test_minimax_conversion.py`
- Modify: `ocr_price/minimax_vision.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_minimax_conversion.py`:

```python
from ocr_price.minimax_vision import convert_to_ocr_format


def test_convert_to_ocr_format_uses_requested_target_city():
    vision = {
        "_source": {"image_path": "线下报价/测试报价.png"},
        "厂家名称": "测试钢厂",
        "报价日期": "2026-05-11",
        "蚌埠": {
            "螺纹": 3180,
            "盘螺": 3450,
            "螺纹为电议": False,
            "盘螺为电议": False,
        },
        "库存情况": [
            {"规格": "9米螺纹18", "状态": "充足", "原始描述": "18"},
        ],
    }

    result = convert_to_ocr_format(vision, target_cities=["蚌埠"])

    assert result["meta"]["target_location"] == "蚌埠"
    assert result["records"] == [
        {
            "region_title": None,
            "location": "蚌埠",
            "rebar_price": 3180,
            "coil_price": 3450,
            "rebar_raw": "3180",
            "coil_raw": "3450",
            "group_index": None,
            "source_row_index": None,
            "header_row_index": None,
        }
    ]
    assert result["_vision_result"]["库存情况"][0]["状态"] == "充足"
```

- [ ] **Step 2: Run test and verify it fails**

Run:

```powershell
pytest tests/test_minimax_conversion.py -q
```

Expected before implementation: failure because `convert_to_ocr_format` currently has no `target_cities` parameter and loops over hardcoded `["合肥", "蚌埠"]`.

- [ ] **Step 3: Update conversion signature and city loop**

In `ocr_price/minimax_vision.py`, replace `convert_to_ocr_format` with:

```python
def convert_to_ocr_format(
    vision_result: dict[str, Any],
    target_cities: list[str] | None = None,
) -> dict[str, Any]:
    """Convert MiniMax vision result to the existing JSON format for compatibility."""
    if target_cities is None:
        target_cities = ["合肥", "蚌埠"]

    source = vision_result.get("_source", {})
    image_path = source.get("image_path", "")

    records: list[dict[str, Any]] = []
    for city_key in target_cities:
        city_data = vision_result.get(city_key)
        if not isinstance(city_data, dict):
            continue
        rebar = city_data.get("螺纹")
        coil = city_data.get("盘螺")
        rebar_elec = bool(city_data.get("螺纹为电议", False))
        coil_elec = bool(city_data.get("盘螺为电议", False))

        if rebar is None and coil is None and not rebar_elec and not coil_elec:
            continue

        records.append({
            "region_title": None,
            "location": city_key,
            "rebar_price": rebar,
            "coil_price": coil,
            "rebar_raw": "电议" if rebar_elec else str(rebar) if rebar is not None else None,
            "coil_raw": "电议" if coil_elec else str(coil) if coil is not None else None,
            "group_index": None,
            "source_row_index": None,
            "header_row_index": None,
        })

    return {
        "meta": {
            "input_file": image_path,
            "provider": "minimax_vision",
            "target_location": ",".join(target_cities),
            "generated_at": __import__("datetime").datetime.now().isoformat(timespec="seconds"),
            "record_count": len(records),
        },
        "company": vision_result.get("厂家名称"),
        "quote_date": vision_result.get("报价日期"),
        "header_row_index": None,
        "group_count": None,
        "records": records,
        "_vision_result": vision_result,
    }
```

- [ ] **Step 4: Pass target cities through the public API**

In `analyze_quote_image_to_ocr_format`, change the return statement to:

```python
    return convert_to_ocr_format(vision_result, target_cities=target_cities)
```

- [ ] **Step 5: Verify conversion tests pass**

Run:

```powershell
pytest tests/test_minimax_conversion.py -q
```

Expected: `1 passed`.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add tests/test_minimax_conversion.py ocr_price/minimax_vision.py
git commit -m "fix: convert minimax quotes for requested cities"
```

---

### Task 4: Add Deterministic Offline Quote Validation

**Files:**
- Create: `ocr_price/offline_validation.py`
- Test: `tests/test_offline_validation.py`

- [ ] **Step 1: Write failing validation tests**

Create `tests/test_offline_validation.py`:

```python
from ocr_price.offline_validation import validate_offline_quote_payload


def test_valid_payload_has_no_blocking_issues():
    payload = {
        "company": "测试钢厂",
        "quote_date": "2026-05-11",
        "records": [
            {"location": "蚌埠", "rebar_price": 3180, "coil_price": 3450},
        ],
        "_vision_result": {
            "库存情况": [
                {"规格": "9米螺纹18", "状态": "充足", "原始描述": "18"},
                {"规格": "9米螺纹20", "状态": "告警", "原始描述": "20配"},
            ]
        },
    }

    result = validate_offline_quote_payload(payload, target_location="蚌埠")

    assert result["block_writeback"] is False
    assert result["issues"] == []


def test_out_of_range_price_blocks_writeback():
    payload = {
        "company": "测试钢厂",
        "records": [
            {"location": "蚌埠", "rebar_price": 180, "coil_price": 3450},
        ],
    }

    result = validate_offline_quote_payload(payload, target_location="蚌埠")

    assert result["block_writeback"] is True
    assert result["issues"][0]["code"] == "price_out_of_range"


def test_missing_target_location_blocks_writeback():
    payload = {
        "company": "测试钢厂",
        "records": [
            {"location": "合肥", "rebar_price": 3180, "coil_price": 3450},
        ],
    }

    result = validate_offline_quote_payload(payload, target_location="蚌埠")

    assert result["block_writeback"] is True
    assert result["issues"][0]["code"] == "missing_target_location"


def test_invalid_inventory_status_blocks_writeback():
    payload = {
        "company": "测试钢厂",
        "records": [
            {"location": "蚌埠", "rebar_price": 3180, "coil_price": 3450},
        ],
        "_vision_result": {
            "库存情况": [
                {"规格": "9米螺纹18", "状态": "很多", "原始描述": "很多"},
            ]
        },
    }

    result = validate_offline_quote_payload(payload, target_location="蚌埠")

    assert result["block_writeback"] is True
    assert result["issues"][0]["code"] == "invalid_inventory_status"
```

- [ ] **Step 2: Run tests and verify import failure**

Run:

```powershell
pytest tests/test_offline_validation.py -q
```

Expected before implementation: failure with `ModuleNotFoundError: No module named 'ocr_price.offline_validation'`.

- [ ] **Step 3: Implement validation module**

Create `ocr_price/offline_validation.py`:

```python
from __future__ import annotations

import re
from typing import Any


VALID_INVENTORY_STATUSES = {"充足", "告警", "缺货"}
MIN_PRICE = 2000
MAX_PRICE = 6000


def _norm_location(value: str) -> str:
    text = re.sub(r"\s+", "", value or "")
    for token in ("省", "市", "地区", "市场", "区域", "报价"):
        text = text.replace(token, "")
    return text


def _price_issue(value: Any, field: str, location: str) -> dict[str, str] | None:
    if value is None:
        return None
    if not isinstance(value, int):
        return {
            "code": "price_not_integer",
            "field": field,
            "location": location,
            "message": f"{location}{field}不是整数价格: {value}",
        }
    if value < MIN_PRICE or value > MAX_PRICE:
        return {
            "code": "price_out_of_range",
            "field": field,
            "location": location,
            "message": f"{location}{field}价格超出合理范围: {value}",
        }
    return None


def _inventory_items(payload: dict[str, Any]) -> list[dict[str, Any]]:
    vision = payload.get("_vision_result")
    if isinstance(vision, dict) and isinstance(vision.get("库存情况"), list):
        return [x for x in vision["库存情况"] if isinstance(x, dict)]
    inventory = payload.get("inventory")
    if isinstance(inventory, list):
        return [x for x in inventory if isinstance(x, dict)]
    return []


def validate_offline_quote_payload(
    payload: dict[str, Any],
    target_location: str,
) -> dict[str, Any]:
    issues: list[dict[str, str]] = []
    records = payload.get("records") if isinstance(payload.get("records"), list) else []
    target = _norm_location(target_location)

    matched_records = []
    for rec in records:
        if not isinstance(rec, dict):
            continue
        location = _norm_location(str(rec.get("location") or ""))
        if target and target not in location and location not in target:
            continue
        matched_records.append(rec)
        for field in ("rebar_price", "coil_price"):
            issue = _price_issue(rec.get(field), field, str(rec.get("location") or target_location))
            if issue:
                issues.append(issue)

    if target and not matched_records:
        issues.append({
            "code": "missing_target_location",
            "field": "records",
            "location": target_location,
            "message": f"未识别到目标地点价格: {target_location}",
        })

    for item in _inventory_items(payload):
        status = str(item.get("状态") or "").strip()
        if status and status not in VALID_INVENTORY_STATUSES:
            issues.append({
                "code": "invalid_inventory_status",
                "field": "库存情况",
                "location": target_location,
                "message": f"库存状态无效: {status}",
            })

    return {
        "block_writeback": any(issue["code"] in {
            "price_not_integer",
            "price_out_of_range",
            "missing_target_location",
            "invalid_inventory_status",
        } for issue in issues),
        "issues": issues,
    }
```

- [ ] **Step 4: Verify validation tests pass**

Run:

```powershell
pytest tests/test_offline_validation.py -q
```

Expected: `4 passed`.

- [ ] **Step 5: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/offline_validation.py tests/test_offline_validation.py
git commit -m "feat: validate offline quote extraction results"
```

---

### Task 5: Apply Validation Before Offline Price Writeback

**Files:**
- Modify: `ocr_price/writeback_image_doc.py`
- Test: extend `tests/test_offline_validation.py`

- [ ] **Step 1: Add source-loading test for invalid JSON**

Append to `tests/test_offline_validation.py`:

```python
import json
from pathlib import Path

from ocr_price.writeback_image_doc import load_source_prices


def test_load_source_prices_skips_invalid_price_payload(tmp_path: Path):
    source = tmp_path / "ocr价格提取_测试钢厂.json"
    source.write_text(
        json.dumps(
            {
                "meta": {"input_file": "线下报价/测试钢厂.png"},
                "company": "测试钢厂",
                "records": [
                    {"location": "蚌埠", "rebar_price": 180, "coil_price": 3450},
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    assert load_source_prices([source], location="蚌埠") == []
```

- [ ] **Step 2: Run the new test and verify it fails**

Run:

```powershell
pytest tests/test_offline_validation.py::test_load_source_prices_skips_invalid_price_payload -q
```

Expected before implementation: failure because invalid price payload is accepted.

- [ ] **Step 3: Import and apply validation**

In `ocr_price/writeback_image_doc.py`, add:

```python
from .offline_validation import validate_offline_quote_payload
```

In `_load_single_source_price`, immediately after reading `data`, add:

```python
    validation = validate_offline_quote_payload(data, target_location=location)
    if validation["block_writeback"]:
        return None
```

- [ ] **Step 4: Make skipped invalid sources visible in prepare reports**

In `load_source_prices`, collect invalid source paths for reporting by changing the exception-swallowing loop to store a module-level report field is not necessary. Keep the implementation narrow: invalid payloads are excluded from writeback and visible through validation tests. The final report task will surface skipped rows from apply reports.

- [ ] **Step 5: Verify source validation tests pass**

Run:

```powershell
pytest tests/test_offline_validation.py -q
```

Expected: all tests in the file pass.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/writeback_image_doc.py tests/test_offline_validation.py
git commit -m "fix: block invalid offline quote sources"
```

---

### Task 6: Fix Inventory Color Writeback and Mapping Support

**Files:**
- Modify: `ocr_price/inventory.py`
- Test: `tests/test_inventory_writeback.py`
- Modify: `ocr_price/writeback_image_doc.py`

- [ ] **Step 1: Write failing inventory test**

Create `tests/test_inventory_writeback.py`:

```python
from pathlib import Path

from openpyxl import Workbook, load_workbook

from ocr_price.inventory import InventoryItem, apply_inventory_to_project


def _make_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    ws["A1"] = "合肥鲲源贸易有限公司钢材报价单"
    ws["P1"] = "徐钢"
    ws["A8"] = "等级"
    ws["B8"] = "规格"
    ws["C8"] = "长度（米）"
    ws["D8"] = "材质"
    ws["A12"] = "抗震三级钢"
    ws["B12"] = "18"
    ws["C12"] = "9"
    ws["D12"] = "HRB400E"
    wb.save(path)


def test_apply_inventory_to_project_colors_matching_cell(tmp_path: Path):
    workbook = tmp_path / "quote.xlsx"
    _make_workbook(workbook)

    result = apply_inventory_to_project(
        project_excel=workbook,
        mill_inventories={
            "徐钢": [
                InventoryItem(
                    product="螺纹",
                    spec="18",
                    length="9",
                    material="HRB400E",
                    status="缺货",
                    note="无货",
                )
            ]
        },
        sheet_name="报价表",
    )

    wb = load_workbook(workbook)
    ws = wb["报价表"]
    assert result["status"] == "ok"
    assert result["applied_count"] == 1
    assert ws["P12"].fill.fgColor.rgb == "00FF0000"
    wb.close()
```

- [ ] **Step 2: Run inventory test**

Run:

```powershell
pytest tests/test_inventory_writeback.py -q
```

Expected: test may pass before the signature fix. Keep it as a regression test for the core color behavior.

- [ ] **Step 3: Extend function signature for mapping JSON**

In `ocr_price/inventory.py`, change the function signature to:

```python
def apply_inventory_to_project(
    project_excel: Path,
    mill_inventories: dict[str, list[InventoryItem]],
    sheet_name: str = "报价表",
    mapping_json_path: Path | None = None,
    clear_existing_colors: bool = True,
) -> dict[str, Any]:
```

- [ ] **Step 4: Add mapping loader helper**

Add this helper above `apply_inventory_to_project`:

```python
def _load_confirmed_mapping(mapping_json_path: Path | None) -> dict[str, str]:
    if not mapping_json_path or not mapping_json_path.exists():
        return {}
    import json

    rows = json.loads(mapping_json_path.read_text(encoding="utf-8"))
    if not isinstance(rows, list):
        return {}
    mapping: dict[str, str] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        status = str(row.get("状态") or "").strip()
        sheet = str(row.get("项目文件Sheet") or "").strip()
        source = str(row.get("最新清单厂家Sheet") or "").strip()
        if status == "已确认匹配" and sheet and source:
            mapping[source] = sheet
    return mapping
```

- [ ] **Step 5: Clear existing fills before applying inventory colors**

Inside `apply_inventory_to_project`, after `ws = wb[sheet_name]`, add:

```python
    if clear_existing_colors:
        empty_fill = PatternFill(fill_type=None)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.fill and cell.fill.fill_type:
                    cell.fill = empty_fill
```

- [ ] **Step 6: Use confirmed mapping for mill names**

Inside `apply_inventory_to_project`, before looping `for mill_name, items in mill_inventories.items():`, add:

```python
    confirmed_mapping = _load_confirmed_mapping(mapping_json_path)
```

Then change the target mill name resolution at the start of that loop to:

```python
    for mill_name, items in mill_inventories.items():
        mapped_target = confirmed_mapping.get(mill_name, mill_name)
        target_col = None
        for mapped_mill, col in mill_to_col.items():
            if _mill_match(mapped_mill, mapped_target) or _mill_match(mapped_mill, mill_name):
                target_col = col
                break
```

- [ ] **Step 7: Keep writeback caller unchanged after signature fix**

In `ocr_price/writeback_image_doc.py`, leave the existing call as:

```python
            inventory_report = apply_inventory_to_project(
                project_excel=project_excel,
                mill_inventories=mill_inventories,
                sheet_name="报价表",
                mapping_json_path=mapping_json_path,
            )
```

This call becomes valid after Step 3.

- [ ] **Step 8: Verify inventory tests pass**

Run:

```powershell
pytest tests/test_inventory_writeback.py -q
```

Expected: `1 passed`.

- [ ] **Step 9: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/inventory.py ocr_price/writeback_image_doc.py tests/test_inventory_writeback.py
git commit -m "fix: apply inventory colors with confirmed mappings"
```

---

### Task 7: Remove Legacy OCR From the Main Offline Image Path

**Files:**
- Modify: `ocr_price/cli.py`
- Modify: `ocr_price/writeback_image_doc.py`
- Modify: `README.md`

- [ ] **Step 1: Update CLI description and provider labels**

In `ocr_price/cli.py`, change the parser description to:

```python
description="Extract steel quote prices and inventory from text files or MiniMax vision image/PDF analysis."
```

Change the `--provider` help text to:

```python
help="Processing mode. Text files use text parser; image/PDF files use MiniMax VLM vision.",
```

- [ ] **Step 2: Remove Paddle fallback imports from offline image text reading**

In `ocr_price/writeback_image_doc.py`, replace `_read_source_text` with:

```python
def _read_source_text(input_file: str, ocr_json: dict[str, Any] | None = None) -> str | None:
    path = Path(input_file)
    if not path.exists():
        return None
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return path.read_text(encoding="utf-8", errors="ignore")
    if suffix in {".jpg", ".jpeg", ".png", ".pdf", ".webp"} and ocr_json:
        vision = ocr_json.get("_vision_result")
        if isinstance(vision, dict) and isinstance(vision.get("库存情况"), list):
            lines = []
            for item in vision["库存情况"]:
                if isinstance(item, dict):
                    spec = str(item.get("规格") or "").strip()
                    status = str(item.get("状态") or "").strip()
                    note = str(item.get("原始描述") or "").strip()
                    if spec:
                        lines.append(f"{spec}（{note or status}）" if note or status else spec)
            return "\n".join(lines) if lines else None
    return None
```

- [ ] **Step 3: Verify no main module imports Paddle**

Run:

```powershell
rg -n "paddle_api|PaddleOCR" ocr_price
```

Expected after this task: matches may remain in `ocr_price/paddle_api.py` and legacy test scripts, but not in `ocr_price/cli.py`, `ocr_price/pipeline.py`, or `ocr_price/writeback_image_doc.py`.

- [ ] **Step 4: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/cli.py ocr_price/writeback_image_doc.py README.md
git commit -m "refactor: use minimax as offline image extraction path"
```

---

### Task 8: Generate Final Markdown Reports From Pipeline Results

**Files:**
- Create: `ocr_price/reporting.py`
- Test: `tests/test_reporting.py`
- Modify: `ocr_price/pipeline.py`

- [ ] **Step 1: Write report rendering test**

Create `tests/test_reporting.py`:

```python
from ocr_price.reporting import render_single_report_markdown


def test_render_single_report_markdown_contains_required_sections():
    result = {
        "project": "项目报价/测试项目.xlsx",
        "mode": "both",
        "started_at": "2026-05-11T09:00:00",
        "ended_at": "2026-05-11T09:01:00",
        "status": "ok",
        "web": {
            "status": "ok",
            "apply_summary": {
                "updated_count": 1,
                "skipped_count": 0,
                "updated_items": [
                    {
                        "项目文件Sheet": "徐钢",
                        "G1": {"old": None, "new": "网价[2026-05-11]"},
                        "G3": {"old": 3300, "new": 3400},
                        "G4": {"old": 3100, "new": 3200},
                    }
                ],
                "skipped_items": [],
            },
        },
        "image_doc": {
            "status": "ok",
            "apply_summary": {
                "updated_count": 1,
                "skipped_count": 0,
                "updated_items": [
                    {
                        "项目文件Sheet": "闽源",
                        "H1": {"old": None, "new": "报价[2026-05-11]"},
                        "H3": {"old": 3320, "new": 3450},
                        "H4": {"old": 3140, "new": 3180},
                    }
                ],
                "skipped_items": [],
            },
        },
    }

    markdown = render_single_report_markdown(result, json_report_path="运行产物/report.json")

    assert "项目报价单文件更新完成" in markdown
    assert "网价更新" in markdown
    assert "图片/文档价更新" in markdown
    assert "徐钢" in markdown
    assert "闽源" in markdown
    assert "运行产物/report.json" in markdown
```

- [ ] **Step 2: Run test and verify import failure**

Run:

```powershell
pytest tests/test_reporting.py -q
```

Expected before implementation: `ModuleNotFoundError`.

- [ ] **Step 3: Implement reporting module**

Create `ocr_price/reporting.py`:

```python
from __future__ import annotations

from pathlib import Path
from typing import Any


def _change(cell: dict[str, Any] | None) -> str:
    if not isinstance(cell, dict):
        return "空 → 空"
    old = cell.get("old")
    new = cell.get("new")
    return f"{old if old is not None else '空'} → {new if new is not None else '空'}"


def _section(title: str, prefix: str, summary: dict[str, Any]) -> list[str]:
    updated = summary.get("updated_items") or []
    skipped = summary.get("skipped_items") or []
    lines = [f"### {title}", ""]
    lines.append(f"**已更新（{len(updated)}家）**：")
    lines.append("")
    if prefix == "G":
        lines.extend([
            "| 序号 | 厂家 | G1 | G3(盘螺) | G4(螺纹) |",
            "|------|------|-----|----------|----------|",
        ])
        for idx, row in enumerate(updated, 1):
            lines.append(
                f"| {idx} | {row.get('项目文件Sheet', '')} | "
                f"{_change(row.get('G1'))} | {_change(row.get('G3'))} | {_change(row.get('G4'))} |"
            )
    else:
        lines.extend([
            "| 序号 | 厂家 | H1 | H3(盘螺) | H4(螺纹) |",
            "|------|------|-----|----------|----------|",
        ])
        for idx, row in enumerate(updated, 1):
            lines.append(
                f"| {idx} | {row.get('项目文件Sheet', '')} | "
                f"{_change(row.get('H1'))} | {_change(row.get('H3'))} | {_change(row.get('H4'))} |"
            )
    lines.append("")
    lines.append(f"**未更新（{len(skipped)}家）**：")
    if skipped:
        for row in skipped:
            lines.append(f"- {row.get('项目文件Sheet', '')} - {row.get('原因', '')}")
    else:
        lines.append("- 无")
    lines.append("")
    return lines


def render_single_report_markdown(result: dict[str, Any], json_report_path: str) -> str:
    project_name = Path(str(result.get("project") or "")).name
    started = str(result.get("started_at") or "")
    ended = str(result.get("ended_at") or "")
    mode = str(result.get("mode") or "")
    status = str(result.get("status") or "")

    lines = [
        "## 项目报价单文件更新完成",
        "",
        f"**执行时间**：{started} 至 {ended}",
        f"**更新模式**：{mode}",
        f"**执行结果**：{status}",
        "",
        f"## 一、{project_name}",
        "",
    ]

    web = result.get("web")
    if isinstance(web, dict) and web.get("status") == "ok":
        lines.extend(_section("1. 网价更新（G1/G3/G4）", "G", web.get("apply_summary") or {}))

    image_doc = result.get("image_doc")
    if isinstance(image_doc, dict) and image_doc.get("status") == "ok":
        lines.extend(_section("2. 图片/文档价更新（H1/H3/H4）", "H", image_doc.get("apply_summary") or {}))

    lines.extend([
        "---",
        "",
        "## 报告文件",
        f"完整JSON报告已保存至：`{json_report_path}`",
        "",
    ])
    return "\n".join(lines)
```

- [ ] **Step 4: Wire Markdown report into single pipeline command**

In `ocr_price/pipeline.py`, import:

```python
from .reporting import render_single_report_markdown
```

After `_json_write(report_out, result)` in the single-command branch, add:

```python
        markdown_report = render_single_report_markdown(result, json_report_path=str(report_out))
        markdown_out = report_out.with_suffix(".md")
        markdown_out.write_text(markdown_report, encoding="utf-8")
        print(f"MarkdownReport: {markdown_out}")
```

- [ ] **Step 5: Verify reporting tests pass**

Run:

```powershell
pytest tests/test_reporting.py -q
```

Expected: `1 passed`.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/reporting.py ocr_price/pipeline.py tests/test_reporting.py
git commit -m "feat: render quote update markdown reports"
```

---

### Task 9: Run Focused Regression Suite

**Files:**
- No file edits expected.

- [ ] **Step 1: Run fast unit tests**

Run:

```powershell
pytest tests/test_location_parse.py tests/test_minimax_conversion.py tests/test_offline_validation.py tests/test_inventory_writeback.py tests/test_reporting.py -q
```

Expected: all tests pass.

- [ ] **Step 2: Run compile check**

Run:

```powershell
python -m compileall -q ocr_price skills\quote-update\scripts
```

Expected: exit code `0`.

- [ ] **Step 3: Avoid broad external-service tests**

Do not use `pytest -q` as the completion gate for this project because existing root test scripts include browser login and external API calls. Use the focused regression suite above as the required local gate.

---

### Task 10: Update Operator Documentation

**Files:**
- Modify: `README.md`
- Modify: `doc/线下报价识别与库存标注标准流程.md`
- Modify: `修复计划.md`

- [ ] **Step 1: Update README implementation summary**

In `README.md`, update the offline quote section to state:

```markdown
线下报价图片/PDF统一使用 MiniMax VLM 做结构化识别，不再使用传统OCR作为主链路。系统会把模型输出转换为固定JSON，再经过价格范围、目标地点、库存状态等确定性校验。校验失败、新厂家、厂家映射冲突、目标地点缺失、电议等情况会阻断回写并进入人工确认。
```

- [ ] **Step 2: Update standard-flow documentation**

In `doc/线下报价识别与库存标注标准流程.md`, replace references that imply PaddleOCR or generic OCR is the main image path with:

```markdown
图片/PDF处理主链路：MiniMax VLM结构化识别 → Python业务规则校验 → 厂家对照人工确认 → Excel回写。传统OCR仅保留为历史诊断脚本，不参与默认执行链路。
```

- [ ] **Step 3: Update repair plan status**

In `修复计划.md`, add a top section:

```markdown
## 2026-05-11 新实施方向

线下报价图片/PDF主链路改为 MiniMax VLM 结构化识别，不再依赖传统OCR。同步修复地点解析、库存颜色标注函数签名、最终Markdown报告、依赖和环境说明。
```

- [ ] **Step 4: Commit or record changes**

Run if Git is available:

```powershell
git add README.md doc/线下报价识别与库存标注标准流程.md 修复计划.md
git commit -m "docs: document minimax-first quote workflow"
```

---

## Final Verification

- [ ] Run:

```powershell
pytest tests/test_location_parse.py tests/test_minimax_conversion.py tests/test_offline_validation.py tests/test_inventory_writeback.py tests/test_reporting.py -q
```

Expected: all focused tests pass.

- [ ] Run:

```powershell
python -m compileall -q ocr_price skills\quote-update\scripts
```

Expected: exit code `0`.

- [ ] Run a dry offline JSON writeback using a copied project workbook and existing generated/fake JSON sources before touching the real project file:

```powershell
Copy-Item "项目报价\安徽合肥-安徽蚌埠-蚌投(春和苑）询价表2026.4.24.xlsx" "$env:TEMP\quote-test.xlsx"
python -m ocr_price.pipeline single --project "$env:TEMP\quote-test.xlsx" --mode image_doc --image-jsons "$env:TEMP\hefeikunyuan_ocr_minyuan.json" --artifact-dir "$env:TEMP\hefeikunyuan-run"
```

Expected: either `Status: pending_confirmation` with a mapping file to confirm, or `Status: ok` with JSON and Markdown report paths. It must not crash with `mapping_json_path` argument errors.

## Self-Review

- Spec coverage: MiniMax-first image/PDF extraction is covered by Tasks 3, 5, and 7. Deterministic validation is covered by Task 4. Location parsing is covered by Task 2. Inventory color fix is covered by Task 6. Markdown reporting is covered by Task 8. Documentation and dependencies are covered by Tasks 1 and 10.
- Placeholder scan: No task uses incomplete placeholders. Code snippets provide concrete function signatures, tests, and expected commands.
- Type consistency: `convert_to_ocr_format(..., target_cities=...)`, `validate_offline_quote_payload(...)`, and `apply_inventory_to_project(..., mapping_json_path=...)` are introduced before later tasks use them.

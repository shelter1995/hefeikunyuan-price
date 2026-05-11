# Dynamic Offline Price Validation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the fixed `2000-6000` offline quote price range rule with a web-price-referenced deviation check, and add guardrails so lower-capability models cannot accidentally corrupt project quote workbooks during routine updates.

**Architecture:** Keep `ocr_price.offline_validation` responsible for source payload shape, target-location presence, integer parsing, and inventory status checks. Move market-reasonableness checks into `ocr_price.writeback_image_doc.apply_writeback`, because that function has the project workbook, confirmed manufacturer mapping, target sheet name, and current `G3/G4` web reference prices available. Add an operator-safe execution layer: defaultable dry-run, explicit write confirmation, pre/post Excel audit, and scripted entrypoints that lower-capability models can follow without hand-editing workbooks.

**Tech Stack:** Python 3.14, openpyxl, pytest, existing `ocr_price` modules, Excel `.xlsx` project quote files.

---

## File Structure

- Modify `ocr_price/offline_validation.py`: remove normal-market fixed range blocking from source JSON validation; keep integer and obvious non-price checks.
- Modify `ocr_price/writeback_image_doc.py`: add web-reference deviation validation before writing `H3/H4`.
- Modify `ocr_price/pipeline.py`: add dry-run propagation and explicit write-confirmation behavior.
- Modify `skills/quote-update/scripts/run_single.py`: expose safe execution flags for model operators.
- Modify `skills/quote-update/scripts/run_batch.py`: expose safe execution flags for model operators.
- Create `ocr_price/audit.py`: verify workbook values and generated reports after update.
- Modify `ocr_price/reporting.py`: ensure skipped rows display deviation details cleanly.
- Modify `tests/test_offline_validation.py`: update fixed-range expectations.
- Create `tests/test_offline_price_deviation.py`: workbook-level tests for `G3/G4` reference checks.
- Create `tests/test_pipeline_safety.py`: tests for dry-run/no-write behavior.
- Create `tests/test_audit.py`: tests for report-to-workbook audit behavior.
- Modify `README.md`: document the new `1000元/吨` and `20%` thresholds.
- Modify `doc/线下报价识别与库存标注标准流程.md`: document reference-price behavior and no-reference fallback.
- Modify `skills/quote-update/指令模板.md`: give lower-capability models a fixed operation script and stopping rules.
- Modify `skills/quote-update/references/workflow.md`: require dry-run first and audit after write.

The repository currently has no `.git` directory in this workspace. If Git is unavailable, skip commit commands and record changed files in the final implementation notes.

---

### Task 1: Relax Source Payload Price Range Validation

**Files:**
- Modify: `ocr_price/offline_validation.py`
- Modify: `tests/test_offline_validation.py`

- [ ] **Step 1: Replace the fixed-range test**

In `tests/test_offline_validation.py`, replace `test_validate_offline_payload_out_of_range_price` with:

```python
def test_validate_offline_payload_allows_market_price_outside_old_fixed_range() -> None:
    payload = _base_payload()
    payload["records"][0]["rebar_price"] = 6500
    result = validate_offline_payload(payload, target_location="蚌埠")
    assert result.is_valid is True
    assert result.errors == []
```

- [ ] **Step 2: Add an obvious non-price guard test**

Append this test to `tests/test_offline_validation.py`:

```python
def test_validate_offline_payload_rejects_non_price_integer() -> None:
    payload = _base_payload()
    payload["records"][0]["coil_price"] = 86
    result = validate_offline_payload(payload, target_location="蚌埠")
    assert result.is_valid is False
    assert any("below hard minimum" in err for err in result.errors)
```

- [ ] **Step 3: Run tests and confirm current behavior fails**

Run:

```powershell
pytest tests/test_offline_validation.py -q
```

Expected before implementation: the `6500` test fails because current `PRICE_MAX = 6000` rejects it.

- [ ] **Step 4: Change fixed range constants into hard dirty-data guard**

In `ocr_price/offline_validation.py`, replace:

```python
PRICE_MIN = 2000
PRICE_MAX = 6000
```

with:

```python
HARD_PRICE_MIN = 1000
HARD_PRICE_MAX = 10000
```

- [ ] **Step 5: Update validation error logic**

In `validate_offline_payload`, replace the fixed range block:

```python
            if not (PRICE_MIN <= price <= PRICE_MAX):
                errors.append(
                    f"records[{idx}].{field} out of range [{PRICE_MIN}, {PRICE_MAX}]"
                )
```

with:

```python
            if price < HARD_PRICE_MIN:
                errors.append(f"records[{idx}].{field} below hard minimum {HARD_PRICE_MIN}")
                continue
            if price > HARD_PRICE_MAX:
                errors.append(f"records[{idx}].{field} above hard maximum {HARD_PRICE_MAX}")
```

- [ ] **Step 6: Verify source validation tests pass**

Run:

```powershell
pytest tests/test_offline_validation.py -q
```

Expected: all tests pass.

- [ ] **Step 7: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/offline_validation.py tests/test_offline_validation.py
git commit -m "fix: relax offline price source validation"
```

---

### Task 2: Add Web-Referenced Deviation Helpers

**Files:**
- Modify: `ocr_price/writeback_image_doc.py`
- Create: `tests/test_offline_price_deviation.py`

- [ ] **Step 1: Write helper tests**

Create `tests/test_offline_price_deviation.py`:

```python
from ocr_price.writeback_image_doc import (
    PriceDeviationConfig,
    _check_price_deviation,
)


def test_price_deviation_allows_value_within_absolute_and_percent_thresholds():
    result = _check_price_deviation(
        offline_price=3600,
        web_price=3300,
        label="盘螺",
        config=PriceDeviationConfig(abs_tolerance=1000, pct_tolerance=0.20),
    )
    assert result is None


def test_price_deviation_blocks_when_absolute_threshold_exceeded():
    result = _check_price_deviation(
        offline_price=4501,
        web_price=3300,
        label="盘螺",
        config=PriceDeviationConfig(abs_tolerance=1000, pct_tolerance=0.20),
    )
    assert result is not None
    assert "盘螺" in result
    assert "差值=1201" in result


def test_price_deviation_blocks_when_percent_threshold_exceeded():
    result = _check_price_deviation(
        offline_price=4050,
        web_price=3000,
        label="螺纹",
        config=PriceDeviationConfig(abs_tolerance=1000, pct_tolerance=0.20),
    )
    assert result is not None
    assert "螺纹" in result
    assert "35.00%" in result


def test_price_deviation_allows_when_web_reference_missing():
    result = _check_price_deviation(
        offline_price=4050,
        web_price=None,
        label="螺纹",
        config=PriceDeviationConfig(abs_tolerance=1000, pct_tolerance=0.20),
    )
    assert result is None
```

- [ ] **Step 2: Run helper tests and confirm import failure**

Run:

```powershell
pytest tests/test_offline_price_deviation.py -q
```

Expected before implementation: import failure for `PriceDeviationConfig` and `_check_price_deviation`.

- [ ] **Step 3: Add config and numeric parser**

In `ocr_price/writeback_image_doc.py`, add `field` to the dataclass import:

```python
from dataclasses import dataclass, field
```

Then add this code after `SourcePrice`:

```python
@dataclass(frozen=True)
class PriceDeviationConfig:
    abs_tolerance: int = 1000
    pct_tolerance: float = 0.20


def _coerce_price(value: Any) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = str(value).replace(",", "").strip()
    match = re.search(r"(?<!\d)(\d{3,5})(?!\d)", text)
    return int(match.group(1)) if match else None
```

- [ ] **Step 4: Add deviation helper**

Add this function after `_coerce_price`:

```python
def _check_price_deviation(
    offline_price: int | None,
    web_price: Any,
    label: str,
    config: PriceDeviationConfig,
) -> str | None:
    if offline_price is None:
        return None
    reference = _coerce_price(web_price)
    if reference is None or reference <= 0:
        return None

    diff = offline_price - reference
    abs_diff = abs(diff)
    pct_diff = abs_diff / reference
    if abs_diff > config.abs_tolerance or pct_diff > config.pct_tolerance:
        return (
            f"{label}线下价与网价偏差过大："
            f"线下={offline_price}，网价={reference}，"
            f"差值={diff}，偏离={pct_diff:.2%}，"
            f"阈值={config.abs_tolerance}元/{config.pct_tolerance:.0%}"
        )
    return None
```

This plan uses `or`: if either `1000元/吨` or `20%` is exceeded, the row is not auto-written and is reported for manual review.

- [ ] **Step 5: Verify helper tests pass**

Run:

```powershell
pytest tests/test_offline_price_deviation.py -q
```

Expected: `4 passed`.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/writeback_image_doc.py tests/test_offline_price_deviation.py
git commit -m "feat: add offline price deviation helper"
```

---

### Task 3: Block Offline Writeback When Deviation Exceeds Thresholds

**Files:**
- Modify: `ocr_price/writeback_image_doc.py`
- Modify: `tests/test_offline_price_deviation.py`

- [ ] **Step 1: Add workbook-level writeback test**

Append to `tests/test_offline_price_deviation.py`:

```python
import json
from pathlib import Path

from openpyxl import Workbook, load_workbook

from ocr_price.writeback_image_doc import apply_writeback


def _make_project(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    mill = wb.create_sheet("测试钢厂")
    mill["G1"] = "网价[2026-05-11]"
    mill["G3"] = 3300
    mill["G4"] = 3100
    mill["H3"] = 3300
    mill["H4"] = 3100
    wb.save(path)


def _write_json(path: Path, payload: object) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def test_apply_writeback_skips_price_when_deviation_is_too_large(tmp_path: Path):
    project = tmp_path / "project.xlsx"
    source = tmp_path / "source.json"
    mapping = tmp_path / "mapping.json"
    report = tmp_path / "report.json"
    _make_project(project)
    _write_json(
        source,
        {
            "meta": {"input_file": "测试钢厂报价.txt"},
            "quote_date": "2026-05-11",
            "records": [
                {"location": "蚌埠", "coil_price": 4600, "rebar_price": 4500},
            ],
        },
    )
    _write_json(
        mapping,
        [
            {
                "项目文件Sheet": "测试钢厂",
                "最新清单厂家Sheet": "测试钢厂",
                "状态": "已确认匹配",
                "说明": "",
            }
        ],
    )

    result = apply_writeback(
        project_excel=project,
        source_json_paths=[source],
        mapping_json_path=mapping,
        location="蚌埠",
        report_out=report,
    )

    wb = load_workbook(project, data_only=False)
    ws = wb["测试钢厂"]
    assert result["updated_count"] == 0
    assert result["skipped_count"] == 1
    assert "偏差过大" in result["skipped"][0]["原因"]
    assert ws["H3"].value == 3300
    assert ws["H4"].value == 3100
    wb.close()
```

- [ ] **Step 2: Run workbook-level test and confirm current behavior fails**

Run:

```powershell
pytest tests/test_offline_price_deviation.py::test_apply_writeback_skips_price_when_deviation_is_too_large -q
```

Expected before implementation: failure because current code writes `H3/H4` without comparing to `G3/G4`.

- [ ] **Step 3: Add pre-write deviation check**

In `ocr_price/writeback_image_doc.py`, inside `apply_writeback`, after:

```python
        ws = wb[sheet]
        old_h1, old_h3, old_h4 = ws["H1"].value, ws["H3"].value, ws["H4"].value
        new_h1 = f"报价[{src.quote_date}]"
        ws["H1"] = new_h1
```

move `ws["H1"] = new_h1` below the deviation check, and insert:

```python
        deviation_config = PriceDeviationConfig()
        deviation_reasons: list[str] = []
        coil_deviation = _check_price_deviation(
            offline_price=src.coil_price,
            web_price=ws["G3"].value,
            label="盘螺",
            config=deviation_config,
        )
        if coil_deviation:
            deviation_reasons.append(coil_deviation)
        rebar_deviation = _check_price_deviation(
            offline_price=src.rebar_price,
            web_price=ws["G4"].value,
            label="螺纹",
            config=deviation_config,
        )
        if rebar_deviation:
            deviation_reasons.append(rebar_deviation)

        if deviation_reasons:
            skipped.append(
                {
                    "项目文件Sheet": sheet,
                    "来源厂家": source_company,
                    "原因": "；".join(deviation_reasons),
                }
            )
            continue

        ws["H1"] = new_h1
```

This avoids changing `H1` when the row is blocked.

- [ ] **Step 4: Verify workbook-level test passes**

Run:

```powershell
pytest tests/test_offline_price_deviation.py::test_apply_writeback_skips_price_when_deviation_is_too_large -q
```

Expected: `1 passed`.

- [ ] **Step 5: Run full deviation tests**

Run:

```powershell
pytest tests/test_offline_price_deviation.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/writeback_image_doc.py tests/test_offline_price_deviation.py
git commit -m "feat: block offline prices far from web reference"
```

---

### Task 4: Report No-Reference Cases Without Blocking

**Files:**
- Modify: `ocr_price/writeback_image_doc.py`
- Modify: `tests/test_offline_price_deviation.py`

- [ ] **Step 1: Add no-reference writeback test**

Append to `tests/test_offline_price_deviation.py`:

```python
def test_apply_writeback_allows_price_when_web_reference_is_missing(tmp_path: Path):
    project = tmp_path / "project.xlsx"
    source = tmp_path / "source.json"
    mapping = tmp_path / "mapping.json"
    report = tmp_path / "report.json"
    _make_project(project)
    wb = load_workbook(project)
    ws = wb["测试钢厂"]
    ws["G3"] = None
    ws["G4"] = None
    wb.save(project)
    wb.close()

    _write_json(
        source,
        {
            "meta": {"input_file": "测试钢厂报价.txt"},
            "quote_date": "2026-05-11",
            "records": [
                {"location": "蚌埠", "coil_price": 4600, "rebar_price": 4500},
            ],
        },
    )
    _write_json(
        mapping,
        [
            {
                "项目文件Sheet": "测试钢厂",
                "最新清单厂家Sheet": "测试钢厂",
                "状态": "已确认匹配",
                "说明": "",
            }
        ],
    )

    result = apply_writeback(
        project_excel=project,
        source_json_paths=[source],
        mapping_json_path=mapping,
        location="蚌埠",
        report_out=report,
    )

    wb = load_workbook(project, data_only=False)
    ws = wb["测试钢厂"]
    assert result["updated_count"] == 1
    assert result["skipped_count"] == 0
    assert "无网价参考" in result["updates"][0]["备注"]
    assert ws["H3"].value == 4600
    assert ws["H4"].value == 4500
    wb.close()
```

- [ ] **Step 2: Run no-reference test and confirm it fails**

Run:

```powershell
pytest tests/test_offline_price_deviation.py::test_apply_writeback_allows_price_when_web_reference_is_missing -q
```

Expected before implementation: failure because update notes do not mention missing web reference.

- [ ] **Step 3: Add missing-reference note helper**

In `ocr_price/writeback_image_doc.py`, add:

```python
def _missing_reference_notes(ws: Any, src: SourcePrice) -> list[str]:
    notes: list[str] = []
    if src.coil_price is not None and _coerce_price(ws["G3"].value) is None:
        notes.append("盘螺无网价参考")
    if src.rebar_price is not None and _coerce_price(ws["G4"].value) is None:
        notes.append("螺纹无网价参考")
    return notes
```

- [ ] **Step 4: Append missing-reference notes to update remark**

Inside `apply_writeback`, after the deviation check and before writing cells, add:

```python
        reference_notes = _missing_reference_notes(ws, src)
```

Then change:

```python
        partial_note = note
```

to:

```python
        partial_note = note
        if reference_notes:
            partial_note += f"({'；'.join(reference_notes)})"
```

- [ ] **Step 5: Verify no-reference test passes**

Run:

```powershell
pytest tests/test_offline_price_deviation.py::test_apply_writeback_allows_price_when_web_reference_is_missing -q
```

Expected: `1 passed`.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/writeback_image_doc.py tests/test_offline_price_deviation.py
git commit -m "feat: report offline prices without web reference"
```

---

### Task 5: Update Documentation and Final Verification

**Files:**
- Modify: `README.md`
- Modify: `doc/线下报价识别与库存标注标准流程.md`

- [ ] **Step 1: Document thresholds in README**

Add this paragraph near the offline quote rules in `README.md`:

```markdown
线下报价价格校验以项目文件中同厂家 sheet 的网价为动态参考：H3(盘螺) 对比 G3，H4(螺纹) 对比 G4。任一价格与网价的差值超过 1000 元/吨或偏离比例超过 20% 时，不自动回写，报告为“线下价与网价偏差过大”，需要人工确认。若 G3/G4 没有可用网价，允许回写，但报告备注“无网价参考”。
```

- [ ] **Step 2: Document thresholds in standard-flow doc**

Add this paragraph in `doc/线下报价识别与库存标注标准流程.md` under image/doc writeback rules:

```markdown
价格偏差校验：线下报价写入 H3/H4 前，必须读取同一厂家 sheet 的 G3/G4 网价作为参考。盘螺使用 H3 vs G3，螺纹使用 H4 vs G4。任一维度超过 1000 元/吨或 20% 偏离时，本厂家本次线下价不自动写入，进入报告待人工确认。没有 G3/G4 参考价时不阻断，但报告必须标注“无网价参考”。
```

- [ ] **Step 3: Run focused regression suite**

Run:

```powershell
pytest tests/test_offline_validation.py tests/test_offline_price_deviation.py tests/test_reporting.py -q
```

Expected: all tests pass.

- [ ] **Step 4: Run broader focused project suite**

Run:

```powershell
pytest tests/test_location_parse.py tests/test_minimax_conversion.py tests/test_offline_validation.py tests/test_offline_price_deviation.py tests/test_inventory_writeback.py tests/test_reporting.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Run compile check**

Run:

```powershell
python -m compileall -q ocr_price skills\quote-update\scripts tests
```

Expected: exit code `0`.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add README.md doc/线下报价识别与库存标注标准流程.md
git commit -m "docs: document web-referenced offline price validation"
```

---

### Task 6: Add Dry-Run and Explicit Write Confirmation Guardrails

**Files:**
- Modify: `ocr_price/writeback_image_doc.py`
- Modify: `ocr_price/pipeline.py`
- Modify: `skills/quote-update/scripts/run_single.py`
- Modify: `skills/quote-update/scripts/run_batch.py`
- Create: `tests/test_pipeline_safety.py`

- [ ] **Step 1: Write dry-run writeback test**

Create `tests/test_pipeline_safety.py`:

```python
import json
from pathlib import Path

from openpyxl import Workbook, load_workbook

from ocr_price.writeback_image_doc import apply_writeback


def _write_json(path: Path, payload: object) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _make_project(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    mill = wb.create_sheet("测试钢厂")
    mill["G3"] = 3300
    mill["G4"] = 3100
    mill["H3"] = 3300
    mill["H4"] = 3100
    wb.save(path)


def test_apply_writeback_dry_run_does_not_modify_workbook(tmp_path: Path):
    project = tmp_path / "project.xlsx"
    source = tmp_path / "source.json"
    mapping = tmp_path / "mapping.json"
    report = tmp_path / "report.json"
    _make_project(project)
    _write_json(
        source,
        {
            "meta": {"input_file": "测试钢厂报价.txt"},
            "quote_date": "2026-05-11",
            "records": [
                {"location": "蚌埠", "coil_price": 3400, "rebar_price": 3200},
            ],
        },
    )
    _write_json(
        mapping,
        [
            {
                "项目文件Sheet": "测试钢厂",
                "最新清单厂家Sheet": "测试钢厂",
                "状态": "已确认匹配",
                "说明": "",
            }
        ],
    )

    result = apply_writeback(
        project_excel=project,
        source_json_paths=[source],
        mapping_json_path=mapping,
        location="蚌埠",
        report_out=report,
        dry_run=True,
    )

    wb = load_workbook(project)
    ws = wb["测试钢厂"]
    assert result["dry_run"] is True
    assert result["updated_count"] == 1
    assert result["backup_file"] is None
    assert ws["H3"].value == 3300
    assert ws["H4"].value == 3100
    wb.close()
```

- [ ] **Step 2: Run dry-run test and confirm it fails**

Run:

```powershell
pytest tests/test_pipeline_safety.py -q
```

Expected before implementation: failure because `apply_writeback` does not accept `dry_run`.

- [ ] **Step 3: Add `dry_run` to image/doc writeback**

In `ocr_price/writeback_image_doc.py`, change the function signature:

```python
def apply_writeback(
    project_excel: Path,
    source_json_paths: list[Path],
    mapping_json_path: Path,
    location: str,
    report_out: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
```

In the save block, replace:

```python
    if updates:
        backup_dir = project_excel.parent / "备份"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = backup_dir / f"{project_excel.stem}.backup_before_image_doc_write_{stamp}.xlsx"
        shutil.copy2(project_excel, backup_file)
        wb.save(project_excel)
        wb.close()
    else:
        backup_file = None
```

with:

```python
    if dry_run:
        backup_file = None
        wb.close()
    elif updates:
        backup_dir = project_excel.parent / "备份"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = backup_dir / f"{project_excel.stem}.backup_before_image_doc_write_{stamp}.xlsx"
        shutil.copy2(project_excel, backup_file)
        wb.save(project_excel)
        wb.close()
    else:
        backup_file = None
        wb.close()
```

Add `"dry_run": dry_run` to the final report dict.

- [ ] **Step 4: Skip inventory coloring during dry-run**

In `ocr_price/writeback_image_doc.py`, before inventory coloring, add:

```python
    if dry_run:
        inventory_report = {"status": "skipped", "reason": "dry-run模式不修改库存颜色"}
    else:
        inventory_report = None
        try:
            ...
```

Move the existing inventory try/except under the `else` branch.

- [ ] **Step 5: Propagate dry-run through pipeline**

In `ocr_price/pipeline.py`, add a `dry_run: bool = False` parameter to `_image_flow` and `run_single`, then pass it into `apply_image_writeback`:

```python
    apply_report = apply_image_writeback(
        project_excel=project,
        source_json_paths=source_jsons,
        mapping_json_path=confirmed_json if reusable or confirmed_json.exists() else pending_json,
        location=_city(location),
        report_out=apply_report_path,
        dry_run=dry_run,
    )
```

- [ ] **Step 6: Add CLI flags**

In `ocr_price/pipeline.py`, add to both `single` and `batch` parsers:

```python
p_single.add_argument("--dry-run", action="store_true", help="预演流程，只生成报告，不修改项目Excel")
p_single.add_argument("--confirm-write", action="store_true", help="明确允许写入项目Excel")
```

For batch, use `p_batch.add_argument(...)`.

Before running a project, enforce:

```python
if not args.dry_run and not args.confirm_write:
    raise SystemExit("为防止误操作，写入项目Excel必须显式传入 --confirm-write；预演请使用 --dry-run。")
```

Pass `dry_run=args.dry_run` into `run_single`.

- [ ] **Step 7: Expose flags in wrapper scripts**

In `skills/quote-update/scripts/run_single.py` and `skills/quote-update/scripts/run_batch.py`, add parser flags:

```python
p.add_argument("--dry-run", action="store_true", help="预演流程，只生成报告，不修改项目Excel")
p.add_argument("--confirm-write", action="store_true", help="明确允许写入项目Excel")
```

When building `cmd`, append:

```python
if args.dry_run:
    cmd.append("--dry-run")
if args.confirm_write:
    cmd.append("--confirm-write")
```

- [ ] **Step 8: Verify dry-run test passes**

Run:

```powershell
pytest tests/test_pipeline_safety.py -q
```

Expected: `1 passed`.

- [ ] **Step 9: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/writeback_image_doc.py ocr_price/pipeline.py skills/quote-update/scripts/run_single.py skills/quote-update/scripts/run_batch.py tests/test_pipeline_safety.py
git commit -m "feat: add dry-run guardrails for quote updates"
```

---

### Task 7: Add Post-Write Workbook Audit

**Files:**
- Create: `ocr_price/audit.py`
- Create: `tests/test_audit.py`
- Modify: `ocr_price/pipeline.py`

- [ ] **Step 1: Write audit test**

Create `tests/test_audit.py`:

```python
from pathlib import Path

from openpyxl import Workbook

from ocr_price.audit import audit_image_doc_updates


def test_audit_image_doc_updates_detects_matching_values(tmp_path: Path):
    workbook = tmp_path / "project.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    mill = wb.create_sheet("测试钢厂")
    mill["H3"] = 3400
    mill["H4"] = 3200
    wb.save(workbook)

    result = audit_image_doc_updates(
        project_excel=workbook,
        updates=[
            {
                "项目文件Sheet": "测试钢厂",
                "H3_new": 3400,
                "H4_new": 3200,
            }
        ],
    )

    assert result["status"] == "ok"
    assert result["mismatch_count"] == 0


def test_audit_image_doc_updates_reports_mismatch(tmp_path: Path):
    workbook = tmp_path / "project.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报价表"
    mill = wb.create_sheet("测试钢厂")
    mill["H3"] = 3300
    mill["H4"] = 3200
    wb.save(workbook)

    result = audit_image_doc_updates(
        project_excel=workbook,
        updates=[
            {
                "项目文件Sheet": "测试钢厂",
                "H3_new": 3400,
                "H4_new": 3200,
            }
        ],
    )

    assert result["status"] == "failed"
    assert result["mismatch_count"] == 1
    assert result["mismatches"][0]["cell"] == "H3"
```

- [ ] **Step 2: Run audit tests and confirm import failure**

Run:

```powershell
pytest tests/test_audit.py -q
```

Expected before implementation: import failure for `ocr_price.audit`.

- [ ] **Step 3: Implement audit module**

Create `ocr_price/audit.py`:

```python
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
```

- [ ] **Step 4: Wire audit into pipeline image flow**

In `ocr_price/pipeline.py`, import:

```python
from .audit import audit_image_doc_updates
```

In `_image_flow`, after `apply_summary = _image_apply_summary(apply_report)`, add:

```python
    audit_report = None
    if not apply_report.get("dry_run") and apply_summary.get("updated_items"):
        audit_report = audit_image_doc_updates(project, apply_report.get("updates") or [])
```

Add `"audit_report": audit_report` to the returned image flow dict.

- [ ] **Step 5: Verify audit tests pass**

Run:

```powershell
pytest tests/test_audit.py -q
```

Expected: `2 passed`.

- [ ] **Step 6: Commit or record changes**

Run if Git is available:

```powershell
git add ocr_price/audit.py ocr_price/pipeline.py tests/test_audit.py
git commit -m "feat: audit image document writeback results"
```

---

### Task 8: Add Fixed Low-Capability Model Operation Rules

**Files:**
- Modify: `skills/quote-update/指令模板.md`
- Modify: `skills/quote-update/references/workflow.md`
- Modify: `README.md`

- [ ] **Step 1: Add safe-operation rules to instruction template**

In `skills/quote-update/指令模板.md`, add this section:

```markdown
## 低能力模型安全执行规则

1. 禁止手动打开或直接编辑项目报价 Excel。
2. 首次执行必须使用 `--dry-run`，只生成报告和待确认表。
3. 出现 `pending_confirmation`、厂家未确认、价格偏差过大、无网价参考、电议、脚本非零退出码时，必须停止并向用户报告，不得自行绕过。
4. 只有用户明确确认后，才允许用同一命令追加 `--confirm-write` 进行真实写入。
5. 写入完成后必须读取 JSON/Markdown 报告和 audit 结果，不得凭模型记忆汇报。
6. 不得编造价格、库存、厂家映射、更新数量。
```

- [ ] **Step 2: Update workflow with dry-run first**

In `skills/quote-update/references/workflow.md`, add this rule near the standard execution flow:

```markdown
### 安全执行顺序

任何模型执行更新时必须先跑 dry-run：

```powershell
python skills/quote-update/scripts/run_single.py --project "<项目文件>" --mode both --dry-run --headless
```

确认无待确认、无偏差异常，并经用户确认后，才允许真实写入：

```powershell
python skills/quote-update/scripts/run_single.py --project "<项目文件>" --mode both --confirm-write --headless
```
```

- [ ] **Step 3: Update README operator guidance**

In `README.md`, add:

```markdown
### 低能力模型/外部Agent执行限制

低能力模型不得直接修改 Excel，也不得手动拼写单元格更新。统一入口是 `skills/quote-update/scripts/run_single.py` 或 `run_batch.py`。默认先 `--dry-run`，用户确认后才可 `--confirm-write`。任何待确认、异常偏差、缺少网价参考、电议或脚本错误都必须停止并报告。
```

- [ ] **Step 4: Commit or record changes**

Run if Git is available:

```powershell
git add skills/quote-update/指令模板.md skills/quote-update/references/workflow.md README.md
git commit -m "docs: add safe operation rules for quote agents"
```

---

## Final Verification

- [ ] `pytest tests/test_location_parse.py tests/test_minimax_conversion.py tests/test_offline_validation.py tests/test_offline_price_deviation.py tests/test_inventory_writeback.py tests/test_reporting.py tests/test_pipeline_safety.py tests/test_audit.py -q`
- [ ] `python -m compileall -q ocr_price skills\quote-update\scripts tests`

Expected outcome: offline source JSON validation no longer rejects normal market movement solely because a price is above `6000`; Excel writeback blocks confirmed offline prices only when they deviate from same-sheet web prices by more than `1000元/吨` or `20%`; no-reference cases are written with an explicit report note; low-capability model runs can preview safely with `--dry-run`, require `--confirm-write` for real writes, and get audited after write.

## Self-Review

- Spec coverage: The user's threshold request (`1000` and `20%`) is implemented in Task 2 and enforced in Task 3. The concern about market movement invalidating fixed ranges is handled in Task 1. Website price as reference is handled by reading same-sheet `G3/G4` during `apply_writeback`. The earlier requirement to prevent lower-capability models from corrupting quote files is covered by Tasks 6, 7, and 8.
- Placeholder scan: No placeholders, TBDs, or unspecified test steps remain.
- Type consistency: `PriceDeviationConfig`, `_coerce_price`, `_check_price_deviation`, `_missing_reference_notes`, `dry_run`, and `audit_image_doc_updates` are introduced before later tasks use them.

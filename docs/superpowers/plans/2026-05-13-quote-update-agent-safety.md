# Quote Update Agent Safety Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让后续 agent 稳定按预期完成报价更新：dry-run 不写 Excel、confirm-write 只复用 dry-run 产物、库存颜色真实写入并在最终报告中按厂家/钢材型号和规格/库存情况逐条列出。

**Architecture:** 把关键安全要求从提示词下沉到脚本和 pipeline：用 manifest 绑定 dry-run 与 confirm-write，用 Excel hash 验证 dry-run 只读，用锁文件阻止并发执行，用官方确认脚本替代临时脚本。报告层只从真实 apply report 和 Excel 校验结果生成内容，不允许 agent 凭 stdout 或记忆汇报。

**Tech Stack:** Python 3, pytest, openpyxl, argparse, pathlib, json, subprocess/Windows process management.

---

## File Structure

- Modify: `ocr_price/pipeline.py`
  - 负责 manifest 创建/读取、Excel hash 校验、pipeline 锁、confirm-write 复用产物、最终结果结构。
- Modify: `ocr_price/reporting.py`
  - 负责 Markdown 最终报告，新增库存颜色明细表，禁止只输出汇总统计。
- Modify: `ocr_price/writeback_image_doc.py`
  - 确保 `inventory_report` 的 `applied` 明细包含厂家、钢材型号、规格、长度、材质、库存情况、Excel 单元格。
- Modify: `skills/quote-update/scripts/run_single.py`
  - 避免包装脚本产生无法管理的子进程；如果仍使用子进程，必须能终止进程树。
- Modify: `skills/quote-update/scripts/run_batch.py`
  - 同 `run_single.py`，批量模式也必须避免超时残留。
- Create: `skills/quote-update/scripts/apply_confirmations.py`
  - 官方确认脚本，替代 agent 自写临时脚本。
- Modify: `skills/quote-update/指令模板.md`
  - 写入低能力模型硬规则：禁止临时脚本、禁止直接调用内部函数、库存报告必须列明细。
- Modify: `skills/quote-update/references/rules.md`
  - 更新最终报告规则，明确库存颜色明细格式。
- Test: `tests/test_pipeline_safety.py`
  - dry-run hash、锁文件、manifest 基础测试。
- Test: `tests/test_pipeline_artifact_apply.py`
  - confirm-write 通过 manifest 复用产物，不走 latest glob。
- Test: `tests/test_reporting.py`
  - 最终报告库存颜色按明细表输出，不只输出汇总。
- Test: `tests/test_inventory_writeback.py`
  - `inventory_report.applied` 明细字段完整。
- Create: `tests/test_apply_confirmations.py`
  - 官方确认脚本的 JSON 修改行为测试。

---

## Task 1: 库存颜色报告改为逐条明细

**Files:**
- Modify: `ocr_price/reporting.py`
- Test: `tests/test_reporting.py`

- [ ] **Step 1: Write the failing reporting test**

Add this test to `tests/test_reporting.py`:

```python
def test_render_single_report_markdown_lists_inventory_details():
    result = {
        "project": "项目报价/测试项目.xlsx",
        "mode": "both",
        "started_at": "2026-05-13T09:00:00",
        "ended_at": "2026-05-13T09:02:00",
        "status": "ok",
        "image_doc": {
            "status": "ok",
            "apply_summary": {
                "updated_count": 0,
                "skipped_count": 0,
                "updated_items": [],
                "skipped_items": [],
            },
            "inventory_report": {
                "status": "ok",
                "cleared_count": 3,
                "applied_count": 2,
                "applied": [
                    {
                        "mill": "桂鑫报价",
                        "sheet_mill": "桂鑫",
                        "product": "螺纹",
                        "spec": "12",
                        "length": "9",
                        "material": "HRB400E",
                        "status": "充足",
                        "cell": "P16",
                    },
                    {
                        "mill": "贵航报价",
                        "sheet_mill": "贵航",
                        "product": "圆钢",
                        "spec": "16",
                        "length": "",
                        "material": "",
                        "status": "告警",
                        "cell": "R18",
                    },
                ],
            },
        },
    }

    markdown = render_single_report_markdown(result, json_report_path="运行产物/report.json")

    assert "### 3. 库存颜色标注明细" in markdown
    assert "| 厂家 | 钢材型号和规格 | 库存情况 | 单元格 |" in markdown
    assert "| 桂鑫 | 螺纹 12 9 HRB400E | 充足 | P16 |" in markdown
    assert "| 贵航 | 圆钢 16 | 告警 | R18 |" in markdown
    assert "蓝色（充足）" not in markdown
    assert "新标注颜色：" not in markdown
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pytest tests/test_reporting.py::test_render_single_report_markdown_lists_inventory_details -v
```

Expected: FAIL because current report outputs aggregate counts instead of the required detail table.

- [ ] **Step 3: Implement detail table rendering**

In `ocr_price/reporting.py`, replace the current inventory aggregate block with helper functions:

```python
def _inventory_spec_text(item: dict[str, Any]) -> str:
    parts = [
        str(item.get("product") or "").strip(),
        str(item.get("spec") or "").strip(),
        str(item.get("length") or "").strip(),
        str(item.get("material") or "").strip(),
    ]
    return " ".join(x for x in parts if x) or "未识别规格"


def _inventory_rows(inventory_report: dict[str, Any]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for item in inventory_report.get("applied", []) or []:
        if not isinstance(item, dict):
            continue
        rows.append(
            {
                "mill": str(item.get("sheet_mill") or item.get("mill") or ""),
                "spec": _inventory_spec_text(item),
                "status": str(item.get("status") or ""),
                "cell": str(item.get("cell") or ""),
            }
        )
    return rows
```

Then render:

```python
inventory_report = image_doc.get("inventory_report") if isinstance(image_doc, dict) else None
if isinstance(inventory_report, dict) and inventory_report.get("status") == "ok":
    rows = _inventory_rows(inventory_report)
    lines.append("")
    lines.append("### 3. 库存颜色标注明细")
    lines.append("")
    lines.append("| 厂家 | 钢材型号和规格 | 库存情况 | 单元格 |")
    lines.append("|------|----------------|----------|--------|")
    if rows:
        for row in rows:
            lines.append(f"| {row['mill']} | {row['spec']} | {row['status']} | {row['cell']} |")
    else:
        lines.append("| 无 | 无 | 无 | 无 |")
    lines.append("")
```

- [ ] **Step 4: Verify reporting test passes**

Run:

```powershell
pytest tests/test_reporting.py -v
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add ocr_price/reporting.py tests/test_reporting.py
git commit -m "report inventory color details"
```

---

## Task 2: 库存写回报告补全明细字段

**Files:**
- Modify: `ocr_price/inventory.py`
- Test: `tests/test_inventory_writeback.py`

- [ ] **Step 1: Write the failing inventory report test**

Extend `tests/test_inventory_writeback.py` after the existing assertions:

```python
    applied = result["applied"][0]
    assert applied["mill"] == "来源钢厂"
    assert applied["sheet_mill"] == "目标钢厂"
    assert applied["product"] == "螺纹"
    assert applied["spec"] == "12"
    assert applied["length"] == "9"
    assert applied["material"] == "HRB400E"
    assert applied["status"] == "告警"
    assert applied["cell"] == "E12"
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pytest tests/test_inventory_writeback.py::test_apply_inventory_uses_confirmed_mapping_and_clears_existing_colors -v
```

Expected: FAIL if `product`, `length`, `material`, or `cell` are missing from `applied`.

- [ ] **Step 3: Implement missing detail fields**

In `ocr_price/inventory.py`, where applied inventory entries are appended, ensure each item includes:

```python
applied.append(
    {
        "mill": source_mill,
        "sheet_mill": sheet_mill,
        "row": row_idx,
        "col": col_idx,
        "cell": ws.cell(row=row_idx, column=col_idx).coordinate,
        "product": item.product,
        "spec": item.spec,
        "length": item.length,
        "material": item.material,
        "status": item.status,
    }
)
```

Use the actual local variable names in `apply_inventory_to_project`; do not change existing color behavior.

- [ ] **Step 4: Verify inventory tests pass**

Run:

```powershell
pytest tests/test_inventory_writeback.py -v
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add ocr_price/inventory.py tests/test_inventory_writeback.py
git commit -m "include inventory detail fields"
```

---

## Task 3: dry-run Excel hash guard

**Files:**
- Modify: `ocr_price/pipeline.py`
- Test: `tests/test_pipeline_safety.py`

- [ ] **Step 1: Add failing hash helper tests**

Add to `tests/test_pipeline_safety.py`:

```python
from ocr_price import pipeline


def test_file_sha256_changes_when_file_changes(tmp_path: Path):
    path = tmp_path / "project.xlsx"
    path.write_bytes(b"before")
    before = pipeline._file_sha256(path)

    path.write_bytes(b"after")
    after = pipeline._file_sha256(path)

    assert before != after


def test_assert_dry_run_unchanged_raises_on_modified_file(tmp_path: Path):
    path = tmp_path / "project.xlsx"
    path.write_bytes(b"before")
    before = pipeline._file_sha256(path)
    path.write_bytes(b"after")

    try:
        pipeline._assert_dry_run_unchanged(path, before)
    except RuntimeError as exc:
        assert "dry-run 修改了项目 Excel" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError when dry-run changes the workbook")
```

- [ ] **Step 2: Run failing tests**

Run:

```powershell
pytest tests/test_pipeline_safety.py::test_file_sha256_changes_when_file_changes tests/test_pipeline_safety.py::test_assert_dry_run_unchanged_raises_on_modified_file -v
```

Expected: FAIL because helpers do not exist.

- [ ] **Step 3: Add hash helpers**

Add to `ocr_price/pipeline.py` near `_json_write`:

```python
import hashlib


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _assert_dry_run_unchanged(project: Path, before_hash: str) -> None:
    after_hash = _file_sha256(project)
    if after_hash != before_hash:
        raise RuntimeError(
            f"dry-run 修改了项目 Excel，禁止继续：{project} before={before_hash} after={after_hash}"
        )
```

- [ ] **Step 4: Wire hash guard into `run_single`**

In `run_single`, before any flow starts:

```python
project_before_hash = _file_sha256(project) if dry_run and project.exists() else None
```

Before returning from `run_single`, including the pending path, call:

```python
if dry_run and project_before_hash is not None:
    _assert_dry_run_unchanged(project, project_before_hash)
```

Also store both hashes in result:

```python
if project_before_hash is not None:
    result["project_excel_hash_before"] = project_before_hash
    result["project_excel_hash_after"] = _file_sha256(project)
```

- [ ] **Step 5: Run safety tests**

Run:

```powershell
pytest tests/test_pipeline_safety.py -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add ocr_price/pipeline.py tests/test_pipeline_safety.py
git commit -m "guard dry-run workbook writes"
```

---

## Task 4: dry-run manifest and confirm-write manifest reuse

**Files:**
- Modify: `ocr_price/pipeline.py`
- Modify: `skills/quote-update/scripts/run_single.py`
- Test: `tests/test_pipeline_artifact_apply.py`

- [ ] **Step 1: Add failing manifest test**

Add to `tests/test_pipeline_artifact_apply.py`:

```python
def test_run_single_confirm_write_uses_manifest_paths(monkeypatch, tmp_path: Path):
    manifest = tmp_path / "manifest.json"
    manifest.write_text(
        json.dumps(
            {
                "project": str(tmp_path / "安徽合肥-安徽蚌埠-测试.xlsx"),
                "artifact_dir": str(tmp_path / "artifacts"),
                "web": {
                    "location": "安徽合肥",
                    "mapping_json": str(tmp_path / "artifacts" / "厂家对照表_安徽合肥_已确认.json"),
                    "raw_price_excel": str(tmp_path / "artifacts" / "安徽合肥2026-05-13建筑钢材原料价格清单.xlsx"),
                },
                "image_doc": {
                    "location": "安徽蚌埠",
                    "mapping_json": str(tmp_path / "artifacts" / "图片文档厂家对照表_安徽蚌埠_已确认.json"),
                    "source_jsons": [str(tmp_path / "artifacts" / "ocr价格提取_桂鑫报价.json")],
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
        artifact_dir=tmp_path / "artifacts",
        dry_run=False,
        confirm_write=True,
        manifest_path=manifest,
    )

    assert result["status"] == "ok"
    assert seen["web_manifest"].endswith("建筑钢材原料价格清单.xlsx")
    assert seen["image_manifest"][0].endswith("ocr价格提取_桂鑫报价.json")
```

- [ ] **Step 2: Run failing manifest test**

Run:

```powershell
pytest tests/test_pipeline_artifact_apply.py::test_run_single_confirm_write_uses_manifest_paths -v
```

Expected: FAIL because `manifest_path`, `_web_apply_from_manifest`, and `_image_apply_from_manifest` do not exist.

- [ ] **Step 3: Add parser argument**

In `ocr_price/pipeline.py`, add to single and batch parsers:

```python
p_single.add_argument("--manifest", help="dry-run生成的manifest路径；confirm-write默认必须复用它")
p_batch.add_argument("--manifest", help="dry-run生成的manifest路径；单文件建议使用，批量后续拆分支持")
```

In `skills/quote-update/scripts/run_single.py`, add:

```python
p.add_argument("--manifest", help="dry-run生成的manifest路径")
```

and append:

```python
if args.manifest:
    cmd.extend(["--manifest", args.manifest])
```

- [ ] **Step 4: Add manifest helpers**

In `ocr_price/pipeline.py`:

```python
def _manifest_path(artifact_dir: Path) -> Path:
    return artifact_dir / "dry_run_manifest.json"


def _write_manifest(path: Path, result: dict[str, Any], artifact_dir: Path) -> None:
    manifest = {
        "project": result.get("project"),
        "artifact_dir": str(artifact_dir),
        "mode": result.get("mode"),
        "started_at": result.get("started_at"),
        "ended_at": result.get("ended_at"),
        "project_excel_hash_before": result.get("project_excel_hash_before"),
        "project_excel_hash_after": result.get("project_excel_hash_after"),
        "web": result.get("web"),
        "image_doc": result.get("image_doc"),
    }
    _json_write(path, manifest)


def _read_manifest(path: Path) -> dict[str, Any]:
    payload = _json_read(path)
    if not isinstance(payload, dict):
        raise RuntimeError(f"manifest不是JSON对象：{path}")
    return payload
```

- [ ] **Step 5: Add manifest apply wrappers**

Implement wrappers initially by validating manifest and delegating to existing artifact apply functions:

```python
def _web_apply_from_manifest(project: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    artifact_dir = Path(str(manifest.get("artifact_dir") or ""))
    web = manifest.get("web") if isinstance(manifest.get("web"), dict) else {}
    location = str(web.get("location") or manifest.get("web_location") or "")
    if not artifact_dir.exists():
        raise RuntimeError(f"manifest产物目录不存在：{artifact_dir}")
    return _web_apply_from_artifacts(project=project, location=location, artifact_dir=artifact_dir)


def _image_apply_from_manifest(project: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    artifact_dir = Path(str(manifest.get("artifact_dir") or ""))
    image_doc = manifest.get("image_doc") if isinstance(manifest.get("image_doc"), dict) else {}
    location = str(image_doc.get("location") or manifest.get("image_location") or "")
    source_jsons = [Path(x) for x in image_doc.get("source_jsons", []) or []]
    return _image_apply_from_artifacts(project=project, location=location, artifact_dir=artifact_dir, source_jsons=source_jsons)
```

Then replace confirm-write default path in `run_single`:

```python
manifest = _read_manifest(manifest_path) if manifest_path else None
if confirm_write and manifest is not None and not refresh_web_artifacts:
    result["web"] = _web_apply_from_manifest(project, manifest)
```

Do the same for image_doc.

- [ ] **Step 6: Generate manifest after dry-run**

After `run_single(...)` returns in CLI `main`, if `args.dry_run`:

```python
manifest_out = _manifest_path(project_artifact_dir)
_write_manifest(manifest_out, result, project_artifact_dir)
print(f"Manifest: {manifest_out}")
```

- [ ] **Step 7: Verify artifact apply tests**

Run:

```powershell
pytest tests/test_pipeline_artifact_apply.py -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

```powershell
git add ocr_price/pipeline.py skills/quote-update/scripts/run_single.py tests/test_pipeline_artifact_apply.py
git commit -m "reuse dry-run manifest for confirm write"
```

---

## Task 5: Pipeline lock to prevent concurrent updates

**Files:**
- Modify: `ocr_price/pipeline.py`
- Test: `tests/test_pipeline_safety.py`

- [ ] **Step 1: Add failing lock test**

Add to `tests/test_pipeline_safety.py`:

```python
def test_pipeline_lock_blocks_existing_active_lock(tmp_path: Path):
    lock_path = tmp_path / ".quote_update.lock"
    lock_path.write_text(
        json.dumps({"pid": 999999, "project": "project.xlsx"}, ensure_ascii=False),
        encoding="utf-8",
    )

    try:
        with pipeline._pipeline_lock(lock_path, project=tmp_path / "project.xlsx"):
            raise AssertionError("lock should block")
    except RuntimeError as exc:
        assert "已有报价更新任务锁" in str(exc)
```

- [ ] **Step 2: Run failing test**

Run:

```powershell
pytest tests/test_pipeline_safety.py::test_pipeline_lock_blocks_existing_active_lock -v
```

Expected: FAIL because `_pipeline_lock` does not exist.

- [ ] **Step 3: Implement conservative lock**

In `ocr_price/pipeline.py`:

```python
from contextlib import contextmanager


@contextmanager
def _pipeline_lock(lock_path: Path, project: Path):
    if lock_path.exists():
        raise RuntimeError(f"已有报价更新任务锁，请先人工确认并清理：{lock_path}")
    payload = {
        "pid": os.getpid(),
        "project": str(project),
        "started_at": datetime.now().isoformat(timespec="seconds"),
    }
    _json_write(lock_path, payload)
    try:
        yield
    finally:
        try:
            lock_path.unlink()
        except FileNotFoundError:
            pass
```

Use conservative behavior first: any lock blocks. Do not auto-delete locks in the first implementation.

- [ ] **Step 4: Wrap single CLI execution**

In `main()` single branch, before `run_single(...)`:

```python
lock_path = artifact_base_dir / ".quote_update.lock"
with _pipeline_lock(lock_path, project_path):
    result = run_single(...)
```

Ensure report writing remains inside the lock context so another process cannot interleave writes.

- [ ] **Step 5: Verify lock tests**

Run:

```powershell
pytest tests/test_pipeline_safety.py -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add ocr_price/pipeline.py tests/test_pipeline_safety.py
git commit -m "add quote update pipeline lock"
```

---

## Task 6: Wrapper scripts must not leak child processes

**Files:**
- Modify: `skills/quote-update/scripts/run_single.py`
- Modify: `skills/quote-update/scripts/run_batch.py`

- [ ] **Step 1: Replace subprocess wrapper with direct module call for single**

In `skills/quote-update/scripts/run_single.py`, avoid launching `python -m ocr_price.pipeline`. Build an argv list and call `ocr_price.pipeline.main()` in-process:

```python
from ocr_price import pipeline


def main() -> int:
    args = _build_parser().parse_args()
    root = _repo_root()
    old_cwd = Path.cwd()
    old_argv = sys.argv[:]
    try:
        os.chdir(root)
        sys.argv = [
            "ocr_price.pipeline",
            "single",
            "--project",
            args.project,
            "--mode",
            args.mode,
            "--account-file",
            args.account_file,
            "--artifact-dir",
            args.artifact_dir,
        ]
        # append the same optional args as before
        return pipeline.main()
    finally:
        sys.argv = old_argv
        os.chdir(old_cwd)
```

Use the existing append logic, but append into `sys.argv` instead of `cmd`.

- [ ] **Step 2: Apply same direct call pattern to batch**

In `skills/quote-update/scripts/run_batch.py`, replace `subprocess.run(cmd, cwd=root).returncode` with the same in-process `pipeline.main()` pattern.

- [ ] **Step 3: Run wrapper smoke checks**

Run:

```powershell
python skills/quote-update/scripts/run_single.py --help
python skills/quote-update/scripts/run_batch.py --help
```

Expected: both show help and return exit code 0.

- [ ] **Step 4: Run pipeline safety tests**

Run:

```powershell
pytest tests/test_pipeline_safety.py tests/test_pipeline_artifact_apply.py -v
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add skills/quote-update/scripts/run_single.py skills/quote-update/scripts/run_batch.py
git commit -m "avoid wrapper process leaks"
```

---

## Task 7: Official confirmation script

**Files:**
- Create: `skills/quote-update/scripts/apply_confirmations.py`
- Test: `tests/test_apply_confirmations.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_apply_confirmations.py`:

```python
from __future__ import annotations

import json
from pathlib import Path

from skills.quote_update.scripts import apply_confirmations


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

    summary = apply_confirmations.apply_web_confirmations(
        mapping,
        confirmed_matches={"徐钢": "徐钢", "六钢": "六安钢铁"},
        skip_new_mills=True,
    )

    rows = json.loads(mapping.read_text(encoding="utf-8"))
    assert rows[0]["状态"] == "已确认匹配"
    assert rows[1]["最新清单厂家Sheet"] == "六安钢铁"
    assert rows[1]["状态"] == "已确认匹配"
    assert rows[2]["状态"] == "已确认不更新"
    assert rows[3]["状态"] == "已确认不更新"
    assert summary["pending_count"] == 0
```

If importing from `skills.quote_update` is not possible because the directory name contains a hyphen, implement tests by invoking the script with `subprocess.run`; do not rename directories in this task.

- [ ] **Step 2: Run failing test**

Run:

```powershell
pytest tests/test_apply_confirmations.py -v
```

Expected: FAIL because script does not exist.

- [ ] **Step 3: Implement script**

Create `skills/quote-update/scripts/apply_confirmations.py` with:

```python
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
```

- [ ] **Step 4: Run tests**

Run:

```powershell
pytest tests/test_apply_confirmations.py -v
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add skills/quote-update/scripts/apply_confirmations.py tests/test_apply_confirmations.py
git commit -m "add official mapping confirmation script"
```

---

## Task 8: Update agent-facing rules

**Files:**
- Modify: `skills/quote-update/指令模板.md`
- Modify: `skills/quote-update/references/rules.md`
- Modify: `skills/quote-update/references/workflow.md`

- [ ] **Step 1: Update low-capability model safety rules**

In `skills/quote-update/指令模板.md`, extend the section `低能力模型安全执行规则`:

```markdown
7. 禁止创建临时 Python 脚本修改厂家对照表、Excel、库存颜色或报告；只能使用 `skills/quote-update/scripts/` 下的正式脚本。
8. 禁止直接调用 `ocr_price.*` 内部函数补写 Excel 或库存颜色。正式脚本失败时必须停止并报告。
9. dry-run 后必须检查输出的 `Manifest:` 路径；confirm-write 必须带 `--manifest` 使用同一次 dry-run 产物。
10. 最终库存颜色报告必须逐条列出：厂家、钢材型号和规格、库存情况、单元格；禁止只写蓝/黄/红总数。
```

- [ ] **Step 2: Update result reporting rules**

In `skills/quote-update/references/rules.md`, replace inventory aggregate requirement with:

```markdown
- **库存颜色明细规则（新增）**：
  - 最终报告必须逐条列出库存颜色标注明细，禁止只输出总统计。
  - 表格格式固定为：`厂家 | 钢材型号和规格 | 库存情况 | 单元格`。
  - `钢材型号和规格` 必须合并产品、规格、长度、材质，例如 `螺纹 12 9 HRB400E`。
  - 如果库存颜色标注失败、缺少 `inventory_report`、或明细为空，任务状态必须标为未通过验收。
```

- [ ] **Step 3: Update workflow commands**

In `skills/quote-update/references/workflow.md`, update confirm-write command:

```powershell
python skills/quote-update/scripts/run_single.py --project "<项目文件>" --mode both --confirm-write --headless --manifest "运行产物/<项目名>/dry_run_manifest.json"
```

Also add:

```markdown
如果 dry-run 没有输出 `Manifest:`，不得执行 confirm-write。
```

- [ ] **Step 4: Review docs**

Run:

```powershell
rg -n "库存颜色统计|蓝色X个|confirm-write|manifest|临时 Python|内部函数" skills/quote-update
```

Expected: old aggregate wording is either removed or explicitly marked as secondary; manifest and no-temporary-script rules are present.

- [ ] **Step 5: Commit**

```powershell
git add skills/quote-update/指令模板.md skills/quote-update/references/rules.md skills/quote-update/references/workflow.md
git commit -m "document quote update safety workflow"
```

---

## Task 9: End-to-end verification checklist

**Files:**
- No new production files.
- Verify all modified tests and a dry-run smoke path.

- [ ] **Step 1: Run focused tests**

Run:

```powershell
pytest tests/test_pipeline_safety.py tests/test_pipeline_artifact_apply.py tests/test_reporting.py tests/test_inventory_writeback.py tests/test_apply_confirmations.py -v
```

Expected: PASS.

- [ ] **Step 2: Run broader tests for touched modules**

Run:

```powershell
pytest tests/test_offline_validation.py tests/test_offline_price_deviation.py tests/test_audit.py tests/test_minimax_conversion.py -v
```

Expected: PASS.

- [ ] **Step 3: Verify dry-run produces manifest and does not modify Excel**

Run on a copy of a project Excel, not the live file:

```powershell
Copy-Item "项目报价/安徽合肥-安徽蚌埠-蚌投(春和苑）询价表2026.4.24.xlsx" "$env:TEMP/quote-update-smoke.xlsx"
python skills/quote-update/scripts/run_single.py --project "$env:TEMP/quote-update-smoke.xlsx" --mode image_doc --dry-run
```

Expected:

```text
Status: prepared
Manifest: 运行产物/quote-update-smoke/dry_run_manifest.json
```

If status is `pending_confirmation`, that is acceptable for dry-run, but manifest must exist and Excel hash before/after must match.

- [ ] **Step 4: Verify final report inventory format from test fixture**

Open the generated Markdown report or run the reporting unit test fixture. Confirm inventory section uses:

```markdown
| 厂家 | 钢材型号和规格 | 库存情况 | 单元格 |
```

Expected: no inventory section that only says `蓝色X个 / 黄色X个 / 红色X个`.

- [ ] **Step 5: Final commit**

```powershell
git status --short
git add ocr_price skills tests docs
git commit -m "harden quote update agent workflow"
```

---

## Self-Review

- Spec coverage:
  - dry-run 不写 Excel：Task 3.
  - confirm-write 复用 dry-run 产物：Task 4.
  - 禁止 refresh 除非明确要求：Task 8 文档化，Task 4 以 manifest 固定默认路径。
  - 进程泄漏防护：Task 5 和 Task 6.
  - 禁止 agent 私自写脚本绕流程：Task 7 和 Task 8.
  - 库存颜色最终报告按厂家、钢材型号和规格、库存情况逐条列出：Task 1 和 Task 2.
  - 结果输出可信：Task 9.
- Placeholder scan:
  - No `TBD`, `TODO`, or unspecified implementation steps.
- Type consistency:
  - `inventory_report.applied` uses `mill`, `sheet_mill`, `product`, `spec`, `length`, `material`, `status`, `cell`.
  - Reporting helper reads the same keys.
  - Manifest helpers use `Path`, `dict[str, Any]`, and existing `run_single` flow.

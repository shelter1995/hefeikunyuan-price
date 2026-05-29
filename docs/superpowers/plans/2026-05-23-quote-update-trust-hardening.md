# Quote Update Trust Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make quote update runs repeatable, complete, auditable, and safe for both strong and weaker model operators.

**Architecture:** Centralize business rules in a small rules module, strengthen dry-run manifest validation before writes, and make reports/logs contract-driven. Keep current CLI entrypoints stable while adding focused helper modules and tests.

**Tech Stack:** Python 3.12, openpyxl, requests, Playwright, pytest.

---

### Task 1: Central Rule Layer

**Files:**
- Create: `ocr_price/rules.py`
- Modify: `ocr_price/offline_validation.py`
- Modify: `ocr_price/writeback_image_doc.py`
- Test: `tests/test_rules_contract.py`

- [ ] Write failing tests for centralized constants and price deviation behavior.
- [ ] Move hard ranges, statuses, confirmed mapping statuses, and deviation logic to `ocr_price/rules.py`.
- [ ] Import rules from validation and writeback modules.
- [ ] Run `pytest tests/test_rules_contract.py tests/test_offline_validation.py tests/test_offline_price_deviation.py -q`.

### Task 2: Manifest Guard

**Files:**
- Modify: `ocr_price/pipeline.py`
- Test: `tests/test_manifest_guard.py`

- [ ] Write failing tests for project mismatch, hash mismatch, and source hash mismatch.
- [ ] Add source hashing to dry-run manifest.
- [ ] Validate project path, mode, dry-run hash, artifact dir, raw web file hash, and image JSON hashes before confirm-write reuse.
- [ ] Run `pytest tests/test_manifest_guard.py tests/test_pipeline_artifact_apply.py tests/test_pipeline_safety.py -q`.

### Task 3: Report Contract And Events

**Files:**
- Create: `ocr_price/events.py`
- Modify: `ocr_price/pipeline.py`
- Modify: `ocr_price/reporting.py`
- Test: `tests/test_reporting.py`
- Test: `tests/test_events.py`

- [ ] Write failing tests for pending reports, reuse notices, inventory errors, and event JSONL output.
- [ ] Add structured pipeline events in result JSON and optional JSONL report artifact.
- [ ] Render pending and skipped states explicitly in Markdown.
- [ ] Run `pytest tests/test_reporting.py tests/test_events.py -q`.

### Task 4: VLM Replay And Engineering Config

**Files:**
- Create: `tests/fixtures/minimax_vlm_response_xugang.json`
- Create: `tests/test_minimax_replay.py`
- Create: `pyproject.toml`
- Modify: `.gitignore`
- Modify: `README.md`

- [ ] Add a fixed MiniMax response fixture and replay conversion test.
- [ ] Add pytest and ruff configuration in `pyproject.toml`.
- [ ] Tighten ignore rules for generated backups and local planning files where appropriate.
- [ ] Document the trusted-run workflow and new manifest guard.
- [ ] Run `pytest tests/ -q`.

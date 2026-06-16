from __future__ import annotations

import json
from pathlib import Path

from ocr_price.offline_validation import validate_offline_payload
from ocr_price.writeback_image_doc import load_source_prices, load_source_prices_with_errors


def _base_payload() -> dict:
    return {
        "quote_date": "2026-05-11",
        "meta": {"input_file": "徐钢4.13.jpg"},
        "records": [
            {"location": "蚌埠", "rebar_price": 3200, "coil_price": 3400},
        ],
        "inventory": [
            {"规格": "9米螺纹12E", "状态": "充足"},
            {"规格": "9米螺纹14E", "状态": "告警"},
            {"规格": "9米螺纹16E", "状态": "缺货"},
        ],
    }


def _write_json(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def test_validate_offline_payload_valid() -> None:
    payload = _base_payload()
    result = validate_offline_payload(payload, target_location="蚌埠")
    assert result.is_valid is True
    assert result.errors == []


def test_validate_offline_payload_allows_market_price_outside_old_fixed_range() -> None:
    payload = _base_payload()
    payload["records"][0]["rebar_price"] = 6500
    result = validate_offline_payload(payload, target_location="蚌埠")
    assert result.is_valid is True
    assert result.errors == []


def test_validate_offline_payload_rejects_non_price_integer() -> None:
    payload = _base_payload()
    payload["records"][0]["coil_price"] = 86
    result = validate_offline_payload(payload, target_location="蚌埠")
    assert result.is_valid is False
    assert any("below hard minimum" in err for err in result.errors)


def test_validate_offline_payload_missing_target_location() -> None:
    payload = _base_payload()
    payload["records"][0]["location"] = "合肥"
    result = validate_offline_payload(payload, target_location="蚌埠")
    assert result.is_valid is False
    assert any("target location not found" in err for err in result.errors)


def test_validate_offline_payload_invalid_inventory_status() -> None:
    payload = _base_payload()
    payload["inventory"][0]["状态"] = "未知"
    result = validate_offline_payload(payload, target_location="蚌埠")
    assert result.is_valid is False
    assert any("invalid status" in err for err in result.errors)


def test_load_source_prices_skips_invalid_payload(tmp_path: Path) -> None:
    valid = _base_payload()
    invalid = _base_payload()
    invalid["meta"]["input_file"] = "金虹4.13.jpg"
    invalid["records"][0]["coil_price"] = 10001

    valid_path = tmp_path / "valid.json"
    invalid_path = tmp_path / "invalid.json"
    _write_json(valid_path, valid)
    _write_json(invalid_path, invalid)

    prices = load_source_prices([valid_path, invalid_path], location="蚌埠")
    assert len(prices) == 1
    assert prices[0].company == "徐钢"


def test_load_source_prices_with_errors_reports_invalid_payload(tmp_path: Path) -> None:
    valid = _base_payload()
    invalid = _base_payload()
    invalid["meta"]["input_file"] = "金虹4.13.jpg"
    invalid["records"][0]["coil_price"] = 10001

    valid_path = tmp_path / "valid.json"
    invalid_path = tmp_path / "invalid.json"
    _write_json(valid_path, valid)
    _write_json(invalid_path, invalid)

    result = load_source_prices_with_errors([valid_path, invalid_path], location="蚌埠")

    assert len(result.prices) == 1
    assert len(result.errors) == 1
    assert result.errors[0]["source_json"].endswith("invalid.json")
    assert "above hard maximum" in result.errors[0]["errors"][0]

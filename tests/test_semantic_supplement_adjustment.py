from __future__ import annotations

import json
from pathlib import Path

from ocr_price import writeback_image_doc
from ocr_price.writeback_image_doc import load_source_prices


def _write_json(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def test_llm_semantic_supplement_adjustment_is_merged_with_base_price(
    tmp_path: Path, monkeypatch
):
    base_json = tmp_path / "ocr价格提取_徐钢.json"
    supplement_txt = tmp_path / "徐钢补充.txt"
    supplement_json = tmp_path / "ocr价格提取_徐钢补充.json"

    supplement_txt.write_text(
        "徐钢蚌埠  滁州区域螺纹表价优惠10出，其他区域螺纹表价优惠20出，盘螺所有区域表价出",
        encoding="utf-8",
    )
    _write_json(
        base_json,
        {
            "meta": {"input_file": str(tmp_path / "徐钢.jpg"), "record_count": 1},
            "company": "徐钢集团",
            "quote_date": "2026-05-20",
            "records": [
                {
                    "location": "蚌埠",
                    "rebar_price": 3270,
                    "coil_price": 3460,
                }
            ],
        },
    )
    _write_json(
        supplement_json,
        {
            "meta": {"input_file": str(supplement_txt), "record_count": 0},
            "company": None,
            "quote_date": None,
            "records": [],
        },
    )

    def fake_interpret(text: str, location: str, company: str | None = None):
        assert "表价优惠10出" in text
        assert location == "蚌埠"
        return {
            "location": "蚌埠",
            "rebar_price": -10,
            "coil_price": 0,
            "is_adjustment": True,
            "is_electronic_negotiation": False,
            "reason": "蚌埠和滁州螺纹优惠10出，盘螺所有区域表价出",
        }

    monkeypatch.setattr(
        writeback_image_doc,
        "interpret_supplement_adjustment_with_llm",
        fake_interpret,
    )

    prices = load_source_prices([base_json, supplement_json], location="蚌埠")

    assert len(prices) == 1
    price = prices[0]
    assert price.company == "徐钢"
    assert price.source_file == "ocr价格提取_徐钢.json+ocr价格提取_徐钢补充.json"
    assert price.rebar_price == 3260
    assert price.coil_price == 3460

    cached = json.loads(supplement_json.read_text(encoding="utf-8"))
    assert cached["_semantic_adjustments"]["蚌埠"]["rebar_price"] == -10
    assert cached["_semantic_adjustments"]["蚌埠"]["coil_price"] == 0

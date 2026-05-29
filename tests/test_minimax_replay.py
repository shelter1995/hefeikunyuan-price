import json
from pathlib import Path

from ocr_price.minimax_vision import convert_to_ocr_format


def test_minimax_fixture_replays_to_stable_ocr_format():
    fixture = Path(__file__).parent / "fixtures" / "minimax_vlm_response_xugang.json"
    vision_result = json.loads(fixture.read_text(encoding="utf-8"))

    converted = convert_to_ocr_format(vision_result, target_cities=["蚌埠"])

    assert converted["company"] == "徐钢"
    assert converted["meta"]["target_location"] == "蚌埠"
    assert converted["meta"]["record_count"] == 1
    assert converted["records"][0]["location"] == "蚌埠"
    assert converted["records"][0]["rebar_price"] == 3187
    assert converted["records"][0]["coil_price"] == 3347
    assert converted["_vision_result"]["库存情况"][1]["状态"] == "告警"

from ocr_price import minimax_vision


def test_convert_to_ocr_format_respects_requested_city_and_keeps_inventory():
    vision_result = {
        "厂家名称": "测试厂家",
        "报价日期": "2026-05-11",
        "合肥": {"螺纹": 3350, "盘螺": 3550, "螺纹为电议": False, "盘螺为电议": False},
        "蚌埠": {"螺纹": 3330, "盘螺": None, "螺纹为电议": False, "盘螺为电议": True},
        "库存情况": [
            {"规格": "16E", "状态": "告警", "原始描述": "82件"},
            {"规格": "18E", "状态": "缺货", "原始描述": "无货"},
        ],
        "_source": {"image_path": "sample.jpg"},
    }

    converted = minimax_vision.convert_to_ocr_format(
        vision_result, target_cities=["蚌埠"]
    )

    assert converted["meta"]["target_location"] == "蚌埠"
    assert converted["meta"]["record_count"] == 1
    assert len(converted["records"]) == 1
    assert converted["records"][0]["location"] == "蚌埠"
    assert converted["_vision_result"]["库存情况"] == vision_result["库存情况"]


def test_analyze_quote_image_to_ocr_format_passes_target_cities(monkeypatch):
    captured: dict[str, object] = {}

    def fake_analyze_quote_image_with_retry(**kwargs):
        captured["analyze_target_cities"] = kwargs.get("target_cities")
        return {
            "厂家名称": "测试厂家",
            "报价日期": "2026-05-11",
            "蚌埠": {"螺纹": 3330, "盘螺": 3530, "螺纹为电议": False, "盘螺为电议": False},
            "库存情况": [{"规格": "16E", "状态": "充足", "原始描述": ""}],
            "_source": {"image_path": "fake.jpg"},
        }

    def fake_convert_to_ocr_format(vision_result, target_cities=None):
        captured["convert_target_cities"] = target_cities
        return {
            "meta": {"target_location": "蚌埠", "record_count": 1},
            "records": [{"location": "蚌埠"}],
            "_vision_result": vision_result,
        }

    monkeypatch.setattr(
        minimax_vision, "analyze_quote_image_with_retry", fake_analyze_quote_image_with_retry
    )
    monkeypatch.setattr(minimax_vision, "convert_to_ocr_format", fake_convert_to_ocr_format)

    result = minimax_vision.analyze_quote_image_to_ocr_format(
        image_path="fake.jpg",
        target_cities=["蚌埠"],
        use_retry=True,
    )

    assert captured["analyze_target_cities"] == ["蚌埠"]
    assert captured["convert_target_cities"] == ["蚌埠"]
    assert result["records"][0]["location"] == "蚌埠"

from ocr_price.web_price import _parse_location


def test_parse_location_dual_keeps_visual_order():
    web_location, image_doc_location = _parse_location("安徽合肥-安徽蚌埠-项目报价单.xlsx")
    assert web_location == "安徽合肥"
    assert image_doc_location == "安徽蚌埠"


def test_parse_location_single_fallback():
    web_location, image_doc_location = _parse_location("江苏南通-项目报价单.xlsx")
    assert web_location == "江苏南通"
    assert image_doc_location == "江苏南通"


def test_parse_location_overlapping_keywords_no_duplicate_span():
    web_location, image_doc_location = _parse_location("安徽合肥项目报价单.xlsx")
    assert web_location == "安徽合肥"
    assert image_doc_location == "安徽合肥"

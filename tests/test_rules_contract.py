from ocr_price.rules import (
    CONFIRMED_SKIP_STATUS,
    CONFIRMED_WRITE_STATUS,
    HARD_PRICE_MAX,
    HARD_PRICE_MIN,
    PRICE_DEVIATION_CONFIG,
    VALID_INVENTORY_STATUSES,
    check_price_deviation,
)


def test_rules_expose_business_safety_constants():
    assert HARD_PRICE_MIN == 1000
    assert HARD_PRICE_MAX == 10000
    assert PRICE_DEVIATION_CONFIG.abs_tolerance == 1000
    assert PRICE_DEVIATION_CONFIG.pct_tolerance == 0.20
    assert VALID_INVENTORY_STATUSES == frozenset({"充足", "告警", "缺货"})
    assert CONFIRMED_WRITE_STATUS == "已确认匹配"
    assert CONFIRMED_SKIP_STATUS == "已确认不更新"


def test_check_price_deviation_blocks_absolute_or_percent_breach():
    result = check_price_deviation(offline_price=5200, web_price=3400, label="螺纹")

    assert result is not None
    assert "螺纹线下价与网价偏差过大" in result
    assert "阈值=1000元/20%" in result


def test_check_price_deviation_allows_missing_reference():
    assert check_price_deviation(offline_price=5200, web_price=None, label="螺纹") is None

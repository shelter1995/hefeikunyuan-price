from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

HARD_PRICE_MIN = 1000
HARD_PRICE_MAX = 10000
VALID_INVENTORY_STATUSES = frozenset({"充足", "告警", "缺货"})

CONFIRMED_WRITE_STATUS = "已确认匹配"
CONFIRMED_SKIP_STATUS = "已确认不更新"


@dataclass(frozen=True)
class PriceDeviationConfig:
    abs_tolerance: int = 1000
    pct_tolerance: float = 0.20


PRICE_DEVIATION_CONFIG = PriceDeviationConfig()


def coerce_price(value: Any) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = str(value).replace(",", "").strip()
    match = re.search(r"(?<!\d)(\d{3,5})(?!\d)", text)
    return int(match.group(1)) if match else None


def check_price_deviation(
    offline_price: int | None,
    web_price: Any,
    label: str,
    config: PriceDeviationConfig = PRICE_DEVIATION_CONFIG,
) -> str | None:
    if offline_price is None:
        return None
    reference = coerce_price(web_price)
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

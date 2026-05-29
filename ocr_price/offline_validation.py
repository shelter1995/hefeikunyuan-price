from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import re

from .rules import HARD_PRICE_MAX, HARD_PRICE_MIN, VALID_INVENTORY_STATUSES


@dataclass
class OfflineValidationResult:
    is_valid: bool
    errors: list[str]


def _normalize_location(text: str) -> str:
    value = re.sub(r"\s+", "", text or "")
    for token in ("省", "市", "地区", "市场", "区域", "报价"):
        value = value.replace(token, "")
    return value


def _matches_location(target: str, location: str) -> bool:
    if not target:
        return True
    t = _normalize_location(target)
    loc = _normalize_location(location)
    if not t or not loc:
        return False
    return t in loc or loc in t


def _to_int_price(value: Any) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str):
        s = value.strip()
        if re.fullmatch(r"\d+", s):
            return int(s)
    return None


def _collect_inventory_items(payload: dict[str, Any]) -> list[Any]:
    items: list[Any] = []
    inventory = payload.get("inventory")
    if isinstance(inventory, list):
        items.extend(inventory)

    zh_inventory = payload.get("库存情况")
    if isinstance(zh_inventory, list):
        items.extend(zh_inventory)

    for key in ("vision_result", "_vision_result"):
        vision_result = payload.get(key)
        if not isinstance(vision_result, dict):
            continue
        vision_inventory = vision_result.get("库存情况")
        if isinstance(vision_inventory, list):
            items.extend(vision_inventory)
    return items


def validate_offline_payload(payload: dict[str, Any], target_location: str) -> OfflineValidationResult:
    errors: list[str] = []

    records_raw = payload.get("records")
    records = records_raw if isinstance(records_raw, list) else []
    if records_raw is not None and not isinstance(records_raw, list):
        errors.append("records must be a list when present")

    found_target = False
    for idx, record in enumerate(records):
        if not isinstance(record, dict):
            errors.append(f"records[{idx}] must be an object")
            continue
        record_location = str(record.get("location") or "")
        if _matches_location(target_location, record_location):
            found_target = True

        for field in ("rebar_price", "coil_price"):
            raw_price = record.get(field)
            if raw_price is None:
                continue
            price = _to_int_price(raw_price)
            if price is None:
                errors.append(f"records[{idx}].{field} must be an integer")
                continue
            if price < HARD_PRICE_MIN:
                errors.append(f"records[{idx}].{field} below hard minimum {HARD_PRICE_MIN}")
                continue
            if price > HARD_PRICE_MAX:
                errors.append(f"records[{idx}].{field} above hard maximum {HARD_PRICE_MAX}")

    if records and target_location and not found_target:
        errors.append("target location not found in records")

    for idx, item in enumerate(_collect_inventory_items(payload)):
        if not isinstance(item, dict):
            errors.append(f"inventory[{idx}] must be an object")
            continue
        status = str(item.get("状态") or item.get("status") or "").strip()
        if not status:
            errors.append(f"inventory[{idx}] missing status")
            continue
        if status not in VALID_INVENTORY_STATUSES:
            errors.append(f"inventory[{idx}] invalid status: {status}")

    return OfflineValidationResult(is_valid=not errors, errors=errors)

from __future__ import annotations

import json
import os
import re
from typing import Any

import requests


class SemanticAdjustmentError(RuntimeError):
    pass


MINIMAX_CHAT_URL = "https://api.minimaxi.com/v1/chat/completions"
MINIMAX_TEXT_MODEL = "MiniMax-M2.7"


def _extract_chat_text(payload: dict[str, Any]) -> str:
    choices = payload.get("choices")
    if isinstance(choices, list) and choices:
        first = choices[0]
        if isinstance(first, dict):
            message = first.get("message")
            if isinstance(message, dict) and isinstance(message.get("content"), str):
                return message["content"].strip()
            if isinstance(first.get("text"), str):
                return first["text"].strip()
    for key in ("text", "content", "result"):
        value = payload.get(key)
        if isinstance(value, str):
            return value.strip()
    return json.dumps(payload, ensure_ascii=False)


def _parse_json_from_text(text: str) -> dict[str, Any] | None:
    md = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if md:
        try:
            return json.loads(md.group(1).strip())
        except json.JSONDecodeError:
            pass
    obj = re.search(r"\{[\s\S]*\}", text)
    if obj:
        try:
            return json.loads(obj.group(0))
        except json.JSONDecodeError:
            pass
    return None


def _coerce_int_or_none(value: Any) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = str(value).strip()
    if text in {"", "null", "None", "无", "不适用"}:
        return None
    match = re.search(r"[-+]?\d+", text)
    return int(match.group(0)) if match else None


def _normalize_semantic_payload(payload: dict[str, Any], location: str) -> dict[str, Any] | None:
    applies = payload.get("applies_to_target", True)
    if applies is False:
        return None

    is_electronic = bool(payload.get("is_electronic_negotiation") or payload.get("电议"))
    is_direct = bool(payload.get("is_direct_price"))
    is_adjustment = bool(payload.get("is_adjustment", not is_direct))

    rebar_value = payload.get("rebar_price", payload.get("螺纹价格"))
    coil_value = payload.get("coil_price", payload.get("盘螺价格"))
    if is_adjustment:
        rebar_value = payload.get("rebar_adjustment", payload.get("螺纹调整", rebar_value))
        coil_value = payload.get("coil_adjustment", payload.get("盘螺调整", coil_value))

    rebar_price = _coerce_int_or_none(rebar_value)
    coil_price = _coerce_int_or_none(coil_value)
    if rebar_price is None and coil_price is None and not is_electronic:
        return None

    return {
        "location": str(payload.get("target_location") or payload.get("location") or location),
        "rebar_price": rebar_price,
        "coil_price": coil_price,
        "is_adjustment": is_adjustment,
        "is_direct_price": is_direct,
        "is_electronic_negotiation": is_electronic,
        "confidence": str(payload.get("confidence") or ""),
        "reason": str(payload.get("reason") or payload.get("依据") or ""),
    }


def build_supplement_adjustment_prompt(text: str, location: str, company: str | None = None) -> str:
    company_line = f"厂家：{company}\n" if company else ""
    return f"""\
你是建筑钢材报价业务助理。请理解下面的“补充报价”文本，把它转换成严格 JSON。

{company_line}目标地区：{location}
补充报价原文：
{text}

任务：
1. 判断原文是否给出了目标地区的螺纹/盘螺直接价格、调价、优惠、上浮、下浮或电议。
2. “优惠10出”“下10”表示调整值 -10；“上10”“加10”表示 +10；“表价出”“不优惠”表示调整值 0。
3. 如果首句或上下文同时列出多个地区，例如“徐钢蚌埠  滁州区域...”，要理解为这些地区同属后续规则的适用对象，不要把目标地区误归入“其他区域”。
4. 只输出 JSON，不要输出解释性文字。

JSON 格式：
{{
  "target_location": "{location}",
  "applies_to_target": true,
  "is_adjustment": true,
  "is_direct_price": false,
  "is_electronic_negotiation": false,
  "rebar_adjustment": -10,
  "coil_adjustment": 0,
  "rebar_price": null,
  "coil_price": null,
  "confidence": "high",
  "reason": "简短说明依据"
}}

字段规则：
- 如果是调整/优惠/上浮/下浮，填 rebar_adjustment 和 coil_adjustment；无变化填 0，未提及填 null。
- 如果是直接价格，is_adjustment=false, is_direct_price=true，并填 rebar_price/coil_price。
- 如果目标地区不适用，applies_to_target=false。
"""


def interpret_supplement_adjustment_with_llm(
    text: str,
    location: str,
    company: str | None = None,
    api_key: str | None = None,
    timeout: int = 60,
) -> dict[str, Any] | None:
    key = api_key or os.getenv("MINIMAX_API_KEY", "").strip()
    if not key:
        return None

    url = os.getenv("MINIMAX_TEXT_URL", MINIMAX_CHAT_URL).strip() or MINIMAX_CHAT_URL
    model = os.getenv("MINIMAX_TEXT_MODEL", MINIMAX_TEXT_MODEL).strip() or MINIMAX_TEXT_MODEL
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "你只做钢材补充报价语义结构化，输出严格 JSON。",
            },
            {
                "role": "user",
                "content": build_supplement_adjustment_prompt(text, location, company),
            },
        ],
        "temperature": 0.1,
        "max_completion_tokens": 800,
    }
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}

    try:
        response = requests.post(url, headers=headers, json=payload, timeout=timeout)
    except requests.RequestException as exc:
        raise SemanticAdjustmentError(f"MiniMax text API request failed: {exc}") from exc
    if not response.ok:
        raise SemanticAdjustmentError(
            f"MiniMax text API failed ({response.status_code}): {response.text[:300]}"
        )

    try:
        raw_payload = response.json()
    except ValueError as exc:
        raise SemanticAdjustmentError(f"MiniMax text API returned non-JSON: {response.text[:300]}") from exc

    raw_text = _extract_chat_text(raw_payload)
    parsed = _parse_json_from_text(raw_text)
    if parsed is None:
        raise SemanticAdjustmentError(f"Failed to parse semantic JSON: {raw_text[:300]}")
    return _normalize_semantic_payload(parsed, location)

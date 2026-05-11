from __future__ import annotations

import base64
import json
import os
import re
from pathlib import Path
from typing import Any

import requests


class MiniMaxVisionError(RuntimeError):
    pass


class MiniMaxVisionClient:
    """Client for MiniMax VLM (Vision-Language Model) API."""

    VLM_URL = "https://api.minimaxi.com/v1/coding_plan/vlm"

    def __init__(self, api_key: str | None = None, timeout: int = 120) -> None:
        self.api_key = api_key or os.getenv("MINIMAX_API_KEY", "").strip()
        if not self.api_key:
            raise MiniMaxVisionError(
                "Missing MINIMAX_API_KEY. Set env variable or pass api_key parameter."
            )
        self.timeout = timeout

    @classmethod
    def from_env(cls) -> MiniMaxVisionClient:
        return cls()

    def analyze_image(
        self,
        image_path: str | Path,
        prompt: str,
        save_raw_path: str | Path | None = None,
    ) -> str:
        path = Path(image_path)
        if not path.exists():
            raise MiniMaxVisionError(f"Image not found: {path}")

        file_bytes = path.read_bytes()
        b64 = base64.b64encode(file_bytes).decode("utf-8")

        suffix = path.suffix.lower()
        mime = "image/jpeg"
        if suffix == ".png":
            mime = "image/png"
        elif suffix == ".webp":
            mime = "image/webp"
        elif suffix == ".gif":
            mime = "image/gif"

        data_url = f"data:{mime};base64,{b64}"

        payload = {
            "prompt": prompt,
            "image_url": data_url,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

        resp = requests.post(
            self.VLM_URL,
            headers=headers,
            json=payload,
            timeout=self.timeout,
        )
        if not resp.ok:
            raise MiniMaxVisionError(
                f"MiniMax VLM API failed ({resp.status_code}): {resp.text[:500]}"
            )

        data = resp.json()
        if save_raw_path:
            raw_path = Path(save_raw_path)
            raw_path.parent.mkdir(parents=True, exist_ok=True)
            raw_path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
            )

        text = _extract_vlm_text(data)
        if not text:
            raise MiniMaxVisionError(
                f"Empty response from MiniMax VLM: {json.dumps(data, ensure_ascii=False)[:500]}"
            )
        return text


def _extract_vlm_text(data: dict[str, Any]) -> str:
    """Extract text content from MiniMax VLM response."""
    for key in ("text", "content", "message", "result", "data"):
        val = data.get(key)
        if isinstance(val, str) and val.strip():
            return val.strip()
        if isinstance(val, dict):
            for inner_key in ("text", "content", "message"):
                inner_val = val.get(inner_key)
                if isinstance(inner_val, str) and inner_val.strip():
                    return inner_val.strip()
        if isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    for inner_key in ("text", "content", "message"):
                        inner_val = item.get(inner_key)
                        if isinstance(inner_val, str) and inner_val.strip():
                            return inner_val.strip()
                elif isinstance(item, str) and item.strip():
                    return item.strip()
    return json.dumps(data, ensure_ascii=False)


def _parse_json_from_text(text: str) -> dict[str, Any] | None:
    """Parse JSON from text, handling markdown code blocks."""
    # Try to extract JSON from markdown code block
    md_match = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if md_match:
        try:
            return json.loads(md_match.group(1).strip())
        except json.JSONDecodeError:
            pass

    # Try to find JSON object directly
    json_match = re.search(r"\{[\s\S]*\}", text)
    if json_match:
        try:
            return json.loads(json_match.group())
        except json.JSONDecodeError:
            pass

    return None


QUOTE_ANALYSIS_PROMPT_TEMPLATE = """\
你是一个钢材报价表分析专家。请仔细分析这张报价图片，提取以下信息并以严格的JSON格式输出。

【图片布局说明】
报价图片通常包含以下区域：
1. 顶部：厂家名称、报价日期、联系电话
2. 顶部/左侧：库存规格表（标注"厂发规格"、"库存"、"抗震"等字样）
3. 中部/右侧：价格表（按地区分列，如合肥、蚌埠等）
4. 底部：加价规则、备注说明

【钢材类型识别】
- 螺纹钢：有9米、12米长度规格，标注"抗震"、"螺纹"、"HRB400E"等
- 盘螺：盘状卷材，规格通常6E、8E、10E、12E
- 线材：HPB300，规格通常6、8、10
- 圆钢：HPB300圆钢，规格通常16、18、20

【价格提取规则 - 非常重要】
1. 价格仅取自表格中直接填写的数字
2. 如果某城市行在表格中不存在或单元格为空白 → 价格填null
3. 如果价格标注为"电议" → 价格填null，电议标记设为true
4. 不要猜测、不要推算、不要从加价规则计算价格
5. 表格列类型根据表头判断：可能是"螺纹+盘螺"、"盘螺+线材"、"仅螺纹"等

【库存提取规则 - 非常重要】
库存区域通常在图片顶部或左侧，有"厂发规格"、"库存"、"抗震"等标题。
请逐行识别，不要遗漏任何规格！

状态判断标准（绝对不可违反）：
1. "无货"、"无" = 缺货（红色）
2. "极少" = 告警（黄色）
3. "少" = 告警（黄色）
4. 规格后带"配"字（如"22E配"、"25E配"、"8E配"）= 告警（黄色），代表需调配，库存不足
5. 【强制规则】规格后面跟有具体件数（如"82件"、"166件"、"406件"、"36件"、"19件"）= 必须标为告警（黄色）！件数代表库存数量有限，不是充足！
   - 错误示例："16（82件）"标为"充足" → 这是严重错误
   - 正确示例："16（82件）"标为"告警"
6. 规格数字后无任何说明文字 = 充足（蓝色），代表库存正常

【最容易犯的错误】
- 将"XX件"错误判断为"充足" → 这是绝对错误的！
- "XX件"是库存有限的明确标志，必须标为告警（黄色）

【库存提取示例 - 请严格按此逻辑】

例1（淮南宏泰式）："螺纹 抗震9米：10无货、12、14、16"
→ 9米螺纹10：缺货（10无货）
→ 9米螺纹12：充足（无说明）
→ 9米螺纹14：充足（无说明）
→ 9米螺纹16：充足（无说明）

例2（徐钢式）："12E、14E极少、16E、18E极少、20E、22E、25E"
→ 12E：充足（无说明）
→ 14E：告警（14E极少）
→ 16E：充足（无说明）
→ 18E：告警（18E极少）
→ 20E：充足（无说明）
→ 22E：充足（无说明）
→ 25E：充足（无说明）

例3（长江式）："12E、14E、16E、18E、20E、22E配、25E配"
→ 12E：充足（无说明）
→ 14E：充足（无说明）
→ 16E：充足（无说明）
→ 18E：充足（无说明）
→ 20E：充足（无说明）
→ 22E：告警（22E配）
→ 25E：告警（25E配）

例4（贵航圆钢式）："16（82件）、18（166件）、20（406件）"
→ 圆钢16：告警（82件）
→ 圆钢18：告警（166件）
→ 圆钢20：告警（406件）

例5（贵航螺纹式）："9米抗震 10-12-14"
→ 9米螺纹10：充足（无说明）
→ 9米螺纹12：充足（无说明）
→ 9米螺纹14：充足（无说明）

例6（桂鑫式）："9米抗震 12-14-16-20-22-25-28"
→ 9米螺纹12：充足（无说明）
→ 9米螺纹14：充足（无说明）
→ 9米螺纹16：充足（无说明）
→ 9米螺纹20：充足（无说明）
→ 9米螺纹22：充足（无说明）
→ 9米螺纹25：充足（无说明）
→ 9米螺纹28：充足（无说明）

【特别注意】
1. 如果图片中无目标城市（合肥、蚌埠）的数据，价格必须填null
2. 如果某规格后无任何文字说明，它就是充足状态
3. "配"字是库存不足的标志，必须标为告警
4. 件数（XX件）是库存有限的标志，必须标为告警
5. 不要遗漏任何规格，要逐行完整读取

请输出以下JSON（不要输出任何其他文字）：
{
  "厂家名称": "从文件名识别的厂家名称",
  "报价日期": "yyyy-MM-dd格式，如无法识别则填null",
  "表格列类型": "螺纹+盘螺 或 盘螺+线材 或 仅螺纹 等",
  "{city1}": {
    "螺纹": 数字或null,
    "盘螺": 数字或null,
    "螺纹为电议": true或false,
    "盘螺为电议": true或false
  },
  "{city2}": {
    "螺纹": 数字或null,
    "盘螺": 数字或null,
    "螺纹为电议": true或false,
    "盘螺为电议": true或false
  },
  "库存情况": [
    {"规格": "如9米螺纹12E", "状态": "充足/告警/缺货", "原始描述": "如极少、无货、82件、22E配等"}
  ],
  "备注": "任何特殊说明，如加价规则、补充说明等"
}"""


def build_analysis_prompt(target_cities: list[str]) -> str:
    """Build the analysis prompt for given target cities."""
    city1 = target_cities[0] if len(target_cities) >= 1 else "合肥"
    city2 = target_cities[1] if len(target_cities) >= 2 else "蚌埠"
    return QUOTE_ANALYSIS_PROMPT_TEMPLATE.replace("{city1}", city1).replace("{city2}", city2)


def analyze_quote_image(
    image_path: str | Path,
    target_cities: list[str] | None = None,
    api_key: str | None = None,
    save_raw_path: str | Path | None = None,
) -> dict[str, Any]:
    """Analyze a quote image using MiniMax VLM and return structured data."""
    if target_cities is None:
        target_cities = ["合肥", "蚌埠"]

    client = MiniMaxVisionClient(api_key=api_key)
    prompt = build_analysis_prompt(target_cities)

    raw_text = client.analyze_image(
        image_path, prompt=prompt, save_raw_path=save_raw_path
    )

    result = _parse_json_from_text(raw_text)
    if result is None:
        raise MiniMaxVisionError(
            f"Failed to parse JSON from MiniMax VLM response: {raw_text[:500]}"
        )

    result["_source"] = {
        "image_path": str(image_path),
        "raw_text": raw_text,
    }
    return result


def analyze_quote_image_with_retry(
    image_path: str | Path,
    target_cities: list[str] | None = None,
    api_key: str | None = None,
    max_retries: int = 3,
    save_raw_path: str | Path | None = None,
) -> dict[str, Any]:
    """Analyze a quote image with multiple retries to improve reliability.
    
    MiniMax VLM output can be non-deterministic. This function calls the API
    multiple times and merges results to reduce false negatives (missing prices).
    """
    if target_cities is None:
        target_cities = ["合肥", "蚌埠"]

    all_results: list[dict[str, Any]] = []
    all_raw_texts: list[str] = []

    for attempt in range(max_retries):
        try:
            result = analyze_quote_image(
                image_path=image_path,
                target_cities=target_cities,
                api_key=api_key,
                save_raw_path=None,  # Don't save raw for retries
            )
            all_results.append(result)
            source = result.get("_source", {})
            all_raw_texts.append(source.get("raw_text", ""))
        except MiniMaxVisionError:
            if attempt == max_retries - 1 and not all_results:
                raise
            continue

    if not all_results:
        raise MiniMaxVisionError(f"All {max_retries} attempts failed for {image_path}")

    # Merge results
    merged = _merge_vision_results(all_results, target_cities)
    merged["_source"] = {
        "image_path": str(image_path),
        "raw_texts": all_raw_texts,
        "attempt_count": len(all_results),
    }

    if save_raw_path:
        raw_path = Path(save_raw_path)
        raw_path.parent.mkdir(parents=True, exist_ok=True)
        raw_path.write_text(
            json.dumps({
                "merged_result": merged,
                "all_results": all_results,
            }, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    return merged


def _merge_vision_results(
    results: list[dict[str, Any]], target_cities: list[str]
) -> dict[str, Any]:
    """Merge multiple vision results, preferring non-null values."""
    if not results:
        return {}

    # Base on first result
    merged = dict(results[0])
    merged.pop("_source", None)

    # Merge company name: prefer non-empty
    for r in results[1:]:
        company = r.get("厂家名称")
        if company and not merged.get("厂家名称"):
            merged["厂家名称"] = company

    # Merge quote date: prefer non-null
    for r in results[1:]:
        date = r.get("报价日期")
        if date and not merged.get("报价日期"):
            merged["报价日期"] = date

    # Merge city prices: prefer non-null values
    for city in target_cities:
        city_key = city
        if city_key not in merged:
            merged[city_key] = {}
        for r in results[1:]:
            if city_key not in r:
                continue
            for field in ["螺纹", "盘螺", "螺纹为电议", "盘螺为电议"]:
                if field in r[city_key]:
                    # Prefer non-null values
                    if r[city_key][field] is not None:
                        if city_key not in merged:
                            merged[city_key] = {}
                        merged[city_key][field] = r[city_key][field]

    # Merge inventory: collect all unique items
    all_inventory: list[dict[str, Any]] = []
    seen_inventory = set()
    for r in results:
        for item in r.get("库存情况", []):
            key = (item.get("规格", ""), item.get("原始描述", ""))
            if key not in seen_inventory:
                seen_inventory.add(key)
                all_inventory.append(item)
    merged["库存情况"] = all_inventory

    # Merge remarks: collect all unique
    all_remarks: list[str] = []
    seen_remarks = set()
    for r in results:
        remark = r.get("备注", "")
        if remark and remark not in seen_remarks:
            seen_remarks.add(remark)
            all_remarks.append(remark)
    merged["备注"] = "; ".join(all_remarks) if all_remarks else ""

    return merged


def convert_to_ocr_format(
    vision_result: dict[str, Any], target_cities: list[str] | None = None
) -> dict[str, Any]:
    """Convert MiniMax vision result to the existing OCR JSON format for compatibility."""
    source = vision_result.get("_source", {})
    image_path = source.get("image_path", "")

    city_keys: list[str] = []
    if target_cities:
        city_keys = list(target_cities)
    else:
        for key, value in vision_result.items():
            if not isinstance(value, dict):
                continue
            if any(
                field in value
                for field in ("螺纹", "盘螺", "螺纹为电议", "盘螺为电议")
            ):
                city_keys.append(key)

    # Build records from city prices
    records: list[dict[str, Any]] = []
    for city_key in city_keys:
        city_data = vision_result.get(city_key)
        if not isinstance(city_data, dict):
            continue
        rebar = city_data.get("螺纹")
        coil = city_data.get("盘螺")
        rebar_elec = city_data.get("螺纹为电议", False)
        coil_elec = city_data.get("盘螺为电议", False)

        # Skip if both null and not electronic negotiation
        if rebar is None and coil is None and not rebar_elec and not coil_elec:
            continue

        records.append({
            "region_title": None,
            "location": city_key,
            "rebar_price": rebar,
            "coil_price": coil,
            "rebar_raw": "电议" if rebar_elec else str(rebar) if rebar is not None else None,
            "coil_raw": "电议" if coil_elec else str(coil) if coil is not None else None,
            "group_index": None,
            "source_row_index": None,
            "header_row_index": None,
        })

    target_location: str | None = None
    if target_cities:
        target_location = target_cities[0]
    elif city_keys:
        target_location = city_keys[0]

    return {
        "meta": {
            "input_file": image_path,
            "provider": "minimax_vision",
            "target_location": target_location,
            "generated_at": __import__("datetime").datetime.now().isoformat(timespec="seconds"),
            "record_count": len(records),
        },
        "company": vision_result.get("厂家名称"),
        "quote_date": vision_result.get("报价日期"),
        "header_row_index": None,
        "group_count": None,
        "records": records,
        "_vision_result": vision_result,
    }


def analyze_quote_image_to_ocr_format(
    image_path: str | Path,
    target_cities: list[str] | None = None,
    api_key: str | None = None,
    save_raw_path: str | Path | None = None,
    use_retry: bool = True,
) -> dict[str, Any]:
    """Analyze image and convert result to OCR-compatible format.
    
    If use_retry is True (default), uses multiple attempts to improve reliability.
    """
    if use_retry:
        vision_result = analyze_quote_image_with_retry(
            image_path=image_path,
            target_cities=target_cities,
            api_key=api_key,
            save_raw_path=save_raw_path,
        )
    else:
        vision_result = analyze_quote_image(
            image_path=image_path,
            target_cities=target_cities,
            api_key=api_key,
            save_raw_path=save_raw_path,
        )
    return convert_to_ocr_format(vision_result, target_cities=target_cities)

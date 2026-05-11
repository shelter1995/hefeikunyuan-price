from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any

import requests


class PaddleOCRApiError(RuntimeError):
    pass


class PaddleOCRApiClient:
    """Thin HTTP client for PaddleOCR layout parsing style APIs."""

    def __init__(
        self,
        base_url: str,
        api_key: str | None = None,
        timeout: int = 90,
        auth_scheme: str = "token",
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.timeout = timeout
        self.auth_scheme = auth_scheme.lower().strip() or "token"

    @classmethod
    def from_env(cls) -> "PaddleOCRApiClient":
        base_url = os.getenv("PADDLEOCR_BASE_URL", "").strip()
        if not base_url:
            raise PaddleOCRApiError(
                "Missing PADDLEOCR_BASE_URL. "
                "Example: http://localhost:8080/layout-parsing"
            )
        api_key = os.getenv("PADDLEOCR_API_KEY", "").strip() or None
        auth_scheme = os.getenv("PADDLEOCR_AUTH_SCHEME", "token").strip()
        return cls(base_url=base_url, api_key=api_key, auth_scheme=auth_scheme)

    def _build_headers(self) -> dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            if self.auth_scheme == "bearer":
                headers["Authorization"] = f"Bearer {self.api_key}"
            elif self.auth_scheme == "token":
                headers["Authorization"] = f"token {self.api_key}"
            else:
                headers["Authorization"] = self.api_key
        return headers

    @staticmethod
    def _detect_file_type(path: Path) -> int:
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            return 0
        return 1

    def predict_layout(
        self,
        file_path: str | Path,
        save_raw_path: str | Path | None = None,
        options: dict[str, Any] | None = None,
    ) -> Any:
        path = Path(file_path)
        if not path.exists():
            raise PaddleOCRApiError(f"Input not found: {path}")

        file_bytes = path.read_bytes()
        payload = {
            "file": base64.b64encode(file_bytes).decode("utf-8"),
            "fileType": self._detect_file_type(path),
        }
        if options:
            payload.update(options)
        resp = requests.post(
            self.base_url,
            headers=self._build_headers(),
            json=payload,
            timeout=self.timeout,
        )
        if not resp.ok:
            raise PaddleOCRApiError(
                f"PaddleOCR API request failed ({resp.status_code}): {resp.text[:300]}"
            )
        data = resp.json()
        if save_raw_path:
            raw_path = Path(save_raw_path)
            raw_path.parent.mkdir(parents=True, exist_ok=True)
            raw_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return data


def _as_float(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _extract_bbox(item: dict[str, Any]) -> list[tuple[float, float]] | None:
    points = item.get("points") or item.get("bbox") or item.get("box")
    if isinstance(points, list) and points:
        parsed: list[tuple[float, float]] = []
        for pt in points:
            if isinstance(pt, (list, tuple)) and len(pt) >= 2:
                x = _as_float(pt[0])
                y = _as_float(pt[1])
                if x is not None and y is not None:
                    parsed.append((x, y))
        if parsed:
            return parsed

    if all(k in item for k in ("x1", "y1", "x2", "y2")):
        x1 = _as_float(item.get("x1"))
        y1 = _as_float(item.get("y1"))
        x2 = _as_float(item.get("x2"))
        y2 = _as_float(item.get("y2"))
        if None not in (x1, y1, x2, y2):
            return [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
    return None


def collect_text_boxes(payload: Any) -> list[dict[str, Any]]:
    """
    Normalize different PaddleOCR JSON layouts into:
    [{"text": "...", "score": 0.99, "bbox": [(x,y)...]}]
    """
    boxes: list[dict[str, Any]] = []

    text_keys = ("text", "rec_text", "transcription", "label", "content")
    score_keys = ("score", "rec_score", "confidence", "prob")

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            text_val = None
            for key in text_keys:
                value = node.get(key)
                if isinstance(value, str) and value.strip():
                    text_val = value.strip()
                    break
            if text_val:
                score = None
                for sk in score_keys:
                    s = _as_float(node.get(sk))
                    if s is not None:
                        score = s
                        break
                boxes.append(
                    {
                        "text": text_val,
                        "score": score,
                        "bbox": _extract_bbox(node),
                    }
                )
            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(payload)
    seen: set[tuple[str, str | None, float | None]] = set()
    deduped: list[dict[str, Any]] = []
    for item in boxes:
        bbox = item.get("bbox")
        bbox_key = None
        if bbox:
            bbox_key = "|".join(f"{p[0]:.2f},{p[1]:.2f}" for p in bbox)
        key = (item["text"], bbox_key, item.get("score"))
        if key not in seen:
            seen.add(key)
            deduped.append(item)
    return deduped


def extract_markdown_texts(payload: Any) -> list[str]:
    texts: list[str] = []
    if not isinstance(payload, dict):
        return texts
    result = payload.get("result")
    if not isinstance(result, dict):
        return texts
    layout_results = result.get("layoutParsingResults")
    if not isinstance(layout_results, list):
        return texts
    for item in layout_results:
        if not isinstance(item, dict):
            continue
        markdown = item.get("markdown")
        if isinstance(markdown, dict):
            text = markdown.get("text")
            if isinstance(text, str) and text.strip():
                texts.append(text.strip())
    return texts


def extract_table_htmls(payload: Any) -> list[str]:
    htmls: list[str] = []
    if not isinstance(payload, dict):
        return htmls
    result = payload.get("result")
    if not isinstance(result, dict):
        return htmls
    layout_results = result.get("layoutParsingResults")
    if not isinstance(layout_results, list):
        return htmls
    for page in layout_results:
        if not isinstance(page, dict):
            continue
        pruned = page.get("prunedResult")
        if not isinstance(pruned, dict):
            continue
        parsing_list = pruned.get("parsing_res_list")
        if not isinstance(parsing_list, list):
            continue
        for block in parsing_list:
            if not isinstance(block, dict):
                continue
            block_content = block.get("block_content")
            if isinstance(block_content, str) and "<table" in block_content.lower():
                htmls.append(block_content)
    return htmls

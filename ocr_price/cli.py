from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from .env_loader import load_env_file
from .minimax_vision import (
    MiniMaxVisionError,
    analyze_quote_image_to_ocr_format,
)
from .parser import (
    parse_price_lines_from_text,
    parse_inventory_from_text,
)


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Extract steel quote prices and inventory from text files or MiniMax vision image/PDF analysis."
    )
    p.add_argument("--input", required=True, help="Input file path (.jpg/.png/.txt)")
    p.add_argument("--location", help="Optional target location filter, e.g. 蚌埠")
    p.add_argument(
        "--provider",
        choices=("text", "minimax"),
        default="minimax",
        help="Processing mode. Text files use text parser; image/PDF files use MiniMax VLM vision.",
    )
    p.add_argument(
        "--env-file",
        default=".env",
        help="Environment file path (default: .env). Set empty to disable.",
    )
    p.add_argument(
        "--output",
        help="Output JSON path. Default: 运行产物/ocr价格提取_<input_stem>.json",
    )
    p.add_argument("--raw-output", help="Optional raw API JSON output path")
    return p


def _default_output(input_path: Path) -> Path:
    return Path("运行产物") / f"ocr价格提取_{input_path.stem}.json"


def main() -> int:
    args = _build_parser().parse_args()
    if args.env_file:
        load_env_file(args.env_file)

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"Input not found: {input_path}")

    provider = args.provider
    if input_path.suffix.lower() == ".txt":
        provider = "text"
    elif input_path.suffix.lower() == ".pdf":
        raise SystemExit("PDF input is not supported. 请先将PDF转换为图片后再识别。")
    elif input_path.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp"}:
        provider = "minimax"

    if provider == "text":
        raw_text = input_path.read_text(encoding="utf-8", errors="ignore")
        parsed = parse_price_lines_from_text(
            raw_text,
            target_location=args.location,
        )
        # Also extract inventory from text files
        inventory = parse_inventory_from_text(raw_text)
        if inventory:
            parsed["inventory"] = inventory
    elif provider == "minimax":
        try:
            result = analyze_quote_image_to_ocr_format(
                image_path=input_path,
                target_cities=[args.location] if args.location else None,
                save_raw_path=args.raw_output,
            )
            parsed = {
                "company": result.get("company"),
                "quote_date": result.get("quote_date"),
                "header_row_index": result.get("header_row_index"),
                "group_count": result.get("group_count"),
                "records": result.get("records", []),
            }
            # Include vision result for inventory extraction
            vision_result = result.get("_vision_result", {})
            if vision_result:
                parsed["_vision_result"] = vision_result
        except MiniMaxVisionError as exc:
            raise SystemExit(str(exc)) from exc
    else:
        raise SystemExit(f"Unknown provider: {provider}")

    output = {
        "meta": {
            "input_file": str(input_path),
            "provider": provider,
            "target_location": args.location,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "record_count": len(parsed.get("records", [])),
        },
        **parsed,
    }

    output_path = Path(args.output) if args.output else _default_output(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Output: {output_path}")
    print(f"Records: {len(parsed.get('records', []))}")
    if parsed.get("company"):
        print(f"Company: {parsed['company']}")
    if parsed.get("quote_date"):
        print(f"Quote Date: {parsed['quote_date']}")
    if parsed.get("inventory"):
        print(f"Inventory items: {len(parsed['inventory'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

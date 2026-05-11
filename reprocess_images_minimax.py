#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使用MiniMax视觉识别重新处理所有图片报价文件，生成兼容的OCR JSON"""

import json
from pathlib import Path
from datetime import datetime

from ocr_price.env_loader import load_env_file
from ocr_price.minimax_vision import (
    MiniMaxVisionError,
    analyze_quote_image_to_ocr_format,
)

load_env_file(".env")


def process_all_images(
    input_dir: str = "线下报价",
    output_dir: str = "运行产物",
    target_cities: list[str] | None = None,
) -> list[Path]:
    if target_cities is None:
        target_cities = ["合肥", "蚌埠"]

    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    image_files = sorted(
        list(input_path.glob("*.jpg")) + list(input_path.glob("*.png"))
    )

    print(f"找到 {len(image_files)} 张图片文件")
    print("=" * 60)

    generated_files: list[Path] = []
    failed_files: list[str] = []

    for img in image_files:
        print(f"\n处理: {img.name}")
        try:
            result = analyze_quote_image_to_ocr_format(
                image_path=img,
                target_cities=target_cities,
                save_raw_path=output_path / f"minimax_raw_{img.stem}.json",
            )

            out_file = output_path / f"ocr价格提取_{img.stem}.json"
            out_file.write_text(
                json.dumps(result, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            generated_files.append(out_file)

            # Print summary
            company = result.get("company") or "未知"
            quote_date = result.get("quote_date") or "未知"
            records = result.get("records", [])
            vision = result.get("_vision_result", {})
            inventory = vision.get("库存情况", [])

            print(f"  [OK] 厂家: {company}, 日期: {quote_date}")
            print(f"       记录数: {len(records)}, 库存项: {len(inventory)}")
            for rec in records:
                loc = rec.get("location", "")
                rebar = rec.get("rebar_price")
                coil = rec.get("coil_price")
                rebar_raw = rec.get("rebar_raw")
                if rebar_raw == "电议":
                    print(f"       {loc}: 螺纹=电议, 盘螺={coil}")
                else:
                    print(f"       {loc}: 螺纹={rebar}, 盘螺={coil}")

        except MiniMaxVisionError as e:
            print(f"  [ERROR] MiniMax视觉识别失败: {e}")
            failed_files.append(img.name)
        except Exception as e:
            print(f"  [ERROR] 处理失败: {e}")
            failed_files.append(img.name)

    print("\n" + "=" * 60)
    print(f"处理完成: {len(generated_files)} 成功, {len(failed_files)} 失败")
    if failed_files:
        print(f"失败文件: {', '.join(failed_files)}")
    print("=" * 60)

    return generated_files


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="使用MiniMax视觉识别处理报价图片")
    p.add_argument("--input-dir", default="线下报价", help="输入图片目录")
    p.add_argument("--output-dir", default="运行产物", help="输出目录")
    p.add_argument(
        "--cities",
        nargs="+",
        default=["合肥", "蚌埠"],
        help="目标城市列表",
    )
    args = p.parse_args()

    process_all_images(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        target_cities=args.cities,
    )

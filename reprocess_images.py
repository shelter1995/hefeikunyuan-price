#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使用PaddleOCR重新处理所有图片报价文件"""

import json
from pathlib import Path
from ocr_price.paddle_api import PaddleOCRApiClient, PaddleOCRApiError, extract_markdown_texts, extract_table_htmls
from ocr_price.env_loader import load_env_file
from ocr_price.parser import parse_html_price_records, parse_markdown_price_records, parse_price_lines_from_text
from ocr_price.paddle_api import collect_text_boxes
from datetime import datetime

# 加载环境配置
load_env_file(".env")
client = PaddleOCRApiClient.from_env()

# 查找所有图片文件
offline_quotes_dir = Path("线下报价")
image_files = list(offline_quotes_dir.glob("**/*.jpg")) + list(offline_quotes_dir.glob("**/*.png"))

print(f"找到 {len(image_files)} 张图片文件")

# OCR options
options = {
    "useDocOrientationClassify": False,
    "useDocUnwarping": False,
    "useTextlineOrientation": False,
    "useChartRecognition": False,
}

# 处理每张图片
for image_file in image_files:
    print(f"\n{'='*60}")
    print(f"处理图片: {image_file}")
    print(f"{'='*60}")
    
    try:
        # 调用PaddleOCR API
        print("正在调用PaddleOCR API...")
        payload = client.predict_layout(image_file, options=options)
        
        # 尝试多种解析方式
        print("正在解析价格信息...")
        
        # 方式1: 从text boxes解析
        text_boxes = collect_text_boxes(payload)
        parsed = None
        
        try:
            parsed = parse_html_price_records(text_boxes, target_location=None)
        except Exception:
            pass
        
        if not parsed or not parsed.get("records"):
            # 方式2: 从HTML表格解析
            html_tables = extract_table_htmls(payload)
            for html in html_tables:
                try:
                    current = parse_html_price_records(html, target_location=None)
                    if current.get("records"):
                        parsed = current
                        print(f"从HTML表格提取到 {len(current['records'])} 条记录")
                        break
                except Exception:
                    pass
        
        if not parsed or not parsed.get("records"):
            # 方式3: 从markdown文本解析
            md_texts = extract_markdown_texts(payload)
            for md in md_texts:
                try:
                    current = parse_markdown_price_records(md, target_location=None)
                    if current.get("records"):
                        parsed = current
                        print(f"从Markdown文本提取到 {len(current['records'])} 条记录")
                        break
                except Exception:
                    pass
        
        if not parsed or not parsed.get("records"):
            # 方式4: 从纯文本解析
            plain_text = "\n".join(item["text"] for item in text_boxes if item.get("text"))
            try:
                parsed = parse_price_lines_from_text(plain_text, target_location=None)
                if parsed.get("records"):
                    print(f"从纯文本提取到 {len(parsed['records'])} 条记录")
            except Exception as e:
                print(f"纯文本解析失败: {e}")
                parsed = {"company": None, "quote_date": None, "records": []}
        
        # 构建输出结果
        output = {
            "meta": {
                "input_file": str(image_file),
                "provider": "paddleocr_api",
                "target_location": None,
                "generated_at": datetime.now().isoformat(timespec="seconds"),
                "record_count": len(parsed.get("records", [])),
            },
            "company": parsed.get("company"),
            "quote_date": parsed.get("quote_date"),
            "header_row_index": parsed.get("header_row_index"),
            "group_count": parsed.get("group_count"),
            "records": parsed.get("records", []),
        }
        
        # 保存结果
        output_path = Path("运行产物") / f"ocr价格提取_{image_file.stem}.json"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
        
        print(f"[OK] 结果已保存到: {output_path}")
        print(f"  厂家: {parsed.get('company') or '未知'}")
        print(f"  报价日期: {parsed.get('quote_date') or '未知'}")
        print(f"  记录数: {len(parsed.get('records', []))}")
        
    except PaddleOCRApiError as e:
        print(f"[ERROR] PaddleOCR API调用失败: {e}")
    except Exception as e:
        print(f"[ERROR] 处理失败: {e}")
        import traceback
        traceback.print_exc()

print(f"\n{'='*60}")
print("所有图片处理完成！")
print(f"{'='*60}")
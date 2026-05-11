#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试PaddleOCR layout-parsing API（使用.env中的配置）"""

from pathlib import Path
from ocr_price.paddle_api import PaddleOCRApiClient, PaddleOCRApiError, extract_markdown_texts, extract_table_htmls

# 使用.env中的配置
from ocr_price.env_loader import load_env_file
load_env_file(".env")

# 从环境变量创建客户端
client = PaddleOCRApiClient.from_env()

print(f"API URL: {client.base_url}")
print(f"API Key: {client.api_key[:20]}..." if client.api_key else "No API Key")

# 查找一个测试图片（线下报价文件夹中的图片）
offline_quotes_dir = Path("线下报价")
test_images = list(offline_quotes_dir.glob("**/*.jpg")) + list(offline_quotes_dir.glob("**/*.png"))

if not test_images:
    print("未找到测试图片！")
else:
    print(f"\n找到 {len(test_images)} 张图片，选择第一张进行测试...")
    test_image = test_images[0]
    print(f"测试图片: {test_image}")
    
    try:
        # 调用PaddleOCR API - 使用CLI中相同的options
        print("正在调用PaddleOCR API...")
        options = {
            "useDocOrientationClassify": False,
            "useDocUnwarping": False,
            "useTextlineOrientation": False,
            "useChartRecognition": False,
        }
        result = client.predict_layout(test_image, options=options)
        
        print(f"\nAPI返回结果结构:")
        print(f"- 键: {list(result.keys())}")
        
        # 提取markdown文本
        markdowns = extract_markdown_texts(result)
        print(f"\n提取到 {len(markdowns)} 个markdown文本块")
        if markdowns:
            print("第一个markdown块:")
            print(markdowns[0][:500])
        
        # 提取表格HTML
        tables = extract_table_htmls(result)
        print(f"\n提取到 {len(tables)} 个表格HTML")
        if tables:
            print("第一个表格HTML（前500字符）:")
            print(tables[0][:500])
        
        print("\nPaddleOCR API测试成功！")
        
        # 保存完整结果
        output_path = Path("运行产物") / f"paddleocr_result_{test_image.stem}.json"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        import json
        output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"完整结果已保存到: {output_path}")
        
    except PaddleOCRApiError as e:
        print(f"\nPaddleOCR API调用失败: {e}")
    except Exception as e:
        print(f"\n未知错误: {e}")
        import traceback
        traceback.print_exc()
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试PaddleOCR job-based API"""

import base64
import json
import time
from pathlib import Path

import requests

# 配置PaddleOCR API
JOB_URL = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs"
TOKEN = "20829c78de212f8a7348f0915b00b2d07007d574"

def submit_ocr_job(image_path: str, token: str) -> str:
    """提交OCR任务，返回job_id"""
    headers = {
        "Authorization": f"token {token}",
        "Content-Type": "application/json"
    }
    
    file_bytes = Path(image_path).read_bytes()
    payload = {
        "file": base64.b64encode(file_bytes).decode("utf-8"),
        "fileType": 1,  # 1 for image
    }
    
    resp = requests.post(JOB_URL, headers=headers, json=payload, timeout=90)
    if not resp.ok:
        raise Exception(f"Job submission failed ({resp.status_code}): {resp.text}")
    
    data = resp.json()
    print(f"Job submission response: {json.dumps(data, ensure_ascii=False, indent=2)}")
    
    if "jobId" not in data:
        raise Exception(f"No jobId in response: {data}")
    
    return data["jobId"]

def get_job_status(job_id: str, token: str) -> dict:
    """获取OCR任务状态"""
    headers = {
        "Authorization": f"token {token}",
    }
    
    status_url = f"{JOB_URL}/{job_id}"
    resp = requests.get(status_url, headers=headers, timeout=30)
    
    if not resp.ok:
        raise Exception(f"Status check failed ({resp.status_code}): {resp.text}")
    
    return resp.json()

def wait_for_job_completion(job_id: str, token: str, max_wait_seconds: int = 300, poll_interval: int = 5) -> dict:
    """等待OCR任务完成"""
    start_time = time.time()
    
    while time.time() - start_time < max_wait_seconds:
        status_data = get_job_status(job_id, token)
        print(f"Job status: {status_data.get('status', 'unknown')}")
        
        if status_data.get("status") == "completed":
            return status_data
        elif status_data.get("status") == "failed":
            raise Exception(f"Job failed: {status_data}")
        
        time.sleep(poll_interval)
    
    raise Exception(f"Job timeout after {max_wait_seconds} seconds")

# 查找测试图片
offline_quotes_dir = Path("线下报价")
test_images = list(offline_quotes_dir.glob("**/*.jpg")) + list(offline_quotes_dir.glob("**/*.png"))

if not test_images:
    print("未找到测试图片！")
else:
    print(f"找到 {len(test_images)} 张图片，选择第一张进行测试...")
    test_image = test_images[0]
    print(f"测试图片: {test_image}")
    
    try:
        # 1. 提交OCR任务
        print("正在提交OCR任务...")
        job_id = submit_ocr_job(str(test_image), TOKEN)
        print(f"任务ID: {job_id}")
        
        # 2. 等待任务完成
        print("等待任务完成...")
        result = wait_for_job_completion(job_id, TOKEN)
        
        print(f"\n任务完成！")
        print(f"结果结构: {list(result.keys())}")
        
        # 保存完整结果
        output_path = Path("运行产物") / f"paddleocr_result_{test_image.stem}.json"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"结果已保存到: {output_path}")
        
    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
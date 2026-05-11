#!/usr/bin/env python3
"""对比测试：PaddleOCR vs MiniMax VLM 价格提取效果"""

import json
import subprocess
import sys
import time
from pathlib import Path

# 测试图片列表
TEST_IMAGES = [
    Path("线下报价/长江报价.png"),
    Path("线下报价/徐钢报价.jpg"),
    Path("线下报价/贵航报价.png"),
    Path("线下报价/淮南宏泰报价.png"),
    Path("线下报价/金虹报价.png"),
    Path("线下报价/桂鑫报价.png"),
]

OUTPUT_DIR = Path("运行产物/ocr对比测试")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def run_paddleocr(image_path: Path) -> dict:
    """用PaddleOCR API提取价格"""
    output_path = OUTPUT_DIR / f"paddle_{image_path.stem}.json"
    cmd = [
        sys.executable, "-m", "ocr_price.cli",
        "--input", str(image_path),
        "--location", "蚌埠",
        "--provider", "api",
        "--output", str(output_path),
    ]
    start = time.time()
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    elapsed = time.time() - start

    if result.returncode != 0:
        return {
            "success": False,
            "error": result.stderr,
            "stdout": result.stdout,
            "elapsed": elapsed,
        }

    try:
        data = json.loads(output_path.read_text(encoding="utf-8"))
        return {
            "success": True,
            "elapsed": elapsed,
            "records": data.get("records", []),
            "company": data.get("company"),
            "quote_date": data.get("quote_date"),
            "record_count": data.get("meta", {}).get("record_count", 0),
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "elapsed": elapsed,
        }


def run_minimax(image_path: Path) -> dict:
    """用MiniMax VLM提取价格"""
    output_path = OUTPUT_DIR / f"minimax_{image_path.stem}.json"
    cmd = [
        sys.executable, "-m", "ocr_price.cli",
        "--input", str(image_path),
        "--location", "蚌埠",
        "--provider", "minimax",
        "--output", str(output_path),
    ]
    start = time.time()
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    elapsed = time.time() - start

    if result.returncode != 0:
        return {
            "success": False,
            "error": result.stderr,
            "stdout": result.stdout,
            "elapsed": elapsed,
        }

    try:
        data = json.loads(output_path.read_text(encoding="utf-8"))
        return {
            "success": True,
            "elapsed": elapsed,
            "records": data.get("records", []),
            "company": data.get("company"),
            "quote_date": data.get("quote_date"),
            "record_count": data.get("meta", {}).get("record_count", 0),
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "elapsed": elapsed,
        }


def extract_bengbu_price(records: list) -> dict:
    """从records中提取蚌埠的价格"""
    for r in records:
        loc = r.get("location", "")
        if "蚌埠" in str(loc):
            return {
                "rebar": r.get("rebar_price"),
                "coil": r.get("coil_price"),
                "rebar_raw": r.get("rebar_raw"),
                "coil_raw": r.get("coil_raw"),
            }
    return {}


def main():
    results = []

    for img in TEST_IMAGES:
        if not img.exists():
            print(f"跳过（文件不存在）: {img}")
            continue

        print(f"\n{'='*60}")
        print(f"测试图片: {img.name}")
        print(f"{'='*60}")

        # PaddleOCR
        print("[1/2] 运行 PaddleOCR API...")
        paddle_result = run_paddleocr(img)
        if paddle_result["success"]:
            print(f"  成功 | 耗时: {paddle_result['elapsed']:.1f}s | 记录数: {paddle_result['record_count']}")
            bb = extract_bengbu_price(paddle_result["records"])
            print(f"  蚌埠价格: 螺纹={bb.get('rebar')}, 盘螺={bb.get('coil')}")
        else:
            print(f"  失败 | 错误: {paddle_result.get('error', 'unknown')[:200]}")

        # MiniMax
        print("[2/2] 运行 MiniMax VLM...")
        minimax_result = run_minimax(img)
        if minimax_result["success"]:
            print(f"  成功 | 耗时: {minimax_result['elapsed']:.1f}s | 记录数: {minimax_result['record_count']}")
            bb = extract_bengbu_price(minimax_result["records"])
            print(f"  蚌埠价格: 螺纹={bb.get('rebar')}, 盘螺={bb.get('coil')}")
        else:
            print(f"  失败 | 错误: {minimax_result.get('error', 'unknown')[:200]}")

        results.append({
            "image": img.name,
            "paddleocr": paddle_result,
            "minimax": minimax_result,
        })

    # 保存完整对比结果
    summary_path = OUTPUT_DIR / "对比结果汇总.json"
    summary_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n\n完整对比结果已保存: {summary_path}")

    # 打印汇总表
    print("\n" + "="*80)
    print("对比汇总表")
    print("="*80)
    print(f"{'图片':<16} {'PaddleOCR':^30} {'MiniMax':^30}")
    print(f"{'':16} {'耗时':>8} {'螺纹':>8} {'盘螺':>8} {'耗时':>8} {'螺纹':>8} {'盘螺':>8}")
    print("-"*80)
    for r in results:
        img_name = r["image"][:14]
        p = r["paddleocr"]
        m = r["minimax"]

        if p["success"]:
            bb_p = extract_bengbu_price(p["records"])
            p_str = f"{p['elapsed']:>7.1f}s {str(bb_p.get('rebar')):>8} {str(bb_p.get('coil')):>8}"
        else:
            p_str = f"{'失败':>7} {'-':>8} {'-':>8}"

        if m["success"]:
            bb_m = extract_bengbu_price(m["records"])
            m_str = f"{m['elapsed']:>7.1f}s {str(bb_m.get('rebar')):>8} {str(bb_m.get('coil')):>8}"
        else:
            m_str = f"{'失败':>7} {'-':>8} {'-':>8}"

        print(f"{img_name:<16} {p_str} {m_str}")
    print("="*80)


if __name__ == "__main__":
    main()

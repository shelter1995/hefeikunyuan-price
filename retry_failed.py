from pathlib import Path
from ocr_price.minimax_vision import analyze_quote_image_to_ocr_format, MiniMaxVisionError
from ocr_price.env_loader import load_env_file
import json
import time

load_env_file('.env')

failed = ['淮南宏泰报价.png', '贵航报价.png', '金虹报价.png', '长江报价.png']

for name in failed:
    img = Path('线下报价') / name
    print(f"\n{'='*60}")
    print(f"重试: {name}")
    print(f"{'='*60}")
    for attempt in range(3):
        try:
            result = analyze_quote_image_to_ocr_format(
                image_path=img,
                target_cities=['合肥', '蚌埠'],
                save_raw_path=Path('运行产物') / f'minimax_raw_{img.stem}.json',
            )
            print(f"  [OK] 尝试{attempt+1}成功")
            company = result.get('company') or '未知'
            records = result.get('records', [])
            vision = result.get('_vision_result', {})
            inventory = vision.get('库存情况', [])
            print(f"       厂家: {company}")
            print(f"       记录数: {len(records)}, 库存项: {len(inventory)}")
            for rec in records:
                loc = rec.get('location', '')
                rebar = rec.get('rebar_price')
                coil = rec.get('coil_price')
                rebar_raw = rec.get('rebar_raw')
                if rebar_raw == '电议':
                    print(f"       {loc}: 螺纹=电议, 盘螺={coil}")
                else:
                    print(f"       {loc}: 螺纹={rebar}, 盘螺={coil}")
            # Save OCR JSON
            out_file = Path('运行产物') / f'ocr价格提取_{img.stem}.json'
            out_file.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
            break
        except Exception as e:
            print(f"  [ERROR] 尝试{attempt+1}失败: {str(e)[:100]}")
            if attempt < 2:
                time.sleep(5)

from ocr_price.minimax_vision import analyze_quote_image
from ocr_price.env_loader import load_env_file
import json

load_env_file('.env')

# 重试淮南宏泰
for i in range(3):
    print(f'=== 尝试 {i+1} ===')
    result = analyze_quote_image('线下报价/淮南宏泰报价.png', ['合肥', '蚌埠'])
    hf = result.get('合肥', {})
    bb = result.get('蚌埠', {})
    print(f"合肥: 螺纹={hf.get('螺纹')}, 盘螺={hf.get('盘螺')}, 电议={hf.get('螺纹为电议')}")
    print(f"蚌埠: 螺纹={bb.get('螺纹')}, 盘螺={bb.get('盘螺')}, 电议={bb.get('螺纹为电议')}")
    print(f"库存项数: {len(result.get('库存情况', []))}")
    print()

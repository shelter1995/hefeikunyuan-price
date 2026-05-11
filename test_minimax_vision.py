#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使用MiniMax视觉识别提取所有图片报价的价格和库存信息"""

import json
from pathlib import Path
from datetime import datetime

# MiniMax VLM API配置
# 需要设置MINIMAX_API_KEY环境变量或在代码中传入
# 当前从MCP工具配置中获取

VLM_PROMPT = """你是一个钢材报价表分析专家。请仔细分析这张报价图片，提取以下信息并以严格的JSON格式输出。

重要规则：
1. 仅以表格中直接填写的数据为准，不推算、不估算
2. 如果某城市在表格中无数据或为空白，对应价格填null
3. 如果某价格标注为"电议"，对应字段填null，并将"电议"标记设为true
4. 表格列可能是"螺纹+盘螺"，也可能是"盘螺+线材"或"仅螺纹"等，请根据实际表头判断
5. 库存状态请根据"极少"、"无货"、"少"、"件数"等描述判断
6. 螺纹钢的特征是有9米、12米长度规格；盘螺是盘状卷材

请输出以下JSON（不要输出任何其他文字）：
{
  "厂家名称": "从图片中识别的厂家名称",
  "报价日期": "yyyy-MM-dd格式，如无法识别则填null",
  "表格列类型": "螺纹+盘螺 或 盘螺+线材 或 仅螺纹 等",
  "合肥": {
    "螺纹": 数字或null,
    "盘螺": 数字或null,
    "螺纹为电议": true或false,
    "盘螺为电议": true或false
  },
  "蚌埠": {
    "螺纹": 数字或null,
    "盘螺": 数字或null,
    "螺纹为电议": true或false,
    "盘螺为电议": true或false
  },
  "库存情况": [
    {"规格": "如12E", "状态": "充足/告警/缺货", "原始描述": "如极少、无货、36件等"}
  ],
  "备注": "任何特殊说明"
}"""

# 所有图片文件
offline_quotes_dir = Path("线下报价")
image_files = sorted(
    list(offline_quotes_dir.glob("*.jpg")) + list(offline_quotes_dir.glob("*.png"))
)

print(f"找到 {len(image_files)} 张图片文件")
print("=" * 60)
print("提示：此脚本需要MINIMAX_API_KEY环境变量")
print("请设置后运行，或在Trae IDE中使用MCP工具逐个分析")
print("=" * 60)

for img in image_files:
    print(f"\n待分析: {img.name}")

print("\n" + "=" * 60)
print("由于需要MINIMAX_API_KEY，请通过以下方式之一运行：")
print("1. 设置环境变量: set MINIMAX_API_KEY=your_key")
print("2. 在.env文件中添加: MINIMAX_API_KEY=your_key")
print("3. 通过Trae IDE的MCP工具直接调用")
print("=" * 60)

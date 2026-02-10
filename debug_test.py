"""快速调试脚本"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib.ExcelHandler import ExcelHandler

# 测试指令解析
handler = ExcelHandler.from_config()

test_instructions = [
    "在A列添加1到36的数字",
    "在A列填充1到36",
    "把A列添加1到36的数字",
]

print("测试指令解析:")
for inst in test_instructions:
    result = handler.parse_instruction(inst)
    print(f"\n指令: {inst}")
    print(f"结果: {result}")

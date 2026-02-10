"""测试魔搭API连接"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib.ExcelHandler import ExcelHandler

print("正在测试API连接...")
print()

try:
    handler = ExcelHandler.from_config()
    print("[OK] 配置加载成功")
    print(f"    API: {handler.api_manager.base_url}")
    print(f"    模型: {handler.api_manager.model}")
    print()
    
    print("测试API调用...")
    success, result, error = handler.api_manager.process_cell(
        cell_content="怎么求和?",
        system_prompt="你是Excel助手,简短回答",
        user_instruction="回答问题",
        temperature=0.7,
        max_tokens=200
    )
    
    if success:
        print("[OK] API调用成功!")
        print()
        print("回答:")
        print(result)
    else:
        print(f"[FAIL] API调用失败: {error}")
        
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()

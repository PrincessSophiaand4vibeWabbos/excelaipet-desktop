"""
测试脚本 - 验证Excel教学桌宠功能
"""
import sys
import os

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """测试导入"""
    print("测试导入模块...")
    
    try:
        from lib.DashScopeAPIManager import DashScopeAPIManager
        print("  [OK] DashScopeAPIManager")
    except Exception as e:
        print(f"  [FAIL] DashScopeAPIManager: {e}")
        return False
    
    try:
        from lib.ExcelDataManager import ExcelDataManager
        print("  [OK] ExcelDataManager")
    except Exception as e:
        print(f"  [FAIL] ExcelDataManager: {e}")
        return False
    
    try:
        from lib.ExcelHandler import ExcelHandler
        print("  [OK] ExcelHandler")
    except Exception as e:
        print(f"  [FAIL] ExcelHandler: {e}")
        return False
    
    return True

def test_config():
    """测试配置加载"""
    print("\n测试配置加载...")
    
    try:
        from lib.ExcelHandler import ExcelHandler
        handler = ExcelHandler.from_config()
        print("  [OK] 配置加载成功")
        return handler
    except Exception as e:
        print(f"  [FAIL] 配置加载失败: {e}")
        return None

def test_api_connection(handler):
    """测试API连接"""
    print("\n测试API连接...")
    
    if handler is None:
        print("  [SKIP] handler未初始化")
        return False
    
    try:
        success, msg = handler.test_api_connection()
        if success:
            print(f"  [OK] {msg}")
        else:
            print(f"  [FAIL] {msg}")
        return success
    except Exception as e:
        print(f"  [FAIL] {e}")
        return False

def test_instruction_parsing(handler):
    """测试指令解析"""
    print("\n测试指令解析...")
    
    if handler is None:
        print("  [SKIP] handler未初始化")
        return
    
    test_cases = [
        "把A列翻译成中文",
        "将B列内容总结",
        "把第1列转成大写",
        "对C列进行分类",
    ]
    
    for instruction in test_cases:
        result = handler.parse_instruction(instruction)
        if result['valid']:
            print(f"  [OK] '{instruction}' -> 列:{result['target_column']}, 操作:{result['ai_prompt']}")
        else:
            print(f"  [FAIL] '{instruction}' -> {result.get('error', '解析失败')}")

def test_excel_operations():
    """测试Excel操作"""
    print("\n测试Excel操作...")
    
    try:
        import pandas as pd
        from lib.ExcelDataManager import ExcelDataManager
        
        # 创建测试数据
        test_data = {
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Message': ['Hello', 'Good morning', 'Thank you']
        }
        
        # 保存测试文件
        test_file = 'd:/pet/test_excel.xlsx'
        df = pd.DataFrame(test_data)
        df.to_excel(test_file, index=False)
        print(f"  [OK] 创建测试文件: {test_file}")
        
        # 测试加载
        dm = ExcelDataManager()
        success, error = dm.load_file(test_file)
        if success:
            print(f"  [OK] 加载文件成功: {dm.get_meta_info()}")
        else:
            print(f"  [FAIL] 加载文件失败: {error}")
            return False
        
        # 测试获取数据
        data = dm.get_column_data('Message')
        print(f"  [OK] 获取列数据: {len(data)}行")
        
        return True
        
    except Exception as e:
        print(f"  [FAIL] {e}")
        return False

if __name__ == "__main__":
    print("=" * 50)
    print("Excel教学桌宠 - 功能测试")
    print("=" * 50)
    
    # 1. 测试导入
    if not test_imports():
        print("\n导入测试失败，请安装依赖: pip install openai pandas openpyxl")
        sys.exit(1)
    
    # 2. 测试配置
    handler = test_config()
    
    # 3. 测试指令解析
    test_instruction_parsing(handler)
    
    # 4. 测试Excel操作
    test_excel_operations()
    
    # 5. 测试API连接(可选)
    print("\n是否测试API连接? (需要网络，可能产生费用)")
    # test_api_connection(handler)
    
    print("\n" + "=" * 50)
    print("测试完成!")
    print("=" * 50)
    print("\n运行桌宠: python main.py")
    print("右键点击桌宠可打开Excel操作菜单")

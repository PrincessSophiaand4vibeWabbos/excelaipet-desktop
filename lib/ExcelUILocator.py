"""
Excel界面定位器 - 定位Excel界面元素并提供引导
"""

import win32gui
import win32api
import win32con
from typing import Dict, Tuple, Optional, List
from dataclasses import dataclass


@dataclass
class UIElement:
    """UI元素定义"""
    name: str           # 元素名称
    keywords: List[str] # 触发关键词
    offset_x: int       # 相对Excel窗口左上角的X偏移
    offset_y: int       # 相对Excel窗口左上角的Y偏移
    description: str    # 操作说明


# Excel 2016/2019/365 常用界面元素位置（相对于窗口左上角）
# 这些是大致位置，可能需要根据实际Excel版本微调
EXCEL_UI_ELEMENTS = {
    # 文件标签
    "file_tab": UIElement("文件", ["文件", "file", "打开", "保存", "新建"], 30, 50, "点击[文件]标签"),
    
    # 开始标签
    "home_tab": UIElement("开始", ["开始", "home", "格式", "字体", "对齐"], 80, 50, "点击[开始]标签"),
    
    # 插入标签
    "insert_tab": UIElement("插入", ["插入", "insert", "图表", "表格", "图片"], 130, 50, "点击[插入]标签"),
    
    # 页面布局标签
    "layout_tab": UIElement("页面布局", ["页面", "布局", "layout", "边距", "方向"], 200, 50, "点击[页面布局]标签"),
    
    # 公式标签
    "formula_tab": UIElement("公式", ["公式", "formula", "函数", "计算"], 280, 50, "点击[公式]标签"),
    
    # 数据标签
    "data_tab": UIElement("数据", ["数据", "data", "筛选", "排序", "分列"], 340, 50, "点击[数据]标签"),
    
    # 审阅标签
    "review_tab": UIElement("审阅", ["审阅", "review", "批注", "保护"], 400, 50, "点击[审阅]标签"),
    
    # 视图标签
    "view_tab": UIElement("视图", ["视图", "view", "冻结", "窗格"], 460, 50, "点击[视图]标签"),
    
    # 开始标签下的功能按钮
    "bold": UIElement("加粗", ["加粗", "粗体", "bold", "B"], 50, 95, "点击[B]加粗按钮"),
    "italic": UIElement("斜体", ["斜体", "italic", "I"], 75, 95, "点击[I]斜体按钮"),
    "underline": UIElement("下划线", ["下划线", "underline", "U"], 100, 95, "点击[U]下划线按钮"),
    "merge_center": UIElement("合并居中", ["合并", "merge", "居中"], 380, 95, "点击[合并后居中]按钮"),
    "auto_sum": UIElement("自动求和", ["求和", "sum", "合计", "相加"], 720, 95, "点击[自动求和]按钮(Σ)"),
    
    # 数据标签下的功能
    "filter": UIElement("筛选", ["筛选", "filter", "过滤"], 200, 95, "点击[筛选]按钮"),
    "sort": UIElement("排序", ["排序", "sort", "升序", "降序"], 150, 95, "点击[排序]按钮"),
    
    # 视图标签下的功能
    "freeze_panes": UIElement("冻结窗格", ["冻结", "freeze", "固定"], 180, 95, "点击[冻结窗格]按钮"),
    
    # 公式标签下的功能
    "insert_function": UIElement("插入函数", ["函数", "fx", "function", "vlookup", "if"], 50, 95, "点击[插入函数]按钮(fx)"),
    
    # 名称框和编辑栏
    "name_box": UIElement("名称框", ["名称框", "单元格地址", "跳转"], 50, 130, "点击名称框可跳转到指定单元格"),
    "formula_bar": UIElement("编辑栏", ["编辑栏", "公式栏", "输入"], 200, 130, "在编辑栏输入内容或公式"),
    
    # 工作表区域
    "cell_area": UIElement("单元格区域", ["单元格", "表格", "数据区"], 100, 200, "这是单元格编辑区域"),
    "sheet_tab": UIElement("工作表标签", ["sheet", "工作表", "标签", "切换"], 100, -30, "点击底部工作表标签切换"),
}


class ExcelUILocator:
    """Excel界面定位器"""
    
    def __init__(self):
        self.excel_hwnd = None
        self.excel_rect = None
    
    def find_excel_window(self) -> bool:
        """查找Excel窗口"""
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                title_lower = title.lower()
                if (
                    'excel' in title_lower
                    or '.xlsx' in title_lower
                    or '.xls' in title_lower
                    or '工作簿' in title
                    or '表格' in title
                    or 'wps' in title_lower
                ):
                    results.append(hwnd)
            return True
        
        results = []
        win32gui.EnumWindows(enum_callback, results)
        
        if results:
            self.excel_hwnd = results[0]
            self.excel_rect = win32gui.GetWindowRect(self.excel_hwnd)
            return True
        return False
    
    def get_excel_position(self) -> Optional[Tuple[int, int, int, int]]:
        """获取Excel窗口位置"""
        if self.find_excel_window():
            return self.excel_rect
        return None
    
    def find_element_by_keywords(self, text: str) -> Optional[UIElement]:
        """根据文本关键词查找UI元素"""
        text_lower = text.lower()
        
        # 按优先级匹配
        best_match = None
        best_score = 0
        
        for key, element in EXCEL_UI_ELEMENTS.items():
            score = 0
            for keyword in element.keywords:
                if keyword.lower() in text_lower:
                    score += 1
            
            if score > best_score:
                best_score = score
                best_match = element
        
        return best_match if best_score > 0 else None
    
    def get_element_screen_position(self, element: UIElement) -> Optional[Tuple[int, int]]:
        """获取元素在屏幕上的绝对位置"""
        if not self.find_excel_window():
            return None
        
        # 计算绝对坐标
        x = self.excel_rect[0] + element.offset_x
        y = self.excel_rect[1] + element.offset_y
        
        # 处理负偏移（如工作表标签在底部）
        if element.offset_y < 0:
            y = self.excel_rect[3] + element.offset_y
        
        return (x, y)
    
    def parse_teaching_response(self, response: str) -> List[Dict]:
        """解析AI教学回答，提取操作步骤和对应UI位置"""
        steps = []
        
        # 按行分割，查找编号步骤
        lines = response.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # 查找对应的UI元素
            element = self.find_element_by_keywords(line)
            
            if element:
                pos = self.get_element_screen_position(element)
                steps.append({
                    'text': line,
                    'element': element,
                    'position': pos
                })
            else:
                steps.append({
                    'text': line,
                    'element': None,
                    'position': None
                })
        
        return steps

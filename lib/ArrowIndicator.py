"""
箭头指示器 - 在屏幕上显示指向目标位置的箭头
"""

import tkinter as tk
from tkinter import Canvas
from typing import Tuple, Optional
import math


class ArrowIndicator:
    """箭头指示器窗口"""
    
    def __init__(self, master=None):
        """初始化箭头指示器"""
        self.window = None
        self.canvas = None
        self.master = master
        self.visible = False
        self.arrow_size = 60
        
    def create_window(self):
        """创建箭头窗口"""
        if self.window is not None:
            return
        
        self.window = tk.Toplevel(self.master)
        self.window.overrideredirect(True)  # 无边框
        self.window.wm_attributes("-topmost", True)  # 始终置顶
        self.window.wm_attributes("-transparentcolor", "white")  # 白色透明
        self.window.geometry(f"{self.arrow_size}x{self.arrow_size}+0+0")
        
        self.canvas = Canvas(
            self.window, 
            width=self.arrow_size, 
            height=self.arrow_size,
            bg="white",
            highlightthickness=0
        )
        self.canvas.pack()
        
        self.window.withdraw()  # 初始隐藏
    
    def draw_arrow(self, direction: str = "down"):
        """绘制箭头"""
        if self.canvas is None:
            self.create_window()
        
        self.canvas.delete("all")
        
        size = self.arrow_size
        center = size // 2
        
        # 箭头颜色
        color = "#FF4444"  # 红色
        outline = "#CC0000"
        
        if direction == "down":
            # 向下的箭头
            points = [
                center, size - 5,      # 箭头尖端
                center - 15, size - 25,  # 左翼
                center - 5, size - 25,   # 左内
                center - 5, 5,           # 左上
                center + 5, 5,           # 右上
                center + 5, size - 25,   # 右内
                center + 15, size - 25,  # 右翼
            ]
        elif direction == "up":
            points = [
                center, 5,
                center - 15, 25,
                center - 5, 25,
                center - 5, size - 5,
                center + 5, size - 5,
                center + 5, 25,
                center + 15, 25,
            ]
        elif direction == "left":
            points = [
                5, center,
                25, center - 15,
                25, center - 5,
                size - 5, center - 5,
                size - 5, center + 5,
                25, center + 5,
                25, center + 15,
            ]
        else:  # right
            points = [
                size - 5, center,
                size - 25, center - 15,
                size - 25, center - 5,
                5, center - 5,
                5, center + 5,
                size - 25, center + 5,
                size - 25, center + 15,
            ]
        
        self.canvas.create_polygon(
            points,
            fill=color,
            outline=outline,
            width=2
        )
    
    def show_at(self, x: int, y: int, direction: str = "down"):
        """在指定位置显示箭头"""
        if self.window is None:
            self.create_window()
        
        self.draw_arrow(direction)
        
        # 根据方向调整位置，使箭头指向目标点
        offset = self.arrow_size
        if direction == "down":
            x -= self.arrow_size // 2
            y -= self.arrow_size
        elif direction == "up":
            x -= self.arrow_size // 2
        elif direction == "left":
            y -= self.arrow_size // 2
        else:  # right
            x -= self.arrow_size
            y -= self.arrow_size // 2
        
        self.window.geometry(f"{self.arrow_size}x{self.arrow_size}+{x}+{y}")
        self.window.deiconify()
        self.visible = True
    
    def hide(self):
        """隐藏箭头"""
        if self.window:
            self.window.withdraw()
        self.visible = False
    
    def destroy(self):
        """销毁窗口"""
        if self.window:
            self.window.destroy()
            self.window = None
            self.canvas = None


class TeachingGuide:
    """教学引导器 - 管理桌宠移动和箭头指示"""
    
    def __init__(self, sprite_controller, master):
        self.sprite = sprite_controller
        self.arrow = ArrowIndicator(master)
        self.current_step = 0
        self.steps = []
        self.guiding = False
    
    def start_guide(self, steps: list):
        """开始引导流程"""
        self.steps = [s for s in steps if s.get('position')]  # 只保留有位置的步骤
        self.current_step = 0
        self.guiding = True
        
        if self.steps:
            self.show_current_step()
    
    def show_current_step(self):
        """显示当前步骤"""
        if self.current_step >= len(self.steps):
            self.stop_guide()
            return
        
        step = self.steps[self.current_step]
        pos = step.get('position')
        
        if pos:
            x, y = pos
            
            # 移动桌宠到目标位置附近
            self.sprite.pos_x = x - 50
            self.sprite.pos_y = y + 30
            self.sprite.UpdateRootWindow()
            
            # 显示箭头
            self.arrow.show_at(x, y, "down")
            
            # 显示说明文字
            element = step.get('element')
            if element:
                self.sprite._show_bubble(
                    f"步骤 {self.current_step + 1}:\n{element.description}",
                    duration=0
                )
    
    def next_step(self):
        """下一步"""
        if not self.guiding:
            return
        
        self.current_step += 1
        if self.current_step < len(self.steps):
            self.show_current_step()
        else:
            self.stop_guide()
    
    def stop_guide(self):
        """停止引导"""
        self.guiding = False
        self.arrow.hide()
        self.current_step = 0
        self.steps = []
        self.sprite._show_bubble("教学完成!", duration=3000)

from os import listdir
from random import choice
from time import time
from pathlib import Path
from queue import Empty, Queue
from threading import Thread, current_thread, main_thread

from PIL import Image, ImageTk
from tkinter import Button, Entry, Frame, Label, Menu, StringVar, Toplevel, filedialog, simpledialog

from lib.WindowHandler import Handler
from lib.CommentGenerator import Commenter
from lib.ExcelHandler import ExcelHandler
from lib.ExcelUILocator import ExcelUILocator
from lib.ArrowIndicator import ArrowIndicator, TeachingGuide


class SpriteController:
    def __init__(self, root, chat_win_root) -> None:
        self.animation_frames: dict[str:list[str]] = self.LoadAnimations()
        
        self.commenter = Commenter()

        self.direction: str = "right"
        self.current_animation = self.animation_frames[f"idle_{self.direction}"]

        self.root = root
        self.chat_window_root = chat_win_root

        self.label = Label(self.root)
        
        self.chat_label = Label(self.chat_window_root)
        self.chat_response = None
        self.chat = False
        
        self.tts_begun = False

        self.moving = False
        
        self.fall    = False
        self.jumping = False

        self.idle_delay     = 0
        self.max_idle_delay = 2000

        self.move_distance     = 0
        self.max_move_distance = 200

        self.move_delay     = 0
        self.max_move_delay = 50

        self.frame_index = 0
        self.max_frame_index = len(self.current_animation) - 1

        self.delay_count = 0

        self.max_delay   = 200
        self.max_delay_0 = 700

        self.fall_frame_delay     = 0
        self.max_fall_frame_delay = 50

        self.pos_x = 20
        self.pos_y = 730

        self.ground_x = 20
        self.ground_y = 730

        self.target_x_max   = 0
        self.target_x       = 0
        self.target_y       = 0

        self.x_border       = 10
        self.x_border_right = 1500
        self.y_border       = 730

        self.jump_init_counter = time()
        self.chat_init_counter = time()

        self.jumped_to_window = False
        self.ui_task_queue: Queue = Queue()

        # Excel功能初始化
        self.excel_handler = None
        self.excel_processing = False
        self.comment_processing = False
        self._init_excel_handler()
        self._setup_context_menu()
        
        # 教学引导初始化
        self.ui_locator = ExcelUILocator()
        self.teaching_guide = TeachingGuide(self, self.root)

        # 左键交互悬浮窗
        self.panel_window = None
        self.panel_visible = False
        self.panel_drag_offset_x = 0
        self.panel_drag_offset_y = 0
        self.selected_excel_file = ""
        self.panel_question_var = StringVar()
        self.panel_instruction_var = StringVar()
        self.panel_comment_var = StringVar()
        self.panel_file_var = StringVar(value="未选择文件")
        self._create_floating_panel()
        
        # 绑定右键事件
        self.label.bind("<Button-3>", self._show_context_menu)
        # 绑定左键点击（打开悬浮窗/引导下一步）
        self.label.bind("<Button-1>", self._on_click)

    def _init_excel_handler(self):
        """初始化Excel处理器"""
        try:
            self.excel_handler = ExcelHandler.from_config()
        except Exception as e:
            print(f"Excel处理器初始化失败: {e}")
            self.excel_handler = None

    def _run_on_ui_thread(self, callback):
        """将UI操作切换到主线程执行"""
        if current_thread() is main_thread():
            callback()
            return
        self.ui_task_queue.put(callback)

    def _flush_ui_tasks(self, max_tasks: int = 30):
        """在动画主循环中消费待执行的UI任务"""
        for _ in range(max_tasks):
            try:
                task = self.ui_task_queue.get_nowait()
            except Empty:
                break
            try:
                task()
            except Exception as e:
                print(f"UI任务执行失败: {e}")

    def _setup_context_menu(self):
        """创建右键菜单"""
        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(
            label="Excel教学问答",
            command=self._ask_excel_question
        )
        self.context_menu.add_command(
            label="Excel操作",
            command=self._open_excel_dialog
        )
        self.context_menu.add_separator()
        self.context_menu.add_command(
            label="截图评论",
            command=self._ask_screenshot_comment
        )
        self.context_menu.add_separator()
        self.context_menu.add_command(
            label="关于",
            command=lambda: self._show_bubble("Excel教学桌宠 v1.0\n右键点击可使用功能")
        )

    def _create_floating_panel(self):
        """创建左键交互悬浮窗"""
        self.panel_window = Toplevel(self.root)
        self.panel_window.withdraw()
        self.panel_window.overrideredirect(True)
        self.panel_window.wm_attributes("-topmost", True)
        self.panel_window.configure(background="#1f2933")
        self.panel_window.geometry(f"390x300+{self.pos_x + 120}+{max(self.pos_y - 120, 80)}")

        header = Frame(self.panel_window, background="#334155", height=34)
        header.pack(fill="x")

        title = Label(
            header,
            text="Excel桌宠助手",
            foreground="#e2e8f0",
            background="#334155",
            font=("Microsoft YaHei", 10, "bold")
        )
        title.pack(side="left", padx=10, pady=6)

        close_btn = Button(
            header,
            text="×",
            command=self._hide_floating_panel,
            foreground="#e2e8f0",
            background="#334155",
            activebackground="#475569",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            width=2
        )
        close_btn.pack(side="right", padx=8, pady=5)

        for widget in (header, title):
            widget.bind("<ButtonPress-1>", self._start_panel_drag)
            widget.bind("<B1-Motion>", self._drag_panel)
            widget.bind("<ButtonRelease-1>", self._end_panel_drag)

        body = Frame(self.panel_window, background="#1f2933")
        body.pack(fill="both", expand=True, padx=10, pady=8)

        file_row = Frame(body, background="#1f2933")
        file_row.pack(fill="x", pady=(0, 6))
        Button(
            file_row,
            text="选择Excel文件",
            command=self._panel_select_file,
            font=("Microsoft YaHei", 9)
        ).pack(side="left")
        Label(
            file_row,
            textvariable=self.panel_file_var,
            foreground="#cbd5e1",
            background="#1f2933",
            anchor="w",
            font=("Microsoft YaHei", 9),
            width=26
        ).pack(side="left", padx=8)

        Label(
            body,
            text="教学问题",
            foreground="#cbd5e1",
            background="#1f2933",
            anchor="w",
            font=("Microsoft YaHei", 9)
        ).pack(fill="x")
        q_row = Frame(body, background="#1f2933")
        q_row.pack(fill="x", pady=(2, 8))
        Entry(q_row, textvariable=self.panel_question_var, font=("Microsoft YaHei", 9)).pack(
            side="left", fill="x", expand=True
        )
        Button(
            q_row, text="问答", command=self._panel_submit_question, font=("Microsoft YaHei", 9), width=8
        ).pack(side="left", padx=(8, 0))

        Label(
            body,
            text="操作指令",
            foreground="#cbd5e1",
            background="#1f2933",
            anchor="w",
            font=("Microsoft YaHei", 9)
        ).pack(fill="x")
        op_row = Frame(body, background="#1f2933")
        op_row.pack(fill="x", pady=(2, 8))
        Entry(op_row, textvariable=self.panel_instruction_var, font=("Microsoft YaHei", 9)).pack(
            side="left", fill="x", expand=True
        )
        Button(
            op_row, text="执行", command=self._panel_submit_operation, font=("Microsoft YaHei", 9), width=8
        ).pack(side="left", padx=(8, 0))

        action_row = Frame(body, background="#1f2933")
        action_row.pack(fill="x")
        Label(
            body,
            text="截图评论指令(可选)",
            foreground="#cbd5e1",
            background="#1f2933",
            anchor="w",
            font=("Microsoft YaHei", 9)
        ).pack(fill="x", pady=(6, 0))
        comment_row = Frame(body, background="#1f2933")
        comment_row.pack(fill="x", pady=(2, 6))
        Entry(comment_row, textvariable=self.panel_comment_var, font=("Microsoft YaHei", 9)).pack(
            side="left", fill="x", expand=True
        )
        Button(
            comment_row, text="截图评论", command=self._panel_submit_screenshot_comment, font=("Microsoft YaHei", 9), width=8
        ).pack(side="left")
        Button(
            action_row, text="开始引导", command=self._start_visual_guide, font=("Microsoft YaHei", 9)
        ).pack(side="left", padx=8)

        self.panel_status_label = Label(
            body,
            text="左键点桌宠可显示/隐藏此窗口",
            foreground="#94a3b8",
            background="#1f2933",
            anchor="w",
            justify="left",
            wraplength=350,
            font=("Microsoft YaHei", 9)
        )
        self.panel_status_label.pack(fill="x", pady=(10, 0))

    def _set_panel_status(self, text: str):
        if hasattr(self, "panel_status_label") and self.panel_status_label:
            self.panel_status_label.configure(text=text)

    def _toggle_floating_panel(self):
        if not self.panel_window:
            return
        if self.panel_visible:
            self._hide_floating_panel()
        else:
            self._show_floating_panel()

    def _show_floating_panel(self):
        if not self.panel_window:
            return
        self.panel_window.deiconify()
        self.panel_visible = True
        self._set_panel_status("可输入问题或操作指令")

    def _hide_floating_panel(self):
        if not self.panel_window:
            return
        self.panel_window.withdraw()
        self.panel_visible = False

    def _start_panel_drag(self, event):
        self.panel_drag_offset_x = event.x_root - self.panel_window.winfo_x()
        self.panel_drag_offset_y = event.y_root - self.panel_window.winfo_y()

    def _drag_panel(self, event):
        x = event.x_root - self.panel_drag_offset_x
        y = event.y_root - self.panel_drag_offset_y
        self.panel_window.geometry(f"+{x}+{y}")

    def _end_panel_drag(self, event):
        _ = event

    def _panel_select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls *.xlsm"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if not file_path:
            return
        self.selected_excel_file = file_path
        self.panel_file_var.set(Path(file_path).name)
        self._set_panel_status(f"已选择文件: {Path(file_path).name}")

    def _panel_submit_question(self):
        question = self.panel_question_var.get().strip()
        if not question:
            self._set_panel_status("请先输入教学问题")
            return
        self._ask_excel_question_with_text(question)

    def _panel_submit_operation(self):
        instruction = self.panel_instruction_var.get().strip()
        if not instruction:
            self._set_panel_status("请先输入操作指令")
            return
        if not self.selected_excel_file:
            self._set_panel_status("请先选择Excel文件")
            self._panel_select_file()
            if not self.selected_excel_file:
                return
        self._start_excel_operation(self.selected_excel_file, instruction)

    def _panel_submit_screenshot_comment(self):
        comment_instruction = self.panel_comment_var.get().strip()
        self.StartScreenshotComment(comment_instruction)

    def _ask_screenshot_comment(self):
        user_prompt = simpledialog.askstring(
            "截图评论",
            "请输入截图评论指令（可选）:\n例如: 请重点点评Excel中的公式错误",
            parent=self.root
        )
        if user_prompt is None:
            return
        self.StartScreenshotComment(user_prompt.strip())

    def _ask_excel_question(self):
        """Excel教学问答"""
        if self.excel_handler is None:
            self._show_bubble("功能未初始化\n请检查配置文件")
            self._set_panel_status("Excel功能未初始化")
            return
        
        if self.excel_processing:
            self._show_bubble("正在处理中...\n请稍候")
            self._set_panel_status("正在处理中，请稍候")
            return
        
        # 获取用户问题
        question = simpledialog.askstring(
            "Excel教学",
            "请输入你的Excel问题:\n(例如: 怎么求和? 如何筛选数据?)",
            parent=self.root
        )
        
        if not question:
            return
        self._ask_excel_question_with_text(question)

    def _ask_excel_question_with_text(self, question: str):
        self.excel_processing = True
        self._show_bubble("思考中...", duration=0)
        self._set_panel_status("教学问答处理中...")

        # 后台调用AI回答
        thread = Thread(
            target=self._get_excel_teaching,
            args=(question,),
            daemon=True
        )
        thread.start()

    def _get_excel_teaching(self, question):
        """调用AI获取Excel教学回答"""
        try:
            system_prompt = """你是一个Excel教学助手。用户会问你Excel相关的问题，请用简洁清晰的方式回答。

回答要求：
1. 用中文回答
2. 给出具体的操作步骤（点哪里、按什么键）
3. 如果涉及公式，给出公式示例
4. 回答控制在150字以内
5. 使用编号列出步骤
6. 每个步骤单独一行，用数字编号如 1. 2. 3."""

            success, result, error = self.excel_handler.api_manager.process_cell(
                cell_content=question,
                system_prompt=system_prompt,
                user_instruction="请回答这个Excel问题",
                temperature=0.7,
                max_tokens=500
            )
            
            if success:
                # 解析回答，查找可以引导的步骤
                steps = self.ui_locator.parse_teaching_response(result)
                guide_steps = [s for s in steps if s.get('position')]
                
                if guide_steps:
                    # 有可引导的步骤，询问是否开始引导
                    self._run_on_ui_thread(lambda: self._offer_guide(result, steps))
                else:
                    # 没有可引导的步骤，直接显示答案
                    self._run_on_ui_thread(lambda: self._show_bubble(result, duration=15000))
                self._run_on_ui_thread(lambda: self._set_panel_status("问答完成，可点击开始引导"))
            else:
                self._run_on_ui_thread(lambda: self._show_bubble(f"回答失败: {error}"))
                self._run_on_ui_thread(lambda: self._set_panel_status("问答失败，请稍后重试"))
                 
        except Exception as e:
            err_msg = str(e)
            self._run_on_ui_thread(lambda msg=err_msg: self._show_bubble(f"错误: {msg}"))
            self._run_on_ui_thread(lambda msg=err_msg: self._set_panel_status(f"问答异常: {msg}"))
        finally:
            self.excel_processing = False

    def _offer_guide(self, answer, steps):
        """提供引导选项"""
        guide_steps = [s for s in steps if s.get('position')]
        
        # 先显示答案
        self._show_bubble(f"{answer}\n\n---\n点击桌宠开始演示引导\n({len(guide_steps)}个步骤)", duration=0)
        
        # 保存步骤，准备引导
        self._pending_guide_steps = steps

    def _start_visual_guide(self):
        """开始可视化引导"""
        if hasattr(self, '_pending_guide_steps') and self._pending_guide_steps:
            self.teaching_guide.start_guide(self._pending_guide_steps)
            self._pending_guide_steps = None
            self._set_panel_status("已开始引导，点击桌宠切换下一步")
        else:
            self._set_panel_status("暂无可引导步骤，请先进行教学问答")

    def _show_context_menu(self, event):
        """显示右键菜单"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _on_click(self, event):
        """左键点击处理"""
        if self.teaching_guide.guiding:
            # 正在引导中，点击进入下一步
            self.teaching_guide.next_step()
            return
        self._toggle_floating_panel()

    def _open_excel_dialog(self):
        """打开Excel文件并执行操作"""
        if self.excel_handler is None:
            self._show_bubble("Excel功能未初始化\n请检查配置文件")
            self._set_panel_status("Excel功能未初始化")
            return
        
        if self.excel_processing:
            self._show_bubble("正在处理中...\n请稍候")
            self._set_panel_status("正在处理中，请稍候")
            return
        
        # 1. 选择文件
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls *.xlsm"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        file_name = Path(file_path).name
        self.selected_excel_file = file_path
        self.panel_file_var.set(file_name)
        self._show_bubble(f"已选择:\n{file_name}")
        self._set_panel_status(f"已选择文件: {file_name}")
        
        # 2. 获取用户指令
        self.root.after(500, lambda: self._get_instruction(file_path, file_name))

    def _get_instruction(self, file_path, file_name):
        """获取用户操作指令"""
        instruction = simpledialog.askstring(
            "Excel操作",
            f"文件: {file_name}\n\n请描述要执行的操作:\n(例如: 把A列翻译成中文)",
            parent=self.root
        )
        
        if not instruction:
            self._show_bubble("操作已取消")
            return

        self._start_excel_operation(file_path, instruction)

    def _start_excel_operation(self, file_path: str, instruction: str):
        # 3. 在后台线程执行操作
        self.excel_processing = True
        self._show_bubble("正在处理中...\n请稍候")
        self._set_panel_status("Excel操作处理中...")
        
        thread = Thread(
            target=self._execute_excel_operation,
            args=(file_path, instruction),
            daemon=True
        )
        thread.start()

    def _execute_excel_operation(self, file_path, instruction):
        """执行Excel操作(在后台线程)"""
        try:
            def progress_update(msg):
                self._run_on_ui_thread(lambda: self._show_bubble(msg))
                self._run_on_ui_thread(lambda m=msg: self._set_panel_status(m))
            
            success, summary = self.excel_handler.execute_operation(
                file_path,
                instruction,
                progress_callback=progress_update
            )
            
            if success:
                result_msg = f"操作完成!\n\n{summary}"
            else:
                result_msg = f"操作失败\n\n{summary}"
            
            self._run_on_ui_thread(lambda: self._show_bubble(result_msg, duration=8000))
            self._run_on_ui_thread(lambda s=summary, ok=success: self._set_panel_status(
                ("操作完成: " if ok else "操作失败: ") + (s.splitlines()[0] if s else "无摘要")
            ))
            
        except Exception as e:
            err_msg = str(e)
            self._run_on_ui_thread(lambda msg=err_msg: self._show_bubble(f"错误: {msg}"))
            self._run_on_ui_thread(lambda msg=err_msg: self._set_panel_status(f"操作异常: {msg}"))
        finally:
            self.excel_processing = False

    def _show_bubble(self, text: str, duration: int = 5000):
        """显示气泡文字"""
        self.chat = True
        self.chat_response = text
        self.UpdateChatWindow()
        
        # 定时隐藏气泡
        if duration > 0:
            self.root.after(duration, self._hide_bubble)

    def _hide_bubble(self):
        """隐藏气泡"""
        if not self.excel_processing:  # 处理中不隐藏
            self.chat = False
            self.chat_window_root.wm_attributes("-alpha", 0)

    def GetScreenshotComment(self):
        # 保留旧入口，默认无额外指令
        self.StartScreenshotComment("")

    def StartScreenshotComment(self, user_instruction: str = ""):
        if self.comment_processing:
            self._set_panel_status("截图评论正在进行中，请稍候")
            return
        if self.excel_processing:
            self._set_panel_status("Excel处理中，稍后再截图评论")
            self._show_bubble("Excel处理中，稍后再截图评论")
            return

        self.moving = False
        self.jumping = False

        self.SetAnimation("sitting")

        last_x = self.pos_x
        last_y = self.pos_y

        self.pos_x = -50
        self.pos_y = -50

        self.UpdateRootWindow()
        self.tts_begun = True
        self.chat = True
        self._show_bubble("截图分析中...", duration=0)
        self.comment_processing = True
        if user_instruction:
            self._set_panel_status(f"截图评论处理中: {user_instruction}")
        else:
            self._set_panel_status("截图评论处理中...")

        thread = Thread(
            target=self._generate_screenshot_comment,
            args=(last_x, last_y, user_instruction),
            daemon=True
        )
        thread.start()

    def _generate_screenshot_comment(self, last_x: int, last_y: int, user_instruction: str):
        try:
            comment_text = self.commenter.GenerateComment(user_instruction=user_instruction)
        except Exception as e:
            comment_text = f"截图评论失败: {e}"

        def _finish_ui():
            self.chat_response = comment_text
            self.pos_x = last_x
            self.pos_y = last_y
            self.tts_begun = False
            self.comment_processing = False
            self._show_bubble(comment_text, duration=12000)
            self.commenter.ThreadedSpeaker()
            self._set_panel_status("截图评论完成")

        self._run_on_ui_thread(_finish_ui)

    def UpdateChatWindowAlpha(self):
        self.chat_window_root.wm_attributes("-alpha", 1 if self.chat else 0)

    def UpdateChatWindow(self):
        if not self.chat:
            self.chat_window_root.wm_attributes("-alpha", 0)
            return
        
        self.chat_window_root.geometry(f"+{self.pos_x+110}+{self.pos_y-20}")

        self.chat_label.configure(text=self.chat_response)
        
        self.chat_label.configure(wraplength=350)
        self.chat_label.configure(anchor="w", justify="left")
        
        # 优化样式：浅黄色背景，深色文字
        self.chat_label.configure(
            background="#FFFACD",
            foreground="#333333",
            font=("Microsoft YaHei", 10),
            padx=12,
            pady=8
        )

        self.chat_label.pack()
        self.chat_window_root.wm_attributes("-alpha", 0.95)
        
        self.chat_window_root.update_idletasks()
        self.chat_window_root.update()

    def UpdateRootWindow(self):
        wind = "100x100"
        self.root.image = self.current_animation[self.frame_index]
        self.root.geometry(f"{wind}+{self.pos_x}+{self.pos_y}")
        self.label.configure(image=self.root.image, background="black")
        self.label.pack()
        self.label.config(cursor="none")

        self.root.update_idletasks()
        self.root.update()

    def SetAnimation(self, name):
        self.current_animation = self.animation_frames[name]
        self.max_frame_index = len(self.current_animation) - 1
        self.frame_index = 0

    def MakeFall(self):
        if self.pos_y != self.ground_y:
            if self.fall_frame_delay == self.max_fall_frame_delay:
                self.pos_y+=1
                self.fall_frame_delay = 0
                self.UpdateRootWindow()
            else:
                self.fall_frame_delay+=1
        else:
            self.fall = False
            self.jumped_to_window = False

    def CheckFallWindow(self):
        if self.fall:
            return
        if not self.jumped_to_window:
            return
        
        new_fg_window_location = Handler.GetForegroundWindowPosition()

        if new_fg_window_location:
            if self.target_x != new_fg_window_location[0] and self.target_y != new_fg_window_location[1]:
                
                self.fall = True
                
                self.moving = False
                self.chat = False
                self.SetAnimation("fall")

    def SetTargetWindow(self):
        target_window_location = Handler.GetForegroundWindowPosition()
        
        if target_window_location:
            self.target_x     = target_window_location[0]
            self.target_y     = target_window_location[1]
            self.target_x_max = target_window_location[2]
            return True
        else:
            return False

    def MakeJump(self):
        if self.jumping:
            if self.pos_x != self.target_x:
                self.pos_x = self.pos_x+1 if self.pos_x < self.target_x else self.pos_x-1
            if self.pos_y != self.target_y:
                self.pos_y = self.pos_y+1 if self.pos_y < self.target_y else self.pos_y-1
            else:
                self.pos_y-=70
                self.jumping = False
                self.jumped_to_window = True
                self.UpdateMovementBorder()
    
    def CheckJump(self) -> None:
        current = int(time() - self.jump_init_counter)
        if current >= 20:
            if self.SetTargetWindow():

                if self.pos_y != self.target_y-70:
                
                    self.jumping = True
                    self.moving = False
                    if self.pos_x > self.target_x:
                        self.SetAnimation("jump_left")
                    else:
                        self.SetAnimation("jump_right")

            self.jump_init_counter = time()
    
    def UpdateMovementBorder(self):
        self.x_border = self.target_x
        self.x_border_right = self.target_x_max

    def InitChat(self):
        
        current = int(time() - self.chat_init_counter)
        
        if current >= 60:
            self.chat_init_counter = time()
            return True
        
        return False
        
    
    def SetIdleAnim(self, direction=None):
        if direction:
            self.direction = direction
        
        self.moving = False
        self.move_distance = 0

        anim = choice((f"idle_{self.direction}", "sitting", "3_idle", "4_idle"))
        
        self.SetAnimation(anim)

    def HandleAnimation(self):
        self._flush_ui_tasks()

        if self.fall:
            self.MakeFall()
            return
        
        if not self.fall:
            self.CheckFallWindow()
        
        if not self.jumping:
            self.CheckJump()

        if self.jumping:
            self.MakeJump()

        if self.moving:
            if self.move_distance != self.max_move_distance:
                if self.move_delay != self.max_move_delay:
                    self.move_delay += 1
                else:
                    self.move_delay = 0
                    self.move_distance += 1

                    if self.direction == "right":
                        if not self.pos_x>=self.x_border_right:
                            self.pos_x+=1
                        else:
                            self.SetIdleAnim("left")
                    else:
                        if self.pos_x!=self.x_border:
                            self.pos_x-=1
                        else:
                            self.SetIdleAnim("right")

            else:
                self.move_distance = 0
                self.moving = False

                change_direction = choice([True, False])
                if change_direction:
                    self.direction = "right" if self.direction == "left" else "left"
                
                anim = choice([f"idle_{self.direction}", "sitting", "3_idle", "4_idle"])
                self.SetAnimation(anim)

                # Excel处理期间暂停自动截图评论，避免与批处理抢占API配额
                if self.InitChat() and (not self.excel_processing):
                    self.chat = True
                    self.GetScreenshotComment()

        if not self.moving:
            if self.idle_delay == self.max_idle_delay:
                self.moving = True
                self.chat = False
                self.idle_delay = 0
                self.SetAnimation(f"move_{self.direction}")
            else:
                self.idle_delay += 1
        self.UpdateFrame()

    def UpdateFrame(self):
        self.frame_index = (
            0 if self.frame_index >= self.max_frame_index else self.frame_index
        )

        if self.delay_count == self.max_delay:
            self.delay_count = 0
            if not len(self.current_animation) == 1:
                self.frame_index += 1
        else:
            self.delay_count += 1

        self.chat_window_root.wm_attributes("-alpha", 1 if self.chat else 0)

        self.UpdateChatWindow()
        if not self.tts_begun:
            self.UpdateRootWindow()
    
    def LoadAnimations(self):
        anim_dict = {}
        
        # 使用绝对路径定位 sprites 目录
        base_dir = Path(__file__).parent.parent
        sprites_dir = base_dir / "sprites"

        if not sprites_dir.exists():
            print(f"Error: Sprites directory not found at {sprites_dir}")
            return {}

        for animation_directory in listdir(str(sprites_dir)):
            # 构造子目录的完整路径
            current_anim_dir = sprites_dir / animation_directory
            
            # 确保是目录
            if not current_anim_dir.is_dir():
                continue

            file_list = listdir(str(current_anim_dir))

            for image_file in file_list:
                # 构造图片完整路径
                image_path = current_anim_dir / image_file
                
                try:
                    pil_image = Image.open(str(image_path))
                    resized_image = pil_image.resize((100,100))
                    tk_image = ImageTk.PhotoImage(resized_image)

                    if animation_directory in anim_dict:
                        anim_dict[animation_directory].append(tk_image)
                    else:
                        anim_dict[animation_directory] = [tk_image]
                except Exception:
                    # 忽略无法打开的文件
                    continue
        return anim_dict

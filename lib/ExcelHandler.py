"""
Excel操作处理器 - 协调API调用和数据管理
解析自然语言指令并执行Excel操作
"""

import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from lib.DashScopeAPIManager import DashScopeAPIManager
from lib.ExcelDataManager import ExcelDataManager


class ExcelHandler:
    """Excel操作协调器"""

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        base_url: Optional[str] = None,
        system_prompt: Optional[str] = None,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None
    ):
        """
        初始化Excel处理器
        
        Args:
            api_key: API密钥
            model: 模型名称
            base_url: API基础URL
            system_prompt: 系统提示词
        """
        config = self._load_settings()

        # 自动回填未传入参数，避免“未接入配置”问题
        api_key = api_key or config.get("api_key", "") or os.environ.get("SK", "")
        model = model or config.get("model", "")
        base_url = base_url or config.get("base_url", "")
        if temperature is None:
            temperature = config.get("temperature", 0.3)
        if max_tokens is None:
            max_tokens = config.get("max_tokens", 500)
        if system_prompt is None:
            system_prompt = config.get("system_prompt")

        if not api_key or not model or not base_url:
            raise ValueError("API配置缺失，请在 config/settings.json 中填写 api_key/model/base_url，或设置环境变量 SK")

        self.api_manager = DashScopeAPIManager(api_key, model, base_url)
        self.data_manager = ExcelDataManager()
        self.logger = logging.getLogger("ExcelHandler")
        self.temperature = temperature
        self.max_tokens = max_tokens
        
        self.system_prompt = system_prompt or (
            "你是一个Excel数据处理助手。根据用户的指令，对单元格内容进行转换处理。"
            "只输出处理结果，不要输出任何解释或额外内容。"
        )

    @staticmethod
    def _load_settings(config_path: Optional[str] = None) -> Dict[str, Any]:
        """读取settings.json，供初始化参数回填"""
        if config_path is None:
            base_dir = Path(__file__).parent.parent
            config_path = str(base_dir / "config" / "settings.json")

        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    @classmethod
    def from_config(cls, config_path: str = None) -> 'ExcelHandler':
        """
        从配置文件创建实例
        
        Args:
            config_path: 配置文件路径
        """
        if config_path is None:
            # 默认配置文件路径
            base_dir = Path(__file__).parent.parent
            config_path = base_dir / "config" / "settings.json"
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        return cls(
            api_key=config['api_key'],
            model=config['model'],
            base_url=config['base_url'],
            system_prompt=config.get('system_prompt'),
            temperature=config.get('temperature', 0.3),
            max_tokens=config.get('max_tokens', 500)
        )

    def parse_instruction(self, instruction: str) -> Dict:
        """
        解析自然语言指令
        
        支持的指令格式:
            - "把A列翻译成中文" (转换操作)
            - "将B列内容总结" (转换操作)
            - "在A列添加1到36的数字" (生成操作)
            - "在B列填充星期一到星期日" (生成操作)
        
        Args:
            instruction: 用户输入的自然语言指令
            
        Returns:
            {
                'valid': True/False,
                'target_column': '列名',
                'operation_type': 'transform' | 'generate',
                'ai_prompt': '给AI的指令',
                'generate_params': {...}  # 生成操作的参数
            }
        """
        instruction = instruction.strip()

        # 复制操作: 把第1列复制到第2、3列 / 将A列复制到B,C列
        if '复制到' in instruction and '列' in instruction:
            left, right = instruction.split('复制到', 1)
            source_tokens = re.findall(r'[A-Za-z]+|\d+', left)
            target_tokens = re.findall(r'[A-Za-z]+|\d+', right)
            if source_tokens and target_tokens:
                source_ref = self._token_to_column_ref(source_tokens[-1])
                target_refs: List[Any] = []
                for token in target_tokens:
                    col_ref = self._token_to_column_ref(token)
                    if col_ref == source_ref:
                        continue
                    if col_ref not in target_refs:
                        target_refs.append(col_ref)
                if target_refs:
                    return {
                        'valid': True,
                        'operation_type': 'copy',
                        'source_column': source_ref,
                        'target_columns': target_refs,
                        'target_column': source_ref,
                        'ai_prompt': f"复制到{','.join([str(x) for x in target_refs])}列"
                    }
         
        # 检查是否是生成操作 - 模式1: "在A列添加xxx"
        pattern1 = r'(?:在|向)?([A-Za-z]+)列(?:中)?(?:添加|填充|写入|生成)(.+)'
        match = re.search(pattern1, instruction, re.IGNORECASE)
        if match:
            col_ref = match.group(1).upper()
            content = match.group(2).strip()
            generate_params = self._parse_generate_content(content)
            return {
                'valid': True,
                'target_column': col_ref,
                'operation_type': 'generate',
                'ai_prompt': content,
                'generate_params': generate_params
            }
        
        # 检查是否是生成操作 - 模式2: "在第1列添加xxx"
        pattern2 = r'(?:在|向)?第(\d+)列(?:中)?(?:添加|填充|写入|生成)(.+)'
        match = re.search(pattern2, instruction, re.IGNORECASE)
        if match:
            col_idx = int(match.group(1)) - 1
            content = match.group(2).strip()
            generate_params = self._parse_generate_content(content)
            return {
                'valid': True,
                'target_column': col_idx,
                'operation_type': 'generate',
                'ai_prompt': content,
                'generate_params': generate_params
            }
        
        # 列名提取模式（转换操作）
        patterns = [
            r'(?:把|将)?(?:第)?([A-Za-z]+)列',  # A列、第A列
            r'(?:把|将)?第(\d+)列',              # 第1列
            r'列([A-Za-z]+)',                    # 列A
        ]
        
        target_column = None
        for pattern in patterns:
            match = re.search(pattern, instruction, re.IGNORECASE)
            if match:
                col_ref = match.group(1)
                if col_ref.isdigit():
                    target_column = int(col_ref) - 1
                else:
                    target_column = col_ref.upper()
                break
        
        if target_column is None:
            return {'valid': False, 'error': '无法识别目标列，请使用如"A列"或"第1列"的格式'}
        
        # 提取操作描述
        operation = instruction
        for pattern in patterns:
            operation = re.sub(pattern, '', operation, flags=re.IGNORECASE)
        
        operation = re.sub(r'^[把将对]', '', operation)
        operation = operation.strip()
        
        if not operation:
            return {'valid': False, 'error': '请描述要执行的操作，如"翻译成中文"'}

        clear_keywords = ('删除', '清空', '置空', '清除')
        if any(k in operation for k in clear_keywords):
            return {
                'valid': True,
                'target_column': target_column,
                'operation_type': 'clear',
                'ai_prompt': operation
            }
        
        return {
            'valid': True,
            'target_column': target_column,
            'operation_type': 'transform',
            'ai_prompt': operation
        }

    def _token_to_column_ref(self, token: str):
        token = token.strip()
        if token.isdigit():
            return int(token) - 1
        return token.upper()

    def _resolve_column_ref(self, ref: Any, columns: List[Any], allow_new: bool = False):
        if isinstance(ref, int):
            if 0 <= ref < len(columns):
                return columns[ref]
            if allow_new:
                return chr(ord('A') + ref) if ref < 26 else f"Col{ref + 1}"
            return None

        if isinstance(ref, str):
            if ref in columns:
                return ref
            if ref.isdigit():
                return self._resolve_column_ref(int(ref) - 1, columns, allow_new=allow_new)
            return ref if allow_new else None

        return None

    def _parse_generate_content(self, content: str) -> Dict:
        """解析生成内容的参数"""
        # 匹配数字范围: "1到36", "1-100"
        num_match = re.search(r'(\d+)\s*(?:到|至|-)\s*(\d+)', content)
        if num_match:
            start, end = int(num_match.group(1)), int(num_match.group(2))
            return {
                'type': 'number_sequence',
                'start': start,
                'end': end,
                'values': list(range(start, end + 1))
            }
        
        # 匹配列表: "星期一到星期日", "周一,周二,周三"
        weekdays_cn = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
        weekdays_short = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
        months = ['一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月']
        
        if '星期' in content:
            return {'type': 'list', 'values': weekdays_cn}
        if '周' in content and ('到' in content or '至' in content):
            return {'type': 'list', 'values': weekdays_short}
        if '月' in content:
            return {'type': 'list', 'values': months}
        
        # 逗号分隔的列表
        if ',' in content or '，' in content:
            items = re.split(r'[,，]', content)
            return {'type': 'list', 'values': [item.strip() for item in items if item.strip()]}
        
        # 默认：需要AI生成
        return {'type': 'ai_generate', 'prompt': content}

    def execute_operation(
        self,
        file_path: str,
        instruction: str,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[bool, str]:
        """
        执行Excel操作
        
        Args:
            file_path: Excel文件路径
            instruction: 自然语言指令
            progress_callback: 进度回调，接收状态文本
            
        Returns:
            (是否成功, 操作摘要/错误信息)
        """
        # 1. 加载文件
        if progress_callback:
            progress_callback("正在加载文件...")
        
        success, error = self.data_manager.load_file(file_path)
        if not success:
            return False, f"文件加载失败: {error}"
        
        meta = self.data_manager.get_meta_info()
        columns = meta['columns']
        
        # 2. 解析指令
        if progress_callback:
            progress_callback("正在解析指令...")
        
        parsed = self.parse_instruction(instruction)
        if not parsed['valid']:
            return False, parsed.get('error', '无法理解指令')

        self.data_manager.last_save_message = ""
        self.data_manager.last_save_success = True

        operation_type = parsed.get('operation_type', 'transform')
        if operation_type == 'copy':
            return self._execute_copy(parsed, columns, meta, progress_callback)
        if operation_type == 'clear':
            return self._execute_clear(parsed, columns, meta, progress_callback)
        
        # 处理列引用
        target_column = parsed['target_column']
        target_column = self._resolve_column_ref(
            target_column,
            columns,
            allow_new=(operation_type == 'generate')
        )
        
        if operation_type == 'generate':
            return self._execute_generate(
                target_column, parsed, meta, progress_callback
            )
        else:
            return self._execute_transform(
                target_column, parsed, columns, progress_callback
            )

    def _execute_clear(
        self,
        parsed: Dict[str, Any],
        columns: List[Any],
        meta: Dict[str, Any],
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[bool, str]:
        """执行清空列内容操作（不调用AI）"""
        target_ref = parsed.get('target_column')
        target_col = self._resolve_column_ref(target_ref, columns, allow_new=False)
        if target_col is None:
            available = ', '.join([str(c) for c in columns[:8]])
            return False, f"目标列不存在: {target_ref}\n可用列: {available}"

        if progress_callback:
            progress_callback("正在清空列内容...")

        # 保留列结构，仅清空单元格内容
        self.data_manager.df[target_col] = None
        self.data_manager.modified = True
        save_ok, save_msg = self.data_manager.save_file()

        row_count = len(self.data_manager.df)
        summary = (
            f"文件: {meta['file_name']}\n"
            f"操作: {parsed.get('ai_prompt', '清空列')}\n"
            f"目标: {target_col}列\n"
            f"清空: {row_count}行"
        )
        if save_msg:
            summary = f"{summary}\n{save_msg}"
        if not save_ok:
            return False, summary
        return True, summary

    def _execute_generate(
        self,
        target_column: str,
        parsed: Dict,
        meta: Dict,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[bool, str]:
        """执行生成操作"""
        if progress_callback:
            progress_callback("正在生成数据...")
        
        generate_params = parsed.get('generate_params', {})
        gen_type = generate_params.get('type', 'ai_generate')
        
        if gen_type in ('number_sequence', 'list'):
            # 直接生成，无需AI
            values = generate_params.get('values', [])
            if not values:
                return False, "无法解析要生成的内容"
            
            # 构建更新列表
            updates = []
            for idx, value in enumerate(values):
                updates.append({
                    'row': idx,
                    'col': target_column,
                    'result': value,
                    'success': True
                })
            
            # 确保列存在
            if target_column not in self.data_manager.df.columns:
                self.data_manager.df[target_column] = None
            
            # 扩展行数
            current_rows = len(self.data_manager.df)
            needed_rows = len(values)
            if needed_rows > current_rows:
                import pandas as pd
                extra_rows = pd.DataFrame(index=range(current_rows, needed_rows))
                self.data_manager.df = pd.concat([self.data_manager.df, extra_rows], ignore_index=True)
            
            # 更新数据
            success_count, fail_count = self.data_manager.update_range(updates, auto_save=True)
            
            summary = self._generate_summary(
                meta['file_name'],
                parsed['ai_prompt'],
                target_column,
                success_count,
                fail_count
            )
            summary = self._append_save_message(summary)
            if success_count == 0:
                return False, f"{summary}\n提示: 生成写入未成功，请检查指令或稍后重试"
            return True, summary
        
        else:
            return self._execute_ai_generate(
                target_column=target_column,
                parsed=parsed,
                meta=meta,
                progress_callback=progress_callback
            )

    def _execute_transform(
        self,
        target_column: str,
        parsed: Dict,
        columns: list,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[bool, str]:
        """执行转换操作"""
        # 验证列存在
        if target_column not in columns:
            available = ', '.join([str(c) for c in columns[:5]])
            if len(columns) > 5:
                available += f' ... (共{len(columns)}列)'
            return False, f"列'{target_column}'不存在\n可用列: {available}"
        
        # 获取目标数据
        if progress_callback:
            progress_callback("正在读取数据...")
        
        cells = self.data_manager.get_column_data(target_column)
        if not cells:
            return False, f"列'{target_column}'为空或没有有效数据"
        
        total_cells = len(cells)
        
        # 批量处理
        def batch_progress(current, total):
            if progress_callback:
                progress_callback(f"处理中 {current}/{total}...")
        
        if progress_callback:
            progress_callback(f"开始处理 {total_cells} 个单元格...")
        
        results = self.api_manager.process_batch(
            cells,
            self.system_prompt,
            parsed['ai_prompt'],
            temperature=self.temperature,
            max_tokens=self.max_tokens,
            progress_callback=batch_progress
        )
        
        # 更新Excel
        if progress_callback:
            progress_callback("正在保存结果...")
        
        success_count, fail_count = self.data_manager.update_range(results, auto_save=True)
        
        # 生成摘要
        meta = self.data_manager.get_meta_info()
        summary = self._generate_summary(
            meta['file_name'],
            parsed['ai_prompt'],
            target_column,
            success_count,
            fail_count
        )
        summary = self._append_save_message(summary)
        if success_count == 0:
            return False, f"{summary}\n提示: 批处理全部失败，可能是API过载(429)，请稍后重试"

        return True, summary

    def _parse_ai_generated_values(self, ai_text: str, row_count: int) -> List[str]:
        """将AI输出解析为行值列表"""
        if row_count <= 0:
            return []

        cleaned = ai_text.strip()
        if cleaned.startswith("```"):
            cleaned = re.sub(r"^```[a-zA-Z]*\n?", "", cleaned)
            cleaned = re.sub(r"\n?```$", "", cleaned).strip()

        values: List[str] = []
        try:
            parsed_json = json.loads(cleaned)
            if isinstance(parsed_json, list):
                values = [str(v).strip() for v in parsed_json if str(v).strip()]
        except Exception:
            pass

        if not values:
            lines = [line.strip() for line in cleaned.splitlines() if line.strip()]
            if len(lines) > 1:
                values = lines
            else:
                values = [
                    item.strip() for item in re.split(r"[,，;；]", cleaned)
                    if item.strip()
                ]

        if not values:
            return []

        if len(values) == 1 and row_count > 1:
            values = values * row_count
        elif len(values) < row_count:
            values.extend([values[-1]] * (row_count - len(values)))
        elif len(values) > row_count:
            values = values[:row_count]

        return values

    def _execute_ai_generate(
        self,
        target_column: str,
        parsed: Dict[str, Any],
        meta: Dict[str, Any],
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[bool, str]:
        """执行AI生成操作：按当前表格行数生成目标列"""
        if progress_callback:
            progress_callback("正在调用AI生成数据...")

        row_count = max(len(self.data_manager.df), 1)
        prompt = parsed.get('ai_prompt', '').strip()
        if not prompt:
            return False, "生成指令为空"

        user_prompt = (
            f"请根据需求生成 {row_count} 条数据。\n"
            f"需求: {prompt}\n"
            "输出要求:\n"
            "1. 只输出数据，不要解释\n"
            "2. 优先输出 JSON 数组，如 [\"值1\",\"值2\"]\n"
            "3. 条数必须与要求一致"
        )

        gen_tokens = max(self.max_tokens, min(2000, row_count * 32))
        success, result, error = self.api_manager.generate_text(
            system_prompt=(
                "你是Excel数据生成助手。"
                "根据用户需求生成指定条数的数据，保持可直接写入单元格。"
            ),
            user_prompt=user_prompt,
            temperature=self.temperature,
            max_tokens=gen_tokens
        )
        if not success or not result:
            return False, f"AI生成失败: {error}"

        values = self._parse_ai_generated_values(result, row_count)
        if not values:
            return False, "AI返回内容无法解析为可写入数据"

        if target_column not in self.data_manager.df.columns:
            self.data_manager.df[target_column] = None

        current_rows = len(self.data_manager.df)
        needed_rows = len(values)
        if needed_rows > current_rows:
            import pandas as pd

            extra_rows = pd.DataFrame(index=range(current_rows, needed_rows))
            self.data_manager.df = pd.concat([self.data_manager.df, extra_rows], ignore_index=True)

        updates = []
        for idx, value in enumerate(values):
            updates.append({
                'row': idx,
                'col': target_column,
                'result': value,
                'success': True
            })

        if progress_callback:
            progress_callback("正在写回Excel...")
        success_count, fail_count = self.data_manager.update_range(updates, auto_save=True)

        summary = self._generate_summary(
            meta['file_name'],
            parsed['ai_prompt'],
            target_column,
            success_count,
            fail_count
        )
        summary = self._append_save_message(summary)
        if success_count == 0:
            return False, f"{summary}\n提示: AI生成未写入有效结果，请稍后重试"
        return True, summary

    def _execute_copy(
        self,
        parsed: Dict[str, Any],
        columns: List[Any],
        meta: Dict[str, Any],
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> Tuple[bool, str]:
        """执行列复制操作，不调用AI"""
        source_ref = parsed.get("source_column")
        target_refs = parsed.get("target_columns", [])

        source_col = self._resolve_column_ref(source_ref, columns, allow_new=False)
        if source_col is None:
            available = ', '.join([str(c) for c in columns[:8]])
            return False, f"源列不存在: {source_ref}\n可用列: {available}"

        target_cols: List[Any] = []
        for ref in target_refs:
            col = self._resolve_column_ref(ref, columns, allow_new=True)
            if col is None or col == source_col:
                continue
            if col not in target_cols:
                target_cols.append(col)

        if not target_cols:
            return False, "未解析到有效目标列"

        if progress_callback:
            progress_callback("正在复制列数据...")

        source_series = self.data_manager.df[source_col]
        for target_col in target_cols:
            if target_col not in self.data_manager.df.columns:
                self.data_manager.df[target_col] = None
            self.data_manager.df[target_col] = source_series

        self.data_manager.modified = True
        save_ok, save_msg = self.data_manager.save_file()

        row_count = len(self.data_manager.df)
        summary = self._generate_copy_summary(
            file_name=meta['file_name'],
            source_column=source_col,
            target_columns=target_cols,
            row_count=row_count,
            save_message=save_msg
        )
        if not save_ok:
            return False, summary
        return True, summary

    def _generate_summary(
        self,
        file_name: str,
        instruction: str,
        column: str,
        success_count: int,
        fail_count: int
    ) -> str:
        """生成操作摘要"""
        lines = [
            f"文件: {file_name}",
            f"操作: {instruction}",
            f"目标: {column}列",
            f"成功: {success_count}行"
        ]
        
        if fail_count > 0:
            lines.append(f"失败: {fail_count}行")
        
        return '\n'.join(lines)

    def _generate_copy_summary(
        self,
        file_name: str,
        source_column: Any,
        target_columns: List[Any],
        row_count: int,
        save_message: str = ""
    ) -> str:
        lines = [
            f"文件: {file_name}",
            "操作: 列复制",
            f"来源: {source_column}列",
            f"目标: {', '.join([str(c) for c in target_columns])}列",
            f"写入: {row_count}行 x {len(target_columns)}列"
        ]
        if save_message:
            lines.append(save_message)
        return '\n'.join(lines)

    def _append_save_message(self, summary: str) -> str:
        msg = self.data_manager.last_save_message
        if msg and msg not in summary:
            return f"{summary}\n{msg}"
        return summary

    def test_api_connection(self) -> Tuple[bool, str]:
        """测试API连接"""
        return self.api_manager.test_connection()

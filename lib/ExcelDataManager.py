"""
Excel数据管理器 - 负责Excel文件的读写操作
基于pandas实现，支持xlsx/xls/csv格式
"""

import logging
from datetime import datetime
from pathlib import Path
from typing import Tuple, List, Dict, Any, Optional

import pandas as pd


class ExcelDataManager:
    """Excel数据管理器"""

    def __init__(self):
        self.df: Optional[pd.DataFrame] = None
        self.file_path: Optional[str] = None
        self.modified = False
        self.logger = logging.getLogger("ExcelDataManager")
        self.last_save_success: bool = True
        self.last_save_path: Optional[str] = None
        self.last_save_message: str = ""

    def _ensure_column_compatible(self, col: str, value: Any) -> None:
        """
        避免向数值列写入字符串时触发未来版本pandas的dtype错误
        """
        if self.df is None:
            return
        if col not in self.df.columns:
            return

        series = self.df[col]
        if pd.api.types.is_object_dtype(series):
            return

        # 当列是数值型而新值是字符串时，先转object，避免FutureWarning
        if isinstance(value, str):
            self.df[col] = series.astype("object")

    def _build_fallback_save_path(self, original_path: Path) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return original_path.with_name(f"{original_path.stem}_pet_saved_{timestamp}{original_path.suffix}")

    def load_file(self, file_path: str) -> Tuple[bool, str]:
        """
        加载Excel或CSV文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            (是否成功, 错误消息)
        """
        try:
            path = Path(file_path)
            if not path.exists():
                return False, f"文件不存在: {file_path}"
            
            file_ext = path.suffix.lower()
            
            if file_ext in ['.xlsx', '.xls', '.xlsm']:
                self.df = pd.read_excel(file_path)
            elif file_ext == '.csv':
                # 尝试多种编码
                for encoding in ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']:
                    try:
                        self.df = pd.read_csv(file_path, encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    return False, "无法识别文件编码"
            else:
                return False, f"不支持的文件格式: {file_ext}"
            
            self.file_path = file_path
            self.modified = False
            self.logger.info(f"已加载文件: {file_path}, {len(self.df)}行")
            return True, ""
            
        except Exception as e:
            self.logger.error(f"加载文件失败: {e}")
            return False, f"加载失败: {str(e)}"

    def save_file(self, file_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        保存文件
        
        Args:
            file_path: 保存路径，为空则覆盖原文件
            
        Returns:
            (是否成功, 错误消息)
        """
        if self.df is None:
            return False, "没有数据可保存"
        
        save_path = file_path or self.file_path
        if not save_path:
            return False, "未指定保存路径"
        
        try:
            path = Path(save_path)
            file_ext = path.suffix.lower()
            
            if file_ext == '.csv':
                self.df.to_csv(save_path, index=False, encoding='utf-8-sig')
            else:
                self.df.to_excel(save_path, index=False, engine='openpyxl')
            
            self.modified = False
            self.last_save_success = True
            self.last_save_path = str(path)
            self.last_save_message = f"已保存文件: {path}"
            self.logger.info(f"已保存文件: {save_path}")
            return True, ""
        except PermissionError:
            # 文件被Excel占用时自动另存，避免操作完全失败
            path = Path(save_path)
            fallback_path = self._build_fallback_save_path(path)
            try:
                file_ext = fallback_path.suffix.lower()
                if file_ext == '.csv':
                    self.df.to_csv(fallback_path, index=False, encoding='utf-8-sig')
                else:
                    self.df.to_excel(fallback_path, index=False, engine='openpyxl')
                self.modified = False
                self.last_save_success = True
                self.last_save_path = str(fallback_path)
                self.last_save_message = f"原文件被占用，已另存为: {fallback_path}"
                self.logger.warning(self.last_save_message)
                return True, self.last_save_message
            except Exception as e:
                self.last_save_success = False
                self.last_save_path = None
                self.last_save_message = f"保存失败(原文件占用且另存失败): {str(e)}"
                self.logger.error(self.last_save_message)
                return False, self.last_save_message
            
        except Exception as e:
            self.last_save_success = False
            self.last_save_path = None
            self.last_save_message = f"保存失败: {str(e)}"
            self.logger.error(f"保存文件失败: {e}")
            return False, f"保存失败: {str(e)}"

    def get_columns(self) -> List[str]:
        """获取所有列名"""
        if self.df is None:
            return []
        return list(self.df.columns)

    def get_column_data(
        self,
        column: str,
        start_row: int = 0,
        end_row: Optional[int] = None
    ) -> List[Dict]:
        """
        获取指定列的数据
        
        Args:
            column: 列名
            start_row: 起始行(从0开始)
            end_row: 结束行(不包含)，为空则到最后一行
            
        Returns:
            [{'row': 0, 'col': 'A', 'content': '...'}]
        """
        if self.df is None:
            return []
        
        if column not in self.df.columns:
            return []
        
        end = end_row if end_row is not None else len(self.df)
        end = min(end, len(self.df))
        
        result = []
        for row_idx in range(start_row, end):
            value = self.df.loc[row_idx, column]
            # 跳过空值
            if pd.isna(value) or str(value).strip() == '':
                continue
            result.append({
                'row': row_idx,
                'col': column,
                'content': str(value)
            })
        
        return result

    def update_cell(self, row: int, col: str, value: Any) -> bool:
        """
        更新单个单元格
        
        Args:
            row: 行索引
            col: 列名
            value: 新值
            
        Returns:
            是否成功
        """
        if self.df is None:
            return False
        
        try:
            self._ensure_column_compatible(col, value)
            self.df.loc[row, col] = value
            self.modified = True
            return True
        except Exception as e:
            self.logger.error(f"更新单元格失败: {e}")
            return False

    def update_range(
        self,
        updates: List[Dict],
        auto_save: bool = True
    ) -> Tuple[int, int]:
        """
        批量更新单元格
        
        Args:
            updates: [{'row': 0, 'col': 'A', 'result': '...', 'success': True}]
            auto_save: 是否自动保存
            
        Returns:
            (成功数量, 失败数量)
        """
        if self.df is None:
            return 0, len(updates)
        
        success_count = 0
        fail_count = 0
        
        for update in updates:
            if not update.get('success', False):
                fail_count += 1
                continue
            
            try:
                row = update['row']
                col = update['col']
                value = update['result']

                self._ensure_column_compatible(col, value)
                self.df.loc[row, col] = value
                success_count += 1
                
            except Exception as e:
                self.logger.error(f"更新失败 [{row},{col}]: {e}")
                fail_count += 1
        
        if success_count > 0:
            self.modified = True
            
            if auto_save and self.file_path:
                ok, msg = self.save_file()
                if not ok:
                    self.logger.error(f"自动保存失败: {msg}")
        
        return success_count, fail_count

    def get_meta_info(self) -> Dict[str, Any]:
        """
        获取文件元信息
        
        Returns:
            {'loaded': True, 'file_name': '...', 'rows': 100, 'columns': [...]}
        """
        if self.df is None:
            return {'loaded': False}
        
        return {
            'loaded': True,
            'file_path': self.file_path,
            'file_name': Path(self.file_path).name if self.file_path else None,
            'rows': len(self.df),
            'columns': list(self.df.columns),
            'modified': self.modified
        }

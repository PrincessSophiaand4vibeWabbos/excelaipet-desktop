import io
from pathlib import Path
import re
from typing import Any, Dict, Optional, Tuple

from PIL import Image

from lib.ExcelUILocator import ExcelUILocator


class LocalExcelVision:
    """本地Excel界面识别器（可选：Ultralytics YOLO）"""

    def __init__(self, config: Dict[str, Any]):
        local_cfg = config.get("local_vision", {}) if isinstance(config, dict) else {}
        self.enabled = bool(local_cfg.get("enabled", False))
        self.model_path = str(local_cfg.get("model_path", "")).strip()
        self.conf_threshold = float(local_cfg.get("conf_threshold", 0.25))
        self.max_det = int(local_cfg.get("max_det", 20))
        self.predict_imgsz = int(local_cfg.get("predict_imgsz", 640))

        self.model = None
        self.error_message = ""
        self._locator = ExcelUILocator()
        self.last_presence: Dict[str, Any] = {}

        if self.enabled:
            self._load_model()

    def _load_model(self):
        if not self.model_path:
            self.error_message = "local_vision.model_path 未配置"
            return

        model_file = Path(self.model_path)
        if not model_file.exists():
            self.error_message = f"本地视觉模型不存在: {model_file}"
            return

        try:
            from ultralytics import YOLO
        except Exception as e:
            self.error_message = f"未安装ultralytics或导入失败: {e}"
            return

        try:
            self.model = YOLO(str(model_file))
        except Exception as e:
            self.error_message = f"加载本地视觉模型失败: {e}"

    def is_ready(self) -> bool:
        return self.enabled and self.model is not None

    @staticmethod
    def _resolve_class_name(names: Any, cls_id: int) -> str:
        if isinstance(names, dict):
            return str(names.get(cls_id, cls_id))
        if isinstance(names, list) and 0 <= cls_id < len(names):
            return str(names[cls_id])
        return str(cls_id)

    def _extract_excel_roi(self, image_np):
        h, w = image_np.shape[:2]
        if not self._locator.find_excel_window() or not self._locator.excel_rect:
            return image_np, "全屏截图"

        left, top, right, bottom = self._locator.excel_rect
        x1 = max(0, min(w - 1, int(left)))
        y1 = max(0, min(h - 1, int(top)))
        x2 = max(0, min(w, int(right)))
        y2 = max(0, min(h, int(bottom)))

        if x2 - x1 < 120 or y2 - y1 < 120:
            return image_np, "全屏截图"

        return image_np[y1:y2, x1:x2], f"Excel窗口({x1},{y1},{x2},{y2})"

    @staticmethod
    def _default_sheet_roi_bounds(width: int, height: int) -> Tuple[int, int, int, int]:
        x1 = int(width * 0.05)
        y1 = int(height * 0.22)
        x2 = int(width * 0.98)
        y2 = int(height * 0.93)
        return x1, y1, x2, y2

    def _get_sheet_bounds(self, roi_np, cell_bbox) -> Tuple[int, int, int, int]:
        h, w = roi_np.shape[:2]
        if cell_bbox is None:
            return self._default_sheet_roi_bounds(w, h)

        x1, y1, x2, y2 = [int(v) for v in cell_bbox]
        x1 = max(0, x1 - 8)
        y1 = max(0, y1 - 8)
        x2 = min(w, x2 + 8)
        y2 = min(h, y2 + 8)
        if x2 - x1 < 60 or y2 - y1 < 60:
            return self._default_sheet_roi_bounds(w, h)
        return x1, y1, x2, y2

    def _build_text_like_map(self, roi_np, cell_bbox):
        try:
            import cv2
        except Exception:
            return None

        x1, y1, x2, y2 = self._get_sheet_bounds(roi_np, cell_bbox)
        sheet = roi_np[y1:y2, x1:x2]
        if sheet.size == 0:
            return None

        gray = cv2.cvtColor(sheet, cv2.COLOR_RGB2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        kx = max(20, sheet.shape[1] // 12)
        ky = max(20, sheet.shape[0] // 12)
        horizontal = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (kx, 1)),
        )
        vertical = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (1, ky)),
        )
        grid = cv2.bitwise_or(horizontal, vertical)
        text_like = cv2.bitwise_and(binary, cv2.bitwise_not(grid))
        return {
            "sheet": sheet,
            "text_like": text_like,
            "vertical": vertical,
            "bounds": (x1, y1, x2, y2),
        }

    @staticmethod
    def _component_stats(binary_map):
        import cv2

        num_labels, _, stats, _ = cv2.connectedComponentsWithStats(binary_map, connectivity=8)
        components = 0
        area_sum = 0
        for idx in range(1, num_labels):
            area = int(stats[idx, cv2.CC_STAT_AREA])
            bw = int(stats[idx, cv2.CC_STAT_WIDTH])
            bh = int(stats[idx, cv2.CC_STAT_HEIGHT])
            if 8 <= area <= 1600 and bw <= 120 and bh <= 80:
                components += 1
                area_sum += area
        total_pixels = float(binary_map.shape[0] * binary_map.shape[1])
        pixel_ratio = area_sum / total_pixels if total_pixels > 0 else 0.0
        return components, pixel_ratio

    def _estimate_data_presence(self, roi_np, cell_bbox):
        features = self._build_text_like_map(roi_np, cell_bbox)
        if features is None:
            return "未知", 0.0, {"components": 0, "pixel_ratio": 0.0}

        text_like = features["text_like"]
        components, pixel_ratio = self._component_stats(text_like)

        density_score = min(1.0, pixel_ratio * 220.0)
        component_score = min(1.0, components / 45.0)
        score = max(density_score, component_score)

        if score >= 0.45:
            label = "有数据"
        elif score >= 0.12:
            label = "疑似有少量数据"
        else:
            label = "未检测到明显数据"

        return label, score, {"components": components, "pixel_ratio": pixel_ratio}

    @staticmethod
    def _split_ranges(indices):
        if not indices:
            return []
        ranges = []
        start = indices[0]
        prev = indices[0]
        for v in indices[1:]:
            if v == prev + 1:
                prev = v
            else:
                ranges.append((start, prev))
                start = v
                prev = v
        ranges.append((start, prev))
        return ranges

    @staticmethod
    def _col_label_to_index(label: str) -> int:
        value = 0
        for ch in label:
            value = value * 26 + (ord(ch) - ord("A") + 1)
        return value

    @staticmethod
    def _chinese_num_to_int(s: str) -> Optional[int]:
        mapping = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
        if not s:
            return None
        if s == "十":
            return 10
        if "十" in s:
            parts = s.split("十")
            left = mapping.get(parts[0], 1 if parts[0] == "" else None)
            right = mapping.get(parts[1], 0 if len(parts) > 1 and parts[1] == "" else None)
            if left is None or right is None:
                return None
            return left * 10 + right
        value = 0
        for ch in s:
            if ch not in mapping:
                return None
            value = value * 10 + mapping[ch]
        return value if value > 0 else None

    def _parse_target_column(self, user_instruction: str) -> Optional[int]:
        if not user_instruction:
            return None
        text = user_instruction.strip().upper()

        m = re.search(r"第\s*(\d+)\s*列", text)
        if m:
            idx = int(m.group(1))
            return idx if idx > 0 else None

        m = re.search(r"第\s*([A-Z]{1,3})\s*列", text)
        if m:
            return self._col_label_to_index(m.group(1))

        m = re.search(r"\b([A-Z]{1,3})\s*列", text)
        if m:
            return self._col_label_to_index(m.group(1))

        m = re.search(r"第\s*([一二三四五六七八九十两零]+)\s*列", user_instruction)
        if m:
            return self._chinese_num_to_int(m.group(1))
        return None

    def _estimate_target_column_presence(self, roi_np, cell_bbox, target_col_idx: int):
        if target_col_idx <= 0:
            return "未知", 0.0, {"components": 0, "pixel_ratio": 0.0, "target_col": target_col_idx}
        features = self._build_text_like_map(roi_np, cell_bbox)
        if features is None:
            return "未知", 0.0, {"components": 0, "pixel_ratio": 0.0, "target_col": target_col_idx}

        text_like = features["text_like"]
        vertical = features["vertical"]
        sh, sw = text_like.shape[:2]

        proj = (vertical > 0).sum(axis=0)
        threshold = max(6, int(sh * 0.25))
        line_idx = [int(i) for i, val in enumerate(proj.tolist()) if int(val) >= threshold]
        line_ranges = self._split_ranges(line_idx)
        centers = [int((a + b) / 2) for a, b in line_ranges]

        col_ranges = []
        if len(centers) >= 2:
            for i in range(len(centers) - 1):
                x1 = max(0, centers[i] + 1)
                x2 = min(sw, centers[i + 1] - 1)
                if x2 - x1 >= 12:
                    col_ranges.append((x1, x2))

        if not col_ranges:
            estimated_cols = 12
            width = max(12, sw // estimated_cols)
            x = 0
            while x < sw:
                x2 = min(sw, x + width)
                if x2 - x >= 12:
                    col_ranges.append((x, x2))
                x = x2

        idx = target_col_idx - 1
        if idx >= len(col_ranges):
            return "未知", 0.0, {"components": 0, "pixel_ratio": 0.0, "target_col": target_col_idx}

        x1, x2 = col_ranges[idx]
        col_text = text_like[:, x1:x2]
        components, pixel_ratio = self._component_stats(col_text)

        density_score = min(1.0, pixel_ratio * 360.0)
        component_score = min(1.0, components / 14.0)
        score = max(density_score, component_score)

        if score >= 0.38:
            label = "有数据"
        elif score >= 0.12:
            label = "疑似有少量数据"
        else:
            label = "未检测到明显数据"
        stats = {
            "components": components,
            "pixel_ratio": pixel_ratio,
            "target_col": target_col_idx,
            "range": (x1, x2),
        }
        return label, score, stats

    def get_data_presence_reply(self, user_instruction: str) -> Optional[str]:
        if not self.last_presence:
            return None

        query_col = self._parse_target_column(user_instruction)
        analyzed_col = self.last_presence.get("target_col")
        if query_col:
            if analyzed_col != query_col:
                return f"喵，第{query_col}列暂时看不清"
            label = self.last_presence.get("target_col_label")
            if label == "有数据":
                return f"喵，第{query_col}列有数据"
            if label == "疑似有少量数据":
                return f"喵，第{query_col}列疑似有少量数据"
            if label == "未检测到明显数据":
                return f"喵，第{query_col}列看起来是空的"
            return f"喵，第{query_col}列暂时看不清"

        overall_label = self.last_presence.get("overall_label")
        if overall_label == "有数据":
            return "喵~当前可见区域有数据"
        if overall_label == "疑似有少量数据":
            return "喵，我看到疑似有少量数据"
        if overall_label == "未检测到明显数据":
            return "喵，当前可见区域像是空表"
        return None

    def analyze(self, image_bytes: bytes, user_instruction: str = "") -> Tuple[bool, Optional[str], Optional[str]]:
        if not self.enabled:
            return False, None, "本地视觉识别未启用"
        if self.model is None:
            return False, None, self.error_message or "本地视觉模型未就绪"

        try:
            import numpy as np
        except Exception as e:
            return False, None, f"缺少numpy依赖: {e}"

        try:
            image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
            image_np = np.array(image)
            roi_np, roi_desc = self._extract_excel_roi(image_np)
            outputs = self.model.predict(
                source=roi_np,
                conf=self.conf_threshold,
                imgsz=self.predict_imgsz,
                max_det=self.max_det,
                verbose=False
            )

            counts: Dict[str, int] = {}
            cell_bbox = None
            if outputs:
                result = outputs[0]
                boxes = getattr(result, "boxes", None)
                if boxes is not None and len(boxes) > 0:
                    names = getattr(result, "names", {}) or getattr(self.model, "names", {})
                    xyxy = boxes.xyxy.tolist()
                    for idx, cls_id in enumerate(boxes.cls.tolist()):
                        cls_idx = int(cls_id)
                        cls_name = self._resolve_class_name(names, cls_idx)
                        counts[cls_name] = counts.get(cls_name, 0) + 1
                        if cls_name == "cell_area":
                            box = xyxy[idx]
                            if cell_bbox is None:
                                cell_bbox = box
                            else:
                                old_area = max(0.0, (cell_bbox[2] - cell_bbox[0]) * (cell_bbox[3] - cell_bbox[1]))
                                new_area = max(0.0, (box[2] - box[0]) * (box[3] - box[1]))
                                if new_area > old_area:
                                    cell_bbox = box

            data_label, data_score, stats = self._estimate_data_presence(roi_np, cell_bbox)
            target_col = self._parse_target_column(user_instruction)
            target_col_label = None
            target_col_score = 0.0
            target_col_stats = {}
            if target_col:
                target_col_label, target_col_score, target_col_stats = self._estimate_target_column_presence(
                    roi_np, cell_bbox, target_col
                )

            parts = [f"视野: {roi_desc}"]
            if counts:
                ordered = sorted(counts.items(), key=lambda x: (-x[1], x[0]))
                parts.append("UI检测: " + "，".join([f"{name}x{cnt}" for name, cnt in ordered]))
            else:
                parts.append("UI检测: 未检测到可识别的Excel界面元素")
            parts.append(
                f"数据判断: {data_label} (score={data_score:.2f}, 连通域={stats['components']}, 覆盖率={stats['pixel_ratio']:.4f})"
            )
            if target_col:
                parts.append(
                    f"目标列判断: 第{target_col}列 {target_col_label} "
                    f"(score={target_col_score:.2f}, 连通域={target_col_stats.get('components', 0)}, "
                    f"覆盖率={target_col_stats.get('pixel_ratio', 0.0):.4f})"
                )
            summary = "；".join(parts)
            self.last_presence = {
                "overall_label": data_label,
                "overall_score": data_score,
                "overall_stats": stats,
                "target_col": target_col,
                "target_col_label": target_col_label,
                "target_col_score": target_col_score,
                "target_col_stats": target_col_stats,
            }
            return True, summary, None
        except Exception as e:
            return False, None, f"本地视觉识别执行失败: {e}"

import io
import json
import os
from pathlib import Path
import re
import time
from threading import Thread
from typing import List, Optional

import pyttsx3
from PIL import ImageGrab

from lib.DashScopeAPIManager import DashScopeAPIManager
from lib.LocalExcelVision import LocalExcelVision


class Commenter:
    def __init__(self):
        # 从配置文件读取配置
        self.config = self.load_config()
        self.API_KEY = self.config.get("api_key", "") or os.environ.get("SK", "")
        self.BASE_URL = self.config.get("base_url", "https://api.moonshot.cn/v1")
        self.MODEL_NAME = self.config.get("model", "moonshot-v1-8k")
        self.VISION_MODEL = self.config.get("vision_model", self.MODEL_NAME)
        self.temperature = self.config.get("temperature", 0.3)

        self.api_manager: Optional[DashScopeAPIManager] = None
        if self.API_KEY:
            self.Configure()
        self.local_vision = LocalExcelVision(self.config)
        self.vision_models = self._build_vision_model_candidates()
        self.active_vision_model = self.vision_models[0] if self.vision_models else self.VISION_MODEL
        self._vision_models_discovered = False

        self.latest_analysis = ""
        self.latest_response = None
        self._api_unavailable_until = 0.0
        self._api_unavailable_reason = ""

    def load_config(self):
        try:
            config_path = Path(__file__).parent.parent / "config" / "settings.json"
            if config_path.exists():
                with open(config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
        return {}

    def Configure(self):
        try:
            self.api_manager = DashScopeAPIManager(
                api_key=self.API_KEY,
                model=self.MODEL_NAME,
                base_url=self.BASE_URL
            )
            return self.api_manager
        except Exception as e:
            print(f"Failed to configure API client: {e}")
            self.api_manager = None
            return None

    def TakeScreenshot(self):
        screenshot = ImageGrab.grab()
        return screenshot

    def _screenshot_to_png_bytes(self, screenshot) -> bytes:
        buffer = io.BytesIO()
        screenshot.save(buffer, format="PNG")
        return buffer.getvalue()

    def _analyze_screen(self, image_bytes: bytes, user_instruction: str = ""):
        """
        代理1：屏幕观察分析代理（视觉）
        """
        # 优先使用本地训练好的Excel识别模型
        if self.local_vision.is_ready():
            ok, result, error = self.local_vision.analyze(image_bytes, user_instruction=user_instruction)
            if ok:
                self.active_vision_model = "local_vision_model"
                return True, result, None
            print(f"Local vision error: {error}")

        if not self.api_manager:
            local_err = self.local_vision.error_message if self.local_vision.enabled else "API未初始化"
            return False, None, local_err
        if self._is_api_temporarily_unavailable():
            return False, None, f"API暂不可用: {self._api_unavailable_reason}"
        self._merge_discovered_vision_models()

        system_prompt = (
            "你是屏幕内容分析代理。请精准识别截图中的软件类型、当前页面、"
            "可能的用户操作意图和可见关键元素。回答使用中文。"
        )
        user_prompt = (
            "请分析这张屏幕截图，按以下4行输出：\n"
            "1) 软件与场景\n"
            "2) 用户正在做什么\n"
            "3) 关键可见元素\n"
            "4) 下一步可执行动作"
        )
        last_error = None
        for vision_model in self.vision_models:
            ok, result, error = self.api_manager.process_image(
                image_bytes=image_bytes,
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                temperature=0.2,
                max_tokens=400,
                model=vision_model
            )
            if ok:
                self.active_vision_model = vision_model
                return True, result, None
            last_error = error
            if self._is_connection_error(error):
                self._mark_api_temporarily_unavailable(error, cooldown_seconds=120)
            if not self._is_vision_unsupported_error(error):
                break
        return False, None, last_error

    def _compose_cat_comment(self, screen_analysis: str, user_instruction: str = ""):
        """
        代理2：桌宠评论代理（文本）
        """
        if not self.api_manager:
            return False, None, "API未初始化"

        system_prompt = (
            "你是一只可爱的Excel桌宠猫。"
            "根据屏幕观察结果给出一句可爱、简短、有帮助的中文评论。"
        )
        user_prompt = (
            "这是观察代理给出的分析：\n"
            f"{screen_analysis}\n\n"
            f"用户额外要求：{user_instruction if user_instruction else '无'}\n\n"
            "请输出1句评论，不超过24个汉字，不要换行，不要解释。"
        )
        return self.api_manager.generate_text(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            temperature=0.7,
            max_tokens=64,
            model=self.MODEL_NAME
        )

    def _fallback_comment(self, reason: str = "") -> str:
        if not self.api_manager:
            return "喵？我还没连上AI服务"
        if self._is_connection_error(reason):
            return "喵，网络有点不稳，我先本地陪你操作"
        if self._is_api_temporarily_unavailable():
            return "喵，API暂时连不上，我先用本地模式"

        success, result, _ = self.api_manager.generate_text(
            system_prompt="你是一只桌宠猫。",
            user_prompt=(
                "视觉识别暂不可用。请给一句通用、可爱、简短的中文提醒，"
                "鼓励用户继续Excel操作。20字以内。"
            ),
            temperature=0.7,
            max_tokens=50
        )
        if success and result:
            return result.strip().replace("\n", " ")
        if reason:
            return "喵呜，屏幕有点看不清"
        return "喵～继续加油处理表格"

    def _is_connection_error(self, error: Optional[str]) -> bool:
        if not error:
            return False
        if self.api_manager:
            return self.api_manager.is_connection_error(error)
        lowered = error.lower()
        return "connection error" in lowered or "timeout" in lowered

    def _mark_api_temporarily_unavailable(self, reason: str, cooldown_seconds: int = 120):
        self._api_unavailable_until = time.time() + cooldown_seconds
        self._api_unavailable_reason = reason

    def _is_api_temporarily_unavailable(self) -> bool:
        return time.time() < self._api_unavailable_until

    def _build_vision_model_candidates(self) -> List[str]:
        """构造视觉模型候选列表（配置优先，其次自动探测）"""
        candidates: List[str] = []
        configured_vision = self.config.get("vision_model", "").strip()
        if configured_vision:
            candidates.append(configured_vision)

        # 兼容Moonshot常见视觉命名，若服务端不支持会自动跳过
        if "moonshot" in self.BASE_URL.lower():
            candidates.extend([
                "moonshot-v1-vision-preview",
                "moonshot-v1-128k-vision-preview",
                "moonshot-v1-32k-vision-preview"
            ])

        # 最后回落到当前文本模型
        candidates.append(self.VISION_MODEL)
        candidates.append(self.MODEL_NAME)

        seen = set()
        uniq: List[str] = []
        for item in candidates:
            key = item.strip()
            if not key or key in seen:
                continue
            seen.add(key)
            uniq.append(key)
        return uniq

    def _merge_discovered_vision_models(self):
        """首次需要视觉时再探测可用模型，避免启动阶段阻塞"""
        if self._vision_models_discovered or not self.api_manager:
            return
        self._vision_models_discovered = True

        discovered = self.api_manager.list_models()
        if not discovered:
            return

        for model_id in discovered:
            lowered = model_id.lower()
            if not any(k in lowered for k in ("vision", "vl", "multimodal", "image")):
                continue
            if model_id not in self.vision_models:
                self.vision_models.insert(0, model_id)

    def _is_vision_unsupported_error(self, error: Optional[str]) -> bool:
        if not error:
            return False
        lowered = error.lower()
        signals = [
            "image input not supported",
            "vision",
            "unsupported",
            "input_image",
            "image_url"
        ]
        return any(s in lowered for s in signals)

    def _is_data_presence_question(self, user_instruction: str) -> bool:
        if not user_instruction:
            return False
        text = user_instruction.strip().lower()
        patterns = [
            r"是否有数据",
            r"有没有数据",
            r"有无数据",
            r"有没有内容",
            r"是否为空",
            r"是不是空表",
            r"空白",
            r"有数据",
            r"没数据",
            r"无数据",
        ]
        return any(re.search(p, text) for p in patterns)

    def _build_data_presence_reply(self, analysis: str) -> Optional[str]:
        if not analysis:
            return None
        if "数据判断: 有数据" in analysis:
            return "喵~当前可见区域有数据"
        if "数据判断: 疑似有少量数据" in analysis:
            return "喵，我看到疑似有少量数据"
        if "数据判断: 未检测到明显数据" in analysis:
            return "喵，当前可见区域像是空表"
        return None

    def _is_ambiguous_column_question(self, user_instruction: str) -> bool:
        if not user_instruction:
            return False
        text = user_instruction.strip()
        has_column = bool(
            re.search(r"第\s*(\d+|[A-Za-z]{1,3}|[一二三四五六七八九十两零]+)\s*列", text)
            or re.search(r"\b[A-Za-z]{1,3}\s*列", text)
        )
        if has_column:
            return False
        ambiguous_signals = ["没数据的列", "空列", "空白列", "这个列", "该列", "哪列"]
        return any(k in text for k in ambiguous_signals)

    def GenerateComment(self, user_instruction: str = ""):
        if not self.api_manager and not self.local_vision.is_ready():
            return "Meow? (API key not configured)"

        try:
            screenshot = self.TakeScreenshot()
            image_bytes = self._screenshot_to_png_bytes(screenshot)
        except Exception as e:
            print(f"Screenshot error: {e}")
            self.latest_response = self._fallback_comment(str(e))
            return self.latest_response

        # Multi-agent协同：先视觉分析，再生成桌宠评论
        ok_analysis, analysis, analysis_error = self._analyze_screen(image_bytes, user_instruction=user_instruction)
        if ok_analysis and analysis:
            self.latest_analysis = analysis

            if self._is_data_presence_question(user_instruction):
                if self._is_ambiguous_column_question(user_instruction):
                    self.latest_response = "喵，请说具体列号，比如第1列或B列"
                    return self.latest_response
                direct_reply = self.local_vision.get_data_presence_reply(user_instruction) or self._build_data_presence_reply(analysis)
                if direct_reply:
                    self.latest_response = direct_reply
                    return self.latest_response

            if self.api_manager and not self._is_api_temporarily_unavailable():
                ok_comment, comment, comment_error = self._compose_cat_comment(analysis, user_instruction=user_instruction)
                if ok_comment and comment:
                    self.latest_response = comment.strip().replace("\n", " ")
                else:
                    print(f"Comment agent error: {comment_error}")
                    if self._is_connection_error(comment_error):
                        self._mark_api_temporarily_unavailable(comment_error or "Connection error", cooldown_seconds=120)
                        local_reply = None
                        if self._is_data_presence_question(user_instruction):
                            local_reply = self.local_vision.get_data_presence_reply(user_instruction)
                        self.latest_response = local_reply or "喵，网络暂时连不上，我先本地陪你"
                        return self.latest_response
                    self.latest_response = self._fallback_comment(comment_error or "")
            else:
                self.latest_response = "喵，我看到了Excel界面"
        else:
            print(f"Vision agent error: {analysis_error}")
            if self._is_connection_error(analysis_error):
                self._mark_api_temporarily_unavailable(analysis_error or "Connection error", cooldown_seconds=120)
                local_reply = self.local_vision.get_data_presence_reply(user_instruction)
                if local_reply:
                    self.latest_response = local_reply
                    return self.latest_response
            self.latest_response = self._fallback_comment(analysis_error or "")

        return self.latest_response

    def ThreadedSpeaker(self):
        _ = Thread(target=self.SpeakComment)
        _.start()  # 使用 start() 而不是 run() 以实现真正的多线程

    def SpeakComment(self):
        try:
            if not self.latest_response:
                return
            tts = pyttsx3.init()
            # 尝试设置中文语音
            voices = tts.getProperty('voices')
            for voice in voices:
                if 'zh' in voice.id.lower() or 'chinese' in voice.name.lower():
                    tts.setProperty('voice', voice.id)
                    break

            tts.say(f'{self.latest_response}')
            tts.runAndWait()
        except Exception as e:
            print(f"TTS Error: {e}")

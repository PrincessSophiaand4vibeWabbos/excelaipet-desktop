"""
魔搭平台API管理器 - 兼容OpenAI接口
用于调用大模型处理Excel文本和屏幕截图
"""

import base64
import logging
import random
import time
from typing import Any, Callable, Dict, List, Optional, Tuple

from openai import OpenAI


class DashScopeAPIManager:
    """OpenAI兼容API管理器（文本+视觉）"""

    def __init__(
        self,
        api_key: str,
        model: str,
        base_url: str,
        timeout: int = 60,
        max_retries: int = 4
    ):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        self.timeout = timeout
        self.max_retries = max_retries
        self.client = OpenAI(api_key=api_key, base_url=base_url, timeout=timeout)
        self.logger = logging.getLogger("DashScopeAPI")

    def _extract_text(self, content: Any) -> str:
        """兼容不同响应格式，统一提取文本"""
        if content is None:
            return ""
        if isinstance(content, str):
            return content.strip()
        if isinstance(content, list):
            parts: List[str] = []
            for item in content:
                if isinstance(item, str):
                    parts.append(item)
                elif isinstance(item, dict):
                    text_value = item.get("text") or item.get("content")
                    if text_value:
                        parts.append(str(text_value))
            return "\n".join([p.strip() for p in parts if p]).strip()
        return str(content).strip()

    def _chat_completion(
        self,
        messages: List[Dict[str, Any]],
        temperature: float = 0.3,
        max_tokens: int = 500,
        model: Optional[str] = None
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """统一聊天调用，带重试"""
        target_model = model or self.model
        last_error: Optional[str] = None

        for attempt in range(self.max_retries + 1):
            try:
                response = self.client.chat.completions.create(
                    model=target_model,
                    messages=messages,
                    temperature=temperature,
                    max_tokens=max_tokens
                )
                result = self._extract_text(response.choices[0].message.content)
                if not result:
                    return False, None, "模型返回为空"
                return True, result, None
            except Exception as e:
                last_error = str(e)
                self.logger.warning(
                    f"API调用失败，第{attempt + 1}/{self.max_retries + 1}次: {last_error}"
                )
                if not self._is_retryable_error(last_error):
                    break
                if attempt < self.max_retries:
                    time.sleep(self._retry_delay_seconds(attempt, last_error))

        return False, None, last_error or "未知错误"

    def _retry_delay_seconds(self, attempt: int, error_text: Optional[str]) -> float:
        """计算重试等待时间（429/过载使用更慢退避）"""
        is_overloaded = self._is_overloaded_error(error_text)
        is_connection = self.is_connection_error(error_text)
        base = 2.0 if is_overloaded else (1.5 if is_connection else 0.8)
        # 1,2,4,8... 指数退避 + 抖动
        delay = base * (2 ** attempt) + random.uniform(0.0, 0.6)
        return min(delay, 18.0)

    def _is_retryable_error(self, error_text: str) -> bool:
        """判断是否值得重试，避免对参数错误反复请求"""
        lowered = (error_text or "").lower()
        non_retry_signals = [
            "invalid_request_error",
            "invalid request",
            "unsupported",
            "401",
            "403",
            "404"
        ]
        return not any(signal in lowered for signal in non_retry_signals)

    def _is_overloaded_error(self, error_text: Optional[str]) -> bool:
        lowered = (error_text or "").lower()
        overload_signals = [
            "429",
            "engine_overloaded_error",
            "overloaded",
            "rate limit",
            "too many requests"
        ]
        return any(signal in lowered for signal in overload_signals)

    def is_connection_error(self, error_text: Optional[str]) -> bool:
        lowered = (error_text or "").lower()
        signals = [
            "connection error",
            "apiconnectionerror",
            "connecterror",
            "connection reset",
            "timeout",
            "timed out",
            "temporarily unavailable",
            "dns",
            "name or service not known",
            "failed to establish a new connection",
            "max retries exceeded",
        ]
        return any(signal in lowered for signal in signals)

    def list_models(self) -> List[str]:
        """获取可用模型列表（若服务端支持）"""
        try:
            response = self.client.models.list()
            return [m.id for m in getattr(response, "data", []) if getattr(m, "id", None)]
        except Exception as e:
            self.logger.info(f"模型列表获取失败: {e}")
            return []

    def generate_text(
        self,
        system_prompt: str,
        user_prompt: str,
        temperature: float = 0.3,
        max_tokens: int = 500,
        model: Optional[str] = None
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """通用文本生成接口"""
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        return self._chat_completion(
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens,
            model=model
        )

    def process_image(
        self,
        image_bytes: bytes,
        system_prompt: str,
        user_prompt: str,
        temperature: float = 0.2,
        max_tokens: int = 400,
        model: Optional[str] = None
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        视觉理解接口（OpenAI兼容的 image_url 格式）
        """
        image_b64 = base64.b64encode(image_bytes).decode("ascii")
        data_url = f"data:image/png;base64,{image_b64}"
        messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": user_prompt},
                    {"type": "image_url", "image_url": {"url": data_url, "detail": "auto"}}
                ]
            }
        ]
        return self._chat_completion(
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens,
            model=model
        )

    def test_connection(self) -> Tuple[bool, str]:
        """测试API连接"""
        success, result, error = self.generate_text(
            system_prompt="你是测试助手。",
            user_prompt="回复：连接正常",
            temperature=0,
            max_tokens=20
        )
        if success:
            return True, f"API连接成功: {result}"
        return False, f"连接失败: {error}"

    def process_cell(
        self,
        cell_content: str,
        system_prompt: str,
        user_instruction: str,
        temperature: float = 0.3,
        max_tokens: int = 500,
        context: Optional[Dict] = None
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """处理单个单元格"""
        full_prompt = f"{user_instruction}\n\n单元格内容:\n{cell_content}"
        if context:
            full_prompt += "\n\n上下文信息:\n"
            for key, value in context.items():
                full_prompt += f"- {key}: {value}\n"

        return self.generate_text(
            system_prompt=system_prompt,
            user_prompt=full_prompt,
            temperature=temperature,
            max_tokens=max_tokens
        )

    def process_batch(
        self,
        cells: List[Dict],
        system_prompt: str,
        user_instruction: str,
        temperature: float = 0.3,
        max_tokens: int = 500,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        delay: float = 0.1
    ) -> List[Dict]:
        """批量处理单元格"""
        results = []
        total = len(cells)

        for idx, cell in enumerate(cells):
            success, result, error = self.process_cell(
                cell_content=cell["content"],
                system_prompt=system_prompt,
                user_instruction=user_instruction,
                temperature=temperature,
                max_tokens=max_tokens,
                context=cell.get("context")
            )
            results.append({
                "row": cell["row"],
                "col": cell["col"],
                "result": result if success else f"错误: {error}",
                "success": success
            })

            if progress_callback:
                progress_callback(idx + 1, total)

            # 遇到过载错误时，提前熔断，避免整批重复失败
            if (not success) and self._is_overloaded_error(error):
                remain = cells[idx + 1 :]
                for tail in remain:
                    results.append({
                        "row": tail["row"],
                        "col": tail["col"],
                        "result": f"错误: {error}",
                        "success": False
                    })
                break

            if idx < total - 1 and delay > 0:
                time.sleep(delay)

        return results

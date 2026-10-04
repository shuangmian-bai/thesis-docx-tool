"""OpenAI 兼容 chat/completions 客户端（仅标准库 urllib）。

DeepSeek / 豆包火山方舟 / 通义千问 / OpenAI 均兼容该协议，故一份实现通用。
设计要点：

- 只做无状态的一问一答 `chat()`，编排（分批、校验）在 correct.py；
- 网络错误 / 429 / 5xx 自动重试（指数退避），4xx 参数错误立即失败；
- 所有对外异常统一为 AIError，信息中不允许出现 api_key（脱敏）；
- 支持可选的 cancel 回调：重试等待间隙检查，已取消则立即中止
  （urllib 单次请求进行中无法强杀，靠较短超时自然结束）。
"""
import json
import socket
import time
import urllib.error
import urllib.request
from typing import Callable, Dict, List, Optional

#: 可重试的网络层异常
_NET_ERRORS = (urllib.error.URLError, socket.timeout, TimeoutError, ConnectionError)
_MAX_RETRY = 2          # 首次 + 2 次重试
_BACKOFF_BASE = 2       # 退避基数秒：2、4


class AIError(Exception):
    """AI 调用失败（配置、网络、HTTP、返回格式）。message 已脱敏。"""


class AIClient:
    """无状态聊天客户端，配置由 docxai.config.load() 提供。"""

    def __init__(self, cfg: Dict[str, object]):
        self.provider = str(cfg.get("provider", "custom"))
        base = str(cfg.get("base_url", "")).rstrip("/")
        if not base:
            raise AIError("配置缺少 base_url")
        # 兼容用户误把完整 chat 地址填进 base_url
        self.url = base if base.endswith("/chat/completions") \
            else base + "/chat/completions"
        self.api_key = str(cfg.get("api_key", ""))
        self.model = str(cfg.get("model", ""))
        try:
            self.timeout = int(cfg.get("timeout", 60))
        except (TypeError, ValueError):
            self.timeout = 60
        if not self.api_key or not self.model:
            raise AIError("配置缺少 api_key 或 model")

    def chat(self, messages: List[Dict[str, str]], *,
             temperature: float = 0.2,
             json_mode: bool = False,
             cancel: Optional[Callable[[], bool]] = None) -> str:
        """发起一次对话，返回助手文本。

        messages: [{"role": "system/user/assistant", "content": "..."}]
        json_mode: 要求模型返回 JSON（各家普遍支持 response_format json_object）。
        cancel:    返回 True 时在重试间隙中止，抛 AIError("已取消")。
        """
        payload = {
            "model": self.model,
            "messages": messages,
            "temperature": temperature,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")

        last_err = None
        for attempt in range(_MAX_RETRY + 1):
            if cancel and cancel():
                raise AIError("已取消")
            try:
                return self._post(data)
            except AIError as e:
                last_err = e
                if not _retryable(e) or attempt == _MAX_RETRY:
                    raise
                wait = _BACKOFF_BASE ** attempt
                # 等待期间也响应取消
                for _ in range(wait * 10):
                    if cancel and cancel():
                        raise AIError("已取消")
                    time.sleep(0.1)
        raise last_err  # 理论不可达

    def _post(self, data: bytes) -> str:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
        }
        req = urllib.request.Request(self.url, data=data, headers=headers,
                                     method="POST")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                body = resp.read().decode("utf-8")
        except urllib.error.HTTPError as e:
            detail = self._http_error_detail(e)
            err = AIError(f"接口返回 HTTP {e.code}：{detail}")
            err.http_code = e.code
            raise err
        except _NET_ERRORS as e:
            # 网络层异常可能带 URL，但不含 key；仍统一措辞，避免泄露内网地址细节
            raise AIError(f"网络连接失败：{type(e).__name__}")

        try:
            obj = json.loads(body)
            return obj["choices"][0]["message"]["content"]
        except (json.JSONDecodeError, KeyError, IndexError, TypeError):
            raise AIError("接口返回内容无法解析（缺少 choices/message/content）")

    @staticmethod
    def _http_error_detail(e: urllib.error.HTTPError) -> str:
        """从错误响应体里提取人类可读信息；任何异常都退化为状态码短语。"""
        try:
            raw = e.read().decode("utf-8", errors="replace")
            obj = json.loads(raw)
            msg = obj.get("error", {}).get("message")
            if isinstance(msg, str) and msg:
                return msg[:300]
        except Exception:
            pass
        return e.reason or "请求被拒绝"


def _retryable(err: AIError) -> bool:
    """429 与 5xx 可重试；401/400 等参数/鉴权错误重试无意义。"""
    code = getattr(err, "http_code", None)
    return code is not None and (code == 429 or code >= 500)

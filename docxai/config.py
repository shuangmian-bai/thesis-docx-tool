"""AI 服务配置：provider 预设与 config/ai.json 读写。

各家大模型（DeepSeek / 火山方舟豆包 / 通义千问 / OpenAI / 自定义）都提供
OpenAI 兼容的 `/chat/completions` 接口，故配置只需 provider、base_url、
api_key、model 四项核心参数 + 超时，客户端一份实现通用。

真实配置写在 config/ai.json（.gitignore 排除，不入库），
config/ai.example.json 是不含 key 的入库示例。
"""
import json
import os
from typing import Dict, Optional

from docxai import AI_CONFIG, AI_CONFIG_EXAMPLE, HERE

#: provider 预设：选 provider 后自动带出 base_url 与默认 model；
#: 豆包/自定义的 model 通常是部署端点 ID，需用户自行填写，故留空占位
PROVIDERS: Dict[str, Dict[str, str]] = {
    "deepseek": {
        "label": "DeepSeek",
        "base_url": "https://api.deepseek.com/v1",
        "model": "deepseek-chat",
    },
    "doubao": {
        "label": "豆包（火山方舟）",
        "base_url": "https://ark.cn-beijing.volces.com/api/v3",
        "model": "",
    },
    "qwen": {
        "label": "通义千问",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen-plus",
    },
    "openai": {
        "label": "OpenAI",
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o-mini",
    },
    "custom": {
        "label": "自定义（OpenAI 兼容接口）",
        "base_url": "",
        "model": "",
    },
}

#: 新建配置时的默认值
DEFAULT_TIMEOUT = 60
REQUIRED_FIELDS = ("base_url", "api_key", "model")


def config_path() -> str:
    """真实配置文件路径。"""
    return AI_CONFIG


def exists() -> bool:
    """config/ai.json 是否存在。"""
    return os.path.exists(AI_CONFIG)


def default_for(provider: str) -> Dict[str, object]:
    """生成某 provider 的一份待填配置（model/base_url 用预设带出）。"""
    p = PROVIDERS.get(provider)
    if p is None:
        raise ValueError(f"未知 provider：{provider}")
    return {
        "provider": provider,
        "base_url": p["base_url"],
        "api_key": "",
        "model": p["model"],
        "timeout": DEFAULT_TIMEOUT,
    }


def load() -> Dict[str, object]:
    """读取配置。文件不存在时给出明确提示（调用方也可先 exists() 判断）。"""
    if not os.path.exists(AI_CONFIG):
        raise ConfigError(
            f"未找到 AI 配置：{os.path.relpath(AI_CONFIG, HERE)}，"
            f"可在 GUI「设置→AI 配置」中填写，或复制 "
            f"{os.path.basename(AI_CONFIG_EXAMPLE)} 为 ai.json 后填写")
    try:
        with open(AI_CONFIG, encoding="utf-8") as fh:
            cfg = json.load(fh)
    except json.JSONDecodeError as e:
        raise ConfigError(f"ai.json 解析失败：{e}")
    if not isinstance(cfg, dict):
        raise ConfigError("ai.json 顶层必须是对象")
    cfg.setdefault("provider", "custom")
    cfg.setdefault("timeout", DEFAULT_TIMEOUT)
    return cfg


def save(cfg: Dict[str, object]) -> None:
    """保存配置到 config/ai.json（覆盖写）。先校验必填项。"""
    problems = validate(cfg)
    if problems:
        raise ConfigError("配置不完整：" + "；".join(problems))
    os.makedirs(os.path.dirname(AI_CONFIG), exist_ok=True)
    with open(AI_CONFIG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def validate(cfg: Dict[str, object]) -> list:
    """返回缺失/不合法字段的中文说明列表；空列表表示通过。"""
    problems = []
    for field in REQUIRED_FIELDS:
        if not str(cfg.get(field, "")).strip():
            names = {"base_url": "接口地址 base_url", "api_key": "API Key",
                     "model": "模型 model"}
            problems.append(f"缺少{names[field]}")
    timeout = cfg.get("timeout", DEFAULT_TIMEOUT)
    try:
        if int(timeout) <= 0:
            raise ValueError
    except (TypeError, ValueError):
        problems.append("timeout 必须是正整数秒")
    return problems


def configured() -> bool:
    """是否已存在一份字段齐全的可用配置（GUI 状态栏/模式开关用）。"""
    try:
        return not validate(load())
    except ConfigError:
        return False


def describe(cfg: Optional[Dict[str, object]] = None) -> str:
    """一行可读描述，如「DeepSeek · deepseek-chat」；未配置时给提示。"""
    if cfg is None:
        if not exists():
            return "AI 未配置"
        try:
            cfg = load()
        except ConfigError:
            return "AI 配置有误"
    provider = str(cfg.get("provider", "custom"))
    label = PROVIDERS.get(provider, {}).get("label", provider)
    model = str(cfg.get("model", "")).strip()
    if not model or not str(cfg.get("api_key", "")).strip():
        return "AI 未配置"
    return f"{label} · {model}"


class ConfigError(Exception):
    """AI 配置缺失或不合法。"""

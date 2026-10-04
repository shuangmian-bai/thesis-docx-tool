"""AI 处理核心包（零 GUI 依赖）。

负责三件事：
- AI 服务配置读写（config/ai.json）与 OpenAI 兼容接口调用；
- 现有块模型（tuple）与可喂给 AI 的 JSON/dict 结构互转、校验；
- 规则解析结果的 AI 结构修正与内容改写编排。

本包只用 Python 标准库，不依赖 PyQt6——GUI（docxgui）与后续终端交互版
都调用本包公开接口，同一套 AI 逻辑只有一份实现。
"""
import os

#: 工具根目录（thesis-docx-tool/）
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: AI 配置文件（个人数据，.gitignore 排除）与入库示例
AI_CONFIG = os.path.join(HERE, "config", "ai.json")
AI_CONFIG_EXAMPLE = os.path.join(HERE, "config", "ai.example.json")

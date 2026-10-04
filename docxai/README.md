# docxai —— AI 结构修正与内容改写核心

无 GUI 依赖的 AI 处理核心，**只用 Python 标准库**（`urllib` / `json`）。
GUI（`docxgui`）与后续的终端交互版都调用本包，同一套 AI 逻辑只此一份实现。

## 定位

规则解析（`docxconvert.parse_docx`）能应对样式规范的 Word，但遇到没套标题样式、
图表题混排的文档会误判块类型。本包不做"纯 AI 重解析"（长文档易丢内容、不可控），
而是 **规则打底 + AI 修正**：规则结果永远先跑，AI 按章只校正结构，校验不过整批保留原文。

## 模块

| 文件 | 职责 |
|---|---|
| `config.py` | `config/ai.json` 读写、provider 预设（DeepSeek/豆包/通义/OpenAI/自定义）、配置校验 |
| `client.py` | OpenAI 兼容 `chat/completions` 客户端：超时、429/5xx 重试退避、错误脱敏、取消回调 |
| `blocks_json.py` | 块 tuple ↔ dict/JSON 互转、schema 校验、AI 返回解析、按 H1 分章 |
| `prompts.py` | 结构修正与内容改写的 system/user 提示词 |
| `correct.py` | 按章分批结构修正（`correct_structure`）与单段改写（`rewrite_text`）编排 |

## 配置

`config/ai.example.json`（入库示例）复制为 `config/ai.json`（不入库）后填写：

```json
{
  "provider": "deepseek",
  "base_url": "https://api.deepseek.com/v1",
  "api_key": "sk-...",
  "model": "deepseek-chat",
  "timeout": 60
}
```

所有 provider 均为 OpenAI 兼容接口，客户端只依赖 base_url / api_key / model；
豆包等用端点 ID 当 model。GUI 也可在「设置 → AI 配置」弹窗里填写。

## 安全护栏（结构修正）

- 按 H1 章分批，发送/收回块数必须一致，否则整批拒绝（防漏块/并块）；
- 每个块过 schema 校验（类型合法、h 的 level 1-4、table 为二维字符串、code 为字符串数组）；
- `img.path` 必须在本次抽取的图片白名单内，AI 无法新增或篡改图片路径；
- 提示词要求 table 单元格文字、img 绑定逐字保留，只允许改块类型与标题级别；
- 被拒批次写入 `CorrectionResult.skipped`（章节+原因），原块保留，不抛断整篇；
- 成功批次的变化块下标在 `changed` 中，由界面标记并支持逐块还原。

## 关键 API

```python
from docxai.config import load, configured
from docxai.client import AIClient
from docxai.correct import correct_structure, rewrite_text

client = AIClient(load())
res = correct_structure(client, blocks, known_images={"images/a.png"},
                        on_progress=lambda done, total, title: None)
# res.blocks / res.changed / res.skipped / res.ok_chapters / res.total_chapters
new_text = rewrite_text(client, "待改文本", "润色为更正式的学术语气")
```

## 依赖与被依赖

- **依赖**：Python 标准库；网络只访问用户配置的 base_url。
- **被谁用**：`docxgui` 的审阅/向导流程；后续终端交互版与 CLI 批处理复用本包。

"""PyQt6 图形界面包（gui 子命令）。

只做 run 的交互式前端：选 docx 与处理模式（默认/AI）→ 审阅扁平化块序列、
手动或 AI 修正 → 写入哈希工作目录章节并由外部进程跑 run 出 Word。

本包是项目内唯一依赖 PyQt6 的模块；核心解析/AI/构建逻辑都在 docxconvert、
docxai、docxflow/docxbuild 中，本包不复制实现。PyQt6 采用延迟导入，
未安装时入口给出安装提示而不是崩溃。
"""

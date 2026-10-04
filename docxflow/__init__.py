"""一键流水线模块（run 子命令）。

把 fig / build / toc / audit 等分步原子命令在内部串成端到端流程，
用户只需一条命令即可从 Markdown（或用户 docx）得到成品论文。
详见 `docxflow/cli.py` 与 `docxflow/README.md`。
"""
import os

#: 工具根目录（thesis-docx-tool/）
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

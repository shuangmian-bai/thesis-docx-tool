"""论文架构图生成模块（内容与样式分离）。

图定义写在 `config/figures/*.md`（markdown 表格 DSL），
样式由 `render.py` 决定。详见 `docxfig/cli.py`。
"""
import os

#: 工具根目录（thesis-docx-tool/）
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

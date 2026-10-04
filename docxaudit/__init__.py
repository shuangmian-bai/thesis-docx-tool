"""成品 docx 格式审计（`audit_docx.py` 的实现）。

`audit_docx.py` 是命令行入口（薄封装），实现按职责分在这几个模块里：

| 模块 | 职责 |
|---|---|
| `common` | 公共层：`Finding` / `Docx` / `Ctx`、路径定位、模块级正则、XML 骨架工具 |
| `expected` | `EXPECTED_DIFFS` 表——登记「有意偏离」，新增一处只改这里 |
| `checks` | 五个检查项，一项一个文件；顺序在 `checks/__init__.py` 里排定 |
| `report` | 结论分级（`classify`）、渲染（`render`）、单项执行（`run_check`） |
| `cli` | 命令行解析与总览输出（`main()`） |

差异**只在报告层降级**：检查函数一律报 `WARN`，命中 `EXPECTED_DIFFS` 才在
`report.classify()` 里改判为「有意偏离」。因此新增一处有意偏离时只改表，
不必回头动检查逻辑。
"""
import os

#: 工具根目录（`thesis-docx-tool/`）——本包在其下一层，故上跳一级
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

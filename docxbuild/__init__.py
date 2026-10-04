"""论文 docx 构建链（`build_docx.py` 的实现）。

`build_docx.py` 是命令行入口（薄封装），实现按职责分在这几个模块里：

| 模块 | 职责 |
|---|---|
| `docinfo` | 论文著录信息：题目、封面字段、版本号——换论文或补填封面只动这里 |
| `layout` | 版式常量：正文区宽度与插图尺寸上限 |
| `mdparse` | Markdown → 块序列（纯文本，不碰 XML） |
| `fragments` | 块 → OOXML 片段（段落、图、表、代码、参考文献） |
| `template` | 模板骨架处理（命名空间、封面、目录、文档属性、无用部件清理） |
| `toc_pages` | 从渲染出的 PDF 反查各级标题页码，写入 `output/toc_pages.json` |
| `cli` | 主流程编排（`main()`） |

依赖是单向的：`docinfo` / `layout` / `mdparse` 不依赖别人，`fragments` 用 `layout`
与 `mdparse`，`template` 用 `fragments`，`cli` 汇总全部。
"""
import os

#: 工具根目录（`thesis-docx-tool/`）——本包在其下一层，故上跳一级
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: wordprocessingml 主命名空间（Clark 记法），解析模板 XML 树时用
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

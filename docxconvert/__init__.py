"""docx 拆解转换链（`convert_docx.py` 的实现）。

把一篇用户撰写的 Word 论文（docx）拆解成与 `章节/*.md` 同格式的 Markdown，
使其可被 `build_docx.py` 重新套回模板生成成品。这是 `docxbuild` 的**逆向**：
`docxbuild` 做 Markdown → docx，本包做 docx → Markdown。

| 模块 | 职责 |
|---|---|
| `parse` | 解析 docx 的 `word/document.xml`，按 body 子元素顺序把段落/表格/图片转成块序列 |
| `extract` | 从 docx 里把图片抽取到目标目录 |
| `markdown` | 块序列渲染成 `章节/*.md` 格式的 Markdown 文本 |
| `cli` | 主流程编排（`main()`） |

块格式与 `docxbuild.mdparse` 完全对齐，保证可双向闭环：

| 标签 | 内容 | 源写法 |
|---|---|---|
| `h` | `(级别, 标题文本)` | `# …` ~ `#### …` |
| `p` | 段落文本 | 连续非空行 |
| `caption` | 表题 / 图题 | `[表4-1　说明]` |
| `ref` | 参考文献条目 | `[1] 作者. 题名…` |
| `code` | 代码行列表 | ``` 围栏块 |
| `table` | 二维单元格列表 | `| a | b |` |
| `img` | `(图片相对路径, 图题说明)` | `![说明](路径)` |

依赖只用标准库（`zipfile` / `xml.etree` / `re` / `os`），不用 python-docx、不用 pandoc，
与 `docxbuild` 保持一致。
"""
import os

#: 工具根目录（`thesis-docx-tool/`）——本包在其下一层，故上跳一级
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: wordprocessingml 主命名空间（Clark 记法）
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: 关系命名空间
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

#: drawingml 命名空间（图片 blip）
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

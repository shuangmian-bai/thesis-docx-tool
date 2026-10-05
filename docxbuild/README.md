# docxbuild — Markdown → docx 构建链

| 项目 | 说明 |
| --- | --- |
| **分层定位** | 构建层 · 正向链路（Markdown → docx） |
| **开发状态** | [通过] |
| **职责** | 把 `章节/*.md` 套用学校论文模板骨架，生成符合格式的 Word 论文 |

## 用途

内部构建模块（由 `run` 编排调用）。输入是按章拆分的 Markdown（`章节/*.md`）和学校模板
（`template.docx`），输出是 `论文_v1_YYYYMMDD.docx`。

核心思路：**保留模板骨架**（封面、承诺书、目录域、页眉页脚、styles/numbering/theme），
只替换中间的示例正文。因 `mc:Ignorable` 前缀问题，不能用 ElementTree 整体序列化，
采用「字符串拼接 + 逐片段序列化」。

## 模块组成

| 文件 | 说明 |
| --- | --- |
| `docinfo.py` | 论文著录信息：题目、版本号（默认占位符；封面由模板骨架提供，不提取不填充） |
| `layout.py` | 版式常量：正文区宽度、插图尺寸上限 |
| `mdparse.py` | Markdown → 块序列（纯文本，不碰 XML） |
| `fragments.py` | 块 → OOXML 片段（段落、标题、有序列表项、图、表、代码、参考文献、图题） |
| `template.py` | 模板骨架处理：命名空间登记、目录重建、样式清理、列表编号注入（封面/承诺书保留模板原样） |
| `tplinspect.py` | 模板只读检视（骨架/封面/承诺书/TOC/样式/编号/页面设置）；无 Qt 依赖 |
| `toc_pages.py` | 从渲染出的 PDF 反查各级标题页码，写入工作目录 toc.json |
| `cli.py` | 主流程编排（`main(argv)`） |

## 块模型（与 docxconvert 共用）

`mdparse.parse_md()` 与 `docxconvert.parse.parse_docx()` 共用同一套块格式，
保证 Markdown ↔ docx 可逆闭环：

| 标签 | 内容 | Markdown 写法 |
|---|---|---|
| `h` | `(级别, 标题文本)` | `#` ~ `####` |
| `p` | 段落文本 | 连续非空行 |
| `ol` | 有序列表的单个条目（连续多个归同一自动编号列表，构建时每段独立从 1 编号） | `1. 文本` |
| `caption` | 表题 / 图题 | `[表4-1　说明]` |
| `ref` | 参考文献条目 | `[1] 作者. 题名…` |
| `code` | 代码行列表 | ``` 围栏块 |
| `table` | 二维单元格列表 | `| a | b |` 管道表格 |
| `img` | `(图片相对路径, 图题说明, cx, cy)` | `![说明](路径){cx=N,cy=N}` |

## 关键 API 与用法

```python
from docxbuild.cli import main
main(["--template", "template.docx"])   # 模块级入口（开发调试用，非面向用户命令）
```

封面与承诺书（含签名）保留模板骨架原样，不提取、不填充、不替换；目录页码由
工作目录的 `toc.json` 提供（由 `toc_pages.py` 从 PDF 量出），缺失时目录页码
留空，在 Word 中按 F9 可重建域结果。

## 依赖与被依赖

- **依赖**：Python 标准库（`zipfile` / `xml.etree` / `re` / `json` / `os` / `datetime`）；
  插图尺寸计算可选 `Pillow`，无图时可缺省。**不用 python-docx、不用 pandoc**。
- **被谁用**：`docxflow` 一键流水线（run 编排调用 `build()` 与 `toc_pages`）。

## 注意点

- 依赖方向单向：`docinfo`/`layout`/`mdparse` 不依赖别人；`fragments` 用 `layout`+`mdparse`；
  `template` 用 `fragments`；`cli` 汇总全部。
- `HERE` 指向工具根目录（`thesis-docx-tool/`），本包在其下一层，故上跳一级。
- 新增块类型需同步改三处：`mdparse.py`（解析）、`fragments.py`（渲染）、
  `docxconvert/parse.py` + `markdown.py`（逆向），保持闭环。
- 有序列表（`ol`）约定：连续的 `ol` 块算一段列表，构建时在 `numbering.xml`
  追加一个新的 `w:num`（复用模板里 `numFmt=decimal` 的 abstractNum，从 1 开始），
  段落只写 pStyle 与 `numPr`，不写编号文字；不同段（中间隔着非 ol 块，含章节
  间空段）各自独立从 1 编号。新 `w:num` 必须排在全部 abstractNum 之后，故统一
  插在 `</w:numbering>` 前。模板没有 decimal 编号定义时该段降级为普通段落
  （统计输出 `有序列表降级 N 段` 警告），不报错。
- `tplinspect.py` 只读模板 zip，不修改模板本体；骨架与样式问题只报告，
  由用户在 Word 中改模板后重新检测复验。
- 章节源过滤：`build()` 只把 `chap_dir`（哈希工作目录的 `章节/`）下**文件名
  以数字开头**的 `.md` 当作章节处理（如 `01_设计思路.md`），其余文件跳过并打印
  警告，防止把模板示例、整篇拆解产物等非章节文件混进来导致正文重复构建。目录里
  全是非数字开头文件时（如 GUI 审阅页导出的单篇论文 md）退化为全量处理。

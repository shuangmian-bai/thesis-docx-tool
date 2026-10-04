# docxconvert — docx → Markdown 拆解链

| 项目 | 说明 |
| --- | --- |
| **分层定位** | 构建层 · 逆向链路（docx → Markdown） |
| **开发状态** | [OK] 已完成 |
| **职责** | 把用户已写好的 Word 论文拆解为与 `章节/*.md` 同格式的 Markdown |

## 用途

`convert_docx.py` 的实现包。输入是一篇 Word 论文（`用户论文.docx`），输出是
Markdown 文件 + 抽取出的图片目录。产出的 Markdown 可直接喂给 `docxbuild`
重新套回模板，形成「撰写 → 拆解 → 二次编辑 → 重建」闭环。

这是 `docxbuild` 的**逆向**：`docxbuild` 做 Markdown → docx，本包做 docx → Markdown。

## 模块组成

| 文件 | 说明 |
| --- | --- |
| `parse.py` | 解析 docx 的 `word/document.xml`，按 body 子元素顺序把段落/表格/图片转成块序列 |
| `extract.py` | 从 docx 的 `word/media/` 里把图片抽取到目标目录 |
| `markdown.py` | 块序列渲染为 `章节/*.md` 格式的 Markdown 文本 |
| `cli.py` | 主流程编排：核心函数 `convert_to()`，命令行 `main(argv)` 只负责参数与默认落点 |

## 解析要点

- 按 `<w:body>` 子元素顺序**混合遍历** `<w:p>` 与 `<w:tbl>`，避免分别遍历导致顺序错乱。
- 标题识别优先用段落样式 ID（`1`~`4` / `Heading 1` 等），无样式时降级看 `<w:outlineLvl>`。
- 有序列表只认真自动编号：读 `pPr/numPr/numId/ilvl`，查 `numbering.xml` 得
  该 numId 的 0 级 `numFmt`；`numFmt=decimal` 且 numId 非 0 才判为 `ol`，
  且该判定优先于标题样式。`numId=0`（取消编号）、bullet 等非 decimal 格式
  一律按正文处理。不看文本里的「1.」前缀——手写伪编号留给 GUI 人工或 AI 改判。
  渲染为 md 时连续 `ol` 块按段重新输出 `1. 2. 3.`，不保留 Word 实际编号。
- 图题（紧跟图片的「图x.y …」）并入图片块的说明字段；表题单独成 `caption` 块。
- 目录条目（样式名 `toc*`）自动跳过。
- `--strip-front`：丢弃第一个以数字开头的一级标题之前的内容（封面、承诺书、目录）。

块格式与 `docxbuild.mdparse` 完全对齐，保证可双向闭环。

## 关键 API 与用法

```python
from docxconvert.cli import main, convert_to
# 命令行入口（产物默认落 output/）
main(["用户论文.docx", "--strip-front"])   # 等价于 python3 main.py convert 用户论文.docx --strip-front
# 指定落点的核心函数（docxflow 用它把拆解产物直接落 config/章节/ 与 config/images/）
# rel_base 必须传 config/：引用写成 images/...，与 build 以 config/ 为基准的解析对齐
convert_to("用户论文.docx", "config/章节/用户论文.md",
           "config/images/用户论文_images", strip_front=True,
           rel_base="config")
```

## 依赖与被依赖

- **依赖**：Python 标准库（`zipfile` / `xml.etree` / `re` / `os` / `shutil`）。
  **不用 python-docx、不用 pandoc**。
- **被谁用**：`main.py` 的 `convert` 子命令；`docxflow` 一键流水线（Word 路线）
  直接调用 `convert_to()`，不重复实现拆解逻辑。

## 注意点

- `HERE` 指向工具根目录（`thesis-docx-tool/`），本包在其下一层，故上跳一级。
- 用户 docx 的标题样式可能不是模板的 `1`/`2`/`3`/`4`，工具会尝试匹配 `Heading 1` 等
  常见样式名；匹配不到的当正文处理，可手工调整 MD。
- 新增块类型需同步改 `docxbuild` 的 `mdparse.py` / `fragments.py`，保持闭环。
- **段落内软换行 `<w:br>`**：转成空格（Markdown 无段落内换行概念，避免段落边界
  被拆散破坏闭环）。
- **空格归一化**：非代码段落的连续空格合并为单空格；代码块（样式 `code`）通过
  `preserve_ws=True` 保留原始缩进。
- **表格单元格多段**：用 `\n` 连接，`docxconvert.markdown` 渲染为 `<br>`，
  `docxbuild.fragments.table_xml` 拆成多个 `<w:p>` 还原。

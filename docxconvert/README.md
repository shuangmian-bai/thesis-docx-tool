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
convert_to("用户论文.docx", "config/章节/用户论文.md",
           "config/images/用户论文_images", strip_front=True)
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

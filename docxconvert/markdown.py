"""块序列 → 章节/*.md 格式的 Markdown 文本。

渲染规则与 `docxbuild.mdparse` 的解析规则一一对应，保证可逆：

- `h` → `#` 数量 = 级别，后跟标题文本
- `p` → 正文段落
- `caption` → `[文本]`（表题；图题已在 parse 阶段并入图片块）
- `ref` → 原样输出 `[1] …`
- `code` → ``` 围栏块
- `table` → 管道表格，首行后加 `| --- |` 分隔行
- `img` → `![说明](路径)`
"""


def _cell(s):
    """转义表格单元格：竖线与换行。"""
    return s.replace("|", "\\|").replace("\n", "<br>")


def render(blocks):
    out = []
    for blk in blocks:
        kind = blk[0]
        if kind == "h":
            out.append(f"{'#' * blk[1]} {blk[2]}")
            out.append("")
        elif kind == "p":
            out.append(blk[1])
            out.append("")
        elif kind == "caption":
            out.append(f"[{blk[1]}]")
            out.append("")
        elif kind == "ref":
            out.append(blk[1])
            out.append("")
        elif kind == "code":
            out.append("```")
            out.extend(blk[1])
            out.append("```")
            out.append("")
        elif kind == "table":
            rows = blk[1]
            if not rows:
                continue
            ncol = max(len(r) for r in rows)
            rows = [r + [""] * (ncol - len(r)) for r in rows]
            out.append("| " + " | ".join(_cell(c) for c in rows[0]) + " |")
            out.append("| " + " | ".join("---" for _ in range(ncol)) + " |")
            for r in rows[1:]:
                out.append("| " + " | ".join(_cell(c) for c in r) + " |")
            out.append("")
        elif kind == "img":
            out.append(f"![{blk[2]}]({blk[1]})")
            out.append("")
    return "\n".join(out).rstrip() + "\n"

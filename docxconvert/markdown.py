"""块序列 → 章节/*.md 格式的 Markdown 文本。

渲染规则与 `docxbuild.mdparse` 的解析规则一一对应，保证可逆：

- `h` → `#` 数量 = 级别，后跟标题文本
- `p` → 正文段落
- `caption` → `[文本]`（表题；图题已在 parse 阶段并入图片块）
- `ref` → 原样输出 `[1] …`
- `ol` → `1. 文本`；连续条目为同一列表，编号在每个连续段重新从 1 计数
- `code` → ``` 围栏块
- `table` → 管道表格，首行后加 `| --- |` 分隔行
- `img` → `![说明](路径)`
"""
import re


def _cell(s):
    """转义表格单元格：竖线与换行。"""
    return s.replace("|", "\\|").replace("\n", "<br>")


def render(blocks):
    out = []
    ol_n = 0          # 当前连续有序列表的序号（遇到非 ol 块归零）
    for blk in blocks:
        kind = blk[0]
        if kind == "ol":
            ol_n += 1
            out.append(f"{ol_n}. {blk[1]}")
            continue
        if ol_n:
            out.append("")          # 列表段结束，与后续块空一行分隔
            ol_n = 0
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
            attr = f"{{sz={blk[2]}}}" if len(blk) > 2 and blk[2] else ""
            out.append(f"```{attr}")
            out.extend(blk[1])
            out.append("```")
            out.append("")
        elif kind == "table":
            rows = blk[1]
            if not rows:
                continue
            # 跨列合并（{span=N} 前缀）占 N 列：ncol 按展开列数计，不再补齐空单元格
            def span_of(c):
                m = re.match(r"^\{span=(\d+)\}", c)
                return int(m.group(1)) if m else 1

            ncol = max(sum(span_of(c) for c in r) for r in rows)
            out.append("| " + " | ".join(_cell(c) for c in rows[0]) + " |")
            out.append("| " + " | ".join("---" for _ in range(ncol)) + " |")
            for r in rows[1:]:
                out.append("| " + " | ".join(_cell(c) for c in r) + " |")
            out.append("")
        elif kind == "img":
            cx, cy = blk[3], blk[4]
            size = f"{{cx={cx},cy={cy}}}" if cx and cy else ""
            out.append(f"![{blk[2]}]({blk[1]}){size}")
            out.append("")
    return "\n".join(out).rstrip() + "\n"

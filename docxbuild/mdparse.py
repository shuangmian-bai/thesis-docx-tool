"""Markdown → 块序列。纯文本解析，不产生任何 XML。

块是元组，首元素是类型标签，`cli.main()` 的正文循环按标签分派到
`fragments` 里对应的生成函数：

| 标签 | 块内容 | 源写法 |
|---|---|---|
| `h` | `(级别, 标题文本)` | `# …` ~ `#### …` |
| `p` | 段落文本 | 连续非空行合并 |
| `quote` | 引用块文本 | `> …` |
| `img` | `(图片路径, 图题说明)` | `![说明](路径)` |
| `caption` | 表题 / 图题 | `[表4-1　说明]` |
| `ref` | 参考文献条目 | `[1] 作者. 题名…` |
| `code` | 代码行列表 | ``` 围栏块 |
| `table` | 二维单元格列表 | `| a | b |` |
"""
import re
import unicodedata


def disp_w(s):
    """估算显示宽度：全角字符按 2 计，半角按 1 计。

    按字符数分配列宽会让中文列偏窄、纯拉丁列偏宽——一个汉字的宽度约等于两个字母。
    """
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in s)


def parse_inline(text):
    """把一行文本切成 (内容, 是否加粗, 是否行内代码) 的列表。"""
    out, i = [], 0
    pattern = re.compile(r"\*\*(.+?)\*\*|`([^`]+)`")
    for m in pattern.finditer(text):
        if m.start() > i:
            out.append((text[i:m.start()], False, False))
        if m.group(1) is not None:
            out.append((m.group(1), True, False))
        else:
            out.append((m.group(2), False, True))
        i = m.end()
    if i < len(text):
        out.append((text[i:], False, False))
    return [r for r in out if r[0]] or [("", False, False)]


def parse_md(text):
    """把 Markdown 文本解析为块序列。"""
    lines = text.split("\n")
    blocks, i = [], 0
    while i < len(lines):
        ln = lines[i]
        s = ln.strip()

        if not s:
            i += 1
            continue

        # 代码块
        if s.startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            blocks.append(("code", buf))
            continue

        # 各级标题
        m = re.match(r"^(#{1,4})\s+(.*)$", s)
        if m:
            blocks.append(("h", len(m.group(1)), m.group(2).strip()))
            i += 1
            continue

        # 图片：![说明](路径)
        m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)$", s)
        if m:
            blocks.append(("img", m.group(2).strip(), m.group(1).strip()))
            i += 1
            continue

        # 表题 / 图题： [表4-1　说明]
        m = re.match(r"^\[((?:表|图)\s*[\d\-–—.]+[^\]]*)\]$", s)
        if m:
            blocks.append(("caption", m.group(1).strip()))
            i += 1
            continue

        # 参考文献条目： [1] 作者. 题名[文献类型]. …
        if re.match(r"^\[\d+\]\s", s):
            blocks.append(("ref", s))
            i += 1
            continue

        # 引用块（用作「说明」提示框）
        if s.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append(("quote", " ".join(buf)))
            continue

        # 表格
        if s.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                raw = lines[i].strip()
                if not re.match(r"^\|[\s:\-|]+\|$", raw):
                    cells = [c.strip() for c in raw.strip("|").split("|")]
                    rows.append(cells)
                i += 1
            if rows:
                blocks.append(("table", rows))
            continue

        # 普通段落：把连续非空行合并为一段
        buf = [s]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith(("#", "|", ">", "```", "!["))
                    or re.match(r"^\[(?:表|图)", nxt)):
                break
            buf.append(nxt)
            i += 1
        blocks.append(("p", " ".join(buf)))
    return blocks

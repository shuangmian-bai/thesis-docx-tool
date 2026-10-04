"""架构图 DSL 解析器（内容层）。

读取 markdown 表格形式的图定义，返回元素列表交给 render 绘制。

DSL 格式（每个图一个 .md 文件）：

    # fig3_1: 注入攻击系统总体架构

    canvas: 1350 1280

    | type | pos | size | color | text | extra |
    |---|---|---|---|---|---|
    | container | 40,100 | 1270,516 | blue | 推送层（PC 端） | |
    | box | 60,176 | 390,120 | blue | 数据源层\\n解复用·提取配置 | |
    | arrow | 675,616 | 675,666 | | | label:VCU1/TCP |

字段说明：
- type: container | box | diamond | arrow | line | text
- pos: x,y（设计单位）
- size: w,h（box/container/diamond）或 x2,y2（arrow/line）
- color: 语义色名（blue/orange/green/grey/purple/red/white），留空用默认
- text: 显示文本，\\n 表示换行
- extra: 可选属性，用 key:value 逗号分隔，如 label:xxx, dashed:true, radius:0, font:tiny
"""
import os
import re


def _parse_pair(s):
    """'40,100' → (40.0, 100.0)。"""
    a, b = s.split(",")
    return float(a.strip()), float(b.strip())


def _parse_extra(s):
    """'label:xxx, dashed:true' → {'label': 'xxx', 'dashed': 'true'}。"""
    out = {}
    if not s or not s.strip():
        return out
    for part in s.split(","):
        part = part.strip()
        if not part or ":" not in part:
            continue
        k, v = part.split(":", 1)
        out[k.strip()] = v.strip()
    return out


def parse(md_path):
    """解析图定义 md 文件，返回 (canvas_h, elements)。"""
    with open(md_path, encoding="utf-8") as fh:
        lines = fh.readlines()

    canvas_h = 1000
    elements = []
    in_table = False

    for ln in lines:
        ln = ln.rstrip("\n")
        # canvas: W H
        m = re.match(r"canvas:\s*(\d+)\s+(\d+)", ln)
        if m:
            canvas_h = int(m.group(2))
            continue
        # 表格行：| a | b | c | d | e | f |
        if ln.startswith("|") and ln.count("|") >= 6:
            cells = [c.strip() for c in ln.strip().strip("|").split("|")]
            # 跳过表头和分隔行
            if cells[0].lower() == "type" or set(cells[0]) <= {"-", ":"}:
                in_table = True
                continue
            if not in_table or len(cells) < 5:
                continue
            typ, pos_s, size_s, color_s, text_s = cells[0], cells[1], cells[2], cells[3], cells[4]
            extra_s = cells[5] if len(cells) > 5 else ""
            text_s = text_s.replace("\\n", "\n")
            el = {
                "type": typ.strip().lower(),
                "pos": _parse_pair(pos_s),
                "size": _parse_pair(size_s) if size_s else (0, 0),
                "color": color_s.strip() or None,
                "text": text_s,
                "extra": _parse_extra(extra_s),
            }
            elements.append(el)

    return canvas_h, elements


def list_figures(fig_dir):
    """列出图目录下所有 .md 定义文件，返回 [(name, path)]。"""
    out = []
    if not os.path.isdir(fig_dir):
        return out
    for fn in sorted(os.listdir(fig_dir)):
        if fn.endswith(".md"):
            name = fn[:-3]
            out.append((name, os.path.join(fig_dir, fn)))
    return out

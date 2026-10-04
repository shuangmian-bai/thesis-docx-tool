"""审阅树的纯数据辅助：块类型中文标签、摘要、可编辑性判定。

不依赖 PyQt6，便于与 docxai 块模型保持单一事实来源；
QTreeWidget 的填充与编辑控件在 review_page.py。
"""
from typing import Tuple

#: 块类型 -> 审阅界面显示的中文标签
KIND_LABELS = {
    "h1": "标题1", "h2": "标题2", "h3": "标题3", "h4": "标题4",
    "p": "正文",
    "quote": "引用",
    "caption": "图题/表题",
    "ref": "文献",
    "code": "代码",
    "table": "表格",
    "img": "图片",
}

#: 类型筛选下拉/复选框的展示顺序
FILTER_KINDS = ("h1", "h2", "h3", "h4", "p", "quote", "caption",
                "ref", "code", "table", "img")

#: 可做 AI/手动文本改写的块类型
TEXT_EDITABLE = ("h", "p", "quote", "caption", "ref")


def kind_key(block: Tuple) -> str:
    """把块映射到 KIND_LABELS 的键（h 按级别拆成 h1-h4）。"""
    return f"h{block[1]}" if block[0] == "h" else block[0]


def label(block: Tuple) -> str:
    """块的中文类型标签，如「标题2」「表格」。"""
    return KIND_LABELS[kind_key(block)]


def summary(block: Tuple, width: int = 38) -> str:
    """树项的一行纯文本摘要（不渲染任何格式）。"""
    kind = block[0]
    if kind == "h":
        text = block[2]
    elif kind in ("p", "quote", "caption", "ref"):
        text = block[1]
    elif kind == "code":
        text = f"{len(block[1])} 行代码：" + " / ".join(block[1][:2])
    elif kind == "table":
        rows = block[1]
        text = f"{len(rows)} 行 x {max(len(r) for r in rows)} 列：" \
               + " | ".join(rows[0][:4])
    else:  # img
        text = f"{block[2]}  ({block[1]})"
    text = " ".join(str(text).split())
    return text if len(text) <= width else text[:width] + "..."


def chapter_title(block: Tuple) -> str:
    """H1 块的章标题文本，非 H1 返回空串。"""
    return block[2] if block[0] == "h" and block[1] == 1 else ""

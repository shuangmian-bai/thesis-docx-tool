"""docxfig：论文架构图生成模块。

**内容与样式分离**：
- 内容（画什么、在哪画）写在 `config/figures/*.md` 的 markdown 表格里；
- 样式（配色、字体、圆角、箭头）由本模块的 render.py 决定，改样式只动代码。

用法：

    python3 main.py fig              # 生成全部图
    python3 main.py fig fig3_1       # 只生成指定图
"""
import os

from docxfig import HERE
from docxfig.parser import parse, list_figures
from docxfig import render as R

# 图定义目录与产物目录
FIG_DIR = os.path.join(HERE, "config", "figures")
OUT_DIR = os.path.join(HERE, "config", "images")

# 字号映射
FONT_MAP = {"box": R.F_BOX, "small": R.F_SMALL, "tiny": R.F_TINY, "head": R.F_HEAD}


def _draw_element(d, el):
    """根据元素类型分派绘制。"""
    t = el["type"]
    x, y = el["pos"]
    w, h = el["size"]
    color = el["color"]
    text = el["text"]
    ex = el["extra"]

    font = FONT_MAP.get(ex.get("font", ""), None)
    radius = int(ex["radius"]) if "radius" in ex else None
    # fill/line 覆盖：值为语义色名，解析为对应十六进制
    fill_override = R._color(ex["fill"])[0] if "fill" in ex else None
    line_override = R._color(ex["line"])[1] if "line" in ex else None

    if t == "container":
        R.container(d, x, y, w, h, text, color or "blue")
    elif t == "box":
        kw = {}
        if font:
            kw["font"] = font
        if radius is not None:
            kw["radius"] = radius
        if fill_override:
            kw["fill"] = fill_override
        if line_override:
            kw["line"] = line_override
        R.box(d, x, y, w, h, text, color or "blue", **kw)
    elif t == "diamond":
        R.diamond(d, x, y, w, h, text, color or "orange", font=font)
    elif t == "arrow":
        aw = int(ex["width"]) if "width" in ex else 4
        ah = int(ex["head"]) if "head" in ex else 16
        R.arrow(d, x, y, w, h, label=ex.get("label") or None,
                color=color, width=aw, head=ah,
                dashed=ex.get("dashed", "").lower() == "true")
    elif t == "line":
        R.line(d, x, y, w, h, color=color)
    elif t == "text":
        R.text(d, x, y, text, font=font, color=color)


def generate(name=None):
    """生成图。name=None 生成全部，否则只生成指定图。"""
    figures = list_figures(FIG_DIR)
    if name:
        figures = [(n, p) for n, p in figures if n == name]
        if not figures:
            raise SystemExit(f"找不到图定义：{name}（在 {FIG_DIR} 下）")

    print(f"生成架构图 → {os.path.relpath(OUT_DIR, HERE)}")
    for fig_name, md_path in figures:
        canvas_h, elements = parse(md_path)
        im, d = R.new_canvas(canvas_h)
        for el in elements:
            _draw_element(d, el)
        R.save(im, OUT_DIR, f"{fig_name}.png")
    print("完成。")


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="从 markdown 定义生成论文架构图。")
    ap.add_argument("name", nargs="?", default=None,
                    help="图名（如 fig3_1），不填则生成全部")
    args = ap.parse_args(argv)
    generate(args.name)

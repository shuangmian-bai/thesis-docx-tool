"""架构图绘制引擎（样式层）。

本模块负责**怎么画**：配色、字体、画布、圆角矩形、菱形、箭头、连线等
所有视觉样式都由代码决定，调用方只描述**画什么**（位置、尺寸、文本、颜色名）。

设计原则：
- 样式常量（颜色、字号、画布宽、缩放比）集中定义，改样式只动这里；
- 坐标、尺寸用**设计单位**传入，渲染时统一乘 SCALE，保证改清晰度不动版面；
- 颜色用语义名（blue/orange/green/grey/purple/red），不直接传十六进制。

依赖：Pillow（可选，无 Pillow 时无法出图）；系统 Noto Sans CJK 字体。
"""
import math
import os

from PIL import Image, ImageChops, ImageDraw, ImageFont

#: 画布宽度（设计单位）。由打印字号倒推：图被缩到版心宽 15.92cm，
#: 要让最小字号 F_TINY(24) 达到纸面 8pt，内容宽须 ≤ 24÷8×451.3 ≈ 1354。
CANVAS_W = 1350

#: 渲染放大倍数，只决定打印清晰度（像素密度），不影响版面。
SCALE = 2

#: 裁剪后四周保留的外边距（设计单位）。
MARGIN = 36

# 字体：Noto Sans CJK，SC 子体（index=2），中文论文必须取 SC 而非 JP。
FONT_PATH = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
FONT_INDEX = 2

# ── 配色（语义名 → (填充色, 边框色)）──
PALETTE = {
    "blue":   ("#EAF2FB", "#3C6FA8"),
    "orange": ("#FDF2E4", "#C97B2C"),
    "green":  ("#EAF6EE", "#3F8F55"),
    "grey":   ("#F3F3F3", "#8A8A8A"),
    "purple": ("#F1EDFA", "#6A55A8"),
    "red":    ("#FCEDED", "#B4453F"),
    "white":  ("#FFFFFF", "#3C6FA8"),
}
# 容器用更浅的底（与内部 box 区分），按语义色取近白 tint
CONTAINER_FILL = {
    "blue":   "#FAFCFF",
    "orange": "#FEFBF7",
    "green":  "#FAFDFB",
    "grey":   "#F8F8F8",
    "purple": "#FBFAFE",
    "red":    "#FEFAFA",
    "white":  "#FFFFFF",
}
C_INK = "#1A1A1A"
C_LINE = "#5A5A5A"


def _color(name):
    """语义色名 → (fill, line)；未命中时退回 blue。"""
    return PALETTE.get(name, PALETTE["blue"])


class ScaledFont:
    """字体句柄：size 是设计字号（布局用它算行高），渲染时已 ×SCALE。"""

    __slots__ = ("size", "real")

    def __init__(self, size):
        self.size = size
        self.real = ImageFont.truetype(FONT_PATH, int(size * SCALE), index=FONT_INDEX)


# 预设字号
F_BOX = ScaledFont(34)
F_SMALL = ScaledFont(28)
F_TINY = ScaledFont(24)
F_HEAD = ScaledFont(36)


class Canvas:
    """按 SCALE 放大绘制的画布：坐标、线宽、半径、字号都用设计单位传入。"""

    def __init__(self, im):
        self.im = im
        self._d = ImageDraw.Draw(im)

    @staticmethod
    def _font(f):
        return f.real if isinstance(f, ScaledFont) else f

    @staticmethod
    def _flat(seq):
        return [v * SCALE for v in seq]

    @staticmethod
    def _pts(seq):
        return [(x * SCALE, y * SCALE) for x, y in seq]

    def textlength(self, s, font=None):
        return self._d.textlength(s, font=self._font(font)) / SCALE

    def text(self, xy, s, font=None, **kw):
        self._d.text((xy[0] * SCALE, xy[1] * SCALE), s, font=self._font(font), **kw)

    def line(self, xy, fill=None, width=1, **kw):
        self._d.line(self._flat(xy), fill=fill, width=max(1, round(width * SCALE)), **kw)

    def polygon(self, xy, fill=None, outline=None, width=1, **kw):
        if outline is not None and width is not None:
            width = max(1, round(width * SCALE))
        self._d.polygon(self._pts(xy), fill=fill, outline=outline, width=width, **kw)

    def rounded_rectangle(self, xy, radius=0, fill=None, outline=None, width=1):
        self._d.rounded_rectangle(self._flat(xy), radius=round(radius * SCALE),
                                  fill=fill, outline=outline,
                                  width=max(1, round(width * SCALE)))


def new_canvas(h):
    im = Image.new("RGB", (CANVAS_W * SCALE, h * SCALE), "white")
    return im, Canvas(im)


def text_w(d, s, f):
    """文本设计宽度；多行时取最长行。"""
    if "\n" in s:
        return max(d.textlength(ln, font=f) for ln in s.split("\n"))
    return d.textlength(s, font=f)


def wrap(d, s, f, max_w):
    """按像素宽度折行：中文逐字折，英文按空格折。"""
    lines, cur = [], ""
    for ch in s:
        if ch == "\n":
            lines.append(cur)
            cur = ""
            continue
        if text_w(d, cur + ch, f) > max_w and cur:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines


def box(d, x, y, w, h, text, color="blue", font=None, radius=14, width=3, pad=18,
        fill=None, line=None):
    """圆角矩形，文本居中自动折行。color 为语义色名；fill/line 可单独覆盖。"""
    base_fill, base_line = _color(color)
    f = font or F_BOX
    use_fill = fill or base_fill
    use_line = line or base_line
    d.rounded_rectangle([x, y, x + w, y + h], radius=radius, fill=use_fill,
                        outline=use_line, width=width)
    lines = wrap(d, text, f, w - 2 * pad)
    lh = f.size + 12
    total = lh * len(lines)
    ty = y + (h - total) // 2
    for ln in lines:
        tw = text_w(d, ln, f)
        d.text((x + (w - tw) // 2, ty), ln, font=f, fill=C_INK)
        ty += lh


def diamond(d, x, y, w, h, text, color="orange", font=None, width=3):
    """判断菱形：x,y 是外包矩形左上角，w,h 是外包宽高。"""
    fill, line = _color(color)
    f = font or F_SMALL
    cx, cy = x + w / 2.0, y + h / 2.0
    d.polygon(
        [(cx, cy - h / 2), (cx + w / 2, cy), (cx, cy + h / 2), (cx - w / 2, cy)],
        fill=fill, outline=line, width=width,
    )
    lines = wrap(d, text, f, w * 0.6)
    lh = f.size + 12
    total = lh * len(lines)
    ty = cy - total / 2
    for ln in lines:
        tw = text_w(d, ln, f)
        d.text((cx - tw / 2, ty), ln, font=f, fill=C_INK)
        ty += lh


def arrow(d, x1, y1, x2, y2, label=None, color=None, width=4, head=16, dashed=False):
    c = _color(color)[1] if color else C_LINE
    f = F_TINY
    if dashed:
        _dashed_line(d, x1, y1, x2, y2, c, width, 16, 12)
    else:
        d.line([x1, y1, x2, y2], fill=c, width=width)
    ang = math.atan2(y2 - y1, x2 - x1)
    p = [(x2, y2),
         (x2 - head * math.cos(ang - 0.42), y2 - head * math.sin(ang - 0.42)),
         (x2 - head * math.cos(ang + 0.42), y2 - head * math.sin(ang + 0.42))]
    d.polygon(p, fill=c)
    if label:
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        tw = text_w(d, label, f)
        vertical = abs(x2 - x1) < abs(y2 - y1)
        if vertical:
            n = label.count("\n") + 1
            d.text((mx + 10, my - n * (f.size + 4) / 2), label, font=f, fill="#444444")
        else:
            d.text((mx - tw / 2, my - f.size - 8), label, font=f, fill="#444444")


def _dashed_line(d, x1, y1, x2, y2, color, width, dash, gap):
    total = math.hypot(x2 - x1, y2 - y1)
    if total == 0:
        return
    ux, uy = (x2 - x1) / total, (y2 - y1) / total
    pos = 0.0
    while pos < total:
        end = min(pos + dash, total)
        d.line([x1 + ux * pos, y1 + uy * pos, x1 + ux * end, y1 + uy * end],
               fill=color, width=width)
        pos = end + gap


def line(d, x1, y1, x2, y2, color=None, width=4):
    """普通连线（无箭头）。"""
    c = _color(color)[1] if color else C_LINE
    d.line([x1, y1, x2, y2], fill=c, width=width)


def text(d, x, y, s, font=None, color=None):
    """在指定位置写文字（左上角对齐）。color 为语义色名或 None(黑)。"""
    f = font or F_SMALL
    c = _color(color)[1] if color else C_INK
    d.text((x, y), s, font=f, fill=c)


def container(d, x, y, w, h, title, color="blue"):
    """带标题的容器框：浅色底 + 标题在左上角。"""
    fill = CONTAINER_FILL.get(color, "#FAFCFF")
    line = _color(color)[1]
    d.rounded_rectangle([x, y, x + w, y + h], radius=20, outline=line,
                        width=4, fill=fill)
    d.text((x + 30, y + 16), title, font=F_HEAD, fill=line)


def trim(im):
    """按非白像素包围盒裁掉四周空白，四边统一留 MARGIN。"""
    m = MARGIN * SCALE
    bbox = ImageChops.difference(im, Image.new("RGB", im.size, "white")).getbbox()
    if bbox is None:
        return im
    x0, y0, x1, y1 = bbox
    return im.crop((max(0, x0 - m), max(0, y0 - m),
                    min(im.size[0], x1 + m), min(im.size[1], y1 + m)))


def save(im, out_dir, name):
    """裁剪、保存，并检测内容是否触到画布边（防止元素被裁）。"""
    os.makedirs(out_dir, exist_ok=True)
    bbox = ImageChops.difference(im, Image.new("RGB", im.size, "white")).getbbox()
    if bbox:
        hit = [n for n, at_edge in
               (("上", bbox[1] <= 0), ("下", bbox[3] >= im.size[1]),
                ("左", bbox[0] <= 0), ("右", bbox[2] >= im.size[0])) if at_edge]
        if hit:
            print(f"  [注意] {name} 内容触到画布{'/'.join(hit)}边，可能有元素被裁掉，请加大画布高度")
    im = trim(im)
    path = os.path.join(out_dir, name)
    im.save(path)
    print(f"  已生成 {name}  {im.size[0]}x{im.size[1]}")
    return path

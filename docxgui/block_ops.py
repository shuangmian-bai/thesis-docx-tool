"""审阅页的块结构操作（纯函数，不依赖 PyQt6）。

类型互转、新建块、插入图片时的文件复制规则集中在这里，与
docxai.blocks_json 的块契约保持一致；review_page 只负责调用与刷新界面。
表格/图片不参与类型互转（结构差异过大，需要时直接插入新块）。
"""
import os
import shutil

from docxgui import blocks_model as bm

#: 允许互转的目标类型键（h 按级别展开；不含 table/img）
CONVERT_KEYS = ("h1", "h2", "h3", "h4", "p", "ol", "quote",
                "caption", "ref", "code")

#: 可插入的块类型键 -> 默认块（img 需调用方给路径，单独走 new_image_block）
INSERT_KINDS = ("p", "h", "ol", "table", "img")


def can_convert(block: tuple) -> bool:
    """文本类、标题、代码可互转；表格、图片不可。"""
    return block[0] == "h" or block[0] in bm.TEXTISH or block[0] == "code"


def _text_of(block: tuple) -> str:
    if block[0] == "h":
        return block[2]
    if block[0] == "code":
        return "\n".join(block[1])
    return block[1]


def convert(block: tuple, key: str) -> tuple:
    """把块转成目标类型键（CONVERT_KEYS 之一），保留可保留的文字内容。"""
    if key not in CONVERT_KEYS:
        raise ValueError(f"不支持的目标类型：{key}")
    text = _text_of(block)
    if key.startswith("h"):
        return ("h", int(key[1:]), text)
    if key == "code":
        return ("code", text.split("\n") if text else [])
    return (key, text)


def new_block(kind: str) -> tuple:
    """插入用的空白块默认值；img 用 new_image_block。"""
    if kind == "p":
        return ("p", "")
    if kind == "h":
        return ("h", 2, "")
    if kind == "ol":
        return ("ol", "")
    if kind == "table":
        return ("table", [["", ""], ["", ""]])
    raise ValueError(f"不支持直接新建：{kind}")


def new_image_block(rel_path: str, caption: str = "") -> tuple:
    return ("img", rel_path, caption)


def copy_image(src: str, images_dir: str, config_dir: str) -> str:
    """把用户选择的本地图片复制进本文档素材目录，返回相对 config 的引用路径。

    与 convert 抽图落盘的引用约定一致（斜杠分隔）。目标已存在且文件大小相同
    视为同一张图，直接复用不复制；大小不同则在文件名后加 _1/_2… 避让，
    不覆盖用户已有素材。
    """
    os.makedirs(images_dir, exist_ok=True)
    base = os.path.basename(src)
    dst = os.path.join(images_dir, base)
    stem, ext = os.path.splitext(base)
    n = 1
    while os.path.exists(dst) and os.path.getsize(dst) != os.path.getsize(src):
        dst = os.path.join(images_dir, f"{stem}_{n}{ext}")
        n += 1
    if not os.path.exists(dst):
        shutil.copyfile(src, dst)
    return os.path.relpath(dst, config_dir).replace(os.sep, "/")

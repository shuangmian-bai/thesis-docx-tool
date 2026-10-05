"""docx → Markdown 拆解模块（内部模块，由 run 编排调用）。

把用户论文 docx 拆解为 Markdown（与 `章节/*.md` 同格式）+ 图片目录，
产物可直接喂给 `docxbuild` 套回模板。

核心流程在 `prepare_blocks()`（run 与 GUI 共用，拿到块供校对）与 `convert_to()`
（独立拆解/开发调试，产物默认落 `output/`）。
"""
import argparse
import os
import re
import zipfile
from typing import Dict, List, Optional, Tuple

from docxconvert import HERE
from docxconvert.extract import extract_images
from docxconvert.markdown import render
from docxconvert.parse import parse_docx

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _rewrite_image_paths(blocks, written, rel_base, images_dir):
    """把图片块里的 zip 内路径改写为相对 rel_base 目录的引用路径。

    rel_base 是引用路径的计算基准，必须与构建端解析图片时的基准一致：
    build 以 `config/` 为基准拼接图片路径，故喂给 build 的 md 也要按
    `config/` 写引用（如 `images/xxx.png`）。
    """
    for blk in blocks:
        if blk[0] == "img" and blk[1] in written:
            full = os.path.join(images_dir, written[blk[1]])
            rel = os.path.relpath(full, rel_base).replace(os.sep, "/")
            blk = list(blk)
            blk[1] = rel
            yield tuple(blk)
        else:
            yield blk


def _strip_front(blocks):
    """丢弃第一个「以数字开头的一级标题」之前的所有块。

    封面、诚信承诺书、目录的标题虽然也是一级标题，但文字不以数字开头；
    正文章节（`1　设计思路`、`2　相关技术`…）以数字开头，以此为起点。
    末尾的「总结」「参考文献」等非数字 H1 在起点之后，照常保留。
    """
    for i, blk in enumerate(blocks):
        if blk[0] == "h" and blk[1] == 1 and blk[2].strip()[:1].isdigit():
            return blocks[i:]
    return blocks


def prepare_blocks(docx_path: str, images_dir: str, *,
                   strip_front: bool = False,
                   rel_base: Optional[str] = None,
                   quiet: bool = False
                   ) -> Tuple[list, str, Dict[str, str]]:
    """解析 docx、抽取图片并改写图片引用路径，但不写 Markdown 文件。

    供 convert_to() 与 GUI 审阅流程共用：GUI 需要先拿到块供人工/AI 校对，
    Markdown 在用户确认后才落盘。

    参数：
        docx_path:   待拆解的 docx。
        images_dir:  图片抽取目录（自动创建）。
        strip_front: 是否丢弃第一个以数字开头的一级标题之前的前置块。
        rel_base:    图片引用路径的计算基准目录；产物要喂给 build 时传
                     `config/`（引用写成 images/xxx.png）；None 时按 images_dir
                     的父目录计算（独立 convert 的原行为）。
        quiet:       静默（GUI 用），不打印过程信息。

    返回：(blocks, images_dir 绝对路径, written {zip路径: 落盘文件名})。
    """
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        raise SystemExit(f"文件不存在：{docx_path}")

    images_dir = os.path.abspath(images_dir)
    rel_base = os.path.abspath(rel_base) if rel_base \
        else os.path.dirname(images_dir)

    blocks, images = parse_docx(docx_path)
    if not quiet:
        print(f"解析得到 {len(blocks)} 个块，图片 {len(images)} 张")

    written: Dict[str, str] = extract_images(docx_path, images, images_dir)
    blocks = list(_rewrite_image_paths(blocks, written, rel_base, images_dir))

    if strip_front:
        before = len(blocks)
        blocks = _strip_front(blocks)
        if not quiet:
            print(f"--strip-front：丢弃 {before - len(blocks)} 个前置块")
    return blocks, images_dir, written


def convert_to(docx_path: str, out_md: str, images_dir: str,
               strip_front: bool = False,
               rel_base: Optional[str] = None
               ) -> Tuple[str, str, Dict[str, int]]:
    """把 docx 拆解为 Markdown 与图片目录（核心流程，不含参数解析）。

    参数：
        docx_path:   待拆解的 docx 绝对/相对路径。
        out_md:      Markdown 产物路径（目录不存在会自动创建）。
        images_dir:  图片抽取目录。
        strip_front: 是否丢弃第一个以数字开头的一级标题之前的前置块。
        rel_base:    md 中图片引用路径的计算基准目录；必须与构建端的解析基准
                     一致——产物要喂给 build 时传 `config/` 目录（引用写成
                     `images/xxx.png`）。None 时按 md 自身所在目录计算
                     （独立 convert 命令的原行为）。

    返回：(md 绝对路径, 图片目录绝对路径, 各类块计数)。
    """
    blocks, images_dir, written = prepare_blocks(
        docx_path, images_dir, strip_front=strip_front, rel_base=rel_base)

    out_md = os.path.abspath(out_md)
    md = render(blocks)
    os.makedirs(os.path.dirname(out_md), exist_ok=True)
    with open(out_md, "w", encoding="utf-8") as fh:
        fh.write(md)

    print(f"已写 {os.path.relpath(out_md, HERE)}")
    if written:
        print(f"图片 {len(written)} 张 → {os.path.relpath(images_dir, HERE)}/")
    # 各类块计数
    counts: Dict[str, int] = {}
    for blk in blocks:
        counts[blk[0]] = counts.get(blk[0], 0) + 1
    print("块统计：" + " · ".join(f"{k} {v}" for k, v in counts.items()))
    return out_md, images_dir, counts


def main(argv: Optional[List[str]] = None):
    ap = argparse.ArgumentParser(
        description="把用户撰写的毕业论文 docx 拆解为 章节/*.md 格式的 Markdown。")
    ap.add_argument("docx", help="用户撰写的论文 Word 文档（.docx）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出的 Markdown 路径（默认与 docx 同名 .md，放在 output/ 下）")
    ap.add_argument("--images-dir", default=None,
                    help="图片输出目录（默认 <md 同名>_images/，与 md 同级）")
    ap.add_argument("--strip-front", action="store_true",
                    help="丢弃第一个一级标题之前的内容（封面、摘要等前置页）")
    args = ap.parse_args(argv)

    base = os.path.splitext(os.path.basename(args.docx))[0]
    out_md = args.output or os.path.join(HERE, "output", f"{base}.md")
    md_dir = os.path.dirname(os.path.abspath(out_md))
    images_dir = args.images_dir or os.path.join(md_dir, f"{base}_images")

    convert_to(args.docx, out_md, images_dir, strip_front=args.strip_front)


if __name__ == "__main__":
    main()

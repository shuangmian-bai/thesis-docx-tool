"""`convert_docx.py` 的主流程。

把用户论文 docx 拆解为 Markdown（与 `章节/*.md` 同格式）+ 图片目录，
产物可直接喂给 `build_docx.py` 套回模板。
"""
import argparse
import os

from docxconvert import HERE
from docxconvert.extract import extract_images
from docxconvert.markdown import render
from docxconvert.parse import parse_docx


def _rewrite_image_paths(blocks, written, md_dir, images_dir):
    """把图片块里的 zip 内路径改写为相对 md 文件的路径。"""
    for blk in blocks:
        if blk[0] == "img" and blk[1] in written:
            full = os.path.join(images_dir, written[blk[1]])
            rel = os.path.relpath(full, md_dir).replace(os.sep, "/")
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


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="把用户撰写的毕业论文 docx 拆解为 章节/*.md 格式的 Markdown。")
    ap.add_argument("docx", help="用户撰写的论文 Word 文档（.docx）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出的 Markdown 路径（默认与 docx 同名 .md，放在论文目录下）")
    ap.add_argument("--images-dir", default=None,
                    help="图片输出目录（默认 <md 同名>_images/，与 md 同级）")
    ap.add_argument("--strip-front", action="store_true",
                    help="丢弃第一个一级标题之前的内容（封面、摘要等前置页）")
    args = ap.parse_args(argv)

    docx_path = os.path.abspath(args.docx)
    if not os.path.exists(docx_path):
        raise SystemExit(f"文件不存在：{docx_path}")

    base = os.path.splitext(os.path.basename(docx_path))[0]
    out_md = args.output or os.path.join(HERE, f"{base}.md")
    out_md = os.path.abspath(out_md)
    md_dir = os.path.dirname(out_md)
    images_dir = args.images_dir or os.path.join(md_dir, f"{base}_images")
    images_dir = os.path.abspath(images_dir)

    blocks, images = parse_docx(docx_path)
    print(f"解析得到 {len(blocks)} 个块，图片 {len(images)} 张")

    written = extract_images(docx_path, images, images_dir)
    blocks = list(_rewrite_image_paths(blocks, written, md_dir, images_dir))

    if args.strip_front:
        before = len(blocks)
        blocks = _strip_front(blocks)
        print(f"--strip-front：丢弃 {before - len(blocks)} 个前置块")

    md = render(blocks)
    os.makedirs(md_dir, exist_ok=True)
    with open(out_md, "w", encoding="utf-8") as fh:
        fh.write(md)

    print(f"已写 {os.path.relpath(out_md, HERE)}")
    if written:
        print(f"图片 {len(written)} 张 → {os.path.relpath(images_dir, HERE)}/")
    # 各类块计数
    counts = {}
    for blk in blocks:
        counts[blk[0]] = counts.get(blk[0], 0) + 1
    print("块统计：" + " · ".join(f"{k} {v}" for k, v in counts.items()))


if __name__ == "__main__":
    main()

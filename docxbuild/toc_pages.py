#!/usr/bin/env python3
"""从渲染出的 PDF 里反查各级标题的页码，写入 `toc_pages.json`。

`build_docx.py` 生成的目录是一个 TOC 域，域结果里预填的页码由本脚本提供。
之所以要绕这一圈，是因为页码只有排版引擎算得出来：先用占位目录生成一遍 docx，
渲染成 PDF，再从 PDF 里量出每个标题落在第几页，回填后重新生成，目录就固定了。

用法：

    python3 main.py build                     # 第一遍：目录页码留空
    soffice --headless --convert-to pdf 论文_v1_*.docx
    python3 main.py toc 论文_v1_*.pdf         # 量页码，写 toc_pages.json
    python3 main.py build                     # 第二遍：目录带上页码
"""
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHAP_DIR = os.path.join(HERE, "config", "章节")
TOC_PAGES = os.path.join(HERE, "output", "toc_pages.json")

# 目录页的判据：带点前导符的目录项会渲染出成串的点
LEADER_RUN = re.compile(r"\.{4,}")


def outline(chap_dir=CHAP_DIR):
    """按文档顺序取出各章的 # 与 ## 标题。

    chap_dir 默认为全局 config/章节/（MD 路线）；Word 路线/闭环的
    哈希工作目录模式由调用方传入工作目录下的 章节/，避免读到别篇论文。
    """
    items = []
    for fn in sorted(f for f in os.listdir(chap_dir) if f.endswith(".md")):
        with open(os.path.join(chap_dir, fn), encoding="utf-8") as fh:
            for ln in fh:
                m = re.match(r"^(#{1,2})\s+(.*?)\s*$", ln)
                if m:
                    items.append((len(m.group(1)), m.group(2)))
    return items


def norm(s):
    """去掉所有空白（含全角空格），便于跨渲染差异匹配。"""
    return re.sub(r"\s+", "", s)


def page_texts(pdf):
    out = subprocess.run(["pdftotext", "-layout", pdf, "-"],
                         capture_output=True, text=True, check=True).stdout
    return out.split("\f")


def find_body_start(pages, items):
    """正文第 1 页在 PDF 里的下标：跳过封面、承诺书与目录页。

    优先用「目录页含成串点前导符」判据；识别不出时，退回到「首页含首个一级标题」。
    首个一级标题从 `items` 里取，不写死具体文字。
    """
    last_toc = -1
    for i, t in enumerate(pages):
        if len(LEADER_RUN.findall(t)) >= 5:
            last_toc = i
    if last_toc >= 0:
        return last_toc + 1
    # 没有识别出目录页（例如目录页码留空导致没有前导符），退回到"首页含首个一级标题"
    first_h1 = next((title for lv, title in items if lv == 1), None)
    if first_h1:
        key = norm(first_h1)
        for i, t in enumerate(pages):
            if key in norm(t):
                return i
    raise SystemExit("无法定位正文起始页")


def main(argv=None):
    args = list(argv if argv is not None else sys.argv[1:])
    if not args:
        raise SystemExit(__doc__)
    # 可选参数：--chap-dir 章节目录 / --out 页码表输出路径（默认全局位置，
    # 哈希工作目录模式由编排层传入工作目录内路径，防止多论文互相污染）
    chap_dir = CHAP_DIR
    out_path = TOC_PAGES
    positional = []
    skip = False
    for i, a in enumerate(args):
        if skip:
            skip = False
            continue
        if a == "--chap-dir" and i + 1 < len(args):
            chap_dir = args[i + 1]
            skip = True
        elif a == "--out" and i + 1 < len(args):
            out_path = args[i + 1]
            skip = True
        else:
            positional.append(a)
    if not positional:
        raise SystemExit(__doc__)
    pdf = positional[0]
    if not os.path.isabs(pdf):
        pdf = os.path.join(HERE, pdf)
    if not os.path.exists(pdf):
        raise SystemExit(f"找不到 PDF：{pdf}")

    pages = page_texts(pdf)
    items = outline(chap_dir)
    start = find_body_start(pages, items)
    print(f"PDF 共 {len(pages)} 页（末页可能为空）；正文起始于 PDF 第 {start + 1} 页")
    print(f"大纲共 {len(items)} 条标题（章 {sum(1 for lv, _ in items if lv == 1)}，"
          f"节 {sum(1 for lv, _ in items if lv == 2)}）")

    pages_norm = [norm(t) for t in pages]
    result, cursor, missing = {}, start, []
    for level, title in items:
        key = norm(title)
        found = None
        # 标题在文档里是有序的，从上次命中的位置继续往后找即可
        for i in range(cursor, len(pages_norm)):
            if key in pages_norm[i]:
                found = i
                break
        if found is None:
            missing.append(title)
            continue
        cursor = found
        result[title] = found - start + 1

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=1)
        fh.write("\n")
    print(f"已写 {os.path.relpath(out_path, HERE)}：{len(result)} 条页码")
    if missing:
        print(f"[注意] {len(missing)} 条标题未在 PDF 中定位到，页码将留空：")
        for t in missing[:10]:
            print(f"    {t}")


if __name__ == "__main__":
    main()

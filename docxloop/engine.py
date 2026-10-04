"""闭环引擎：对输入 docx 执行 convert → build，然后全盘对比。

每个输入文件用其内容 SHA256 哈希（前 16 位）作为隔离标识，创建独立工作目录，
避免多个文件共享 config/章节、config/images 造成污染，支持多线程并发。

工作目录结构（默认 .cache/loop_work/）：
    .cache/loop_work/
    ├── index.json              # {哈希: 原文件名}
    └── <hash16>/
        ├── src.md              # convert 产物
        ├── images/             # 抽取的图片
        ├── 章节/               # build 输入
        ├── output/             # build 输出
        └── toc.json            # 目录页码
"""
import hashlib
import json
import os
import shutil
from typing import Dict, List, Optional

from docxloop.compare import CompareResult, compare

#: 闭环工作根目录
WORK_ROOT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    ".cache", "loop_work")


def file_hash(path: str) -> str:
    """计算文件 SHA256，取前 16 位十六进制。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def work_dir_for(hash16: str) -> str:
    """返回哈希对应的工作目录路径。"""
    return os.path.join(WORK_ROOT, hash16)


def load_index() -> Dict[str, str]:
    """加载哈希索引文件。"""
    idx_path = os.path.join(WORK_ROOT, "index.json")
    if os.path.exists(idx_path):
        with open(idx_path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_index(index: Dict[str, str]):
    """保存哈希索引文件。"""
    os.makedirs(WORK_ROOT, exist_ok=True)
    with open(os.path.join(WORK_ROOT, "index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def resolve_input(spec: str) -> Optional[str]:
    """把命令行参数解析为实际文件路径。

    spec 可以是文件路径或 16 位哈希：
    - 文件路径：计算哈希，创建/定位工作目录，返回原路径
    - 哈希：从索引查原文件名；若工作目录不存在，返回 None
    """
    # 16 位十六进制视为哈希
    if len(spec) == 16 and all(c in "0123456789abcdef" for c in spec.lower()):
        index = load_index()
        orig = index.get(spec)
        if orig and os.path.exists(work_dir_for(spec)):
            return orig
        return None
    # 文件路径
    if os.path.isfile(spec) and spec.endswith(".docx"):
        return os.path.abspath(spec)
    return None


def run_closed_loop(src_path: str, template_path: str,
                    reuse: bool = True) -> CompareResult:
    """对单个 docx 跑闭环，返回对比结果。

    reuse=True 时，若哈希对应的工作目录已存在且有输出 docx，直接用已有产物对比
    （适合反复调整对比引擎）；False 时强制重新 convert→build。
    """
    h = file_hash(src_path)
    wdir = work_dir_for(h)
    os.makedirs(wdir, exist_ok=True)

    # 记录索引（存绝对路径，哈希复用时能直接定位原文件）
    src_path = os.path.abspath(src_path)
    index = load_index()
    index[h] = src_path
    save_index(index)

    out_docx = None
    if reuse:
        out_dir = os.path.join(wdir, "output")
        if os.path.isdir(out_dir):
            for f in os.listdir(out_dir):
                if f.endswith(".docx"):
                    out_docx = os.path.join(out_dir, f)
                    break

    if out_docx is None:
        # ── convert ──
        from docxconvert.cli import convert_to, extract_cover, extract_template_title
        from docxbuild.docinfo import VERSION, save_cover_to
        md_out = os.path.join(wdir, "src.md")
        img_dir = os.path.join(wdir, "images")
        os.makedirs(img_dir, exist_ok=True)
        # rel_base=wdir：md 里图片引用写成 images/xxx.png，build 时 IMG_BASE=wdir
        convert_to(src_path, md_out, img_dir, strip_front=True, rel_base=wdir)

        # 提取封面字段存到工作目录 cover.json（与 docxflow Word 路线一致），
        # 每个论文独立，不污染全局；build 时通过 cover_path 加载
        cover = extract_cover(src_path)
        tpl_title = extract_template_title(src_path)
        save_cover_to(os.path.join(wdir, "cover.json"), {
            "title": cover.get("成果名称", ""),
            "version": VERSION,
            "template_title": tpl_title,
            "cover": cover,
        })

        # ── build ──
        from docxbuild.cli import build
        chap_dir = os.path.join(wdir, "章节")
        os.makedirs(chap_dir, exist_ok=True)
        shutil.copy(md_out, os.path.join(chap_dir, "00_src.md"))
        out_docx = build(
            template=template_path,
            chap_dir=chap_dir,
            output_dir=os.path.join(wdir, "output"),
            toc_pages=os.path.join(wdir, "toc.json"),
            images_base=wdir,
            cover_path=os.path.join(wdir, "cover.json"),
        )

    return compare(src_path, out_docx)


def run_batch(input_specs: List[str], template_path: str,
              reuse: bool = True) -> List[CompareResult]:
    """批量跑闭环。input_specs 可以是文件、文件夹或哈希。"""
    paths = []
    for spec in input_specs:
        resolved = resolve_input(spec)
        if resolved:
            paths.append(resolved)
        elif os.path.isdir(spec):
            for f in sorted(os.listdir(spec)):
                if f.endswith(".docx") and not f.startswith("~$"):
                    paths.append(os.path.join(spec, f))
        else:
            print(f"[注意] 无法解析输入：{spec}"
                  f"（文件不存在或哈希对应的工作目录不存在）")

    results = []
    for p in paths:
        results.append(run_closed_loop(p, template_path, reuse=reuse))
    return results

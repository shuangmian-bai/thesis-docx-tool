"""一键流水线：`run` 子命令的编排实现（应用层）。

面向用户只有一条命令、两种输入（同一入口，参数区分）：

    python3 main.py run 论文.docx [--template 模板.docx]   # 拆解 word → md 存哈希工作目录 → 构建
    python3 main.py run <16位哈希> [--template 模板.docx]  # 复用工作目录 md 重新构建

本模块的 run_pipeline() 是**全流程唯一编排实现**，命令行 / GUI / 闭环测试
三个入口只负责解析参数后调用它，不得各自实现 convert/build/章节写入等
流程逻辑（子项目 CLAUDE.md 项目特定规则第 8 条）。

所有路径统一走「哈希工作目录」，不存在 MD 路线的全局目录分支——这是
「要么一起出 bug、要么全部正常」的保证：一条目录结构、一套产物落点。

内部自动串联以下阶段，用户不需要手工分步运行，也无任何开关参数：

    1. 判定输入（word 或哈希）→ 统一落到哈希工作目录
    2. 前置检查：模板、章节源（哈希路线），缺件一次性报齐
    3. 清理工作目录旧产物（output/、toc.json）
    4. 拆解（仅 word 输入）：docx → 工作目录 src.md + 章节/*.md + images/
    5. 架构图：config/figures/ 有 DSL 定义时自动渲染 PNG（无则跳过）
    6. 构建：build 第一遍，生成占位目录的成品 docx
    7. 页码回填：本机有 soffice 与 pdftotext 时自动 docx→PDF→量页码→第二遍构建；
       缺任一则跳过（Word 中 Ctrl+A 后 F9 也可更新目录），不阻断出稿
    8. 审计：对最终成品跑 audit，报告照打，审计结论不改变流水线成败
"""
import argparse
import importlib.util
import os
import re
import shutil
import subprocess
from typing import List, Optional

from docxflow import HERE
from docxflow.workdir import is_hash16, prepare_work_dir, resolve_hash
from docxbuild.cli import OUTPUT_DIR
from docxfig.parser import list_figures

# 编排层固定路径（与各原子模块中的定义同值；集中在此供前置检查使用）
CONFIG_DIR = os.path.join(HERE, "config")
FIG_DIR = os.path.join(CONFIG_DIR, "figures")
DEFAULT_TEMPLATE = os.path.join(CONFIG_DIR, "template.docx")

# soffice 使用独立用户配置目录，避免与用户正在运行的 LibreOffice 抢 profile
LO_PROFILE = "file:///tmp/thesis-docx-tool-lo-profile"


def _step(title: str) -> None:
    """打印阶段标题。"""
    print(f"\n{'=' * 30} {title} {'=' * 30}")


# ── 应用层：章节源写入，命令行 / GUI 共用 ──

def _safe_filename(title: str) -> str:
    """章节标题转合法文件名：去掉路径分隔符等非法字符，空标题回退为「章节」。"""
    name = re.sub(r'[\\/:*?"<>|\r\n]+', "_", title).strip()
    return name or "章节"


def write_chapters_from_blocks(blocks, chap_dir: str) -> List[str]:
    """把块序列按 H1 拆成多章，写入指定章节目录（哈希工作目录的 章节/）。

    这是**全项目唯一的章节落盘实现**：命令行 Word 路线与 GUI 审阅页统一走这里，
    保证各入口行为完全等价。同时把整篇渲染写到 chap_dir 父目录（工作目录）的
    src.md，保证命令行 / GUI 的工作目录产物一致。
    返回写入的文件名列表。
    """
    from docxai.blocks_json import split_chapters
    from docxconvert.markdown import render

    target = chap_dir
    # 清理目标目录下数字开头的旧 md
    if os.path.isdir(target):
        for fn in os.listdir(target):
            if fn.endswith(".md") and fn[:1].isdigit():
                os.remove(os.path.join(target, fn))
    os.makedirs(target, exist_ok=True)
    chapters = split_chapters(blocks)
    written = []
    for i, ch in enumerate(chapters, 1):
        ch_blocks = [blocks[idx] for idx in ch["indexes"]]
        title = _safe_filename(ch["title"])
        fn = f"{i:02d}_{title}.md"
        with open(os.path.join(target, fn), "w", encoding="utf-8") as fh:
            fh.write(render(ch_blocks))
        written.append(fn)
    # 整篇渲染写到工作目录 src.md（与章节并列，供回溯与闭环比对）
    with open(os.path.join(os.path.dirname(target), "src.md"),
              "w", encoding="utf-8") as fh:
        fh.write(render(blocks))
    return written


def _clean_work_dir_products(work_dir: str) -> None:
    """清空哈希工作目录的构建产物（output/、toc.json），防脏数据污染。

    应用层唯一清理入口；检查通过后才调用，避免检查失败误删上一轮可用成品。
    """
    out_dir = os.path.join(work_dir, "output")
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    toc = os.path.join(work_dir, "toc.json")
    if os.path.exists(toc):
        os.remove(toc)


def _precheck(template: str, work_dir: str, have_docx: bool) -> None:
    """开工前一次性列出所有缺失件，避免跑到半途才报错。

    have_docx=True（word 路线）时章节源将由拆解生成，不检查；
    have_docx=False（哈希路线）时检查工作目录的 章节/ 已有拆解产物。
    """
    problems: List[str] = []
    if not os.path.exists(template):
        problems.append(f"模板不存在：{os.path.relpath(template, HERE)}"
                        f"（放 config/template.docx 或用 --template 指定）")
    if not have_docx:
        chap_dir = os.path.join(work_dir, "章节")
        if not os.path.isdir(chap_dir) or not [
                f for f in os.listdir(chap_dir) if f.endswith(".md")]:
            problems.append(f"找不到章节 Markdown：{os.path.relpath(chap_dir, HERE)}/"
                            " 下没有 .md 文件（该哈希对应的工作目录无拆解产物，"
                            "请先传入文件路径）")
    if problems:
        msg = "\n  ".join(problems)
        raise SystemExit(f"开工前检查未通过：\n  {msg}")


def _fig_stage() -> bool:
    """渲染架构图（config/figures/ 有 DSL 才渲染）。返回是否实际生成。"""
    figures = list_figures(os.path.normpath(FIG_DIR))
    if not figures:
        print("config/figures/ 下无图定义，跳过架构图阶段")
        return False
    if importlib.util.find_spec("PIL") is None:
        raise SystemExit("存在架构图定义但未安装 Pillow，无法渲染："
                         "请先执行 pip install -r requirements.txt")
    from docxfig.cli import generate
    generate(None)
    return True


def _build(template: str, work_dir: str) -> str:
    """跑一遍 build，返回成品 docx 绝对路径（统一走哈希工作目录）。"""
    from docxbuild.cli import build as build_fn
    return build_fn(
        template=template,
        chap_dir=os.path.join(work_dir, "章节"),
        output_dir=os.path.join(work_dir, "output"),
        toc_pages=os.path.join(work_dir, "toc.json"),
        images_base=work_dir,
    )


def _toc_stage(product: str, template: str, work_dir: str) -> bool:
    """自动转 PDF、量页码并第二遍构建。缺 soffice/pdftotext 自动跳过。"""
    missing = [t for t in ("soffice", "pdftotext") if shutil.which(t) is None]
    if missing:
        print(f"[注意] 缺少 {', '.join(missing)}，跳过目录页码回填（不影响出稿）。")
        print("       安装后重跑本命令可自动回填；或在 Word 中按 Ctrl+A 后 F9 更新目录。")
        return False

    # PDF 落在工作目录的 output/，页码表写工作目录 toc.json
    out_dir = os.path.dirname(product)
    print("正在用 LibreOffice 把成品转为 PDF ...")
    proc = subprocess.run(
        ["soffice", f"-env:UserInstallation={LO_PROFILE}", "--headless",
         "--convert-to", "pdf", "--outdir", out_dir, product],
        capture_output=True, text=True, timeout=300)
    pdf = os.path.splitext(product)[0] + ".pdf"
    if proc.returncode != 0 or not os.path.exists(pdf):
        print(f"[注意] docx 转 PDF 失败，跳过页码回填：{proc.stderr.strip() or proc.stdout.strip()}")
        return False

    from docxbuild.toc_pages import main as toc_main
    toc_main([pdf, "--chap-dir", os.path.join(work_dir, "章节"),
              "--out", os.path.join(work_dir, "toc.json")])

    print("页码表已生成，第二遍构建以回填目录 ...")
    return _build(template, work_dir) is not None


def _audit_stage(product: str, template: str, template_explicit: bool) -> None:
    """成品审计：报告照打，任何问题都不影响流水线已产出的成品。

    用户显式指定外部模板时把 --template 透传给 audit（审计必须拿同一个模板比，
    audit 同时会停用只对默认模板登记的预期偏离表）；用默认模板时不传，
    保留 audit 的默认行为。
    """
    _step("阶段 · 成品格式审计")
    argv = ["audit", "--product", product]
    if template_explicit:
        argv += ["--template", template]
    try:
        from docxaudit.cli import main as audit_main
        audit_main(argv)
    except SystemExit as e:
        # audit 以退出码 1 表达「有注意项」，对流水线而言只是报告，不当失败
        if e.code:
            print("[注意] 审计存在需关注项，见上方报告；成品不受影响，可按报告逐项确认。")
    except Exception as e:  # 审计是只读附加环节，不因其自身异常拖垮整条流水线
        print(f"[注意] 审计阶段异常，已跳过：{e}")


def _convert_stage(docx: str, work_dir: str) -> None:
    """拆解用户 docx：解析 → src.md + 章节/*.md（固定剥离前置封面/承诺书/目录）。

    与 GUI 审阅页共用同一份块处理链路（prepare_blocks →
    write_chapters_from_blocks），保证命令行与 GUI 的拆解产物完全一致。
    """
    from docxconvert.cli import prepare_blocks

    # rel_base 传 work_dir：图片引用写成 images/x.png，
    # build 时 images_base=work_dir，与工作目录结构匹配
    images_dir = os.path.join(work_dir, "images")
    blocks, _, _ = prepare_blocks(docx, images_dir,
                                  strip_front=True,
                                  rel_base=work_dir)
    # 按 H1 拆分写入工作目录的 章节/，同时把整篇渲染落到 src.md
    # （src.md 由 write_chapters_from_blocks 内部一并落盘，保证 CLI/GUI 产物一致）
    written = write_chapters_from_blocks(
        blocks, chap_dir=os.path.join(work_dir, "章节"))
    print(f"  已写入 {len(written)} 个章节 → "
          f"{os.path.relpath(os.path.join(work_dir, '章节'), HERE)}/")


def run_pipeline(source: str, template: Optional[str] = None, *,
                 no_audit: bool = False,
                 copy_to_output: bool = True) -> str:
    """全流程唯一编排实现：从 word（或哈希工作目录）到成品 Word。

    统一入口，无模式分叉——source 为 word 文件路径或 16 位哈希，内部统一落到
    「哈希工作目录」这一条目录结构：

    - word 文件：按内容哈希创建/定位工作目录 → 拆解 → 构建
    - 16 位哈希：定位已有工作目录 → 构建（跳过拆解，复用已拆解/编辑的 md）

    template=None 时用默认模板 config/template.docx。

    no_audit / copy_to_output 是**内部参数**，仅供 loop 等程序内调用
    （闭环测试跳过审计、产物不回拷全局 output/），命令行 main 不暴露。

    返回成品 docx 绝对路径。
    """
    template_explicit = template is not None
    template = os.path.abspath(template) if template else DEFAULT_TEMPLATE

    # 判定输入类型，统一落到哈希工作目录（work_dir 永远非 None）
    docx = None
    if is_hash16(source):
        work_dir = resolve_hash(source.lower())
        if not work_dir:
            raise SystemExit(f"哈希 {source.lower()} 对应的工作目录不存在，请先传入文件路径")
    else:
        if not os.path.isfile(source):
            raise SystemExit(f"输入不存在：{source}"
                             "（不是文件路径；若是哈希，对应工作目录不存在，请先传入文件路径）")
        docx = source
        work_dir = prepare_work_dir(docx)
    print(f"工作目录（哈希隔离）：{os.path.relpath(work_dir, HERE)}")

    _step("阶段 · 开工前检查")
    _precheck(template, work_dir, docx is not None)
    print("检查通过。")

    # 检查通过后统一清理工作目录旧产物（output/、toc.json），防脏数据污染
    _clean_work_dir_products(work_dir)

    if docx:
        _step("阶段 · 拆解用户 Word 为章节 Markdown")
        _convert_stage(docx, work_dir)
        print(f"拆解产物在工作目录中，修改章节后可用 "
              f"`run {os.path.basename(work_dir)}` 重新构建。")

    _step("阶段 · 架构图渲染")
    _fig_stage()

    _step("阶段 · 构建 Word（第一遍）")
    product = _build(template, work_dir)

    _step("阶段 · 目录页码回填")
    _toc_stage(product, template, work_dir)

    if not no_audit:
        _audit_stage(product, template, template_explicit)

    # 把成品从工作目录复制到全局 output/，方便用户取用
    if copy_to_output and os.path.exists(product):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        final = os.path.join(OUTPUT_DIR, os.path.basename(product))
        shutil.copy2(product, final)
        product = final

    _step("完成")
    print(f"成品：{os.path.relpath(product, HERE)}")
    return product


def main(argv: Optional[List[str]] = None) -> int:
    """命令行入口：只做参数解析，流程一律交 run_pipeline。

    面向用户只暴露两种用法：`run 论文.docx`（word 路线）与 `run <16位哈希>`
    （复用工作目录），模板用 `--template` 指定。无任何其它开关。
    """
    ap = argparse.ArgumentParser(
        description="一键流水线：从用户 docx（或已拆解的哈希工作目录）端到端生成"
                    "成品 Word，内部自动完成拆解、架构图、构建、页码回填与审计。")
    ap.add_argument("source",
                    help="用户论文 docx 路径，或 16 位哈希（复用已有工作目录重建）")
    ap.add_argument("--template", default=None,
                    help=f"模板 docx 路径（默认 {DEFAULT_TEMPLATE}）")
    args = ap.parse_args(argv)

    product = run_pipeline(args.source, args.template)
    return 0 if product else 1

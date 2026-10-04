"""一键流水线：`run` 子命令的编排实现（应用层）。

面向用户只有一条命令、三条输入路径（同一入口，参数区分）：

    python3 main.py run                 # MD 路线：config/章节/*.md → 成品 docx
    python3 main.py run 用户论文.docx    # Word 路线：先拆解入哈希工作目录，其余相同
    python3 main.py run <16位哈希>      # 复用已有工作目录（等价 --work-dir）

本模块的 run_pipeline() 是**全流程唯一编排实现**，命令行 / GUI / 闭环测试
三个入口只负责解析参数后调用它，不得各自实现 convert/build/章节写入等
流程逻辑（子项目 CLAUDE.md 项目特定规则第 8 条）。

内部自动串联以下阶段，用户不需要手工分步运行：

    1. 前置检查：模板、签名图、章节源，缺件一次性报齐
    2. 拆解（仅 Word 路线）：docx → 工作目录 src.md + 章节/*.md + images/
    3. 架构图：config/figures/ 有 DSL 定义时自动渲染 PNG（无则跳过）
    4. 构建：build 第一遍，生成占位目录的成品 docx
    5. 页码回填：本机有 soffice 与 pdftotext 时自动 docx→PDF→量页码→第二遍构建；
       缺任一则跳过（Word 中 Ctrl+A 后 F9 也可更新目录），不阻断出稿
    6. 审计：对最终成品跑 audit，报告照打，审计结论不改变流水线成败

分步子命令（fig/build/toc/audit/convert）仍保留，供高级用户单步调试。
"""
import argparse
import importlib.util
import os
import re
import shutil
import subprocess
from typing import List, Optional

from docxflow import HERE
from docxflow.workdir import (WORK_ROOT, is_hash16, prepare_work_dir,
                              resolve_hash)
from docxbuild.cli import CHAP_DIR, OUTPUT_DIR
from docxbuild.docinfo import VERSION
from docxfig.parser import list_figures

# 编排层固定路径（与各原子模块中的定义同值；集中在此供前置检查使用）
CONFIG_DIR = os.path.join(HERE, "config")
FIG_DIR = os.path.join(CONFIG_DIR, "figures")
CONFIG_IMAGES = os.path.join(CONFIG_DIR, "images")
SIGNATURE = os.path.join(CONFIG_IMAGES, "signature.png")
DEFAULT_TEMPLATE = os.path.join(CONFIG_DIR, "template.docx")

# soffice 使用独立用户配置目录，避免与用户正在运行的 LibreOffice 抢 profile
LO_PROFILE = "file:///tmp/thesis-docx-tool-lo-profile"

# [兼容别名] 工作根目录的唯一权威定义在 docxflow.workdir.WORK_ROOT，
# 此处仅为旧引用保留，新代码一律从 docxflow.workdir 导入
RUN_WORK_ROOT = WORK_ROOT


def _step(title: str) -> None:
    """打印阶段标题。"""
    print(f"\n{'=' * 30} {title} {'=' * 30}")


# ── 应用层：章节源准备（清理 + 写入），命令行 / GUI / 闭环共用 ──

def clear_old_chapters() -> None:
    """清理 config/章节/ 下所有数字开头的 .md 文件，避免新旧章节混杂。

    非数字开头的文件（如模板示例、整篇拆解产物）不动，由 build 自行跳过。
    本函数是应用层唯一的 config 章节清理入口，出 bug 只修此处。
    """
    if not os.path.isdir(CHAP_DIR):
        return
    for fn in os.listdir(CHAP_DIR):
        if fn.endswith(".md") and fn[:1].isdigit():
            os.remove(os.path.join(CHAP_DIR, fn))


def _safe_filename(title: str) -> str:
    """章节标题转合法文件名：去掉路径分隔符等非法字符，空标题回退为「章节」。"""
    name = re.sub(r'[\\/:*?"<>|\r\n]+', "_", title).strip()
    return name or "章节"


def write_chapters_from_blocks(blocks, chap_dir: Optional[str] = None) -> List[str]:
    """把块序列按 H1 拆成多章，写入章节目录（先清理旧分章）。

    chap_dir 为 None 时写入 config/章节/（MD 路线）；
    传入时写入指定目录（Word 路线/GUI/闭环的哈希工作目录），避免多论文污染。
    这是**全项目唯一的章节落盘实现**：GUI 审阅页、命令行 Word 路线、闭环
    测试统一走这里，保证各入口行为完全等价。
    返回写入的文件名列表。
    """
    from docxai.blocks_json import split_chapters
    from docxconvert.markdown import render

    target = chap_dir or CHAP_DIR
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
    return written


def _precheck(template: str, have_docx: bool, md_target: Optional[str],
              force: bool, work_dir: Optional[str] = None) -> None:
    """开工前一次性列出所有缺失件，避免跑到半途才报错。

    章节源检查看的是**实际生效的章节目录**：
    work_dir 模式下检查工作目录的 章节/，否则检查 config/章节/。
    """
    problems: List[str] = []
    if not os.path.exists(template):
        problems.append(f"模板不存在：{os.path.relpath(template, HERE)}"
                        f"（放 config/template.docx 或用 --template 指定）")
    if not os.path.exists(SIGNATURE):
        problems.append("承诺书签名图不存在：config/images/signature.png（请自备）")
    if not have_docx:
        chap_dir = os.path.join(work_dir, "章节") if work_dir else CHAP_DIR
        if not os.path.isdir(chap_dir) or not [
                f for f in os.listdir(chap_dir) if f.endswith(".md")]:
            problems.append(f"找不到章节 Markdown：{os.path.relpath(chap_dir, HERE)}/"
                            " 下没有 .md 文件"
                            "（直接写 MD，或用 `run 用户论文.docx` 传入 Word）")
    elif md_target and os.path.exists(md_target) and not force:
        problems.append(f"拆解目标已存在：{os.path.relpath(md_target, HERE)}"
                        f"（确认覆盖请加 --force）")
    if problems:
        msg = "\n  ".join(problems)
        raise SystemExit(f"开工前检查未通过：\n  {msg}")

    # 封面字段：Word 路线从 docx 自动提取存工作目录；GUI 路线在模板预览页填写；
    # MD 路线使用占位符，均可正常出稿（不再依赖全局 config/cover.json）


def _fig_stage(no_fig: bool) -> bool:
    """渲染架构图。返回是否实际生成。"""
    if no_fig:
        print("已跳过架构图阶段（--no-fig）")
        return False
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


def _build(template: str, work_dir: Optional[str] = None) -> str:
    """跑一遍 build，返回成品 docx 绝对路径。

    work_dir 为 None 时走旧的全局目录（MD 路线）；
    传入时用哈希工作目录（Word 路线），chap_dir/output_dir/toc_pages/images_base
    都指向工作目录，cover_path 指向工作目录的 cover.json，避免多论文共享造成污染。
    """
    from docxbuild.cli import build as build_fn
    if work_dir:
        out = build_fn(
            template=template,
            chap_dir=os.path.join(work_dir, "章节"),
            output_dir=os.path.join(work_dir, "output"),
            toc_pages=os.path.join(work_dir, "toc.json"),
            images_base=work_dir,
            cover_path=os.path.join(work_dir, "cover.json"),
        )
    else:
        out = build_fn(template=template)
    return out


def _toc_stage(product: str, template: str, no_toc: bool,
               work_dir: Optional[str] = None) -> bool:
    """自动转 PDF、量页码并第二遍构建。返回是否完成页码回填。"""
    if no_toc:
        print("已跳过目录页码回填（--no-toc）；在 Word 中按 Ctrl+A 后 F9 可更新目录")
        return False
    missing = [t for t in ("soffice", "pdftotext") if shutil.which(t) is None]
    if missing:
        print(f"[注意] 缺少 {', '.join(missing)}，跳过目录页码回填（不影响出稿）。")
        print("       安装后重跑本命令可自动回填；或在 Word 中按 Ctrl+A 后 F9 更新目录。")
        return False

    # PDF 落在成品旁边（工作目录模式下进工作目录的 output/），
    # 页码表与章节目录同样指向工作目录，不碰全局位置（防止多论文污染）
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
    if work_dir:
        toc_main([pdf, "--chap-dir", os.path.join(work_dir, "章节"),
                  "--out", os.path.join(work_dir, "toc.json")])
    else:
        toc_main([pdf])

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


def _convert_stage(docx: str, work_dir: str, keep_front: bool) -> None:
    """拆解用户 docx：解析 → src.md + 封面 cover.json + 章节/*.md。

    与 GUI 审阅页共用同一份块处理链路（prepare_blocks →
    write_chapters_from_blocks），保证命令行与 GUI 的拆解产物完全一致。
    """
    from docxbuild.docinfo import save_cover_to
    from docxconvert.cli import (extract_cover, extract_template_title,
                                 prepare_blocks)
    from docxconvert.markdown import render

    # rel_base 传 work_dir：图片引用写成 images/x.png，
    # build 时 images_base=work_dir，与工作目录结构匹配
    images_dir = os.path.join(work_dir, "images")
    blocks, _, _ = prepare_blocks(docx, images_dir,
                                  strip_front=not keep_front,
                                  rel_base=work_dir)
    with open(os.path.join(work_dir, "src.md"), "w", encoding="utf-8") as f:
        f.write(render(blocks))

    # 提取封面字段并存到工作目录的 cover.json（每个论文独立，不污染全局）
    cover = extract_cover(docx)
    tpl_title = extract_template_title(docx)
    save_cover_to(os.path.join(work_dir, "cover.json"), {
        "title": cover.get("成果名称", ""),
        "version": VERSION,
        "template_title": tpl_title,
        "cover": cover,
    })
    print(f"  已提取封面字段 {len(cover)} 项 → cover.json"
          f"（承诺书题目：{tpl_title or '未识别'}）")

    # 按 H1 拆分写入工作目录的 章节/（先清理旧分章，与 GUI 同一实现）
    written = write_chapters_from_blocks(
        blocks, chap_dir=os.path.join(work_dir, "章节"))
    print(f"  已写入 {len(written)} 个章节 → "
          f"{os.path.relpath(os.path.join(work_dir, '章节'), HERE)}/")


def run_pipeline(docx: Optional[str] = None,
                 work_dir: Optional[str] = None,
                 template: Optional[str] = None,
                 force: bool = False,
                 keep_front: bool = False,
                 no_fig: bool = False,
                 no_toc: bool = False,
                 no_audit: bool = False,
                 copy_to_output: bool = True) -> str:
    """全流程唯一编排实现：从章节源（或用户 docx）到成品 Word。

    命令行 run / GUI 构建 / 闭环测试统一调用本函数，区别只在参数：

    - docx=None, work_dir=None：MD 路线，直接用 config/章节/
    - docx 传入：Word 路线，先拆解到哈希工作目录（work_dir 未指定时按
      内容哈希自动创建）再构建
    - docx=None, work_dir 传入：复用工作目录模式（GUI 审阅后构建、
      `run <哈希>`），跳过拆解直接构建
    - template=None 时用默认模板 config/template.docx；闭环测试传
      输入文件本身，验证拆解→重建的无损性

    返回成品 docx 绝对路径。
    """
    template_explicit = template is not None
    template = os.path.abspath(template) if template else DEFAULT_TEMPLATE

    # 工作目录优先级：显式 work_dir > docx 哈希自动创建 > None（MD 路线）
    if docx and not work_dir:
        work_dir = prepare_work_dir(docx)
        print(f"工作目录（哈希隔离）：{os.path.relpath(work_dir, HERE)}")
    elif work_dir:
        print(f"工作目录（指定）：{os.path.relpath(work_dir, HERE)}")

    md_target = os.path.join(work_dir, "src.md") if (docx and work_dir) else None

    _step("阶段 · 开工前检查")
    _precheck(template, bool(docx), md_target, force, work_dir)
    print("检查通过。")

    if docx:
        _step("阶段 · 拆解用户 Word 为章节 Markdown")
        _convert_stage(docx, work_dir, keep_front)

    _step("阶段 · 架构图渲染")
    _fig_stage(no_fig)

    _step("阶段 · 构建 Word（第一遍）")
    product = _build(template, work_dir)

    _step("阶段 · 目录页码回填")
    _toc_stage(product, template, no_toc, work_dir)

    if not no_audit:
        _audit_stage(product, template, template_explicit)

    # Word 路线 / 复用工作目录：把成品从工作目录复制到 output/，方便用户取用
    if copy_to_output and work_dir and os.path.exists(product):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        final = os.path.join(OUTPUT_DIR, os.path.basename(product))
        shutil.copy2(product, final)
        product = final

    _step("完成")
    print(f"成品：{os.path.relpath(product, HERE)}")
    return product


def main(argv: Optional[List[str]] = None) -> int:
    """命令行入口：只做参数解析与哈希/路径判定，流程一律交 run_pipeline。"""
    ap = argparse.ArgumentParser(
        description="一键流水线：从 config/章节/*.md（或传入的用户 docx）"
                    "端到端生成成品 Word，内部自动完成架构图、构建、页码回填与审计。")
    ap.add_argument("docx", nargs="?", default=None,
                    help="用户已写好的论文 docx 路径，或 16 位哈希（复用已有"
                         "工作目录，等价 --work-dir）；不传则使用 config/章节/")
    ap.add_argument("--template", default=None,
                    help="外部模板 docx 路径（默认 config/template.docx）；"
                         "指定后构建与审计都使用该模板")
    ap.add_argument("--force", action="store_true",
                    help="Word 路线下允许覆盖工作目录中已存在的拆解产物")
    ap.add_argument("--keep-front", action="store_true",
                    help="Word 路线下保留 docx 原有前置页（默认剥离封面/承诺书/目录，"
                         "因模板自带这些部分）")
    ap.add_argument("--no-fig", action="store_true", help="跳过架构图渲染")
    ap.add_argument("--no-toc", action="store_true", help="跳过目录页码自动回填（不转 PDF）")
    ap.add_argument("--no-audit", action="store_true", help="跳过成品格式审计")
    ap.add_argument("--work-dir", default=None,
                    help="指定哈希工作目录（GUI 审阅后构建用，跳过拆解，"
                         "直接从该目录的 章节/ 构建）")
    args = ap.parse_args(argv)

    # 位置参数两种形态（全局规范）：文件路径 或 16 位哈希
    docx = args.docx
    work_dir = args.work_dir
    if docx and is_hash16(docx):
        h = docx.lower()
        if not work_dir:
            work_dir = resolve_hash(h)
        if not work_dir:
            raise SystemExit(f"哈希 {h} 对应的工作目录不存在，请先传入文件路径")
        # 哈希模式等价于 --work-dir：跳过拆解，直接从工作目录的 章节/ 构建
        docx = None
    elif docx and not os.path.isfile(docx):
        raise SystemExit(f"输入不存在：{docx}（不是文件路径；若是哈希，"
                         "对应工作目录不存在，请先传入文件路径）")

    product = run_pipeline(docx=docx, work_dir=work_dir,
                           template=args.template, force=args.force,
                           keep_front=args.keep_front, no_fig=args.no_fig,
                           no_toc=args.no_toc, no_audit=args.no_audit)
    if docx:
        print(f"拆解产物在工作目录 {os.path.relpath(prepare_work_dir(docx), HERE)} 中，"
              "修改章节后可用 `python3 main.py run <哈希>` 重跑。")
    return 0 if product else 1

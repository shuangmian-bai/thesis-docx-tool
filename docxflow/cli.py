"""一键流水线：`run` 子命令的编排实现。

面向用户只有一条命令、两条路径：

    python3 main.py run                 # MD 路线：config/章节/*.md → 成品 docx
    python3 main.py run 用户论文.docx    # Word 路线：先拆解入 config/章节/，其余同 MD 路线

内部自动串联以下阶段，用户不需要手工分步运行：

    1. 前置检查：模板、签名图、封面信息、章节源，缺件一次性报齐
    2. 拆解（仅 Word 路线）：docx → config/章节/<名>.md，图片 → config/images/<名>_images/
    3. 架构图：config/figures/ 有 DSL 定义时自动渲染 PNG（无则跳过）
    4. 构建：build 第一遍，生成占位目录的成品 docx
    5. 页码回填：本机有 soffice 与 pdftotext 时自动 docx→PDF→量页码→第二遍构建；
       缺任一则跳过（Word 中 Ctrl+A 后 F9 也可更新目录），不阻断出稿
    6. 审计：对最终成品跑 audit，报告照打，审计结论不改变流水线成败

分步子命令（fig/build/toc/audit/convert）仍保留，供高级用户单步调试；
新增任何需要多步完成的能力，必须在本编排层接进来，不得把串联负担丢给用户
（子项目 CLAUDE.md 项目特定规则）。
"""
import argparse
import datetime
import importlib.util
import os
import shutil
import subprocess
from typing import List, Optional

from docxflow import HERE
from docxbuild.cli import CHAP_DIR, OUTPUT_DIR
from docxbuild.docinfo import VERSION
from docxconvert.cli import convert_to
from docxfig.parser import list_figures

# 编排层固定路径（与各原子模块中的定义同值；集中在此供前置检查使用）
CONFIG_DIR = os.path.join(HERE, "config")
FIG_DIR = os.path.join(CONFIG_DIR, "figures")
CONFIG_IMAGES = os.path.join(CONFIG_DIR, "images")
SIGNATURE = os.path.join(CONFIG_IMAGES, "signature.png")
COVER_JSON = os.path.join(CONFIG_DIR, "cover.json")
DEFAULT_TEMPLATE = os.path.join(CONFIG_DIR, "template.docx")

# soffice 使用独立用户配置目录，避免与用户正在运行的 LibreOffice 抢 profile
LO_PROFILE = "file:///tmp/thesis-docx-tool-lo-profile"


def _step(title: str) -> None:
    """打印阶段标题。"""
    print(f"\n{'=' * 30} {title} {'=' * 30}")


def _precheck(template: str, have_docx: bool, md_target: Optional[str],
              force: bool) -> None:
    """开工前一次性列出所有缺失件，避免跑到半途才报错。"""
    problems: List[str] = []
    if not os.path.exists(template):
        problems.append(f"模板不存在：{os.path.relpath(template, HERE)}"
                        f"（放 config/template.docx 或用 --template 指定）")
    if not os.path.exists(SIGNATURE):
        problems.append("承诺书签名图不存在：config/images/signature.png（请自备）")
    if not have_docx:
        if not os.path.isdir(CHAP_DIR) or not [
                f for f in os.listdir(CHAP_DIR) if f.endswith(".md")]:
            problems.append("找不到章节 Markdown：config/章节/ 下没有 .md 文件"
                            "（直接写 MD，或用 `run 用户论文.docx` 传入 Word）")
    elif md_target and os.path.exists(md_target) and not force:
        problems.append(f"拆解目标已存在：{os.path.relpath(md_target, HERE)}"
                        f"（确认覆盖请加 --force）")
    if problems:
        msg = "\n  ".join(problems)
        raise SystemExit(f"开工前检查未通过：\n  {msg}")

    # cover.json 缺失不阻断：占位符也能出稿，只是封面字段待填
    if not os.path.exists(COVER_JSON):
        print("[注意] config/cover.json 不存在，封面字段将使用占位符"
              "（可复制 config/cover.example.json 后填写）")


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


def _build(template: str) -> str:
    """跑一遍 build，返回成品 docx 绝对路径。"""
    from docxbuild.cli import main as build_main
    build_main(["--template", template])
    today = datetime.date.today().strftime("%Y%m%d")
    return os.path.join(OUTPUT_DIR, f"论文_{VERSION}_{today}.docx")


def _toc_stage(product: str, template: str, no_toc: bool) -> bool:
    """自动转 PDF、量页码并第二遍构建。返回是否完成页码回填。"""
    if no_toc:
        print("已跳过目录页码回填（--no-toc）；在 Word 中按 Ctrl+A 后 F9 可更新目录")
        return False
    missing = [t for t in ("soffice", "pdftotext") if shutil.which(t) is None]
    if missing:
        print(f"[注意] 缺少 {', '.join(missing)}，跳过目录页码回填（不影响出稿）。")
        print("       安装后重跑本命令可自动回填；或在 Word 中按 Ctrl+A 后 F9 更新目录。")
        return False

    print("正在用 LibreOffice 把成品转为 PDF ...")
    proc = subprocess.run(
        ["soffice", f"-env:UserInstallation={LO_PROFILE}", "--headless",
         "--convert-to", "pdf", "--outdir", OUTPUT_DIR, product],
        capture_output=True, text=True, timeout=300)
    pdf = os.path.splitext(product)[0] + ".pdf"
    if proc.returncode != 0 or not os.path.exists(pdf):
        print(f"[注意] docx 转 PDF 失败，跳过页码回填：{proc.stderr.strip() or proc.stdout.strip()}")
        return False

    from docxbuild.toc_pages import main as toc_main
    toc_main([pdf])

    print("页码表已生成，第二遍构建以回填目录 ...")
    # 第二遍产物路径与第一遍相同（版本号+日期），直接覆盖占位目录版本
    from docxbuild.cli import main as build_main
    build_main(["--template", template])
    return True


def _audit_stage(product: str) -> None:
    """成品审计：报告照打，任何问题都不影响流水线已产出的成品。"""
    _step("阶段 · 成品格式审计")
    try:
        from docxaudit.cli import main as audit_main
        audit_main(["audit", "--product", product])
    except SystemExit as e:
        # audit 以退出码 1 表达「有注意项」，对流水线而言只是报告，不当失败
        if e.code:
            print("[注意] 审计存在需关注项，见上方报告；成品不受影响，可按报告逐项确认。")
    except Exception as e:  # 审计是只读附加环节，不因其自身异常拖垮整条流水线
        print(f"[注意] 审计阶段异常，已跳过：{e}")


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="一键流水线：从 config/章节/*.md（或传入的用户 docx）"
                    "端到端生成成品 Word，内部自动完成架构图、构建、页码回填与审计。")
    ap.add_argument("docx", nargs="?", default=None,
                    help="用户已写好的论文 docx：先自动拆解入 config/章节/ 再生成；"
                         "不传则直接使用 config/章节/ 下现有 Markdown")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE,
                    help="模板 docx 路径（默认 config/template.docx）")
    ap.add_argument("--force", action="store_true",
                    help="Word 路线下允许覆盖 config/章节/ 中已存在的同名拆解产物")
    ap.add_argument("--keep-front", action="store_true",
                    help="Word 路线下保留 docx 原有前置页（默认剥离封面/承诺书/目录，"
                         "因模板自带这些部分）")
    ap.add_argument("--no-fig", action="store_true", help="跳过架构图渲染")
    ap.add_argument("--no-toc", action="store_true", help="跳过目录页码自动回填（不转 PDF）")
    ap.add_argument("--no-audit", action="store_true", help="跳过成品格式审计")
    args = ap.parse_args(argv)

    template = os.path.abspath(args.template)

    # Word 路线的拆解落点：章节源与图片都归 config/，build 以 config/ 为根解析图片相对路径
    md_target = None
    images_dir = None
    if args.docx:
        base = os.path.splitext(os.path.basename(args.docx))[0]
        md_target = os.path.join(CHAP_DIR, f"{base}.md")
        images_dir = os.path.join(CONFIG_IMAGES, f"{base}_images")

    _step("阶段 · 开工前检查")
    _precheck(template, bool(args.docx), md_target, args.force)
    print("检查通过。")

    if args.docx:
        _step("阶段 · 拆解用户 Word 为章节 Markdown")
        convert_to(args.docx, md_target, images_dir,
                   strip_front=not args.keep_front)

    _step("阶段 · 架构图渲染")
    _fig_stage(args.no_fig)

    _step("阶段 · 构建 Word（第一遍）")
    product = _build(template)

    _step("阶段 · 目录页码回填")
    _toc_stage(product, template, args.no_toc)

    if not args.no_audit:
        _audit_stage(product)

    _step("完成")
    print(f"成品：{os.path.relpath(product, HERE)}")
    if args.docx:
        print("拆解得到的章节 Markdown 在 config/章节/ 中，可直接修改，"
              "之后重跑 `python3 main.py run`（不带 docx）即可重新出稿。")
    return 0

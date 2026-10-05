"""闭环测试命令行入口（开发调试工具，非面向用户入口）。

用法：
    # 用文件路径跑闭环（自动计算哈希并创建隔离工作目录）
    python3 -m docxloop.cli 论文.docx

    # 用哈希跑闭环（直接定位已有工作目录，不重新 convert→build）
    python3 -m docxloop.cli abc123def4567890

    # 输入文件夹（批量，每个文件独立哈希隔离）
    python3 -m docxloop.cli ./papers/

    # 强制重新构建（不复用已有产物）
    python3 -m docxloop.cli 论文.docx --no-reuse
"""
import os
import sys


def main(argv=None):
    argv = argv or sys.argv[1:]
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    import argparse
    ap = argparse.ArgumentParser(description="闭环测试引擎（哈希隔离，支持多线程）")
    ap.add_argument("input", help="输入 docx 文件 / 文件夹 / 16位哈希")
    ap.add_argument("--template", default=None,
                    help="模板 docx 路径；默认每个输入文件自己作自己的模板"
                         "（闭环验证拆解→重建的无损性，模板与输入同源）")
    ap.add_argument("--output", default=None,
                    help="HTML 报告输出路径（默认 .cache/loop_report.html）")
    ap.add_argument("--no-reuse", action="store_true",
                    help="强制重新 convert→build，不复用已有工作目录产物")
    args = ap.parse_args(argv)

    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    template = os.path.abspath(args.template) if args.template else None
    if template and not os.path.exists(template):
        print(f"模板不存在：{template}", file=sys.stderr)
        return 1

    from docxloop.engine import resolve_input, run_batch

    # 先判断输入是否可解析
    from docxflow.workdir import is_hash16, resolve_hash
    if not resolve_input(args.input) and not os.path.isdir(args.input):
        # 可能是哈希但工作目录不存在，或文件不存在
        if is_hash16(args.input):
            print(f"哈希 {args.input} 对应的工作目录不存在，请先传入文件路径",
                  file=sys.stderr)
        else:
            print(f"输入不存在：{args.input}", file=sys.stderr)
        return 1

    print(f"输入: {args.input}")
    print(f"模板: {template or '每个输入文件本身'}")
    results = run_batch([args.input], template, reuse=not args.no_reuse)

    if not results:
        print("未找到可处理的 docx 文件", file=sys.stderr)
        return 1

    print(f"\n完成 {len(results)} 个文件的闭环测试")
    for r in results:
        status = "通过" if r.passed else "未通过"
        print(f"  [{status}] {os.path.basename(r.src_path)}: "
              f"{r.total} 处差异（{r.unexpected} 非预期 / {r.total - r.unexpected} 预期）")
        # 人工验收入口：给出参与对比的两份 Word 完整路径，可直接打开逐项核对
        print(f"      原文: {r.src_path}")
        print(f"      成品: {r.out_path}")

    out_path = args.output or os.path.join(here, ".cache", "loop_report.html")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    from docxloop.report import render
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(render(results, title=f"闭环测试报告 - {os.path.basename(args.input)}"))
    print(f"报告已生成: {out_path}")

    return 0 if all(r.passed for r in results) else 1


if __name__ == "__main__":
    sys.exit(main())

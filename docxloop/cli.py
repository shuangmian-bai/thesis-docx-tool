"""闭环测试命令行入口：word 扁平化 → 合成 → 严格对比。

用法：
    python3 -m docxloop.cli 论文.docx        # 单文件
    python3 -m docxloop.cli ./papers/        # 文件夹批量
    python3 -m docxloop.cli 论文.docx --no-reuse   # 强制重新扁平化+合成
"""
import os
import sys


def main(argv=None):
    argv = argv or sys.argv[1:]
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    import argparse
    ap = argparse.ArgumentParser(description="闭环测试：word 扁平化→合成→严格对比")
    ap.add_argument("input", help="输入 docx 文件 / 文件夹")
    ap.add_argument("--no-reuse", action="store_true",
                    help="强制重新扁平化+合成（改了扁平化器/渲染器后用）")
    args = ap.parse_args(argv)

    from docxloop.engine import run_loop
    results = run_loop(args.input, reuse=not args.no_reuse)
    if not results:
        print("未找到可处理的 docx 文件", file=sys.stderr)
        return 1

    for r in results:
        status = "通过" if r.passed else "未通过"
        print(f"  [{status}] {os.path.basename(r.src_path)}: "
              f"{r.total} 处差异")
        print(f"      原文: {r.src_path}")
        print(f"      成品: {r.out_path}")

    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_path = os.path.join(here, ".cache", "loop_report.html")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    from docxloop.report import render
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(render(results, title="闭环测试报告"))
    print(f"报告已生成: {out_path}")

    return 0 if all(r.passed for r in results) else 1


if __name__ == "__main__":
    sys.exit(main())

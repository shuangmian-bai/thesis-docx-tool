#!/usr/bin/env python3
"""校验成品 docx 是否严格贴合模板格式。

论文构建链的最后一环：`build_docx.py` 套模板生成 docx，`toc_pages.py` 回填目录页码，
本脚本负责**核对成品有没有跑偏**。做法是把成品与模板逐部件、逐样式、逐段落地比，
再把差异分成三类：通过、**有意偏离**（登记在 `EXPECTED_DIFFS` 表里，见 `README.md`
的「格式约定」）、需关注。

差异只在报告层降级：检查函数一律报 `WARN`，`EXPECTED_DIFFS` 表命中才改判为有意偏离。
新增一处有意偏离时只改表，不动检查逻辑。

用法（内部模块，由 run 编排调用；独立调试用 python3 -m docxaudit.cli）：

    python3 -m docxaudit.cli                    # 等同 all，给总览
    python3 -m docxaudit.cli parts styles       # 只跑指定项
    python3 -m docxaudit.cli --list             # 列出全部检查项
    python3 -m docxaudit.cli -v body            # 展开通过项的细节
    python3 -m docxaudit.cli --product 论文_v1_20260916.docx    # 指定成品

退出码：有「需关注」或「失败」项时为 1；`--strict` 时连「有意偏离」也算失败，
便于接进提交前钩子。
"""
import os
import sys
from collections import Counter

from docxaudit.checks import CHECKS, CHECK_MAP
from docxaudit.common import (
    EXPECTED, FAIL, OK, WARN, Ctx, find_product, find_template, fmt_ts,
)
from docxaudit.report import render, run_check


def main(argv):
    verbose = strict = False
    product = template = None
    wanted = []

    args = argv[1:]
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("-h", "--help"):
            print(__doc__)
            return 0
        elif a == "--list":
            print("检查项：")
            for name, _, _, desc in CHECKS:
                print(f"  {name:<8} {desc}")
            print(f"  {'all':<8} 依次跑完以上全部，给总览")
            return 0
        elif a in ("-v", "--verbose"):
            verbose = True
        elif a == "--strict":
            strict = True
        elif a in ("--product", "--template"):
            i += 1
            if i >= len(args):
                raise SystemExit(f"{a} 需要一个路径参数")
            if a == "--product":
                product = args[i]
            else:
                template = args[i]
        elif a.startswith("-"):
            raise SystemExit(f"未知参数：{a}（--help 看用法）")
        else:
            wanted.append(a)
        i += 1

    for name in wanted:
        if name != "all" and name not in CHECK_MAP:
            raise SystemExit(f"未知检查项：{name}（--list 看全部）")

    tpl_path = find_template(template)
    prod_path = find_product(product)
    # 换过模板后，针对原模板登记的预期偏离不再适用，一律当作真差异看
    ctx = Ctx(tpl_path, prod_path, use_expected=(template is None))
    if template:
        print("注：指定了 --template，已停用预期偏离表（登记项只对原模板成立）\n")

    print(f"模板：{tpl_path}")
    print(f"成品：{prod_path}（{fmt_ts(os.path.getmtime(prod_path))}）")
    print("=" * 74)

    names = [n for n, _, _, _ in CHECKS] if (not wanted or "all" in wanted) else wanted
    totals = Counter()
    for name in names:
        findings = run_check(name, ctx)
        totals.update(f.level for f in findings)
        print(f"\n【{CHECK_MAP[name][0]}】")
        print("\n".join(render(findings, verbose)))

    print("\n" + "=" * 74)
    print(f"共 {sum(totals.values())} 项："
          f"[通过] {totals[OK]} / "
          f"[有意偏离] {totals[EXPECTED]} / "
          f"[注意] {totals[WARN]} 需关注 / "
          f"[失败] {totals[FAIL]}")
    if totals[EXPECTED] and not verbose:
        print("（有意偏离的判据见 README「格式约定」，"
              "以及 docxaudit/expected.py 的 EXPECTED_DIFFS 表）")

    ctx.close()
    bad = totals[WARN] + totals[FAIL] + (totals[EXPECTED] if strict else 0)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))

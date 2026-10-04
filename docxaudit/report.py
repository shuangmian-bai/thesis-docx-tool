"""报告层：结论分级、渲染、单项执行。

`classify` 是「有意偏离」唯一的判定点——检查函数一律报 `WARN`，命中
`EXPECTED_DIFFS` 才在这里改判。所以新增一处有意偏离只改表，不动检查逻辑。
"""
from dataclasses import replace

from docxaudit.checks import CHECK_MAP
from docxaudit.common import EXPECTED, FAIL, LEVEL_MARK, WARN, Finding
from docxaudit.expected import EXPECTED_DIFFS


def classify(findings, use_expected):
    """把登记在 `EXPECTED_DIFFS` 表里的差异降级为「有意偏离」。

    差异只在报告层降级，检查函数一律报 WARN——新增一处有意偏离时只改表，
    不必回头动检查逻辑。
    """
    if not use_expected:
        return findings
    out = []
    for f in findings:
        if f.level == WARN and f.key in EXPECTED_DIFFS:
            out.append(replace(f, level=EXPECTED,
                               msg=f"{f.msg}｜{EXPECTED_DIFFS[f.key]}"))
        else:
            out.append(f)
    return out


def render(findings, verbose):
    """把 Findings 渲染成文本行。通过项只占一行，问题项自动展开 detail。"""
    lines = []
    for f in findings:
        lines.append(f"  {LEVEL_MARK[f.level]} {f.msg}")
        if f.detail and (verbose or f.level in (WARN, FAIL)):
            lines.extend(f"      {ln}" for ln in f.detail.splitlines())
    return lines


def run_check(name, ctx):
    """跑一项检查。单项出错不该拖垮整轮——捕获后报成一条失败。"""
    _, fn, _ = CHECK_MAP[name]
    try:
        return classify(fn(ctx), ctx.use_expected)
    except Exception as exc:                                  # noqa: BLE001
        return [Finding(f"{name}.crash", FAIL, f"检查项 {name} 执行失败：{exc!r}")]

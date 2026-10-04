"""前置警示：成品是不是比章节源还旧。"""
import os

from docxaudit.common import CHAP_DIR, OK, WARN, Finding, fmt_ts, rel_delta


def check_stale(ctx):
    """成品是不是比章节源还旧。

    README 警告过：跨零点重跑会写出新文件名，若渲染时用了旧稿，会「拿旧稿排版而
    不自知」。这条比任何格式比对都该先跑——比对得再细，比的是旧稿也没意义。
    """
    out = []
    prod_m = os.path.getmtime(ctx.prod_path)
    newest, newest_m = None, 0.0
    for fn in os.listdir(CHAP_DIR):
        if not fn.endswith(".md"):
            continue
        m = os.path.getmtime(os.path.join(CHAP_DIR, fn))
        if m > newest_m:
            newest, newest_m = fn, m
    if newest is None:
        return [Finding("stale", WARN, f"{CHAP_DIR} 下没有 .md 源文件")]

    stamp = f"成品 {fmt_ts(prod_m)}"
    if newest_m > prod_m:
        out.append(Finding(
            "stale.newer_source", WARN,
            f"章节源比成品新 {rel_delta(newest_m - prod_m)}，成品可能是旧稿",
            f"最新源：{newest}（{fmt_ts(newest_m)}）\n{stamp}\n"
            f"改完正文要重跑 build_docx.py 再比对，否则比的是旧稿"))
    else:
        out.append(Finding("stale", OK,
                           f"{stamp}，晚于全部章节源（最近改动 {newest}）"))
    return out

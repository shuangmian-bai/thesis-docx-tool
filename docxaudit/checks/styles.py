"""样式层：逐样式差异、docDefaults、latentStyles、TOC 大小写、使用统计。"""
import re
from collections import Counter

from docxaudit.common import (
    DEFAULTS_RE, FAIL, LATENT_RE, OK, STYLE_RE, WARN,
    Finding, iter_paras, pstyle, tag_counter,
)


def check_styles(ctx):
    """样式定义与使用：逐样式差异、docDefaults、latentStyles、TOC 大小写、使用统计。"""
    out = []
    xt, xp = ctx.styles_xml("tpl"), ctx.styles_xml("prod")
    st = {m.group(1): m.group(0) for m in STYLE_RE.finditer(xt)}
    sp = {m.group(1): m.group(0) for m in STYLE_RE.finditer(xp)}

    miss, extra = sorted(set(st) - set(sp)), sorted(set(sp) - set(st))
    if miss:
        out.append(Finding("styles.missing", FAIL, f"成品缺少样式定义：{miss}"))
    if extra:
        out.append(Finding("styles.extra", WARN, f"成品多出样式定义：{extra}"))
    if not miss and not extra:
        out.append(Finding("styles.ids", OK, f"{len(sp)} 个样式 ID 与模板完全一致"))

    dt, dp = DEFAULTS_RE.search(xt), DEFAULTS_RE.search(xp)
    if dt and dp and dt.group(0) == dp.group(0):
        s = re.search(r'<w:sz w:val="(\d+)"', dt.group(0))
        out.append(Finding("styles.docdefaults", OK,
                           f"docDefaults 一致（默认字号 {s.group(1) if s else '?'} 半磅）"))
    else:
        out.append(Finding("styles.docdefaults", WARN, "docDefaults 不一致",
                           f"模板：{(dt.group(0) if dt else '—')[:300]}\n"
                           f"成品：{(dp.group(0) if dp else '—')[:300]}"))

    lt, lp = LATENT_RE.search(xt), LATENT_RE.search(xp)
    if lt and lp and lt.group(0) == lp.group(0):
        out.append(Finding("styles.latent", OK, "latentStyles 一致"))
    else:
        out.append(Finding("styles.latent", WARN, "latentStyles 不一致"))

    changed = [sid for sid in sorted(set(st) & set(sp)) if st[sid] != sp[sid]]
    toc_sids = [s for s in changed if s in ("TOC1", "TOC2")]
    other = [s for s in changed if s not in ("TOC1", "TOC2")]
    if toc_sids:
        detail = []
        for sid, tag in (("TOC1", "caps"), ("TOC2", "smallCaps")):
            a = bool(re.search(rf"<w:{tag}\s*/>", st.get(sid, "")))
            b = bool(re.search(rf"<w:{tag}\s*/>", sp.get(sid, "")))
            detail.append(f"{sid}.{tag}：模板 {'有' if a else '无'} → 成品 {'有' if b else '无'}")
        detail.append(f"styles.xml 字符数：模板 {len(xt)} / 成品 {len(xp)}（差 {len(xt) - len(xp)}）")
        out.append(Finding("styles.toc", WARN,
                           f"TOC 样式的强制大小写标签被移除（{', '.join(toc_sids)}）",
                           "\n".join(detail)))
    if other:
        out.append(Finding("styles.changed", WARN,
                           f"{len(other)} 个样式定义与模板不同：{other}",
                           "\n\n".join(
                               f"── {sid} ──\n模板多出：{dict(tag_counter(st[sid]) - tag_counter(sp[sid]))}\n"
                               f"成品多出：{dict(tag_counter(sp[sid]) - tag_counter(st[sid]))}"
                               for sid in other)))
    if not changed:
        out.append(Finding("styles.defs", OK, "全部样式定义逐字节一致"))

    # ── 段落样式使用统计（两篇论文内容不同，只看「有没有用模板没有的样式」）──
    ct = Counter(s for s in (pstyle(p) for p in iter_paras(ctx.doc_xml("tpl"))) if s)
    cp = Counter(s for s in (pstyle(p) for p in iter_paras(ctx.doc_xml("prod"))) if s)
    used_no_def = sorted(set(cp) - set(sp))
    if used_no_def:
        out.append(Finding("styles.undefined", FAIL,
                           f"成品用了没有定义的样式：{used_no_def}"))
    else:
        out.append(Finding("styles.used", OK,
                           f"成品用到 {len(cp)} 种段落样式，全部有定义"
                           f"（模板 {len(ct)} 种）",
                           "\n".join(f"{k:<6} 模板 {ct.get(k, 0):>4} 段 / 成品 {cp[k]:>4} 段"
                                     for k in sorted(cp, key=lambda k: -cp[k]))))
    return out

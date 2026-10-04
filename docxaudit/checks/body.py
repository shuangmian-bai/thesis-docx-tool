"""正文层：保留区骨架、分块、分节符、页面设置、表格分页、直接格式化、按样式分组。"""
import re
from collections import Counter

from docxaudit.common import (
    OK, WARN, Finding, iter_blocks, iter_paras, norm, pstyle,
    reserved_parts, skeleton, tail_sectpr,
)


def check_body(ctx):
    """正文层：保留区骨架、分块、分节符、页面设置、表格分页、直接格式化、按样式分组。"""
    out = []
    rt, rp = ctx.doc_xml("tpl"), ctx.doc_xml("prod")
    at, mt, ct_ = reserved_parts(rt)
    ap, mp, cp_ = reserved_parts(rp)

    # ── 保留区：比结构骨架，不比文字（封面字段、目录条目本就与模板不同）──
    bad = [lbl for lbl, x, y in (("目录之前（封面/承诺书）", at, ap),
                                 ("分节符段落", ct_, cp_))
           if skeleton(x) != skeleton(y)]
    if bad:
        detail = []
        for lbl, x, y in (("目录之前（封面/承诺书）", at, ap), ("分节符段落", ct_, cp_)):
            if lbl not in bad:
                continue
            bt_, bp_ = iter_blocks(x), iter_blocks(y)
            detail.append(f"── {lbl} ── 模板 {len(bt_)} 块 / 成品 {len(bp_)} 块")
            for i, (u, v) in enumerate(zip(bt_, bp_)):
                if skeleton(u) != skeleton(v):
                    detail.append(f"   第 {i} 块不同\n"
                                  f"     模板：{skeleton(u)[:180]}\n"
                                  f"     成品：{skeleton(v)[:180]}")
                    break
        out.append(Finding("body.reserved.structure", WARN,
                           f"保留区结构与模板不同：{'、'.join(bad)}", "\n".join(detail)))
    else:
        out.append(Finding("body.reserved", OK,
                           f"保留区结构与模板一致（封面/承诺书 {len(iter_blocks(at))} 块 + "
                           f"分节符段；目录 {len(iter_blocks(mt))} 块按论文重写，不参与比对）"))

    # ── 末尾 sectPr 与页面设置 ──
    st_, sp_ = tail_sectpr(rt), tail_sectpr(rp)
    if norm(st_) == norm(sp_):
        out.append(Finding("body.sectpr", OK, "body 级 sectPr 与模板一致"))
    else:
        out.append(Finding("body.sectpr", WARN, "body 级 sectPr 与模板不同",
                           f"模板：{st_}\n成品：{sp_}"))

    for lbl, xml in (("模板", st_), ("成品", sp_)):
        m = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"', xml)
        g = re.search(r"<w:pgMar ([^/]*)/>", xml)
        if m and g:
            w_, h_ = int(m.group(1)), int(m.group(2))
            mar = dict(re.findall(r'w:(\w+)="(-?\d+)"', g.group(1)))
            content = w_ - int(mar.get("left", 0)) - int(mar.get("right", 0))
            out.append(Finding(f"body.pagesize.{lbl}", OK,
                               f"{lbl}：A4 {w_}×{h_} twips"
                               f"（{w_ / 1440 * 25.4:.1f}×{h_ / 1440 * 25.4:.1f} mm），"
                               f"正文宽 {content} twips"))
        else:
            out.append(Finding(f"body.pagesize.{lbl}", WARN, f"{lbl}：sectPr 里找不到页面尺寸"))

    # ── 表格不跨页（只查正文表格，保留区的学籍表不归本检查管）──
    body_start = rp.rindex(cp_) + len(cp_)
    n_tbl, bad_tbl = _table_keep(iter_blocks(rp[body_start:]))
    if bad_tbl:
        out.append(Finding("body.table.keep", WARN,
                           f"{len(bad_tbl)} / {n_tbl} 个正文表格缺「不跨页」标记",
                           "\n".join(bad_tbl)))
    else:
        out.append(Finding("body.table.keep", OK,
                           f"{n_tbl} 个正文表格均带「不跨页」标记（表题与整表绑为一块）"))

    # ── 直接格式化审计 ──
    ht, hp = _direct_fmt(iter_paras(rt)), _direct_fmt(iter_paras(rp))
    out.append(Finding("body.directfmt", OK,
                       f"正文直接格式化：{len(hp)} 种写法 / {sum(hp.values())} 处"
                       f"（模板 {len(ht)} 种 / {sum(ht.values())} 处）",
                       "\n".join(f"x{v:<5} {k}" for k, v in hp.most_common())))

    new = set(hp) - set(ht)
    jc_both = sorted(k for k in new if "w:jc" in k and '"both"' in k)
    runfont = sorted(k for k in new if "Consolas" in k)
    unexpected = sorted(new - set(jc_both) - set(runfont))
    if jc_both:
        out.append(Finding("body.directfmt.jc", WARN,
                           f"正文 {hp[jc_both[0]]} 处显式两端对齐（jc=both）"))
    if runfont:
        out.append(Finding("body.directfmt.runfont", WARN,
                           f"行内等宽字体 {sum(hp[k] for k in runfont)} 处",
                           "\n".join(f"x{hp[k]:<5} {k}" for k in runfont)))
    if unexpected:
        out.append(Finding("body.directfmt.unexpected", WARN,
                           f"成品引入了 {len(unexpected)} 种模板没有的直接格式写法",
                           "\n".join(f"x{hp[k]:<5} {k}" for k in unexpected)))

    # ── 按样式分组比对直接格式化（同一样式下才可比）──
    tps, pps = _by_style(iter_paras(rt)), _by_style(iter_paras(rp))
    rows, drift = [], []
    for sid in sorted(set(pps) & set(tps), key=lambda s: -len(pps[s])):
        kt_ = {re.match(r"<w:(\w+)", k).group(1) for k in tps[sid]}
        kp_ = {re.match(r"<w:(\w+)", k).group(1) for k in pps[sid]}
        rows.append(f"{sid:<6} 模板 {len(tps[sid]):>3} 段 / 成品 {len(pps[sid]):>3} 段  "
                    f"属性类别 {sorted(kp_)}")
        if kp_ - kt_:
            drift.append(f"{sid}：成品多 {sorted(kp_ - kt_)}，模板多 {sorted(kt_ - kp_)}")
    if drift:
        out.append(Finding("body.bystyle.drift", WARN,
                           f"{len(drift)} 种样式下的直接格式属性超出模板用法",
                           "\n".join(drift)))
    else:
        out.append(Finding("body.bystyle", OK,
                           f"{len(rows)} 种共有样式下，直接格式属性类别均未超出模板用法",
                           "\n".join(rows)))
    return out


def _table_keep(blocks):
    """检查正文表格是否都带「不跨页」标记，返回 (表格数, 问题列表)。

    构建侧的写法见 `docxbuild/fragments.table_xml`：每行 `cantSplit`（行内不断页）、
    除末行外每个单元格段落 `keepNext`（与下一行同页）、表题段落也 `keepNext`。
    这里只复核**标记齐不齐**——标记齐不等于渲染后一定同页（表格高过一页时
    keepNext 会失效），那要看导出的 PDF。
    """
    bad, n = [], 0
    for i, b in enumerate(blocks):
        if not b.startswith("<w:tbl>"):
            continue
        n += 1
        rows = re.findall(r"<w:tr(?:\s[^>]*)?>.*?</w:tr>", b, re.S)
        miss = []
        if any("<w:cantSplit/>" not in r for r in rows):
            miss.append("行缺 cantSplit")
        if any("<w:keepNext/>" not in r for r in rows[:-1]):
            miss.append("行间缺 keepNext")
        prev = blocks[i - 1] if i else ""
        if pstyle(prev) != "aa" or "<w:keepNext/>" not in prev:
            miss.append("表题缺 keepNext")
        if miss:
            bad.append(f"第 {n} 个表：{'、'.join(miss)}")
    return n, bad


def _direct_fmt(paras):
    """统计段落里绕过样式的直接格式：段落级 jc/ind/spacing/shd/bdr，run 级字体字号等。"""
    hits = Counter()
    for p in paras:
        ppr = re.search(r"<w:pPr>(.*?)</w:pPr>", p, re.S)
        if ppr:
            body = re.sub(r"<w:pStyle[^>]*/>", "", ppr.group(1))
            for tag in ("jc", "ind", "spacing", "shd", "bdr"):
                for m in re.finditer(rf"<w:{tag}\b[^>]*/?>", body):
                    hits[f"p.{tag}: {m.group(0)}"] += 1
        for rm in re.finditer(r"<w:rPr>(.*?)</w:rPr>", p, re.S):
            body = re.sub(r"<w:rStyle[^>]*/>", "", rm.group(1))
            for tag in ("b", "bCs", "i", "sz", "szCs", "color", "rFonts", "u", "shd"):
                for m in re.finditer(rf"<w:{tag}\b[^>]*/?>", body):
                    hits[f"r.{tag}: {m.group(0)}"] += 1
    return hits


def _by_style(paras):
    """把段落按其样式 ID 分组，返回 {样式: Counter(直接格式化标签)}。"""
    groups = {}
    for p in paras:
        sid = pstyle(p) or "(无)"
        ppr = re.search(r"<w:pPr>(.*?)</w:pPr>", p, re.S)
        tags = Counter()
        if ppr:
            body = re.sub(r"<w:pStyle[^>]*/>", "", ppr.group(1))
            tags += Counter(m.group(0) for m in
                            re.finditer(r"<w:(jc|ind|spacing|shd|bdr|outlineLvl)\b[^>]*/?>", body))
        for rm in re.finditer(r"<w:rPr>(.*?)</w:rPr>", p, re.S):
            body = re.sub(r"<w:rStyle[^>]*/>", "", rm.group(1))
            tags += Counter(m.group(0) for m in re.finditer(
                r"<w:(b|bCs|i|sz|szCs|color|rFonts|u|shd|vertAlign)\b[^>]*/?>", body))
        groups.setdefault(sid, Counter()).update(tags)
    return groups

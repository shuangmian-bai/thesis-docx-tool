"""属性与设置：core.xml、settings.xml、编号、字体表/主题、命名空间声明。"""
import re

from docxaudit.common import OK, WARN, XMLNS_DECL_RE, Finding, iter_paras, tag_counter


def check_meta(ctx):
    """文档属性与设置：core.xml、settings.xml、编号、字体表/主题、命名空间声明。"""
    out = []

    # ── core.xml ──
    a, b = ctx.tpl.text("docProps/core.xml"), ctx.prod.text("docProps/core.xml")
    fields = ("dc:title", "dc:creator", "cp:lastModifiedBy", "cp:revision",
              "dcterms:created", "dcterms:modified")
    vals = {}
    for tag in fields:
        va = re.search(rf"<{tag}[^>]*>(.*?)</{tag}>", a)
        vb = re.search(rf"<{tag}[^>]*>(.*?)</{tag}>", b)
        vals[tag] = (va.group(1) if va else None, vb.group(1) if vb else None)
    # 只允许著录信息与时间戳随论文变化，其余字段（如 revision）仍须一致
    free = {"dc:title", "dc:creator", "cp:lastModifiedBy",
            "dcterms:created", "dcterms:modified"}
    strict_bad = [t for t in fields if t not in free and vals[t][0] != vals[t][1]]
    ea = set(re.findall(r"<(\w+:[\w-]+)", a))
    eb = set(re.findall(r"<(\w+:[\w-]+)", b))
    changed = [t for t in free if vals[t][0] != vals[t][1]]
    if strict_bad or (ea - eb):
        out.append(Finding("meta.core.structure", WARN,
                           f"core.xml 结构或受限字段有变：{strict_bad or ''} "
                           f"模板独有元素 {sorted(ea - eb)}"))
    elif changed:
        # 只有著录信息与时间戳变了，属预期（本表登记为有意偏离，见 EXPECTED_DIFFS）
        out.append(Finding("meta.core.authorship", WARN, "core.xml 著录信息随论文替换（预期）",
                           "\n".join(f"{t:<20} 模板={vals[t][0]!r} 成品={vals[t][1]!r}"
                                     for t in fields)))
    else:
        out.append(Finding("meta.core", OK, "core.xml 全部字段与模板一致"))

    # ── settings.xml ──
    sa, sb = ctx.tpl.text("word/settings.xml"), ctx.prod.text("word/settings.xml")
    if sa == sb:
        out.append(Finding("meta.settings", OK, "settings.xml 逐字节一致"))
    else:
        ta, tb = tag_counter(sa), tag_counter(sb)
        added, removed = dict(tb - ta), dict(ta - tb)
        uf = re.search(r"<w:updateFields[^>]*/>", sb)
        if removed:
            out.append(Finding("meta.settings.removed", WARN,
                               f"settings.xml 少了模板有的元素：{removed}"))
        if set(added) == {"updateFields"} and uf:
            out.append(Finding("meta.settings.updatefields", WARN,
                               "settings.xml 多一个 <w:updateFields/>（预期）",
                               f"{uf.group(0)}——让 Word 打开时提示重建目录域"))
        elif added:
            out.append(Finding("meta.settings.added", WARN,
                               f"settings.xml 多出未预期的元素：{added}"))

    # ── numbering / fontTable / theme / webSettings ──
    for part in ["word/numbering.xml", "word/fontTable.xml", "word/webSettings.xml"] + \
                sorted(n for n in ctx.tpl.names() if n.startswith("word/theme/")):
        if not (ctx.tpl.has(part) and ctx.prod.has(part)):
            out.append(Finding(f"meta.{part}.missing", WARN, f"{part}：模板有="
                               f"{ctx.tpl.has(part)} 成品有={ctx.prod.has(part)}"))
        elif ctx.tpl.raw(part) != ctx.prod.raw(part):
            out.append(Finding(f"meta.{part}.diff", WARN, f"{part} 与模板不一致"))
    num_same = (ctx.tpl.has("word/numbering.xml") and ctx.prod.has("word/numbering.xml")
                and ctx.tpl.raw("word/numbering.xml") == ctx.prod.raw("word/numbering.xml"))
    n_numpr = sum(1 for p in iter_paras(ctx.doc_xml("prod")) if "<w:numPr>" in p)
    if num_same and n_numpr == 0:
        out.append(Finding("meta.numbering", OK,
                           "numbering.xml 一致，正文未使用编号列表（同模板）"))
    elif not num_same:
        out.append(Finding("meta.numbering.diff", WARN, "numbering.xml 与模板不一致"))

    # ── 命名空间声明分布 ──
    def decls(raw):
        return dict(XMLNS_DECL_RE.findall(raw[raw.index("<w:body"):]))
    dt_, dp_ = decls(ctx.doc_xml("tpl")), decls(ctx.doc_xml("prod"))
    nsnum = len(re.findall(r"xmlns:ns[0-9]+=", ctx.doc_xml("prod")))
    detail = (f"模板：{dt_}\n成品：{dp_}\n"
              f"说明：本模块就地声明的 a / pic / a14 由 ElementTree 提到外层元素上，"
              f"属性值相同、语义等价；模板声明在内层图形元素上。")
    if nsnum:
        out.append(Finding("meta.xmlns.nsnum", WARN,
                           f"成品出现 {nsnum} 处自动编号的 nsN 前缀", detail))
    elif set(dp_) - set(dt_):
        out.append(Finding("meta.xmlns.extra", WARN,
                           f"成品多出命名空间前缀：{sorted(set(dp_) - set(dt_))}", detail))
    else:
        out.append(Finding("meta.xmlns", OK,
                           f"内联命名空间前缀与模板同为 {sorted(dt_)}", detail))
    return out

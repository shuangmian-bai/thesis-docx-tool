"""zip 部件层：清单、逐字节、Content_Types、关系、媒体、页眉页脚、序列化风格。"""
import os
import re

from docxaudit.common import FAIL, OK, WARN, Finding

# 封面用图：不随论文内容变化，模板与成品必须逐字节一致。
FIXED_MEDIA = ("word/media/image1.png",)

# 差异属于「内容随论文变化」而非「格式跑偏」的部件
CONTENT_PARTS = frozenset({
    "word/document.xml",
    "docProps/core.xml",
    "word/settings.xml",
    "word/styles.xml",
    "word/_rels/document.xml.rels",
    "_rels/.rels",
})


def check_parts(ctx):
    """zip 部件层面：清单、逐字节、Content_Types、关系、媒体、页眉页脚、序列化风格。"""
    out = []
    nt, np_ = set(ctx.tpl.names()), set(ctx.prod.names())

    # ── 部件清单 ──
    only_t, only_p = sorted(nt - np_), sorted(np_ - nt)
    tpl_media = [n for n in only_t if n.startswith("word/media/")
                 or n.startswith("docProps/thumbnail")]
    tpl_other = [n for n in only_t if n not in tpl_media]
    prod_media = [n for n in only_p if n.startswith("word/media/")]
    prod_other = [n for n in only_p if n not in prod_media]

    if tpl_media:
        out.append(Finding("parts.only_tpl.media", WARN,
                           f"仅模板有 {len(tpl_media)} 个示例部件（插图/缩略图）",
                           "\n".join(tpl_media)))
    if prod_media:
        out.append(Finding("parts.only_prod.media", WARN,
                           f"仅成品有 {len(prod_media)} 张论文插图",
                           "\n".join(prod_media)))
    if tpl_other:
        out.append(Finding("parts.only_tpl.other", WARN,
                           f"仅模板有 {len(tpl_other)} 个部件：{', '.join(tpl_other)}"))
    if prod_other:
        out.append(Finding("parts.only_prod.other", WARN,
                           f"仅成品有 {len(prod_other)} 个部件：{', '.join(prod_other)}"))
    if not only_t and not only_p:
        out.append(Finding("parts.names", OK, "部件清单与模板一致"))

    # ── 共有部件逐字节 ──
    same, diff = [], []
    for n in sorted(nt & np_):
        (same if ctx.tpl.raw(n) == ctx.prod.raw(n) else diff).append(n)
    content_diff = [n for n in diff if n in CONTENT_PARTS
                    or (n.startswith("word/media/") and n not in FIXED_MEDIA)]
    other_diff = [n for n in diff if n not in content_diff]
    if content_diff:
        out.append(Finding("parts.diff.content", WARN,
                           f"{len(content_diff)} 个内容类部件随论文变化",
                           "\n".join(f"{n}（模板 {len(ctx.tpl.raw(n))} B / "
                                     f"成品 {len(ctx.prod.raw(n))} B）"
                                     for n in content_diff)))
    if other_diff:
        out.append(Finding("parts.diff.other", WARN,
                           f"{len(other_diff)} 个格式类部件与模板不一致："
                           f"{', '.join(other_diff)}"))
    if not diff:
        out.append(Finding("parts.bytes", OK, "共有部件全部逐字节一致"))
    elif same:
        out.append(Finding("parts.bytes", OK, f"{len(same)} 个格式类部件逐字节一致"))

    # ── 封面与承诺书用图 ──
    for n in FIXED_MEDIA:
        if not ctx.prod.has(n):
            out.append(Finding("parts.fixed_media.missing", FAIL, f"成品缺失固定用图 {n}"))
        elif ctx.tpl.raw(n) != ctx.prod.raw(n):
            out.append(Finding("parts.fixed_media.diff", FAIL, f"固定用图 {n} 与模板不一致"))

    # ── [Content_Types].xml ──
    if not (ctx.tpl.has("[Content_Types].xml") and ctx.prod.has("[Content_Types].xml")):
        out.append(Finding("parts.contenttypes.missing", FAIL, "缺少 [Content_Types].xml"))
    else:
        a, b = ctx.tpl.text("[Content_Types].xml"), ctx.prod.text("[Content_Types].xml")
        if a == b:
            out.append(Finding("parts.contenttypes", OK, "[Content_Types].xml 逐字节一致"))
        else:
            ea = set(re.findall(r'Extension="([^"]+)"', a))
            eb = set(re.findall(r'Extension="([^"]+)"', b))
            pa = set(re.findall(r'PartName="([^"]+)"', a))
            pb = set(re.findall(r'PartName="([^"]+)"', b))
            out.append(Finding("parts.contenttypes.diff", WARN,
                               "[Content_Types].xml 与模板不一致",
                               f"仅模板声明的扩展名：{sorted(ea - eb)}\n"
                               f"仅成品声明的扩展名：{sorted(eb - ea)}\n"
                               f"仅模板声明的 PartName：{sorted(pa - pb)}\n"
                               f"仅成品声明的 PartName：{sorted(pb - pa)}"))

    # ── 成品 media 是否都在 Content_Types 里声明 ──
    # 上一版脚本在这里用了被覆盖的变量，把 [Content_Types].xml 的内容当成了
    # document.xml.rels，于是恒报「未声明」。这里用独立变量，结论才是真的。
    ct_xml = ctx.prod.text("[Content_Types].xml")
    declared = set(re.findall(r'Extension="([^"]+)"', ct_xml))
    media = [n for n in ctx.prod.names() if n.startswith("word/media/")]
    media_ext = {os.path.splitext(n)[1][1:].lower() for n in media}
    if media_ext <= declared:
        out.append(Finding("parts.media.declared", OK,
                           f"{len(media)} 个媒体文件的扩展名均已在 Content_Types 声明"))
    else:
        out.append(Finding("parts.media.declared", FAIL,
                           f"媒体扩展名未在 Content_Types 声明："
                           f"{sorted(media_ext - declared)}"))

    # ── 根关系与文档关系 ──
    ra = set(re.findall(r"<Relationship [^>]*/>", ctx.tpl.text("_rels/.rels")))
    rb = set(re.findall(r"<Relationship [^>]*/>", ctx.prod.text("_rels/.rels")))
    if ra == rb:
        out.append(Finding("parts.rels.root", OK, "_rels/.rels 关系集合一致"))
    else:
        out.append(Finding("parts.rels.root", WARN, "_rels/.rels 关系集合不同",
                           "仅模板：" + ("\n".join(sorted(ra - rb)) or "（无）") +
                           "\n仅成品：" + ("\n".join(sorted(rb - ra)) or "（无）")))

    da = set(re.findall(r"<Relationship [^>]*/>", ctx.tpl.text("word/_rels/document.xml.rels")))
    db = set(re.findall(r"<Relationship [^>]*/>", ctx.prod.text("word/_rels/document.xml.rels")))
    na = {r for r in da if "image" not in r}
    nb = {r for r in db if "image" not in r}
    if na == nb:
        out.append(Finding("parts.rels.doc", OK,
                           f"document.xml.rels 非图片关系一致"
                           f"（图片关系 模板 {len(da - na)} / 成品 {len(db - nb)}）"))
    else:
        out.append(Finding("parts.rels.doc", WARN, "document.xml.rels 非图片关系不同",
                           "仅模板：" + ("\n".join(sorted(na - nb)) or "（无）") +
                           "\n仅成品：" + ("\n".join(sorted(nb - na)) or "（无）")))

    # ── 页眉页脚 ──
    hf = sorted(n for n in nt if "header" in n or "footer" in n)
    bad_hf = []
    for n in hf:
        if not ctx.prod.has(n):
            bad_hf.append(f"{n}：成品缺失")
        elif ctx.tpl.raw(n) != ctx.prod.raw(n):
            bad_hf.append(f"{n}：与模板不一致")
    if bad_hf:
        out.append(Finding("parts.headerfooter", WARN, "页眉页脚与模板不符",
                           "\n".join(bad_hf)))
    elif hf:
        out.append(Finding("parts.headerfooter", OK, f"{len(hf)} 个页眉页脚部件逐字节一致"))

    # ── 序列化风格 ──
    doc = ctx.prod.text("word/document.xml")
    n_space = doc.count(" />")
    n_nsnum = len(re.findall(r"xmlns:ns[0-9]+=", doc))
    if n_space or n_nsnum:
        out.append(Finding("parts.serialize", WARN,
                           "ElementTree 的序列化痕迹未清干净",
                           f"自闭合标签带空格 ' />'：{n_space} 处（模板 0）\n"
                           f"自动编号的 ns 前缀：{n_nsnum} 处（模板 0）"))
    else:
        out.append(Finding("parts.serialize", OK,
                           "无 ET 序列化痕迹（自闭合无空格、无 nsN 前缀）"))
    return out

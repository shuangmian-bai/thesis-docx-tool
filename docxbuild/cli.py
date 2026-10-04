"""把 `章节/*.md` 合成 Word 论文（docx）。

思路：**保留模板的骨架，只替换正文内容**。
`template.docx` 的正文 XML 里，前 57 个直接子元素是封面、诚信承诺书、目录域与分节符，
最后一个是带页眉页脚引用的 `sectPr`。本脚本把中间那一段（模板自带的示例正文）换成
由 Markdown 生成的段落，其余部分原样保留——包括封面、承诺书、活目录域、页眉页脚引用、
styles.xml / numbering.xml / theme / fontTable 与字体设置。

这样做的代价是**不能用 ElementTree 整体重新序列化文档**：`mc:Ignorable` 里列出的那些
前缀（w14、w15、wp14…）在重新序列化时会被丢掉，文档会损坏。因此这里采用
「字符串拼接 + 逐片段序列化」的方式，并把根标签上的 xmlns 声明原样搬运过去。

用法：

    python3 main.py build
    python3 main.py build --template /path/to/其它模板.docx   # 换模板

产物：`output/论文_v1_YYYYMMDD.docx`（日期取当天），并把各章合并为 `output/论文正文.md`。
默认模板为 `config/template.docx`（使用者自备），可用 `--template` 覆盖。
"""
import argparse
import datetime
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

from docxbuild import HERE, W
from docxbuild.docinfo import COVER, TITLE, VERSION
from docxbuild.fragments import (
    caption_xml, code_xml, heading_xml, image_xml, list_item_xml, para_xml,
    ref_xml, table_xml,
)
from docxbuild.mdparse import parse_md
from docxbuild.template import (
    TPL_SIGN_PART, add_update_fields, check_prefixes, declared_namespaces,
    document_namespaces, fill_cover, fit_signature, fix_core_props, free_toc_case,
    list_numbering, make_list_num_xml, register_all, shrink_toc, tidy_frag,
)

CHAP_DIR = os.path.join(HERE, "config", "章节")
# 默认模板：config/template.docx（使用者自备，不入库）；也可用 --template 指定
DEFAULT_TEMPLATE = os.path.join(HERE, "config", "template.docx")
# 目录页码表：由 toc_pages.py 从渲染出的 PDF 里提取，缺失时目录页码留空
TOC_PAGES = os.path.join(HERE, "output", "toc_pages.json")
# 产物输出目录
OUTPUT_DIR = os.path.join(HERE, "output")


def build(template=None, chap_dir=None, output_dir=None, toc_pages=None,
          images_base=None):
    """核心构建流程（不含命令行解析）。

    所有路径参数均可注入，便于批量/多线程场景下用哈希隔离的工作目录调用，
    避免多个任务共享 config/章节、config/images 造成污染。

    参数为 None 时回退到全局默认值（兼容旧的命令行与单线程调用）。
    """
    TEMPLATE = template or DEFAULT_TEMPLATE
    CHAP = chap_dir or CHAP_DIR
    OUT = output_dir or OUTPUT_DIR
    TOC = toc_pages or TOC_PAGES
    # 图片解析基准：md 里的相对路径（如 images/xxx.png）相对于此目录解析。
    # 默认为项目根 config/（章节与图片都在 config/ 下）；批量场景传工作目录。
    IMG_BASE = images_base or os.path.join(HERE, "config")

    today = datetime.date.today().strftime("%Y%m%d")
    os.makedirs(OUT, exist_ok=True)
    out_docx = os.path.join(OUT, f"论文_{VERSION}_{today}.docx")
    out_md = os.path.join(OUT, "论文正文.md")

    if not os.path.isdir(CHAP):
        raise SystemExit(f"找不到章节目录：{CHAP}（请在该目录下放 *.md）")
    all_md = sorted(f for f in os.listdir(CHAP) if f.endswith(".md"))
    if not all_md:
        raise SystemExit("章节目录里没有 .md 文件")
    # 章节文件约定以数字开头（如 01_设计思路.md）。若存在此类文件，只处理它们，
    # 跳过模板示例、整篇拆解产物等非章节文件，避免正文被重复构建。
    # 目录里全是非数字开头的文件时（如审阅页导出的单篇论文 md），退化为全量处理。
    numbered = [f for f in all_md if f[:1].isdigit()]
    if numbered:
        skipped = [f for f in all_md if not f[:1].isdigit()]
        if skipped:
            print(f"  注意：章节目录下有 {len(skipped)} 个非章节文件已跳过"
                  f"（章节文件名需以数字开头，如 01_设计思路.md）："
                  + "、".join(skipped))
        chapters = numbered
    else:
        chapters = all_md

    # ── 1. 合并 Markdown，供版本管理与比对 ──
    merged = []
    for fn in chapters:
        with open(os.path.join(CHAP, fn), encoding="utf-8") as fh:
            merged.append(fh.read().rstrip() + "\n")
    with open(out_md, "w", encoding="utf-8") as fh:
        fh.write("\n\n".join(merged))
    print(f"已写 {os.path.relpath(out_md, HERE)}（{len(chapters)} 章）")

    # ── 2. 读模板 ──
    if not os.path.exists(TEMPLATE):
        raise SystemExit(f"找不到模板：{os.path.relpath(TEMPLATE, HERE)}（请放到 config/template.docx 或用 --template 指定）")
    zin = zipfile.ZipFile(TEMPLATE)
    raw = zin.read("word/document.xml").decode("utf-8")
    ns, root_tag = declared_namespaces(raw)
    all_ns = document_namespaces(raw)
    register_all(all_ns)
    print(f"模板根标签声明了 {len(ns)} 个命名空间前缀；"
          f"含正文就地声明共 {len(all_ns)} 个，已全部登记")

    doc = ET.fromstring(raw)
    body = doc.find(f"{W}body")
    kids = list(body)

    # 分节符段落：pPr 里含 sectPr 的那一段；其前是封面/承诺书/目录/分节符，其后是模板正文
    boundary = None
    for i, c in enumerate(kids):
        ppr = c.find(f"{W}pPr") if c.tag == f"{W}p" else None
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            boundary = i
    if boundary is None:
        raise SystemExit("模板里没找到分节符段落")
    final_sectpr = kids[-1]
    if final_sectpr.tag != f"{W}sectPr":
        raise SystemExit("模板最后的元素不是 sectPr")
    print(f"模板正文区：子元素 [{boundary + 1} .. {len(kids) - 2}]，"
          f"共 {len(kids) - 2 - boundary} 个（将被替换）")

    keep = kids[:boundary + 1]

    # ── 3. 填封面与承诺书、重建目录 ──
    # 目录要用到标题清单，先扫一遍各章的 # / ## 标题
    outline = []
    for fn in chapters:
        with open(os.path.join(CHAP, fn), encoding="utf-8") as fh:
            for ln in fh:
                m = re.match(r"^(#{1,2})\s+(.*?)\s*$", ln)
                if m:
                    outline.append((len(m.group(1)), m.group(2)))

    pages = {}
    if os.path.exists(TOC):
        with open(TOC, encoding="utf-8") as fh:
            pages = json.load(fh)
        print(f"已读取 {TOC}（{len(pages)} 条页码）")

    log = []
    fill_cover(keep, log)
    keep = shrink_toc(keep, log, outline, pages) or keep
    for line in log:
        print(" ", line)

    # ── 4. 换掉承诺书里的示例签名 ──
    # 模板的示例签名是模板自带的，属保留区的一部分。关系文件给出它的 rId，
    # 模板 XML 给出它的浮动尺寸，两处都要改（见 template.fit_signature）。
    rels_xml = zin.read("word/_rels/document.xml.rels").decode("utf-8")
    sign_path = os.path.join(HERE, "config", "images", "signature.png")
    if not os.path.exists(sign_path):
        raise SystemExit(f"签名图不存在：{os.path.relpath(sign_path, HERE)}")
    sign_cx, sign_cy = fit_signature(keep, rels_xml, sign_path)
    sign_png = open(sign_path, "rb").read()
    print(f"  承诺书签名已换为 config/images/signature.png"
          f"（{sign_cx / 360000:.2f} × {sign_cy / 360000:.2f} cm）")

    # ── 5. 生成正文，并登记图片关系 ──
    used_ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
    next_rel = max(used_ids) + 1 if used_ids else 1

    # 有序列表编号规划：每个连续 ol 段发一个新 numId（从 1 重编号）。
    # 模板没有 numbering.xml 或缺 decimal 定义时，ol 降级为普通段落并计数提示。
    numbering_xml = ""
    if "word/numbering.xml" in zin.namelist():
        numbering_xml = zin.read("word/numbering.xml").decode("utf-8")
    max_num_id, list_abs = list_numbering(numbering_xml)
    next_num_id = max_num_id + 1
    new_nums_xml = []
    cur_list_id = None
    n_ol_fallback = 0

    media, new_rels, body_xml = {}, [], []
    n_img = n_tbl = n_code = n_ref = n_ol = n_ol_group = 0

    for fn in chapters:
        with open(os.path.join(CHAP, fn), encoding="utf-8") as fh:
            blocks = parse_md(fh.read())
        for blk in blocks:
            kind = blk[0]
            if kind != "ol":
                cur_list_id = None      # 连续 ol 段在任何非 ol 块处中断
            if kind == "h":
                body_xml.append(heading_xml(blk[1], blk[2]))
            elif kind in ("p", "quote"):
                # 模板正文里没有引用块这种结构，引用一律按普通正文段落排
                body_xml.append(para_xml(blk[1], style="a0"))
            elif kind == "caption":
                # 表题排在表格上方，与整表绑成一块，避免表格被分页截断时表题留在上一页
                body_xml.append(caption_xml(blk[1], keep_next=blk[1].startswith("表")))
            elif kind == "ref":
                body_xml.append(ref_xml(blk[1]))
                n_ref += 1
            elif kind == "code":
                body_xml.append(code_xml(blk[1]))
                n_code += 1
            elif kind == "table":
                body_xml.append(table_xml(blk[1]))
                n_tbl += 1
            elif kind == "ol":
                n_ol += 1
                if list_abs is not None:
                    if cur_list_id is None:
                        # 新的连续列表段：发新 numId，编号自动从 1 开始
                        cur_list_id = next_num_id
                        next_num_id += 1
                        new_nums_xml.append(
                            make_list_num_xml(cur_list_id, list_abs))
                        n_ol_group += 1
                    body_xml.append(list_item_xml(blk[1], cur_list_id))
                else:
                    # 模板不支持自动编号：降级正文，不静默丢内容
                    body_xml.append(para_xml(blk[1], style="a0"))
                    n_ol_fallback += 1
            elif kind == "img":
                # 图片路径相对于 IMG_BASE（默认为 config/，批量场景为工作目录）
                path = os.path.normpath(os.path.join(IMG_BASE, blk[1]))
                if not os.path.exists(path):
                    raise SystemExit(f"插图不存在：{blk[1]}（应放在 config/images/ 下）")
                name = os.path.basename(path)
                rid = f"rId{next_rel}"
                next_rel += 1
                media[f"word/media/{name}"] = open(path, "rb").read()
                new_rels.append(f'<Relationship Id="{rid}" '
                                f'Type="http://schemas.openxmlformats.org/'
                                f'officeDocument/2006/relationships/image" '
                                f'Target="media/{name}"/>')
                body_xml.append(image_xml(rid, path, name, blk[3], blk[4]))
                if blk[2]:
                    body_xml.append(caption_xml(blk[2]))
                n_img += 1
        # 章与章之间留一个空段作分隔；末章之后不留，否则会多出一整页空白；
        # 分隔段同时把跨章的连续列表切断（下一章的 ol 段另发 numId）
        cur_list_id = None
        if fn != chapters[-1]:
            body_xml.append('<w:p><w:pPr><w:pStyle w:val="a0"/></w:pPr></w:p>')

    # ── 6. 拼接 document.xml ──
    for k, frag in enumerate(body_xml):
        check_prefixes(frag, ns, f"生成片段 #{k}")
    for c in keep:
        if isinstance(c, tuple):
            check_prefixes(c[1], ns, "目录提示段")
    body_open = raw.index("<w:body")
    head = raw[:raw.index(">", body_open) + 1]
    tail = raw[raw.rindex("</w:body>"):]

    parts = [head]
    for c in keep:
        if isinstance(c, tuple):            # 直接拼接的原始片段
            parts.append(c[1])
        else:
            parts.append(tidy_frag(
                ET.tostring(c, encoding="unicode"), ns))
    parts.extend(body_xml)
    parts.append(tidy_frag(
        ET.tostring(final_sectpr, encoding="unicode"), ns))
    parts.append(tail)
    new_doc = "".join(parts)

    # 自检：解析一遍，确认没有破坏 XML
    try:
        ET.fromstring(new_doc)
    except ET.ParseError as e:
        dbg = os.path.join(HERE, ".debug_document.xml")
        with open(dbg, "w", encoding="utf-8") as fh:
            fh.write(new_doc)
        line, col = e.position
        lines = new_doc.split("\n")
        off = sum(len(x) + 1 for x in lines[:line - 1]) + col
        raise SystemExit(
            f"XML 校验失败：{e}\n调试文件：{dbg}\n"
            f"出错位置前后 400 字符：\n..."
            f"{new_doc[max(0, off - 400):off + 400]}...")
    if new_doc.count("<w:body") != 1 or new_doc.count("</w:body>") != 1:
        raise SystemExit("生成的 document.xml 结构异常")

    # ── 7. 清理模板遗留的无用部件 ──
    # 模板自带的 30 张示例插图与 1.6 MB 的封面缩略图已无人引用，留着白白撑大文件。
    protected = set()
    for item in zin.infolist():
        fn = item.filename
        if not fn.endswith(".rels") or fn == "word/_rels/document.xml.rels":
            continue
        for m in re.finditer(r'Target="([^"]+)"', zin.read(fn).decode("utf-8")):
            protected.add(os.path.basename(m.group(1)))

    # 保留下来的正文区（封面等）仍在引用的图片也不能删
    kept_text = "".join(
        c[1] if isinstance(c, tuple) else ET.tostring(c, encoding="unicode")
        for c in keep)
    old_target = {m.group(1): m.group(2) for m in re.finditer(
        r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml)}
    for rid in re.findall(r'r:(?:embed|id|link)="(rId\d+)"', kept_text):
        if rid in old_target:
            protected.add(os.path.basename(old_target[rid]))

    drop = {"docProps/thumbnail.emf"}
    mine = {os.path.basename(k) for k in media}
    for item in zin.infolist():
        base = os.path.basename(item.filename)
        if (item.filename.startswith("word/media/")
                and base not in protected and base not in mine):
            drop.add(item.filename)
    dropped_base = {os.path.basename(d) for d in drop}

    # ── 8. 写 docx：只替换必要的部件，其余按字节复制 ──
    settings = add_update_fields(zin.read("word/settings.xml").decode("utf-8"))
    rels_new = re.sub(
        r"<Relationship[^>]*/>",
        lambda m: "" if os.path.basename(
            re.search(r'Target="([^"]*)"', m.group(0)).group(1)
        ) in dropped_base else m.group(0),
        rels_xml)
    rels_new = rels_new.replace(
        "</Relationships>", "".join(new_rels) + "</Relationships>")

    replace = {
        "word/document.xml": new_doc.encode("utf-8"),
        "word/settings.xml": settings.encode("utf-8"),
        "word/_rels/document.xml.rels": rels_new.encode("utf-8"),
        # 签名图占的是模板示例签名的部件名，直接按字节覆盖；它被承诺书引用，
        # 故上面第 7 步的清理会把它计入 protected 而不会误删
        f"word/media/{TPL_SIGN_PART}": sign_png,
        "word/styles.xml": free_toc_case(
            zin.read("word/styles.xml").decode("utf-8")).encode("utf-8"),
    }
    if new_nums_xml:
        # 新增的列表编号实例插到根元素末尾（w:num 必须排在 abstractNum 之后）
        replace["word/numbering.xml"] = numbering_xml.replace(
            "</w:numbering>", "".join(new_nums_xml) + "</w:numbering>",
            1).encode("utf-8")
    if "docProps/core.xml" in zin.namelist():
        core = zin.read("docProps/core.xml").decode("utf-8")
        replace["docProps/core.xml"] = fix_core_props(
            core, TITLE, COVER["学生姓名"]).encode("utf-8")
    content_types = zin.read("[Content_Types].xml").decode("utf-8")
    if "thumbnail" in drop or any("thumbnail" in d for d in drop):
        root_rels = zin.read("_rels/.rels").decode("utf-8")
        root_rels = re.sub(r"<Relationship[^>]*thumbnail[^>]*/>", "", root_rels)
        replace["_rels/.rels"] = root_rels.encode("utf-8")
        content_types = re.sub(
            r'<Override[^>]*thumbnail\.emf[^>]*/>', "", content_types)
        replace["[Content_Types].xml"] = content_types.encode("utf-8")
    replace.update(media)
    zin.close()

    with zipfile.ZipFile(out_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        src = zipfile.ZipFile(TEMPLATE)
        for item in src.infolist():
            if item.filename in drop:
                continue
            data = replace.pop(item.filename, None)
            if data is None:
                data = src.read(item.filename)
            zout.writestr(item, data)
        for name, data in replace.items():
            zout.writestr(name, data)
        src.close()

    print(f"  已清理模板遗留部件 {len(drop)} 个"
          f"（含无引用的示例插图与封面缩略图）")

    size = os.path.getsize(out_docx) / 1024
    print(f"\n已生成 {os.path.relpath(out_docx, HERE)}  ({size:.0f} KB)")
    print(f"  章节 {len(chapters)} 个 · 插图 {n_img} 张 · 表格 {n_tbl} 个 · "
          f"代码块 {n_code} 个 · 参考文献 {n_ref} 条 · "
          f"有序列表 {n_ol_group} 段/{n_ol} 条")
    if n_ol_fallback:
        print(f"  注意：模板缺少 numbering.xml 或 decimal 编号定义，"
              f"{n_ol_fallback} 条列表项已按普通正文排版；"
              f"在模板中补编号定义后重新构建即可恢复自动编号")
    if pages:
        print("  目录已预填静态页码；在 Word 中按 Ctrl+A 后按 F9 可重建为由排版引擎计算的域结果")
    else:
        print("  目录页码为空，跑 toc_pages.py 后重新生成即可填入（见该脚本的用法说明）")
    return out_docx


def main(argv=None):
    """命令行入口：解析参数后调用 build()。"""
    ap = argparse.ArgumentParser(
        description="把 章节/*.md 合成 Word 论文（套用模板骨架）。")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE,
                    help=f"模板 docx 路径（默认 {DEFAULT_TEMPLATE}）")
    args = ap.parse_args(argv)
    build(template=args.template)

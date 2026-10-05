"""模板 docx 的只读检视：骨架部件、封面字段、样式、编号、页面设置。

这是「模板预览」的事实来源：GUI 只负责展示，判断逻辑全在本模块，
不依赖 PyQt6，命令行也可调用 `inspect_template()`。

设计原则（见 README「设计理念」）：本模块**不直接修改模板 OOXML**——骨架与
样式的修复在 Word 里完成后替换模板文件，这里负责检出问题并支持修复后重新验证。
封面与承诺书（含签名）保留模板原样，不提供 cover.json / signature.png 修复入口。
"""
import os
import zipfile
import xml.etree.ElementTree as ET

from docxbuild import W

#: 构建端硬依赖的段落样式（与 fragments.py 一致）
REQUIRED_STYLES = ("1", "2", "3", "4", "a0")
OPTIONAL_STYLES = ("aa", "ac", "ad", "code")


def _text_of(p):
    return "".join(t.text or "" for t in p.iter(f"{W}t")).strip()


def inspect_template(path: str) -> dict:
    """检视模板，返回结构化报告 dict（含 problems 与 skeleton 预览）。"""
    rep = {"path": os.path.abspath(path), "exists": os.path.exists(path),
           "problems": [], "skeleton": [], "parts": {},
           "styles": {}, "page": {}, "numbering": {}}
    if not rep["exists"]:
        rep["problems"].append(("error", f"模板文件不存在：{rep['path']}"))
        return rep

    try:
        zin = zipfile.ZipFile(path)
        names = set(zin.namelist())
    except (OSError, zipfile.BadZipFile) as e:
        rep["problems"].append(("error", f"无法作为 docx 打开：{e}"))
        return rep

    rep["parts"] = {
        "styles": "word/styles.xml" in names,
        "numbering": "word/numbering.xml" in names,
        "header": sorted(n for n in names if n.startswith("word/header")),
        "footer": sorted(n for n in names if n.startswith("word/footer")),
        "media": sorted(n for n in names if n.startswith("word/media/")),
        "core": "docProps/core.xml" in names,
    }

    if "word/document.xml" not in names:
        rep["problems"].append(("error", "缺少 word/document.xml"))
        zin.close()
        return rep

    doc = ET.fromstring(zin.read("word/document.xml"))
    body = doc.find(f"{W}body")
    kids = list(body) if body is not None else []

    # ── 骨架边界：含 sectPr 的分节符段落 ──
    boundary = None
    for i, c in enumerate(kids):
        ppr = c.find(f"{W}pPr") if c.tag == f"{W}p" else None
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            boundary = i
            break
    if boundary is None:
        rep["problems"].append(("error", "找不到分节符段落（封面/承诺书/目录的边界标记）"))
    final = kids[-1] if kids else None
    if final is None or final.tag != f"{W}sectPr":
        rep["problems"].append(("error", "正文最后的元素不是 sectPr（页面设置丢失）"))
    kept = kids[:boundary + 1] if boundary is not None else []
    rep["skeleton_count"] = len(kept)

    # ── 骨架纯文本预览（不渲染 Word 格式）──
    for c in kept:
        if c.tag == f"{W}p":
            txt = _text_of(c)
            if txt:
                rep["skeleton"].append(("段落", txt[:48]))
        elif c.tag == f"{W}tbl":
            first = ""
            tcs = c.findall(f"{W}tr/{W}tc")
            if tcs:
                first = _text_of(tcs[0])
            rep["skeleton"].append(("表格", (first[:40] or "封面/字段表")))

    # ── 封面表格字段（只列模板里实际有的标签，不再按预定义集合判缺失）──
    cover_tbl = next((c for c in kept if c.tag == f"{W}tbl"), None)
    if cover_tbl is None:
        rep["problems"].append(("warn", "保留区找不到封面表格"))
    else:
        labels = []
        for row in cover_tbl.findall(f"{W}tr"):
            tcs = row.findall(f"{W}tc")
            if len(tcs) >= 2:
                lab = _text_of(tcs[0])
                if lab:
                    labels.append(lab)
        rep["cover"] = {"labels": labels}

    # ── 目录域 ──
    kept_xml = "".join(ET.tostring(c, encoding="unicode") for c in kept)
    rep["has_toc"] = ("TOC" in kept_xml and "instrText" in kept_xml)
    if not rep["has_toc"]:
        rep["problems"].append(("warn", "保留区未发现 TOC 目录域，成品不会有自动目录"))

    # ── 样式 ──
    style_ids = set()
    if rep["parts"]["styles"]:
        st_root = ET.fromstring(zin.read("word/styles.xml"))
        style_ids = {s.get(f"{W}styleId") for s in st_root
                     if s.get(f"{W}styleId")}
    for sid in REQUIRED_STYLES:
        if sid not in style_ids:
            rep["problems"].append((
                "error", f"缺少必需样式 styleId={sid}"
                         + ("（标题样式）" if sid.isdigit() else "（正文样式 a0）")))
    for sid in OPTIONAL_STYLES:
        if sid not in style_ids:
            rep["problems"].append((
                "warn", f"缺少样式 styleId={sid}，相关内容将无法按模板格式排版"))
    rep["styles"] = {"present": sorted(s for s in style_ids if s),
                     "required": list(REQUIRED_STYLES),
                     "optional": list(OPTIONAL_STYLES)}

    # ── 编号定义（有序列表）──
    if rep["parts"]["numbering"]:
        nraw = zin.read("word/numbering.xml").decode("utf-8")
        decimal = ('<w:numFmt w:val="decimal"' in nraw)
        rep["numbering"] = {
            "decimal": decimal,
            "num_count": nraw.count("<w:num "),
        }
        if not decimal:
            rep["problems"].append((
                "warn", "numbering.xml 无 decimal 编号定义，有序列表将降级为普通正文"))
    else:
        rep["numbering"] = {"decimal": False, "num_count": 0}
        rep["problems"].append((
            "warn", "模板缺少 numbering.xml，有序列表将降级为普通正文"))

    # ── 页面设置（twips → cm）──
    if final is not None and final.tag == f"{W}sectPr":
        sz = final.find(f"{W}pgSz")
        mg = final.find(f"{W}pgMar")
        if sz is not None:
            rep["page"] = {
                "width_cm": round(int(sz.get(f"{W}w", 0)) / 567, 2),
                "height_cm": round(int(sz.get(f"{W}h", 0)) / 567, 2),
            }
        if mg is not None:
            rep["page"].update({
                "margin_cm": {
                    k: round(int(mg.get(f"{W}{k}", 0)) / 567, 2)
                    for k in ("top", "bottom", "left", "right")},
            })

    zin.close()
    rep["problems"].sort(key=lambda p: 0 if p[0] == "error" else 1)
    return rep

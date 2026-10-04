"""模板 docx 的只读检视：骨架部件、封面字段、样式、编号、页面设置。

这是「模板预览与修复」的事实来源：GUI 只负责展示与落配置，判断逻辑全在本模块，
不依赖 PyQt6，命令行也可调用 `inspect_template()`。

设计原则（见 README「设计理念」）：本模块**不直接修改模板 OOXML**——骨架与
样式的修复在 Word 里完成后替换模板文件，这里负责检出问题、给出修复入口
（cover.json 封面字段、signature.png 签名图）并支持修复后重新验证。
"""
import json
import os
import zipfile
import xml.etree.ElementTree as ET

from docxbuild import W
from docxbuild.docinfo import COVER, TEMPLATE_TITLE
from docxbuild.template import TPL_SIGN_PART

#: 使用者侧的承诺书签名图落点（build 的 fit_signature 读这个路径；
#: TPL_SIGN_PART 是模板内部件名，二者不是一回事）
SIGNATURE_NAME = "signature.png"

#: 构建端硬依赖的段落样式（与 fragments.py 一致）
REQUIRED_STYLES = ("1", "2", "3", "4", "a0")
OPTIONAL_STYLES = ("aa", "ac", "ad", "code")


def _text_of(p):
    return "".join(t.text or "" for t in p.iter(f"{W}t")).strip()


def inspect_template(path: str) -> dict:
    """检视模板，返回结构化报告 dict（含 problems 与 skeleton 预览）。"""
    rep = {"path": os.path.abspath(path), "exists": os.path.exists(path),
           "problems": [], "skeleton": [], "parts": {}, "cover": {},
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
    rep["parts"]["has_signature_part"] = (
        f"word/media/{TPL_SIGN_PART}" in names)

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

    # ── 封面表格字段 ──
    cover_tbl = next((c for c in kept if c.tag == f"{W}tbl"), None)
    if cover_tbl is None:
        rep["problems"].append(("error", "保留区找不到封面表格（fill_cover 将中断构建）"))
    else:
        labels = []
        for row in cover_tbl.findall(f"{W}tr"):
            tcs = row.findall(f"{W}tc")
            if len(tcs) >= 2:
                lab = _text_of(tcs[0])
                if lab:
                    labels.append(lab)
        rep["cover"]["labels"] = labels
        missing = [k for k in COVER if k not in labels]
        rep["cover"]["missing"] = missing
        for k in missing:
            rep["problems"].append(
                ("warn", f"封面表格缺少字段行「{k}」，该字段构建时不会被填入"))

    # ── 诚信承诺书题目占位符 ──
    has_commit_title = any(
        TEMPLATE_TITLE in _text_of(c) for c in kept if c.tag == f"{W}p")
    rep["commitment_placeholder"] = has_commit_title
    if not has_commit_title:
        rep["problems"].append((
            "warn", "承诺书里未找到模板题目占位符，构建时承诺书题目不会被替换"))

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

    # ── 签名部件 ──
    if not rep["parts"]["has_signature_part"]:
        rep["problems"].append((
            "warn", f"模板关系中找不到承诺书签名部件 {TPL_SIGN_PART}，"
                    "fit_signature 将中断构建"))

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


# ── 修复入口：cover.json 读写（不碰模板本体）────────────────────────
def cover_json_path(tool_root: str) -> str:
    return os.path.join(tool_root, "config", "cover.json")


def load_cover_form(tool_root: str) -> dict:
    """读 cover.json 为表单 dict；不存在时返回空值表单（键固定）。"""
    form = {"title": "", "template_title": "", "version": "",
            "cover": {k: "" for k in COVER}}
    p = cover_json_path(tool_root)
    if os.path.exists(p):
        with open(p, encoding="utf-8") as fh:
            cfg = json.load(fh)
        form["title"] = cfg.get("title", "")
        form["template_title"] = cfg.get("template_title", "")
        form["version"] = str(cfg.get("version", ""))
        for k in COVER:
            form["cover"][k] = cfg.get("cover", {}).get(k, "")
    return form


def save_cover_form(tool_root: str, form: dict):
    """把表单落为 config/cover.json（空字段也保留，便于继续填写）。"""
    p = cover_json_path(tool_root)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    cfg = {"title": form["title"].strip(),
           "template_title": form["template_title"].strip(),
           "cover": {k: form["cover"].get(k, "").strip() for k in COVER}}
    if form.get("version", "").strip():
        cfg["version"] = form["version"].strip()
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    return p


def signature_path(tool_root: str) -> str:
    return os.path.join(tool_root, "config", "images", SIGNATURE_NAME)

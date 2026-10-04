"""解析 docx 的 document.xml → 块序列。

按 `<w:body>` 的子元素顺序**混合遍历** `<w:p>` 与 `<w:tbl>`，避免分别遍历
`paragraphs` / `tables` 导致段落与表格的先后顺序错乱（这是 python-docx 常见坑，
我们不用它，但同样要注意顺序）。

块格式与 `docxbuild.mdparse` 完全对齐：标题、正文、图题、表题、表格、代码、参考文献、
有序列表、图片。图片块里存的是 docx 内部的 zip 路径（如 `word/media/image1.png`），
由 `cli` 在抽取图片后改写为相对路径。

有序列表**只读 Word 的自动编号属性**（pPr/numPr + numbering.xml 的 numFmt=decimal），
不靠 "1." 文本前缀猜测；手动敲出来的伪编号仍按正文解析，由 GUI 人工/AI 改判为 ol。
"""
import os
import re
import zipfile
import xml.etree.ElementTree as ET

from docxconvert import A, R, W

# 常见模板里的标题样式 ID（标题用数字 1~4）
_TPL_HEADING = {"1": 1, "2": 2, "3": 3, "4": 4}
_TPL_CAPTION = {"aa"}
_TPL_CODE = {"code"}


def _load_styles(zin):
    """读 styles.xml，返回 (style_id → 名称, style_id → 标题级别)。

    标题级别优先从样式名 `heading N` 解析；模板里的数字 ID 直接用内置表。
    """
    names, heading = {}, {}
    try:
        raw = zin.read("word/styles.xml").decode("utf-8")
    except KeyError:
        return names, heading
    root = ET.fromstring(raw)
    for st in root:
        sid = st.get(f"{W}styleId")
        if not sid:
            continue
        name_el = st.find(f"{W}name")
        name = name_el.get(f"{W}val") if name_el is not None else ""
        names[sid] = name
        low = name.lower().strip()
        m = re.match(r"heading\s*(\d+)", low)
        if m:
            heading[sid] = min(int(m.group(1)), 4)
        elif sid in _TPL_HEADING:
            heading[sid] = _TPL_HEADING[sid]
    return names, heading


def _para_style(p):
    """段落的 pStyle 值，无则 None。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return None
    pstyle = ppr.find(f"{W}pStyle")
    return pstyle.get(f"{W}val") if pstyle is not None else None


def _load_numbering(zin):
    """读 numbering.xml，返回 {numId: 0 级编号格式}（如 decimal / bullet）。

    numId 经 abstractNumId 间接引用 abstractNum；只取 ilvl=0 的 numFmt，
    供段落级分类（decimal → ol，其余暂按正文）。
    """
    try:
        raw = zin.read("word/numbering.xml").decode("utf-8")
    except KeyError:
        return {}
    abs_fmt = {}
    for m in re.finditer(
            r'<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"(.*?)</w:abstractNum>',
            raw, re.S):
        lvl0 = re.search(r'<w:lvl\b[^>]*w:ilvl="0".*?</w:lvl>', m.group(2), re.S)
        fmt = None
        if lvl0:
            fm = re.search(r'<w:numFmt w:val="([^"]+)"', lvl0.group(0))
            fmt = fm.group(1) if fm else None
        abs_fmt[m.group(1)] = fmt
    num_fmt = {}
    for m in re.finditer(r'<w:num\b[^>]*w:numId="(\d+)"(.*?)</w:num>|<w:num\b[^>]*w:numId="(\d+)"[^>]*/>',
                         raw, re.S):
        nid = m.group(1) or m.group(3)
        body = m.group(2) or ""
        am = re.search(r'<w:abstractNumId w:val="(\d+)"', body)
        if am:
            num_fmt[nid] = abs_fmt.get(am.group(1))
    return num_fmt


def _para_num(p):
    """段落的自动编号 (numId, ilvl)；无 numPr 返回 (None, 0)。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return None, 0
    numpr = ppr.find(f"{W}numPr")
    if numpr is None:
        return None, 0
    nid_el = numpr.find(f"{W}numId")
    ilvl_el = numpr.find(f"{W}ilvl")
    nid = nid_el.get(f"{W}val") if nid_el is not None else None
    try:
        ilvl = int(ilvl_el.get(f"{W}val")) if ilvl_el is not None else 0
    except (TypeError, ValueError):
        ilvl = 0
    return nid, ilvl


def _outline_level(p):
    """段落的大纲级别（pPr/outlineLvl），无则 None。用作无标题样式时的兜底。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return None
    ol = ppr.find(f"{W}outlineLvl")
    if ol is None:
        return None
    val = ol.get(f"{W}val")
    try:
        return min(int(val) + 1, 4)
    except (TypeError, ValueError):
        return None


def _para_text(p):
    """提取段落纯文本：处理 <w:t> / <w:tab> / <w:br>，合并多余空白。"""
    parts = []
    for node in p.iter():
        tag = node.tag
        if tag == f"{W}t":
            parts.append(node.text or "")
        elif tag == f"{W}tab":
            parts.append("\t")
        elif tag == f"{W}br":
            parts.append("\n")
    text = "".join(parts)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def _para_image(p, rels):
    """若段落含嵌入图片，返回其在 zip 内的路径（如 word/media/image1.png），否则 None。"""
    for blip in p.iter(f"{A}blip"):
        rid = blip.get(f"{R}embed")
        if rid and rid in rels:
            target = rels[rid]
            if target.startswith("/"):
                return target.lstrip("/")
            return "word/" + target
    return None


def _table_rows(tbl):
    """把 <w:tbl> 解析为二维单元格列表；单元格内多段用换行连接。"""
    rows = []
    for tr in tbl.findall(f"{W}tr"):
        cells = []
        for tc in tr.findall(f"{W}tc"):
            paras = []
            for p in tc.findall(f"{W}p"):
                t = _para_text(p)
                if t:
                    paras.append(t)
            cells.append("\n".join(paras))
        rows.append(cells)
    return rows


# 图题/表题的文本特征：以「图」或「表」开头，后跟数字编号
_CAPTION_RE = re.compile(r"^\s*(?:\[)?(图|表)\s*[\d\-–—.]+")
# 参考文献条目：[1] …
_REF_RE = re.compile(r"^\s*\[\d+\]\s")


def parse_docx(docx_path):
    """解析 docx，返回 (blocks, images)。

    blocks 与 mdparse 输出对齐；images 是 {zip内路径: 建议文件名}，
    由 extract 模块落盘后再把图片块的路径改写为相对路径。
    """
    zin = zipfile.ZipFile(docx_path)
    names, heading_lv = _load_styles(zin)
    num_fmt = _load_numbering(zin)

    caption_ids = ({sid for sid, nm in names.items()
                    if "图表标题" in nm or "caption" in nm.lower()}
                   | _TPL_CAPTION)
    code_ids = ({sid for sid, nm in names.items() if nm.lower() == "code"}
                | _TPL_CODE)
    toc_ids = {sid for sid, nm in names.items() if nm.lower().startswith("toc")}

    rels_raw = zin.read("word/_rels/document.xml.rels").decode("utf-8")
    rels = {m.group(1): m.group(2) for m in
            re.finditer(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"',
                        rels_raw)}

    doc = ET.fromstring(zin.read("word/document.xml"))
    body = doc.find(f"{W}body")

    blocks = []
    images = {}          # zip 内路径 → 建议文件名
    code_buf = []

    def flush_code():
        if code_buf:
            blocks.append(("code", list(code_buf)))
            code_buf.clear()

    for child in body:
        if child.tag == f"{W}p":
            sid = _para_style(child)
            if sid in toc_ids:          # 跳过目录条目
                continue

            img = _para_image(child, rels)
            if img:
                flush_code()
                images[img] = os.path.basename(img)
                blocks.append(("img", img, ""))
                continue

            text = _para_text(child)
            if not text:
                if code_buf:             # 代码块里的空行保留
                    code_buf.append("")
                continue

            if sid in code_ids:
                code_buf.append(text)
                continue
            flush_code()

            # 自动编号优先于标题判定：编号段落（decimal）归有序列表，
            # numId=0 是 Word 显式「取消编号」，不能归入
            num_id, _ilvl = _para_num(child)
            if num_id and num_id != "0" and num_fmt.get(num_id) == "decimal":
                blocks.append(("ol", text))
                continue

            lvl = heading_lv.get(sid) if sid else None
            if lvl is None:
                lvl = _outline_level(child)
            if lvl is not None:
                blocks.append(("h", lvl, text))
            elif sid in caption_ids or _CAPTION_RE.match(text):
                blocks.append(("caption", text))
            elif _REF_RE.match(text):
                blocks.append(("ref", text))
            else:
                blocks.append(("p", text))

        elif child.tag == f"{W}tbl":
            flush_code()
            rows = _table_rows(child)
            if rows:
                blocks.append(("table", rows))

    flush_code()
    zin.close()

    # 把紧跟图片后的「图题」caption 合并进图片块的说明字段，
    # 与源 MD 写法 `![图x.y 说明](path)` 对齐；表题保持独立块。
    merged = []
    i = 0
    while i < len(blocks):
        blk = blocks[i]
        if (blk[0] == "img" and i + 1 < len(blocks)
                and blocks[i + 1][0] == "caption"
                and blocks[i + 1][1].lstrip("[").startswith("图")):
            merged.append(("img", blk[1], blocks[i + 1][1]))
            i += 2
        else:
            merged.append(blk)
            i += 1
    return merged, images

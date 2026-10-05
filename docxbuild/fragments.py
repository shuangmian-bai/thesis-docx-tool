"""块 → OOXML 片段。

每个生成函数返回一段可直接拼进 `document.xml` 的字符串。片段一律**自带所需的
命名空间声明**，因为 ElementTree 不认识 `a` / `pic` 这类只在正文就地声明的前缀；
`template.check_prefixes()` 在拼接前会逐段核对，缺声明在这里就报错。

段落样式全部沿用模板自带的（`a0` 正文、`aa` 图表题、`a0` 参考文献、`code` 代码、
`ac` 表格、`ad` 表文），不自行发明；直接格式只保留模板里出现过的写法，例外见
本目录 `README.md` 的「格式约定」。
"""
import re

from docxbuild.layout import CONTENT_W, MAX_IMG_H, MAX_IMG_W
from docxbuild.mdparse import disp_w, parse_inline

_uid = [1000]


def next_id():
    _uid[0] += 1
    return _uid[0]


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _split_fmt(text):
    """解析文本开头的 `{k=v}` 格式前缀，返回 (格式字典, 剩余文本)。

    拆解端 _para_text 把段落级格式（jc/ind firstLine）写成 `{jc=center,ind=482}`
    前缀；本函数还原，供各段落渲染函数写成 `<w:jc/>` / `<w:ind/>`。
    """
    m = re.match(r"^\{([^}]*)\}", text)
    if not m:
        return {}, text
    fmt = {}
    for kv in m.group(1).split(","):
        if "=" in kv:
            k, v = kv.split("=", 1)
            fmt[k] = v
    return fmt, text[m.end():]


def runs_xml(text):
    """生成内联 run 序列。"""
    out = []
    for content, bold, code in parse_inline(text):
        rpr = []
        if bold:
            rpr.append("<w:b/>")
        if code:
            rpr.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>')
        rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>" if rpr else ""
        out.append(f'<w:r>{rpr_xml}<w:t xml:space="preserve">{esc(content)}</w:t></w:r>')
    return "".join(out)


def para_xml(text, style=None, jc=None, extra_ppr=""):
    fmt, text = _split_fmt(text)
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if extra_ppr:
        ppr.append(extra_ppr)
    if fmt.get("ind"):
        ppr.append(f'<w:ind w:firstLine="{fmt["ind"]}"/>')
    if jc or fmt.get("jc"):
        ppr.append(f'<w:jc w:val="{jc or fmt["jc"]}"/>')
    ppr_xml = f"<w:pPr>{''.join(ppr)}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}{runs_xml(text)}</w:p>"


def heading_xml(level, text):
    style = str(min(level, 4))
    # 章标题（一级）每次从新页开始
    extra = "<w:pageBreakBefore/>" if level == 1 else ""
    return para_xml(text, style=style, extra_ppr=extra)


def caption_xml(text, keep_next=False):
    """图表标题：居中、加粗、中英文字体全由样式 aa 给。

    模板的题注段落只挂 `<w:pStyle w:val="aa"/>`，段落属性一概不写
    （个别题注手动加了字号属模板特例，不在这里统一补）。
    不补 `jc`、`spacing`、`sz` 等多余直接格式。

    `keep_next` 供**表题**用：表题排在表格上方，要与整表同页（见 `table_xml`）。
    图题不能加——它在图片下方，加 `keepNext` 会与后面的正文段落粘连，
    反而可能把图与图题拆到两页。
    """
    return para_xml(text, style="aa",
                    extra_ppr="<w:keepNext/>" if keep_next else "")


def ref_xml(text):
    """参考文献条目：悬挂缩进，写法照模板的参考文献段落，一字不差。

    模板给的是 `left=425 / hangingChars=177 / hanging=425` 三件套，行距沿用
    a0 样式自带的值，不另写 `spacing`。`hanging` 与样式的 `firstLine` 同属
    段落首行属性，直接格式里的 `hanging` 会压过样式，不必再显式清零 `firstLine`。
    条目里拉丁文与 URL 长，用左对齐而非两端对齐，避免含中文的行被拉散。
    """
    ppr = ('<w:pStyle w:val="a0"/>'
           '<w:ind w:left="425" w:hangingChars="177" w:hanging="425"/>'
           '<w:jc w:val="left"/>')
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs_xml(text)}</w:p>"


def list_item_xml(text, num_id):
    """有序列表条目：挂 numbering.xml 里的自动编号（decimal），文字样式沿用正文 a0。

    缩进、编号文字（"1."）与计数全部由编号定义给出，片段里不写编号文本；
    同一段连续条目共用一个 numId，新段落在 cli 里另发新 numId 并从 1 重编号。
    """
    fmt, text = _split_fmt(text)
    ind_fl = f'<w:ind w:firstLine="{fmt["ind"]}"/>' if fmt.get("ind") else ""
    jc_xml = f'<w:jc w:val="{fmt["jc"]}"/>' if fmt.get("jc") else ""
    ppr = (f'<w:pStyle w:val="a0"/>'
           f'{ind_fl}{jc_xml}'
           f'<w:numPr><w:ilvl w:val="0"/>'
           f'<w:numId w:val="{num_id}"/></w:numPr>')
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs_xml(text)}</w:p>"


def code_xml(lines, sz=None):
    rpr = f'<w:rPr><w:sz w:val="{sz}"/></w:rPr>' if sz else ""
    return "".join(
        f'<w:p><w:pPr><w:pStyle w:val="code"/></w:pPr>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{esc(ln) if ln else " "}</w:t></w:r></w:p>'
        for ln in lines
    )


def table_xml(rows):
    """生成表格：整体沿用模板的表样式 ac，首行为表头，且**整表不跨页**。

    模板里表格不写直接边框与底纹，全部由样式 ac 提供：左右无外框，上下框较粗（sz=12），
    内部细线（sz=4）；表头由 firstRow 条件格式给出居中、加粗与灰底（BFBFBF），
    因此这里只需挂上 tblStyle 并用 cnfStyle 把首行标记为 firstRow。
    单元格正文用表文样式 ad（无首行缩进、单倍行距、10.5 磅）。

    **表格不跨页**靠两条属性组合——OOXML 没有「整表保持同页」的开关：
    每行 `<w:trPr>` 加 `cantSplit`（行内不断页），除末行外每个单元格段落加
    `keepNext`（与下一行同页）。末行不加，否则表会与后面的正文段落粘连；
    表题由 `caption_xml(keep_next=True)` 绑住，于是「表题 + 整表」成为不可分割的一块。
    模板里没有这两处写法，属有意引入的直接格式（见本目录 README「格式约定」）。

    首行的 `tblHeader` 保留：表格不跨页时它不触发，留作表格长到超过一页时的兜底。
    """
    # 每行按 span 展开：(单元格, span)；ncol = 最大展开列数；不补齐空单元格
    def span_of(c):
        m = re.match(r"^\{span=(\d+)\}", c)
        return int(m.group(1)) if m else 1

    expanded = [[(c, span_of(c)) for c in r] for r in rows]
    ncol = max(sum(s for _, s in er) for er in expanded)

    # 列宽按内容需要分配：span 单元格的需求宽度均分给 span 列，每列取最大
    # 需求，总宽按比例归一化到正文宽度。
    CHAR_W = 105      # 10.5 磅下一个半角字符的宽度（twips）
    CELL_PAD = 240    # 单元格左右边距之和

    def need(text):
        return disp_w(text or "") * CHAR_W + CELL_PAD

    col_reqs = [[] for _ in range(ncol)]
    for er in expanded:
        ci = 0
        for c, span in er:
            w = need(c) / span
            for _ in range(span):
                col_reqs[ci].append(w)
                ci += 1
    widths = [max(reqs) if reqs else CELL_PAD for reqs in col_reqs]

    total = sum(widths)
    if total != CONTENT_W:
        # 取整误差（或需求本身就超宽）统一消化到最后一列
        widths = [int(w * CONTENT_W / total) for w in widths]
        widths[-1] += CONTENT_W - sum(widths)

    # 与模板保持一致：gridCol 用 dxa（合计 = 正文宽度），tcW 用百分比（合计 5000 = 100%）
    pcts = [max(1, round(5000 * w / CONTENT_W)) for w in widths]

    parts = [
        "<w:tbl><w:tblPr>",
        '<w:tblStyle w:val="ac"/><w:tblW w:w="5000" w:type="pct"/>',
        '<w:tblLayout w:type="fixed"/>',
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1"'
        ' w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
        "</w:tblPr><w:tblGrid>",
    ]
    parts += [f'<w:gridCol w:w="{w}"/>' for w in widths]
    parts.append("</w:tblGrid>")

    last = len(rows) - 1
    for ri, er in enumerate(expanded):
        if ri == 0:
            # cnfStyle 与 tblLook 的 firstRow 共同触发样式 ac 的表头格式。
            # CT_TrPr 是 sequence：cantSplit 排在 cnfStyle 之后、tblHeader 之前。
            trpr = ('<w:trPr><w:cnfStyle w:val="100000000000" w:firstRow="1"'
                    ' w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
                    ' w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0"'
                    ' w:firstRowFirstColumn="0" w:firstRowLastColumn="0"'
                    ' w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/>'
                    '<w:cantSplit/><w:tblHeader/></w:trPr>')
        else:
            trpr = "<w:trPr><w:cantSplit/></w:trPr>"
        parts.append(f"<w:tr>{trpr}")
        # 除末行外，行内段落与下一行同页，整表因而不可分页；末行不加，
        # 否则表格会与后面的正文段落粘连。keepNext 在 CT_PPr 里紧跟 pStyle。
        keep = "" if ri == last else "<w:keepNext/>"
        ci = 0
        for cell, span in er:
            fmt, cell_body = _split_fmt(cell)
            if span > 1:
                # 跨列合并：tcW 取 span 列百分比之和，gridSpan 标记跨列
                total_pct = sum(pcts[ci:ci + span])
                tcpr = (f'<w:tcPr><w:tcW w:w="{total_pct}" w:type="pct"/>'
                        f'<w:gridSpan w:val="{span}"/><w:hideMark/></w:tcPr>')
            else:
                tcpr = f'<w:tcPr><w:tcW w:w="{pcts[ci]}" w:type="pct"/><w:hideMark/></w:tcPr>'
            # 单元格内的 <br>（来自 markdown 表格换行）拆成多个段落，
            # 避免被 esc 转义成字面量 <br>
            paras = cell_body.split("<br>") if cell_body else [""]
            body_parts = []
            for p in paras:
                fmt2, p = _split_fmt(p)
                ind_fl = f'<w:ind w:firstLine="{fmt2["ind"]}"/>' if fmt2.get("ind") else ""
                jc_xml = f'<w:jc w:val="{fmt2["jc"]}"/>' if fmt2.get("jc") else ""
                body_parts.append(
                    f'<w:p><w:pPr><w:pStyle w:val="ad"/>{ind_fl}{jc_xml}{keep}</w:pPr>'
                    f'{runs_xml(p)}</w:p>')
            body = "".join(body_parts)
            parts.append(f"<w:tc>{tcpr}{body}</w:tc>")
            ci += span
        parts.append("</w:tr>")
    parts.append("</w:tbl>")
    return "".join(parts)


def image_xml(rid, path, name, cx=0, cy=0):
    """居中插入图片。

    cx/cy 为 EMU；若均为 0，则按图片像素（96 dpi）与页面可用宽高自适应缩放。
    若传入了 cx/cy（来自模板原文档的 <wp:extent>），则直接使用该尺寸，
    保证 convert→build 闭环后图片显示尺寸与原文档一致。
    """
    from PIL import Image
    if cx and cy:
        cx, cy = int(cx), int(cy)
    else:
        with Image.open(path) as im:
            px_w, px_h = im.size
        # 按 96 dpi 换算，再按可用宽度/高度收缩
        cx, cy = px_w * 9525, px_h * 9525
        scale = min(MAX_IMG_W / cx, MAX_IMG_H / cy, 1.0)
        cx, cy = int(cx * scale), int(cy * scale)
    i = next_id()
    return (
        # 只写 jc=center：模板的插图段落也这么写，且不再补 spacing（同图题段落，
        # 多余的直接格式会压过样式，见 README「格式约定」）。
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        "<w:r><w:drawing>"
        # 模板根标签只声明了 w / wp / r 等前缀，a 与 pic 是在正文里就地声明的，
        # 因此这里必须自带声明，否则序列化出来是 unbound prefix。
        '<wp:inline xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{i}" name="图片{i}"/>'
        "<wp:cNvGraphicFramePr>"
        '<a:graphicFrameLocks noChangeAspect="1"/>'
        "</wp:cNvGraphicFramePr>"
        "<a:graphic>"
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:pic>"
        f'<pic:nvPicPr><pic:cNvPr id="{i}" name="{esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        "<pic:spPr>"
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>"
        "</w:drawing></w:r></w:p>"
    )

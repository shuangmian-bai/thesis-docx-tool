"""模板骨架处理：只动该动的地方，其余原样保留。

`template.docx` 的正文 XML 里，分节符段落之前是封面、诚信承诺书、目录域，
之后是模板自带的示例正文。构建时**保留前者、替换后者**——封面与承诺书
（含签名）一律保留模板原样，不再由代码填充或替换。这里要处理的是
「保留区」里少数按论文变化的部分：

- 目录域的缓存结果（`shrink_toc`）；
- TOC 样式的强制大小写标签（`free_toc_case`）；
- `settings.xml` 的域更新开关（`add_update_fields`）。

另外还要保证拼接出来的 XML 合法、书写风格与模板一致（`check_prefixes` /
`tidy_frag`），并登记命名空间前缀（`register_all`）。
"""
import re
import xml.etree.ElementTree as ET

from docxbuild import W
from docxbuild.fragments import esc, para_xml, runs_xml
from docxbuild.layout import CONTENT_W


def declared_namespaces(raw):
    start = raw.index("<w:document")
    root_tag = raw[start:raw.index(">", start) + 1]
    return dict(re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', root_tag)), root_tag


def document_namespaces(raw):
    """收集全文出现过的 xmlns 声明，含正文里的就地声明。

    模板只在根标签声明常用前缀，`a` / `pic` / `a14` 这些绘图命名空间是在正文的
    图形元素上就地声明的。`register_all` 若只拿到根标签那批，ElementTree 就不认识
    `a` / `pic`，序列化封面与承诺书的图形时会自己编号，吐出 `ns4` / `ns5` ——
    属性值一模一样，但前缀名与模板对不上。这里把全文前缀一并交给它。
    """
    return dict(re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', raw))


def register_all(ns):
    for prefix, uri in ns.items():
        if prefix in ("xml", "xmlns"):
            continue
        try:
            ET.register_namespace(prefix, uri)
        except (ValueError, KeyError):
            pass


def check_prefixes(frag, ns, where):
    """确认片段用到的命名空间前缀都已声明（根标签上的或片段自带的）。

    模板根标签只声明了一部分前缀，`a` / `pic` 之类是在正文里就地声明的。
    少了声明不会在拼接时报错，只会在最后解析时抛出难以定位的 unbound prefix，
    所以这里逐个片段先查一遍。
    """
    declared = set(ns) | set(re.findall(r'xmlns:([A-Za-z0-9]+)=', frag))
    used = set(re.findall(r'<([A-Za-z0-9]+):', frag))
    missing = used - declared - {"xml"}
    if missing:
        raise SystemExit(f"{where} 使用了未声明的前缀：{sorted(missing)}")


def tidy_frag(frag, ns):
    """把 ElementTree 序列化出来的片段整理成模板的书写风格。

    两件事：
    1. 去掉根标签已声明过的同名同 URI 前缀声明，避免重复。
    2. 自闭合标签写成 `<w:x/>`。ElementTree 吐的是 `<w:x />`（斜杠前带空格），
       与模板不一致。只在标签内替换——正则要求标签体内不含尖括号，属性值里的
       `>` 已被转义成 `&gt;`，所以文本节点里的 ` />` 不会被误伤。
    """
    def repl(m):
        return "" if ns.get(m.group(1)) == m.group(2) else m.group(0)

    frag = re.sub(r'\s+xmlns:([A-Za-z0-9]+)="([^"]*)"', repl, frag)
    return re.sub(r"<([^<>]*?)\s+/>", r"<\1/>", frag)


def list_numbering(numbering_xml):
    """分析模板 numbering.xml，返回 (现有最大 numId, 可复用的 decimal abstractNumId)。

    供有序列表构建：每个连续列表段新建一个 w:num 引用该 abstractNum，
    用 startOverride 从 1 重编号。没有任何 decimal 定义（或缺部件）时
    abstractNumId 为 None，调用方降级为普通段落并提示。
    """
    if not numbering_xml:
        return 0, None
    nums = [int(x) for x in re.findall(
        r'<w:num\b[^>]*w:numId="(\d+)"', numbering_xml)]
    decimal_abs = None
    for m in re.finditer(
            r'<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"(.*?)</w:abstractNum>',
            numbering_xml, re.S):
        lvl0 = re.search(r'<w:lvl\b[^>]*w:ilvl="0".*?</w:lvl>', m.group(2), re.S)
        if lvl0 is not None and '<w:numFmt w:val="decimal"' in lvl0.group(0):
            decimal_abs = m.group(1)
            break
    return (max(nums) if nums else 0), decimal_abs


def make_list_num_xml(num_id, abstract_id):
    """生成一个新的 decimal 自动编号实例（独立计数，从 1 开始）。"""
    return (f'<w:num w:numId="{num_id}">'
            f'<w:abstractNumId w:val="{abstract_id}"/>'
            f'<w:lvlOverride w:ilvl="0">'
            f'<w:startOverride w:val="1"/></w:lvlOverride></w:num>')


def free_toc_case(styles_xml):
    """去掉目录样式的强制大小写，让目录条目按原样显示。

    模板给 TOC1 配了 caps、给 TOC2 配了 smallCaps（西文目录的全大写习惯），
    落在中文目录上会把条目里的英文专有名词一并转换——`Android` 渲染成
    `ANDROID`、`Magisk` 渲染成 `MAGISK`、`SELinux` 渲染成 `SELINUX`，
    与正文的命名规范不一致。目录本身是 TOC 域，在 Word 里按 F9 重建后仍会
    套用这两个样式，所以必须在样式上改，逐个 run 覆盖是治标。
    """
    for sid, tag in (("TOC1", "caps"), ("TOC2", "smallCaps")):
        m = re.search(
            rf'<w:style [^>]*w:styleId="{sid}"[^>]*>.*?</w:style>', styles_xml, re.S)
        if not m:
            continue
        block = m.group(0)
        fixed = block.replace(f"<w:{tag}/>", "").replace(f"<w:{tag} />", "")
        if fixed != block:
            styles_xml = styles_xml.replace(block, fixed)
    return styles_xml


def toc_entry_xml(level, text, page):
    """生成一条静态目录项：标题 + 制表位 + 页码。"""
    style = "TOC1" if level == 1 else "TOC2"
    ppr = (f'<w:pStyle w:val="{style}"/>'
           f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{CONTENT_W}"/></w:tabs>')
    runs = runs_xml(text)
    runs += ('<w:r><w:tab/></w:r>'
             f'<w:r><w:t>{esc(str(page))}</w:t></w:r>')
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def shrink_toc(kids, log, outline=(), pages=None):
    """用静态目录项替换目录域的缓存结果。

    目录仍是一个 TOC 域，在 Word 里按 F9 会重新生成（页码由 Word 的排版引擎算出）；
    这里预先填好一版，是为了让不更新域的阅读器（LibreOffice、部分预览工具）也能
    看到完整目录。pages 为 {标题: 页码}，缺省时页码留空。
    """
    begin = None
    for i, c in enumerate(kids):
        if c.tag != f"{W}p":
            continue
        if any(it.text and "TOC" in it.text for it in c.iter(f"{W}instrText")):
            begin = i
            break
    if begin is None:
        log.append("目录域结构未识别，跳过（目录保持模板原样）")
        return None

    # 目录域内部还嵌着每条目录项的 PAGEREF 域，必须按嵌套深度找它自己的结束位置
    depth, end = 0, None
    for i in range(begin, len(kids)):
        c = kids[i]
        if c.tag != f"{W}p":
            continue
        for fc in c.iter(f"{W}fldChar"):
            t = fc.get(f"{W}fldCharType")
            if t == "begin":
                depth += 1
            elif t == "end":
                depth -= 1
                if depth == 0:
                    end = i
        if end is not None:
            break
    if end is None or end <= begin + 1:
        log.append("目录域结构未识别，跳过（目录保持模板原样）")
        return None

    old = end - begin - 1
    if not outline:
        hint = para_xml("【目录为自动域：请在 Word 中按 Ctrl+A 全选后按 F9 更新】",
                        style="TOC1")
        log.append(f"目录域缓存结果已清空（原 {old} 条），待 F9 重建")
        return kids[:begin + 1] + [("RAW", hint)] + kids[end:]

    pages = pages or {}
    filled = sum(1 for lv, t in outline if t in pages)
    entries = [("RAW", toc_entry_xml(lv, t, pages.get(t, ""))) for lv, t in outline]
    log.append(f"目录已填入 {len(entries)} 条静态条目"
               f"（原 {old} 条），其中 {filled} 条带页码；"
               f"在 Word 中按 F9 可重建为域结果")
    return kids[:begin + 1] + entries + kids[end:]


def add_update_fields(settings_xml):
    """在 settings.xml 的合法位置插入 updateFields，使 Word 打开时自动更新域。"""
    if "updateFields" in settings_xml:
        return settings_xml
    # 按 CT_Settings 的元素顺序，updateFields 紧接在 hdrShapeDefaults 之前
    for anchor in ("<w:hdrShapeDefaults", "<w:footnotePr", "<w:endnotePr",
                   "<w:compat", "<w:rsids", "</w:settings>"):
        pos = settings_xml.find(anchor)
        if pos != -1 and anchor != "</w:settings>":
            return (settings_xml[:pos] + '<w:updateFields w:val="true"/>'
                    + settings_xml[pos:])
    pos = settings_xml.rindex("</w:settings>")
    return settings_xml[:pos] + '<w:updateFields w:val="true"/>' + settings_xml[pos:]

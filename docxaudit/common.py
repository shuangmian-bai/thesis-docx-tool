"""公共层：docx 只读封装、审计上下文、路径定位与 XML 比对工具。

检查项（`checks/`）与报告层（`report`）都建立在这一层上。这里的工具分三类：

- **读**：`Docx`（带缓存的 zip 只读封装）、`Ctx`（一次审计的全部输入）；
- **定位**：`find_template` / `find_product`；
- **比对**：`iter_blocks` / `iter_paras` / `pstyle` / `tag_counter` / `norm` /
  `skeleton` / `reserved_parts` / `tail_sectpr`。
"""
import datetime
import os
import re
import zipfile
from collections import Counter
from dataclasses import dataclass

from docxaudit import HERE

CHAP_DIR = os.path.join(HERE, "config", "章节")

OK, EXPECTED, WARN, FAIL = "ok", "expected", "warn", "fail"
# 状态标记一律用纯文字，不用符号（开源子项目硬规则：禁止装饰符号）
LEVEL_MARK = {OK: "[通过]", EXPECTED: "[有意偏离]", WARN: "[注意]", FAIL: "[失败]"}

# ── 模块级正则 ────────────────────────────────────────────────────
BLOCK_RE = re.compile(r"<w:(p|tbl)\b(?![^>]*/>).*?</w:\1>", re.S)
PARA_RE = re.compile(r"<w:p\b(?![^>]*/>).*?</w:p>", re.S)
STYLE_RE = re.compile(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>.*?</w:style>', re.S)
DEFAULTS_RE = re.compile(r"<w:docDefaults>.*?</w:docDefaults>", re.S)
LATENT_RE = re.compile(r"<w:latentStyles[^>]*>.*?</w:latentStyles>|<w:latentStyles[^>]*/>", re.S)
TAG_RE = re.compile(r"<w:(\w+)[ />]")
XMLNS_DECL_RE = re.compile(r'\s+xmlns:([A-Za-z0-9]+)="([^"]*)"')


@dataclass
class Finding:
    """一条检查结论。

    `key` 是差异的稳定标识，用来匹配 `EXPECTED` 表；`msg` 是一句话结论；
    `detail` 是排查用的展开内容，只在需要时打印。
    """
    key: str
    level: str
    msg: str
    detail: str = ""


class Docx:
    """一个 docx（zip）的只读封装，带文本与条目名缓存。

    一次审计要反复读同一批部件（styles.xml 会被 styles / body / meta 三项轮流读），
    每次 `read().decode()` 都是解压 + 解码，缓存下来省一大截。
    """

    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path)
        self._z = zipfile.ZipFile(path)
        self._names = None
        self._nameset = None
        self._text = {}

    def names(self):
        if self._names is None:
            self._names = self._z.namelist()
        return self._names

    def has(self, part):
        if self._nameset is None:
            self._nameset = set(self.names())
        return part in self._nameset

    def raw(self, part):
        return self._z.read(part)

    def text(self, part):
        if part not in self._text:
            self._text[part] = self._z.read(part).decode("utf-8")
        return self._text[part]

    def close(self):
        self._z.close()


class Ctx:
    """一次审计的全部输入。"""

    def __init__(self, tpl_path, prod_path, use_expected=True):
        self.tpl_path = tpl_path
        self.prod_path = prod_path
        self.tpl = Docx(tpl_path)
        self.prod = Docx(prod_path)
        # 换过模板后，`EXPECTED` 表里针对本模板的登记就不再适用
        self.use_expected = use_expected

    def close(self):
        self.tpl.close()
        self.prod.close()

    def doc_xml(self, which):
        return (self.tpl if which == "tpl" else self.prod).text("word/document.xml")

    def styles_xml(self, which):
        return (self.tpl if which == "tpl" else self.prod).text("word/styles.xml")


def find_template(path=None):
    """定位模板 docx。

    默认取 `config/template.docx`（使用者自备，不入库）；也可通过 `--template` 指定。
    """
    if path:
        p = path if os.path.isabs(path) else os.path.abspath(path)
    else:
        p = os.path.join(HERE, "config", "template.docx")
    if not os.path.exists(p):
        raise SystemExit(f"找不到模板：{p}（请自备模板放到 config/template.docx，或用 --template 指定）")
    return p


def find_product(path=None):
    """定位成品 docx。

    默认取 `output/` 下 **mtime 最新**的 `*.docx`。
    按 mtime 而非文件名字典序——`v10` 字典序排在 `v9` 前面，会取错。
    """
    if path:
        p = path if os.path.isabs(path) else os.path.join(HERE, path)
        if not os.path.exists(p):
            raise SystemExit(f"找不到成品：{p}")
        return p
    out_dir = os.path.join(HERE, "output")
    if not os.path.isdir(out_dir):
        raise SystemExit(f"output/ 目录不存在，先跑 build 子命令")
    cands = [os.path.join(out_dir, f) for f in os.listdir(out_dir)
             if f.lower().endswith(".docx")]
    if not cands:
        raise SystemExit(f"output/ 下没有成品 docx，先跑 build 子命令")
    return max(cands, key=os.path.getmtime)


def iter_blocks(raw):
    """按文档顺序取出顶层块：`<w:p>` 或整个 `<w:tbl>`。

    表格算一块，内部的单元格段落不单独返回——块序列比对要的是「第 N 块」这种
    一一对应关系，把表格拆开会与模板对不齐。
    """
    return [m.group(0) for m in BLOCK_RE.finditer(raw)]


def iter_paras(raw):
    """取出**所有**段落，含表格单元格里的。用于按样式统计，不做块对齐。"""
    return [m.group(0) for m in PARA_RE.finditer(raw)]


def pstyle(block):
    """取段落的样式 ID，没有则返回 None。"""
    m = re.search(r'<w:pStyle w:val="([^"]+)"', block)
    return m.group(1) if m else None


def tag_counter(xml):
    return Counter(TAG_RE.findall(xml))


def norm(x):
    """规范化 XML 书写风格：自闭合标签前空格、标签间空白、属性间多余空格。

    只做无副作用的白化，用来区分「序列化风格差异」与「真正的格式差异」。
    """
    x = re.sub(r"\s*/>", "/>", x)
    x = re.sub(r">\s+<", "><", x)
    x = re.sub(r'"\s+', '" ', x)
    x = re.sub(r'\s+(?=[\w:]+=")', " ", x)
    return x.strip()


def skeleton(x):
    """抽出保留区的**结构骨架**：去掉文本内容、命名空间声明与图形尺寸，只留标签与属性。

    保留区里封面字段、承诺书题目、目录条目都是按论文填的，签名图也是按作者换的，
    文本与图形尺寸必然与模板不同；要比的是「结构与格式有没有被破坏」，
    不是「文字或图一不一样」。
    """
    x = XMLNS_DECL_RE.sub("", x)
    # 文本标签整体规范化成 <w:t/>：空文本在 OOXML 里有三种等价写法——
    # `<w:t/>`、`<w:t></w:t>`、`<w:t xml:space="preserve"></w:t>`，Word 与 ElementTree
    # 各写一种（`set_para_text()` 会给首 run 加 xml:space）。只清内容不规范化标签本身，
    # 会把这种序列化差异误报成结构差异。
    # `(?:(?!/>)[^>])*` 用来区分自闭合与成对形态：若图省事写成 `[^>]*>`，`[^>]*` 会把
    # `<w:t/>` 里的 `/` 一并吃掉、当成开标签，再向后找下一个 `</w:t>`，把中间内容整段删掉。
    for tag in ("w:t", "w:instrText"):
        x = re.sub(rf"<{tag}\b(?:(?!/>)[^>])*(?:/>|>.*?</{tag}>)", f"<{tag}/>", x, flags=re.S)
    # 图形尺寸也不比：承诺书的签名图按论文作者替换后，`wp:extent` 与 `a:ext` 上的
    # cx / cy 必然与模板不同（模板那张是示例签名）。尺寸属「图怎么摆」，
    # 是内容层的事；图本身由 parts 层的媒体检查覆盖。这里仍管结构与其余属性。
    for _ in range(2):
        x = re.sub(r'(<(?:wp:extent|a:ext)\b[^>]*?)\s+c[xy]="\d+"', r"\1", x)
    return norm(x)


def reserved_parts(raw):
    """把保留区拆成三段：`(目录之前, 目录段, 分节符段)`。

    保留区 = `<w:body>` 开头到含 `sectPr` 的段落，构建时原样搬过来。但中间那段
    是 TOC 域，目录条目由 `build_docx.py` 按论文重写（本论文 65 条），条目数必然
    与模板示例目录不同——拿它比块数没有意义，故单独拎出来只报数、不判定。
    """
    b0 = raw.index("<w:body")
    b0 = raw.index(">", b0) + 1
    sect = raw.index("<w:sectPr", b0)
    ms = list(re.finditer(r"<w:p[ >/]", raw[b0:sect]))
    sect_p0 = b0 + ms[-1].start()
    sect_p1 = raw.index("</w:p>", sect) + len("</w:p>")

    # 目录域的起点 = 含 `instrText TOC` 的那个段落
    toc = re.search(r"<w:instrText[^>]*>[^<]*\bTOC\b", raw[b0:sect_p0])
    if toc:
        head = list(re.finditer(r"<w:p[ >/]", raw[b0:toc.start()]))
        cut = b0 + head[-1].start() if head else b0 + toc.start()
    else:
        cut = sect_p0
    return raw[b0:cut], raw[cut:sect_p0], raw[sect_p0:sect_p1]


def tail_sectpr(raw):
    """body 末尾那个带页眉页脚引用的 `sectPr`——页面设置的真身。"""
    i = raw.rindex("<w:sectPr")
    m = re.search(r"<w:sectPr\b[^>]*/>|<w:sectPr\b.*?</w:sectPr>", raw[i:], re.S)
    return m.group(0) if m else raw[i:]


def fmt_ts(t):
    return datetime.datetime.fromtimestamp(t).strftime("%Y-%m-%d %H:%M")


def rel_delta(sec):
    if sec < 90:
        return f"{int(sec)} 秒"
    if sec < 5400:
        return f"{sec / 60:.1f} 分钟"
    return f"{sec / 3600:.1f} 小时"

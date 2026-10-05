"""严格对比：原文与合成的正文区，内容与格式完全对应。

不标记、不跳过、不容差——任何差异都是扁平化器或渲染器的缺陷，对比失败
即修复对应模块，不允许用「预期差异」掩盖。
"""
import re
import zipfile
from dataclasses import dataclass, field
from typing import List

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# w:t 可能带 xml:space 等属性；`[^>]*` 会误吞 <w:tabs> 等同前缀标签，
# 必须限定属性前是空白
T_RE = r"<w:t(?:\s[^>]*)?>(.*?)</w:t>"


@dataclass
class Diff:
    area: str            # body
    location: str        # 位置描述（段落序号.字段 等）
    src: str             # 原值
    out: str             # 输出值


@dataclass
class CompareResult:
    src_path: str
    out_path: str
    diffs: List[Diff] = field(default_factory=list)

    @property
    def total(self) -> int:
        return len(self.diffs)

    @property
    def passed(self) -> bool:
        return not self.diffs


def _read_xml(z, name):
    try:
        return z.read(name).decode("utf-8")
    except KeyError:
        return ""


def _attr(elem_xml, tag, val_attr="w:val"):
    m = re.search(rf"<w:{tag}[^>]*{val_attr}=\"([^\"]+)\"", elem_xml)
    return m.group(1) if m else None


def _split_body(doc_xml):
    """按分节符段落把 body 切成 [前置区（含分节符）, 正文区]。

    与 build 的骨架边界同标准：`pPr` 内含 `sectPr` 的段落是骨架的最后一个
    元素，其后才是正文区（扁平化→合成的作用范围）。
    """
    body_m = re.search(r"<w:body>(.*?)</w:body>", doc_xml, re.S)
    body = body_m.group(1) if body_m else ""
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>|<w:p\b[^>]*/>", body, re.S):
        p = m.group(0)
        ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
        if ppr and "<w:sectPr" in ppr.group(0):
            return body[:m.end()], body[m.end():]
    return "", body


def _extract_paras(xml):
    """提取段落列表，每段含 text 和全部格式属性。"""
    paras = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.S):
        p = m.group(0)
        text = "".join(re.findall(T_RE, p, re.S)).strip()
        run_m = re.search(r"<w:r\b.*?</w:r>", p, re.S)
        run = run_m.group(0) if run_m else ""
        paras.append({
            "text": text,
            "pstyle": _attr(p, "pStyle"),
            "jc": _attr(p, "jc"),
            "spacing_line": _attr(p, "spacing", "w:line"),
            "spacing_before": _attr(p, "spacing", "w:before"),
            "spacing_after": _attr(p, "spacing", "w:after"),
            "ind_left": _attr(p, "ind", "w:left"),
            "ind_firstLine": _attr(p, "ind", "w:firstLine"),
            "rfonts": _attr(run, "rFonts", "w:eastAsia") or _attr(run, "rFonts", "w:ascii"),
            "sz": _attr(run, "sz"),
            "b": "1" if re.search(r"<w:b/>|<w:b w:val=\"1\"/>", run) else "0",
        })
    return paras


def _extract_tables(xml):
    """提取表格：单元格文本与段落格式（jc/ind 前缀），保留合并信息的差异可见。"""
    tables = []
    for m in re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", xml, re.S):
        tbl = m.group(0)
        rows = []
        for tr in re.findall(r"<w:tr[ >].*?</w:tr>", tbl, re.S):
            cells = []
            for tc in re.findall(r"<w:tc[ >].*?</w:tc>", tr, re.S):
                cell_paras = []
                for pm in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", tc, re.S):
                    p = pm.group(0)
                    t = "".join(re.findall(T_RE, p, re.S)).strip()
                    if not t:
                        continue
                    jc = _attr(p, "jc")
                    ind = _attr(p, "ind", "w:firstLine")
                    fmt = []
                    if jc:
                        fmt.append(f"jc={jc}")
                    if ind:
                        fmt.append(f"ind={ind}")
                    cell_paras.append(
                        ("{" + ",".join(fmt) + "}" if fmt else "") + t)
                cells.append("\n".join(cell_paras))
            rows.append(cells)
        tables.append(rows)
    return tables


def _extract_images(xml):
    """提取图片尺寸列表（cx, cy EMU）。"""
    return [tuple(int(x) for x in em.groups())
            for em in re.finditer(r"<wp:extent cx=\"(\d+)\" cy=\"(\d+)\"", xml)]


def compare(src_path: str, out_path: str) -> CompareResult:
    """严格对比原文与合成的正文区：内容与格式完全对应。"""
    result = CompareResult(src_path=src_path, out_path=out_path)

    with zipfile.ZipFile(src_path) as zs, zipfile.ZipFile(out_path) as zo:
        src_doc = _read_xml(zs, "word/document.xml")
        out_doc = _read_xml(zo, "word/document.xml")
        _, src_body = _split_body(src_doc)
        _, out_body = _split_body(out_doc)

        # 段落：文本 + 全部格式（样式/对齐/行距/段距/缩进/字体/字号/加粗）
        src_paras = [p for p in _extract_paras(src_body) if p["text"]]
        out_paras = [p for p in _extract_paras(out_body) if p["text"]]
        keys = ["text", "pstyle", "jc", "spacing_line", "spacing_before",
                "spacing_after", "ind_left", "ind_firstLine", "rfonts", "sz", "b"]
        for i, (a, b) in enumerate(zip(src_paras, out_paras)):
            for k in keys:
                if a[k] != b[k]:
                    result.diffs.append(Diff(
                        "body", f"段落#{i}.{k}", f"{a[k]}", f"{b[k]}"))
        if len(src_paras) != len(out_paras):
            result.diffs.append(Diff(
                "body", "段落总数", str(len(src_paras)), str(len(out_paras))))

        # 表格：单元格文本与段落格式
        src_tbls = _extract_tables(src_body)
        out_tbls = _extract_tables(out_body)
        for i, (a, b) in enumerate(zip(src_tbls, out_tbls)):
            if a != b:
                result.diffs.append(Diff(
                    "body", f"表格#{i}", str(a)[:120], str(b)[:120]))
        if len(src_tbls) != len(out_tbls):
            result.diffs.append(Diff(
                "body", "表格总数", str(len(src_tbls)), str(len(out_tbls))))

        # 图片尺寸
        src_imgs = _extract_images(src_body)
        out_imgs = _extract_images(out_body)
        if src_imgs != out_imgs:
            result.diffs.append(Diff(
                "body", "图片尺寸", str(src_imgs)[:80], str(out_imgs)[:80]))

    return result

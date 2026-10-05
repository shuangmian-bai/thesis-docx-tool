"""全盘对比：原文档 vs 输出文档的所有内容与格式。

对比项分六个区域：
1. 正文区（分节符段落之后）：段落文本、格式属性、表格、图片尺寸
2. 封面/承诺书（前置区非目录段落）：文本
3. 目录（TOC 域缓存条目）：文本与数量必须一致；页码与制表位为预期差异
4. 页眉/页脚：文本
5. 样式表：关键样式定义

骨架边界与 build 同标准：`pPr` 内含 `sectPr` 的分节符段落之前为前置区
（封面、承诺书、目录域），之后为正文区。

每个差异生成 Diff 对象，包含：区域、位置、原值、输出值、是否预期、预期原因。
预期差异规则内联在本模块（`check_expected`），不再依赖独立规则文件。
"""
import re
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# w:t 可能带 xml:space 等属性；`[^>]*` 会误吞 <w:tabs> 等同前缀标签，
# 必须限定属性前是空白
T_RE = r"<w:t(?:\s[^>]*)?>(.*?)</w:t>"


@dataclass
class Diff:
    area: str            # body / cover / toc / header / footer / styles
    location: str        # 位置描述（段落序号/字段名等）
    src: str             # 原值
    out: str             # 输出值
    expected: bool = False
    expected_reason: str = ""


@dataclass
class CompareResult:
    src_path: str
    out_path: str
    diffs: List[Diff] = field(default_factory=list)
    stats: dict = field(default_factory=dict)

    @property
    def total(self) -> int:
        return len(self.diffs)

    @property
    def unexpected(self) -> int:
        return sum(1 for d in self.diffs if not d.expected)

    @property
    def passed(self) -> bool:
        return self.unexpected == 0


def check_expected(area: str, what: str):
    """登记已知且接受的偏离。返回 (是否预期, 原因)。

    目录页码与制表位由 Word 排版引擎在打开/重建域时计算，静态回填值与
    原文不可能逐字相同，属结构性预期差异。
    """
    if area == "toc" and what in ("页码", "制表位"):
        return True, "页码由排版引擎计算；Word 中 Ctrl+A 后 F9 重建即同步"
    return False, ""


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

    与 build 的骨架边界同标准：`pPr` 内含 `sectPr` 的段落是骨架的最后
    一个元素，其后才是模板示例正文（构建时被替换的部分）。源文档若没有
    分节符段落，则整篇按正文处理。
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
    """提取段落列表，每段含 text 和关键格式属性。"""
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
            "raw": p,
        })
    return paras


def _extract_tables(xml):
    tables = []
    for m in re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", xml, re.S):
        tbl = m.group(0)
        rows = []
        for tr in re.findall(r"<w:tr[ >].*?</w:tr>", tbl, re.S):
            cells = []
            for tc in re.findall(r"<w:tc[ >].*?</w:tc>", tr, re.S):
                t = "".join(re.findall(T_RE, tc, re.S)).strip()
                cells.append(t)
            rows.append(cells)
        tables.append(rows)
    return tables


def _extract_images(xml):
    """提取图片尺寸列表（cx, cy EMU）及所在段落文本。"""
    imgs = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.S):
        p = m.group(0)
        text = "".join(re.findall(T_RE, p, re.S)).strip()
        for em in re.finditer(r"<wp:extent cx=\"(\d+)\" cy=\"(\d+)\"", p):
            imgs.append({"cx": int(em.group(1)), "cy": int(em.group(2)), "near": text})
    return imgs


def _extract_toc_entries(front_xml):
    """提取目录域缓存的静态条目：[(pstyle, title, page, tab_pos)]。

    静态条目是 pStyle 为 TOC1/TOC2 的段落：标题文本 + 右对齐制表位 + 页码。
    页码取段落内最后一个纯数字文本节点；制表位取 w:tabs 中最后一个 w:tab
    的 pos。对比语义：标题、级别、数量必须一致；页码与制表位由排版引擎
    决定，差异记预期。
    """
    entries = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", front_xml, re.S):
        p = m.group(0)
        pstyle = _attr(p, "pStyle")
        if pstyle not in ("TOC1", "TOC2"):
            continue
        texts = [t for t in re.findall(T_RE, p, re.S) if t.strip()]
        page = ""
        if texts and re.fullmatch(r"\d+", texts[-1].strip()):
            page = texts.pop().strip()
        title = "".join(texts).strip()
        tabs = re.findall(r"<w:tab [^>]*w:pos=\"(\d+)\"", p)
        entries.append((pstyle, title, page, tabs[-1] if tabs else ""))
    return entries


def _extract_hf(z, part_name):
    """提取页眉/页脚文本。"""
    xml = _read_xml(z, part_name)
    if not xml:
        return ""
    return " | ".join(t for t in re.findall(T_RE, xml, re.S) if t.strip())


def compare(src_path: str, out_path: str) -> CompareResult:
    """全盘对比原文档与输出文档，返回 CompareResult。"""
    result = CompareResult(src_path=src_path, out_path=out_path)

    with zipfile.ZipFile(src_path) as zs, zipfile.ZipFile(out_path) as zo:
        src_doc = _read_xml(zs, "word/document.xml")
        out_doc = _read_xml(zo, "word/document.xml")

        src_front, src_body = _split_body(src_doc)
        out_front, out_body = _split_body(out_doc)

        # ===== 1. 正文区对比 =====
        src_paras = [p for p in _extract_paras(src_body) if p["text"]]
        out_paras = [p for p in _extract_paras(out_body) if p["text"]]
        fmt_keys = ["pstyle", "jc", "spacing_line", "spacing_before", "spacing_after",
                    "ind_left", "ind_firstLine", "rfonts", "sz", "b"]
        for i, (a, b) in enumerate(zip(src_paras, out_paras)):
            if a["text"] != b["text"]:
                result.diffs.append(Diff(
                    "body", f"段落#{i}", a["text"][:60], b["text"][:60]))
            else:
                for k in fmt_keys:
                    if a[k] != b[k]:
                        result.diffs.append(Diff(
                            "body", f"段落#{i}.{k}",
                            f"{a[k]}", f"{b[k]}"))
        if len(src_paras) != len(out_paras):
            result.diffs.append(Diff(
                "body", "段落总数", str(len(src_paras)), str(len(out_paras))))

        # 正文表格
        src_tbls = _extract_tables(src_body)
        out_tbls = _extract_tables(out_body)
        for i, (a, b) in enumerate(zip(src_tbls, out_tbls)):
            if a != b:
                result.diffs.append(Diff(
                    "body", f"表格#{i}", str(a)[:80], str(b)[:80]))
        if len(src_tbls) != len(out_tbls):
            result.diffs.append(Diff(
                "body", "表格总数", str(len(src_tbls)), str(len(out_tbls))))

        # 正文图片（尺寸应原样保留，差异一律记非预期）
        src_imgs = _extract_images(src_body)
        out_imgs = _extract_images(out_body)
        for i, (a, b) in enumerate(zip(src_imgs, out_imgs)):
            if (a["cx"], a["cy"]) != (b["cx"], b["cy"]):
                result.diffs.append(Diff(
                    "body", f"图片#{i}尺寸",
                    f"{a['cx']}x{a['cy']}", f"{b['cx']}x{b['cy']}"))
        if len(src_imgs) != len(out_imgs):
            result.diffs.append(Diff(
                "body", "图片总数", str(len(src_imgs)), str(len(out_imgs))))

        # ===== 2. 封面/承诺书（前置区非目录段落） =====
        # 目录条目由下面的 TOC 专项对比负责，此处排除避免重复报告
        src_front_paras = _extract_paras(src_front)
        out_front_paras = _extract_paras(out_front)
        src_cover = [p["text"] for p in src_front_paras
                     if p["text"] and p["pstyle"] not in ("TOC1", "TOC2")]
        out_cover = [p["text"] for p in out_front_paras
                     if p["text"] and p["pstyle"] not in ("TOC1", "TOC2")]
        for i, (a, b) in enumerate(zip(src_cover, out_cover)):
            if a != b:
                result.diffs.append(Diff("cover", f"前置段落#{i}", a[:60], b[:60]))
        if len(src_cover) != len(out_cover):
            result.diffs.append(Diff(
                "cover", "前置段落总数", str(len(src_cover)), str(len(out_cover))))

        # ===== 3. 目录（标记检查：文本与数量必须一致，页码/制表位为预期差异） =====
        src_toc = _extract_toc_entries(src_front)
        out_toc = _extract_toc_entries(out_front)
        if len(src_toc) != len(out_toc):
            result.diffs.append(Diff(
                "toc", "条目总数", str(len(src_toc)), str(len(out_toc))))
        for i, (a, b) in enumerate(zip(src_toc, out_toc)):
            if a[1] != b[1]:
                # 去空白后相同：原文目录域缓存未随标题编辑更新（多见于编号后
                # 空格的有无），F9 重建即同步，属预期；文字实质不同则非预期
                if re.sub(r"\s+", "", a[1]) == re.sub(r"\s+", "", b[1]):
                    exp, reason = True, ("原文目录域缓存未随标题编辑更新（仅空白差异）；"
                                         "F9 重建后与正文标题一致")
                else:
                    exp, reason = False, ""
                result.diffs.append(Diff(
                    "toc", f"条目#{i}文本", a[1][:60], b[1][:60],
                    expected=exp, expected_reason=reason))
            if a[0] != b[0]:
                result.diffs.append(Diff("toc", f"条目#{i}级别", a[0], b[0]))
            if a[2] != b[2]:
                exp, reason = check_expected("toc", "页码")
                result.diffs.append(Diff(
                    "toc", f"条目#{i}页码", a[2], b[2],
                    expected=exp, expected_reason=reason))
            if a[3] != b[3]:
                exp, reason = check_expected("toc", "制表位")
                result.diffs.append(Diff(
                    "toc", f"条目#{i}制表位", a[3], b[3],
                    expected=exp, expected_reason=reason))

        # ===== 4. 页眉/页脚 =====
        for part in ["word/header1.xml", "word/footer1.xml"]:
            src_hf = _extract_hf(zs, part)
            out_hf = _extract_hf(zo, part)
            area = "header" if "header" in part else "footer"
            if src_hf != out_hf:
                result.diffs.append(Diff(area, part, src_hf[:80], out_hf[:80]))

        # ===== 5. 样式表关键样式 =====
        src_styles = _read_xml(zs, "word/styles.xml")
        out_styles = _read_xml(zo, "word/styles.xml")
        for sid in ["a0", "aa", "1", "2", "3", "4"]:
            src_m = re.search(rf'<w:style w:type="paragraph" w:styleId="{sid}">.*?</w:style>',
                              src_styles, re.S)
            out_m = re.search(rf'<w:style w:type="paragraph" w:styleId="{sid}">.*?</w:style>',
                              out_styles, re.S)
            if src_m and out_m:
                src_sz = _attr(src_m.group(0), "sz")
                out_sz = _attr(out_m.group(0), "sz")
                if src_sz != out_sz:
                    result.diffs.append(Diff(
                        "styles", f"样式 {sid} 字号", str(src_sz), str(out_sz)))

    result.stats = {
        "body_paras": len(src_paras),
        "body_tables": len(src_tbls),
        "body_images": len(src_imgs),
        "cover_paras": len(src_cover),
        "toc_entries": len(src_toc),
    }
    return result

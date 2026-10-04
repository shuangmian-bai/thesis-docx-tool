"""全盘对比：原文档 vs 输出文档的所有内容与格式。

对比项分六个区域：
1. 正文区（一级标题之后）：段落文本、格式属性、表格、图片尺寸
2. 封面：字段文本
3. 目录：条目文本与数量
4. 承诺书：签名图（标记预期差异）
5. 页眉/页脚：文本
6. 样式表：关键样式定义

每个差异生成 Diff 对象，包含：区域、位置、原值、输出值、是否预期、预期原因。
"""
import re
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional

from docxloop.rules import check_expected

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass
class Diff:
    area: str            # body / cover / toc / promise / header / footer / styles
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


def _read_xml(z, name):
    try:
        return z.read(name).decode("utf-8")
    except KeyError:
        return ""


def _attr(elem_xml, tag, val_attr="w:val"):
    m = re.search(rf"<w:{tag}[^>]*{val_attr}=\"([^\"]+)\"", elem_xml)
    return m.group(1) if m else None


def _split_body(doc_xml):
    """把 document.xml 的 body 按一级标题切成 [前置, 正文]。

    一级标题用 pStyle="1" 判定（封面、目录、承诺书都在前置区）。
    """
    body_m = re.search(r"<w:body>(.*?)</w:body>", doc_xml, re.S)
    body = body_m.group(1) if body_m else ""
    split_idx = 0
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", body, re.S):
        pstyle = _attr(m.group(0), "pStyle")
        if pstyle == "1":
            split_idx = m.start()
            break
    return body[:split_idx], body[split_idx:]


def _extract_paras(xml):
    """提取段落列表，每段含 text 和关键格式属性。"""
    paras = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.S):
        p = m.group(0)
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p)).strip()
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
                t = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", tc)).strip()
                cells.append(t)
            rows.append(cells)
        tables.append(rows)
    return tables


def _extract_images(xml):
    """提取图片尺寸列表（cx, cy EMU）及所在段落文本。"""
    imgs = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.S):
        p = m.group(0)
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p)).strip()
        for em in re.finditer(r"<wp:extent cx=\"(\d+)\" cy=\"(\d+)\"", p):
            imgs.append({"cx": int(em.group(1)), "cy": int(em.group(2)), "near": text})
    return imgs


def _extract_hf(z, part_name):
    """提取页眉/页脚文本。"""
    xml = _read_xml(z, part_name)
    if not xml:
        return ""
    return " | ".join(t for t in re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml) if t.strip())


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

        # 正文图片
        src_imgs = _extract_images(src_body)
        out_imgs = _extract_images(out_body)
        for i, (a, b) in enumerate(zip(src_imgs, out_imgs)):
            if (a["cx"], a["cy"]) != (b["cx"], b["cy"]):
                exp, reason = check_expected("img", b)
                result.diffs.append(Diff(
                    "body", f"图片#{i}尺寸",
                    f"{a['cx']}x{a['cy']}", f"{b['cx']}x{b['cy']}",
                    expected=exp, expected_reason=reason))
        if len(src_imgs) != len(out_imgs):
            result.diffs.append(Diff(
                "body", "图片总数", str(len(src_imgs)), str(len(out_imgs))))

        # ===== 2. 封面/目录/承诺书（前置区）对比 =====
        src_front_paras = _extract_paras(src_front)
        out_front_paras = _extract_paras(out_front)
        # 封面：对比非空文本段落
        src_cover = [p["text"] for p in src_front_paras if p["text"]]
        out_cover = [p["text"] for p in out_front_paras if p["text"]]
        for i, (a, b) in enumerate(zip(src_cover, out_cover)):
            if a != b:
                exp, reason = check_expected("cover", {"text": b})
                result.diffs.append(Diff(
                    "cover", f"前置段落#{i}", a[:60], b[:60],
                    expected=exp, expected_reason=reason))
        if len(src_cover) != len(out_cover):
            result.diffs.append(Diff(
                "cover", "前置段落总数", str(len(src_cover)), str(len(out_cover))))

        # ===== 3. 页眉/页脚 =====
        for part in ["word/header1.xml", "word/footer1.xml"]:
            src_hf = _extract_hf(zs, part)
            out_hf = _extract_hf(zo, part)
            area = "header" if "header" in part else "footer"
            if src_hf != out_hf:
                result.diffs.append(Diff(area, part, src_hf[:80], out_hf[:80]))

        # ===== 4. 样式表关键样式 =====
        src_styles = _read_xml(zs, "word/styles.xml")
        out_styles = _read_xml(zo, "word/styles.xml")
        # 对比 aa / a0 / 1 / 2 样式的 sz 和 jc
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
    }
    return result

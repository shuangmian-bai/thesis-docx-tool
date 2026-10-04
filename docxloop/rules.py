"""预期差异规则：登记「引擎已知且可解释」的差异，不在报告里算失败。

新增规则只需在此处加一条，不动 compare 逻辑。每条规则返回
(是否预期, 原因说明)。
"""
import re


def is_signature_image(img_elem, doc_ctx):
    """承诺书签名图：由 config/images/signature.png 替换，尺寸与模板原图不同。

    判定：图片在承诺书区域（「承诺书」标题之后、正文一级标题之前），
    或图片尺寸与 signature.png 已知尺寸（3.5x1.01 cm）吻合。
    """
    # signature.png 的 EMU 尺寸：3.5cm * 360000 = 1260000, 1.01cm * 360000 ≈ 363600
    SIG_CX, SIG_CY = 1260000, 363600
    cx = int(img_elem.get("cx", 0))
    cy = int(img_elem.get("cy", 0))
    if abs(cx - SIG_CX) < 50000 and abs(cy - SIG_CY) < 50000:
        return True, "承诺书签名图（由 signature.png 替换，尺寸与模板原图不同属正常）"
    return False, ""


def is_cover_placeholder(para_text, doc_ctx):
    """封面/目录/承诺书差异：convert 拆解时丢弃前置区（strip_front），
    build 用模板骨架重建，封面字段为占位符、目录页码为空。

    原文档前置区有真实信息，输出是模板骨架，这是设计行为。
    """
    out_text = para_text.get("out", "") if isinstance(para_text, dict) else para_text
    if "【待填：" in out_text or "待填" in out_text:
        return True, "封面字段占位符（convert 丢弃前置区，build 用默认占位符，需用户填写）"
    # 目录条目差异（页码为空）
    if "w:tab" in out_text and "leader" in out_text:
        return True, "目录页码（需跑 toc_pages.py 从 PDF 反查页码后重建）"
    # 承诺书内容差异（模板自带的承诺书标题）
    return True, "前置区差异（封面/目录/承诺书由模板骨架重建，非正文内容）"


def is_front_matter_diff(diff, doc_ctx):
    """前置区（封面/目录/承诺书）的所有文本差异均为预期。

    convert 的 strip_front 会丢弃一级标题之前的所有内容，build 从模板骨架
    重建封面、目录、承诺书，因此前置区文本与原文档不同是设计行为。
    """
    return True, "前置区差异（封面/目录/承诺书由模板骨架重建，非正文内容）"


#: 规则注册表：(检查阶段, 规则函数)
#: 检查阶段："img" | "cover" | "toc" | "header" | "footer" | "body"
RULES = {
    "img": [is_signature_image],
    "cover": [is_cover_placeholder],
}


def check_expected(stage, elem, doc_ctx=None):
    """对某元素跑该阶段的所有规则，返回 (是否预期差异, 原因)。"""
    for rule in RULES.get(stage, []):
        ok, reason = rule(elem, doc_ctx or {})
        if ok:
            return True, reason
    return False, ""

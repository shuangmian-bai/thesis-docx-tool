"""块模型 tuple ↔ dict/JSON 双向转换与严格校验。

AI 只能读写 JSON（dict）结构，不能直接碰内部 tuple；本模块是两者之间唯一的
边界，校验失败的 AI 输出一律不得转回块，从结构上保证"AI 胡来就整批丢弃"。

块定义与 docxbuild.mdparse / docxconvert.parse 完全对齐：

    ("h", 级别1-4, 文本)
    ("p"|"quote"|"caption"|"ref"|"ol", 文本)
    ("code", [行文本, ...])
    ("table", [[单元格, ...], ...])
    ("img", 图片路径, 图题文本, cx, cy)  # cx/cy 为 EMU，0 表示按图片像素自适应

其中 `ol` 是有序列表（1. 2. 3.）的**单个条目**：连续多个 ol 块在构建时
归为同一个自动编号列表，拆成逐条目块是为了与块数护栏兼容（AI 只改分类，
不增删块、不并块）。

对应的 dict 形态（喂给 AI / 从 AI 收回）：

    {"t":"h","level":1,"text":"..."}
    {"t":"p","text":"..."}
    {"t":"ol","text":"..."}
    {"t":"code","lines":["..."]}
    {"t":"table","rows":[["..."]]}
    {"t":"img","path":"images/x.png","caption":"...","cx":0,"cy":0}
"""
from typing import Dict, List, Optional, Tuple

Block = Tuple
TEXT_KINDS = ("p", "quote", "caption", "ref", "ol")


class BlockSchemaError(Exception):
    """AI 返回的块不符合 schema。消息含批次内位置与原因，便于界面展示。"""


def block_to_dict(block: Block) -> Dict[str, object]:
    """tuple 块 → dict（发送给 AI / GUI 编辑的中间形态）。"""
    kind = block[0]
    if kind == "h":
        return {"t": "h", "level": int(block[1]), "text": block[2]}
    if kind in TEXT_KINDS:
        return {"t": kind, "text": block[1]}
    if kind == "code":
        return {"t": "code", "lines": list(block[1])}
    if kind == "table":
        return {"t": "table", "rows": [list(r) for r in block[1]]}
    if kind == "img":
        return {"t": "img", "path": block[1], "caption": block[2],
                "cx": int(block[3]), "cy": int(block[4])}
    raise BlockSchemaError(f"未知块类型：{kind!r}")


def blocks_to_dicts(blocks: List[Block]) -> List[Dict[str, object]]:
    """整块序列 → dict 列表。"""
    return [block_to_dict(b) for b in blocks]


def dict_to_block(d: Dict[str, object]) -> Block:
    """dict → tuple 块；不合 schema 抛 BlockSchemaError。"""
    problems = validate_block(d)
    if problems:
        raise BlockSchemaError("；".join(problems))
    t = d["t"]
    if t == "h":
        return ("h", int(d["level"]), d["text"])
    if t in TEXT_KINDS:
        return (t, d["text"])
    if t == "code":
        return ("code", list(d["lines"]))
    if t == "table":
        return ("table", [list(r) for r in d["rows"]])
    return ("img", d["path"], d["caption"], int(d.get("cx", 0)), int(d.get("cy", 0)))


def validate_block(d: object) -> List[str]:
    """校验单个 dict 块，返回问题列表（空列表=合法）。只认结构，不改内容。"""
    if not isinstance(d, dict):
        return ["块不是对象"]
    t = d.get("t")
    if not isinstance(t, str):
        return ["缺少或非法的类型字段 t"]

    if t == "h":
        probs = _check_text(d)
        lv = d.get("level")
        if not isinstance(lv, int) or isinstance(lv, bool) or not (1 <= lv <= 4):
            probs.append("h 的 level 必须是 1-4 的整数")
        return probs
    if t in TEXT_KINDS:
        return _check_text(d)
    if t == "code":
        lines = d.get("lines")
        if not isinstance(lines, list) or not all(isinstance(x, str) for x in lines):
            return ["code 的 lines 必须是字符串列表"]
        return []
    if t == "table":
        rows = d.get("rows")
        if not isinstance(rows, list) or not rows:
            return ["table 的 rows 必须是非空列表"]
        if not all(isinstance(r, list) and all(isinstance(c, str) for c in r)
                   for r in rows):
            return ["table 的每一行必须是字符串列表"]
        return []
    if t == "img":
        probs = []
        if not isinstance(d.get("path"), str) or not d["path"]:
            probs.append("img 的 path 必须是非空字符串")
        if not isinstance(d.get("caption"), str):
            probs.append("img 的 caption 必须是字符串")
        for k in ("cx", "cy"):
            v = d.get(k, 0)
            if not isinstance(v, int) or isinstance(v, bool) or v < 0:
                probs.append(f"img 的 {k} 必须是非负整数")
        return probs
    return [f"未知块类型 {t!r}（允许 h/p/quote/caption/ref/ol/code/table/img）"]


def _check_text(d: Dict[str, object]) -> List[str]:
    return [] if isinstance(d.get("text"), str) else ["text 必须是字符串"]


def parse_ai_payload(raw: str) -> List[Dict[str, object]]:
    """解析 AI 返回文本为 dict 列表。

    兼容三种常见返回：裸 JSON 数组、{"blocks":[...]} 对象、带 ```json 围栏。
    解析失败抛 BlockSchemaError，调用方按"整批丢弃"处理。
    """
    import json
    import re
    text = raw.strip()
    fence = re.match(r"^```(?:json)?\s*(.*?)\s*```$", text, re.S)
    if fence:
        text = fence.group(1).strip()
    try:
        obj = json.loads(text)
    except json.JSONDecodeError as e:
        raise BlockSchemaError(f"AI 返回不是合法 JSON：{e}")
    if isinstance(obj, dict):
        obj = obj.get("blocks")
    if not isinstance(obj, list):
        raise BlockSchemaError("AI 返回顶层必须是块数组，或含 blocks 数组的对象")
    return obj


def dicts_to_blocks(dicts: List[Dict[str, object]], *,
                    expected_count: Optional[int] = None,
                    known_images: Optional[set] = None,
                    ) -> List[Block]:
    """整批 dict → tuple 块，并做两道护栏校验：

    - expected_count：块数必须与输入一致（防止 AI 漏块/并块/自由发挥）；
    - known_images：img 的 path 必须在给定白名单内（AI 不得新增/篡改图片路径）。

    任一不过抛 BlockSchemaError，调用方整批保留原块。
    """
    if expected_count is not None and len(dicts) != expected_count:
        raise BlockSchemaError(
            f"块数不一致：输入 {expected_count} 块，AI 返回 {len(dicts)} 块"
            "（疑似漏块或并块），整批保留原文")
    blocks = []
    for i, d in enumerate(dicts):
        problems = validate_block(d)
        if problems:
            raise BlockSchemaError(f"第 {i + 1} 块不合规：" + "；".join(problems))
        if d["t"] == "img" and known_images is not None \
                and d["path"] not in known_images:
            raise BlockSchemaError(
                f"第 {i + 1} 块的图片路径 {d['path']!r} 不在抽取图片集合内，"
                "整批保留原文")
        blocks.append(dict_to_block(d))
    return blocks


def split_chapters(blocks: List[Block]
                   ) -> List[Dict[str, object]]:
    """按 H1（level==1）把块序列切成章，保证各章拼接可无损还原原序列。

    返回 [{"title": 章标题, "indexes": [全局下标...]}]；
    第一个 H1 之前的块（如摘要）归入标题为"前置内容"的首章。
    """
    chapters: List[Dict[str, object]] = []
    cur_title = "前置内容"
    cur_idx: List[int] = []

    def flush():
        if cur_idx:
            chapters.append({"title": cur_title, "indexes": cur_idx.copy()})

    for i, b in enumerate(blocks):
        if b[0] == "h" and b[1] == 1:
            flush()
            cur_title = b[2]
            cur_idx = []
        cur_idx.append(i)
    flush()
    return chapters

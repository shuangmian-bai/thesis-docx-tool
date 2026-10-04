"""规则解析结果的 AI 编排：结构修正（按章分批）与内容改写。

策略是"规则打底 + AI 修正"：docxconvert 先产出完整块序列，AI 只对结构误判
做校正。每章一批，返回块数/schema/图片路径任一校验不过，**整批保留原文**，
内容零丢失；成功的批次标记变化块，供 GUI 高亮与逐块还原。

本模块不依赖 PyQt6，GUI 与后续终端交互版共用；通过回调上报进度与支持取消。
"""
import json
from dataclasses import dataclass, field
from typing import Callable, List, Optional, Set, Tuple

from docxai import blocks_json as bj
from docxai.client import AIClient, AIError
from docxai.prompts import rewrite_messages, structure_messages

Block = Tuple


@dataclass
class CorrectionResult:
    """一次整篇结构修正的结果。"""
    blocks: List[Block]                      # 修正后的块序列（被拒批次保持原样）
    changed: Set[int] = field(default_factory=set)      # 相对原块发生变化的全局下标
    skipped: List[dict] = field(default_factory=list)   # [{chapter, reason}] 被拒批次
    ok_chapters: int = 0                     # AI 成功处理的批次数
    total_chapters: int = 0                  # 总批次数

    def mark_meta(self, ok: int, total: int) -> None:
        self.ok_chapters, self.total_chapters = ok, total


def correct_structure(client: AIClient, blocks: List[Block], *,
                      known_images: Optional[Set[str]] = None,
                      on_progress: Optional[Callable[[int, int, str], None]] = None,
                      cancel: Optional[Callable[[], bool]] = None
                      ) -> CorrectionResult:
    """按章分批让 AI 修正块类型/层级。

    参数：
        blocks:       规则解析得到的完整块序列（会复制后修改，不改入参）。
        known_images: 图片路径白名单（img.path 必须在内），None 表示不校验。
        on_progress:  回调 (已完成批次数, 总批次数, 章节标题)。
        cancel:       返回 True 则在批次间隙中止（已完成的批次保留）。

    任何一批失败都记录到 skipped 并保留该批原块，不向上抛，保证整篇可用。
    """
    work = list(blocks)
    chapters = bj.split_chapters(work)
    result = CorrectionResult(blocks=work)
    total = len(chapters)
    ok = 0

    for done, ch in enumerate(chapters, 1):
        if cancel and cancel():
            break
        title = ch["title"]
        idxs: List[int] = ch["indexes"]
        if on_progress:
            on_progress(done - 1, total, title)

        old_chapter = [work[i] for i in idxs]
        # 纯 H1 标题章（只有标题本身）没有可校正的正文，跳过调用省 token
        if len(old_chapter) <= 1 and old_chapter[0][0] == "h":
            ok += 1
            if on_progress:
                on_progress(done, total, title)
            continue

        dicts = bj.blocks_to_dicts(old_chapter)
        try:
            raw = client.chat(structure_messages(title, dicts),
                              temperature=0.0, json_mode=True, cancel=cancel)
            new_chapter = bj.dicts_to_blocks(
                bj.parse_ai_payload(raw),
                expected_count=len(old_chapter),
                known_images=known_images)
        except (AIError, bj.BlockSchemaError) as e:
            result.skipped.append({"chapter": title, "reason": str(e)})
        else:
            ok += 1
            for gi, (old, new) in enumerate(zip(old_chapter, new_chapter)):
                if old != new:
                    work[idxs[gi]] = new
                    result.changed.add(idxs[gi])

        if on_progress:
            on_progress(done, total, title)

    result.mark_meta(ok, total)
    return result


def rewrite_text(client: AIClient, text: str, instruction: str, *,
                 cancel: Optional[Callable[[], bool]] = None) -> str:
    """让 AI 按指令改写单段文本，返回改后纯文本。

    AI 须返回 {"text": "..."}；解析失败抛 AIError，调用方保留原文本。
    """
    raw = client.chat(rewrite_messages(text, instruction),
                      temperature=0.4, json_mode=True, cancel=cancel)
    try:
        obj = json.loads(_strip_fence(raw))
        out = obj["text"]
    except (json.JSONDecodeError, KeyError, TypeError):
        raise AIError("改写结果无法解析（缺少 text 字段），已保留原文")
    if not isinstance(out, str) or not out.strip():
        raise AIError("改写结果为空，已保留原文")
    return out.strip()


def _strip_fence(text: str) -> str:
    """去掉可能的 ```json 围栏（双保险，json_mode 下一般没有）。"""
    import re
    m = re.match(r"^```(?:json)?\s*(.*?)\s*```$", text.strip(), re.S)
    return m.group(1).strip() if m else text.strip()

"""AI 提示词集中管理。

两类任务：
- 结构修正：规则解析可能把标题判成正文、图题表题错位等，AI 只修"分类与层级"，
  不得改写实质文字、不得增删块、不得动表格内容与图片路径；
- 内容改写：用户对选中文本下指令（润色、缩写、改语气），只返回改后文本。
"""
import json
from typing import Dict, List

SYSTEM_STRUCTURE = (
    "你是毕业论文 Word 文档结构解析的校对器。输入是一个有序的文档块数组，"
    "每块形如 {\"t\":类型,...}。规则解析器可能误判块类型，你只做结构修正：\n"
    "1. 把被误判为普通正文(p)的标题改为 h，level 取 1-4；把误判为标题的正文改回 p；\n"
    "2. 识别图题/表题：caption 类型（正文中以\"图\"或\"表\"加编号开头的说明句）；\n"
    "3. 参考文献条目用 ref；代码段用 code（lines 为行数组）；表格用 table（rows 为二维字符串数组）；\n"
    "4. 属于同一组 1. 2. 3. 有序编号并列内容的每个段落，改为 ol 类型（text 为该条文字，"
    "去掉手写的\"1.\"\"2)\"等编号前缀）；列表中的每条仍是独立块，严禁把多条合并成一块；"
    "不是并列编号内容的普通段落即使以数字开头也保持 p；\n"
    "5. img 块、table 的单元格文字一律原样保留，禁止修改 path/caption 绑定与表格内容；\n"
    "严禁增删任何块、改变块的数量与顺序，严禁润色或改写正文的实质内容，"
    "除块类型与标题 level 外其余字段逐字照抄。"
    "只输出修正后的块数组 JSON，不要输出解释或 markdown 代码围栏。"
)

SYSTEM_REWRITE = (
    "你是中文论文写作助手。按用户指令修改给定文本，只返回 JSON "
    "{\"text\": \"修改后的文本\"}。保持学术语气，不改变事实与数据，"
    "不要输出解释、标题符号或代码围栏。"
)


def structure_messages(chapter_title: str,
                       dicts: List[Dict[str, object]]
                       ) -> List[Dict[str, str]]:
    """构造一章的结构修正对话。"""
    user = (f"以下是论文章节「{chapter_title}」的 {len(dicts)} 个文档块，"
            "按系统规则修正块类型后原样输出同数量的块数组 JSON：\n"
            + json.dumps(dicts, ensure_ascii=False))
    return [
        {"role": "system", "content": SYSTEM_STRUCTURE},
        {"role": "user", "content": user},
    ]


def rewrite_messages(text: str, instruction: str) -> List[Dict[str, str]]:
    """构造单段文本的内容改写对话。"""
    user = (f"修改要求：{instruction}\n"
            f"待修改文本：\n{text}")
    return [
        {"role": "system", "content": SYSTEM_REWRITE},
        {"role": "user", "content": user},
    ]

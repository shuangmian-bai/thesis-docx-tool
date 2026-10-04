"""论文著录信息：题目、封面字段与版本号。

这些是**论文特定的内容**，与构建逻辑无关——换论文、补填封面只改本模块或同目录下
的 `cover.json`。`template.fill_cover()` 用 `COVER` 填封面表格、用 `TEMPLATE_TITLE`
定位诚信承诺书里的示例题目；`cli.main()` 用 `TITLE` 与 `COVER["学生姓名"]` 写
`docProps/core.xml`。

开源版本**不含任何个人数据**：所有字段默认为占位符，真实信息由使用者在
`config/cover.json` 里填写（`.gitignore` 排除，不入库）。
"""
import json
import os

#: 版本号，进产物文件名（`论文_v1_YYYYMMDD.docx`）
VERSION = "v1"

TITLE = "【待填：论文题目】"

# ── 封面与承诺书中需要替换的字段 ───────────────────────────────
# 开源默认值全为占位符；真实信息写在 cover.json 里（不入库）。
COVER = {
    "成果名称": TITLE,
    "学生姓名": "【待填：姓名】",
    "学    号": "【待填：学号】",
    "专    业": "【待填：专业】",
    "二级学院": "【待填：二级学院】",
    "班    级": "【待填：班级】",
    "指导老师": "【待填：指导老师】",
}

#: 模板承诺书里出现的示例题目（用于定位替换）
TEMPLATE_TITLE = "【待填：模板承诺书示例题目】"

#: config/cover.json，填写后覆盖上面的占位符
_COVER_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "config", "cover.json")


def _load_cover_json():
    """若存在 cover.json，用其中的字段覆盖占位符。"""
    global TITLE, TEMPLATE_TITLE, VERSION
    if not os.path.exists(_COVER_JSON):
        return
    try:
        with open(_COVER_JSON, encoding="utf-8") as fh:
            cfg = json.load(fh)
    except (OSError, json.JSONDecodeError) as e:
        raise SystemExit(f"cover.json 解析失败：{e}")
    if "title" in cfg:
        TITLE = cfg["title"]
        COVER["成果名称"] = TITLE
    if "template_title" in cfg:
        TEMPLATE_TITLE = cfg["template_title"]
    if "version" in cfg:
        VERSION = str(cfg["version"])
    if "cover" in cfg and isinstance(cfg["cover"], dict):
        for k, v in cfg["cover"].items():
            if k in COVER:
                COVER[k] = v


_load_cover_json()

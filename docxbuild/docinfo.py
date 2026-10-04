"""论文著录信息：题目、封面字段与版本号。

这些是**论文特定的内容**，与构建逻辑无关。`template.fill_cover()` 用 `COVER`
填封面表格、用 `TEMPLATE_TITLE` 定位诚信承诺书里的示例题目；`cli.build()`
用 `TITLE` 与 `COVER["学生姓名"]` 写 `docProps/core.xml`。

**封面字段来源（不再用全局 config/cover.json）**：
- Word 路线：`convert` 时从 docx 封面表格提取字段，存入哈希工作目录的
  `cover.json`；`build` 时通过 `cover_path` 参数加载，每个论文独立、互不污染。
- GUI 路线：模板预览页修改封面后存到当前论文的工作目录。
- MD 路线（不传 docx）：使用本模块的占位符默认值。

开源版本**不含任何个人数据**：所有字段默认为占位符。
"""
import json
import os

#: 版本号，进产物文件名（`论文_v1_YYYYMMDD.docx`）
VERSION = "v1"

TITLE = "【待填：论文题目】"

# ── 封面与承诺书中需要替换的字段 ───────────────────────────────
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


def reset_cover():
    """把封面字段重置为占位符默认值。"""
    global TITLE, TEMPLATE_TITLE, VERSION
    TITLE = "【待填：论文题目】"
    TEMPLATE_TITLE = "【待填：模板承诺书示例题目】"
    VERSION = "v1"
    COVER.update({
        "成果名称": TITLE,
        "学生姓名": "【待填：姓名】",
        "学    号": "【待填：学号】",
        "专    业": "【待填：专业】",
        "二级学院": "【待填：二级学院】",
        "班    级": "【待填：班级】",
        "指导老师": "【待填：指导老师】",
    })


def load_cover_from(path: str):
    """从指定路径的 cover.json 加载封面字段，覆盖当前值。

    文件不存在时静默跳过（保持占位符）；解析失败则报错。
    """
    global TITLE, TEMPLATE_TITLE, VERSION
    if not os.path.exists(path):
        return
    try:
        with open(path, encoding="utf-8") as fh:
            cfg = json.load(fh)
    except (OSError, json.JSONDecodeError) as e:
        raise SystemExit(f"封面配置解析失败（{path}）：{e}")
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


def save_cover_to(path: str, form: dict):
    """把封面表单落为指定路径的 cover.json。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = {
        "title": form.get("title", "").strip(),
        "version": form.get("version", "v1").strip() or "v1",
        "template_title": form.get("template_title", "").strip(),
        "cover": {k: form["cover"].get(k, "").strip() for k in COVER},
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

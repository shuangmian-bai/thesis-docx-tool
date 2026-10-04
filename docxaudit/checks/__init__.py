"""五项检查，一项一个文件。

`CHECKS` 是唯一的清单：命令行参数 `audit_docx.py <名>`、`--list` 输出、总览的表头
标签都从它派生，新增一项只加一行、加一个模块。
"""
from docxaudit.checks.body import check_body
from docxaudit.checks.meta import check_meta
from docxaudit.checks.parts import check_parts
from docxaudit.checks.stale import check_stale
from docxaudit.checks.styles import check_styles

#: (命令名, 总览标签, 检查函数, 一句话说明)
CHECKS = [
    ("stale", "stale ", check_stale, "成品是否比章节源还旧（前置警示）"),
    ("parts", "parts ", check_parts, "zip 部件：清单、逐字节、Content_Types、关系、媒体、页眉页脚"),
    ("styles", "styles", check_styles, "样式：逐样式、docDefaults、TOC 大小写、使用统计"),
    ("body", "body  ", check_body, "正文：保留区、分块、分节符、页面设置、表格分页、直接格式化"),
    ("meta", "meta  ", check_meta, "属性与设置：core.xml、settings.xml、编号、命名空间"),
]

CHECK_MAP = {name: (label, fn, desc) for name, label, fn, desc in CHECKS}

# docxaudit — 成品 docx 格式审计

| 项目 | 说明 |
| --- | --- |
| **分层定位** | 质量层 · 审计（成品 vs 模板） |
| **开发状态** | [OK] 已完成 |
| **职责** | 逐项检查生成的 docx 是否贴合模板格式，避免排版跑偏 |

## 用途

`audit_docx.py` 的实现包。把成品 docx 与学校模板逐部件、逐样式、逐段落地比对，
差异分四类：`[通过]`、`[有意偏离]`（登记在 `EXPECTED_DIFFS`）、`[注意]`、`[失败]`。

退出码：有「注意」或「失败」项为 1；`--strict` 时「有意偏离」也算失败。

## 模块组成

| 文件 | 说明 |
| --- | --- |
| `common.py` | 公共层：`Finding` / `Docx` / `Ctx`、模板与成品路径定位、XML 骨架比对工具 |
| `expected.py` | `EXPECTED_DIFFS` 表——登记「有意偏离」，新增一处只改这里 |
| `checks/` | 五个检查项，一项一个文件；顺序在 `checks/__init__.py` 里排定 |
| `report.py` | 结论分级（`classify`）、渲染（`render`）、单项执行（`run_check`） |
| `cli.py` | 命令行解析与总览输出（`main(argv)` → 退出码） |

## 检查项

| 命令名 | 说明 |
|---|---|
| `stale` | 成品是否比章节源还旧（前置警示） |
| `parts` | zip 部件：清单、逐字节、Content_Types、关系、媒体、页眉页脚 |
| `styles` | 样式：逐样式、docDefaults、TOC 大小写、使用统计 |
| `body` | 正文：保留区、分块、分节符、页面设置、表格分页、直接格式化 |
| `meta` | 属性与设置：core.xml、settings.xml、编号、命名空间 |

## 审计思路

差异**只在报告层降级**：检查函数一律报 `WARN`，命中 `EXPECTED_DIFFS` 才在
`report.classify()` 里改判为「有意偏离」。因此新增一处有意偏离时只改表，
不必回头动检查逻辑。

## 关键 API 与用法

```python
from docxaudit.cli import main
main(["--list"])               # 列出全部检查项
main(["parts", "styles"])      # 只查指定项
main([])                       # 总览（全部五项）
```

默认成品取本目录下 mtime 最新的 `*.docx`（排除 `template.docx`）。

## 依赖与被依赖

- **依赖**：Python 标准库（`zipfile` / `xml.etree` / `re` / `os` / `difflib`）。
- **被谁用**：`main.py` 的 `audit` 子命令。

## 注意点

- `HERE` 指向工具根目录（`thesis-docx-tool/`），本包在其下一层，故上跳一级。
- 模板与成品路径由 `common.py` 的 `find_template` / `find_product` 自动定位，
  也可用 `--template` / `--product` 显式指定。
- 新增检查项：在 `checks/` 加一个文件，在 `checks/__init__.py` 的 `CHECKS` 里加一行，
  若是有意偏离再在 `expected.py` 登记。

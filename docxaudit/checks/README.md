# checks — 审计检查项集合

| 项目 | 说明 |
| --- | --- |
| **分层定位** | 质量层 · 审计子模块 |
| **开发状态** | [通过] |
| **职责** | 五项格式审计检查，一项一个文件 |

## 用途

`docxaudit` 的检查项集合。每个文件是一个检查函数，输入审计上下文 `Ctx`，
输出 `Finding` 列表。检查项的清单、顺序、命令行参数名都由本目录的 `__init__.py`
统一管理，新增一项只加一个文件 + 一行登记。

## 检查项清单

| 文件 | 命令名 | 说明 |
| --- | --- | --- |
| `stale.py` | `stale` | 成品是否比章节源还旧（前置警示，旧稿结论不作数） |
| `parts.py` | `parts` | zip 部件：清单、逐字节、Content_Types、关系、媒体、页眉页脚 |
| `styles.py` | `styles` | 样式：逐样式、docDefaults、TOC 大小写、使用统计 |
| `body.py` | `body` | 正文：保留区、分块、分节符、页面设置、表格分页、直接格式化 |
| `meta.py` | `meta` | 属性与设置：core.xml、settings.xml、编号、命名空间 |

## 新增检查项

1. 新建一个文件，如 `checks/foo.py`，定义 `check_foo(ctx) -> list[Finding]`。
2. 在 `checks/__init__.py` 的 `CHECKS` 里加一行：
   ```python
   ("foo", "foo  ", check_foo, "一句话说明"),
   ```
3. 若是「有意偏离」（已知且可接受的差异），在 `docxaudit/expected.py` 的
   `EXPECTED_DIFFS` 表里登记，使其从「注意」降级为「有意偏离」。

## 依赖与被依赖

- **依赖**：`docxaudit.common`（`Finding`、`Ctx`、XML 工具）、`docxaudit.expected`
  （偏离登记表）。
- **被谁用**：`docxaudit.report.run_check` 按 `CHECKS` 顺序逐项执行。

## 注意点

- 检查函数一律报 `WARN`，降级为「有意偏离」由 `report.classify()` 统一处理，
  不要在检查函数里直接判「有意偏离」。
- `Finding.key` 必须与 `EXPECTED_DIFFS` 的键一一对应，改 key 时务必同步两处。

# docxloop —— 闭环测试引擎

对输入 docx 执行「流水线出稿 → 全盘对比」，验证拆解（convert）与构建（build）
的无损性，产出可视化报告供人工验收。

## 用法

```bash
python3 main.py loop 论文.docx        # 单文件
python3 main.py loop ./papers/        # 文件夹批量
python3 main.py loop <16位哈希>       # 复用已有工作目录（与 run 共用）
python3 main.py loop 论文.docx --no-reuse   # 强制重跑流水线（改了 convert/build 后）
python3 main.py loop 论文.docx --template 模板.docx   # 指定模板
```

模板默认是**每个输入文件自己**：闭环验证的是「拆解→重建」的无损性，
模板与输入同源才能排除格式差异干扰；需要验证对某个学校模板的贴合度时再传
`--template`。

## 架构约定（硬性）

本模块**只是入口/执行层，不含任何流程逻辑**：

- 出稿一律调用应用层 `docxflow.cli.run_pipeline()`，与命令行 `run`、GUI
  构建是同一条流水线，产物必然一致；不得在本模块重写 convert/build/
  封面提取/章节写入的任何环节。
- 哈希隔离与索引复用 `docxflow.workdir`（与 run/gui 共用
  `.cache/run_work/`），同一篇论文在三条入口下是同一个工作目录。
- 本模块自己的职责只有三件：输入解析（文件/文件夹/哈希）、触发流水线、
  全盘对比出报告。
- 闭环产物保留在工作目录 `output/` 下，不回拷全局 `output/`；
  终端摘要与 HTML 报告都必须给出原文与成品的完整路径，供人工打开两份
  Word 逐项核验（子项目 CLAUDE.md 项目特定规则第 9 条）。

## 模块组成

| 文件 | 职责 |
|---|---|
| `cli.py` | 参数解析、结果摘要、报告落盘 |
| `engine.py` | 输入解析（`resolve_input`）、单文件/批量闭环编排（`run_closed_loop` / `run_batch`） |
| `compare.py` | 全盘对比：正文/封面/目录/承诺书/页眉/页脚/样式，产出 Diff 列表；预期差异规则（目录页码与制表位）内联在本文件 |
| `report.py` | 自包含 HTML 报告（差异表格 + 原文/成品路径） |

报告默认写 `.cache/loop_report.html`。

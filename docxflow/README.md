# docxflow —— 一键流水线（run）

面向用户的**唯一主入口**。把拆解 / 架构图 / 构建 / 页码回填 / 审计等内部阶段
串成端到端流程，用户一条命令得到成品 Word，不需要手工分步运行。

## 两条输入路径（同一入口，参数区分）

```bash
# Word 路线：用户 docx → 自动拆解入哈希工作目录 → 成品 Word
python3 main.py run 用户论文.docx

# 哈希路线：16 位哈希 → 复用已有工作目录重跑（跳过拆解）
python3 main.py run <哈希>
```

两条路径统一走「哈希工作目录」（`.cache/run_work/<哈希>/`）：Word 路线先按内容
哈希创建工作目录，拆解产物落 `src.md`（整篇）+ `章节/`（按 H1 拆分的数字开头 md）
+ `images/`（抽取图片），再构建；哈希路线直接复用已有工作目录的 `章节/` 构建。
默认剥离原 docx 的封面/承诺书/目录（由模板骨架提供），成品复制到全局 `output/`。

## 模块组成

| 文件 | 职责 |
|---|---|
| `cli.py` | `run_pipeline()` 全流程**唯一编排实现** + `main()` 参数解析；章节落盘唯一实现 `write_chapters_from_blocks()` |
| `workdir.py` | 哈希工作目录唯一实现：`file_hash` / `prepare_work_dir` / 索引 `index.json` 读写 / 哈希定位 |

命令行 `run`、GUI 构建、闭环测试（docxloop）三个入口都只解析参数后调用
`run_pipeline()`，不得各自实现流程逻辑；工作目录管理一律经 `workdir.py`。

## 流水线阶段

| 阶段 | 动作 | 跳过条件 |
|---|---|---|
| 开工前检查 | 模板、章节源（哈希路线）一次性报齐 | 不可跳过 |
| 拆解（仅 Word 路线） | `prepare_blocks()` → `src.md` + `write_chapters_from_blocks()` 落 `章节/` | 哈希路线无此阶段 |
| 架构图 | 调 `docxfig.generate()` 渲染 `config/figures/*.md` | 无图定义 |
| 构建（第一遍） | 调 `docxbuild.cli.build()` 生成成品 docx | 不可跳过 |
| 目录页码回填 | soffice 转 PDF → `toc_pages` 量页码（写工作目录 `toc.json`）→ 第二遍构建覆盖 | 缺 soffice/pdftotext 自动跳过 |
| 成品审计 | 调 `docxaudit` 对最终成品出报告；结论不影响流水线成败 | 闭环测试内部传 `no_audit=True` |

外部工具（`soffice`、`pdftotext`）缺失时页码回填**自动降级**：照常出稿，
目录页码留空，在 Word 中按 Ctrl+A 后 F9 即可由排版引擎更新；安装工具后重跑 `run`
可自动回填。

## 参数

```
python3 main.py run <论文.docx | 16位哈希> [--template 模板.docx]
```

面向用户只有「word/哈希 + --template」两个参数，无任何 `--no-*` 开关。
`run_pipeline(source, template=None, *, no_audit=False, copy_to_output=True)` 的
`no_audit` / `copy_to_output` 是**内部关键字参数**，仅供闭环测试（docxloop）
程序内调用，命令行不暴露。

## 设计约定

- **三层架构中的应用层**：
  - **功能模块层**（`docxconvert` / `docxbuild` / `docxfig` 等）只做一件事，
    不感知清理旧产物、多入口一致性等流程问题。
  - **应用层**（本包）负责把功能模块串成完整流程，并承担「构建前清理旧产物」
    「章节源准备」「哈希工作目录管理」等跨模块流程逻辑，每样只在此实现一处。
  - **入口/执行层**（`main.py` 命令行、`docxgui` GUI、`docxloop` 闭环测试）
    只调用本包公开接口（`run_pipeline()` / `write_chapters_from_blocks()` /
    `workdir.py`），不自行实现任何流程逻辑——闭环测试连 convert/build 的
    调用顺序都不得重写。
- **只编排，不重写**：本包不实现任何解析/构建逻辑，只调用各原子模块的公开接口，
  同一职责全项目只有一份实现。
- **哈希隔离（全局硬性）**：所有路径一律用内容哈希（SHA256 前 16 位）
  建/定位工作目录；索引 `index.json`（哈希→原文件路径）与目录并列存放；
  页码表、章节源、图片全部落工作目录，不碰全局可变目录。
- **产物命名沿用 build**：成品为 `论文_<版本号>_<当天日期>.docx`；
  第二遍构建路径相同，直接覆盖第一遍；成品同时复制一份到全局 `output/` 方便取用。
- **新增任何需要多步完成的能力，必须接进本流水线**，不得要求用户手工串联，
  也不新增面向用户的分步子命令（见子项目 CLAUDE.md 项目特定规则）。

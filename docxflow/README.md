# docxflow —— 一键流水线（run）

面向用户的**唯一主入口**。把 fig / convert / build / toc / audit 五个分步原子命令
在内部串成端到端流程，用户一条命令得到成品 Word，不需要手工分步运行。

## 两条路径

```bash
# MD 路线：章节 Markdown（含文字、架构图引用、自定义插图）→ 成品 Word
python3 main.py run

# Word 路线：用户已写好的 docx → 自动拆解入 config/章节/ → 成品 Word
python3 main.py run 用户论文.docx
```

Word 路线默认剥离原 docx 的封面/承诺书/目录（这些由模板提供），需要保留加 `--keep-front`。
拆解产物落 `config/章节/<docx名>.md`，图片落 `config/images/<docx名>_images/`；
同名 Markdown 已存在时拒绝覆盖，确认覆盖加 `--force`。

## 流水线阶段

| 阶段 | 动作 | 跳过条件 |
|---|---|---|
| 开工前检查 | 模板、签名图、章节源、覆盖冲突一次性报齐 | 不可跳过 |
| 拆解（仅 Word 路线） | 调 `docxconvert.convert_to()` 落到 `config/` | 不传 docx 时无此阶段 |
| 架构图 | 调 `docxfig.generate()` 渲染 `config/figures/*.md` | 无图定义，或 `--no-fig` |
| 构建（第一遍） | 调 `docxbuild.cli.main()` 生成成品 docx | 不可跳过 |
| 目录页码回填 | soffice 转 PDF → `toc_pages` 量页码 → 第二遍构建覆盖 | 缺 soffice/pdftotext 自动跳过，或 `--no-toc` |
| 成品审计 | 调 `docxaudit` 对最终成品出报告；结论不影响流水线成败 | `--no-audit` |

外部工具（`soffice`、`pdftotext`）缺失时页码回填**自动降级**：照常出稿，
目录页码留空，在 Word 中按 Ctrl+A 后 F9 即可由排版引擎更新；安装工具后重跑 `run`
可自动回填。

## 参数

```
python3 main.py run [用户论文.docx]
    [--template 模板.docx]   # 外部模板，默认 config/template.docx；构建与审计都跟随
    [--force]                # 允许覆盖已有的同名拆解 Markdown
    [--keep-front]           # Word 路线保留原 docx 前置页
    [--no-fig] [--no-toc] [--no-audit]   # 关闭对应阶段
```

## 设计约定

- **三层架构中的应用层**：
  - **功能模块层**（`docxconvert` / `docxbuild` / `docxfig` 等）只做一件事，
    不感知清理旧产物、多入口一致性等流程问题。
  - **应用层**（本包）负责把功能模块串成完整流程，并承担「构建前清理旧产物」
    「章节源准备」等跨模块流程逻辑。清理逻辑只在此实现一处，出 bug 只修一处。
  - **入口/执行层**（`main.py`、`docxgui`）只调用本包公开接口，不自行实现
    清理、拆分章节、写入文件等流程逻辑。
- **只编排，不重写**：本包不实现任何解析/构建逻辑，只调用各原子模块的公开接口
  （`convert_to()`、`generate()`、各模块 `cli.main()`），同一职责全项目只有一份实现。
- **旧产物不得残留**：Word 路线与 GUI 路线在写入章节源前，由本包统一清理
  `config/章节/` 中已有的数字开头 `.md`，避免新旧章节混杂；`output/` 旧 docx
  由同名覆盖处理。
- **产物命名沿用 build**：成品为 `output/论文_<版本号>_<当天日期>.docx`；
  第二遍构建路径相同，直接覆盖第一遍。
- 分步子命令保留是为了单步调试与高级用途；**新增任何需要多步完成的能力，
  必须接进本流水线**，不得要求用户手工串联（见子项目 CLAUDE.md 项目特定规则）。

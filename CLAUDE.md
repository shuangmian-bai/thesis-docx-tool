# CLAUDE.md

本文件指导在 `thesis-docx-tool/` 子项目内工作。这是一个**开源项目**，
目标上传至 GitHub 公开仓库，**绝对不能包含任何个人数据**。

## 定位

`thesis-docx-tool` 是一套论文 Word 排版与拆解工具，做五件事：

1. **build**：把按章拆分的 Markdown（`章节/*.md`）套用学校论文模板骨架，生成符合格式的 Word 论文；
2. **convert**：把用户已写好的 Word 论文反向拆解为同样格式的 Markdown，便于二次编辑；
3. **audit**：审计生成的 docx 是否贴合模板格式，避免排版跑偏；
4. **toc**：从渲染出的 PDF 反查目录页码，回填 TOC 域；
5. **fig**：把 `figures/*.md` 中用表格 DSL 描述的架构图渲染为 PNG，供 build 嵌入。

## 目录结构

```
thesis-docx-tool/
├── main.py                 # 统一入口（build / convert / audit / toc / fig 子命令）
├── config/                 # 素材与配置模块（个人数据在 .gitignore 排除）
│   ├── cover.example.json  #   封面信息示例（复制为 cover.json 后填写）
│   ├── cover.json          #   真实封面信息（不入库）
│   ├── template.docx       #   学校论文模板（使用者自备，不入库）
│   ├── 章节/               #   论文正文 Markdown 源（不入库）
│   ├── figures/            #   架构图 DSL 定义（*.md，不入库，见 docxfig/README.md）
│   └── images/             #   论文插图与签名图（不入库；fig 子命令生成的 PNG 也在此）
├── output/                 # 生成产物模块（.gitignore 排除，不入库）
│   ├── 论文_*.docx         #   build 输出
│   ├── 论文正文.md         #   合并后的正文
│   └── toc_pages.json      #   目录页码表
├── docxbuild/              # Markdown → docx（含目录页码回填），见 docxbuild/README.md
├── docxconvert/            # docx → Markdown（逆向拆解），见 docxconvert/README.md
├── docxaudit/              # docx 格式审计（对比模板），见 docxaudit/README.md
│   └── checks/             # 五项审计检查，见 docxaudit/checks/README.md
├── docxfig/                # 架构图生成（Markdown 表格 DSL → PNG），见 docxfig/README.md
├── README.md               # 用法、流程、模块与块模型
└── .gitignore              # 排除所有个人数据与生成产物
```

根目录只放入口（`main.py`）、文档（`README.md`、`CLAUDE.md`）与 `.gitignore`；
素材与配置归 `config/`，生成产物归 `output/`（继承全局第 7 条）。

## 模块文档索引

每个模块目录都附带 `README.md`，改动对应模块前先读：

| 位置 | 内容 |
|---|---|
| `docxbuild/README.md` | 构建链：模块组成、块模型、依赖方向、注意点 |
| `docxconvert/README.md` | 拆解链：解析要点、模块组成、与 build 的闭环关系 |
| `docxaudit/README.md` | 审计：检查项、审计思路、EXPECTED_DIFFS 机制 |
| `docxaudit/checks/README.md` | 五项检查清单、新增检查项步骤 |
| `docxfig/README.md` | 架构图：表格 DSL 语法、渲染引擎、内容/样式分离约定 |

## 个人数据排除（开源硬规则）

本项目上传 GitHub 公开仓库，**以下内容一律不得入库**（已由 `.gitignore` 排除）：

| 类别 | 内容 | 排除方式 |
|---|---|---|
| 封面信息 | 姓名、学号、班级、专业、指导老师 | `config/cover.json` 不入库，只留 `config/cover.example.json` 占位 |
| 学校模板 | `template.docx`（可能含版权） | `config/template.docx` 使用者自备，不入库 |
| 论文正文 | 章节 Markdown 源 | `config/章节/` 不入库 |
| 架构图定义 | 架构图 DSL（`*.md`） | `config/figures/` 不入库 |
| 论文插图 | 插图与签名图（含 fig 生成的 PNG） | `config/images/` 不入库 |
| 生成产物 | docx / pdf / 正文 md / 页码表 | `output/` 整个目录不入库 |

提交前务必用 `git status` 确认暂存区不含上述文件。**开源仓库里出现个人数据视为泄露事故。**

## 继承全局规则（必读）

本子项目继承仓库根 `CLAUDE.md` 的**全部**全局规则：

- **通用约定**（第 1–10 条）：禁止 git、文件边界、虚拟环境、低耦合、文档同步、
  优先读文档、文件规模、文档仅供参考、混合项目分离、禁止表情符号。
- **代码规范**：中文注释、优先标准库、类型标注、命名、模块边界、异常处理等。

全局规则的最新版本以仓库根 `CLAUDE.md` 为准，本文件不重复抄写；
若本节与根文件冲突，以根文件为准。

## 项目特定规则（必读）

以下规则仅适用于 `thesis-docx-tool/`：

1. **标准库-only（build/convert/audit/toc）**：构建与拆解仅依赖 Python 标准库（`zipfile`、`xml.etree`、`struct` 等），
   **不引入** `python-docx`、`pandoc` 等第三方库；`build` 插图尺寸可选 `Pillow`，无 Pillow 时回退到默认尺寸。
   `fig` 子命令需要 `Pillow`（PIL）渲染 PNG。
2. **根目录整洁**：根目录只放入口（`main.py`）、文档（`README.md`、`CLAUDE.md`）与
   `.gitignore`；素材与配置归 `config/`，生成产物归 `output/`（继承全局第 7 条）。
3. **子包必须有 README.md**：每个模块目录必须有 `README.md`
   （对全局第 6 条「优先读文档」的具体落实）。
4. **个人数据走配置**：真实封面信息放 `config/cover.json`（已在 `.gitignore` 排除），
   代码中不得硬编码姓名、学号等。

## 与父仓库的关系

本项目是父仓库（`jumo_hook/`）下的一个**开源子模块**，同时存在于两个语境：

- **父仓库（私有）**：本目录连同个人数据（`config/cover.json`、`config/章节/`、`config/images/`、
  `config/template.docx`）一起被父仓库跟踪，用于开发调试与毕设产出。父仓库 `.gitignore`
  只排除本目录的 `.gitignore` 文件本身，不重复排除个人数据。
- **独立开源仓库（公开）**：本目录单独发布到 GitHub 公开仓库，此时本目录的 `.gitignore`
  生效，排除所有个人数据与生成产物，只含通用工具代码与示例配置。

编码规范、文档约定、通用规则与父仓库**统一同步**，不因私有/开源语境而差异。
开源仓库里出现个人数据视为泄露事故。

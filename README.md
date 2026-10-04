# thesis-docx-tool

论文 Word 排版与拆解工具。日常使用只需一个子命令 **run**，一条命令端到端出成品：

- **run**：一键流水线。自动串联 拆解（可选）→ 架构图 → 构建 → 转 PDF 量目录页码 → 二次构建 → 格式审计，
  支持两条路径：直接写 Markdown，或传入已写好的 Word 自动拆解后重排。

另有 **gui** 图形界面（run 的交互前端）：选 Word 与处理模式（默认/AI），
在界面里审阅扁平化结果——标题↔正文等误判直接改类型、插入表格与图片、
修正 1./2./3. 有序列表，手动或让 AI 修正后一键出稿；还内置模板预览与修复页，
适合解析别人样式混乱的 Word 与核对学校模板：

- **gui**：PyQt6 三页向导（选择 → 审阅修正 → 构建），AI 能力由 `docxai/` 提供，
  模板检视由 `docxbuild/tplinspect.py` 提供

其余五个是分步原子命令，供单步调试与高级用途（run 在内部编排它们）：

- **build**：把按章拆分的 Markdown（`章节/*.md`）套用学校论文模板骨架，生成符合格式的 Word 论文
- **convert**：把已写好的 Word 论文反向拆解为同格式 Markdown，便于二次编辑
- **audit**：审计生成的 docx 是否贴合模板格式，避免排版跑偏
- **toc**：从渲染出的 PDF 反查目录页码，回填 TOC 域
- **fig**：把 `figures/*.md` 中用表格 DSL 描述的架构图渲染为 PNG，供 build 嵌入

命令行核心仅依赖 Python 标准库（`zipfile` / `xml.etree` / `re` / `json` /
`subprocess` / `urllib`），**不用 python-docx、不用 pandoc、不用任何 AI SDK**。
`build` 的插图尺寸计算可选 `Pillow`，无图时可缺省；`fig` 子命令需要 `Pillow`；
`gui` 子命令需要 PyQt6（见下）。

## 设计理念：格式确定，内容自由

论文写作包含两类性质完全不同的问题：

- **内容**（写什么）：人和 AI 擅长，表达自由、每篇不同；
- **格式**（排成什么样）：学校模板固定且严格，不该靠人或 AI 临场发挥。

AI 能写出好内容，却无法稳定对齐每个学校的字体、行距、标题级别、图表题注与目录页码。
本工具因此把格式从写作过程中彻底剥离，让它成为一条**确定的链路**，只需三个动作：

1. **模板扁平化（定骨架）**——学校模板中固定的部分：封面、诚信承诺书、目录域、
   页眉页脚、样式表，原样保留为骨架；把“正文长什么样”抽象成一套与模板无关的
   纯文本块模型（标题 / 正文 / 有序列表 / 图题表题 / 文献 / 代码 / 表格 / 图片）。
   模板不是只读黑盒：`gui` 的「模板预览与修复」页可逐项检视保留骨架、封面字段、
   承诺书占位符、TOC 域、样式、编号定义与页面设置，发现问题可在页内修复封面字段/
   签名图，或在 Word 中修正模板文件后一键重新检测复验。
2. **论文扁平化（转内容）**——把任意来源的 Word（即使样式混乱、标题没套样式）
   解析成同一套块序列。默认按 Word 样式规则解析；样式不规范时由 AI **只修正结构
   分类**（含把并列段落改判为 1./2./3. 有序列表），不整篇重解析、不改实质内容。
   结果在 `gui` 中按章展开为纯文本，可逐块核对、改类型/级别/文字、插入或删除块
   （段落/标题/有序列表项/表格/图片，图片选文件后自动复制进素材目录）、让 AI 再
   修正；结构操作可逐步撤销，文字改动可逐块还原，**确认扁平化准确后才继续**。
3. **确定性复原（出成品）**——纯文本块序列 + 模板骨架 → 成品 Word。这一步是纯规则
   的确定性过程，**不经过 AI**：同一份块序列配同一个模板，必然得到同一份格式；
   随后自动转 PDF 量页码、二次构建回填目录、格式审计，全程可重复、可审计。

**关键前提：扁平化结果必须可修复——模板和论文都不例外。**
任何自动扁平化都不保证一次 100% 准确，所以本工具不把它做成“丢进去就出稿”的黑盒：
无论模板侧还是论文侧，扁平化产物都必须是**看得见、改得动、改完能重新验证**的中间结果，
确认准确后才允许进入确定性复原。论文侧由 `gui` 审阅页承担修复（改类型、插入/删除块、
逐块编辑 + AI 结构修正 + 撤销/还原）；模板侧由 `gui` 模板预览与修复页承担检视
（骨架纯文本序列、封面字段、样式、编号、页面设置），封面字段与签名图可在页内直接
修复，骨架与样式本体在 Word 中改好模板文件后回页内重新检测，并由 `audit` 的格式
差异报告再次复核。两侧扁平化都准确，最终格式才可信。

由此得到一条明确分工：**人和 AI 只对内容负责，格式由模板与确定性流水线保证。**

- 换学校？换一个 `config/template.docx` 即可，内容与块序列零改动；
- 格式不对？只查两个确定环节——模板骨架是否正确、扁平化分类是否准确——
  两侧中间结果都可检视、可修复，不靠反复试格式，也不让 AI 调格式；
- 内容随便改？改纯文本 Markdown，或在 `gui` 里让 AI 改写，重跑一次 `run`，
  格式自动复原。

块模型是 Word 与文本之间唯一的共同语言：`convert`（Word → 块）与 `build`
（块 → Word）共用同一套定义，已验证 Markdown → docx → Markdown 块级无损——
这正是“扁平化之后能正确复原为模板”这一前提成立的技术保证。在这个前提下写作，
格式不再消耗精力，事半功倍。

## 安装

```bash
git clone <你的仓库地址>
cd thesis-docx-tool
pip install -r requirements.txt
```

一条命令装齐全部第三方依赖（仅两个：Pillow、PyQt6）。各模式需要的模块：

- `run` / `build` / `convert` / `audit` / `toc` 以及 AI 核心（docxai）：**纯标准库**，零第三方模块；
- `fig`：需要 Pillow（`build` 的插图尺寸计算也可选 Pillow，缺失时回退默认尺寸）；
- `gui`：需要 PyQt6（AI 调用本身只用标准库 urllib，不需要任何 AI SDK）。

## 快速开始

### 1. 准备素材（只需一次）

```bash
# 学校论文模板放到 config/，命名为 template.docx（不入库）
cp /path/to/学校论文模板.docx config/template.docx

# 封面信息：复制示例后填写真实信息（cover.json 不入库）
cp config/cover.example.json config/cover.json
# 编辑 config/cover.json，填入题目、姓名、学号等

# 承诺书签名图放到 config/images/signature.png（自备，不入库）
```

### 2. 一条命令出成品（二选一）

路径一：直接写 Markdown。在 `config/章节/` 下按章建文件（如 `01_设计思路.md`，
格式见下「块模型」），架构图 DSL 放 `config/figures/`，自定义插图放 `config/images/`，然后：

```bash
python3 main.py run
```

**章节文件命名约定**：文件名必须以数字开头（如 `01_设计思路.md`、`2_相关技术.md`）。
build 只处理以数字开头的文件，其余（模板示例、整篇拆解产物等）自动跳过并提示，
避免把非章节文件混进来导致正文重复构建。若目录里全是非数字开头的文件
（如 GUI 审阅页导出的单篇论文 md），则退化为全量处理。

路径二：传入已写好的 Word，自动拆解为章节 Markdown 后直接套模板出成品：

```bash
python3 main.py run 用户论文.docx          # 默认剥离原封面/目录（模板自带）
python3 main.py run 用户论文.docx --force  # 拆解产物已存在、确认覆盖时
```

`run` 会依次自动完成：架构图渲染 → 构建第一遍 → LibreOffice 转 PDF → 量目录页码 →
第二遍构建（页码回填）→ 成品格式审计，产物为 `output/论文_v1_YYYYMMDD.docx`。
拆解得到的 Markdown 在 `config/章节/` 中，改完后不带 docx 重跑 `run` 即可重新出稿。

本机没有 LibreOffice（`soffice`）或 `pdftotext` 时，页码回填自动跳过、照常出稿，
在 Word 中按 Ctrl+A 后 F9 即可更新目录；可用 `--no-fig` / `--no-toc` / `--no-audit`
关闭对应阶段。

路径三（推荐用于别人写的、样式不规范的 Word）：图形界面交互式处理。

```bash
python3 main.py gui
```

向导中选择 docx 与处理模式——**默认模式**直接按样式规则解析（离线零成本），
**AI 模式**在规则解析后自动按章调 AI 修正结构分类；进入审阅页后可按章/类型展开
纯文本块列表，用「修改类型」下拉修正标题↔正文等误判，用「插入块」新增正文/标题/
有序列表项/表格/图片，用「删除当前块」与「撤销结构操作」整理结构，或选中块让 AI
按指令改写，文字改动可逐块还原。模板区的「模板预览与修复...」按钮可在解析前先
核对学校模板。确认结构无误后点「确认并生成 Word」，界面自动执行 run
并实时显示日志。首次使用 AI 模式请在菜单「设置 → AI 配置」填写服务商与 API Key
（支持 DeepSeek、豆包火山方舟、通义千问、OpenAI 及任意 OpenAI 兼容接口，
配置保存在不入库的 `config/ai.json`）。

### 高级：分步命令

需要单步调试时才用各原子命令，等价于 `run` 内部的各阶段：

```bash
python3 main.py fig                      # 只渲染架构图
python3 main.py build                    # 只构建（不转 PDF、不审计）
python3 main.py convert 用户论文.docx --strip-front   # 只拆解，产物落 output/
soffice --headless --convert-to pdf output/论文_v1_*.docx
python3 main.py toc output/论文_v1_*.pdf
python3 main.py audit                    # 总览；audit --list 列出全部检查项
```

## 子命令详解

### run（一键流水线，推荐）

```
python3 main.py run [用户论文.docx]
    [--template 模板.docx] [--force] [--keep-front]
    [--no-fig] [--no-toc] [--no-audit]
```

- 不带 docx：直接使用 `config/章节/*.md` 出成品（MD 路线）。
- 带 docx：先把该 docx 拆解为 `config/章节/<名>.md`（图片入 `config/images/<名>_images/`），
  再出成品（Word 路线）；默认剥离原封面/承诺书/目录，`--keep-front` 保留；
  同名 Markdown 已存在时拒绝覆盖，`--force` 允许。
- 开工前一次性检查模板、签名图、章节源是否齐备，缺件直接报齐，不跑到半途才失败。
- `--template` 可指定任意外部模板（默认 `config/template.docx`），两遍构建与成品审计
  都使用该模板；外部模板下审计会自动停用只对默认模板登记的预期偏离表，按真实差异报告。
- 自动转 PDF 量页码依赖 LibreOffice 与 poppler-utils，缺失则跳过回填、不影响出稿。
- 阶段细节与设计约定见 `docxflow/README.md`。

### build

```
python3 main.py build [--template 模板.docx]
```

- 默认模板为 `config/template.docx`，可用 `--template` 指定其它模板。
- 产物：`output/论文_v1_YYYYMMDD.docx` + `output/论文正文.md`（各章合并视图）。
- 封面字段从 `config/cover.json` 读取；承诺书签名图取自 `config/images/signature.png`（自备）。

### convert

```
python3 main.py convert 用户论文.docx [-o 输出.md] [--images-dir 图片目录] [--strip-front]
```

- 把 docx 拆解为与 `章节/*.md` 同格式的 Markdown，图片抽取到独立目录。
- `--strip-front`：丢弃第一个以数字开头的一级标题之前的内容（封面、承诺书、目录）。
- 产出的 Markdown 可直接喂给 `build`，形成闭环。

### audit

```
python3 main.py audit [检查项...] [--product 成品.docx] [--template 模板.docx] [-v] [--strict] [--list]
```

- 默认成品取 `output/` 下 mtime 最新的 `*.docx`（排除 `config/template.docx`）。
- 退出码：有「注意」或「失败」项为 1；`--strict` 时「有意偏离」也算失败。

### toc

```
python3 main.py toc 成品.pdf
```

- 需要系统安装 `poppler-utils`（提供 `pdftotext`）。
- 把各级标题的页码写入 `output/toc_pages.json`，供 `build` 第二遍使用。

### fig

```
python3 main.py fig [图名]
```

- 把 `config/figures/` 下的 `*.md` 表格 DSL 渲染为 PNG，输出到 `config/images/`。
- 不带参数时渲染全部图；带参数时只渲染指定图（省略 `.md` 后缀）。
- 架构图采用**内容与样式分离**：DSL 只描述画什么、画在哪（坐标、尺寸、文本、语义色），
  配色、字体、圆角、箭头等视觉样式由 `docxfig/render.py` 统一决定。
- 详细语法见 `docxfig/README.md`。

### gui（图形界面）

```
python3 main.py gui
```

- 需要 PyQt6（含在 `requirements.txt` 中）；未安装时命令给出安装提示。
- 三页向导：选择 docx/处理模式/模板 → 审阅与修正块序列（手动编辑 + AI 修正）
  → 导出 `config/章节/<名>.md` 并自动执行 run，日志实时显示、可终止。
- 审阅页修复手段：类型互转（标题↔正文、改正文为有序列表等）、插入块
  （正文/标题/有序列表项/表格/图片，图片选文件后复制进本文档素材目录）、
  删除块；插入/删除/改类型/AI 整批改均可逐步撤销，文字改动可逐块还原。
- 「模板预览与修复」页（向导按钮或菜单「设置」）：只读检视骨架部件序列、封面
  字段、承诺书题目占位符、TOC 域、样式、decimal 编号、页面尺寸与页边距，按
  错误/警告列出问题；封面字段可直接写入 `config/cover.json`、签名图可替换
  `config/images/signature.png`；骨架与样式本体请在 Word 中修改模板文件后
  点「重新检测」复验，工具不直接改写模板内部 XML。
- AI 配置在菜单「设置 → AI 配置」，保存在 `config/ai.json`（不入库）；
  AI 只修正结构不改实质内容，每批返回过块数/schema/图片路径三重校验，
  不过整批保留原文。详见 `docxgui/README.md` 与 `docxai/README.md`。

## 块模型（build 与 convert 共用）

`docxbuild.mdparse.parse_md()` 与 `docxconvert.parse.parse_docx()` 共用同一套块格式，
保证 Markdown ↔ docx 可逆闭环：

| 标签 | 内容 | Markdown 写法 |
|---|---|---|
| `h` | `(级别, 标题文本)` | `#` ~ `####` |
| `p` | 段落文本 | 连续非空行 |
| `ol` | 有序列表的单个条目（连续多个归为同一自动编号列表） | `1. 文本` |
| `caption` | 表题 / 图题 | `[表4-1　说明]` |
| `ref` | 参考文献条目 | `[1] 作者. 题名…` |
| `code` | 代码行列表 | ``` 围栏块 |
| `table` | 二维单元格列表 | `| a | b |` 管道表格 |
| `img` | `(图片相对路径, 图题说明)` | `![说明](路径)` |

图片路径**一律相对 `config/` 目录**解析（不是相对 md 文件）：素材照片放进
`config/images/` 后，md 里直接写 `![图x.y 说明](images/照片.png)` 即可；
`fig` 生成的架构图 PNG 也在该目录，引用方式相同（如 `images/fig3_1.png`）。
`run 用户论文.docx` 自动拆解出的图片落 `config/images/<docx名>_images/`，
引用同样写成 `images/<docx名>_images/xxx.png`。

## 模块结构

```
main.py  ── 统一入口，按子命令分发
  │
  ├─ run     → docxflow/    （一键流水线，编排下面四个模块，命令行主入口）
  ├─ gui     → docxgui/     （PyQt6 界面，审阅/修正后经外部进程执行 run）
  │             └─ docxai/  （AI 结构修正/内容改写核心，纯标准库，GUI 与终端版共用）
  ├─ build   → docxbuild/   （Markdown → docx + 目录页码回填）
  ├─ convert → docxconvert/ （docx → Markdown）
  ├─ audit   → docxaudit/   （docx 格式审计）
  └─ fig     → docxfig/     （架构图 DSL → PNG）
```

### docxflow/ —— 一键流水线

| 模块 | 职责 |
|---|---|
| `cli` | run 主流程编排：前置检查、拆解（Word 路线）、fig、build、PDF 量页码二遍构建、audit |

只编排、不重写：所有实际工作都调用各原子模块的公开接口，同一职责全项目只有一份实现。
新增需要多步完成的能力必须接进本流水线，不让用户手工串联。详见 `docxflow/README.md`。

### docxbuild/ —— Markdown → docx

| 模块 | 职责 |
|---|---|
| `docinfo` | 论文著录信息：从 `config/cover.json` 读题目、封面字段、版本号（默认占位符） |
| `layout` | 版式常量：正文区宽度、插图尺寸上限、签名图宽度 |
| `mdparse` | Markdown → 块序列（纯文本，不碰 XML） |
| `fragments` | 块 → OOXML 片段（段落、标题、有序列表项、图、表、代码、参考文献、图题） |
| `template` | 模板骨架处理：命名空间登记、封面填充、目录重建、签名图替换、样式清理、列表编号注入 |
| `tplinspect` | 模板只读检视：骨架/封面字段/承诺书/TOC/样式/编号/页面设置，cover.json 与签名图修复入口（无 Qt 依赖） |
| `toc_pages` | 从渲染出的 PDF 反查各级标题页码，写入 `output/toc_pages.json` |
| `cli` | 主流程编排（`main(argv)`） |

**核心思路**：保留模板骨架（封面、承诺书、目录域、页眉页脚、styles/numbering/theme），
只替换中间的示例正文。因 `mc:Ignorable` 前缀问题，不能用 ElementTree 整体序列化，
采用「字符串拼接 + 逐片段序列化」。

依赖方向：`docinfo`/`layout`/`mdparse` 不依赖别人；`fragments` 用 `layout`+`mdparse`；
`template` 用 `fragments`；`cli` 汇总全部。

### docxconvert/ —— docx → Markdown

| 模块 | 职责 |
|---|---|
| `parse` | 解析 docx 的 `document.xml`，按 body 子元素顺序把段落/表格/图片转成块序列 |
| `extract` | 从 docx 抽取图片到目标目录 |
| `markdown` | 块序列渲染为 `章节/*.md` 格式的 Markdown 文本 |
| `cli` | 主流程编排（`main(argv)`） |

**解析要点**：

- 按 `<w:body>` 子元素顺序**混合遍历** `<w:p>` 与 `<w:tbl>`，避免分别遍历导致顺序错乱。
- 标题识别优先用段落样式 ID（`1`~`4` / `Heading 1` 等），无样式时降级看 `<w:outlineLvl>`。
- 有序列表只读 Word 自动编号属性（`pPr/numPr` + `numbering.xml` 的 `numFmt=decimal`），
  不靠「1.」文本前缀猜测；手写伪编号仍按正文解析，由 GUI 人工或 AI 改判为 `ol`。
- 图题（紧跟图片的「图x.y …」）并入图片块的说明字段；表题单独成 `caption` 块。
- 目录条目（样式名 `toc*`）自动跳过。

### docxaudit/ —— docx 格式审计

| 模块 | 职责 |
|---|---|
| `common` | docx 只读封装、审计上下文、模板/成品路径定位、XML 比对工具 |
| `expected` | 有意偏离登记表（`EXPECTED_DIFFS`），命中则从「注意」降级为「有意偏离」 |
| `checks/` | 五项检查：`stale`（新旧）、`parts`（部件）、`styles`（样式）、`body`（正文）、`meta`（元数据） |
| `report` | 检查结论分级与渲染 |
| `cli` | 主流程编排（`main(argv)` → 退出码） |

**审计思路**：把成品与模板逐部件、逐样式、逐段落地比对，差异分四类：
`[通过]`、`[有意偏离]`（登记在 `EXPECTED_DIFFS`）、`[注意]`、`[失败]`。

### docxfig/ —— 架构图生成

| 模块 | 职责 |
|---|---|
| `render` | PIL 绘制引擎：画布、矩形、菱形、箭头、连线、文字、容器、裁切、保存 |
| `parser` | 解析 `figures/*.md` 的表格 DSL（canvas + 元素表） |
| `cli` | 主流程编排（`generate(name)` → 输出 PNG 到 `config/images/`） |

**内容与样式分离**：DSL 只描述内容（画什么、坐标、尺寸、文本、语义色），
配色、字体、圆角、箭头等视觉样式集中在 `render.py` 的常量中，改样式只动代码。
DSL 语法见 `docxfig/README.md`。

### docxai/ —— AI 结构修正与内容改写核心

| 模块 | 职责 |
|---|---|
| `config` | `config/ai.json` 读写与 provider 预设（DeepSeek/豆包/通义/OpenAI/自定义） |
| `client` | OpenAI 兼容 `chat/completions` 客户端（urllib）：超时、429/5xx 重试、错误脱敏 |
| `blocks_json` | 块 tuple ↔ dict/JSON 边界：转换、schema 校验、AI 返回解析、按 H1 分章 |
| `prompts` | 结构修正/内容改写提示词 |
| `correct` | 按章分批结构修正、单段内容改写编排（校验不过整批保留原文） |

**规则打底 + AI 修正**：规则解析永远先跑出完整结果，AI 只按章修结构，
块数不一致、schema 非法、图片路径越界的返回一律整批拒绝，保证内容零丢失。
纯标准库、无 Qt 依赖，GUI 与后续终端交互版共用。详见 `docxai/README.md`。

### docxgui/ —— PyQt6 图形界面

| 模块 | 职责 |
|---|---|
| `app` | 入口与 PyQt6 延迟导入（缺失给安装提示） |
| `main_window` | 三页编排、菜单、状态栏 |
| `wizard_page` | docx/模式（默认/AI）/模板选择、打开模板预览与修复页 |
| `review_page` | 分类块树、各类型块编辑器、类型互转、插入/删除块与撤销、AI 修正/改写、逐块还原 |
| `blocks_model` | 类型标签与摘要（无 Qt 依赖） |
| `block_ops` | 块结构操作纯函数（类型互转、新建块/图片块），无 Qt 依赖 |
| `template_page` | 模板预览与修复对话框：检视报告、封面字段/签名图修复 |
| `settings_dialog` | AI 配置弹窗（预设、掩码、连接测试） |
| `workers` | QThread（解析/AI）与 QProcess（外部 run） |

构建阶段由外部进程执行 `python3 main.py run`，界面不重写流水线。
详见 `docxgui/README.md`。

## 数据流闭环

```
                        ┌── run（一键流水线）──────────────────────┐
                        │                                          │
config/figures/*.md ──fig──▶ config/images/*.png                   │
                                 │                                 ▼
章节/*.md ──────────────────build──▶ 论文.docx ──PDF/toc──▶ build（页码回填）──▶ audit
    ▲                                │
    │                           convert
    └──────── 二次编辑后重 run ◀──────┘（Word 路线直接由 run 内部完成拆解入 config/章节/）
```

`build` 与 `convert` 共用块模型，已验证 Markdown → docx → Markdown 块级无损。

## 验证与回归测试（规范）

扁平化（`convert`：Word → 块）与渲染引擎（`build`：块 → Word）的正确性，
统一用**现成模板 docx 走闭环对比**来验证，不靠肉眼看输出：

**测试步骤**：

1. 取一个内容完整的 docx 作为输入（推荐直接用 `config/template.docx` 或任意
   已写好的论文 docx），记为 `输入.docx`。
2. 拆解：`python3 main.py convert 输入.docx --strip-front`，产物落 `output/`，
   得到 `输入.md` 与抽取的图片。
3. 用拆解产物出稿：把 `输入.md` 放入 `config/章节/`（确保文件名以数字开头），
   跑 `python3 main.py build --template config/template.docx`，得到 `输出.docx`。
4. 对比 `输入.docx` 与 `输出.docx` 的**正文纯文本**（去掉 XML 标签后逐段比对）：
   - 段落文本应一致（顺序、文字、数量）；
   - 标题层级（1~4 级）应一致；
   - 表格单元格内容、图片数量与引用关系应一致；
   - 有序列表条目数与内容应一致。
5. 若有差异，定位是 `convert` 扁平化丢了内容，还是 `build` 渲染丢了内容，
   再针对对应模块修复。

**为什么用模板文件测**：模板自带封面、承诺书、目录域、样式、编号定义、签名图等
全部骨架部件，是最复杂的输入；用它走一遍闭环能同时覆盖解析与渲染两条链路，
比手写最小样例更能暴露真实问题。

**自动化脚本**（放到 `.cache/` 或临时目录，不入库）：

```python
import re, zipfile
from docxconvert.cli import convert_to
from docxbuild.cli import main as build_main

def texts(docx):
    with zipfile.ZipFile(docx) as z:
        doc = z.read("word/document.xml").decode("utf-8")
    return ["".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p))
            for p in re.findall(r"<w:p\b.*?</w:p>", doc, re.S)]

# convert 输入.docx → 临时 md → build → 输出.docx，再 texts() 对比
```

## 依赖清单

| 依赖 | 用途 | 是否必需 |
|---|---|---|
| Python 3.8+ | 命令行运行环境 | 必需（gui 建议 3.9+） |
| Pillow | build 时计算插图显示尺寸；fig 时渲染架构图 PNG | fig 必需，build 无图可缺省 |
| PyQt6 | gui 图形界面（docxai AI 核心不需要） | 仅 gui 模式必需 |
| LibreOffice (`soffice`) | docx 转 PDF（量目录页码用） | 量页码时必需 |
| poppler-utils (`pdftotext`) | 从 PDF 提取文本量页码 | 量页码时必需 |
| OpenAI 兼容大模型 API | gui 的 AI 模式/AI 改写（DeepSeek/豆包/通义/OpenAI 等） | 仅 AI 功能必需 |

## 常见问题

- **build 报找不到模板**：把模板放到 `config/template.docx` 或用 `--template` 指定。
- **build 报找不到签名图**：在 `config/images/signature.png` 放签名图，或从 `config/cover.json`/模板调整。
- **convert 后标题层级不对**：用户 docx 的标题样式可能不是模板的 `1`/`2`/`3`/`4`，
  工具会尝试匹配 `Heading 1` 等常见样式名；匹配不到的当正文处理。
  可用 `python3 main.py gui` 选 AI 模式自动修正结构，或在审阅页手动调整后再出稿。

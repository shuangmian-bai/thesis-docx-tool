# thesis-docx-tool

论文 Word 排版与拆解工具。日常使用只需一个子命令 **run**，一条命令端到端出成品：

- **run**：一键流水线。自动串联 拆解（可选）→ 架构图 → 构建 → 转 PDF 量目录页码 → 二次构建 → 格式审计，
  支持两条路径：直接写 Markdown，或传入已写好的 Word 自动拆解后重排。

其余五个是分步原子命令，供单步调试与高级用途（run 在内部编排它们）：

- **build**：把按章拆分的 Markdown（`章节/*.md`）套用学校论文模板骨架，生成符合格式的 Word 论文
- **convert**：把已写好的 Word 论文反向拆解为同格式 Markdown，便于二次编辑
- **audit**：审计生成的 docx 是否贴合模板格式，避免排版跑偏
- **toc**：从渲染出的 PDF 反查目录页码，回填 TOC 域
- **fig**：把 `figures/*.md` 中用表格 DSL 描述的架构图渲染为 PNG，供 build 嵌入

依赖仅 Python 标准库（`zipfile` / `xml.etree` / `re` / `json` / `subprocess`），
**不用 python-docx、不用 pandoc**。`build` 的插图尺寸计算可选 `Pillow`，无图时可缺省；
`fig` 子命令需要 `Pillow` 渲染 PNG。

## 安装

```bash
git clone <你的仓库地址>
cd thesis-docx-tool
pip install -r requirements.txt   # 仅 Pillow；只用 build/convert/audit/toc 且无插图时可跳过
```

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

## 块模型（build 与 convert 共用）

`docxbuild.mdparse.parse_md()` 与 `docxconvert.parse.parse_docx()` 共用同一套块格式，
保证 Markdown ↔ docx 可逆闭环：

| 标签 | 内容 | Markdown 写法 |
|---|---|---|
| `h` | `(级别, 标题文本)` | `#` ~ `####` |
| `p` | 段落文本 | 连续非空行 |
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
  ├─ run     → docxflow/    （一键流水线，编排下面四个模块，面向用户的主入口）
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
| `fragments` | 块 → OOXML 片段（段落、标题、图、表、代码、参考文献、图题） |
| `template` | 模板骨架处理：命名空间登记、封面填充、目录重建、签名图替换、样式清理 |
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

**内容与样式分离**：DSL 只描述内容（画什么、坐标、尺寸、文本、语义色名），
配色、字体、圆角、箭头等视觉样式集中在 `render.py` 的常量中，改样式只动代码。
DSL 语法见 `docxfig/README.md`。

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

## 依赖清单

| 依赖 | 用途 | 是否必需 |
|---|---|---|
| Python 3.8+ | 运行环境 | 必需 |
| Pillow | build 时计算插图显示尺寸；fig 时渲染架构图 PNG | fig 必需，build 无图可缺省 |
| LibreOffice (`soffice`) | docx 转 PDF（量目录页码用） | 量页码时必需 |
| poppler-utils (`pdftotext`) | 从 PDF 提取文本量页码 | 量页码时必需 |

## 常见问题

- **build 报找不到模板**：把模板放到 `config/template.docx` 或用 `--template` 指定。
- **build 报找不到签名图**：在 `config/images/signature.png` 放签名图，或从 `config/cover.json`/模板调整。
- **convert 后标题层级不对**：用户 docx 的标题样式可能不是模板的 `1`/`2`/`3`/`4`，
  工具会尝试匹配 `Heading 1` 等常见样式名；匹配不到的当正文处理，可手工调整 MD。

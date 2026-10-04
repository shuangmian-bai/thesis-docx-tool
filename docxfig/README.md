# docxfig —— 架构图生成模块

把论文架构图从「写死在 Python 代码里」改为**内容与样式分离**：

- **内容层**：`config/figures/*.md` 用 markdown 表格描述画什么、画在哪（坐标、尺寸、文本、语义色）。
- **样式层**：配色、字体、画布、圆角、箭头等视觉样式由本模块的 `render.py` 决定，改样式只动代码。

## 用法

```bash
python3 main.py fig              # 生成 config/figures/ 下全部图
python3 main.py fig fig3_1       # 只生成指定图（省略 .md 后缀）
```

输出 PNG 到 `config/images/`，文件名与 md 同名（如 `fig3_1_architecture.md` → `fig3_1_architecture.png`），
供 `build` 嵌入论文。

## DSL 语法

每个图一个 `.md` 文件，由两部分组成：

### 1. 画布声明

```
canvas: 1350 1280
```

第一个数是画布宽（固定取 1350，由版心宽度与最小字号倒推，见 `render.py` 注释），
第二个数是画布高（按内容自定）。

### 2. 元素表格

六列：`type | pos | size | color | text | extra`

| 列 | 说明 |
|---|---|
| `type` | `container` / `box` / `diamond` / `arrow` / `line` / `text` |
| `pos` | `x,y`（设计单位，左上角） |
| `size` | `w,h`（box/container/diamond）或 `x2,y2`（arrow/line 的终点） |
| `color` | 语义色名：`blue`/`orange`/`green`/`grey`/`purple`/`red`/`white`，留空用默认 |
| `text` | 显示文本，`\n` 换行（box 内自动折行） |
| `extra` | 可选属性，`key:value` 逗号分隔 |

`extra` 支持的键：

| 键 | 适用元素 | 说明 |
|---|---|---|
| `font` | box/diamond/text | `box`(34) / `small`(28) / `tiny`(24) / `head`(36) |
| `radius` | box | 圆角半径，`0` 为直角 |
| `fill` | box | 单独覆盖填充色（语义色名），如 `fill:white` |
| `line` | box | 单独覆盖边框色（语义色名），如 `line:orange` |
| `dashed` | arrow | `true` 画虚线箭头 |
| `label` | arrow | 箭头中点文字 |
| `width` / `head` | arrow | 线宽 / 箭头大小（默认 4 / 16） |

### 元素类型说明

- **container**：带标题的容器框，浅底（`CONTAINER_FILL`）+ 左上角标题，用于分组。
- **box**：圆角矩形，文本居中自动折行；`radius:0` 变直角；`fill`/`line` 可单独覆盖。
- **diamond**：判断菱形，x/y 为外包矩形左上角。
- **arrow**：带箭头连线，`size` 列为终点 `x2,y2`；`dashed:true` 画虚线。
- **line**：普通连线（无箭头）。
- **text**：在指定位置写文字（左上角对齐）。

### 示例

```markdown
# fig3_1: 注入攻击系统总体架构

canvas: 1350 1280

| type | pos | size | color | text | extra |
|---|---|---|---|---|---|
| container | 40,100 | 1270,516 | blue | 推送层（PC 端） | |
| box | 60,176 | 390,120 | blue | 数据源层\n解复用·提取配置 | |
| arrow | 675,616 | 675,666 | | | label:VCU1/TCP |
```

## 模块结构

| 文件 | 职责 |
|---|---|
| `render.py` | PIL 绘制引擎（样式层）：画布、box、diamond、arrow、line、text、container、trim、save |
| `parser.py` | 解析 `figures/*.md` 的表格 DSL，返回元素列表 |
| `cli.py` | 编排：`generate(name)` 遍历图定义 → 解析 → 绘制 → 保存 |
| `__init__.py` | 定义 `HERE = thesis-docx-tool/` 根目录 |

## 样式常量（改样式只动这里）

定义在 `render.py` 顶部：

| 常量 | 值 | 说明 |
|---|---|---|
| `CANVAS_W` | 1350 | 画布宽（设计单位），由版心与最小字号倒推 |
| `SCALE` | 2 | 渲染放大倍数，只影响清晰度不影响版面 |
| `MARGIN` | 36 | 裁切后四周外边距 |
| `PALETTE` | 7 语义色 | `(填充色, 边框色)` 字典 |
| `CONTAINER_FILL` | 7 语义色 | 容器浅底色（近白 tint） |
| `F_BOX`/`F_SMALL`/`F_TINY`/`F_HEAD` | 34/28/24/36 | 预设字号 |
| `FONT_PATH` | NotoSansCJK-Regular.ttc, index=2 | 中文字体（SC 简体） |

## 设计约定

- **坐标用设计单位**：所有 `pos`/`size`/`width`/`radius` 都用设计单位传入，
  `render.py` 内部统一乘 `SCALE`，保证改清晰度（`SCALE`）不动版面。
- **颜色用语义名**：DSL 中不写十六进制，只写 `blue`/`orange` 等语义名；
  具体色值在 `PALETTE` 中统一定义。
- **容器与内部 box 不同底**：`container` 用 `CONTAINER_FILL`（近白），
  内部 `box` 用 `PALETTE` 的填充色，保证层次分明。

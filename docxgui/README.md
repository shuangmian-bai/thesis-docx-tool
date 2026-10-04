# docxgui —— PyQt6 图形界面（gui 子命令）

run 的交互式前端：选择 Word 文档与处理模式，审阅扁平化（块序列）结果并手动或
借助 AI 修正，确认后由外部进程执行 run 一键出稿。

启动：

```bash
pip install -r requirements.txt
python3 main.py gui
```

本包是项目内**唯一**依赖 PyQt6 的模块；未安装时入口打印安装指引并退出，
命令行各子命令不受影响。AI 逻辑（`docxai`）、解析（`docxconvert`）、
构建编排（`docxflow`）都不在本包内重复实现。

## 三页流程

1. **向导页 `wizard_page.py`**：选 docx、选处理模式（默认/AI）、选模板
   （可外部指定，与 `run --template` 等价）、可选丢弃前置封面内容；
   「模板预览与修复...」按钮打开模板对话框；
   AI 配置状态实时显示，入口在菜单「设置 → AI 配置」。
2. **审阅页 `review_page.py`**：
   - 左：按 H1 章分组的块树，纯文本展示，前缀类型标签（标题1-4/正文/有序列表/
     引用/图题表题/文献/代码/表格/图片），支持类型筛选；
   - 右：块编辑器——标题改级别、文本/代码直接改、表格可增删行列、
     图片显示缩略图（路径只读，图题可改）；
   - 结构修复（纯函数在 `block_ops.py`）：「修改类型」下拉做类型互转
     （表格/图片不可转）；「插入块 ▾」插入正文/标题/有序列表项/表格，
     「插入图片」选本地文件后自动复制进本文档素材目录（本文档无素材目录时
     用 `config/images/<docx名>_images/`，重名自动加序号），块引用写相对
     config 的路径；「删除当前块」有确认框；「撤销结构操作」回退最近一次
     插入/删除/类型转换/AI 整批改（文字编辑不在其列，靠逐块「还原」）；
   - AI 操作：整篇结构修正（按章分批、进度可取消）、单块按指令改写，
     AI 改动标记 `[AI 已改]` 且可逐块「还原」；
   - 确认前一切修改只在内存，不落盘。
3. **构建页 `build_page.py`**：确认后块序列写成 `config/章节/<docx名>.md`
   （图片在解析阶段已抽签到 `config/images/<docx名>_images/`），随后
   QProcess 执行 `python3 main.py run [--template ...]`，日志实时滚动、
   可终止，结束后可打开 output 目录。

## 模块

| 文件 | 职责 |
|---|---|
| `app.py` | 入口；延迟导入 PyQt6，缺失给安装提示 |
| `main_window.py` | 三页 QStackedWidget 编排、菜单、状态栏、页面间数据传递 |
| `wizard_page.py` | docx/模式/模板选择，`open_template` 信号打开模板对话框 |
| `review_page.py` | 块树、类型筛选、各类块编辑器、类型互转/插入/删除/撤销、AI 操作与还原 |
| `block_ops.py` | 块结构操作纯函数（can_convert/convert/new_block/new_image_block，无 Qt 依赖） |
| `template_page.py` | 模板预览与修复对话框：`docxbuild.tplinspect` 报告、cover.json 与签名图修复、`template_changed` 信号 |
| `blocks_model.py` | 块类型中文标签、摘要、可编辑性判定（无 Qt 依赖） |
| `settings_dialog.py` | AI 配置弹窗（provider 预设、key 掩码、连接测试） |
| `workers.py` | ParseWorker / StructureWorker / RewriteWorker（QThread）与 RunProcess（QProcess） |

## 线程模型

- 解析与 AI 调用在 QThread 中执行，结果/进度/失败一律经信号回传 UI；
- run 构建在独立 QProcess 中执行，界面只收 stdout 流，构建逻辑零重写；
- UI 线程不做任何阻塞操作；API Key 不出现在日志与错误信息中。

## 依赖与被依赖

- **依赖**：PyQt6（含在统一的 requirements.txt 中）、`docxai`、`docxconvert`、`docxflow`。
- **被谁用**：仅 `main.py gui` 子命令；后续终端交互版直接复用 `docxai`，
  不经过本包。

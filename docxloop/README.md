# docxloop —— 闭环测试引擎

对输入 docx 执行「扁平化 → 合成 → 严格对比」，验证扁平化器（word→块）与
渲染器（块→word）完全对应：内容与格式一致才算通过。

## 用法

```bash
python3 -m docxloop.cli 论文.docx        # 单文件
python3 -m docxloop.cli ./papers/        # 文件夹批量
python3 -m docxloop.cli 论文.docx --no-reuse   # 强制重新扁平化+合成
```

模板默认是**每个输入文件自己**：验证的是「拆解→重建」的无损性，模板与输入
同源才能排除格式差异干扰。

## 对比口径

只对比**正文区**（分节符段落之后，扁平化→合成的作用范围），逐项严格对应：

- 段落：文本、样式（pStyle）、对齐（jc）、行距/段距（spacing）、缩进（ind）、
  字体（rFonts）、字号（sz）、加粗（b）
- 表格：单元格文本、段落格式（jc/ind）、跨列合并（gridSpan）
- 图片：尺寸（cx/cy EMU）

**无标记、无跳过、无容差**——任何差异都是扁平化器或渲染器的缺陷，修复缺陷
而非掩盖；报告里所有差异红色标注。

## 模块组成

| 文件 | 职责 |
|---|---|
| `cli.py` | 参数解析、结果摘要、报告落盘 |
| `engine.py` | 输入解析、单文件/批量编排（run_closed_loop / run_loop） |
| `compare.py` | 严格对比：正文区段落/表格/图片逐项对应 |
| `report.py` | 自包含 HTML 报告 |

报告默认写 `.cache/loop_report.html`。

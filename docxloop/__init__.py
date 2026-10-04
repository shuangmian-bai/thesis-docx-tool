"""闭环测试引擎：convert → build → 全盘对比 → 可视化报告。

输入一组 docx（文件夹或单文件），逐个走拆解→构建闭环，对比原文档与输出文档的
**全部内容与格式**（正文、封面、目录、承诺书、页眉页脚、样式、图片尺寸），
生成 HTML 可视化报告供人验收。

特殊元素（如承诺书签名图）由 `rules.py` 登记为「预期差异」，在报告里标注原因，
**不是跳过检查**——所有元素都会被检查，只是差异被解释。

用法：
    python3 main.py loop 输入文件夹 [--template 模板.docx] [--output 报告.html]
    python3 main.py loop 单个.docx
"""
import os

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

"""闭环测试引擎：验证扁平化（word→块）与合成（块→word）对正文区完全对应。

word 扁平化 → 合成 → 严格对比：内容与格式完全对应才算通过，无标记、无跳过、
无容差——任何差异都是扁平化器或渲染器的缺陷，修复缺陷而非掩盖。
"""
import os

#: 工具根目录（thesis-docx-tool/）
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

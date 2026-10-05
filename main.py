#!/usr/bin/env python3
"""thesis-docx-tool 统一入口。

面向用户只有两个子命令：

  run   一键流水线：从用户 docx（或已拆解的哈希工作目录）端到端生成成品 Word
  gui   PyQt6 图形界面：选 docx 与模板 → 审阅/修复（可 AI 修正）→ 一键出稿

run 的两种用法（同一入口，靠第一个参数区分）：

    python3 main.py run 论文.docx [--template 模板.docx]   # 拆解 word → md 存哈希工作目录 → 构建
    python3 main.py run <16位哈希> [--template 模板.docx]  # 复用工作目录 md 重新构建

拆解产物落在哈希工作目录（.cache/run_work/<hash>/），修改章节后传哈希重跑即可。
封面、承诺书、签名图等由模板骨架提供，工具不提取、不填充、不缺件报错。

gui 需先 pip install -r requirements.txt 安装 PyQt6（run 不依赖 PyQt6）。
"""
import sys


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        return 0

    cmd = sys.argv[1]
    rest = sys.argv[2:]

    if cmd == "run":
        from docxflow.cli import main as run_main
        return run_main(rest) or 0
    elif cmd == "gui":
        from docxgui.app import main as gui_main
        return gui_main(rest)
    else:
        print(f"未知子命令：{cmd}", file=sys.stderr)
        print(__doc__, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

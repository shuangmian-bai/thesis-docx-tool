#!/usr/bin/env python3
"""thesis-docx-tool 统一入口。

桌面应用：直接运行（双击 exe 或 python main.py，不带参数）默认进入 gui 图形界面。

面向用户只有两个子命令：

  run   一键流水线：从「工作目录」端到端生成成品 Word（word/哈希/目录 三种来源）
  gui   PyQt6 图形界面：选 docx 与模板 → 审阅/修复（可 AI 修正）→ 一键出稿

run 的三种用法（同一入口，靠第一个参数区分，本质都是「目录模式」）：

    python3 main.py run 论文.docx [--template 模板.docx]   # word → 哈希目录 → 拆解 → 构建
    python3 main.py run <16位哈希> [--template 模板.docx]  # 定位目录 → 构建（复用 md）
    python3 main.py run 目录 [--template 模板.docx]        # 目录（含 章节/*.md）→ 直接构建

「目录」是核心契约：一个含「章节/*.md」（+ 可选 figures/、images/）的独立目录。
word 拆解、AI 生成、手写都是「填充这个目录」的方式；哈希只是 word 的自动命名
（内容 SHA256 前 16 位）。封面、承诺书、签名图等由模板骨架提供，不提取不填充。

gui 需先 pip install -r requirements.txt 安装 PyQt6（run 不依赖 PyQt6）。
"""
import sys


def main():
    if len(sys.argv) < 2:
        # 无参数默认 gui（桌面应用双击即用，带 cmd 窗口时日志可见）
        from docxgui.app import main as gui_main
        return gui_main([])
    if sys.argv[1] in ("-h", "--help"):
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

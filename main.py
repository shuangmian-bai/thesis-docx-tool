#!/usr/bin/env python3
"""thesis-docx-tool 统一入口。

日常使用只需一个子命令 **run**（一键流水线，端到端出成品）：

    python3 main.py run                 # config/章节/*.md → 成品 Word
    python3 main.py run 用户论文.docx    # 先拆解用户 Word，再出成品

其余七个是分步原子命令/图形入口，供单步调试、高级用途与交互式处理：

  build     把 `章节/*.md` 合成 Word 论文（套用模板骨架）
  convert   把用户撰写的论文 docx 拆解为 `章节/*.md` 格式的 Markdown
  audit     审计成品 docx 是否贴合模板格式
  loop      闭环测试：convert→build 后全盘对比原文档与输出文档，生成可视化报告
  toc       从渲染出的 PDF 反查标题页码，写入 `toc_pages.json`
  fig       从 `config/figures/*.md` 生成论文架构图（内容与样式分离）
  gui       PyQt6 图形界面：选 docx 与处理模式（默认/AI）→ 审阅扁平化结果
            → 手动或 AI 修正 → 一键 run 出稿（gui 需先 pip install -r requirements.txt 装 PyQt6）

用法：

    python3 main.py run [用户论文.docx] [--template 模板.docx] [--no-fig/--no-toc/--no-audit]
    python3 main.py gui
    python3 main.py build [--template 模板.docx]
    python3 main.py convert 用户论文.docx [--strip-front]
    python3 main.py audit [parts styles ...] [--product 成品.docx]
    python3 main.py loop <文件/文件夹/哈希> [--template 模板.docx] [--no-reuse]
    python3 main.py toc 成品.pdf
    python3 main.py fig [图名]

各子命令的详细参数见 `python3 main.py <子命令> --help`。
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
    elif cmd == "build":
        from docxbuild.cli import main as build_main
        return build_main(rest) or 0
    elif cmd == "convert":
        from docxconvert.cli import main as convert_main
        return convert_main(rest) or 0
    elif cmd == "audit":
        from docxaudit.cli import main as audit_main
        return audit_main(["audit_docx.py"] + rest)
    elif cmd == "loop":
        from docxloop.cli import main as loop_main
        return loop_main(rest) or 0
    elif cmd == "toc":
        from docxbuild.toc_pages import main as toc_main
        return toc_main(rest) or 0
    elif cmd == "fig":
        from docxfig.cli import main as fig_main
        return fig_main(rest) or 0
    elif cmd == "gui":
        from docxgui.app import main as gui_main
        return gui_main(rest)
    else:
        print(f"未知子命令：{cmd}", file=sys.stderr)
        print(__doc__, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

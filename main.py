#!/usr/bin/env python3
"""thesis-docx-tool 统一入口。

四个子命令对应论文排版与拆解的完整流程：

  build     把 `章节/*.md` 合成 Word 论文（套用模板骨架）
  convert   把用户撰写的论文 docx 拆解为 `章节/*.md` 格式的 Markdown
  audit     审计成品 docx 是否贴合模板格式
  toc       从渲染出的 PDF 反查标题页码，写入 `toc_pages.json`

用法：

    python3 main.py build [--template 模板.docx]
    python3 main.py convert 用户论文.docx [--strip-front]
    python3 main.py audit [parts styles ...] [--product 成品.docx]
    python3 main.py toc 成品.pdf

各子命令的详细参数见 `python3 main.py <子命令> --help`。
"""
import sys


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        return 0

    cmd = sys.argv[1]
    rest = sys.argv[2:]

    if cmd == "build":
        from docxbuild.cli import main as build_main
        return build_main(rest) or 0
    elif cmd == "convert":
        from docxconvert.cli import main as convert_main
        return convert_main(rest) or 0
    elif cmd == "audit":
        from docxaudit.cli import main as audit_main
        return audit_main(["audit_docx.py"] + rest)
    elif cmd == "toc":
        from docxbuild.toc_pages import main as toc_main
        return toc_main(rest) or 0
    else:
        print(f"未知子命令：{cmd}", file=sys.stderr)
        print(__doc__, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

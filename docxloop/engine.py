"""闭环引擎：对输入 docx 扁平化→合成→严格对比。

不含任何流程逻辑：扁平化与合成都调用应用层唯一实现（run_pipeline），
本模块只负责「输入解析 → 触发流水线 → 严格对比」。
"""
import os
from typing import List, Optional

from docxloop.compare import CompareResult, compare


def run_closed_loop(src_path: str, reuse: bool = True) -> CompareResult:
    """对单个 docx 跑闭环：扁平化 → 合成（自己作模板）→ 严格对比。

    模板用输入文件本身——验证的是「拆解→重建」的无损性，模板与输入同源
    才能排除格式差异干扰。
    """
    from docxflow.workdir import file_hash, work_dir_for
    from docxflow.cli import run_pipeline

    wdir = work_dir_for(file_hash(src_path))
    out_docx = None
    if reuse:
        out_dir = os.path.join(wdir, "output")
        if os.path.isdir(out_dir):
            for f in sorted(os.listdir(out_dir)):
                if f.endswith(".docx"):
                    out_docx = os.path.join(out_dir, f)

    if out_docx is None:
        # 走应用层唯一流水线（自己作模板）；跳过审计（只读报告无信息量）、
        # 产物留在工作目录
        out_docx = run_pipeline(src_path, src_path,
                                no_audit=True, copy_to_output=False)

    return compare(src_path, out_docx)


def run_loop(spec: str, reuse: bool = True) -> List[CompareResult]:
    """解析输入（文件/文件夹），逐个跑闭环。"""
    paths: List[str] = []
    if os.path.isfile(spec) and spec.endswith(".docx"):
        paths.append(os.path.abspath(spec))
    elif os.path.isdir(spec):
        for f in sorted(os.listdir(spec)):
            if f.endswith(".docx") and not f.startswith("~$"):
                paths.append(os.path.abspath(os.path.join(spec, f)))
    else:
        print(f"[注意] 无法解析输入：{spec}（应为 docx 文件或文件夹）")
        return []

    return [run_closed_loop(p, reuse=reuse) for p in paths]

"""闭环引擎：对输入 docx 跑应用层流水线，然后全盘对比原文档与成品。

本模块**不含任何流程逻辑**：convert/build/封面提取/章节写入等一律调用
docxflow 应用层的 run_pipeline()（与命令行 run、GUI 构建同一条流水线），
这里只负责「输入解析 → 触发流水线 → 全盘对比」。

哈希隔离与索引复用 docxflow.workdir 的唯一实现（与 run/gui 共用
.cache/run_work/），同一论文在三条入口下看到的是同一个工作目录。
"""
import os
from typing import List, Optional

from docxflow.workdir import (file_hash, is_hash16, load_index, resolve_hash,
                              source_for_hash, work_dir_for)
from docxloop.compare import CompareResult, compare


def resolve_input(spec: str) -> Optional[str]:
    """把命令行参数解析为实际文件路径。

    spec 可以是文件路径或 16 位哈希：
    - 文件路径：返回绝对路径（流水线内部会算哈希定位工作目录）
    - 哈希：从索引查原文件路径；工作目录或原文件不存在时返回 None
    """
    if is_hash16(spec):
        h = spec.lower()
        if resolve_hash(h) is None:
            return None
        return source_for_hash(h)
    if os.path.isfile(spec) and spec.endswith(".docx"):
        return os.path.abspath(spec)
    return None


def run_closed_loop(src_path: str, template_path: Optional[str] = None,
                    reuse: bool = True) -> CompareResult:
    """对单个 docx 跑闭环，返回对比结果。

    template_path 为 None 时用输入文件本身作模板——闭环测试验证的是
    「拆解→重建」的无损性，模板与输入同源才能排除格式差异干扰。
    reuse=True 时，若哈希对应的工作目录已存在且有输出 docx，直接用已有产物对比
    （适合反复调整对比引擎）；False 时强制重跑流水线。
    """
    template_path = template_path or src_path
    h = file_hash(src_path)
    wdir = work_dir_for(h)

    out_docx = None
    if reuse:
        out_dir = os.path.join(wdir, "output")
        if os.path.isdir(out_dir):
            for f in sorted(os.listdir(out_dir)):
                if f.endswith(".docx"):
                    out_docx = os.path.join(out_dir, f)

    if out_docx is None:
        # 走应用层唯一流水线（与 run 行为完全一致，仅模板与输出落点不同）；
        # no_audit：audit 是只读报告且对比已由本模块负责，跑两遍没有信息量；
        # force：--no-reuse 明确要求重新拆解，允许覆盖已有产物；
        # copy_to_output：闭环产物留在工作目录即可，不占用全局 output/
        from docxflow.cli import run_pipeline
        out_docx = run_pipeline(
            docx=src_path, work_dir=wdir, template=template_path,
            force=True, no_audit=True, copy_to_output=False)

    return compare(src_path, out_docx)


def run_batch(input_specs: List[str], template_path: Optional[str] = None,
              reuse: bool = True) -> List[CompareResult]:
    """批量跑闭环。input_specs 可以是文件、文件夹或哈希。

    template_path 为 None 时每个输入文件各自作自己的模板。
    """
    paths = []
    for spec in input_specs:
        resolved = resolve_input(spec)
        if resolved:
            paths.append(resolved)
        elif os.path.isdir(spec):
            for f in sorted(os.listdir(spec)):
                if f.endswith(".docx") and not f.startswith("~$"):
                    paths.append(os.path.join(spec, f))
        else:
            print(f"[注意] 无法解析输入：{spec}"
                  f"（文件不存在或哈希对应的工作目录不存在）")

    results = []
    for p in paths:
        results.append(run_closed_loop(p, template_path, reuse=reuse))
    return results

"""哈希工作目录管理（应用层唯一实现）。

run / loop / gui 三条入口共用同一套哈希隔离机制，本模块是唯一实现处：
- 以输入 docx 的 SHA256 前 16 位作为工作目录名，防止多论文共享可变目录互相污染；
- 索引文件 index.json（哈希 → 原文件绝对路径）与工作目录并列存放，供哈希定位回溯；
- 任何入口不得自行计算哈希、创建目录或维护索引。

目录结构（.cache/run_work/）：
    .cache/run_work/
    ├── index.json              # {哈希: 原文件绝对路径}
    └── <hash16>/
        ├── src.md              # convert 产物（整篇 Markdown）
        ├── images/             # 抽取的图片
        ├── 章节/               # build 输入（按 H1 拆分的数字开头 md）
        ├── output/             # build 输出
        └── toc.json            # 目录页码表
"""
import hashlib
import json
import os
from typing import Dict, Optional

from docxflow import HERE

#: 哈希隔离工作根目录（run / loop / gui 共用）
WORK_ROOT = os.path.join(HERE, ".cache", "run_work")


def file_hash(path: str) -> str:
    """计算文件 SHA256，取前 16 位十六进制。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def is_hash16(spec: str) -> bool:
    """判断命令行参数是否为 16 位哈希（十六进制）。"""
    return len(spec) == 16 and all(c in "0123456789abcdef" for c in spec.lower())


def work_dir_for(hash16: str) -> str:
    """返回哈希对应的工作目录路径（不要求存在）。"""
    return os.path.join(WORK_ROOT, hash16)


def load_index() -> Dict[str, str]:
    """加载哈希索引文件（哈希 → 原文件绝对路径）。"""
    idx_path = os.path.join(WORK_ROOT, "index.json")
    if os.path.exists(idx_path):
        with open(idx_path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_index(index: Dict[str, str]) -> None:
    """保存哈希索引文件。"""
    os.makedirs(WORK_ROOT, exist_ok=True)
    with open(os.path.join(WORK_ROOT, "index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def prepare_work_dir(docx_path: str) -> str:
    """按 docx 内容哈希创建/定位工作目录并登记索引，返回工作目录路径。

    run 的 Word 路线、GUI 选文件、loop 的输入定位统一走这里。
    """
    docx_path = os.path.abspath(docx_path)
    h = file_hash(docx_path)
    wdir = work_dir_for(h)
    os.makedirs(wdir, exist_ok=True)
    index = load_index()
    index[h] = docx_path
    save_index(index)
    return wdir


def resolve_hash(hash16: str) -> Optional[str]:
    """哈希定位：工作目录存在则返回其路径，否则返回 None。

    索引中有原文件路径记录时也返回 (工作目录) —— 调用方负责区分
    「继续构建」（只需工作目录）与「重新拆解」（需原文件仍存在）。
    """
    wdir = work_dir_for(hash16)
    return wdir if os.path.isdir(wdir) else None


def source_for_hash(hash16: str) -> Optional[str]:
    """从索引查哈希对应的原文件绝对路径，不存在或文件已删除返回 None。"""
    orig = load_index().get(hash16)
    if orig and os.path.isfile(orig):
        return orig
    return None

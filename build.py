#!/usr/bin/env python3
"""打包脚本 —— 把入口 main.py 打成免 Python 的 Windows 多文件程序（带 cmd 窗口）。

用法（在 thesis-docx-tool/ 目录下执行）：
    .venv\\Scripts\\python.exe build.py

产物 dist/thesis-docx-tool/：
    thesis-docx-tool.exe       双击即用：默认进入 gui，带 cmd 窗口（run/gui 日志可见）
    _internal/                 Python 运行时 + 全部第三方与自建模块
    config/                    用户数据目录：模板 template.docx、AI 配置 ai.json、
                               figures/ 架构图 DSL、images/ 图片、章节/ 章节源，由用户放入

设计要点：
  - console + onedir：带 cmd 窗口（--console，不用 --windowed），多文件模式启动快。
  - config/ 是「用户数据」不打包：模板/AI 配置/章节源自备，产物只预建空目录；
    模板缺失时程序会明确提示「放 config/template.docx 或 --template 指定」。
  - 自建包一律 --collect-submodules：防止静态分析漏掉子模块。
  - 构建前后自动清理残留：build/、生成的 .spec、__pycache__ / *.pyc。
  - soffice / pdftotext 是系统级外部依赖，不打包；缺失时页码回填自动降级。

分发：整个 dist/thesis-docx-tool/ 目录拷到目标机器即可，无需安装 Python。
"""

import os
import shutil
import subprocess
import sys
import time

# 控制台可能是 GBK，打印中文会 UnicodeEncodeError，统一切到 UTF-8
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
NAME = "thesis-docx-tool"
ENTRY = "main.py"

DIST_DIR = os.path.join(ROOT, "dist")
BUILD_DIR = os.path.join(ROOT, "build")
APP_DIR = os.path.join(DIST_DIR, NAME)

# PyInstaller 实际输出到暂存目录，再由 install() 搬进 dist/<NAME>/。
# 原因：dist/<NAME>/ 常被 IDE 或资源管理器占用——目录内容可以清空，
# 但目录本身删不掉，而 PyInstaller 的 COLLECT 会强制 rmtree 目标目录并失败。
STAGE_DIR = os.path.join(BUILD_DIR, "_dist")
STAGE_APP = os.path.join(STAGE_DIR, NAME)

# 自建包：整包收集全部子模块
PACKAGES = (
    "docxflow", "docxbuild", "docxconvert", "docxaudit",
    "docxfig", "docxai", "docxgui", "docxloop",
)

# 明确用不到、即便装了也不打进去的第三方库
EXCLUDES = (
    "tkinter", "PyQt5", "PySide2", "PySide6",
    "matplotlib", "pytest", "IPython",
)

# 产物里预建的用户数据子目录（模板放 config/ 根，其余按目录结构）
CONFIG_SUBDIRS = ("章节", "figures", "images")

# 清理 __pycache__ 时跳过的目录：虚拟环境 / 构建产物 / 用户数据 / 测试样本
SOURCE_SKIP_DIRS = {
    ".venv", ".venv_win", "dist", "build", "config", "output", ".cache", ".git",
    "test_word",
}


def _human(nbytes: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if nbytes < 1024 or unit == "GB":
            return "%.1f %s" % (nbytes, unit)
        nbytes /= 1024.0
    return "%.1f GB" % nbytes


def _dir_size(path: str) -> int:
    total = 0
    for dirpath, _dirnames, filenames in os.walk(path):
        for fn in filenames:
            fp = os.path.join(dirpath, fn)
            try:
                total += os.path.getsize(fp)
            except OSError:
                pass
    return total


def _rm(path: str) -> None:
    if os.path.isdir(path) and not os.path.islink(path):
        shutil.rmtree(path, ignore_errors=True)
    elif os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def _clean_pycache(root: str, skip_top: set) -> int:
    """删除 root 下所有 __pycache__ 与 *.pyc / *.pyo，返回删除条目数。"""
    removed = 0
    for dirpath, dirnames, filenames in os.walk(root, topdown=True):
        rel = os.path.relpath(dirpath, root)
        top = rel.split(os.sep)[0]
        if top in skip_top:
            dirnames[:] = []
            continue

        for name in list(dirnames):
            if name == "__pycache__":
                _rm(os.path.join(dirpath, name))
                dirnames.remove(name)
                removed += 1

        for fn in filenames:
            if fn.endswith((".pyc", ".pyo")):
                _rm(os.path.join(dirpath, fn))
                removed += 1
    return removed


def clean_before() -> None:
    """清理上一次构建留下的产物，避免旧文件混进新包。"""
    print("[清理] 构建前：build/、dist/%s/ 内容、*.spec" % NAME)
    _rm(BUILD_DIR)
    # 只清空 dist/<NAME>/ 的内容，不删目录本身（可能被 IDE/资源管理器占用）
    if os.path.isdir(APP_DIR):
        for entry in os.listdir(APP_DIR):
            _rm(os.path.join(APP_DIR, entry))
    for fn in os.listdir(ROOT):
        if fn.endswith(".spec") and os.path.isfile(os.path.join(ROOT, fn)):
            print("       删除 %s" % fn)
            _rm(os.path.join(ROOT, fn))


def build() -> None:
    """调用 PyInstaller 生成多文件、带 cmd 窗口的程序。"""
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",              # 覆盖旧产物不再询问
        "--clean",                  # 清 PyInstaller 自身缓存
        "--console",                # 带 cmd 窗口（run/gui 日志可见）
        "--onedir",                 # 多文件模式（非单文件）
        "--noupx",                  # UPX 压 Qt 的 DLL 易出问题，直接不用
        "--name", NAME,
        "--paths", ROOT,            # 让自建包在分析期可被 import
        "--distpath", STAGE_DIR,    # 先出到暂存目录，见 install()
        "--workpath", BUILD_DIR,
        "--specpath", BUILD_DIR,    # .spec 生成到 build/，随构建残留一起清掉
        "--log-level", "WARN",
    ]
    for pkg in PACKAGES:
        cmd += ["--collect-submodules", pkg]
    for mod in EXCLUDES:
        cmd += ["--exclude-module", mod]
    cmd.append(ENTRY)

    print("[构建] PyInstaller 开始打包 %s（多文件 / 带 cmd 窗口）" % ENTRY)
    print("       入口 -> dist/%s/%s.exe" % (NAME, NAME))
    started = time.time()
    ret = subprocess.call(cmd, cwd=ROOT)
    if ret != 0:
        raise SystemExit("[构建] 失败，PyInstaller 退出码 %d" % ret)
    print("[构建] 完成，用时 %.1f 秒" % (time.time() - started))


def install() -> None:
    """把暂存产物搬进 dist/<NAME>/（目录被占用时也能成功）。"""
    if not os.path.isdir(STAGE_APP):
        raise SystemExit("[安装] 找不到暂存产物 %s" % os.path.relpath(STAGE_APP, ROOT))

    os.makedirs(APP_DIR, exist_ok=True)
    for entry in os.listdir(APP_DIR):
        _rm(os.path.join(APP_DIR, entry))
    for entry in os.listdir(STAGE_APP):
        shutil.move(os.path.join(STAGE_APP, entry), os.path.join(APP_DIR, entry))
    print("[安装] 产物 -> dist/%s/" % NAME)


def make_config_dir() -> None:
    """预建 config/ 用户数据目录（模板、AI 配置、章节源由用户放入）。"""
    cfg = os.path.join(APP_DIR, "config")
    os.makedirs(cfg, exist_ok=True)
    for sub in CONFIG_SUBDIRS:
        os.makedirs(os.path.join(cfg, sub), exist_ok=True)
    print("[预建] config/ 用户数据目录（放 template.docx、ai.json 等）")


def clean_after() -> None:
    """清掉构建残留：build/（含 .spec）、以及产物内的 __pycache__ / *.pyc。"""
    print("[清理] 构建后：build/、__pycache__、*.pyc/*.pyo")
    _rm(BUILD_DIR)
    n_src = _clean_pycache(ROOT, SOURCE_SKIP_DIRS)
    n_out = _clean_pycache(APP_DIR, set())
    print("       源码清理 %d 项，产物清理 %d 项" % (n_src, n_out))


def verify() -> None:
    """校验关键文件都在，避免打出个跑不起来的包。"""
    required = [
        os.path.join(APP_DIR, NAME + ".exe"),
        os.path.join(APP_DIR, "_internal", "base_library.zip"),
        os.path.join(APP_DIR, "config"),
    ]
    missing = [os.path.relpath(p, ROOT) for p in required if not os.path.exists(p)]
    if missing:
        for p in missing:
            print("[校验] 缺失：%s" % p)
        raise SystemExit("[校验] 失败，产物不完整")

    print("[校验] 关键文件齐备")
    print("[校验] 依赖模块：")
    for sub in ("PyQt6", "PIL", "av"):
        p = os.path.join(APP_DIR, "_internal", sub)
        print("       %-10s %s" % (sub, "OK" if os.path.exists(p) else "未打包（非必需）"))


def main() -> None:
    print("=" * 60)
    print("打包 %s（多文件 / 带 cmd 窗口 / 免 Python）" % NAME)
    print("=" * 60)

    clean_before()
    build()
    install()
    make_config_dir()
    clean_after()
    verify()

    print("=" * 60)
    print("产物：%s" % os.path.relpath(APP_DIR, ROOT))
    print("体积：%s" % _human(_dir_size(APP_DIR)))
    print("分发：整个目录拷到目标机器，双击 %s.exe 即可，无需安装 Python。" % NAME)
    print("      首次使用把学校模板放到 config/template.docx；")
    print("      soffice/pdftotext 缺失时页码回填自动降级，不影响出稿。")
    print("=" * 60)


if __name__ == "__main__":
    main()

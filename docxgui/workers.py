"""后台任务：解析（QThread）、AI 结构修正/内容改写（QThread）、构建（QProcess）。

UI 线程不跑任何阻塞任务；结果一律通过信号回传，取消通过标志位或 kill 进程。
"""
import os
import sys
from typing import List, Optional, Set

from PyQt6.QtCore import QObject, QProcess, QThread, pyqtSignal

from docxai.client import AIClient, AIError
from docxai.correct import CorrectionResult, correct_structure, rewrite_text
from docxconvert.cli import prepare_blocks


class ParseWorker(QThread):
    """规则解析 + 图片抽取（不写 md），供审阅页使用。"""

    finished_blocks = pyqtSignal(list, str)   # blocks, images_dir
    failed = pyqtSignal(str)

    def __init__(self, docx_path: str, images_dir: str, rel_base: str,
                 strip_front: bool):
        super().__init__()
        self._args = (docx_path, images_dir, rel_base, strip_front)

    def run(self):
        try:
            blocks, images_dir, _ = prepare_blocks(
                self._args[0], self._args[1],
                strip_front=self._args[3], rel_base=self._args[2], quiet=True)
            self.finished_blocks.emit(blocks, images_dir)
        except Exception as e:   # 解析失败要回到界面提示，不让线程静默死掉
            self.failed.emit(str(e))


class StructureWorker(QThread):
    """整篇 AI 结构修正（内部按章分批，护栏在 docxai.correct）。"""

    progress = pyqtSignal(int, int, str)       # done, total, chapter
    finished_result = pyqtSignal(object)       # CorrectionResult
    failed = pyqtSignal(str)

    def __init__(self, client: AIClient, blocks: list,
                 known_images: Optional[Set[str]] = None):
        super().__init__()
        self._client = client
        self._blocks = blocks
        self._images = known_images
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            result: CorrectionResult = correct_structure(
                self._client, self._blocks, known_images=self._images,
                on_progress=lambda d, t, name: self.progress.emit(d, t, name),
                cancel=lambda: self._cancel)
            self.finished_result.emit(result)
        except AIError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit(f"AI 结构修正异常：{e}")


class RewriteWorker(QThread):
    """单段文本的 AI 内容改写。"""

    finished_text = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, client: AIClient, text: str, instruction: str):
        super().__init__()
        self._client, self._text, self._instruction = client, text, instruction
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            out = rewrite_text(self._client, self._text, self._instruction,
                               cancel=lambda: self._cancel)
            self.finished_text.emit(out)
        except AIError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit(f"AI 改写异常：{e}")


class RunProcess(QObject):
    """以外部进程执行 `python3 main.py run`，流式回传日志，可终止。

    构建逻辑零重写：这里只起进程，输出什么日志界面就显示什么，
    与命令行行为永远一致。
    """

    log = pyqtSignal(str)
    finished_ok = pyqtSignal(bool)   # True 退出码 0

    def __init__(self, tool_root: str):
        super().__init__()
        self._proc: Optional[QProcess] = None
        self._root = tool_root

    def start(self, template: Optional[str], hash16: str):
        program = sys.executable
        args = [os.path.join(self._root, "main.py"), "run", hash16]
        if template:
            args += ["--template", template]
        self._proc = QProcess(self)
        self._proc.setWorkingDirectory(self._root)
        self._proc.setProcessChannelMode(
            QProcess.ProcessChannelMode.MergedChannels)
        self._proc.readyReadStandardOutput.connect(self._on_output)
        self._proc.finished.connect(self._on_finished)
        self._proc.start(program, args)

    def _on_output(self):
        data = bytes(self._proc.readAllStandardOutput()).decode("utf-8", "replace")
        if data:
            self.log.emit(data)

    def _on_finished(self, code, _status):
        self.finished_ok.emit(code == 0)

    def stop(self):
        if self._proc is not None and self._proc.state() != QProcess.ProcessState.NotRunning:
            self._proc.kill()
